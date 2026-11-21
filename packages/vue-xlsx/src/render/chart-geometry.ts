import { scaleBand, scalePoint } from "d3-scale";
import type {
  ChartAxisGeometry,
  ChartAxisTickGeometry,
  ChartGeometryFrame,
  ChartGeometryLabel,
  ChartGeometryLimitation,
  ChartGeometryPoint,
  ChartGeometryProvider,
  ChartGeometryRect,
  ChartGeometryRequest,
  ChartGeometrySnapshot,
  ChartLegendGeometry,
  ChartLegendItemGeometry,
  ChartLegendPosition,
  ChartSeriesGeometry,
  ChartSeriesPointGeometry,
} from "@arcships/office-interaction";
import type { XlsxChart, XlsxChartAxis, XlsxChartSeries } from "@arcships/xlsx-core";
import { buildLayout, getLegendItems } from "./chart-legend";
import {
  buildComboGroups,
  buildPieEntries,
  getCategoryLabels,
  getCategoryNumericValues,
  isComboChart,
  isHistogramLikeChart,
  isHistogramLikeSeries,
  normalizeRenderableChartType,
  resolveCategoryBandPadding,
  resolveRenderableSeriesValue,
} from "./chart-data";
import { isLikelyDateFormatCode } from "./chart-date";
import {
  buildDateAxisTicks,
  buildNumericTickValues,
  createDateScale,
  createValueScale,
  resolveAxisDomainWithChartOverrides,
  resolveAxisLogBase,
  resolveCategoryLabelLayout,
  resolveDateAxisDomain,
  resolveLogAxisDomain,
} from "./chart-axis";
import { createChartAxisTickFormatter } from "./chart-data";
import {
  estimateSvgTextWidth,
  normalizeLegendPosition,
  safeNumber,
} from "./chart-shared";
import type {
  ChartLabels,
  ChartLayout,
  ChartRendererPalette,
  LegendItem,
  PlotRect,
} from "./chart-types";

export const XLSX_CHART_GEOMETRY_PROVIDER_ID = "xlsx-svg-renderer";

const LEGEND_FONT_SIZE = 10;
const LEGEND_SWATCH_SIZE = 8;
const LEGEND_TEXT_OFFSET = 12;
const LEGEND_LINE_HEIGHT = 12;
const LEGEND_ITEM_GAP = 6;
const LEGEND_PADDING = 8;

const DEFAULT_GEOMETRY_PALETTE: ChartRendererPalette = {
  border: "#cbd5e1",
  mutedText: "#64748b",
  surface: "#ffffff",
  text: "#0f172a",
};

export interface XlsxChartGeometryOptions {
  /** Chart frame in chart-local CSS pixels. The renderer rounds it like `buildLayout`. */
  frame: ChartGeometryFrame;
  /** Colors are irrelevant to geometry; kept so legend items use the same palette. */
  palette?: ChartRendererPalette;
  labels?: ChartLabels;
}

const TEXT_METRICS_LIMITATION: ChartGeometryLimitation = {
  code: "xlsx.text-metrics-estimated",
  detail: "Label boxes use the renderer's estimated SVG text metrics; glyph bounds depend on the loaded font.",
  confidence: "approximate",
};

const MANUAL_LAYOUT_LIMITATION: ChartGeometryLimitation = {
  code: "xlsx.manual-layout-not-parsed",
  detail: "xlsx-core does not surface c:plotArea/c:layout/c:manualLayout yet; this provider reports the automatic renderer layout.",
  confidence: "approximate",
};

const THIRD_DIMENSION_LIMITATION: ChartGeometryLimitation = {
  code: "xlsx.3d-projection-provider",
  detail: "3-D marks are projected by the renderer; no 2-D rectangle or point geometry is exposed for them.",
  confidence: "provider",
};

function makeRect(
  x: number,
  y: number,
  width: number,
  height: number,
): ChartGeometryRect {
  return { x, y, width, height };
}

function pointOfRect(rect: ChartGeometryRect): ChartGeometryPoint {
  return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
}

function unionRects(rects: readonly ChartGeometryRect[]): ChartGeometryRect {
  const left = Math.min(...rects.map((rect) => rect.x));
  const top = Math.min(...rects.map((rect) => rect.y));
  const right = Math.max(...rects.map((rect) => rect.x + rect.width));
  const bottom = Math.max(...rects.map((rect) => rect.y + rect.height));
  return makeRect(left, top, right - left, bottom - top);
}

function estimateLabelRect(
  text: string,
  anchor: ChartGeometryPoint,
  textAnchor: "start" | "middle" | "end",
): ChartGeometryRect {
  const width = estimateSvgTextWidth(text, LEGEND_FONT_SIZE);
  const x = textAnchor === "middle" ? anchor.x - width / 2 : textAnchor === "end" ? anchor.x - width : anchor.x;
  return makeRect(x, anchor.y - LEGEND_FONT_SIZE, width, LEGEND_LINE_HEIGHT + 2);
}

function estimateRotatedLabelRect(
  text: string,
  anchor: ChartGeometryPoint,
): ChartGeometryRect {
  const width = estimateSvgTextWidth(text, LEGEND_FONT_SIZE);
  const projected = width * Math.SQRT1_2;
  return makeRect(
    anchor.x - projected - 7,
    anchor.y - 7,
    projected + 9,
    projected + 9,
  );
}

function normalizeTickLabelPosition(axis: XlsxChartAxis | null | undefined) {
  return axis?.labelPosition === "high" || axis?.labelPosition === "low" || axis?.labelPosition === "none"
    ? axis.labelPosition
    : "nextTo";
}

function legendPositionOf(rawPosition: string | undefined, layoutPosition: string | undefined): ChartLegendPosition {
  const normalized = normalizeLegendPosition(layoutPosition ?? rawPosition);
  switch (normalized) {
    case "left":
    case "right":
    case "top":
    case "bottom":
      return normalized;
    default:
      return "unknown";
  }
}

function resolveComboSeriesIndex(chart: XlsxChart, series: XlsxChartSeries): number {
  const byId = chart.series.findIndex((candidate) => candidate.id === series.id);
  if (byId >= 0) {
    return byId;
  }
  const flattened = buildComboGroups(chart).flatMap((group) => group.series);
  const byOrder = flattened.indexOf(series);
  return byOrder >= 0 && byOrder < chart.series.length ? byOrder : 0;
}

function resolveLegendEntryIdentity(
  chart: XlsxChart,
  chartType: string,
  index: number,
): Pick<ChartLegendItemGeometry, "seriesIndex" | "pointIndex"> {
  if (chartType === "Pie" || chartType === "Pie3D" || chartType === "PieExploded"
    || chartType === "Doughnut" || chartType === "BarOfPie") {
    const entries = chartType === "BarOfPie"
      ? getCategoryLabels(chart)
          .map((label, entryIndex) => ({
            index: entryIndex,
            value: safeNumber(chart.series[0]?.values[entryIndex]) ?? 0,
          }))
          .filter((entry) => entry.value > 0)
      : buildPieEntries(chart);
    return { pointIndex: entries[index]?.index ?? index };
  }
  if (chartType === "Sunburst" || chartType === "Treemap") {
    return {};
  }
  if (isComboChart(chart)) {
    const groupSeries = buildComboGroups(chart).flatMap((group) => group.series);
    const series = groupSeries[index];
    return series ? { seriesIndex: resolveComboSeriesIndex(chart, series) } : {};
  }
  return index < chart.series.length ? { seriesIndex: index } : {};
}

function buildLegendGeometry(
  chart: XlsxChart,
  chartType: string,
  layout: ChartLayout,
  legendItems: LegendItem[],
): ChartLegendGeometry | undefined {
  if (!chart.legend || legendItems.length === 0) {
    return undefined;
  }
  const metrics = layout.legendMetrics;
  if (!metrics) {
    return undefined;
  }
  const position = legendPositionOf(chart.legend.position, layout.legendPosition);
  const items: ChartLegendItemGeometry[] = [];
  const entryWidth = (index: number) => {
    const lines = metrics.entryLines[index] ?? [legendItems[index]?.label ?? ""];
    return LEGEND_TEXT_OFFSET
      + Math.max(...lines.map((line) => estimateSvgTextWidth(line, LEGEND_FONT_SIZE)))
      + 8;
  };

  if (position === "left" || position === "right") {
    const x = position === "right"
      ? layout.plot.left + layout.plot.width + 8 + (layout.secondaryAxisWidth ?? 0)
      : 8;
    let cursorY = layout.plot.top + 6;
    legendItems.forEach((item, index) => {
      const lines = metrics.entryLines[index] ?? [item.label];
      const rowHeight = metrics.rowHeights[index] ?? (lines.length * LEGEND_LINE_HEIGHT + LEGEND_ITEM_GAP);
      const width = entryWidth(index);
      const maxLineWidth = Math.max(...lines.map((line) => estimateSvgTextWidth(line, LEGEND_FONT_SIZE)));
      const labelAnchor = { x: x + LEGEND_TEXT_OFFSET, y: cursorY };
      items.push({
        index,
        ...resolveLegendEntryIdentity(chart, chartType, index),
        label: item.label,
        rect: makeRect(x, cursorY - 7, width, rowHeight),
        swatchRect: makeRect(x, cursorY - 7, LEGEND_SWATCH_SIZE, LEGEND_SWATCH_SIZE),
        labelAnchor,
        lines,
        labelRect: makeRect(labelAnchor.x, labelAnchor.y - LEGEND_FONT_SIZE, maxLineWidth, lines.length * LEGEND_LINE_HEIGHT + 2),
        confidence: "approximate",
      });
      cursorY += rowHeight;
    });
  } else {
    const startY = position === "top"
      ? layout.titleHeight + LEGEND_PADDING
      : Math.max(layout.plot.top + layout.plot.height + 6, layout.height - metrics.height + LEGEND_PADDING);
    let rowTop = startY;
    metrics.rows.forEach((row, rowIndex) => {
      const rowHeight = metrics.rowHeights[rowIndex] ?? LEGEND_LINE_HEIGHT + 2;
      const totalWidth = row.reduce((sum, index) => sum + entryWidth(index) + LEGEND_ITEM_GAP, 0);
      let cursorX = Math.max(8, (layout.width - totalWidth) / 2);
      row.forEach((index) => {
        const item = legendItems[index];
        if (!item) {
          return;
        }
        const lines = metrics.entryLines[index] ?? [item.label];
        const width = entryWidth(index);
        const x = cursorX;
        cursorX += width + LEGEND_ITEM_GAP;
        const labelAnchor = { x: x + LEGEND_TEXT_OFFSET, y: rowTop + 9 };
        const maxLineWidth = Math.max(...lines.map((line) => estimateSvgTextWidth(line, LEGEND_FONT_SIZE)));
        items.push({
          index,
          ...resolveLegendEntryIdentity(chart, chartType, index),
          label: item.label,
          rect: makeRect(x, rowTop + 2, width, rowHeight),
          swatchRect: makeRect(x, rowTop + 2, LEGEND_SWATCH_SIZE, LEGEND_SWATCH_SIZE),
          labelAnchor,
          lines,
          labelRect: makeRect(labelAnchor.x, labelAnchor.y - LEGEND_FONT_SIZE, maxLineWidth, lines.length * LEGEND_LINE_HEIGHT + 2),
          confidence: "approximate",
        });
      });
      rowTop += rowHeight;
    });
  }

  const rect = items.length > 0
    ? unionRects(items.map((item) => item.rect))
    : makeRect(0, 0, metrics.width, metrics.height);
  return {
    position,
    rect,
    items,
    confidence: "approximate",
  };
}

function buildCategoryAxisGeometry(input: {
  axis: XlsxChartAxis | null | undefined;
  kind: "category" | "date";
  labels: readonly string[];
  positions: readonly number[];
  plot: PlotRect;
  isHorizontal: boolean;
}): ChartAxisGeometry {
  const { axis, kind, labels, positions, plot, isHorizontal } = input;
  const labelPosition = normalizeTickLabelPosition(axis);
  const showLabels = axis?.delete !== true && labelPosition !== "none";
  const tickLabelSkip = typeof axis?.tickLabelSkip === "number" && axis.tickLabelSkip > 0
    ? Math.max(1, Math.floor(axis.tickLabelSkip))
    : 1;
  const labelLayout = isHorizontal
    ? {
        rotate: false,
        step: Math.max(
          1,
          tickLabelSkip,
          Math.ceil(12 / Math.max(1, plot.height / Math.max(1, labels.length))),
        ),
      }
    : resolveCategoryLabelLayout([...labels], plot.width, tickLabelSkip);
  const labelsAtHighEnd = labelPosition === "high";
  const ticks: ChartAxisTickGeometry[] = labels.map((label, index) => {
    const position = isHorizontal
      ? { x: plot.left, y: positions[index] ?? plot.top }
      : { x: positions[index] ?? plot.left, y: plot.top + plot.height };
    const labelVisible = showLabels && index % labelLayout.step === 0;
    if (!labelVisible) {
      return {
        index,
        categoryIndex: index,
        label,
        position,
        labelVisible: false,
        labelConfidence: "approximate",
      };
    }
    if (isHorizontal) {
      const labelAnchor = {
        x: labelsAtHighEnd ? plot.left + plot.width + 6 : plot.left - 6,
        y: position.y + 3,
      };
      return {
        index,
        categoryIndex: index,
        label,
        position,
        labelAnchor,
        labelRect: estimateLabelRect(label, labelAnchor, labelsAtHighEnd ? "start" : "end"),
        labelVisible: true,
        labelConfidence: "approximate",
      };
    }
    const labelAnchor = labelsAtHighEnd
      ? { x: position.x, y: plot.top - 4 }
      : { x: position.x, y: plot.top + plot.height + 13 };
    return {
      index,
      categoryIndex: index,
      label,
      position,
      labelAnchor,
      labelRect: labelLayout.rotate
        ? estimateRotatedLabelRect(label, labelAnchor)
        : estimateLabelRect(label, labelAnchor, labelsAtHighEnd ? "middle" : labelLayout.rotate ? "end" : "middle"),
      labelVisible: true,
      labelConfidence: "approximate",
    };
  });
  const line = axis?.delete === true
    ? undefined
    : isHorizontal
      ? {
          start: { x: plot.left, y: plot.top },
          end: { x: plot.left, y: plot.top + plot.height },
        }
      : {
          start: { x: plot.left, y: plot.top + plot.height },
          end: { x: plot.left + plot.width, y: plot.top + plot.height },
        };
  return {
    axisId: axis?.id,
    kind,
    title: axis?.title,
    ...(line ? { line } : {}),
    ticks,
    confidence: "exact",
  };
}

type ValueAxisPlacement = "left" | "right" | "bottom" | "top";

function buildValueAxisGeometry(input: {
  axis: XlsxChartAxis | null | undefined;
  chart: XlsxChart;
  ticks: readonly number[];
  scale: (value: number) => number;
  plot: PlotRect;
  placement: ValueAxisPlacement;
  percentValuesAreScaled: boolean;
  line?: { start: ChartGeometryPoint; end: ChartGeometryPoint };
  zeroLine?: { start: ChartGeometryPoint; end: ChartGeometryPoint };
}): ChartAxisGeometry {
  const {
    axis,
    chart,
    ticks,
    scale,
    plot,
    placement,
    percentValuesAreScaled,
    line,
    zeroLine,
  } = input;
  const labelPosition = normalizeTickLabelPosition(axis);
  const showLabels = axis?.delete !== true && labelPosition !== "none";
  const labelSkip = typeof axis?.tickLabelSkip === "number" && axis.tickLabelSkip > 0
    ? Math.max(1, Math.floor(axis.tickLabelSkip))
    : 1;
  const formatTick = createChartAxisTickFormatter(axis, chart, { percentValuesAreScaled });
  const vertical = placement === "left" || placement === "right";
  const ticksGeometry: ChartAxisTickGeometry[] = ticks.map((value, index) => {
    const offset = scale(value);
    const position = vertical
      ? { x: placement === "right" ? plot.left + plot.width : plot.left, y: offset }
      : { x: offset, y: placement === "top" ? plot.top : plot.top + plot.height };
    const labelVisible = showLabels && index % labelSkip === 0;
    const label = formatTick(value);
    if (!labelVisible) {
      return {
        index,
        value,
        label,
        position,
        labelVisible: false,
        labelConfidence: "approximate",
      };
    }
    let labelAnchor: ChartGeometryPoint;
    let textAnchor: "start" | "middle" | "end";
    if (placement === "left") {
      labelAnchor = { x: plot.left - 6, y: offset + 3 };
      textAnchor = "end";
    } else if (placement === "right") {
      labelAnchor = { x: plot.left + plot.width + 6, y: offset + 3 };
      textAnchor = "start";
    } else if (placement === "top") {
      labelAnchor = { x: offset, y: plot.top - 4 };
      textAnchor = "middle";
    } else {
      labelAnchor = { x: offset, y: plot.top + plot.height + 14 };
      textAnchor = "middle";
    }
    return {
      index,
      value,
      label,
      position,
      labelAnchor,
      labelRect: estimateLabelRect(label, labelAnchor, textAnchor),
      labelVisible: true,
      labelConfidence: "approximate",
    };
  });
  return {
    axisId: axis?.id,
    kind: "value",
    title: axis?.title,
    ...(line ? { line } : {}),
    ...(zeroLine ? { zeroLine } : {}),
    ticks: ticksGeometry,
    confidence: "exact",
  };
}

function createValueScaleRange(
  axis: XlsxChartAxis | null | undefined,
  domain: { min: number; max: number },
  isHorizontal: boolean,
  plot: PlotRect,
) {
  const reverse = axis?.orientation === "maxMin";
  if (isHorizontal) {
    return createValueScale(
      axis,
      domain.min,
      domain.max,
      reverse ? [plot.left + plot.width, plot.left] : [plot.left, plot.left + plot.width],
    );
  }
  return createValueScale(
    axis,
    domain.min,
    domain.max,
    reverse ? [plot.top, plot.top + plot.height] : [plot.top + plot.height, plot.top],
  );
}

function valueAxisLine(
  placement: ValueAxisPlacement,
  plot: PlotRect,
): { start: ChartGeometryPoint; end: ChartGeometryPoint } {
  switch (placement) {
    case "right":
      return {
        start: { x: plot.left + plot.width, y: plot.top },
        end: { x: plot.left + plot.width, y: plot.top + plot.height },
      };
    case "bottom":
      return {
        start: { x: plot.left, y: plot.top + plot.height },
        end: { x: plot.left + plot.width, y: plot.top + plot.height },
      };
    case "top":
      return {
        start: { x: plot.left, y: plot.top },
        end: { x: plot.left + plot.width, y: plot.top },
      };
    default:
      return {
        start: { x: plot.left, y: plot.top },
        end: { x: plot.left, y: plot.top + plot.height },
      };
  }
}

function valueAxisZeroLine(
  placement: ValueAxisPlacement,
  zero: number,
  plot: PlotRect,
): { start: ChartGeometryPoint; end: ChartGeometryPoint } {
  if (placement === "left" || placement === "right") {
    return {
      start: { x: plot.left, y: zero },
      end: { x: plot.left + plot.width, y: zero },
    };
  }
  return {
    start: { x: zero, y: plot.top },
    end: { x: zero, y: plot.top + plot.height },
  };
}

function resolveZeroPosition(
  scale: (value: number) => number,
  fallback: number,
): number {
  const value = scale(0);
  return Number.isFinite(value) ? value : fallback;
}

function seriesName(series: XlsxChartSeries, index: number) {
  const name = series.name?.trim();
  return name && name.length > 0 ? name : `Series ${index + 1}`;
}

function markerRect(position: ChartGeometryPoint, series: XlsxChartSeries): ChartGeometryRect {
  const radius = Math.max(3, (series.markerSize ?? 7) * 0.5);
  return makeRect(position.x - radius, position.y - radius, radius * 2, radius * 2);
}

function valueDomainForSeries(
  axis: XlsxChartAxis | null | undefined,
  values: readonly number[],
  includeZero: boolean,
) {
  if (resolveAxisLogBase(axis)) {
    const positive = values.filter((value) => value > 0);
    return resolveLogAxisDomain(
      axis,
      positive.length > 0 ? Math.min(...positive) : 1,
      positive.length > 0 ? Math.max(...positive) : 10,
    );
  }
  return resolveAxisDomainWithChartOverrides(
    axis,
    values.length > 0 ? Math.min(...values) : 0,
    values.length > 0 ? Math.max(...values) : 1,
    includeZero,
  );
}

function buildBarGeometry(
  chart: XlsxChart,
  chartType: string,
  layout: ChartLayout,
  limitations: ChartGeometryLimitation[],
): { axes: ChartAxisGeometry[]; series: ChartSeriesGeometry[] } {
  if (chart.is3d) {
    limitations.push(THIRD_DIMENSION_LIMITATION);
    return { axes: [], series: [] };
  }
  const sourceCategories = getCategoryLabels(chart);
  if (sourceCategories.length === 0 || chart.series.length === 0) {
    return { axes: [], series: [] };
  }
  const isHorizontal = /^Bar(?!OfPie)/.test(chartType);
  const shouldReverseCategories = isHorizontal && chart.categoryAxis?.orientation !== "maxMin";
  const categories = shouldReverseCategories ? sourceCategories.slice().reverse() : sourceCategories;
  const isPercentStacked = chartType === "ColumnPercentStacked" || chartType === "BarPercentStacked";
  const isStacked = chartType === "ColumnStacked" || chartType === "BarStacked" || isPercentStacked;
  const isHistogramLike = isHistogramLikeChart(chart) && chartType === "ColumnClustered";
  const categoryCount = categories.length;
  const seriesCount = chart.series.length;
  const plot = layout.plot;
  const categoryBandPadding = isHistogramLike
    ? { inner: 0, outer: 0 }
    : resolveCategoryBandPadding(chart.gapWidth);
  const categoryScale = scaleBand<string>()
    .domain(categories)
    .range(isHorizontal ? [plot.top, plot.top + plot.height] : [plot.left, plot.left + plot.width])
    .paddingInner(categoryBandPadding.inner)
    .paddingOuter(categoryBandPadding.outer);
  const seriesScale = scaleBand<string>()
    .domain(Array.from({ length: seriesCount }, (_, index) => String(index)))
    .range([0, categoryScale.bandwidth()])
    .paddingInner(isHistogramLike ? 0 : 0.16)
    .paddingOuter(isHistogramLike ? 0 : 0.08);
  const matrix = chart.series.map((series) => {
    const values = Array.from(
      { length: categoryCount },
      (_, categoryIndex) => safeNumber(series.values[categoryIndex]) ?? 0,
    );
    return shouldReverseCategories ? values.reverse() : values;
  });
  const rawMatrix = matrix.map((row) => row.slice());
  if (isPercentStacked) {
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex += 1) {
      const total = matrix.reduce((sum, row) => sum + Math.abs(row[categoryIndex] ?? 0), 0);
      for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex += 1) {
        matrix[seriesIndex][categoryIndex] = total > 0
          ? ((matrix[seriesIndex][categoryIndex] ?? 0) / total) * 100
          : 0;
      }
    }
  }
  let minValue = Number.POSITIVE_INFINITY;
  let maxValue = Number.NEGATIVE_INFINITY;
  if (isStacked) {
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex += 1) {
      let positive = 0;
      let negative = 0;
      for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex += 1) {
        const value = matrix[seriesIndex][categoryIndex] ?? 0;
        if (value >= 0) {
          positive += value;
        } else {
          negative += value;
        }
      }
      maxValue = Math.max(maxValue, positive);
      minValue = Math.min(minValue, negative);
    }
  } else {
    for (const row of matrix) {
      for (const value of row) {
        maxValue = Math.max(maxValue, value);
        minValue = Math.min(minValue, value);
      }
    }
  }
  if (!Number.isFinite(minValue) || !Number.isFinite(maxValue)) {
    minValue = 0;
    maxValue = 1;
  }
  const valueDomain = valueDomainForSeries(chart.valueAxis, matrix.flat(), true);
  const scale = createValueScaleRange(chart.valueAxis, valueDomain, isHorizontal, plot);
  const ticks = resolveAxisLogBase(chart.valueAxis)
    ? valueDomain.ticks
    : isPercentStacked
      ? buildNumericTickValues(valueDomain.min, valueDomain.max, chart.valueAxis?.majorUnit ?? 20)
      : valueDomain.ticks;
  const categoryPositions = categories.map((category) => (
    (categoryScale(category) ?? 0) + categoryScale.bandwidth() / 2
  ));
  const categoryAxis = buildCategoryAxisGeometry({
    axis: chart.categoryAxis,
    kind: chart.categoryAxis?.axisType === "date" ? "date" : "category",
    labels: categories,
    positions: categoryPositions,
    plot,
    isHorizontal,
  });
  const valuePlacement: ValueAxisPlacement = isHorizontal
    ? (normalizeTickLabelPosition(chart.valueAxis) === "high" ? "top" : "bottom")
    : (normalizeTickLabelPosition(chart.valueAxis) === "high" ? "right" : "left");
  const zero = resolveZeroPosition(scale, isHorizontal ? plot.left : plot.top + plot.height);
  const valueAxis = buildValueAxisGeometry({
    axis: chart.valueAxis,
    chart,
    ticks,
    scale,
    plot,
    placement: valuePlacement,
    percentValuesAreScaled: isPercentStacked,
    line: chart.valueAxis?.delete === true ? undefined : valueAxisLine(valuePlacement, plot),
    zeroLine: chart.valueAxis?.delete === true
      ? undefined
      : valueAxisZeroLine(valuePlacement, zero, plot),
  });

  const positive = Array.from({ length: categoryCount }, () => 0);
  const negative = Array.from({ length: categoryCount }, () => 0);
  const seriesGeometry: ChartSeriesGeometry[] = chart.series.map((series, seriesIndex) => ({
    index: seriesIndex,
    seriesId: series.id,
    name: seriesName(series, seriesIndex),
    shape: "bar",
    categoryAxisIndex: 0,
    valueAxisIndex: 1,
    points: [],
    confidence: "exact",
  }));
  for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex += 1) {
    const points: ChartSeriesPointGeometry[] = [];
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex += 1) {
      const value = matrix[seriesIndex][categoryIndex] ?? 0;
      const rawValue = rawMatrix[seriesIndex][categoryIndex] ?? value;
      let start = 0;
      let end = value;
      if (isStacked) {
        if (value >= 0) {
          start = positive[categoryIndex];
          positive[categoryIndex] += value;
          end = positive[categoryIndex];
        } else {
          start = negative[categoryIndex];
          negative[categoryIndex] += value;
          end = negative[categoryIndex];
        }
      }
      const category = categories[categoryIndex];
      const categoryStart = categoryScale(category) ?? 0;
      const barThickness = isHistogramLike
        ? categoryScale.bandwidth()
        : isStacked
          ? categoryScale.bandwidth()
          : seriesScale.bandwidth();
      const barOffset = isHistogramLike || isStacked
        ? 0
        : (seriesScale(String(seriesIndex)) ?? 0);
      const logBase = resolveAxisLogBase(chart.valueAxis);
      const scaleStart = logBase ? Math.max(start, valueDomain.min) : start;
      const scaleEnd = logBase ? Math.max(end, valueDomain.min) : end;
      let rect: ChartGeometryRect;
      if (isHorizontal) {
        const x1 = scale(scaleStart);
        const x2 = scale(scaleEnd);
        rect = makeRect(
          Math.min(x1, x2),
          categoryStart + barOffset,
          Math.max(1, Math.abs(x2 - x1)),
          Math.max(1, barThickness),
        );
      } else {
        const y1 = scale(scaleStart);
        const y2 = scale(scaleEnd);
        rect = makeRect(
          categoryStart + barOffset,
          Math.min(y1, y2),
          Math.max(1, barThickness),
          Math.max(1, Math.abs(y2 - y1)),
        );
      }
      const sourceIndex = shouldReverseCategories ? categoryCount - 1 - categoryIndex : categoryIndex;
      points.push({
        index: sourceIndex,
        categoryIndex,
        value: rawValue,
        position: pointOfRect(rect),
        rect,
        confidence: "exact",
        rectConfidence: "exact",
      });
    }
    seriesGeometry[seriesIndex].points = points;
  }
  return { axes: [categoryAxis, valueAxis], series: seriesGeometry };
}

function buildLineGeometry(
  chart: XlsxChart,
  chartType: string,
  layout: ChartLayout,
  limitations: ChartGeometryLimitation[],
): { axes: ChartAxisGeometry[]; series: ChartSeriesGeometry[] } {
  if (chart.is3d) {
    limitations.push(THIRD_DIMENSION_LIMITATION);
    return { axes: [], series: [] };
  }
  const categories = getCategoryLabels(chart);
  if (categories.length === 0 || chart.series.length === 0) {
    return { axes: [], series: [] };
  }
  const plot = layout.plot;
  const isAreaChart = chartType === "Area" || chartType === "AreaStacked" || chartType === "AreaPercentStacked";
  const isStackedArea = chartType === "AreaStacked" || chartType === "AreaPercentStacked";
  const isStackedLine = chartType === "LineStacked" || chartType === "LinePercentStacked";
  const isStackedSeries = isStackedArea || isStackedLine;
  const isPercentStackedSeries = chartType === "AreaPercentStacked" || chartType === "LinePercentStacked";
  const resolvedValuesBySeries = chart.series.map((series) => (
    categories.map((_, categoryIndex) => resolveRenderableSeriesValue(
      series.values[categoryIndex],
      chart.displayBlanksAs,
    ))
  ));
  const categoryTotals = isPercentStackedSeries
    ? categories.map((_, categoryIndex) => {
        const total = resolvedValuesBySeries.reduce(
          (sum, values) => sum + Math.abs(values[categoryIndex] ?? 0),
          0,
        );
        return Math.abs(total) < 1e-9 ? 1 : total;
      })
    : null;
  const positive = Array.from({ length: categories.length }, () => 0);
  const negative = Array.from({ length: categories.length }, () => 0);
  const stackedPoints = chart.series.map((_, seriesIndex) => categories.map((_, categoryIndex) => {
    const rawValue = resolvedValuesBySeries[seriesIndex]?.[categoryIndex] ?? null;
    if (rawValue == null) {
      return { defined: false, value: null as number | null, y: null as number | null };
    }
    const value = isPercentStackedSeries
      ? (rawValue / (categoryTotals?.[categoryIndex] ?? 1)) * 100
      : rawValue;
    if (value >= 0) {
      const start = positive[categoryIndex];
      const end = start + value;
      positive[categoryIndex] = end;
      return { defined: true, value, y: end };
    }
    const start = negative[categoryIndex];
    const end = start + value;
    negative[categoryIndex] = end;
    return { defined: true, value, y: end };
  }));
  const extents = stackedPoints.flatMap((points) => points.flatMap((point) => (
    point.defined && point.y != null ? [point.y] : []
  )));
  if (extents.length === 0) {
    return { axes: [], series: [] };
  }
  let minValue = Math.min(...extents);
  let maxValue = Math.max(...extents);
  if (isAreaChart) {
    minValue = Math.min(0, minValue);
  } else if ((chartType === "Line" || isStackedLine) && chart.valueAxis?.crosses === "autoZero") {
    minValue = Math.min(0, minValue);
  }
  let valueDomain;
  if (isPercentStackedSeries) {
    const explicitMin = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min)
      ? Number(chart.valueAxis.min)
      : Math.min(0, minValue);
    const explicitMax = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max)
      ? Number(chart.valueAxis.max)
      : Math.max(100, maxValue);
    const max = explicitMax <= explicitMin ? explicitMin + 1 : explicitMax;
    valueDomain = {
      min: explicitMin,
      max,
      ticks: buildNumericTickValues(explicitMin, max, chart.valueAxis?.majorUnit ?? 20),
    };
  } else if (resolveAxisLogBase(chart.valueAxis)) {
    const positiveValues = extents.filter((value) => value > 0);
    valueDomain = resolveLogAxisDomain(
      chart.valueAxis,
      positiveValues.length > 0 ? Math.min(...positiveValues) : minValue,
      maxValue,
    );
  } else {
    valueDomain = resolveAxisDomainWithChartOverrides(
      chart.valueAxis,
      minValue,
      maxValue,
      isAreaChart || chart.valueAxis?.crosses === "autoZero",
    );
  }
  const scale = createValueScale(chart.valueAxis, valueDomain.min, valueDomain.max, [plot.top + plot.height, plot.top]);
  const categoryNumericValues = getCategoryNumericValues(chart);
  const date1904 = chart.raw && typeof chart.raw === "object"
    && (chart.raw as Record<string, unknown>).date1904 === true;
  const isDateAxis = chart.categoryAxis?.axisType === "date"
    || (
      isLikelyDateFormatCode(chart.categoryAxis?.numberFormat?.formatCode)
      && categoryNumericValues.length > 0
      && categoryNumericValues.every((value) => value != null)
    );
  const canUseDateScale = isDateAxis
    && categoryNumericValues.length === categories.length
    && categoryNumericValues.every((value) => value != null);
  const dateDomain = canUseDateScale
    ? resolveDateAxisDomain(chart.categoryAxis, categoryNumericValues)
    : null;
  const dateScale = dateDomain
    ? createDateScale(dateDomain.min, dateDomain.max, [plot.left, plot.left + plot.width], date1904)
    : null;
  const dateTicks = dateDomain
    ? buildDateAxisTicks(chart.categoryAxis, dateDomain.min, dateDomain.max, date1904)
    : [];
  const pointScale = canUseDateScale
    ? null
    : scalePoint<string>().domain(categories).range([plot.left, plot.left + plot.width]);
  const categoryPositions = canUseDateScale
    ? categoryNumericValues.map((value) => dateScale?.(value as number) ?? plot.left)
    : categories.map((category) => pointScale?.(category) ?? plot.left);
  const axisCategoryLabels = canUseDateScale ? dateTicks.map((tick) => tick.label) : categories;
  const axisCategoryPositions = canUseDateScale
    ? dateTicks.map((tick) => dateScale?.(tick.serial) ?? plot.left)
    : categoryPositions;
  const categoryAxis = buildCategoryAxisGeometry({
    axis: chart.categoryAxis,
    kind: canUseDateScale ? "date" : "category",
    labels: axisCategoryLabels,
    positions: axisCategoryPositions,
    plot,
    isHorizontal: false,
  });
  const valuePlacement: ValueAxisPlacement = normalizeTickLabelPosition(chart.valueAxis) === "high"
    ? "right"
    : "left";
  const zero = resolveZeroPosition(scale, plot.top + plot.height);
  const valueAxis = buildValueAxisGeometry({
    axis: chart.valueAxis,
    chart,
    ticks: valueDomain.ticks,
    scale,
    plot,
    placement: valuePlacement,
    percentValuesAreScaled: isPercentStackedSeries,
    line: chart.valueAxis?.delete === true ? undefined : valueAxisLine(valuePlacement, plot),
    zeroLine: chart.valueAxis?.delete === true
      ? undefined
      : valueAxisZeroLine(valuePlacement, zero, plot),
  });
  const seriesGeometry: ChartSeriesGeometry[] = chart.series.map((series, seriesIndex) => {
    const points: ChartSeriesPointGeometry[] = [];
    categories.forEach((_, categoryIndex) => {
      const point = stackedPoints[seriesIndex]?.[categoryIndex];
      if (!point?.defined || point.y == null) {
        return;
      }
      const position = {
        x: categoryPositions[categoryIndex] ?? plot.left,
        y: scale(point.y),
      };
      points.push({
        index: categoryIndex,
        categoryIndex,
        value: point.value,
        position,
        rect: markerRect(position, series),
        confidence: "exact",
        rectConfidence: "approximate",
      });
    });
    return {
      index: seriesIndex,
      seriesId: series.id,
      name: seriesName(series, seriesIndex),
      shape: isAreaChart ? "area" : "line",
      categoryAxisIndex: 0,
      valueAxisIndex: 1,
      points,
      confidence: "exact",
    };
  });
  return { axes: [categoryAxis, valueAxis], series: seriesGeometry };
}

function buildComboGeometry(
  chart: XlsxChart,
  layout: ChartLayout,
  limitations: ChartGeometryLimitation[],
): { axes: ChartAxisGeometry[]; series: ChartSeriesGeometry[] } {
  const groups = buildComboGroups(chart);
  const columnGroup = groups.find((group) => group.chartType.startsWith("Column"));
  const lineGroup = groups.find((group) => group.chartType.startsWith("Line") || group.chartType.startsWith("Area"));
  if (!columnGroup || !lineGroup) {
    limitations.push({
      code: "xlsx.combo-shape-provider",
      detail: "The SVG renderer only composes a Column* group with a Line*/Area* group; this combination has no geometric layout.",
      confidence: "provider",
    });
    return { axes: [], series: [] };
  }
  if (columnGroup.is3d || lineGroup.is3d || chart.is3d) {
    limitations.push(THIRD_DIMENSION_LIMITATION);
    return { axes: [], series: [] };
  }
  const isAreaGroup = lineGroup.chartType.startsWith("Area");
  const primaryChart: XlsxChart = {
    ...chart,
    categoryAxis: columnGroup.categoryAxis ?? chart.categoryAxis,
    chartType: columnGroup.chartType,
    is3d: false,
    series: columnGroup.series,
    valueAxis: columnGroup.valueAxis ?? chart.valueAxis,
  };
  const categories = getCategoryLabels(primaryChart);
  if (categories.length === 0) {
    return { axes: [], series: [] };
  }
  const plot = layout.plot;
  const groupGapWidth = typeof columnGroup.raw?.gapWidth === "number" && Number.isFinite(columnGroup.raw.gapWidth)
    ? columnGroup.raw.gapWidth
    : columnGroup.gapWidth ?? chart.gapWidth;
  const histogramColumns = columnGroup.series.some((series) => isHistogramLikeSeries(series));
  const categoryBandPadding = histogramColumns
    ? { inner: 0, outer: 0 }
    : resolveCategoryBandPadding(groupGapWidth);
  const categoryScale = scaleBand<string>()
    .domain(categories)
    .range([plot.left, plot.left + plot.width])
    .paddingInner(categoryBandPadding.inner)
    .paddingOuter(categoryBandPadding.outer);
  const seriesScale = scaleBand<string>()
    .domain(Array.from({ length: columnGroup.series.length }, (_, index) => String(index)))
    .range([0, categoryScale.bandwidth()])
    .paddingInner(histogramColumns ? 0 : 0.16)
    .paddingOuter(histogramColumns ? 0 : 0.08);
  const categoryPositions = categories.map((category) => (
    (categoryScale(category) ?? plot.left) + categoryScale.bandwidth() / 2
  ));
  const resolveGlobalSeriesIndex = (series: XlsxChartSeries) => resolveComboSeriesIndex(chart, series);
  const primaryValues = columnGroup.series.flatMap((series) => (
    series.values.map((value) => safeNumber(value)).filter((value): value is number => value != null)
  ));
  const secondaryValues = lineGroup.series.flatMap((series) => (
    series.values.map((value) => safeNumber(value)).filter((value): value is number => value != null)
  ));
  if (primaryValues.length === 0 || secondaryValues.length === 0) {
    return { axes: [], series: [] };
  }
  const columnValueAxis = columnGroup.valueAxis ?? chart.valueAxis ?? null;
  const lineValueAxis = lineGroup.valueAxis ?? chart.valueAxis ?? null;
  const sharedValueAxis = columnValueAxis != null && lineValueAxis != null
    ? (columnValueAxis.id != null && lineValueAxis.id != null
        ? columnValueAxis.id === lineValueAxis.id
        : columnValueAxis === lineValueAxis)
    : columnValueAxis === lineValueAxis;
  const resolveComboDomain = (
    axis: typeof columnValueAxis,
    values: number[],
    includeZero: boolean,
  ) => {
    if (resolveAxisLogBase(axis)) {
      const positiveValues = values.filter((value) => value > 0);
      return resolveLogAxisDomain(
        axis,
        positiveValues.length > 0 ? Math.min(...positiveValues) : 1,
        positiveValues.length > 0 ? Math.max(...positiveValues) : 10,
      );
    }
    return resolveAxisDomainWithChartOverrides(
      axis,
      typeof axis?.min === "number" ? axis.min : Math.min(...values),
      typeof axis?.max === "number" ? axis.max : Math.max(...values),
      includeZero,
    );
  };
  const primaryDomain = resolveComboDomain(
    columnValueAxis,
    sharedValueAxis ? [...primaryValues, ...secondaryValues] : primaryValues,
    true,
  );
  const secondaryDomain = sharedValueAxis
    ? primaryDomain
    : resolveComboDomain(lineValueAxis, secondaryValues, false);
  const primaryScale = createValueScale(columnValueAxis, primaryDomain.min, primaryDomain.max, [plot.top + plot.height, plot.top]);
  const secondaryScale = sharedValueAxis
    ? primaryScale
    : createValueScale(lineValueAxis, secondaryDomain.min, secondaryDomain.max, [plot.top + plot.height, plot.top]);
  const primaryPlacement: ValueAxisPlacement = normalizeTickLabelPosition(columnValueAxis) === "high" ? "right" : "left";
  const categoryAxis = buildCategoryAxisGeometry({
    axis: columnGroup.categoryAxis ?? chart.categoryAxis,
    kind: (columnGroup.categoryAxis ?? chart.categoryAxis)?.axisType === "date" ? "date" : "category",
    labels: categories,
    positions: categoryPositions,
    plot,
    isHorizontal: false,
  });
  const primaryZero = resolveZeroPosition(primaryScale, plot.top + plot.height);
  const primaryValueAxis = buildValueAxisGeometry({
    axis: columnValueAxis,
    chart,
    ticks: primaryDomain.ticks,
    scale: primaryScale,
    plot,
    placement: primaryPlacement,
    percentValuesAreScaled: columnGroup.chartType.includes("PercentStacked"),
    line: columnValueAxis?.delete === true ? undefined : valueAxisLine(primaryPlacement, plot),
    zeroLine: columnValueAxis?.delete === true
      ? undefined
      : valueAxisZeroLine(primaryPlacement, primaryZero, plot),
  });
  const secondaryValueAxis = sharedValueAxis
    ? undefined
    : buildValueAxisGeometry({
        axis: lineValueAxis,
        chart,
        ticks: secondaryDomain.ticks,
        scale: secondaryScale,
        plot,
        placement: "right",
        percentValuesAreScaled: lineGroup.chartType.includes("PercentStacked"),
        line: valueAxisLine("right", plot),
      });
  const axes = secondaryValueAxis
    ? [categoryAxis, primaryValueAxis, secondaryValueAxis]
    : [categoryAxis, primaryValueAxis];
  const comboValues = [...primaryValues, ...secondaryValues];
  if (comboValues.some((value) => value < 0)) {
    limitations.push({
      code: "xlsx.combo-negative-zero-clamped",
      detail: "The SVG combo renderer clamps negative column values to the zero baseline; the provider reports the painted (clamped) rectangles.",
      confidence: "approximate",
    });
  }
  const seriesByIndex = new Map<number, ChartSeriesGeometry>();
  const ensureSeries = (globalIndex: number, series: XlsxChartSeries, shape: ChartSeriesGeometry["shape"], valueAxisIndex: number) => {
    const existing = seriesByIndex.get(globalIndex);
    if (existing) {
      return existing;
    }
    const created: ChartSeriesGeometry = {
      index: globalIndex,
      seriesId: series.id,
      name: seriesName(series, globalIndex),
      shape,
      categoryAxisIndex: 0,
      valueAxisIndex,
      points: [],
      confidence: "exact",
    };
    seriesByIndex.set(globalIndex, created);
    return created;
  };
  columnGroup.series.forEach((series, seriesIndex) => {
    const globalIndex = resolveGlobalSeriesIndex(series);
    const geometry = ensureSeries(globalIndex, series, "bar", 1);
    const points: ChartSeriesPointGeometry[] = [];
    categories.forEach((category, categoryIndex) => {
      const value = safeNumber(series.values[categoryIndex]) ?? 0;
      const categoryStart = categoryScale(category) ?? plot.left;
      const barWidth = Math.max(1, histogramColumns ? categoryScale.bandwidth() : seriesScale.bandwidth());
      const x = categoryStart + (histogramColumns ? 0 : (seriesScale(String(seriesIndex)) ?? 0));
      const logBaseline = resolveAxisLogBase(columnValueAxis) ? primaryDomain.min : 0;
      const scaledValue = Math.max(logBaseline, value);
      const scaledBaseline = Math.max(logBaseline, 0);
      const y = primaryScale(scaledValue);
      const zeroY = primaryScale(scaledBaseline);
      const height = Math.max(1, Math.abs(zeroY - primaryScale(scaledValue)));
      const rect = makeRect(x, Math.min(y, zeroY), barWidth, height);
      points.push({
        index: categoryIndex,
        categoryIndex,
        value,
        position: pointOfRect(rect),
        rect,
        confidence: "exact",
        rectConfidence: "exact",
      });
    });
    geometry.points = [...geometry.points, ...points];
  });
  lineGroup.series.forEach((series) => {
    const globalIndex = resolveGlobalSeriesIndex(series);
    const geometry = ensureSeries(globalIndex, series, isAreaGroup ? "area" : "line", secondaryValueAxis ? 2 : 1);
    const points: ChartSeriesPointGeometry[] = [];
    categories.forEach((category, categoryIndex) => {
      const value = safeNumber(series.values[categoryIndex]);
      if (value == null) {
        return;
      }
      const position = {
        x: (categoryScale(category) ?? plot.left) + categoryScale.bandwidth() / 2,
        y: secondaryScale(value),
      };
      points.push({
        index: categoryIndex,
        categoryIndex,
        value,
        position,
        rect: markerRect(position, series),
        confidence: "exact",
        rectConfidence: "approximate",
      });
    });
    geometry.points = [...geometry.points, ...points];
  });
  const series = [...seriesByIndex.values()].sort((left, right) => left.index - right.index);
  return { axes, series };
}

function buildAxesAndSeries(
  chart: XlsxChart,
  chartType: string,
  layout: ChartLayout,
  limitations: ChartGeometryLimitation[],
): { axes: ChartAxisGeometry[]; series: ChartSeriesGeometry[] } {
  if (isComboChart(chart)) {
    return buildComboGeometry(chart, layout, limitations);
  }
  if (/^(Column|Bar)(Clustered|Stacked|PercentStacked)$/.test(chartType)) {
    return buildBarGeometry(chart, chartType, layout, limitations);
  }
  if (/^(Line|Area)(Stacked|PercentStacked)?$/.test(chartType)) {
    return buildLineGeometry(chart, chartType, layout, limitations);
  }
  limitations.push({
    code: `xlsx.${chartType.replace(/[^A-Za-z0-9]+/g, "-").toLowerCase()}-provider`,
    detail: `The XLSX geometry provider does not expose internal geometry for ${chartType} charts yet; layout and legend geometry are still available.`,
    confidence: "provider",
  });
  return { axes: [], series: [] };
}

/**
 * Compute chart-internal geometry from the XLSX SVG renderer's own layout
 * rules. Coordinates are chart-local CSS pixels. `exact` means "the renderer
 * paints this exact rectangle/anchor"; text boxes carry `approximate` because
 * SVG text metrics are estimated. It never claims Microsoft Excel pixel parity.
 */
export function computeXlsxChartGeometry(
  chart: XlsxChart,
  options: XlsxChartGeometryOptions,
): ChartGeometrySnapshot {
  if (!chart || typeof chart !== "object") {
    throw new TypeError("computeXlsxChartGeometry requires an XlsxChart");
  }
  if (!Number.isFinite(options.frame?.width) || options.frame.width <= 0
    || !Number.isFinite(options.frame?.height) || options.frame.height <= 0) {
    throw new RangeError("chart geometry frame must be positive finite pixels");
  }
  const frame: ChartGeometryFrame = {
    width: Math.max(80, Math.round(options.frame.width)),
    height: Math.max(60, Math.round(options.frame.height)),
  };
  const palette = options.palette ?? DEFAULT_GEOMETRY_PALETTE;
  const chartType = normalizeRenderableChartType(chart);
  const legendItems = getLegendItems(chart, chartType, palette, options.labels);
  const layout = buildLayout(chart, { ...frame, left: 0, top: 0 }, legendItems);
  const limitations: ChartGeometryLimitation[] = [
    TEXT_METRICS_LIMITATION,
    MANUAL_LAYOUT_LIMITATION,
  ];
  const legend = buildLegendGeometry(chart, chartType, layout, legendItems);
  const { axes, series } = buildAxesAndSeries(chart, chartType, layout, limitations);
  const titleLabel: ChartGeometryLabel | undefined = chart.title
    ? {
        text: chart.title,
        anchor: { x: layout.width / 2, y: layout.titleHeight >= 30 ? 19 : 16 },
        rect: makeRect(
          Math.max(6, layout.width / 2 - estimateSvgTextWidth(chart.title, 12) / 2),
          (layout.titleHeight >= 30 ? 19 : 16) - 12,
          Math.min(estimateSvgTextWidth(chart.title, 12), Math.max(40, layout.width - 12)),
          14,
        ),
        confidence: "approximate",
      }
    : undefined;
  const exact = !limitations.some((limitation) => limitation.confidence === "provider")
    && axes.every((axis) => axis.confidence === "exact")
    && series.every((entry) => entry.confidence === "exact")
    && layout.plot.width > 0
    && layout.plot.height > 0;
  return {
    chartId: chart.id,
    provider: XLSX_CHART_GEOMETRY_PROVIDER_ID,
    format: "xlsx",
    space: "chart-local-px",
    frame,
    layout: {
      chartArea: makeRect(0, 0, layout.width, layout.height),
      plotArea: makeRect(layout.plot.left, layout.plot.top, layout.plot.width, layout.plot.height),
      ...(chart.title ? { title: makeRect(0, 0, layout.width, layout.titleHeight) } : {}),
      ...(titleLabel ? { titleLabel } : {}),
      ...(legend ? { legend } : {}),
      confidence: "exact",
    },
    axes,
    series,
    limitations,
    exact,
  };
}

export interface XlsxChartGeometryProviderOptions {
  palette?: ChartRendererPalette;
  labels?: ChartLabels;
}

/**
 * Create a framework-agnostic provider bound to the XLSX SVG renderer.
 *
 * The provider rejects requests for a different `chartId` instead of returning
 * geometry for the wrong chart; that keeps Surface caches safe when a workbook
 * revision swaps charts under the same frame.
 */
export function createXlsxChartGeometryProvider(
  options: XlsxChartGeometryProviderOptions = {},
): ChartGeometryProvider<XlsxChart> {
  return {
    id: XLSX_CHART_GEOMETRY_PROVIDER_ID,
    format: "xlsx",
    getChartGeometry(chart: XlsxChart, request: ChartGeometryRequest) {
      if (request.chartId && request.chartId !== chart.id) {
        return undefined;
      }
      const snapshot = computeXlsxChartGeometry(chart, {
        frame: request.frame,
        palette: options.palette,
        labels: options.labels,
      });
      if (request.detail === "layout") {
        return {
          ...snapshot,
          axes: [],
          series: [],
          exact: false,
        };
      }
      return snapshot;
    },
  };
}
