/** @jsxImportSource vue */
import type { XlsxChart, XlsxImageRect } from "@arcships/xlsx-core";
import {
  buildComboGroups,
  buildHierarchyData,
  buildPieEntries,
  buildRegionMapLegendItems,
  buildStockLegendItems,
  buildSurfaceLegendItems,
  chartElementDataProps,
  chartPointColor,
  chartSeriesColor,
  clamp,
  createChartAxisTickFormatter,
  estimateSvgTextWidth,
  getCategoryLabels,
  getComboLegendSeries,
  isComboChart,
  normalizeCategoryLabel,
  normalizeLegendPosition,
  resolveChartAxisTextColor,
  resolveComboSecondaryValueAxis,
  safeNumber,
  wrapSvgText,
} from "./chart-renderer";
import type { ChartLayout, ChartLegendMetrics, ChartRendererPalette, LegendItem } from "./chart-renderer";
import type { ChartLabels } from "./chart-types";
import { resolveChartSeriesLabel } from "./chart-shared";

const LEGEND_FONT_SIZE = 10;
const LEGEND_SWATCH_SIZE = 8;
const LEGEND_TEXT_OFFSET = 12;
const LEGEND_LINE_HEIGHT = 12;
const LEGEND_ITEM_GAP = 6;
const LEGEND_PADDING = 8;

export function getLegendItems(chart: XlsxChart, chartType: string, palette: ChartRendererPalette, labels?: ChartLabels): LegendItem[] {
  if (!chart.legend) {
    return [];
  }
  if (isComboChart(chart)) {
    return getComboLegendSeries(chart, labels);
  }
  if (chartType === "Stock") {
    return buildStockLegendItems(chart, palette, labels);
  }
  if (chartType === "Surface") {
    return buildSurfaceLegendItems(chart, palette);
  }
  if (chartType === "RegionMap") {
    return buildRegionMapLegendItems(chart);
  }
  if (chartType === "Sunburst" || chartType === "Treemap") {
    const hierarchyData = buildHierarchyData(chart);
    if (!hierarchyData?.children) {
      return [];
    }
    return hierarchyData.children.map((child, index) => ({
      color: chartSeriesColor(chart, child.colorIndex ?? index),
      label: child.name
    }));
  }
  if (chartType === "Pie" || chartType === "Pie3D" || chartType === "PieExploded" || chartType === "Doughnut" || chartType === "BarOfPie") {
    if (chartType === "BarOfPie") {
      const categories = getCategoryLabels(chart);
      const values = chart.series[0]?.values ?? [];
      return categories
        .map((label, index) => ({
          color: chartPointColor(chart, index),
          label: normalizeCategoryLabel(label),
          value: safeNumber(values[index]) ?? 0
        }))
        .filter((entry) => entry.value > 0)
        .map((entry) => ({
          color: entry.color,
          label: entry.label
        }));
    }
    return buildPieEntries(chart).map((entry) => ({
      color: entry.color,
      label: entry.label
    }));
  }
  const isXyLegend = (
    chartType === "Scatter"
    || chartType === "ScatterLines"
    || chartType === "ScatterSmooth"
    || chartType === "Bubble"
  );
  return chart.series.map((series, index) => ({
    color: isXyLegend
      ? (series.lineColor ?? series.markerColor ?? series.color ?? chartSeriesColor(chart, index))
      : chartSeriesColor(chart, index),
    label: resolveChartSeriesLabel(series, index, labels)
  }));
}

function usesCartesianAxes(chart: XlsxChart) {
  if (isComboChart(chart)) {
    return true;
  }
  const chartType = chart.chartType;
  if (chartType === "BarOfPie") {
    return false;
  }
  return /^(Column|Bar|Line|Area|Scatter|Bubble|Stock|Waterfall)/.test(chartType);
}

function isHorizontalBarChart(chart: XlsxChart) {
  return /^Bar(?!OfPie)/.test(chart.chartType);
}

function estimateValueAxisLabelWidth(chart: XlsxChart, width: number) {
  const valueAxis = chart.valueAxis;
  if (!valueAxis || valueAxis.labelPosition === "none" || valueAxis.delete === true) {
    return 0;
  }
  const formatter = createChartAxisTickFormatter(valueAxis, chart, {
    percentValuesAreScaled: /PercentStacked/.test(chart.chartType)
  });
  const candidates: number[] = [0];
  if (typeof valueAxis.min === "number" && Number.isFinite(valueAxis.min)) {
    candidates.push(valueAxis.min);
  }
  if (typeof valueAxis.max === "number" && Number.isFinite(valueAxis.max)) {
    candidates.push(valueAxis.max);
  }
  for (const series of chart.series) {
    for (const rawValue of series.values) {
      const value = safeNumber(rawValue);
      if (value != null) {
        candidates.push(value);
      }
      if (candidates.length >= 200) {
        break;
      }
    }
    if (candidates.length >= 200) {
      break;
    }
  }
  const maxLabelWidth = candidates.reduce(
    (widest, value) => Math.max(widest, estimateSvgTextWidth(formatter(value), LEGEND_FONT_SIZE)),
    0
  );
  return Math.min(maxLabelWidth + 12, Math.max(40, width * 0.36));
}

function estimateSecondaryAxisLabelWidth(chart: XlsxChart, width: number) {
  const axis = resolveComboSecondaryValueAxis(chart);
  if (!axis || axis.labelPosition === "none" || axis.delete === true) {
    return 0;
  }
  const formatter = createChartAxisTickFormatter(axis, chart);
  const candidates: number[] = [0];
  if (typeof axis.min === "number" && Number.isFinite(axis.min)) {
    candidates.push(axis.min);
  }
  if (typeof axis.max === "number" && Number.isFinite(axis.max)) {
    candidates.push(axis.max);
  }
  const group = buildComboGroups(chart).find((entry) => entry.valueAxis === axis);
  for (const series of group?.series ?? chart.series) {
    for (const rawValue of series.values) {
      const value = safeNumber(rawValue);
      if (value != null) {
        candidates.push(value);
      }
      if (candidates.length >= 200) {
        break;
      }
    }
    if (candidates.length >= 200) {
      break;
    }
  }
  const maxLabelWidth = candidates.reduce(
    (widest, value) => Math.max(widest, estimateSvgTextWidth(formatter(value), LEGEND_FONT_SIZE)),
    0
  );
  const titleWidth = typeof axis.title === "string" && axis.title.trim().length > 0 ? 14 : 0;
  return Math.min(maxLabelWidth + 18 + titleWidth, Math.max(40, width * 0.3));
}

function estimateCategoryAxisLabelWidth(chart: XlsxChart, width: number) {
  const categoryAxis = chart.categoryAxis;
  if (!categoryAxis || categoryAxis.labelPosition === "none" || categoryAxis.delete === true) {
    return 0;
  }
  const maxLabelWidth = getCategoryLabels(chart).reduce(
    (widest, label) => Math.max(widest, estimateSvgTextWidth(label, LEGEND_FONT_SIZE)),
    0
  );
  return Math.min(maxLabelWidth + 12, Math.max(40, width * 0.42));
}

function estimateCategoryLabelBottomReserve(chart: XlsxChart, width: number) {
  const categories = getCategoryLabels(chart);
  if (categories.length === 0) {
    return 28;
  }
  const maxLabelWidth = categories.reduce(
    (widest, label) => Math.max(widest, estimateSvgTextWidth(label, LEGEND_FONT_SIZE)),
    0
  );
  const plotWidth = Math.max(80, width - 60);
  const visibleCount = Math.max(
    1,
    Math.ceil(categories.length / Math.max(1, Math.floor(chart.categoryAxis?.tickLabelSkip ?? 1)))
  );
  const slotWidth = plotWidth / visibleCount;
  if (maxLabelWidth <= slotWidth * 0.92) {
    return 28;
  }
  return Math.max(28, Math.min(104, 18 + maxLabelWidth * Math.SQRT1_2));
}

export function buildLegendMetrics(
  items: LegendItem[],
  legendPosition: string | undefined,
  width: number,
  height: number
): ChartLegendMetrics | undefined {
  if (items.length === 0) {
    return undefined;
  }
  const horizontal = legendPosition !== "left" && legendPosition !== "right";
  if (horizontal) {
    const maxTextWidth = Math.max(40, Math.min(140, width * 0.45));
    const entryLines = items.map((item) => wrapSvgText(item.label, maxTextWidth, LEGEND_FONT_SIZE));
    const entryWidths = entryLines.map((lines) => (
      LEGEND_TEXT_OFFSET + Math.max(...lines.map((line) => estimateSvgTextWidth(line, LEGEND_FONT_SIZE)))
    ));
    const rows: number[][] = [];
    const rowHeights: number[] = [];
    let currentRow: number[] = [];
    let currentWidth = 0;
    entryWidths.forEach((entryWidth, index) => {
      const itemWidth = entryWidth + LEGEND_ITEM_GAP + 8;
      if (currentRow.length > 0 && currentWidth + itemWidth > Math.max(60, width - 16)) {
        rows.push(currentRow);
        rowHeights.push(
          Math.max(...currentRow.map((rowIndex) => entryLines[rowIndex]?.length ?? 1)) * LEGEND_LINE_HEIGHT + 2
        );
        currentRow = [];
        currentWidth = 0;
      }
      currentRow.push(index);
      currentWidth += itemWidth;
    });
    if (currentRow.length > 0) {
      rows.push(currentRow);
      rowHeights.push(
        Math.max(...currentRow.map((rowIndex) => entryLines[rowIndex]?.length ?? 1)) * LEGEND_LINE_HEIGHT + 2
      );
    }
    return {
      entryLines,
      height: Math.min(height, rowHeights.reduce((sum, rowHeight) => sum + rowHeight, 0) + LEGEND_PADDING),
      maxTextWidth,
      rowHeights,
      rows,
      width: Math.max(60, Math.min(width - 16, Math.max(...entryWidths.map((entryWidth) => entryWidth + 8))))
    };
  }

  const maxTextWidth = Math.max(48, Math.min(160, width * 0.4));
  const entryLines = items.map((item) => wrapSvgText(item.label, maxTextWidth, LEGEND_FONT_SIZE));
  const rowHeights = entryLines.map((lines) => (
    Math.max(1, lines.length) * LEGEND_LINE_HEIGHT + LEGEND_ITEM_GAP
  ));
  const contentWidth = Math.max(
    0,
    ...entryLines.map((lines) => (
      LEGEND_TEXT_OFFSET + Math.max(...lines.map((line) => estimateSvgTextWidth(line, LEGEND_FONT_SIZE)))
    ))
  );
  return {
    entryLines,
    height: rowHeights.reduce((sum, rowHeight) => sum + rowHeight, 0) + LEGEND_PADDING,
    maxTextWidth,
    rowHeights,
    rows: items.map((_, index) => [index]),
    width: clamp(contentWidth + LEGEND_PADDING * 2, 60, width * 0.5)
  };
}

export function buildLayout(chart: XlsxChart, rect: XlsxImageRect, legendItems: LegendItem[]): ChartLayout {
  const width = Math.max(80, Math.round(rect.width));
  const height = Math.max(60, Math.round(rect.height));
  const isSurfaceChart = chart.chartType === "Surface"
    || (chart.raw && typeof chart.raw === "object" && typeof (chart.raw as Record<string, unknown>).xmlChartType === "string" && String((chart.raw as Record<string, unknown>).xmlChartType).includes("surface"));
  const titleHeight = chart.title ? (isSurfaceChart ? 30 : 24) : 8;
  const legendPosition = normalizeLegendPosition(chart.legend?.position);
  const legendVertical = legendItems.length > 0 && (legendPosition === "right" || legendPosition === "left");
  const compactWidth = width <= 280;
  const compactHeight = height <= 200;
  const finalLegendPosition = compactWidth || compactHeight
    ? (legendVertical ? "bottom" : legendPosition)
    : legendPosition;
  const legendMetrics = buildLegendMetrics(legendItems, finalLegendPosition, width, height);
  const legendWidth = legendMetrics?.width ?? 0;
  const legendHeight = legendMetrics?.height ?? 0;

  const cartesian = usesCartesianAxes(chart);
  const isHorizontalBar = isHorizontalBarChart(chart);
  const valueLabelWidth = cartesian ? estimateValueAxisLabelWidth(chart, width) : 0;
  const secondaryAxisWidth = cartesian ? estimateSecondaryAxisLabelWidth(chart, width) : 0;
  const categoryLabelWidth = cartesian && isHorizontalBar ? estimateCategoryAxisLabelWidth(chart, width) : 0;
  const baseLeft = Math.max(42, isHorizontalBar ? categoryLabelWidth + 12 : valueLabelWidth + 12);
  const baseBottom = cartesian && !isHorizontalBar
    ? estimateCategoryLabelBottomReserve(chart, width)
    : 28;

  const plotLeft = baseLeft + (finalLegendPosition === "left" ? legendWidth : 0);
  const plotRight = 16 + secondaryAxisWidth + (finalLegendPosition === "right" ? legendWidth : 0);
  const plotTop = titleHeight + (finalLegendPosition === "top" ? legendHeight + 4 : 0);
  const plotBottom = baseBottom + (finalLegendPosition === "bottom" ? legendHeight + 4 : 0);

  return {
    height,
    legendItems,
    legendMetrics,
    legendPosition: finalLegendPosition,
    plot: {
      height: Math.max(40, height - plotTop - plotBottom),
      left: plotLeft,
      top: plotTop,
      width: Math.max(40, width - plotLeft - plotRight)
    },
    secondaryAxisWidth,
    titleHeight,
    width
  };
}

export function renderLegend(chart: XlsxChart, layout: ChartLayout, palette: ChartRendererPalette) {
  if (!chart.legend || layout.legendItems.length === 0) {
    return null;
  }
  const textColor = resolveChartAxisTextColor(chart);
  const legendPos = layout.legendPosition;
  const items = layout.legendItems;
  const metrics = layout.legendMetrics ?? buildLegendMetrics(items, legendPos, layout.width, layout.height);
  if (!metrics) {
    return null;
  }
  const entryWidth = (index: number) => {
    const lines = metrics.entryLines[index] ?? [items[index]?.label ?? ""];
    return LEGEND_TEXT_OFFSET + Math.max(...lines.map((line) => estimateSvgTextWidth(line, LEGEND_FONT_SIZE))) + 8;
  };

  if (legendPos === "left" || legendPos === "right") {
    const x = legendPos === "right"
      ? layout.plot.left + layout.plot.width + 8 + (layout.secondaryAxisWidth ?? 0)
      : 8;
    let cursorY = layout.plot.top + 6;
    return (
      <g>
        {items.map((item, index) => {
          const lines = metrics.entryLines[index] ?? [item.label];
          const rowHeight = metrics.rowHeights[index] ?? (lines.length * LEGEND_LINE_HEIGHT + LEGEND_ITEM_GAP);
          const y = cursorY;
          cursorY += rowHeight;
          return (
              <g
                {...(index < chart.series.length ? chartElementDataProps(index) : {})}
                data-xlsx-chart-element-kind={index < chart.series.length ? "legendEntry" : undefined}
                key={`legend-${index}`}
                transform={`translate(${x}, ${y})`}
              >
              <rect fill={item.color} height={LEGEND_SWATCH_SIZE} rx={1.2} ry={1.2} width={LEGEND_SWATCH_SIZE} x={0} y={-7} />
              {lines.map((line, lineIndex) => (
                <text
                  fill={textColor}
                  font-size={LEGEND_FONT_SIZE}
                  key={`legend-label-${index}-${lineIndex}`}
                  x={LEGEND_TEXT_OFFSET}
                  y={lineIndex * LEGEND_LINE_HEIGHT}
                >
                  {line}
                </text>
              ))}
            </g>
          );
        })}
      </g>
    );
  }

  const startY = legendPos === "top"
    ? layout.titleHeight + LEGEND_PADDING
    : Math.max(layout.plot.top + layout.plot.height + 6, layout.height - metrics.height + LEGEND_PADDING);
  let rowTop = startY;
  return (
    <g>
      {metrics.rows.map((row, rowIndex) => {
        const rowHeight = metrics.rowHeights[rowIndex] ?? LEGEND_LINE_HEIGHT + 2;
        const totalWidth = row.reduce((sum, index) => sum + entryWidth(index) + LEGEND_ITEM_GAP, 0);
        let cursorX = Math.max(8, (layout.width - totalWidth) / 2);
        const node = (
          <g key={`legend-row-${rowIndex}`}>
            {row.map((index) => {
              const item = items[index];
              if (!item) {
                return null;
              }
              const lines = metrics.entryLines[index] ?? [item.label];
              const x = cursorX;
              cursorX += entryWidth(index) + LEGEND_ITEM_GAP;
              return (
                <g
                  {...(index < chart.series.length ? chartElementDataProps(index) : {})}
                  data-xlsx-chart-element-kind={index < chart.series.length ? "legendEntry" : undefined}
                  key={`legend-${index}`}
                  transform={`translate(${x}, ${rowTop + 9})`}
                >
                  <rect fill={item.color} height={LEGEND_SWATCH_SIZE} rx={1.2} ry={1.2} width={LEGEND_SWATCH_SIZE} x={0} y={-7} />
                  {lines.map((line, lineIndex) => (
                    <text
                      fill={textColor}
                      font-size={LEGEND_FONT_SIZE}
                      key={`legend-label-${index}-${lineIndex}`}
                      x={LEGEND_TEXT_OFFSET}
                      y={lineIndex * LEGEND_LINE_HEIGHT}
                    >
                      {line}
                    </text>
                  ))}
                </g>
              );
            })}
          </g>
        );
        rowTop += rowHeight;
        return node;
      })}
    </g>
  );
}
