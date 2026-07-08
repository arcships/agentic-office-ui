import type { Workbook } from "@dukelib/sheets-wasm";
import type {
  XlsxChart,
  XlsxChartAxis,
  XlsxChartDataLabels,
  XlsxChartLegend,
  XlsxChartSeries,
  XlsxChartTypeGroup,
  XlsxChartWall,
  XlsxChartsheet,
  XlsxImageAnchor,
  XlsxWorkbookTab,
  XlsxThemePalette,
} from "../types";

import {
  normalizeHexColor,
  EMU_PER_PIXEL,
} from "./chart-colors";

import {
  resolveReferenceValues,
  resolveChartReferenceLabel,
  resolveSeriesName,
  buildA1RangeFormula,
  resolveReferenceSheet,
  normalizeChartReference,
  resolveReferenceRowPaths,
} from "./chart-series";

import {
  applyBuiltinChartDefaults,
} from "./chart-styles";

import {
  normalizeChartAnchor,
  normalizeChartDataLabels,
} from "./chart-export";

import {
  getLocalChildren,
  normalizeLegendPosition,
} from "./chart-parser";

export const PRIMARY_CHART_TYPE_LOCAL_NAMES = [
  "barChart",
  "lineChart",
  "line3DChart",
  "stockChart",
  "radarChart",
  "scatterChart",
  "pieChart",
  "pie3DChart",
  "doughnutChart",
  "areaChart",
  "area3DChart",
  "bar3DChart",
  "ofPieChart",
  "bubbleChart",
  "surfaceChart",
  "surface3DChart"
] as const;

export type WorkbookChartOrigin = {
  anchorIndex: number;
  anchor: XlsxImageAnchor | null;
  chartPath: string | null;
  chartKind: "classic" | "modern";
  drawingPath: string;
  workbookSheetIndex: number;
};

export type WorkbookChartAssets = {
  chartOriginsById: Map<string, WorkbookChartOrigin>;
  chartsByWorkbookSheetIndex: XlsxChart[][];
  chartsheets: XlsxChartsheet[];
  tabs: XlsxWorkbookTab[];
};

export type ChartStyleAppearance = {
  axisLabelColor?: string;
  axisLineColor?: string;
  chartAreaBorderColor?: string;
  chartAreaFillColor?: string;
  chartAreaNoFill?: boolean;
  paletteOffset?: number;
  textColor?: string;
  titleColor?: string;
};

export function findPrimaryChartTypeNode(plotAreaNode: Element | null) {
  if (!plotAreaNode) {
    return null;
  }

  for (const localName of PRIMARY_CHART_TYPE_LOCAL_NAMES) {
    const node = getLocalChildren(plotAreaNode, localName)[0];
    if (node) {
      return node;
    }
  }

  return null;
}

export function resolveScatterChartType(scatterStyle: string | null | undefined) {
  switch (scatterStyle) {
    case "line":
    case "lineMarker":
      return "ScatterLines";
    case "smooth":
    case "smoothMarker":
      return "ScatterSmooth";
    default:
      return "Scatter";
  }
}

function normalizeChartExLegend(raw: unknown): XlsxChartLegend | null {
  if (!raw || typeof raw !== "object") {
    return null;
  }

  const legend = raw as Record<string, unknown>;
  const position = typeof legend.pos === "string"
    ? normalizeLegendPosition(String(legend.pos))
    : undefined;
  return {
    overlay: typeof legend.overlay === "boolean" ? legend.overlay : undefined,
    position,
    raw: legend
  };
}

function humanizeChartExLayoutLabel(layout: string | undefined) {
  if (!layout) {
    return undefined;
  }
  return layout
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/[-_]+/g, " ")
    .trim()
    .replace(/\b\w/g, (match) => match.toUpperCase());
}

function normalizeChartExAxis(raw: unknown): XlsxChartAxis | null {
  if (!raw || typeof raw !== "object") {
    return null;
  }

  const axis = raw as Record<string, unknown>;
  const scaling = axis.scaling && typeof axis.scaling === "object"
    ? axis.scaling as Record<string, unknown>
    : null;
  const numberFormat = axis.numberFormat && typeof axis.numberFormat === "object"
    ? axis.numberFormat as Record<string, unknown>
    : null;

  return {
    delete: typeof axis.hidden === "boolean" ? axis.hidden : undefined,
    id: typeof axis.id === "number" && Number.isFinite(axis.id) ? axis.id : undefined,
    crossId: typeof axis.crossId === "number" && Number.isFinite(axis.crossId) ? axis.crossId : undefined,
    majorGridlines: axis.majorGridlines != null ? true : undefined,
    majorUnit: typeof scaling?.majorUnit === "number" ? scaling.majorUnit : undefined,
    max: typeof scaling?.max === "number" ? scaling.max : undefined,
    min: typeof scaling?.min === "number" ? scaling.min : undefined,
    minorGridlines: axis.minorGridlines != null ? true : undefined,
    minorUnit: typeof scaling?.minorUnit === "number" ? scaling.minorUnit : undefined,
    numberFormat: numberFormat
      ? {
          formatCode: typeof numberFormat.formatCode === "string" ? numberFormat.formatCode : undefined,
          sourceLinked: typeof numberFormat.sourceLinked === "boolean" ? numberFormat.sourceLinked : undefined
        }
      : undefined,
    raw: axis,
    position: typeof axis.position === "string" ? axis.position : undefined,
    tickLabelSkip: typeof axis.tickLabelSkip === "number" ? axis.tickLabelSkip : undefined,
    tickMarkSkip: typeof axis.tickMarkSkip === "number" ? axis.tickMarkSkip : undefined
  };
}

export function resolveChartExLayoutChartType(layout: string | undefined) {
  switch (layout) {
    case "boxWhisker":
      return "BoxWhisker";
    case "clusteredColumn":
      return "ColumnClustered";
    case "funnel":
      return "Funnel";
    case "paretoLine":
      return "Line";
    case "regionMap":
      return "RegionMap";
    case "sunburst":
      return "Sunburst";
    case "treemap":
      return "Treemap";
    case "waterfall":
      return "Waterfall";
    default:
      return layout ? `Unsupported(cx:${layout})` : "ColumnClustered";
  }
}

function resolveChartExSeriesLayout(raw: unknown) {
  if (!raw || typeof raw !== "object") {
    return undefined;
  }
  const record = raw as Record<string, unknown>;
  return typeof record.layout === "string"
    ? record.layout
    : typeof record.layoutId === "string"
      ? record.layoutId
      : undefined;
}

function resolveChartExSeriesAxisIds(raw: unknown) {
  if (!raw || typeof raw !== "object") {
    return [];
  }
  const record = raw as Record<string, unknown>;
  if (Array.isArray(record.axisIds)) {
    return record.axisIds.filter((value): value is number => typeof value === "number" && Number.isFinite(value));
  }
  if (Array.isArray(record.axisId)) {
    return record.axisId.flatMap((value) => {
      if (typeof value === "number" && Number.isFinite(value)) {
        return [value];
      }
      if (value && typeof value === "object" && typeof (value as { val?: unknown }).val === "number") {
        return [(value as { val: number }).val];
      }
      return [];
    });
  }
  if (typeof record.axisId === "number" && Number.isFinite(record.axisId)) {
    return [record.axisId];
  }
  return [];
}

function niceHistogramStep(value: number) {
  if (!Number.isFinite(value) || value <= 0) {
    return 1;
  }
  const exponent = Math.floor(Math.log10(value));
  const scale = 10 ** exponent;
  const normalized = value / scale;
  if (normalized <= 1) {
    return scale;
  }
  if (normalized <= 2) {
    return scale * 2;
  }
  if (normalized <= 5) {
    return scale * 5;
  }
  return scale * 10;
}

function formatHistogramBinLabel(lower: number, upper: number, index: number, closedRight: boolean) {
  const leftBracket = closedRight
    ? (index === 0 ? "[" : "(")
    : "[";
  const rightBracket = closedRight ? "]" : ")";
  return `${leftBracket}${Number(lower.toFixed(6))},${Number(upper.toFixed(6))}${rightBracket}`;
}

type ChartExHistogramBin = {
  count: number;
  label: string;
  lower: number;
  upper: number;
};

function buildChartExHistogramBins(values: number[], rawSeries: unknown, sortByFrequency: boolean) {
  if (values.length === 0) {
    return [] as ChartExHistogramBin[];
  }

  const rawRecord = rawSeries && typeof rawSeries === "object" ? rawSeries as Record<string, unknown> : null;
  const layoutProperties = rawRecord?.layoutPr && typeof rawRecord.layoutPr === "object"
    ? rawRecord.layoutPr as Record<string, unknown>
    : null;
  const rawBinning = layoutProperties?.binning && typeof layoutProperties.binning === "object"
    ? layoutProperties.binning as Record<string, unknown>
    : null;
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  const explicitWidth = typeof rawBinning?.binWidth === "number" && Number.isFinite(rawBinning.binWidth) && rawBinning.binWidth > 0
    ? rawBinning.binWidth
    : typeof rawBinning?.width === "number" && Number.isFinite(rawBinning.width) && rawBinning.width > 0
      ? rawBinning.width
      : undefined;
  const explicitCount = typeof rawBinning?.binCount === "number" && Number.isFinite(rawBinning.binCount) && rawBinning.binCount > 0
    ? rawBinning.binCount
    : typeof rawBinning?.count === "number" && Number.isFinite(rawBinning.count) && rawBinning.count > 0
      ? rawBinning.count
      : undefined;
  const closedRight = rawBinning?.intervalClosed === "r" || rawBinning?.intervalClosed === "right";
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
  const variance = values.reduce((sum, value) => sum + ((value - mean) ** 2), 0) / Math.max(1, values.length);
  const standardDeviation = Math.sqrt(Math.max(0, variance));
  const allIntegers = values.every((value) => Math.abs(value - Math.round(value)) < 1e-9);
  const scottWidth = standardDeviation > 0
    ? (3.49 * standardDeviation) / Math.cbrt(values.length)
    : undefined;
  const fallbackWidth = explicitCount != null
    ? (maxValue - minValue) / Math.max(1, explicitCount)
    : scottWidth ?? ((maxValue - minValue) / Math.max(1, Math.ceil(Math.log2(values.length) + 1)));
  const roughWidth = explicitWidth
    ?? (allIntegers
      ? Math.max(1, Math.ceil(Math.max(fallbackWidth, 1e-6)))
      : niceHistogramStep(Math.max(fallbackWidth, 1e-6)));
  const binWidth = Math.max(roughWidth, 1e-6);
  const start = explicitWidth != null || explicitCount != null
    ? Math.floor(minValue / binWidth) * binWidth
    : minValue;
  const end = Math.max(start + binWidth, start + Math.ceil((maxValue - start) / binWidth) * binWidth);
  const binCount = Math.max(1, Math.ceil((end - start) / binWidth));
  const bins = Array.from({ length: binCount }, (_, index) => {
    const lower = start + binWidth * index;
    const upper = lower + binWidth;
    return {
      count: 0,
      label: formatHistogramBinLabel(lower, upper, index, closedRight),
      lower,
      upper
    } satisfies ChartExHistogramBin;
  });

  values.forEach((value) => {
    if (!Number.isFinite(value)) {
      return;
    }
    const offset = (value - start) / binWidth;
    let binIndex = Math.floor(offset);
    if (closedRight && Math.abs(offset - Math.round(offset)) < 1e-9 && value > start) {
      binIndex -= 1;
    }
    if (value >= end) {
      binIndex = bins.length - 1;
    }
    if (value <= start) {
      binIndex = 0;
    }
    const target = bins[Math.max(0, Math.min(bins.length - 1, binIndex))];
    if (target) {
      target.count += 1;
    }
  });

  if (sortByFrequency) {
    bins.sort((left, right) => (
      right.count - left.count
      || left.lower - right.lower
    ));
  }
  return bins;
}

export function buildChartExHistogramSeries(
  series: XlsxChartSeries,
  rawSeries: unknown,
  sortByFrequency: boolean
) {
  const layout = resolveChartExSeriesLayout(rawSeries);
  const rawRecord = rawSeries && typeof rawSeries === "object" ? rawSeries as Record<string, unknown> : null;
  const hasBinning = Boolean(
    layout === "clusteredColumn"
    && rawRecord?.layoutPr
    && typeof rawRecord.layoutPr === "object"
    && (rawRecord.layoutPr as Record<string, unknown>).binning != null
  );
  if (!hasBinning) {
    return series;
  }

  const numericValues = series.values.filter((value): value is number => typeof value === "number" && Number.isFinite(value));
  if (numericValues.length === 0) {
    return series;
  }

  const bins = buildChartExHistogramBins(numericValues, rawSeries, sortByFrequency);
  if (bins.length === 0) {
    return series;
  }

  return {
    ...series,
    categories: bins.map((bin) => bin.label),
    categoriesRef: null,
    raw: {
      ...series.raw,
      chartExHistogramBins: bins,
      chartExSourceValues: numericValues
    },
    values: bins.map((bin) => bin.count)
  };
}

export function buildChartExParetoLineSeries(series: XlsxChartSeries, sourceRaw: unknown, index: number) {
  const counts = series.values.map((value) => (
    typeof value === "number" && Number.isFinite(value) ? value : 0
  ));
  const total = counts.reduce((sum, value) => sum + value, 0);
  let running = 0;
  const cumulative = counts.map((value) => {
    running += value;
    return total > 0 ? (running / total) * 100 : 0;
  });
  return {
    ...series,
    color: undefined,
    lineColor: undefined,
    markerColor: undefined,
    markerLineColor: undefined,
    markerSize: 7,
    markerSymbol: "circle",
    name: typeof (sourceRaw as { text?: unknown } | null)?.text === "string"
      ? (sourceRaw as { text: string }).text
      : "Pareto",
    raw: {
      ...(series.raw ?? {}),
      chartExLayout: "paretoLine",
      source: sourceRaw && typeof sourceRaw === "object" ? sourceRaw as Record<string, unknown> : undefined
    },
    values: cumulative
  };
}

export function resolveChartExTextFormula(raw: unknown) {
  if (typeof raw === "string" && raw.length > 0) {
    return raw;
  }
  if (!raw || typeof raw !== "object") {
    return undefined;
  }
  const record = raw as Record<string, unknown>;
  if (typeof record.formula === "string" && record.formula.length > 0) {
    return record.formula;
  }
  if (typeof record.text === "string" && record.text.length > 0) {
    return record.text;
  }
  if (typeof record.value === "string" && record.value.length > 0) {
    return record.value;
  }
  return undefined;
}

function resolveChartExTitleText(raw: unknown) {
  if (typeof raw === "string" && raw.length > 0) {
    return raw;
  }
  if (!raw || typeof raw !== "object") {
    return undefined;
  }
  const record = raw as Record<string, unknown>;
  if (typeof record.text === "string" && record.text.length > 0) {
    return record.text;
  }
  const nestedText = record.text && typeof record.text === "object"
    ? resolveChartExTextFormula(record.text)
    : undefined;
  if (nestedText) {
    return nestedText;
  }
  return typeof record.value === "string" && record.value.length > 0 ? record.value : undefined;
}

function resolveChartExFallbackCategoryReference(
  workbook: Workbook,
  fallbackSheetIndex: number,
  valueFormula: string | undefined
) {
  if (!valueFormula) {
    return null;
  }

  const resolved = resolveReferenceSheet(workbook, fallbackSheetIndex, valueFormula);
  if (!resolved.sheet || !resolved.range || resolved.range.start.col <= 0) {
    return null;
  }

  return normalizeChartReference({
    formula: buildA1RangeFormula(
      resolved.sheetName,
      {
        col: resolved.range.start.col - 1,
        row: resolved.range.start.row
      },
      {
        col: resolved.range.start.col - 1,
        row: resolved.range.end.row
      }
    )
  });
}

export function normalizeChartExSeries(
  workbook: Workbook,
  workbookSheetIndex: number,
  chartId: string,
  raw: unknown,
  dataById: Map<number, Record<string, unknown>>,
  index: number,
  chartType?: XlsxChart["chartType"]
): XlsxChartSeries {
  const series = raw && typeof raw === "object" ? raw as Record<string, unknown> : {};
  const dataId = typeof series.dataId === "number" ? series.dataId : null;
  const dataEntry = dataId != null ? dataById.get(dataId) ?? null : null;
  const dimensions = Array.isArray(dataEntry?.dimensions)
    ? dataEntry.dimensions.filter((value): value is Record<string, unknown> => Boolean(value && typeof value === "object"))
    : [];
  const categoryDimension = dimensions.find((dimension) => dimension.dimType === "cat")
    ?? dimensions.find((dimension) => dimension.dimType === "name")
    ?? null;
  const valueDimension = dimensions.find((dimension) => (
    dimension.dimType === "val"
    || dimension.dimType === "y"
    || dimension.dimType === "colorVal"
    || dimension.dimType === "size"
  ))
    ?? dimensions.find((dimension) => dimension !== categoryDimension)
    ?? categoryDimension;
  const categoryDimensionFormula = typeof categoryDimension?.formula === "string" ? categoryDimension.formula : undefined;
  const valueDimensionFormula = typeof valueDimension?.formula === "string" ? valueDimension.formula : undefined;
  const fallbackCategoryRef = (
    (chartType === "Sunburst" || chartType === "Treemap")
    && !categoryDimension
    && typeof valueDimensionFormula === "string"
  )
    ? resolveChartExFallbackCategoryReference(workbook, workbookSheetIndex, valueDimensionFormula)
    : null;
  const categoriesRef = categoryDimension
    ? normalizeChartReference({
        formula: categoryDimensionFormula
      })
    : fallbackCategoryRef;
  const valuesRef = valueDimension
    ? normalizeChartReference({
        formula: valueDimensionFormula
      })
    : null;
  const resolvedValueCells = resolveReferenceValues(workbook, workbookSheetIndex, valuesRef, "value");
  const values = resolvedValueCells.map((value) => (
    typeof value === "number" && Number.isFinite(value) ? value : null
  ));
  const colorStrings = chartType === "RegionMap" && valueDimension?.dimType === "colorStr"
    ? resolvedValueCells.map((value) => {
        if (typeof value === "string") {
          const trimmed = value.trim();
          return trimmed.length > 0 ? trimmed : null;
        }
        if (typeof value === "number" && Number.isFinite(value)) {
          return String(value);
        }
        return null;
      })
    : [];
  const categories = resolveReferenceValues(workbook, workbookSheetIndex, categoriesRef, "category");
  const hierarchyCategories = (
    chartType === "Sunburst" || chartType === "Treemap"
  )
    ? resolveReferenceRowPaths(workbook, workbookSheetIndex, categoriesRef)
    : [];
  const seriesTextFormula = resolveChartExTextFormula(series.text);
  const shapeProperties = series.shapeProperties && typeof series.shapeProperties === "object"
    ? series.shapeProperties as Record<string, unknown>
    : undefined;
  const rawFillColor = typeof shapeProperties?.solidFillHex === "string"
    ? normalizeHexColor(shapeProperties.solidFillHex)
    : null;
  const rawLineColor = typeof shapeProperties?.lineColorHex === "string"
    ? normalizeHexColor(shapeProperties.lineColorHex)
    : null;

  return {
    bubbleSizeRef: null,
    bubbleSizes: [],
    categories,
    categoriesRef,
    color: rawFillColor ?? undefined,
    dataPoints: Array.isArray(series.dataPoints) ? series.dataPoints : [],
    dataPointStyles: undefined,
    formatIdx: typeof series.formatIdx === "number" ? series.formatIdx : undefined,
    hidden: typeof series.hidden === "boolean" ? series.hidden : undefined,
    id: `${chartId}-series-${index}`,
    invertIfNegative: undefined,
    lineColor: rawLineColor ?? rawFillColor ?? undefined,
    lineWidthPx: typeof shapeProperties?.lineWidth === "number"
      ? Math.max(1, Number(shapeProperties.lineWidth) / EMU_PER_PIXEL)
      : undefined,
    marker: undefined,
    markerColor: rawFillColor ?? undefined,
    markerLineColor: rawLineColor ?? rawFillColor ?? undefined,
    markerSize: undefined,
    markerSymbol: undefined,
    name: typeof series.text === "string"
      ? series.text
      : seriesTextFormula
        ? resolveSeriesName(workbook, workbookSheetIndex, seriesTextFormula)
      : resolveChartReferenceLabel(workbook, workbookSheetIndex, valuesRef, `Series ${index + 1}`),
    negativeColor: undefined,
    negativeLineColor: undefined,
    raw: {
      ...series,
      chartExColorStrings: colorStrings,
      chartExHierarchyCategories: hierarchyCategories,
      data: dataEntry,
      dimType: typeof valueDimension?.dimType === "string" ? valueDimension.dimType : undefined
    },
    shapeProperties,
    smooth: undefined,
    values,
    valuesRef
  };
}

function collapseChartExPointSeries(chartType: XlsxChart["chartType"], series: XlsxChartSeries[]) {
  if (chartType !== "Funnel" && chartType !== "Waterfall") {
    if (
      (chartType === "Sunburst" || chartType === "Treemap")
      && series.length > 1
      && series.every((entry) => {
        const raw = entry.raw && typeof entry.raw === "object" ? entry.raw as Record<string, unknown> : null;
        return raw?.dimType === "size";
      })
    ) {
      const primarySeries = series.find((entry) => entry.hidden !== true) ?? series[0] ?? null;
      if (!primarySeries) {
        return series;
      }
      return [
        {
          ...primarySeries,
          dataPoints: [],
          hidden: false
        }
      ];
    }
    return series;
  }

  const primarySeries = series.find((entry) => entry.hidden !== true) ?? series[0] ?? null;
  if (!primarySeries) {
    return series;
  }

  return [
    {
      ...primarySeries,
      categories: [],
      categoriesRef: null,
      dataPoints: [],
      hidden: false
    }
  ];
}

export function normalizeChartExChart(
  workbook: Workbook,
  workbookSheetIndex: number,
  visibleSheetIndex: number,
  raw: unknown,
  index: number,
  themePalette?: XlsxThemePalette | null
): XlsxChart {
  const chart = raw && typeof raw === "object" ? raw as Record<string, unknown> : {};
  const plotArea = chart.plotArea && typeof chart.plotArea === "object"
    ? chart.plotArea as Record<string, unknown>
    : {};
  const rawSeries = Array.isArray(plotArea.series) ? plotArea.series : [];
  const seriesLayouts = rawSeries.map(resolveChartExSeriesLayout);
  const dataEntries = Array.isArray(chart.data) ? chart.data : [];
  const dataById = new Map<number, Record<string, unknown>>();
  dataEntries.forEach((entry) => {
    if (!entry || typeof entry !== "object") {
      return;
    }
    const record = entry as Record<string, unknown>;
    if (typeof record.id === "number") {
      dataById.set(record.id, record);
    }
  });
  const axes = Array.isArray(plotArea.axes)
    ? plotArea.axes.map(normalizeChartExAxis).filter((value): value is XlsxChartAxis => Boolean(value))
    : [];
  const primaryLayout = typeof chart.layout === "string"
    ? chart.layout
    : seriesLayouts.find((value): value is string => typeof value === "string" && value.length > 0);
  const fallbackTitle = humanizeChartExLayoutLabel(primaryLayout);
  const chartTitle = resolveChartExTitleText(chart.title) ?? (chart.title != null ? "Chart Title" : fallbackTitle);
  const chartType = resolveChartExLayoutChartType(primaryLayout);
  const normalizedSeries = rawSeries.map((entry, seriesIndex) => (
    normalizeChartExSeries(workbook, workbookSheetIndex, `chart-ex-${workbookSheetIndex}-${index}`, entry, dataById, seriesIndex, chartType)
  ));
  const clusteredColumnSeriesIndex = seriesLayouts.findIndex((layout) => layout === "clusteredColumn");
  const hasParetoLine = seriesLayouts.includes("paretoLine");
  const clusteredColumnAxisIds = clusteredColumnSeriesIndex >= 0
    ? resolveChartExSeriesAxisIds(rawSeries[clusteredColumnSeriesIndex])
    : [];
  const paretoLineSeriesIndex = seriesLayouts.findIndex((layout) => layout === "paretoLine");
  const paretoLineAxisIds = paretoLineSeriesIndex >= 0
    ? resolveChartExSeriesAxisIds(rawSeries[paretoLineSeriesIndex])
    : [];
  const primaryHistogramSeries = clusteredColumnSeriesIndex >= 0
    ? buildChartExHistogramSeries(normalizedSeries[clusteredColumnSeriesIndex] ?? normalizedSeries[0], rawSeries[clusteredColumnSeriesIndex], hasParetoLine)
    : null;
  const synthesizedParetoSeries = (
    hasParetoLine
    && primaryHistogramSeries
    && primaryHistogramSeries.values.length > 0
  )
    ? buildChartExParetoLineSeries(primaryHistogramSeries, rawSeries[paretoLineSeriesIndex], paretoLineSeriesIndex)
    : null;
  const resolvedSeries = synthesizedParetoSeries
    ? [primaryHistogramSeries!, synthesizedParetoSeries]
    : primaryHistogramSeries
      ? [
          primaryHistogramSeries,
          ...normalizedSeries.filter((_, seriesIndex) => seriesIndex !== clusteredColumnSeriesIndex)
        ]
      : collapseChartExPointSeries(chartType, normalizedSeries);
  const resolvedChartType = primaryHistogramSeries ? "ColumnClustered" : chartType;
  const resolvedGapWidth = primaryHistogramSeries ? 0 : undefined;
  const typeGroups = synthesizedParetoSeries
    ? [
        {
          axisIds: clusteredColumnAxisIds,
          chartType: "ColumnClustered",
          gapWidth: 0,
          raw: {
            gapWidth: 0,
            layout: "clusteredColumn"
          },
          series: [primaryHistogramSeries!]
        },
        {
          axisIds: paretoLineAxisIds,
          chartType: "Line",
          raw: {
            layout: "paretoLine"
          },
          series: [synthesizedParetoSeries]
        }
      ]
    : [];
  const normalizedChart: XlsxChart = {
    anchor: normalizeChartAnchor(chart.anchor),
    autoTitleDeleted: undefined,
    axes,
    axisLabelColor: undefined,
    axisLineColor: undefined,
    categoryAxis: axes[0] ?? null,
    chartAreaBorderColor: undefined,
    chartAreaFillColor: undefined,
    chartColorPalette: undefined,
    chartColorPaletteOffset: undefined,
    chartExLayout: primaryLayout,
    chartPath: undefined,
    chartStyleId: undefined,
    chartType: resolvedChartType,
    dataLabels: rawSeries.length > 0 && rawSeries[0] && typeof rawSeries[0] === "object"
      ? normalizeChartDataLabels((rawSeries[0] as Record<string, unknown>).dataLabels)
      : null,
    displayBlanksAs: undefined,
    editable: true,
    firstSliceAngle: undefined,
    fontFamily: undefined,
    gapWidth: resolvedGapWidth,
    holeSize: undefined,
    id: `chart-ex-${workbookSheetIndex}-${index}`,
    is3d: undefined,
    legend: normalizeChartExLegend(chart.legend),
    name: chartTitle,
    overlap: undefined,
    plotVisibleOnly: undefined,
    raw: chart,
    radarStyle: undefined,
    scatterStyle: undefined,
    roundedCorners: undefined,
    shape3d: undefined,
    seriesAxis: null,
    series: resolvedSeries,
    sheetIndex: visibleSheetIndex,
    showDlblsOverMax: undefined,
    sideWall: null,
    backWall: null,
    bubbleScale: undefined,
    bubble3d: undefined,
    floor: null,
    surfaceMaterial: undefined,
    textColor: undefined,
    title: chartTitle,
    titleColor: undefined,
    titleFontFamily: undefined,
    typeGroups,
    valueAxis: axes.find((axis) => axis.numberFormat || axis.majorGridlines) ?? axes[1] ?? null,
    varyColors: typeof chart.valueColors === "boolean" ? chart.valueColors : undefined,
    view3d: undefined,
    wireframe: undefined,
    workbookSheetIndex,
    zIndex: index
  };

  applyBuiltinChartDefaults(normalizedChart, themePalette);
  return normalizedChart;
}

