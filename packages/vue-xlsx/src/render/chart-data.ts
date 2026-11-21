/** @jsxImportSource vue */
import type { XlsxChart, XlsxChartAxis, XlsxChartSeries } from "@arcships/xlsx-core";
import { safeNumber, normalizeCategoryLabel, clamp, normalizeRendererHexColor, lightenColor, darkenColor, resolveChartSeriesLabel } from "./chart-shared";
import { formatExcelDateSerial, isLikelyDateFormatCode } from "./chart-date";
import {
  chartImageFillUrl,
  chartPointImageFillKey,
  chartSeriesImageFillKey,
  sanitizeChartImageSource,
} from "./chart-image-fills";
import type { ChartLabels, ComboRenderableGroup } from "./chart-types";

export function chartSeriesColor(chart: XlsxChart, seriesIndex: number) {
  const series = chart.series[seriesIndex];
  // A picture fill replaces the solid fill entirely; renderers paint the
  // `url(#…)` pattern the chart's `<defs>` define.
  if (sanitizeChartImageSource(series?.fillImage)) {
    return chartImageFillUrl(chart, chartSeriesImageFillKey(seriesIndex));
  }
  const paletteColor = chart.chartColorPalette?.[seriesIndex % Math.max(1, chart.chartColorPalette.length)];
  return series?.color ?? series?.lineColor ?? paletteColor ?? chart.textColor ?? "#222222";
}

export function chartSeriesStrokeColor(chart: XlsxChart, seriesIndex: number) {
  const series = chart.series[seriesIndex];
  const paletteColor = chart.chartColorPalette?.[seriesIndex % Math.max(1, chart.chartColorPalette.length)];
  return series?.lineColor ?? series?.color ?? paletteColor ?? chart.textColor ?? "#222222";
}

export function chartPointFillImage(chart: XlsxChart, seriesIndex: number, pointIndex: number) {
  const pointStyle = chart.series[seriesIndex]?.dataPointStyles?.find((entry) => entry.index === pointIndex);
  return sanitizeChartImageSource(pointStyle?.fillImage)
    ? chartImageFillUrl(chart, chartPointImageFillKey(seriesIndex, pointIndex))
    : undefined;
}

export function chartPointColor(chart: XlsxChart, pointIndex: number, seriesIndex = 0) {
  const pointStyle = chart.series[seriesIndex]?.dataPointStyles?.find((entry) => entry.index === pointIndex);
  if (pointStyle && sanitizeChartImageSource(pointStyle.fillImage)) {
    return chartImageFillUrl(chart, chartPointImageFillKey(seriesIndex, pointIndex));
  }
  if (pointStyle?.color) {
    return pointStyle.color;
  }
  if (sanitizeChartImageSource(chart.series[seriesIndex]?.fillImage)) {
    return chartImageFillUrl(chart, chartSeriesImageFillKey(seriesIndex));
  }
  const rawPoint = chart.series[seriesIndex]?.dataPoints?.[pointIndex];
  if (rawPoint && typeof rawPoint === "object") {
    const pointRecord = rawPoint as Record<string, unknown>;
    if (typeof pointRecord.color === "string") {
      return pointRecord.color;
    }
    if (typeof pointRecord.fillColor === "string") {
      return pointRecord.fillColor;
    }
    if (typeof pointRecord.solidFillHex === "string") {
      const normalized = pointRecord.solidFillHex.replace(/^#/, "");
      if (/^[0-9a-fA-F]{6}$/.test(normalized)) {
        return `#${normalized.toLowerCase()}`;
      }
    }
    const shapeProperties = pointRecord.shapeProperties && typeof pointRecord.shapeProperties === "object"
      ? pointRecord.shapeProperties as Record<string, unknown>
      : null;
    if (shapeProperties && typeof shapeProperties.solidFillHex === "string") {
      const normalized = shapeProperties.solidFillHex.replace(/^#/, "");
      if (/^[0-9a-fA-F]{6}$/.test(normalized)) {
        return `#${normalized.toLowerCase()}`;
      }
    }
  }
  const palette = chart.chartColorPalette;
  if (palette && palette.length > 0) {
    const offset = chart.chartColorPaletteOffset ?? 0;
    return palette[(pointIndex + offset) % palette.length] ?? palette[pointIndex % palette.length];
  }
  return chartSeriesColor(chart, seriesIndex);
}

export function isHistogramLikeSeries(series: XlsxChartSeries | null | undefined) {
  const raw = series?.raw && typeof series.raw === "object"
    ? series.raw as Record<string, unknown>
    : null;
  return Array.isArray(raw?.chartExHistogramBins) && raw.chartExHistogramBins.length > 0;
}

export function isHistogramLikeChart(chart: XlsxChart) {
  return chart.series.some((series) => isHistogramLikeSeries(series));
}

export function normalizeBuiltinPieStyleId(styleId: number | undefined) {
  if (typeof styleId !== "number" || !Number.isFinite(styleId)) {
    return null;
  }
  return styleId >= 100 ? styleId - 100 : styleId;
}

export function getBuiltinPiePalette(chart: XlsxChart, seriesIndex: number) {
  // A series picture fill paints every slice with the same image; a built-in
  // palette would hide it.
  if (sanitizeChartImageSource(chart.series[seriesIndex]?.fillImage)) {
    return null;
  }
  const normalized = normalizeBuiltinPieStyleId(chart.chartStyleId);
  if (normalized !== 32) {
    return null;
  }
  const baseColor = chart.series[seriesIndex]?.color
    ?? chart.series[seriesIndex]?.lineColor
    ?? chart.chartColorPalette?.[0]
    ?? null;
  if (!baseColor) {
    return null;
  }
  return [
    lightenColor(baseColor, 0.16),
    darkenColor(baseColor, 0.42),
    baseColor,
    darkenColor(baseColor, 0.18),
    lightenColor(baseColor, 0.08),
    darkenColor(baseColor, 0.3)
  ];
}

export function resolvePiePointColor(chart: XlsxChart, pointIndex: number, seriesIndex = 0) {
  const pointStyle = chart.series[seriesIndex]?.dataPointStyles?.find((entry) => entry.index === pointIndex);
  if (pointStyle && sanitizeChartImageSource(pointStyle.fillImage)) {
    return chartImageFillUrl(chart, chartPointImageFillKey(seriesIndex, pointIndex));
  }
  if (pointStyle?.color) {
    return pointStyle.color;
  }
  const builtinPalette = getBuiltinPiePalette(chart, seriesIndex);
  if (builtinPalette && builtinPalette.length > 0) {
    return builtinPalette[pointIndex % builtinPalette.length] ?? builtinPalette[0];
  }
  return chartPointColor(chart, pointIndex, seriesIndex);
}

export function selectPrimaryPieSeriesIndex(chart: XlsxChart) {
  let bestIndex = 0;
  let bestScore = Number.NEGATIVE_INFINITY;

  chart.series.forEach((series, index) => {
    let positiveCount = 0;
    let total = 0;
    for (const rawValue of series.values) {
      const value = safeNumber(rawValue);
      if (value != null && value > 0) {
        positiveCount += 1;
        total += value;
      }
    }
    const score = positiveCount * 1_000_000 + total;
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestIndex;
}

export function chartSeriesBarColors(
  chart: XlsxChart,
  seriesIndex: number,
  value: number,
  negativeFillMode: "chartArea" | "none" | "series"
) {
  const series = chart.series[seriesIndex];
  const defaultFill = chartSeriesColor(chart, seriesIndex);
  const defaultStroke = chartSeriesStrokeColor(chart, seriesIndex);
  const defaultStrokeWidth = typeof series?.lineWidthPx === "number" && Number.isFinite(series.lineWidthPx)
    ? Math.max(0.8, Math.min(4, series.lineWidthPx))
    : 1;
  if (value < 0 && series?.invertIfNegative) {
    const resolvedNegativeFill = negativeFillMode === "none"
      ? "none"
      : series.negativeColor
        ?? (negativeFillMode === "chartArea" ? chart.chartAreaFillColor : undefined)
        ?? defaultFill;
    return {
      fill: resolvedNegativeFill,
      stroke: series.negativeLineColor ?? defaultStroke,
      strokeWidth: defaultStrokeWidth
    };
  }
  return {
    fill: defaultFill,
    stroke: defaultStroke,
    strokeWidth: defaultStrokeWidth
  };
}

export function resolveCategoryBandPadding(gapWidth: number | undefined) {
  const normalizedGap = typeof gapWidth === "number" && Number.isFinite(gapWidth)
    ? clamp(gapWidth, 0, 500)
    : 150;
  const inner = clamp(normalizedGap / (100 + normalizedGap), 0.05, 0.88);
  const outer = clamp(inner * 0.5, 0, 0.45);
  return { inner, outer };
}

export function normalizeChartMarkerSymbol(value: string | undefined) {
  if (!value || value === "none") {
    return "none";
  }
  if (value === "auto") {
    return "circle";
  }
  return value;
}

export function normalizeRenderableChartType(chart: XlsxChart) {
  if (chart.chartType === "ScatterSmooth") {
    return "ScatterSmooth";
  }
  if (chart.chartType === "Pie" && chart.is3d) {
    return "Pie3D";
  }
  if (
    chart.chartType === "Pie"
    && chart.series.some((series) => Array.isArray(series.dataPoints) && series.dataPoints.some((point) => (
      point != null
      && typeof point === "object"
      && "explosion" in point
      && typeof (point as { explosion?: unknown }).explosion === "number"
      && ((point as { explosion?: number }).explosion ?? 0) > 0
    )))
  ) {
    return "PieExploded";
  }
  if (chart.chartType === "Unsupported(c:ofPieChart)") {
    return "BarOfPie";
  }
  return chart.chartType;
}

export function estimateReferencePointCount(formula: string | undefined) {
  if (!formula) {
    return 0;
  }
  const bang = formula.lastIndexOf("!");
  const rawRange = (bang >= 0 ? formula.slice(bang + 1) : formula).replace(/\$/g, "");
  const match = /^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/.exec(rawRange.trim());
  if (!match) {
    return 0;
  }
  const startRow = Number(match[2]);
  const endRow = Number(match[4]);
  if (!Number.isFinite(startRow) || !Number.isFinite(endRow)) {
    return 0;
  }
  return Math.abs(endRow - startRow) + 1;
}

export {
  excelSerialToDate,
  formatExcelDateSerial,
  formatExcelSerial,
  isDateLikeFormat,
  isLikelyDateFormatCode
} from "./chart-date";
export type { ExcelSerialFormatOptions } from "./chart-date";

export type ParsedExcelNumberFormat = {
  currencySuffix: string;
  currencySymbol: string;
  grouping: boolean;
  kind: "currency" | "date" | "number" | "percent";
  maxDecimals: number;
  minDecimals: number;
  prefix: string;
  suffix: string;
};

const EXCEL_CURRENCY_SYMBOLS = ["R$", "$", "€", "£", "¥", "￥", "₹", "₩", "₽", "₪", "₫", "฿"];

function splitExcelFormatSections(formatCode: string) {
  const sections: string[] = [];
  let current = "";
  let quoted = false;
  let bracketed = false;
  for (const character of Array.from(formatCode)) {
    if (character === "\"" && !bracketed) {
      quoted = !quoted;
    } else if (character === "[" && !quoted) {
      bracketed = true;
    } else if (character === "]" && !quoted) {
      bracketed = false;
    }
    if (character === ";" && !quoted && !bracketed) {
      sections.push(current);
      current = "";
      continue;
    }
    current += character;
  }
  sections.push(current);
  return sections;
}

function findCurrencySymbol(value: string) {
  for (const symbol of EXCEL_CURRENCY_SYMBOLS) {
    if (value.includes(symbol)) {
      return symbol;
    }
  }
  return "";
}

export function parseExcelNumberFormat(formatCode: string | undefined): ParsedExcelNumberFormat {
  const raw = (formatCode ?? "").trim();
  if (isLikelyDateFormatCode(raw)) {
    return {
      currencySuffix: "",
      currencySymbol: "",
      grouping: false,
      kind: "date",
      maxDecimals: 0,
      minDecimals: 0,
      prefix: "",
      suffix: ""
    };
  }

  const section = splitExcelFormatSections(raw)[0] ?? "";
  let currencySymbol = "";
  let text = section.replace(/\[([^\]]*)\]/g, (_match, inner: string) => {
    const content = String(inner);
    if (content.startsWith("$")) {
      const symbol = content.slice(1).split("-")[0] ?? "";
      if (symbol) {
        currencySymbol = symbol;
        return symbol;
      }
    }
    return "";
  });
  text = text
    .replace(/\\(.)/g, "$1")
    .replace(/"([^"]*)"/g, "$1")
    .replace(/[_*]./g, "");

  const percent = text.includes("%");
  if (!currencySymbol) {
    currencySymbol = findCurrencySymbol(text);
  }

  const numericMatch = /[0#?][0#?,.]*/.exec(text);
  if (!numericMatch || numericMatch.index == null) {
    return {
      currencySuffix: "",
      currencySymbol,
      grouping: /,/.test(text),
      kind: currencySymbol ? "currency" : "number",
      maxDecimals: 0,
      minDecimals: 0,
      prefix: text.trim(),
      suffix: ""
    };
  }

  const numericStart = numericMatch.index;
  const numericEnd = numericStart + numericMatch[0].length;
  const numericPattern = text.slice(numericStart, numericEnd);
  const decimalMatch = /\.([0#?]+)/.exec(numericPattern);
  const decimalPlaceholders = decimalMatch?.[1] ?? "";
  const minDecimals = (decimalPlaceholders.match(/0/g) ?? []).length;
  const maxDecimals = decimalPlaceholders.length;
  const grouping = /,/.test(numericPattern.split(".")[0] ?? numericPattern);

  const prefixText = text.slice(0, numericStart).trim();
  const suffixText = text.slice(numericEnd).replace(/%/g, "").trim();
  const prefix = currencySymbol && prefixText.includes(currencySymbol)
    ? currencySymbol
    : prefixText;
  const suffix = currencySymbol && suffixText.includes(currencySymbol)
    ? currencySymbol
    : suffixText;

  return {
    currencySuffix: currencySymbol && suffixText.includes(currencySymbol) ? currencySymbol : "",
    currencySymbol,
    grouping,
    kind: percent ? "percent" : currencySymbol ? "currency" : "number",
    maxDecimals,
    minDecimals,
    prefix: prefix.replace(/[()]/g, "").trim(),
    suffix: suffix.replace(/[()]/g, "").trim()
  };
}

export function formatDefaultTickValue(value: number) {
  if (Math.abs(value) >= 1000) {
    return value.toLocaleString(undefined, { maximumFractionDigits: 0 });
  }
  return value.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

export function formatExcelNumberValue(value: number, formatCode: string | undefined, use1904 = false) {
  const raw = (formatCode ?? "").trim();
  if (!raw || /^general$/i.test(raw) || /^@+$/.test(raw)) {
    return null;
  }
  if (isLikelyDateFormatCode(raw)) {
    return formatExcelDateSerial(value, raw, use1904);
  }
  const parsed = parseExcelNumberFormat(raw);
  if (parsed.kind === "percent") {
    return new Intl.NumberFormat(undefined, {
      maximumFractionDigits: parsed.maxDecimals,
      minimumFractionDigits: parsed.minDecimals,
      style: "percent",
      useGrouping: parsed.grouping
    }).format(value);
  }
  const digits = new Intl.NumberFormat(undefined, {
    maximumFractionDigits: parsed.maxDecimals,
    minimumFractionDigits: parsed.minDecimals,
    useGrouping: parsed.grouping
  }).format(Math.abs(value));
  const sign = value < 0 ? "-" : "";
  return `${sign}${parsed.prefix}${digits}${parsed.suffix}`;
}

export type ChartAxisTickFormatterOptions = {
  defaultFormatter?: (value: number) => string;
  percentValuesAreScaled?: boolean;
};

export function createChartAxisTickFormatter(
  axis: XlsxChartAxis | null | undefined,
  chart: XlsxChart | null | undefined,
  options: ChartAxisTickFormatterOptions = {}
) {
  const defaultFormatter = options.defaultFormatter ?? formatDefaultTickValue;
  const formatCode = axis?.numberFormat?.formatCode?.trim();
  if (!formatCode || /^general$/i.test(formatCode) || /^@+$/.test(formatCode)) {
    if (options.percentValuesAreScaled) {
      const percentFormatter = new Intl.NumberFormat(undefined, {
        maximumFractionDigits: 0,
        minimumFractionDigits: 0,
        style: "percent"
      });
      return (value: number) => percentFormatter.format(value / 100);
    }
    return defaultFormatter;
  }

  const rawRecord = chart?.raw && typeof chart.raw === "object"
    ? chart.raw as Record<string, unknown>
    : null;
  const use1904 = rawRecord?.date1904 === true;
  if (isLikelyDateFormatCode(formatCode)) {
    return (value: number) => formatExcelDateSerial(value, formatCode, use1904);
  }

  const parsed = parseExcelNumberFormat(formatCode);
  if (parsed.kind === "percent") {
    const percentFormatter = new Intl.NumberFormat(undefined, {
      maximumFractionDigits: parsed.maxDecimals,
      minimumFractionDigits: parsed.minDecimals,
      style: "percent",
      useGrouping: parsed.grouping
    });
    if (options.percentValuesAreScaled) {
      return (value: number) => percentFormatter.format(value / 100);
    }
    return (value: number) => percentFormatter.format(value);
  }

  const numberFormatter = new Intl.NumberFormat(undefined, {
    maximumFractionDigits: parsed.maxDecimals,
    minimumFractionDigits: parsed.minDecimals,
    useGrouping: parsed.grouping
  });
  return (value: number) => {
    const sign = value < 0 ? "-" : "";
    return `${sign}${parsed.prefix}${numberFormatter.format(Math.abs(value))}${parsed.suffix}`;
  };
}

export function formatCategoryLabel(chart: XlsxChart, value: unknown, fallback = "") {
  const rawRecord = chart.raw && typeof chart.raw === "object"
    ? chart.raw as Record<string, unknown>
    : null;
  const numeric = typeof value === "number" && Number.isFinite(value)
    ? value
    : null;
  const formatCode = chart.categoryAxis?.numberFormat?.formatCode;
  if (numeric != null && formatCode) {
    const formatted = formatExcelNumberValue(numeric, formatCode, rawRecord?.date1904 === true);
    if (formatted != null) {
      return formatted;
    }
  }

  const normalized = normalizeCategoryLabel(value);
  if (normalized.length > 0) {
    return normalized;
  }
  return fallback;
}

export function getCategoryLabels(chart: XlsxChart) {
  const primaryCategories = chart.series[0]?.categories ?? [];
  const includeReferenceCategoryCount = chart.plotVisibleOnly !== true;
  const referenceCategoryCount = includeReferenceCategoryCount
    ? Math.max(
      0,
      ...chart.series.map((series) => Math.max(
        estimateReferencePointCount(series.categoriesRef?.formula),
        estimateReferencePointCount(series.valuesRef?.formula)
      ))
    )
    : 0;
  const categoryCount = Math.max(
    referenceCategoryCount,
    primaryCategories.length,
    ...chart.series.map((series) => Math.max(series.categories.length, series.values.length))
  );
  if (categoryCount <= 0) {
    return [];
  }
  const hasAnyExplicitCategory = chart.series.some((series) => (
    series.categories.some((value) => normalizeCategoryLabel(value).length > 0)
  ));
  const fallbackToImplicitOrdinal = chart.series.some((series) => {
    const categoriesLength = series.categories.length;
    if (categoriesLength === 0) {
      return false;
    }
    return series.categories.every((value) => normalizeCategoryLabel(value).length === 0);
  }) || !hasAnyExplicitCategory;
  return Array.from({ length: categoryCount }, (_, categoryIndex) => {
    const primary = primaryCategories[categoryIndex];
    if (primary != null) {
      const normalizedPrimary = normalizeCategoryLabel(primary);
      if (normalizedPrimary.length > 0) {
        return formatCategoryLabel(chart, primary, normalizedPrimary);
      }
    }
    const fallback = chart.series
      .map((series) => series.categories[categoryIndex])
      .find((value) => normalizeCategoryLabel(value).length > 0);
    if (fallback != null) {
      return formatCategoryLabel(chart, fallback, normalizeCategoryLabel(fallback));
    }
    return fallbackToImplicitOrdinal ? String(categoryIndex + 1) : "";
  });
}

export function getCategoryNumericValues(chart: XlsxChart) {
  const categoryCount = Math.max(
    getCategoryLabels(chart).length,
    ...chart.series.map((series) => series.categories.length)
  );
  return Array.from({ length: categoryCount }, (_, categoryIndex) => {
    for (const series of chart.series) {
      const numeric = coerceLooseNumber(series.categories[categoryIndex]);
      if (numeric != null) {
        return numeric;
      }
    }
    return null;
  });
}

export function computePercentStackedTotals(
  valuesBySeries: Array<Array<number | null>>,
  categoryCount: number
) {
  return Array.from({ length: categoryCount }, (_, categoryIndex) => {
    const total = valuesBySeries.reduce((sum, seriesValues) => (
      sum + Math.abs(seriesValues[categoryIndex] ?? 0)
    ), 0);
    return Math.abs(total) < 1e-9 ? 1 : total;
  });
}

export function isComboChart(chart: XlsxChart) {
  const typeGroups = chart.typeGroups ?? [];
  if (typeGroups.length < 2) {
    return false;
  }
  const distinctChartTypes = new Set(typeGroups.map((group) => group.chartType));
  return distinctChartTypes.size > 1;
}

export function getComboLegendSeries(chart: XlsxChart, labels?: ChartLabels) {
  if (!isComboChart(chart)) {
    return chart.series.map((series, index) => ({
      color: chartSeriesColor(chart, index),
      label: resolveChartSeriesLabel(series, index, labels)
    }));
  }
  return (chart.typeGroups ?? []).flatMap((group) => (
    group.series.map((series, seriesIndex) => ({
      color: series.lineColor ?? series.markerColor ?? series.color ?? chartSeriesColor(chart, seriesIndex),
      label: resolveChartSeriesLabel(series, seriesIndex, labels)
    }))
  ));
}

export function findAxisForGroup(
  chart: XlsxChart,
  axisIds: number[],
  positionMatcher: (position: string | undefined) => boolean,
  allowAnyMatch = false,
  axisTypes?: Array<XlsxChartAxis["axisType"]>
) {
  const matchesAxisType = (axis: XlsxChartAxis) => (
    axisTypes == null
    || axis.axisType == null
    || axisTypes.includes(axis.axisType)
  );
  const positionedMatch = chart.axes.find((axis) => (
    axis.id != null
    && axisIds.includes(axis.id)
    && matchesAxisType(axis)
    && positionMatcher(axis.position)
  ));
  if (positionedMatch) {
    return positionedMatch;
  }
  if (!allowAnyMatch) {
    return null;
  }
  return chart.axes.find((axis) => (
    axis.id != null
    && axisIds.includes(axis.id)
    && matchesAxisType(axis)
  )) ?? null;
}

export function buildComboGroups(chart: XlsxChart): ComboRenderableGroup[] {
  return (chart.typeGroups ?? []).map((group) => {
    const axisIds = group.axisIds ?? [];
    const categoryAxis = findAxisForGroup(
      chart,
      axisIds,
      (position) => position === "b" || position === "t",
      false,
      ["category", "date"]
    )
      ?? chart.categoryAxis
      ?? null;
    const valueAxis = findAxisForGroup(
      chart,
      axisIds,
      (position) => position === "l" || position === "r",
      true,
      ["value"]
    )
      ?? chart.valueAxis
      ?? null;
    return {
      axisIds,
      categoryAxis,
      chartType: group.chartType,
      dataLabels: group.dataLabels,
      gapWidth: group.gapWidth,
      is3d: group.is3d,
      raw: group.raw,
      series: group.series,
      valueAxis
    };
  }).filter((group) => group.series.length > 0);
}

/**
 * Returns the value axis used by a second combo group when it differs from the
 * first group's axis. A combo chart that declares both groups on the same
 * `axId` has no secondary axis.
 */
export function resolveComboSecondaryValueAxis(chart: XlsxChart): XlsxChartAxis | null {
  const groups = buildComboGroups(chart);
  if (groups.length < 2) {
    return null;
  }
  const primary = groups[0]?.valueAxis ?? null;
  for (const group of groups.slice(1)) {
    const axis = group.valueAxis;
    if (!axis) {
      continue;
    }
    if (primary && (
      axis === primary
      || (axis.id != null && primary.id != null && axis.id === primary.id)
    )) {
      continue;
    }
    return axis;
  }
  return null;
}

export function resolveRenderableSeriesValue(rawValue: unknown, displayBlanksAs: string | undefined) {
  const numeric = safeNumber(rawValue);
  if (numeric != null) {
    return numeric;
  }
  return displayBlanksAs === "zero" ? 0 : null;
}

export function coerceLooseNumber(value: unknown): number | null {
  const strict = safeNumber(value);
  if (strict != null) {
    return strict;
  }
  if (Array.isArray(value)) {
    for (const entry of value) {
      const nested = coerceLooseNumber(entry);
      if (nested != null) {
        return nested;
      }
    }
    return null;
  }
  if (value && typeof value === "object") {
    const record = value as Record<string, unknown>;
    for (const key of ["x", "value", "v", "num", "number", "raw"]) {
      const nested = coerceLooseNumber(record[key]);
      if (nested != null) {
        return nested;
      }
    }
    return null;
  }
  if (typeof value === "string") {
    const match = /-?\d+(?:\.\d+)?/.exec(value);
    if (!match) {
      return null;
    }
    const parsed = Number(match[0]);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

export function buildPieEntries(chart: XlsxChart, seriesIndex = selectPrimaryPieSeriesIndex(chart)) {
  const categories = getCategoryLabels(chart);
  const values = chart.series[seriesIndex]?.values ?? [];
  return values
    .map((rawValue, index) => ({
      color: resolvePiePointColor(chart, index, seriesIndex),
      index,
      label: normalizeCategoryLabel(categories[index]),
      value: Math.max(0, safeNumber(rawValue) ?? 0)
    }))
    .filter((entry) => entry.value > 0);
}
