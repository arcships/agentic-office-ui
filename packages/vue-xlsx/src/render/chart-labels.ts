/** @jsxImportSource vue */
import type { XlsxChart, XlsxChartDataLabels, XlsxChartSeries } from "@arcships/xlsx-core";
import { normalizeCategoryLabel, safeNumber } from "./chart-shared";

export type ResolvedChartDataLabel = XlsxChartDataLabels & {
  fontSizePt?: number;
};

function pointLabelFor(labels: XlsxChartDataLabels, pointIndex: number | undefined) {
  if (pointIndex == null) {
    return null;
  }
  return labels.pointLabels?.find((entry) => entry.index === pointIndex) ?? null;
}

export function resolveDataLabelsForPoint(
  chart: XlsxChart,
  pointIndex: number | undefined,
  series?: XlsxChartSeries | null
): ResolvedChartDataLabel | null {
  const base = chart.dataLabels;
  if (!base) {
    return null;
  }
  const pointLabel = pointLabelFor(base, pointIndex);
  if (pointLabel?.deleted) {
    return null;
  }
  if (!pointLabel) {
    return base;
  }
  return {
    ...base,
    ...pointLabel
  };
}

export function hasVisibleDataLabels(labels: XlsxChartDataLabels | null | undefined) {
  if (!labels) {
    return false;
  }
  return Boolean(
    labels.showValue
    || labels.showSeriesName
    || labels.showCategoryName
    || labels.showBubbleSize
    || labels.showPercent
    || labels.showLegendKey
  );
}

export function formatDataLabelNumber(value: number, formatCode?: string) {
  if (!Number.isFinite(value)) {
    return "";
  }
  const normalized = (formatCode ?? "").toLowerCase();
  if (!normalized || normalized === "general") {
    return formatGeneralNumber(value);
  }
  const isPercent = normalized.includes("%");
  const numeric = isPercent ? value * 100 : value;
  const decimalMatch = /\.(0+)/.exec(normalized);
  const decimals = decimalMatch ? decimalMatch[1].length : 0;
  const useGrouping = /#,##|,##/.test(normalized);
  const formatted = numeric.toLocaleString(undefined, {
    maximumFractionDigits: decimals,
    minimumFractionDigits: decimals,
    useGrouping
  });
  return isPercent ? `${formatted}%` : formatted;
}

function formatGeneralNumber(value: number) {
  if (Number.isInteger(value)) {
    return String(value);
  }
  return value.toLocaleString(undefined, {
    maximumFractionDigits: 2
  });
}

export function buildDataLabelText(
  chart: XlsxChart,
  series: XlsxChartSeries | null | undefined,
  labels: XlsxChartDataLabels,
  options: {
    categoryLabel?: string;
    percentValue?: boolean;
    totalValue?: number | null;
    value?: number | null;
  }
) {
  const parts: string[] = [];
  if (labels.showSeriesName && series?.name) {
    parts.push(series.name);
  }
  if (labels.showCategoryName) {
    const category = normalizeCategoryLabel(options.categoryLabel);
    if (category.length > 0) {
      parts.push(category);
    }
  }

  const value = typeof options.value === "number" && Number.isFinite(options.value)
    ? options.value
    : safeNumber(options.value);
  if (value != null) {
    if (labels.showValue) {
      if (options.percentValue) {
        parts.push(`${formatGeneralNumber(value)}%`);
      } else {
        parts.push(formatDataLabelNumber(value, labels.numberFormat?.formatCode));
      }
    } else if (labels.showPercent) {
      const total = typeof options.totalValue === "number" && Number.isFinite(options.totalValue)
        ? Math.abs(options.totalValue)
        : null;
      if (total != null && total > 0) {
        parts.push(`${formatGeneralNumber((value / total) * 100)}%`);
      }
    }
  }
  return parts.join(", ");
}

export function resolveDataLabelPosition(labels: XlsxChartDataLabels | null | undefined, fallback: string) {
  const position = labels?.position;
  return typeof position === "string" && position.length > 0 ? position : fallback;
}
