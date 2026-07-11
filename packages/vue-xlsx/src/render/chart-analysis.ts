/** @jsxImportSource vue */
import type { HierarchyNode, HierarchyRectangularNode } from "d3-hierarchy";
import { treemapBinary, treemapDice } from "d3-hierarchy";
import type { XlsxChart, XlsxChartSeries } from "@arcships/xlsx-core";
import { safeNumber, normalizeCategoryLabel, clamp, lightenColor } from "./chart-shared";
import { chartSeriesColor, chartPointColor, getCategoryLabels } from "./chart-data";
import type { ChartHierarchyDatum, ChartStage, BoxWhiskerStats } from "./chart-types";

export function resolveChartStageSubtotal(series: XlsxChart["series"][number]) {
  const raw = series.raw && typeof series.raw === "object"
    ? series.raw as Record<string, unknown>
    : null;
  const layoutProperties = raw?.layoutProperties && typeof raw.layoutProperties === "object"
    ? raw.layoutProperties as Record<string, unknown>
    : null;
  const subtotals = Array.isArray(layoutProperties?.subtotals) ? layoutProperties.subtotals : [];
  return subtotals.length > 0 || layoutProperties?.aggregation === true;
}

export function buildChartStages(chart: XlsxChart) {
  if (chart.chartType === "Funnel" || chart.chartType === "Waterfall") {
    const primarySeriesIndex = Math.max(0, chart.series.findIndex((series) => series.hidden !== true));
    const primarySeries = chart.series[primarySeriesIndex] ?? null;
    if (!primarySeries) {
      return [];
    }

    const labels = getCategoryLabels(chart);
    return primarySeries.values
      .map((rawValue, index): ChartStage | null => {
        const value = safeNumber(rawValue);
        if (value == null || !Number.isFinite(value)) {
          return null;
        }
        return {
          color: chart.varyColors
            ? chartPointColor(chart, index, primarySeriesIndex)
            : chartSeriesColor(chart, primarySeriesIndex),
          isSubtotal: false,
          label: normalizeCategoryLabel(labels[index]) || String(index + 1),
          value
        };
      })
      .filter((stage): stage is ChartStage => stage != null);
  }

  return chart.series
    .map((series, index): ChartStage | null => {
      const value = series.values.reduce<number>((sum, entry) => sum + (safeNumber(entry) ?? 0), 0);
      if (!Number.isFinite(value)) {
        return null;
      }
      const formula = typeof series.valuesRef?.formula === "string" ? series.valuesRef.formula : "";
      const label = series.name
        ?? (formula.length > 0 ? formula.replace(/^.*!/, "").replace(/\$/g, "") : `Series ${index + 1}`);
      return {
        color: chartSeriesColor(chart, typeof series.formatIdx === "number" ? series.formatIdx : index),
        isSubtotal: resolveChartStageSubtotal(series),
        label,
        value
      };
    })
    .filter((stage): stage is ChartStage => stage != null);
}

export function buildHierarchyData(chart: XlsxChart) {
  const root: ChartHierarchyDatum = {
    children: [],
    name: chart.title ?? chart.name ?? "Root"
  };
  const rootChildren = root.children ?? [];
  const topLevelIndexByName = new Map<string, number>();
  const primaryHierarchySeries = chart.series.find((series) => {
    const raw = series.raw && typeof series.raw === "object" ? series.raw as Record<string, unknown> : null;
    return Array.isArray(raw?.chartExHierarchyCategories);
  }) ?? null;
  const primaryHierarchyPaths = (() => {
    if (!primaryHierarchySeries) {
      return null;
    }
    const raw = primaryHierarchySeries.raw as Record<string, unknown>;
    return Array.isArray(raw.chartExHierarchyCategories)
      ? raw.chartExHierarchyCategories.map((entry) => Array.isArray(entry)
        ? entry.map((value) => normalizeCategoryLabel(value)).filter((value) => value.length > 0)
        : [])
      : null;
  })();
  const rowCount = primaryHierarchyPaths
    ? primaryHierarchyPaths.length
    : Math.max(0, ...chart.series.map((series) => series.values.length));

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const path = primaryHierarchyPaths?.[rowIndex] ?? chart.series
      .map((series) => normalizeCategoryLabel(series.categories[rowIndex] ?? series.values[rowIndex]))
      .filter((value) => value.length > 0);
    if (path.length === 0) {
      continue;
    }

    let current = root;
    path.forEach((part, levelIndex) => {
      current.children = current.children ?? [];
      const preserveDuplicateLeaf = levelIndex === path.length - 1 && path.length === 1;
      let next = preserveDuplicateLeaf
        ? undefined
        : current.children.find((child) => child.name === part);
      if (!next) {
        next = { children: [], name: part };
        if (levelIndex === 0) {
          const topLevelIndex = topLevelIndexByName.size;
          topLevelIndexByName.set(part, topLevelIndex);
          next.colorIndex = topLevelIndex;
        } else {
          next.colorIndex = current.colorIndex;
        }
        current.children.push(next);
      }
      current = next;
    });

    const valueSource = primaryHierarchySeries ?? chart.series[chart.series.length - 1] ?? null;
    current.value = (current.value ?? 0) + Math.max(0.0001, safeNumber(valueSource?.values[rowIndex]) ?? 1);
  }

  return rootChildren.length > 0 ? root : null;
}

export function resolveBoxWhiskerQuartileMethod(series: XlsxChartSeries) {
  const raw = series.raw && typeof series.raw === "object" ? series.raw as Record<string, unknown> : null;
  const layoutProperties = raw?.layoutProperties && typeof raw.layoutProperties === "object"
    ? raw.layoutProperties as Record<string, unknown>
    : null;
  const statistics = layoutProperties?.statistics && typeof layoutProperties.statistics === "object"
    ? layoutProperties.statistics as Record<string, unknown>
    : null;
  return statistics?.quartileMethod === "inclusive" ? "inclusive" : "exclusive";
}

export function resolveBoxWhiskerVisibility(series: XlsxChartSeries) {
  const raw = series.raw && typeof series.raw === "object" ? series.raw as Record<string, unknown> : null;
  const layoutProperties = raw?.layoutProperties && typeof raw.layoutProperties === "object"
    ? raw.layoutProperties as Record<string, unknown>
    : null;
  const visibility = layoutProperties?.visibility && typeof layoutProperties.visibility === "object"
    ? layoutProperties.visibility as Record<string, unknown>
    : null;
  return {
    meanLine: visibility?.meanLine === true,
    meanMarker: visibility?.meanMarker !== false,
    nonoutliers: visibility?.nonoutliers === true,
    outliers: visibility?.outliers !== false
  };
}

export function computePercentile(sortedValues: number[], percentile: number, method: "exclusive" | "inclusive") {
  const count = sortedValues.length;
  if (count === 0) {
    return 0;
  }
  if (count === 1) {
    return sortedValues[0] ?? 0;
  }

  const rank = method === "exclusive"
    ? percentile * (count + 1)
    : 1 + percentile * (count - 1);
  if (rank <= 1) {
    return sortedValues[0] ?? 0;
  }
  if (rank >= count) {
    return sortedValues[count - 1] ?? 0;
  }

  const lowerIndex = Math.floor(rank) - 1;
  const upperIndex = Math.ceil(rank) - 1;
  const fraction = rank - Math.floor(rank);
  const lower = sortedValues[Math.max(0, lowerIndex)] ?? sortedValues[0] ?? 0;
  const upper = sortedValues[Math.max(0, upperIndex)] ?? sortedValues[count - 1] ?? 0;
  return lower + (upper - lower) * fraction;
}

export function computeBoxWhiskerStats(series: XlsxChartSeries): BoxWhiskerStats | null {
  const sortedValues = series.values
    .map((value) => safeNumber(value))
    .filter((value): value is number => value != null)
    .sort((left, right) => left - right);
  if (sortedValues.length === 0) {
    return null;
  }

  const quartileMethod = resolveBoxWhiskerQuartileMethod(series);
  const q1 = computePercentile(sortedValues, 0.25, quartileMethod);
  const median = computePercentile(sortedValues, 0.5, quartileMethod);
  const q3 = computePercentile(sortedValues, 0.75, quartileMethod);
  const iqr = q3 - q1;
  const lowerFence = q1 - iqr * 1.5;
  const upperFence = q3 + iqr * 1.5;
  const visiblePoints = sortedValues.filter((value) => value >= lowerFence && value <= upperFence);
  const outliers = sortedValues.filter((value) => value < lowerFence || value > upperFence);
  const lowerWhisker = visiblePoints[0] ?? sortedValues[0] ?? 0;
  const upperWhisker = visiblePoints[visiblePoints.length - 1] ?? sortedValues[sortedValues.length - 1] ?? 0;
  const mean = sortedValues.reduce((sum, value) => sum + value, 0) / sortedValues.length;

  return {
    lowerFence,
    lowerWhisker,
    max: sortedValues[sortedValues.length - 1] ?? 0,
    mean,
    median,
    min: sortedValues[0] ?? 0,
    outliers,
    q1,
    q3,
    upperFence,
    upperWhisker,
    visiblePoints
  };
}

export function resolveHierarchyNodeColor(chart: XlsxChart, node: HierarchyNode<ChartHierarchyDatum>) {
  const lineage = node.ancestors().reverse();
  const topLevel = lineage[1];
  const baseIndex = topLevel?.data.colorIndex ?? 0;
  const baseColor = chartSeriesColor(chart, baseIndex);
  const depth = Math.max(0, node.depth - 1);
  return depth === 0 ? baseColor : lightenColor(baseColor, clamp(depth * 0.16, 0, 0.5));
}

export function resolveTreemapNodeColor(chart: XlsxChart, node: HierarchyNode<ChartHierarchyDatum>) {
  const lineage = node.ancestors().reverse();
  const topLevel = lineage[1];
  return chartSeriesColor(chart, topLevel?.data.colorIndex ?? 0);
}

export function excelTreemapTile(node: HierarchyRectangularNode<ChartHierarchyDatum>, x0: number, y0: number, x1: number, y1: number) {
  if (!node.children || node.children.length === 0) {
    return;
  }
  if (node.depth === 0) {
    treemapDice(node, x0, y0, x1, y1);
    node.children.forEach((child) => excelTreemapTile(child, child.x0, child.y0, child.x1, child.y1));
    return;
  }
  treemapBinary(node, x0, y0, x1, y1);
  node.children.forEach((child) => excelTreemapTile(child, child.x0, child.y0, child.x1, child.y1));
}
