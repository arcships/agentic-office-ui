/** @jsxImportSource vue */
import type { XlsxChart, XlsxChartElementSelection } from "@extend-ai/xlsx-core";
import type { ChartElementDataOptions } from "./chart-types";

export function chartElementDataProps(seriesIndex: number, pointIndex?: number, options?: ChartElementDataOptions) {
  return {
    "data-xlsx-chart-point-index": typeof pointIndex === "number" ? String(pointIndex) : undefined,
    "data-xlsx-chart-selection-mode": options?.selectionMode,
    "data-xlsx-chart-series-index": String(seriesIndex),
    style: {
      cursor: "pointer",
      pointerEvents: "all"
    } as Record<string, string | number>
  };
}

export function barChartElementDataProps(seriesIndex: number, pointIndex: number) {
  return chartElementDataProps(seriesIndex, pointIndex, { selectionMode: "seriesFirst" });
}

export function resolveChartElementTarget(target: EventTarget | null) {
  if (!(target instanceof Element)) {
    return null;
  }

  return target.closest("[data-xlsx-chart-series-index]") as HTMLElement | SVGElement | null;
}

export function resolveChartSelectionFromTarget(
  chart: XlsxChart,
  target: EventTarget | null,
  selectedChartElement: XlsxChartElementSelection | null | undefined
): XlsxChartElementSelection | null {
  const element = resolveChartElementTarget(target);
  if (!element) {
    return selectedChartElement?.chartId === chart.id
      ? { chartId: chart.id, kind: "chart" }
      : { chartId: chart.id, kind: "chart" };
  }

  const seriesIndex = Number(element.dataset.xlsxChartSeriesIndex);
  if (!Number.isInteger(seriesIndex) || seriesIndex < 0 || seriesIndex >= chart.series.length) {
    return null;
  }

  const series = chart.series[seriesIndex];
  if (!series) {
    return null;
  }

  const rawPointIndex = element.dataset.xlsxChartPointIndex;
  const rawElementKind = element.dataset.xlsxChartElementKind;
  const selectionMode = element.dataset.xlsxChartSelectionMode;
  const pointIndex = rawPointIndex == null || rawPointIndex === "" ? null : Number(rawPointIndex);
  const hasPoint = pointIndex != null && Number.isInteger(pointIndex) && pointIndex >= 0;
  const sameSelectedSeries = selectedChartElement?.chartId === chart.id
    && selectedChartElement.kind !== "chart"
    && selectedChartElement.seriesIndex === seriesIndex;
  if (hasPoint && (selectionMode !== "seriesFirst" || sameSelectedSeries)) {
    return {
      chartId: chart.id,
      kind: "point",
      pointIndex,
      seriesId: series.id,
      seriesIndex
    };
  }

  return {
    chartId: chart.id,
    kind: rawElementKind === "legendEntry" ? "legendEntry" : "series",
    seriesId: series.id,
    seriesIndex
  };
}

export function isSelectedChartSeries(
  selectedChartElement: XlsxChartElementSelection | null | undefined,
  chartId: string,
  seriesIndex: number
) {
  return selectedChartElement?.chartId === chartId
    && selectedChartElement.kind !== "chart"
    && selectedChartElement.seriesIndex === seriesIndex;
}

export function isSelectedChartPoint(
  selectedChartElement: XlsxChartElementSelection | null | undefined,
  chartId: string,
  seriesIndex: number,
  pointIndex: number
) {
  return selectedChartElement?.chartId === chartId
    && selectedChartElement.kind === "point"
    && selectedChartElement.seriesIndex === seriesIndex
    && selectedChartElement.pointIndex === pointIndex;
}

export function isSelectedChartPointOrSeries(
  selectedChartElement: XlsxChartElementSelection | null | undefined,
  chartId: string,
  seriesIndex: number,
  pointIndex: number
) {
  return isSelectedChartPoint(selectedChartElement, chartId, seriesIndex, pointIndex)
    || (
      isSelectedChartSeries(selectedChartElement, chartId, seriesIndex)
      && selectedChartElement?.kind !== "point"
    );
}

export function renderSelectionRectHandles(
  key: string,
  left: number,
  top: number,
  width: number,
  height: number,
  color = "#64748b"
) {
  return renderSelectionPointHandles(
    key,
    [
      { x: left, y: top },
      { x: left + width, y: top },
      { x: left, y: top + height },
      { x: left + width, y: top + height }
    ],
    color
  );
}

export function renderSelectionPointHandles(
  key: string,
  points: Array<{ x: number; y: number }>,
  color = "#64748b"
) {
  const radius = 3;

  return (
    <g key={key} pointer-events="none">
      {points.map((point, index) => (
        <circle
          cx={point.x}
          cy={point.y}
          fill="#ffffff"
          key={`${key}-handle-${index}`}
          r={radius}
          stroke={color}
          stroke-width={1.1}
          style={{ filter: "drop-shadow(0 1px 2px rgba(15, 23, 42, 0.18))" }}
          vector-effect="non-scaling-stroke"
        />
      ))}
    </g>
  );
}
