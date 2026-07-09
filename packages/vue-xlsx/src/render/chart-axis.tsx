/** @jsxImportSource vue */
import type { XlsxChart, XlsxChartAxis } from "@extend-ai/xlsx-core";
import {
  getCategoryLabels,
  lightenColor,
  normalizeCategoryLabel,
  resolveChartAxisTextColor,
  resolveSurfacePlotRect,
} from "./chart-renderer";
import type { ChartLayout, ChartRendererPalette, PlotRect } from "./chart-renderer";

export function formatTickValue(value: number) {
  if (Math.abs(value) >= 1000) {
    return value.toLocaleString(undefined, { maximumFractionDigits: 0 });
  }
  return value.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

export function formatPercentTickValue(value: number) {
  return `${formatTickValue(value)}%`;
}

export function buildNiceStep(minValue: number, maxValue: number, preferredTicks = 5) {
  const span = Math.max(1e-6, maxValue - minValue);
  const roughStep = span / Math.max(1, preferredTicks);
  const magnitude = 10 ** Math.floor(Math.log10(roughStep));
  const normalized = roughStep / magnitude;
  const niceNormalized = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  return niceNormalized * magnitude;
}

export function buildNumericTickValues(minValue: number, maxValue: number, majorUnit?: number) {
  if (!Number.isFinite(minValue) || !Number.isFinite(maxValue) || maxValue <= minValue) {
    return [];
  }
  const step = typeof majorUnit === "number" && majorUnit > 0
    ? majorUnit
    : buildNiceStep(minValue, maxValue, 5);
  if (!Number.isFinite(step) || step <= 0) {
    return [];
  }
  const start = Math.floor(minValue / step) * step;
  const values: number[] = [];
  for (let current = start; current <= maxValue + step * 0.001; current += step) {
    if (current >= minValue - step * 0.001 && current <= maxValue + step * 0.001) {
      values.push(Number(current.toFixed(8)));
    }
  }
  return values;
}

export function resolveNumericAxisDomain(minValue: number, maxValue: number, majorUnit?: number, includeZero = false) {
  if (!Number.isFinite(minValue) || !Number.isFinite(maxValue)) {
    return {
      max: 1,
      min: 0,
      ticks: [0, 1]
    };
  }
  let domainMin = includeZero ? Math.min(0, minValue) : minValue;
  let domainMax = includeZero ? Math.max(0, maxValue) : maxValue;
  if (domainMax <= domainMin) {
    domainMax = domainMin + 1;
  }
  const step = typeof majorUnit === "number" && majorUnit > 0
    ? majorUnit
    : buildNiceStep(domainMin, domainMax, 5);
  const roundedMin = typeof majorUnit === "number" && majorUnit > 0
    ? domainMin
    : Math.floor(domainMin / step) * step;
  const roundedMax = typeof majorUnit === "number" && majorUnit > 0
    ? domainMax
    : Math.ceil(domainMax / step) * step;
  const finalMin = includeZero ? Math.min(0, roundedMin) : roundedMin;
  const finalMax = roundedMax <= finalMin ? finalMin + step : roundedMax;
  const ticks = buildNumericTickValues(finalMin, finalMax, step);
  return {
    max: finalMax,
    min: finalMin,
    ticks: ticks.length > 0 ? ticks : [finalMin, finalMax]
  };
}

export function resolveAxisDomainWithChartOverrides(
  axis: XlsxChartAxis | null | undefined,
  minValue: number,
  maxValue: number,
  includeZero = false
) {
  const hasExplicitMin = typeof axis?.min === "number" && Number.isFinite(axis.min);
  const hasExplicitMax = typeof axis?.max === "number" && Number.isFinite(axis.max);
  const rawMin = hasExplicitMin ? Number(axis?.min) : minValue;
  const rawMax = hasExplicitMax ? Number(axis?.max) : maxValue;
  const domain = resolveNumericAxisDomain(rawMin, rawMax, axis?.majorUnit, includeZero);
  return {
    hasExplicitMax,
    hasExplicitMin,
    majorUnit: axis?.majorUnit,
    max: hasExplicitMax ? Number(axis?.max) : domain.max,
    min: hasExplicitMin ? Number(axis?.min) : domain.min,
    ticks: (hasExplicitMin || hasExplicitMax)
      ? buildNumericTickValues(
          hasExplicitMin ? Number(axis?.min) : domain.min,
          hasExplicitMax ? Number(axis?.max) : domain.max,
          axis?.majorUnit
        )
      : domain.ticks
  };
}

export function renderSurfaceAxes(chart: XlsxChart, layout: ChartLayout) {
  const plot = resolveSurfacePlotRect(chart, layout);
  const categories = getCategoryLabels(chart);
  const seriesLabels = chart.series.map((series, index) => normalizeCategoryLabel(series.name) || `Q${index + 1}`);
  const labelColor = resolveChartAxisTextColor(chart);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? "#888888";
  const rowCount = Math.max(1, seriesLabels.length);
  const columnCount = Math.max(1, categories.length);
  const columnPositions = categories.map((_, index) => (
    plot.left + (columnCount <= 1 ? plot.width / 2 : (index / (columnCount - 1)) * plot.width)
  ));
  const rowPositions = seriesLabels.map((_, index) => (
    plot.top + plot.height - (rowCount <= 1 ? plot.height / 2 : (index / (rowCount - 1)) * plot.height)
  ));

  return (
    <g>
      <rect
        fill="none"
        height={plot.height}
        stroke={lightenColor(axisColor, 0.18)}
        stroke-width={0.8}
        width={plot.width}
        x={plot.left}
        y={plot.top}
      />
      {categories.map((label, index) => (
        <text
          key={`surface-x-label-${index}`}
          fill={labelColor}
          font-size={10}
          text-anchor="middle"
          x={columnPositions[index] ?? plot.left}
          y={plot.top + plot.height + 14}
        >
          {label}
        </text>
      ))}
      {seriesLabels.map((label, index) => (
        <text
          key={`surface-y-label-${index}`}
          fill={labelColor}
          font-size={10}
          text-anchor="start"
          x={plot.left + plot.width + 8}
          y={(rowPositions[index] ?? plot.top) + 3}
        >
          {label}
        </text>
      ))}
    </g>
  );
}

export function renderCartesianAxes(
  chart: XlsxChart,
  palette: ChartRendererPalette,
  plot: PlotRect,
  isHorizontal: boolean,
  categoryLabels: string[],
  categoryPositions: number[],
  valueTicks: number[],
  mapValue: (value: number) => number,
  formatValueTick: (value: number) => string = formatTickValue
) {
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);
  const zeroPosition = mapValue(0);

  return (
    <g>
      {valueTicks.map((tick) => {
        const valuePosition = mapValue(tick);
        if (isHorizontal) {
          return (
            <g key={`grid-v-${tick}`}>
              <line
                stroke={lightenColor(axisColor, 0.7)}
                stroke-width={1}
                x1={valuePosition}
                x2={valuePosition}
                y1={plot.top}
                y2={plot.top + plot.height}
              />
              <text
                fill={labelColor}
                font-size={10}
                text-anchor="middle"
                x={valuePosition}
                y={plot.top + plot.height + 14}
              >
                {formatValueTick(tick)}
              </text>
            </g>
          );
        }
        return (
          <g key={`grid-h-${tick}`}>
            <line
              stroke={lightenColor(axisColor, 0.7)}
              stroke-width={1}
              x1={plot.left}
              x2={plot.left + plot.width}
              y1={valuePosition}
              y2={valuePosition}
            />
            <text
              fill={labelColor}
              font-size={10}
              text-anchor="end"
              x={plot.left - 6}
              y={valuePosition + 3}
            >
              {formatValueTick(tick)}
            </text>
          </g>
        );
      })}
      {categoryPositions.map((position, index) => {
        const label = categoryLabels[index] ?? "";
        if (isHorizontal) {
          return (
            <text
              key={`cat-y-${index}`}
              fill={labelColor}
              font-size={10}
              text-anchor="end"
              x={plot.left - 6}
              y={position + 3}
            >
              {label}
            </text>
          );
        }
        return (
          <text
            key={`cat-x-${index}`}
            fill={labelColor}
            font-size={10}
            text-anchor="middle"
            x={position}
            y={plot.top + plot.height + 14}
          >
            {label}
          </text>
        );
      })}
      {isHorizontal ? (
        <>
          <line
            stroke={axisColor}
            stroke-width={1.2}
            x1={plot.left}
            x2={plot.left + plot.width}
            y1={plot.top + plot.height}
            y2={plot.top + plot.height}
          />
          <line
            stroke={axisColor}
            stroke-width={1.2}
            x1={zeroPosition}
            x2={zeroPosition}
            y1={plot.top}
            y2={plot.top + plot.height}
          />
        </>
      ) : (
        <>
          <line
            stroke={axisColor}
            stroke-width={1.2}
            x1={plot.left}
            x2={plot.left}
            y1={plot.top}
            y2={plot.top + plot.height}
          />
          <line
            stroke={axisColor}
            stroke-width={1.2}
            x1={plot.left}
            x2={plot.left + plot.width}
            y1={zeroPosition}
            y2={zeroPosition}
          />
        </>
      )}
    </g>
  );
}
