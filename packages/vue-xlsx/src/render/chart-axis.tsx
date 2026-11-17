/** @jsxImportSource vue */
import { scaleLinear, scaleLog, scaleTime } from "d3-scale";
import type { XlsxChart, XlsxChartAxis } from "@arcships/xlsx-core";
import {
  excelSerialToDate,
  formatExcelDateSerial,
  getCategoryLabels,
  isLikelyDateFormatCode,
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

export function resolveAxisLogBase(axis: XlsxChartAxis | null | undefined) {
  const base = axis?.logBase;
  return typeof base === "number" && Number.isFinite(base) && base > 1 ? base : null;
}

function normalizePowerValue(value: number) {
  return Number(value.toPrecision(12));
}

export function buildLogTickValues(minValue: number, maxValue: number, base: number, maxTicks = 8) {
  if (
    !Number.isFinite(minValue)
    || !Number.isFinite(maxValue)
    || !Number.isFinite(base)
    || base <= 1
    || minValue <= 0
    || maxValue <= minValue
  ) {
    return [];
  }
  const logMin = Math.log(minValue) / Math.log(base);
  const logMax = Math.log(maxValue) / Math.log(base);
  const startExponent = Math.ceil(logMin - 1e-9);
  const endExponent = Math.floor(logMax + 1e-9);
  if (endExponent < startExponent) {
    return [];
  }
  const exponentCount = endExponent - startExponent + 1;
  const step = exponentCount > maxTicks ? Math.ceil(exponentCount / maxTicks) : 1;
  const values: number[] = [];
  for (let exponent = startExponent; exponent <= endExponent; exponent += step) {
    values.push(normalizePowerValue(base ** exponent));
  }
  return values;
}

export function resolveLogAxisDomain(
  axis: XlsxChartAxis | null | undefined,
  minValue: number,
  maxValue: number,
  fallbackMin = 1
) {
  const base = resolveAxisLogBase(axis) ?? 10;
  const hasExplicitMin = typeof axis?.min === "number" && Number.isFinite(axis.min) && axis.min > 0;
  const hasExplicitMax = typeof axis?.max === "number" && Number.isFinite(axis.max) && axis.max > 0;
  const rawMin = hasExplicitMin
    ? Number(axis?.min)
    : (Number.isFinite(minValue) && minValue > 0 ? minValue : fallbackMin);
  const rawMax = hasExplicitMax
    ? Number(axis?.max)
    : (Number.isFinite(maxValue) && maxValue > rawMin ? maxValue : rawMin * base);
  const domainMin = hasExplicitMin
    ? rawMin
    : normalizePowerValue(base ** Math.floor(Math.log(rawMin) / Math.log(base) + 1e-9));
  let domainMax = hasExplicitMax
    ? rawMax
    : normalizePowerValue(base ** Math.ceil(Math.log(rawMax) / Math.log(base) - 1e-9));
  if (domainMax <= domainMin) {
    domainMax = normalizePowerValue(domainMin * base);
  }
  const ticks = buildLogTickValues(domainMin, domainMax, base);
  return {
    max: domainMax,
    min: domainMin,
    ticks: ticks.length >= 2 ? ticks : [domainMin, domainMax]
  };
}

export function createValueScale(
  axis: XlsxChartAxis | null | undefined,
  minValue: number,
  maxValue: number,
  range: [number, number]
): (value: number) => number {
  const logBase = resolveAxisLogBase(axis);
  const scale = logBase
    ? scaleLog().base(logBase).domain([minValue, maxValue]).range(range)
    : scaleLinear().domain([minValue, maxValue]).range(range);
  return (value: number) => scale(value);
}

export function resolveDateAxisDomain(
  axis: XlsxChartAxis | null | undefined,
  values: Array<number | null | undefined>
) {
  const numericValues = values.filter((value): value is number => (
    typeof value === "number" && Number.isFinite(value)
  ));
  const explicitMin = typeof axis?.min === "number" && Number.isFinite(axis.min) ? axis.min : null;
  const explicitMax = typeof axis?.max === "number" && Number.isFinite(axis.max) ? axis.max : null;
  const min = explicitMin ?? (numericValues.length > 0 ? Math.min(...numericValues) : 0);
  const rawMax = explicitMax ?? (numericValues.length > 0 ? Math.max(...numericValues) : min + 1);
  const max = rawMax <= min ? min + 1 : rawMax;
  return { max, min };
}

export function createDateScale(
  minSerial: number,
  maxSerial: number,
  range: [number, number],
  use1904 = false
): (serial: number) => number {
  const scale = scaleTime()
    .domain([
      excelSerialToDate(minSerial, use1904).getTime(),
      excelSerialToDate(maxSerial, use1904).getTime()
    ])
    .range(range);
  return (serial: number) => scale(excelSerialToDate(serial, use1904).getTime());
}

const MILLISECONDS_PER_DAY = 86_400_000;

function excelSerialFromDate(date: Date, use1904: boolean) {
  const baseUtc = use1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return (date.getTime() - baseUtc) / MILLISECONDS_PER_DAY;
}

function resolveDateTickStep(spanDays: number, unit: "days" | "months" | "years", preferredStep: number | undefined) {
  const candidates = unit === "days"
    ? [1, 2, 7, 14, 30]
    : unit === "months"
      ? [1, 2, 3, 6]
      : [1, 2, 5, 10];
  if (typeof preferredStep === "number" && preferredStep > 0) {
    return preferredStep;
  }
  const spanUnits = unit === "days" ? spanDays : unit === "months" ? spanDays / 30 : spanDays / 365;
  return candidates.find((candidate) => spanUnits / candidate <= 8) ?? candidates[candidates.length - 1] ?? 1;
}

export function buildDateAxisTicks(
  axis: XlsxChartAxis | null | undefined,
  minSerial: number,
  maxSerial: number,
  use1904 = false
) {
  if (!Number.isFinite(minSerial) || !Number.isFinite(maxSerial) || maxSerial <= minSerial) {
    return [] as Array<{ label: string; serial: number; time: number }>;
  }
  const spanDays = maxSerial - minSerial;
  const rawUnit = axis?.majorTimeUnit?.toLowerCase();
  const unit = rawUnit === "days" || rawUnit === "months" || rawUnit === "years"
    ? rawUnit
    : spanDays <= 45
      ? "days"
      : spanDays <= 900
        ? "months"
        : "years";
  const step = resolveDateTickStep(spanDays, unit, axis?.majorUnit);
  const startDate = excelSerialToDate(minSerial, use1904);
  let cursor: Date;
  if (unit === "days") {
    cursor = new Date(startDate.getTime());
  } else if (unit === "months") {
    cursor = new Date(Date.UTC(startDate.getUTCFullYear(), startDate.getUTCMonth(), 1));
    if (cursor.getTime() < startDate.getTime()) {
      cursor = new Date(Date.UTC(startDate.getUTCFullYear(), startDate.getUTCMonth() + 1, 1));
    }
  } else {
    cursor = new Date(Date.UTC(startDate.getUTCFullYear() + (startDate.getUTCMonth() > 0 || startDate.getUTCDate() > 1 ? 1 : 0), 0, 1));
  }

  const ticks: Array<{ label: string; serial: number; time: number }> = [];
  const formatCode = axis?.numberFormat?.formatCode;
  for (let guard = 0; guard < 64; guard += 1) {
    const serial = excelSerialFromDate(cursor, use1904);
    if (serial > maxSerial + 1e-6) {
      break;
    }
    if (serial >= minSerial - 1e-6) {
      const label = formatCode && isLikelyDateFormatCode(formatCode)
        ? formatExcelDateSerial(serial, formatCode, use1904)
        : (() => {
            if (unit === "years") {
              return String(cursor.getUTCFullYear());
            }
            if (unit === "months") {
              return cursor.toLocaleDateString(undefined, { month: "short", year: "numeric", timeZone: "UTC" });
            }
            return cursor.toLocaleDateString(undefined, { day: "numeric", month: "numeric", timeZone: "UTC" });
          })();
      ticks.push({ label, serial, time: cursor.getTime() });
    }
    if (unit === "days") {
      cursor = new Date(cursor.getTime() + MILLISECONDS_PER_DAY * step);
    } else if (unit === "months") {
      cursor = new Date(Date.UTC(cursor.getUTCFullYear(), cursor.getUTCMonth() + step, 1));
    } else {
      cursor = new Date(Date.UTC(cursor.getUTCFullYear() + step, 0, 1));
    }
  }
  return ticks;
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
  const rawZeroPosition = mapValue(0);
  const zeroPosition = Number.isFinite(rawZeroPosition)
    ? rawZeroPosition
    : (isHorizontal ? plot.left : plot.top + plot.height);

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
