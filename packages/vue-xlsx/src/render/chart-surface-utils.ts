/** @jsxImportSource vue */
import { line as d3Line, curveCatmullRom, curveLinear } from "d3-shape";
import type { XlsxChart } from "@arcships/xlsx-core";
import { clamp, mixRgbColor, lightenColor, darkenColor, normalizeRendererHexColor, safeNumber } from "./chart-shared";
import { getCategoryLabels } from "./chart-data";
import { formatTickValue } from "./chart-axis";
import type { ChartRendererPalette, ChartLayout, SurfaceDomain, LegendItem, SurfacePathPoint, SurfaceContourSegment } from "./chart-types";

export function resolveSurfaceBaseColor(chart: XlsxChart, palette: ChartRendererPalette) {
  return chart.chartColorPalette?.[0]
    ?? chart.series[0]?.color
    ?? chart.series[0]?.lineColor
    ?? chart.axisLineColor
    ?? chart.textColor
    ?? palette.text;
}

export function normalizeBuiltinSurfaceStyleId(styleId: number | undefined) {
  if (typeof styleId !== "number" || !Number.isFinite(styleId)) {
    return null;
  }
  return styleId >= 100 ? styleId - 100 : styleId;
}

export function hasExplicitSurfaceBaseColor(chart: XlsxChart) {
  const primarySeriesColor = normalizeRendererHexColor(chart.series[0]?.color ?? chart.series[0]?.lineColor);
  if (!primarySeriesColor) {
    return null;
  }
  const paletteColor = normalizeRendererHexColor(chart.chartColorPalette?.[0]);
  return paletteColor && paletteColor === primarySeriesColor ? null : primarySeriesColor;
}

export function buildMonochromeSurfacePalette(baseColor: string, count: number) {
  if (count <= 3) {
    return [
      lightenColor(baseColor, 0.22),
      baseColor,
      darkenColor(baseColor, 0.2)
    ];
  }
  return [
    lightenColor(baseColor, 0.3),
    lightenColor(baseColor, 0.14),
    baseColor,
    darkenColor(baseColor, 0.1),
    darkenColor(baseColor, 0.22)
  ];
}

export function getBuiltinSurfacePalette(chart: XlsxChart) {
  const normalized = normalizeBuiltinSurfaceStyleId(chart.chartStyleId);
  const explicitBaseColor = hasExplicitSurfaceBaseColor(chart);
  if (normalized === 26) {
    return buildMonochromeSurfacePalette(explicitBaseColor ?? "#ff006e", 3);
  }
  if (normalized === 34 && explicitBaseColor) {
    return buildMonochromeSurfacePalette(explicitBaseColor, 3);
  }
  if (normalized === 34 || (chart.wireframe === true && normalized == null)) {
    return ["#5b9bd5", "#ed7d31", "#a5a5a5"];
  }
  if (normalized === 35 || normalized === 36 || (chart.wireframe !== true && normalized == null)) {
    return ["#2f5597", "#4472c4", "#5b9bd5", "#8faadc", "#d9e2f3"];
  }
  return null;
}

export function shouldPreferBuiltinSurfacePalette(chart: XlsxChart) {
  const normalized = normalizeBuiltinSurfaceStyleId(chart.chartStyleId);
  const rawChartType = chart.raw && typeof chart.raw === "object" && typeof (chart.raw as Record<string, unknown>).xmlChartType === "string"
    ? String((chart.raw as Record<string, unknown>).xmlChartType)
    : "";
  return (
    (rawChartType === "surfaceChart" || rawChartType === "surface3DChart")
    && (normalized === 26 || normalized === 34 || normalized === 35 || normalized === 36)
  );
}

export function isContourSurfaceChart(chart: XlsxChart) {
  const rawChartType = chart.raw && typeof chart.raw === "object" && typeof (chart.raw as Record<string, unknown>).xmlChartType === "string"
    ? (chart.raw as Record<string, unknown>).xmlChartType
    : "";
  if (rawChartType === "surfaceChart") {
    return true;
  }
  if (rawChartType === "surface3DChart") {
    return false;
  }
  return chart.chartType === "Surface" && chart.is3d !== true;
}

export function getSurfaceBandCount(chart: XlsxChart) {
  const raw = chart.raw && typeof chart.raw === "object" ? chart.raw as Record<string, unknown> : null;
  const explicitBandCount = typeof raw?.bandFormatCount === "number" && Number.isFinite(raw.bandFormatCount)
    ? raw.bandFormatCount
    : null;
  if (explicitBandCount != null && explicitBandCount > 0 && !(isContourSurfaceChart(chart) && chart.wireframe !== true)) {
    return explicitBandCount;
  }
  const builtinPalette = getBuiltinSurfacePalette(chart);
  if (shouldPreferBuiltinSurfacePalette(chart) && builtinPalette && builtinPalette.length > 0) {
    return builtinPalette.length;
  }
  if (chart.chartColorPalette && chart.chartColorPalette.length > 1) {
    return chart.chartColorPalette.length;
  }
  if (builtinPalette && builtinPalette.length > 0) {
    return builtinPalette.length;
  }
  return chart.wireframe ? 3 : 5;
}

export function getSurfaceColorStops(chart: XlsxChart, palette: ChartRendererPalette) {
  const builtinPalette = getBuiltinSurfacePalette(chart);
  if (shouldPreferBuiltinSurfacePalette(chart) && builtinPalette && builtinPalette.length >= 2) {
    return builtinPalette;
  }
  const explicitStops = (chart.chartColorPalette ?? []).filter((value): value is string => typeof value === "string" && value.length > 0);
  if (explicitStops.length >= 2) {
    return explicitStops;
  }
  if (builtinPalette && builtinPalette.length >= 2) {
    return builtinPalette;
  }
  const baseColor = resolveSurfaceBaseColor(chart, palette);
  return [
    darkenColor(baseColor, 0.42),
    darkenColor(baseColor, 0.24),
    baseColor,
    lightenColor(baseColor, 0.18),
    lightenColor(baseColor, 0.34),
    lightenColor(baseColor, 0.5)
  ];
}

export function resolveSurfaceBandPaletteColor(chart: XlsxChart, palette: ChartRendererPalette, domain: SurfaceDomain, value: number) {
  const stops = getSurfaceColorStops(chart, palette);
  if (stops.length === 0) {
    return resolveSurfaceBaseColor(chart, palette);
  }
  const bandIndex = resolveSurfaceBandIndex(domain, value);
  return stops[Math.min(stops.length - 1, Math.max(0, bandIndex))] ?? stops[stops.length - 1] ?? resolveSurfaceBaseColor(chart, palette);
}

export function getSurfaceDomain(chart: XlsxChart): SurfaceDomain | null {
  const numericValues = chart.series.flatMap((series) => (
    series.values
      .map((value) => safeNumber(value))
      .filter((value): value is number => value != null)
  ));
  if (numericValues.length === 0) {
    return null;
  }
  const explicitMin = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min)
    ? chart.valueAxis.min
    : null;
  const explicitMax = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max)
    ? chart.valueAxis.max
    : null;
  const rawMin = Math.min(...numericValues);
  const rawMax = Math.max(...numericValues);
  const bandCount = Math.max(1, getSurfaceBandCount(chart));
  const spanBase = Math.max(1e-6, rawMax - Math.min(0, rawMin));
  const roughStep = spanBase / bandCount;
  const magnitude = 10 ** Math.floor(Math.log10(Math.max(roughStep, 1e-6)));
  const normalized = roughStep / magnitude;
  const niceNormalized = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  const step = typeof chart.valueAxis?.majorUnit === "number" && chart.valueAxis.majorUnit > 0
    ? chart.valueAxis.majorUnit
    : niceNormalized * magnitude;
  const minValue = explicitMin ?? (rawMin >= 0 ? 0 : Math.floor(rawMin / step) * step);
  const maxValue = explicitMax ?? Math.ceil(rawMax / step) * step;
  const safeMax = maxValue <= minValue ? minValue + step : maxValue;
  const ticks: number[] = [];
  for (let current = minValue; current <= safeMax + step * 0.001; current += step) {
    ticks.push(Number(current.toFixed(8)));
  }
  return {
    maxValue,
    minValue,
    safeMax,
    ticks: ticks.length >= 2 ? ticks : [minValue, safeMax]
  };
}

export function resolveSurfaceColor(chart: XlsxChart, palette: ChartRendererPalette, ratio: number) {
  const stops = getSurfaceColorStops(chart, palette);
  if (stops.length === 0) {
    return resolveSurfaceBaseColor(chart, palette);
  }
  if (stops.length === 1) {
    return stops[0];
  }
  const clamped = clamp(ratio, 0, 1) * (stops.length - 1);
  const lowerIndex = Math.floor(clamped);
  const upperIndex = Math.min(stops.length - 1, lowerIndex + 1);
  const mixRatio = clamped - lowerIndex;
  return mixRgbColor(stops[lowerIndex] ?? stops[0], stops[upperIndex] ?? stops[stops.length - 1], mixRatio);
}

export function resolveSurfaceBandColor(chart: XlsxChart, palette: ChartRendererPalette, domain: SurfaceDomain, value: number) {
  return resolveSurfaceBandPaletteColor(chart, palette, domain, value);
}

export function resolveSurfaceBandIndex(domain: SurfaceDomain, value: number) {
  const ticks = domain.ticks;
  for (let index = 0; index < ticks.length - 1; index += 1) {
    const end = ticks[index + 1] ?? domain.safeMax;
    if (value <= end || index === ticks.length - 2) {
      return index;
    }
  }
  return Math.max(0, ticks.length - 2);
}

export function buildSurfaceSmoothPath(points: SurfacePathPoint[], smooth: boolean) {
  if (points.length < 2) {
    return "";
  }
  return d3Line<SurfacePathPoint>()
    .x((point) => point.x)
    .y((point) => point.y)
    .curve(points.length >= 3 && smooth ? curveCatmullRom.alpha(0.5) : curveLinear)(points) ?? "";
}

export function splitSurfacePointRuns<T extends { hasValue: boolean; x: number; y: number }>(points: T[]) {
  const runs: SurfacePathPoint[][] = [];
  let currentRun: SurfacePathPoint[] = [];
  points.forEach((point) => {
    if (!point.hasValue) {
      if (currentRun.length >= 2) {
        runs.push(currentRun);
      }
      currentRun = [];
      return;
    }
    const nextPoint = { x: point.x, y: point.y };
    const previous = currentRun[currentRun.length - 1];
    if (!previous || Math.abs(previous.x - nextPoint.x) > 0.01 || Math.abs(previous.y - nextPoint.y) > 0.01) {
      currentRun.push(nextPoint);
    }
  });
  if (currentRun.length >= 2) {
    runs.push(currentRun);
  }
  return runs;
}

export function connectSurfaceContourSegments(segments: SurfaceContourSegment[]) {
  const remaining = [...segments];
  const epsilon = 0.75;
  const within = (left: SurfacePathPoint, right: SurfacePathPoint) => (
    Math.abs(left.x - right.x) <= epsilon && Math.abs(left.y - right.y) <= epsilon
  );
  const paths: SurfacePathPoint[][] = [];

  while (remaining.length > 0) {
    const seed = remaining.shift();
    if (!seed) {
      break;
    }
    const chain: SurfacePathPoint[] = [seed.start, seed.end];
    let changed = true;
    while (changed) {
      changed = false;
      for (let index = remaining.length - 1; index >= 0; index -= 1) {
        const candidate = remaining[index];
        if (!candidate) {
          continue;
        }
        if (within(chain[chain.length - 1] ?? candidate.start, candidate.start)) {
          chain.push(candidate.end);
        } else if (within(chain[chain.length - 1] ?? candidate.start, candidate.end)) {
          chain.push(candidate.start);
        } else if (within(chain[0] ?? candidate.start, candidate.end)) {
          chain.unshift(candidate.start);
        } else if (within(chain[0] ?? candidate.start, candidate.start)) {
          chain.unshift(candidate.end);
        } else {
          continue;
        }
        remaining.splice(index, 1);
        changed = true;
      }
    }
    const deduped = chain.filter((point, index) => {
      if (index === 0) {
        return true;
      }
      const previous = chain[index - 1];
      return !previous || !within(previous, point);
    });
    if (deduped.length >= 2) {
      paths.push(deduped);
    }
  }

  return paths;
}

export function getSurfaceWireframePalette(chart: XlsxChart, palette: ChartRendererPalette) {
  const raw = chart.raw && typeof chart.raw === "object" ? chart.raw as Record<string, unknown> : null;
  const explicit = Array.isArray(raw?.bandFormatLineColors)
    ? raw.bandFormatLineColors.filter((color): color is string => typeof color === "string" && color.length > 0)
    : [];
  if (explicit.length > 0) {
    return explicit;
  }
  return getSurfaceColorStops(chart, palette);
}

export function resolveSurfaceWireframeColor(chart: XlsxChart, palette: ChartRendererPalette, domain: SurfaceDomain, value: number) {
  const stops = getSurfaceWireframePalette(chart, palette);
  if (stops.length === 0) {
    return resolveSurfaceBandPaletteColor(chart, palette, domain, value);
  }
  const bandIndex = resolveSurfaceBandIndex(domain, value);
  return stops[Math.min(stops.length - 1, Math.max(0, bandIndex))] ?? resolveSurfaceBandPaletteColor(chart, palette, domain, value);
}

export function resolveSurfacePlotRect(chart: XlsxChart, layout: ChartLayout) {
  if (!isContourSurfaceChart(chart)) {
    return layout.plot;
  }
  const columnCount = Math.max(1, getCategoryLabels(chart).length);
  const rowCount = Math.max(1, chart.series.length);
  const targetAspect = Math.max(0.72, columnCount / Math.max(1, rowCount));
  const widthScale = chart.wireframe ? 0.78 : 0.84;
  const heightScale = chart.wireframe ? 0.72 : 0.8;
  let width = layout.plot.width * widthScale;
  let height = layout.plot.height * heightScale;
  if (width / Math.max(1e-6, height) > targetAspect) {
    width = height * targetAspect;
  } else {
    height = width / Math.max(1e-6, targetAspect);
  }
  return {
    height,
    left: layout.plot.left + (layout.plot.width - width) / 2,
    top: layout.plot.top + (layout.plot.height - height) / 2,
    width
  };
}

export function buildSurfaceLegendItems(chart: XlsxChart, palette: ChartRendererPalette) {
  const domain = getSurfaceDomain(chart);
  if (!domain) {
    return [];
  }
  const items: LegendItem[] = [];
  for (let index = 0; index < domain.ticks.length - 1; index += 1) {
    const start = domain.ticks[index] ?? domain.minValue;
    const end = domain.ticks[index + 1] ?? domain.safeMax;
    const midpoint = start + (end - start) * 0.5;
    const ratio = (midpoint - domain.minValue) / Math.max(1e-6, domain.safeMax - domain.minValue);
    items.push({
      color: resolveSurfaceColor(chart, palette, ratio),
      label: `${formatTickValue(start)}-${formatTickValue(end)}`
    });
  }
  return items.reverse();
}
