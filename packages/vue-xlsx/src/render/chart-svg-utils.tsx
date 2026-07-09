/** @jsxImportSource vue */
import { symbol as d3Symbol, symbolCircle, symbolCross, symbolDiamond, symbolSquare, symbolStar, symbolTriangle } from "d3-shape";
import type { XlsxChart } from "@extend-ai/xlsx-core";
import { clamp, lightenColor, normalizeRendererHexColor, DEFAULT_CHART_TEXT_COLOR, buildChartFontFamily, truncateSvgText, resolveChartMutedTextColor } from "./chart-shared";
import { chartSeriesColor } from "./chart-data";
import type { ChartRendererPalette, ChartLayout, LegendItem } from "./chart-types";

export const TWO_PI = Math.PI * 2;

// ---- toSvgNumber ----

export function toSvgNumber(value: number) {
  return Number(value.toFixed(3));
}

// ---- Stock chart ----

export function indexByName(series: XlsxChart["series"], matcher: RegExp) {
  const index = series.findIndex((entry) => matcher.test((entry.name ?? "").toLowerCase()));
  return index >= 0 ? index : null;
}

export function resolveStockRoleIndices(chart: XlsxChart) {
  const roles = {
    close: indexByName(chart.series, /close|last|end/),
    high: indexByName(chart.series, /high|max|top/),
    low: indexByName(chart.series, /low|min|bottom/),
    open: indexByName(chart.series, /open|start/),
    volume: indexByName(chart.series, /volume|vol/)
  };
  const usedIndices = new Set<number>(Object.values(roles).filter((value): value is number => value != null));
  const remainingIndices = chart.series.map((_, index) => index).filter((index) => !usedIndices.has(index));
  const canonicalOrder = chart.series.length >= 4
    ? (["open", "high", "low", "close"] as const)
    : (["high", "low", "close"] as const);
  canonicalOrder.forEach((role) => {
    if (roles[role] == null) {
      const nextIndex = remainingIndices.shift();
      if (nextIndex != null) {
        roles[role] = nextIndex;
      }
    }
  });
  return roles;
}

export function resolveStockPalette(chart: XlsxChart, axisColor: string) {
  const raw = chart.raw && typeof chart.raw === "object" ? chart.raw as Record<string, unknown> : null;
  const highLowLines = raw?.highLowLines && typeof raw.highLowLines === "object"
    ? raw.highLowLines as Record<string, unknown>
    : null;
  const shapeProperties = highLowLines?.shapeProperties && typeof highLowLines.shapeProperties === "object"
    ? highLowLines.shapeProperties as Record<string, unknown>
    : null;
  const lineColor = normalizeRendererHexColor(shapeProperties?.lineColorHex) ?? axisColor ?? "#333333";
  const chartStyleId = typeof chart.chartStyleId === "number" ? chart.chartStyleId : null;
  const closeAccent = chartStyleId != null && chartStyleId >= 128 ? "#c0504d" : lineColor;
  const lowAccent = chartStyleId != null && chartStyleId >= 128 ? "#d9a3a0" : lightenColor(lineColor, 0.45);
  return {
    closeAccent,
    downFill: lightenColor(lineColor, 0.4),
    lineColor,
    lowAccent,
    openAccent: lightenColor(lineColor, 0.22),
    upFill: "#ffffff",
    volumeFill: lightenColor(lineColor, 0.26)
  };
}

export function buildStockLegendItems(chart: XlsxChart, palette: ChartRendererPalette): LegendItem[] {
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const roles = resolveStockRoleIndices(chart);
  const stockPalette = resolveStockPalette(chart, axisColor);
  return chart.series.map((series, index) => {
    let color = stockPalette.lineColor;
    if (roles.volume === index) {
      color = stockPalette.volumeFill;
    } else if (roles.open === index) {
      color = stockPalette.openAccent;
    } else if (roles.low === index) {
      color = stockPalette.lowAccent;
    } else if (roles.close === index) {
      color = stockPalette.closeAccent;
    }
    return {
      color,
      label: series.name ?? `Series ${index + 1}`
    };
  });
}

// ---- Marker symbol ----

export function markerSymbolPath(symbol: string, size: number) {
  if (symbol === "none") {
    return "";
  }
  const symbolType = (() => {
    switch (symbol) {
      case "diamond":
        return symbolDiamond;
      case "square":
        return symbolSquare;
      case "triangle":
      case "triangle-up":
        return symbolTriangle;
      case "cross":
      case "plus":
        return symbolCross;
      case "star":
        return symbolStar;
      case "circle":
      default:
        return symbolCircle;
    }
  })();
  return d3Symbol().type(symbolType).size(size * size * Math.PI)() ?? "";
}

// ---- Pie 3D helpers ----

export function normalizePieArc(startAngle: number, endAngle: number) {
  let start = startAngle;
  let end = endAngle;
  while (end < start) {
    end += TWO_PI;
  }
  return { end, start };
}

export function resolvePieFrontSegments(startAngle: number, endAngle: number): Array<[number, number]> {
  const { start, end } = normalizePieArc(startAngle, endAngle);
  const segments: Array<[number, number]> = [];
  const minBand = Math.floor(start / TWO_PI) - 1;
  const maxBand = Math.ceil(end / TWO_PI) + 1;
  for (let band = minBand; band <= maxBand; band += 1) {
    const frontStart = Math.PI / 2 + band * TWO_PI;
    const frontEnd = Math.PI * 1.5 + band * TWO_PI;
    const segmentStart = Math.max(start, frontStart);
    const segmentEnd = Math.min(end, frontEnd);
    if (segmentEnd > segmentStart) {
      segments.push([segmentStart, segmentEnd]);
    }
  }
  return segments;
}

export function pieEllipsePoint(
  centerX: number,
  centerY: number,
  radius: number,
  tilt: number,
  angle: number,
  depth = 0
) {
  return {
    x: centerX + Math.sin(angle) * radius,
    y: centerY - Math.cos(angle) * radius * tilt + depth
  };
}

export function buildPieOuterWallPath(
  centerX: number,
  centerY: number,
  radius: number,
  tilt: number,
  depth: number,
  startAngle: number,
  endAngle: number,
  fullArc = false
) {
  const ry = radius * tilt;
  const segments = fullArc ? [[startAngle, endAngle] as [number, number]] : resolvePieFrontSegments(startAngle, endAngle);
  return segments.map(([segmentStart, segmentEnd]) => {
    const topStart = pieEllipsePoint(centerX, centerY, radius, tilt, segmentStart, 0);
    const topEnd = pieEllipsePoint(centerX, centerY, radius, tilt, segmentEnd, 0);
    const bottomStart = pieEllipsePoint(centerX, centerY, radius, tilt, segmentStart, depth);
    const bottomEnd = pieEllipsePoint(centerX, centerY, radius, tilt, segmentEnd, depth);
    const largeArc = segmentEnd - segmentStart > Math.PI ? 1 : 0;
    return [
      `M ${toSvgNumber(topStart.x)} ${toSvgNumber(topStart.y)}`,
      `A ${toSvgNumber(radius)} ${toSvgNumber(ry)} 0 ${largeArc} 1 ${toSvgNumber(topEnd.x)} ${toSvgNumber(topEnd.y)}`,
      `L ${toSvgNumber(bottomEnd.x)} ${toSvgNumber(bottomEnd.y)}`,
      `A ${toSvgNumber(radius)} ${toSvgNumber(ry)} 0 ${largeArc} 0 ${toSvgNumber(bottomStart.x)} ${toSvgNumber(bottomStart.y)}`,
      "Z"
    ].join(" ");
  });
}

export function isPieFrontFacingAngle(angle: number) {
  const normalized = ((angle % TWO_PI) + TWO_PI) % TWO_PI;
  return normalized > Math.PI / 2 && normalized < Math.PI * 1.5;
}

export function buildPieRadialWallPath(
  centerX: number,
  centerY: number,
  radius: number,
  tilt: number,
  depth: number,
  angle: number,
  forceVisible = false
) {
  if (!forceVisible && !isPieFrontFacingAngle(angle)) {
    return null;
  }
  const topCenter = { x: centerX, y: centerY };
  const bottomCenter = { x: centerX, y: centerY + depth };
  const topEdge = pieEllipsePoint(centerX, centerY, radius, tilt, angle, 0);
  const bottomEdge = pieEllipsePoint(centerX, centerY, radius, tilt, angle, depth);
  return [
    `M ${toSvgNumber(topCenter.x)} ${toSvgNumber(topCenter.y)}`,
    `L ${toSvgNumber(topEdge.x)} ${toSvgNumber(topEdge.y)}`,
    `L ${toSvgNumber(bottomEdge.x)} ${toSvgNumber(bottomEdge.y)}`,
    `L ${toSvgNumber(bottomCenter.x)} ${toSvgNumber(bottomCenter.y)}`,
    "Z"
  ].join(" ");
}

// ---- renderTitle ----

export function renderTitle(chart: XlsxChart, layout: ChartLayout, palette: ChartRendererPalette) {
  if (!chart.title) {
    return null;
  }
  const fontSize = 12;
  const text = truncateSvgText(chart.title, Math.max(40, layout.width - 12), fontSize);
  const baselineY = layout.titleHeight >= 30 ? 19 : 16;
  return (
    <text
      fill={chart.titleColor ?? chart.textColor ?? DEFAULT_CHART_TEXT_COLOR}
      font-family={buildChartFontFamily(chart.titleFontFamily ?? chart.fontFamily)}
      font-size={fontSize}
      font-weight={600}
      text-anchor="middle"
      x={layout.width / 2}
      y={baselineY}
    >
      {text}
    </text>
  );
}

// ---- renderUnsupported ----

export function renderUnsupported(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string) {
  return (
    <g>
      <rect
        fill={lightenColor(chart.chartAreaFillColor ?? palette.surface, 0.02)}
        height={layout.plot.height}
        stroke={lightenColor(chart.axisLineColor ?? palette.border, 0.2)}
        stroke-dasharray="3 3"
        width={layout.plot.width}
        x={layout.plot.left}
        y={layout.plot.top}
      />
      <text
        fill={resolveChartMutedTextColor(chart)}
        font-size={11}
        text-anchor="middle"
        x={layout.plot.left + layout.plot.width / 2}
        y={layout.plot.top + layout.plot.height / 2}
      >
        {`Unsupported chart type: ${chartType}`}
      </text>
    </g>
  );
}
