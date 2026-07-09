import { hierarchy as d3Hierarchy, partition as d3Partition, treemap as d3Treemap, treemapBinary, treemapDice, treemapSquarify } from "d3-hierarchy";
import { geoIdentity, geoMercator, geoNaturalEarth1, geoPath } from "d3-geo";
import { defineComponent, type VNode } from "vue";
import { scaleBand, scaleLinear, scalePoint } from "d3-scale";
import {
  arc as d3Arc,
  area as d3Area,
  curveCatmullRom,
  curveLinear,
  curveLinearClosed,
  line as d3Line,
  pie as d3Pie,
  symbol as d3Symbol,
  symbolCircle,
  symbolCross,
  symbolDiamond,
  symbolSquare,
  symbolStar,
  symbolTriangle
} from "d3-shape";
import type { CurveFactory } from "d3-shape";
import { feature as topojsonFeature } from "topojson-client";
import countiesAlbers10m from "us-atlas/counties-albers-10m.json";
import countries50m from "world-atlas/countries-50m.json";
import { MemoSurfaceChartComposite } from "./surface-regl";
import type { HierarchyNode, HierarchyRectangularNode } from "d3-hierarchy";
import type { Feature, FeatureCollection, Geometry } from "geojson";
import type { XlsxChart, XlsxChartAxis, XlsxChartElementSelection, XlsxChartSeries, XlsxChartTypeGroup, XlsxImageRect } from "@extend-ai/xlsx-core";


import { renderBarChart, renderWaterfallChart, renderFunnelChart, renderBoxWhiskerChart } from "./chart-bar";
import { renderLineOrAreaChart, renderComboChart, renderRadarChart, renderStockChart } from "./chart-line";
import { renderPieChart, renderBarOfPieChart, renderSunburstChart, renderTreemapChart } from "./chart-pie";
import { renderScatterChart, renderBubbleChart } from "./chart-scatter";
import { renderSurfaceChart, renderSurfaceHitOverlay, renderRegionMapChart } from "./chart-surface";
import { buildLayout, getLegendItems, renderLegend } from "./chart-legend";
import { buildNumericTickValues, formatTickValue, renderSurfaceAxes, resolveAxisDomainWithChartOverrides } from "./chart-axis";

export { buildLayout, getLegendItems, renderLegend } from "./chart-legend";
export { buildNiceStep, buildNumericTickValues, formatPercentTickValue, formatTickValue, renderCartesianAxes, renderSurfaceAxes, resolveAxisDomainWithChartOverrides, resolveNumericAxisDomain } from "./chart-axis";

export type ChartRendererPalette = {
  border: string;
  mutedText: string;
  surface: string;
  text: string;
};

export type ChartSvgProps = {
  chart: XlsxChart;
  onChartElementDoubleClick?: (selection: XlsxChartElementSelection, event: MouseEvent) => void;
  onChartElementPointerDown?: (selection: XlsxChartElementSelection, event: PointerEvent) => void;
  palette: ChartRendererPalette;
  rect: XlsxImageRect;
  selectedChartElement?: XlsxChartElementSelection | null;
};

export type LegendItem = {
  color: string;
  label: string;
};

export type PlotRect = {
  height: number;
  left: number;
  top: number;
  width: number;
};

export type ChartLayout = {
  height: number;
  legendItems: LegendItem[];
  legendPosition: string | undefined;
  plot: PlotRect;
  titleHeight: number;
  width: number;
};

export type BarRect = {
  capEnd?: boolean;
  capStart?: boolean;
  categoryIndex: number;
  bottomScale?: number;
  color: string;
  depthOrder?: number;
  depthOffsetX?: number;
  depthOffsetY?: number;
  depthX?: number;
  depthY?: number;
  gradientId?: string;
  height: number;
  invertedNegative?: boolean;
  isHorizontal: boolean;
  key: string;
  left: number;
  shape3d?: string;
  seriesIndex: number;
  stroke: string;
  strokeWidth: number;
  topScale?: number;
  value: number;
  width: number;
  top: number;
};

export type ChartElementDataOptions = {
  selectionMode?: "seriesFirst";
};

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

export type ChartHierarchyDatum = {
  children?: ChartHierarchyDatum[];
  colorIndex?: number;
  name: string;
  value?: number;
};

export type ComboRenderableGroup = {
  axisIds: number[];
  categoryAxis: XlsxChartAxis | null;
  chartType: string;
  gapWidth?: number;
  is3d?: boolean;
  raw?: Record<string, unknown>;
  series: XlsxChartSeries[];
  valueAxis: XlsxChartAxis | null;
};

export type ChartStage = {
  color: string;
  isSubtotal: boolean;
  label: string;
  value: number;
};

export type BoxWhiskerStats = {
  lowerFence: number;
  lowerWhisker: number;
  max: number;
  mean: number;
  median: number;
  min: number;
  outliers: number[];
  q1: number;
  q3: number;
  upperFence: number;
  upperWhisker: number;
  visiblePoints: number[];
};

export type SurfaceDomain = {
  maxValue: number;
  minValue: number;
  safeMax: number;
  ticks: number[];
};

export type RegionMapFeature = Feature<Geometry, { name?: string; regionSet?: "country" | "us-state"; stateCode?: string }>;

export const WORLD_COUNTRY_FEATURES = ((topojsonFeature(
  countries50m as unknown as Parameters<typeof topojsonFeature>[0],
  (countries50m as { objects: { countries: unknown } }).objects.countries as Parameters<typeof topojsonFeature>[1]
) as unknown) as FeatureCollection<Geometry, { name?: string }>).features as RegionMapFeature[];

export const US_STATE_NAME_BY_ID: Record<string, { code: string; name: string }> = {
  "01": { code: "AL", name: "Alabama" },
  "02": { code: "AK", name: "Alaska" },
  "04": { code: "AZ", name: "Arizona" },
  "05": { code: "AR", name: "Arkansas" },
  "06": { code: "CA", name: "California" },
  "08": { code: "CO", name: "Colorado" },
  "09": { code: "CT", name: "Connecticut" },
  "10": { code: "DE", name: "Delaware" },
  "11": { code: "DC", name: "District of Columbia" },
  "12": { code: "FL", name: "Florida" },
  "13": { code: "GA", name: "Georgia" },
  "15": { code: "HI", name: "Hawaii" },
  "16": { code: "ID", name: "Idaho" },
  "17": { code: "IL", name: "Illinois" },
  "18": { code: "IN", name: "Indiana" },
  "19": { code: "IA", name: "Iowa" },
  "20": { code: "KS", name: "Kansas" },
  "21": { code: "KY", name: "Kentucky" },
  "22": { code: "LA", name: "Louisiana" },
  "23": { code: "ME", name: "Maine" },
  "24": { code: "MD", name: "Maryland" },
  "25": { code: "MA", name: "Massachusetts" },
  "26": { code: "MI", name: "Michigan" },
  "27": { code: "MN", name: "Minnesota" },
  "28": { code: "MS", name: "Mississippi" },
  "29": { code: "MO", name: "Missouri" },
  "30": { code: "MT", name: "Montana" },
  "31": { code: "NE", name: "Nebraska" },
  "32": { code: "NV", name: "Nevada" },
  "33": { code: "NH", name: "New Hampshire" },
  "34": { code: "NJ", name: "New Jersey" },
  "35": { code: "NM", name: "New Mexico" },
  "36": { code: "NY", name: "New York" },
  "37": { code: "NC", name: "North Carolina" },
  "38": { code: "ND", name: "North Dakota" },
  "39": { code: "OH", name: "Ohio" },
  "40": { code: "OK", name: "Oklahoma" },
  "41": { code: "OR", name: "Oregon" },
  "42": { code: "PA", name: "Pennsylvania" },
  "44": { code: "RI", name: "Rhode Island" },
  "45": { code: "SC", name: "South Carolina" },
  "46": { code: "SD", name: "South Dakota" },
  "47": { code: "TN", name: "Tennessee" },
  "48": { code: "TX", name: "Texas" },
  "49": { code: "UT", name: "Utah" },
  "50": { code: "VT", name: "Vermont" },
  "51": { code: "VA", name: "Virginia" },
  "53": { code: "WA", name: "Washington" },
  "54": { code: "WV", name: "West Virginia" },
  "55": { code: "WI", name: "Wisconsin" },
  "56": { code: "WY", name: "Wyoming" }
};

export const US_STATE_FEATURES = ((topojsonFeature(
  countiesAlbers10m as unknown as Parameters<typeof topojsonFeature>[0],
  (countiesAlbers10m as { objects: { states: unknown } }).objects.states as Parameters<typeof topojsonFeature>[1]
) as unknown) as FeatureCollection<Geometry, { name?: string }>).features.map((feature) => {
  const id = typeof feature.id === "string" ? feature.id : String(feature.id ?? "");
  const state = US_STATE_NAME_BY_ID[id];
  return {
    ...feature,
    properties: {
      ...(feature.properties ?? {}),
      name: state?.name,
      regionSet: "us-state" as const,
      stateCode: state?.code
    }
  };
}) as RegionMapFeature[];

export const REGION_MAP_COUNTRY_ALIASES = new Map<string, string>([
  ["us", "united states of america"],
  ["usa", "united states of america"],
  ["u s a", "united states of america"],
  ["united states", "united states of america"],
  ["united states america", "united states of america"],
  ["u s", "united states of america"],
  ["uk", "united kingdom"],
  ["u k", "united kingdom"],
  ["uae", "united arab emirates"],
  ["u a e", "united arab emirates"],
  ["south korea", "korea, south"],
  ["north korea", "korea, north"],
  ["russia", "russian federation"],
  ["vietnam", "viet nam"],
  ["czech republic", "czechia"],
  ["ivory coast", "cote d'ivoire"],
  ["côte divoire", "cote d'ivoire"]
]);

export const REGION_MAP_US_STATE_ALIASES = new Map<string, string>([
  ["district of columbia", "district of columbia"],
  ["washington dc", "district of columbia"],
  ["washington d c", "district of columbia"],
  ["dc", "district of columbia"],
  ["d c", "district of columbia"]
]);

export const REGION_MAP_FEATURES_BY_KEY = (() => {
  const byKey = new Map<string, RegionMapFeature>();
  WORLD_COUNTRY_FEATURES.forEach((feature) => {
    const name = typeof feature.properties?.name === "string" ? feature.properties.name : "";
    const key = normalizeRegionMapKey(name);
    if (key.length > 0) {
      byKey.set(key, feature);
    }
  });
  return byKey;
})();

export const REGION_MAP_US_STATE_FEATURES_BY_KEY = (() => {
  const byKey = new Map<string, RegionMapFeature>();
  US_STATE_FEATURES.forEach((feature) => {
    const name = typeof feature.properties?.name === "string" ? feature.properties.name : "";
    const key = normalizeRegionMapKey(name);
    if (key.length > 0) {
      byKey.set(key, feature);
    }
    const stateCode = normalizeRegionMapKey(feature.properties?.stateCode);
    if (stateCode.length > 0) {
      byKey.set(stateCode, feature);
    }
  });
  REGION_MAP_US_STATE_ALIASES.forEach((value, key) => {
    const feature = byKey.get(normalizeRegionMapKey(value));
    if (feature) {
      byKey.set(normalizeRegionMapKey(key), feature);
    }
  });
  return byKey;
})();

export function parseRgbColor(color: string) {
  const match = /^#?([0-9a-f]{6})$/i.exec(color);
  if (!match) {
    return null;
  }
  return {
    blue: Number.parseInt(match[1].slice(4, 6), 16),
    green: Number.parseInt(match[1].slice(2, 4), 16),
    red: Number.parseInt(match[1].slice(0, 2), 16)
  };
}

export function mixRgbColor(color: string, mixWith: string, ratio: number) {
  const base = parseRgbColor(color);
  const target = parseRgbColor(mixWith);
  if (!base || !target) {
    return color;
  }
  const clamped = Math.max(0, Math.min(1, ratio));
  const mixChannel = (left: number, right: number) => Math.round(left + (right - left) * clamped);
  return `#${[
    mixChannel(base.red, target.red),
    mixChannel(base.green, target.green),
    mixChannel(base.blue, target.blue)
  ].map((value) => value.toString(16).padStart(2, "0")).join("")}`;
}

export function lightenColor(color: string, ratio: number) {
  return mixRgbColor(color, "#ffffff", ratio);
}

export function darkenColor(color: string, ratio: number) {
  return mixRgbColor(color, "#000000", ratio);
}

export function clamp(value: number, min: number, max: number) {
  return Math.max(min, Math.min(max, value));
}

export function normalizeRegionMapKey(value: unknown) {
  if (typeof value !== "string") {
    return "";
  }
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/&/g, " and ")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

export const DEFAULT_CHART_FONT_STACK = [
  "\"Aptos\"",
  "Calibri",
  "Carlito",
  "\"Segoe UI\"",
  "Tahoma",
  "Arial",
  "sans-serif"
].join(", ");

export const DEFAULT_CHART_TEXT_COLOR = "#000000";

export const DEFAULT_CHART_MUTED_TEXT_COLOR = "#7f7f7f";

export function escapeCssFontFamilyToken(value: string) {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }
  if (
    trimmed.startsWith("\"")
    || trimmed.startsWith("'")
    || /^(serif|sans-serif|monospace|cursive|fantasy|system-ui|math|emoji|fangsong)$/i.test(trimmed)
  ) {
    return trimmed;
  }
  return /\s/.test(trimmed) ? `"${trimmed.replace(/"/g, "\\\"")}"` : trimmed;
}

export function buildChartFontFamily(fontFamily: string | undefined) {
  if (!fontFamily || fontFamily.trim().length === 0) {
    return DEFAULT_CHART_FONT_STACK;
  }
  const tokens = fontFamily
    .split(",")
    .map(escapeCssFontFamilyToken)
    .filter((token) => token.length > 0);
  if (tokens.length === 0) {
    return DEFAULT_CHART_FONT_STACK;
  }
  return [...tokens, "\"Segoe UI\"", "Tahoma", "Arial", "sans-serif"].join(", ");
}

export function resolveChartTextColor(chart: XlsxChart) {
  return chart.textColor ?? chart.titleColor ?? DEFAULT_CHART_TEXT_COLOR;
}

export function resolveChartAxisTextColor(chart: XlsxChart) {
  return chart.axisLabelColor ?? chart.textColor ?? chart.titleColor ?? DEFAULT_CHART_TEXT_COLOR;
}

export function resolveChartMutedTextColor(chart: XlsxChart) {
  return chart.textColor ?? chart.axisLabelColor ?? DEFAULT_CHART_MUTED_TEXT_COLOR;
}

export function safeNumber(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === "string") {
    const normalized = value.trim().replace(/,/g, "");
    if (!normalized) {
      return null;
    }
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

export function chartSeriesColor(chart: XlsxChart, seriesIndex: number) {
  const series = chart.series[seriesIndex];
  const paletteColor = chart.chartColorPalette?.[seriesIndex % Math.max(1, chart.chartColorPalette.length)];
  return series?.color ?? series?.lineColor ?? paletteColor ?? chart.textColor ?? "#222222";
}

export function chartSeriesStrokeColor(chart: XlsxChart, seriesIndex: number) {
  const series = chart.series[seriesIndex];
  const paletteColor = chart.chartColorPalette?.[seriesIndex % Math.max(1, chart.chartColorPalette.length)];
  return series?.lineColor ?? series?.color ?? paletteColor ?? chart.textColor ?? "#222222";
}

export function chartPointColor(chart: XlsxChart, pointIndex: number, seriesIndex = 0) {
  const pointStyle = chart.series[seriesIndex]?.dataPointStyles?.find((entry) => entry.index === pointIndex);
  if (pointStyle?.color) {
    return pointStyle.color;
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

export function resolveRegionMapFeature(value: unknown, featureSet: "country" | "us-state" = "country") {
  const rawKey = normalizeRegionMapKey(value);
  if (!rawKey) {
    return null;
  }
  if (featureSet === "us-state") {
    const canonicalKey = REGION_MAP_US_STATE_ALIASES.get(rawKey) ?? rawKey;
    return REGION_MAP_US_STATE_FEATURES_BY_KEY.get(canonicalKey) ?? null;
  }
  const canonicalKey = REGION_MAP_COUNTRY_ALIASES.get(rawKey) ?? rawKey;
  return REGION_MAP_FEATURES_BY_KEY.get(canonicalKey) ?? null;
}

export function resolveRegionMapBaseColor(chart: XlsxChart, seriesIndex: number) {
  return chart.series[seriesIndex]?.color
    ?? chart.series[seriesIndex]?.lineColor
    ?? chart.chartColorPalette?.[0]
    ?? "#ff006e";
}

export function resolveRegionMapDataColor(chart: XlsxChart, seriesIndex: number) {
  const pointColor = normalizeRendererHexColor(chartPointColor(chart, 0, seriesIndex));
  if (pointColor) {
    return pointColor;
  }

  const palette = Array.isArray(chart.chartColorPalette) ? chart.chartColorPalette : [];
  if (palette.length > 0) {
    const offset = chart.chartColorPaletteOffset ?? 0;
    const paletteColor = normalizeRendererHexColor(
      palette[((offset % palette.length) + palette.length) % palette.length]
    );
    if (paletteColor) {
      return paletteColor;
    }
  }

  return normalizeRendererHexColor(chart.series[seriesIndex]?.color ?? chart.series[seriesIndex]?.lineColor)
    ?? "#4f81bd";
}

export function resolveRegionMapValueColors(series: XlsxChartSeries | null | undefined) {
  const raw = series?.raw && typeof series.raw === "object"
    ? series.raw as Record<string, unknown>
    : null;
  const colors = Array.isArray(raw?.valueColors)
    ? raw.valueColors
      .map((value) => normalizeRendererHexColor(value))
      .filter((value): value is string => Boolean(value))
    : [];
  return colors.length >= 2 ? colors : null;
}

export function resolveRegionMapColorStrings(series: XlsxChartSeries | null | undefined) {
  const raw = series?.raw && typeof series.raw === "object"
    ? series.raw as Record<string, unknown>
    : null;
  const values = Array.isArray(raw?.chartExColorStrings)
    ? raw.chartExColorStrings
      .map((value) => (typeof value === "string" ? value.trim() : ""))
      .filter((value): value is string => value.length > 0)
    : [];
  return values.length > 0 ? values : null;
}

export function resolveRegionMapLayoutProperties(series: XlsxChartSeries | null | undefined) {
  const raw = series?.raw && typeof series.raw === "object"
    ? series.raw as Record<string, unknown>
    : null;
  return raw?.layoutProperties && typeof raw.layoutProperties === "object"
    ? raw.layoutProperties as Record<string, unknown>
    : null;
}

export function resolveRegionMapFeatureSet(labels: string[], geography: Record<string, unknown> | null) {
  const cultureRegion = typeof geography?.cultureRegion === "string"
    ? geography.cultureRegion.trim().toUpperCase()
    : "";
  const countryMatches = labels.filter((label) => resolveRegionMapFeature(label, "country") != null).length;
  const stateMatches = labels.filter((label) => resolveRegionMapFeature(label, "us-state") != null).length;
  if (cultureRegion === "US" && stateMatches > 0 && stateMatches >= countryMatches) {
    return "us-state" as const;
  }
  return "country" as const;
}

export function getRegionMapBaseFeatures(featureSet: "country" | "us-state") {
  return featureSet === "us-state" ? US_STATE_FEATURES : WORLD_COUNTRY_FEATURES;
}

export function resolveRegionMapValueColorFromStops(stops: string[], ratio: number) {
  if (stops.length === 0) {
    return "#4f81bd";
  }
  if (stops.length === 1) {
    return stops[0];
  }
  const clamped = clamp(ratio, 0, 1);
  const scaled = clamped * (stops.length - 1);
  const lowerIndex = Math.floor(scaled);
  const upperIndex = Math.min(stops.length - 1, lowerIndex + 1);
  const mixRatio = scaled - lowerIndex;
  return mixRgbColor(stops[lowerIndex] ?? stops[0], stops[upperIndex] ?? stops[stops.length - 1], mixRatio);
}

export function resolveRegionMapValueColor(chart: XlsxChart, seriesIndex: number, ratio: number) {
  const explicitStops = resolveRegionMapValueColors(chart.series[seriesIndex] ?? null);
  if (explicitStops) {
    return resolveRegionMapValueColorFromStops(explicitStops, ratio);
  }
  const baseColor = resolveRegionMapDataColor(chart, seriesIndex);
  return resolveRegionMapValueColorFromStops([
    lightenColor(baseColor, 0.82),
    lightenColor(baseColor, 0.28),
    darkenColor(baseColor, 0.08)
  ], ratio);
}

export function resolveRegionMapNoDataColor(chart: XlsxChart, seriesIndex: number) {
  const baseColor = resolveRegionMapBaseColor(chart, seriesIndex);
  return normalizeRendererHexColor(baseColor) ?? "#ff006e";
}

export function buildRegionMapLegendItems(chart: XlsxChart): LegendItem[] {
  const primarySeriesIndex = Math.max(0, chart.series.findIndex((series) => series.hidden !== true));
  const categoricalValues = resolveRegionMapColorStrings(chart.series[primarySeriesIndex] ?? null);
  if (categoricalValues) {
    const uniqueValues = Array.from(new Set(categoricalValues));
    return uniqueValues.map((value, index) => ({
      color: chartPointColor(chart, index, primarySeriesIndex),
      label: value
    }));
  }
  const values = (chart.series[primarySeriesIndex]?.values ?? [])
    .map((value) => safeNumber(value))
    .filter((value): value is number => value != null);
  if (values.length === 0) {
    return [];
  }
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  const ticks = buildNumericTickValues(minValue, maxValue, undefined).slice(0, 5);
  return ticks.slice(0, -1).map((tick, index) => {
    const nextTick = ticks[index + 1] ?? maxValue;
    const midpoint = tick + (nextTick - tick) * 0.5;
    const ratio = (midpoint - minValue) / Math.max(1e-6, maxValue - minValue);
    return {
      color: resolveRegionMapValueColor(chart, primarySeriesIndex, ratio),
      label: `${formatTickValue(tick)}-${formatTickValue(nextTick)}`
    };
  });
}

export function normalizeBuiltinPieStyleId(styleId: number | undefined) {
  if (typeof styleId !== "number" || !Number.isFinite(styleId)) {
    return null;
  }
  return styleId >= 100 ? styleId - 100 : styleId;
}

export function getBuiltinPiePalette(chart: XlsxChart, seriesIndex: number) {
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

export function normalizeLegendPosition(position: string | undefined) {
  switch (position) {
    case "bottom":
      return "bottom";
    case "left":
      return "left";
    case "right":
      return "right";
    case "top":
      return "top";
    case "b":
      return "bottom";
    case "l":
      return "left";
    case "r":
      return "right";
    case "t":
      return "top";
    default:
      return position;
  }
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

export function normalizeCategoryLabel(value: unknown) {
  if (value == null) {
    return "";
  }
  return String(value)
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .trim();
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

export function isComboChart(chart: XlsxChart) {
  const typeGroups = chart.typeGroups ?? [];
  if (typeGroups.length < 2) {
    return false;
  }
  const distinctChartTypes = new Set(typeGroups.map((group) => group.chartType));
  return distinctChartTypes.size > 1;
}

export function getComboLegendSeries(chart: XlsxChart) {
  if (!isComboChart(chart)) {
    return chart.series.map((series, index) => ({
      color: chartSeriesColor(chart, index),
      label: series.name ?? `Series ${index + 1}`
    }));
  }
  return (chart.typeGroups ?? []).flatMap((group) => (
    group.series.map((series, seriesIndex) => ({
      color: series.lineColor ?? series.markerColor ?? series.color ?? chartSeriesColor(chart, seriesIndex),
      label: series.name ?? `Series ${seriesIndex + 1}`
    }))
  ));
}

export function findAxisForGroup(
  chart: XlsxChart,
  axisIds: number[],
  positionMatcher: (position: string | undefined) => boolean,
  allowAnyMatch = false
) {
  const positionedMatch = chart.axes.find((axis) => (
    axis.id != null
    && axisIds.includes(axis.id)
    && positionMatcher(axis.position)
  ));
  if (positionedMatch) {
    return positionedMatch;
  }
  if (!allowAnyMatch) {
    return null;
  }
  return chart.axes.find((axis) => axis.id != null && axisIds.includes(axis.id)) ?? null;
}

export function buildComboGroups(chart: XlsxChart): ComboRenderableGroup[] {
  return (chart.typeGroups ?? []).map((group) => {
    const axisIds = group.axisIds ?? [];
    const categoryAxis = findAxisForGroup(chart, axisIds, (position) => position === "b" || position === "t")
      ?? chart.categoryAxis
      ?? null;
    const valueAxis = findAxisForGroup(chart, axisIds, (position) => position === "l" || position === "r", true)
      ?? chart.valueAxis
      ?? null;
    return {
      axisIds,
      categoryAxis,
      chartType: group.chartType,
      gapWidth: group.gapWidth,
      is3d: group.is3d,
      raw: group.raw,
      series: group.series,
      valueAxis
    };
  }).filter((group) => group.series.length > 0);
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

export type SurfacePathPoint = { x: number; y: number };

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

export type SurfaceContourSegment = {
  end: SurfacePathPoint;
  start: SurfacePathPoint;
};

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

export function normalizeRendererHexColor(value: unknown) {
  if (typeof value !== "string") {
    return null;
  }
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }
  const normalized = trimmed.startsWith("#") ? trimmed.slice(1) : trimmed;
  return /^[0-9a-f]{6}$/i.test(normalized) ? `#${normalized.toLowerCase()}` : null;
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

export function resolve3dFrameOffsets(chart: XlsxChart, baseDepthX = 11, baseDepthY = 8) {
  const depthRatio = clamp((chart.view3d?.depthPercent ?? 100) / 100, 0.35, 3.2);
  const rotX = clamp(chart.view3d?.rotX ?? 20, -80, 80);
  const rotY = clamp(chart.view3d?.rotY ?? 20, -80, 80);
  const usePerspective = chart.view3d?.rAngAx === false;
  const perspectiveStrength = clamp((chart.view3d?.perspective ?? (usePerspective ? 26 : 0)) / 100, 0, 1);
  const projectedOrigin = projectCartesian3dPoint(
    0,
    0,
    0,
    rotX * (Math.PI / 180),
    rotY * (Math.PI / 180),
    usePerspective,
    perspectiveStrength
  );
  const projectedDepthPoint = projectCartesian3dPoint(
    0,
    0,
    1,
    rotX * (Math.PI / 180),
    rotY * (Math.PI / 180),
    usePerspective,
    perspectiveStrength
  );
  const depthDirectionX = projectedDepthPoint.x - projectedOrigin.x;
  const depthDirectionY = projectedDepthPoint.y - projectedOrigin.y;
  const horizontalBias = clamp(1.42 + Math.abs(rotY) / 48 + depthRatio * 0.06, 1.5, 1.95);
  const verticalBias = clamp(1.18 + Math.abs(rotX) / 78, 1.22, 1.5);
  const biasedDepthDirectionX = depthDirectionX * horizontalBias;
  const biasedDepthDirectionY = depthDirectionY * verticalBias;
  const depthDirectionMagnitude = Math.hypot(biasedDepthDirectionX, biasedDepthDirectionY) || 1;
  const horizontalFactor = clamp(Math.abs(rotY) / 22, 0.55, 2.2);
  const verticalFactor = clamp(Math.abs(rotX) / 18, 0.45, 1.9);
  const legacyDepthX = Math.max(6, baseDepthX * depthRatio * horizontalFactor) * (rotY < 0 ? -1 : 1);
  const legacyDepthY = -Math.max(4, baseDepthY * depthRatio * verticalFactor);
  const depthMagnitude = Math.max(10, Math.hypot(legacyDepthX, legacyDepthY) * 1.34);
  const depthX = biasedDepthDirectionX / depthDirectionMagnitude * depthMagnitude;
  const depthY = biasedDepthDirectionY / depthDirectionMagnitude * depthMagnitude;
  return {
    depthRatio,
    depthX,
    depthY,
    insetBottom: Math.max(6, Math.abs(depthY) + 5),
    insetLeft: depthX < 0 ? Math.abs(depthX) + 4 : 0,
    insetRight: depthX > 0 ? depthX + 4 : 0,
    insetTop: Math.max(4, Math.abs(depthY) + 2)
  };
}

export function isLikelyDateFormatCode(formatCode: string | undefined) {
  if (!formatCode) {
    return false;
  }
  const normalized = formatCode
    .toLowerCase()
    .replace(/\[[^\]]*]/g, " ")
    .replace(/"[^"]*"/g, " ")
    .replace(/\\./g, " ");
  return /(?:d|m|y)/.test(normalized);
}

export function excelSerialToDate(serial: number, use1904: boolean) {
  const wholeDays = Math.trunc(serial);
  const milliseconds = Math.round((serial - wholeDays) * 86_400_000);
  const baseUtc = use1904
    ? Date.UTC(1904, 0, 1)
    : Date.UTC(1899, 11, 30);
  return new Date(baseUtc + wholeDays * 86_400_000 + milliseconds);
}

export function formatExcelDateSerial(value: number, formatCode: string | undefined, use1904: boolean) {
  const date = excelSerialToDate(value, use1904);
  const normalized = (formatCode ?? "").toLowerCase();
  const options: Intl.DateTimeFormatOptions = {};

  if (/yyyy/.test(normalized)) {
    options.year = "numeric";
  } else if (/yy/.test(normalized)) {
    options.year = "2-digit";
  }

  if (/mmmm/.test(normalized)) {
    options.month = "long";
  } else if (/mmm/.test(normalized)) {
    options.month = "short";
  } else if (/(^|[^a-z])m([^a-z]|$)|(^|[^a-z])mm([^a-z]|$)/.test(normalized)) {
    options.month = "numeric";
  }

  if (/dddd/.test(normalized)) {
    options.weekday = "long";
  } else if (/ddd/.test(normalized)) {
    options.weekday = "short";
  }

  if (/d/.test(normalized)) {
    options.day = "numeric";
  }

  if (Object.keys(options).length === 0) {
    options.month = "short";
    options.day = "numeric";
  }

  return new Intl.DateTimeFormat(undefined, options).format(date);
}

export function formatCategoryLabel(chart: XlsxChart, value: unknown, fallback = "") {
  const rawRecord = chart.raw && typeof chart.raw === "object"
    ? chart.raw as Record<string, unknown>
    : null;
  const numeric = typeof value === "number" && Number.isFinite(value)
    ? value
    : null;
  const formatCode = chart.categoryAxis?.numberFormat?.formatCode;
  if (numeric != null && isLikelyDateFormatCode(formatCode)) {
    return formatExcelDateSerial(numeric, formatCode, rawRecord?.date1904 === true);
  }

  const normalized = normalizeCategoryLabel(value);
  if (normalized.length > 0) {
    return normalized;
  }
  return fallback;
}

export function truncateSvgText(value: string, maxWidth: number, fontSize = 10) {
  if (!value || maxWidth <= 0) {
    return "";
  }
  const charWidth = Math.max(4.2, fontSize * 0.56);
  const maxChars = Math.max(1, Math.floor(maxWidth / charWidth));
  if (value.length <= maxChars) {
    return value;
  }
  if (maxChars <= 1) {
    return "…";
  }
  return `${value.slice(0, Math.max(0, maxChars - 1)).trimEnd()}…`;
}

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

export const TWO_PI = Math.PI * 2;

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
    if (segmentEnd - segmentStart > 1e-4) {
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

export function toSvgNumber(value: number) {
  return Number(value.toFixed(3));
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

export function scaleProjectedVector(vector: { x: number; y: number }, scale: number) {
  return {
    x: vector.x * scale,
    y: vector.y * scale
  };
}

export function sampleProjectedEllipseArc(
  center: { x: number; y: number },
  axisA: { x: number; y: number },
  axisB: { x: number; y: number },
  startAngle: number,
  endAngle: number
) {
  const steps = Math.max(12, Math.ceil(Math.abs(endAngle - startAngle) / (Math.PI / 14)));
  const points: Array<{ x: number; y: number }> = [];
  for (let step = 0; step <= steps; step += 1) {
    const ratio = step / steps;
    const angle = startAngle + ((endAngle - startAngle) * ratio);
    points.push({
      x: center.x + (axisA.x * Math.cos(angle)) + (axisB.x * Math.sin(angle)),
      y: center.y + (axisA.y * Math.cos(angle)) + (axisB.y * Math.sin(angle))
    });
  }
  return points;
}

export function buildRibbonSvgPath(startArc: Array<{ x: number; y: number }>, endArc: Array<{ x: number; y: number }>) {
  if (startArc.length === 0 || endArc.length === 0) {
    return "";
  }
  return buildLinearSvgPath([...startArc, ...endArc.slice().reverse()], true);
}

export function renderRadialFrustum(
  bar: BarRect,
  normalizedShape: "cone" | "cylinder",
  frontFill: string,
  sideFill: string,
  topFill: string
) {
  const depthX = bar.depthX ?? (bar.isHorizontal ? 10 : 9);
  const depthY = bar.depthY ?? -7;
  const frontX = bar.left + (bar.depthOffsetX ?? 0);
  const frontY = bar.top + (bar.depthOffsetY ?? 0);
  const frontW = bar.width;
  const frontH = bar.height;
  const centerX = frontX + frontW / 2;
  const centerY = frontY + frontH / 2;
  const showStartCap = bar.capStart !== false;
  const showEndCap = bar.capEnd !== false;
  const startScale = normalizedShape === "cone" ? clamp(bar.bottomScale ?? 1, 0, 1) : 1;
  const endScale = normalizedShape === "cone" ? clamp(bar.topScale ?? 1, 0, 1) : 1;
  const radialDepthScale = normalizedShape === "cylinder"
    ? (bar.isHorizontal ? 0.52 : 0.48)
    : (bar.isHorizontal ? 0.7 : 0.58);
  const depthAxis = scaleProjectedVector({ x: depthX / 2, y: depthY / 2 }, radialDepthScale);
  const secondaryStrokeWidth = Math.max(0.6, bar.strokeWidth * 0.65);

  if (bar.isHorizontal) {
    const startCenter = { x: frontX + depthAxis.x, y: centerY + depthAxis.y };
    const endCenter = { x: frontX + frontW + depthAxis.x, y: centerY + depthAxis.y };
    const startAxisA = { x: 0, y: (frontH * startScale) / 2 };
    const endAxisA = { x: 0, y: (frontH * endScale) / 2 };
    const startAxisB = scaleProjectedVector(depthAxis, startScale);
    const endAxisB = scaleProjectedVector(depthAxis, endScale);
    const backStartArc = sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, 0, Math.PI);
    const backEndArc = sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, 0, Math.PI);
    const frontStartArc = sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, Math.PI, Math.PI * 2);
    const frontEndArc = sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, Math.PI, Math.PI * 2);
    const backWallPath = buildRibbonSvgPath(backEndArc, backStartArc);
    const frontWallPath = buildRibbonSvgPath(frontEndArc, frontStartArc);
    const endCapPath = endScale > 0.018
      ? buildLinearSvgPath(sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, 0, Math.PI * 2), true)
      : "";
    const startCapPath = startScale > 0.028
      ? buildLinearSvgPath(sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, 0, Math.PI * 2), true)
      : "";
    return (
      <g key={`${bar.key}-3d-horizontal-${normalizedShape}`}>
        {backWallPath ? <path d={backWallPath} fill={sideFill} stroke={bar.stroke} stroke-width={secondaryStrokeWidth} /> : null}
        {showEndCap && endCapPath ? <path d={endCapPath} fill={topFill} stroke={bar.stroke} stroke-width={secondaryStrokeWidth} /> : null}
        {frontWallPath ? <path d={frontWallPath} fill={frontFill} stroke={bar.stroke} stroke-width={bar.strokeWidth} /> : null}
        {showStartCap && startCapPath ? <path d={startCapPath} fill={frontFill} stroke={bar.stroke} stroke-width={secondaryStrokeWidth} /> : null}
      </g>
    );
  }

  const startCenter = { x: centerX + depthAxis.x, y: frontY + frontH + depthAxis.y };
  const endCenter = { x: centerX + depthAxis.x, y: frontY + depthAxis.y };
  const startAxisA = { x: (frontW * startScale) / 2, y: 0 };
  const endAxisA = { x: (frontW * endScale) / 2, y: 0 };
  const startAxisB = scaleProjectedVector(depthAxis, startScale);
  const endAxisB = scaleProjectedVector(depthAxis, endScale);
  const backStartArc = sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, 0, Math.PI);
  const backEndArc = sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, 0, Math.PI);
  const frontStartArc = sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, Math.PI, Math.PI * 2);
  const frontEndArc = sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, Math.PI, Math.PI * 2);
  const backWallPath = buildRibbonSvgPath(backEndArc, backStartArc);
  const frontWallPath = buildRibbonSvgPath(frontEndArc, frontStartArc);
  const endCapPath = endScale > 0.018
    ? buildLinearSvgPath(sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, 0, Math.PI * 2), true)
    : "";
  return (
    <g key={`${bar.key}-3d-${normalizedShape}`}>
      {backWallPath ? <path d={backWallPath} fill={sideFill} stroke={bar.stroke} stroke-width={secondaryStrokeWidth} /> : null}
      {showEndCap && endCapPath ? <path d={endCapPath} fill={topFill} stroke={bar.stroke} stroke-width={secondaryStrokeWidth} /> : null}
      {frontWallPath ? <path d={frontWallPath} fill={frontFill} stroke={bar.stroke} stroke-width={bar.strokeWidth} /> : null}
    </g>
  );
}

export function renderExtrudedRect(bar: BarRect) {
  const normalizedShape = (() => {
    switch (bar.shape3d) {
      case "cone":
      case "coneToMax":
        return "cone";
      case "cylinder":
        return "cylinder";
      case "pyramid":
      case "pyramidToMax":
        return "pyramid";
      default:
        return "box";
    }
  })();
  const depthX = bar.depthX ?? (bar.isHorizontal ? 10 : 9);
  const depthY = bar.depthY ?? -7;
  const frontX = bar.left + (bar.depthOffsetX ?? 0);
  const frontY = bar.top + (bar.depthOffsetY ?? 0);
  const frontW = bar.width;
  const frontH = bar.height;
  const frontX2 = frontX + frontW;
  const frontY2 = frontY + frontH;

  const sideAnchorX = frontX2;
  const sideDepthX = depthX;
  const centerX = frontX + frontW / 2;
  const bottomScale = clamp(bar.bottomScale ?? 1, 0.04, 1);
  const topScale = clamp(bar.topScale ?? 1, 0.04, 1);
  const bottomHalfWidth = (frontW * bottomScale) / 2;
  const topHalfWidth = (frontW * topScale) / 2;
  const bottomLeft = centerX - bottomHalfWidth;
  const bottomRight = centerX + bottomHalfWidth;
  const topLeft = centerX - topHalfWidth;
  const topRight = centerX + topHalfWidth;
  const showStartCap = bar.capStart !== false;
  const showEndCap = bar.capEnd !== false;

  const topFace = `${frontX},${frontY} ${frontX2},${frontY} ${frontX2 + sideDepthX},${frontY + depthY} ${frontX + sideDepthX},${frontY + depthY}`;
  const sideFace = `${sideAnchorX},${frontY} ${sideAnchorX},${frontY2} ${sideAnchorX + sideDepthX},${frontY2 + depthY} ${sideAnchorX + sideDepthX},${frontY + depthY}`;
  const frontFill = bar.gradientId ? `url(#${bar.gradientId})` : bar.color;
  const sideFill = bar.invertedNegative ? bar.color : darkenColor(bar.color, 0.22);
  const topFill = bar.invertedNegative ? lightenColor(bar.color, 0.04) : lightenColor(bar.color, 0.24);

  if (normalizedShape === "cylinder" || normalizedShape === "cone") {
    return renderRadialFrustum(bar, normalizedShape, frontFill, sideFill, topFill);
  }

  if (normalizedShape === "box") {
    return (
      <g key={`${bar.key}-3d`}>
        <polygon fill={sideFill} points={sideFace} stroke={bar.stroke} stroke-width={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <polygon fill={topFill} points={topFace} stroke={bar.stroke} stroke-width={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <rect fill={frontFill} height={frontH} stroke={bar.stroke} stroke-width={bar.strokeWidth} width={frontW} x={frontX} y={frontY} />
      </g>
    );
  }

  if (bar.isHorizontal) {
    const centerY = frontY + frontH / 2;
    const startHalfHeight = (frontH * clamp(bar.bottomScale ?? 1, 0.04, 1)) / 2;
    const endHalfHeight = (frontH * clamp(bar.topScale ?? 1, 0.04, 1)) / 2;
    const startTop = centerY - startHalfHeight;
    const startBottom = centerY + startHalfHeight;
    const endTop = centerY - endHalfHeight;
    const endBottom = centerY + endHalfHeight;
    const topFacePoints = `${frontX},${startTop} ${frontX2},${endTop} ${frontX2 + sideDepthX},${endTop + depthY} ${frontX + sideDepthX},${startTop + depthY}`;
    const farSidePoints = `${frontX2},${endTop} ${frontX2},${endBottom} ${frontX2 + sideDepthX},${endBottom + depthY} ${frontX2 + sideDepthX},${endTop + depthY}`;

    const taperedFarFace = `${frontX2},${endTop} ${frontX2 + sideDepthX},${endTop + depthY} ${frontX2 + sideDepthX},${endBottom + depthY} ${frontX2},${endBottom}`;
    const frontPolygon = `${frontX},${startTop} ${frontX2},${endTop} ${frontX2},${endBottom} ${frontX},${startBottom}`;

    return (
      <g key={`${bar.key}-3d-horizontal-${normalizedShape}`}>
        <polygon fill={sideFill} points={taperedFarFace} stroke={bar.stroke} stroke-width={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <polygon fill={topFill} points={topFacePoints} stroke={bar.stroke} stroke-width={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <polygon fill={frontFill} points={frontPolygon} stroke={bar.stroke} stroke-width={bar.strokeWidth} />
      </g>
    );
  }

  const taperedTopFace = `${topLeft},${frontY} ${topRight},${frontY} ${topRight + sideDepthX},${frontY + depthY} ${topLeft + sideDepthX},${frontY + depthY}`;
  const taperedSideFace = `${topRight},${frontY} ${bottomRight},${frontY2} ${bottomRight + sideDepthX},${frontY2 + depthY} ${topRight + sideDepthX},${frontY + depthY}`;
  const pyramidFrontFace = `${topLeft},${frontY} ${topRight},${frontY} ${bottomRight},${frontY2} ${bottomLeft},${frontY2}`;

  return (
    <g key={`${bar.key}-3d-${normalizedShape}`}>
      <polygon fill={sideFill} points={taperedSideFace} stroke={bar.stroke} stroke-width={Math.max(0.6, bar.strokeWidth * 0.65)} />
      {showEndCap ? (
        <polygon fill={topFill} points={taperedTopFace} stroke={bar.stroke} stroke-width={Math.max(0.6, bar.strokeWidth * 0.65)} />
      ) : null}
      <polygon fill={frontFill} points={pyramidFrontFace} stroke={bar.stroke} stroke-width={bar.strokeWidth} />
    </g>
  );
}

export function buildLinearSvgPath(points: Array<{ x: number; y: number }>, close = false) {
  if (points.length === 0) {
    return "";
  }
  const commands = points.map((point, index) => (
    `${index === 0 ? "M" : "L"} ${toSvgNumber(point.x)} ${toSvgNumber(point.y)}`
  ));
  if (close) {
    commands.push("Z");
  }
  return commands.join(" ");
}

export function projectCartesian3dPoint(
  x: number,
  y: number,
  z: number,
  rotXRad: number,
  rotYRad: number,
  usePerspective: boolean,
  perspectiveStrength: number
) {
  const cosX = Math.cos(rotXRad);
  const sinX = Math.sin(rotXRad);
  const cosY = Math.cos(rotYRad);
  const sinY = Math.sin(rotYRad);

  const x1 = x * cosY + z * sinY;
  const z1 = -x * sinY + z * cosY;
  const y1 = y * cosX - z1 * sinX;
  const z2 = y * sinX + z1 * cosX;
  const perspective = usePerspective
    ? 1 / Math.max(0.22, 1 + z2 * (0.26 + perspectiveStrength * 0.54))
    : 1;

  return {
    depth: z2,
    x: x1 * perspective,
    y: y1 * perspective
  };
}

export function renderLineOrAreaChart3d(
  chart: XlsxChart,
  palette: ChartRendererPalette,
  layout: ChartLayout,
  categories: string[],
  stackedPointsBySeries: Array<Array<{ defined: boolean; y: number | null; y0: number | null; y1: number | null }>>,
  minValue: number,
  maxValue: number,
  isAreaChart: boolean,
  isStackedSeries: boolean,
  selectedChartElement?: XlsxChartElementSelection | null
) {
  const plot = layout.plot;
  const valueDomain = resolveAxisDomainWithChartOverrides(
    chart.valueAxis,
    minValue,
    maxValue,
    isAreaChart || chart.valueAxis?.crosses === "autoZero"
  );
  minValue = valueDomain.min;
  maxValue = valueDomain.max;
  const valueSpan = Math.max(1e-6, maxValue - minValue);
  const seriesCount = Math.max(1, chart.series.length);
  const categoryCount = Math.max(1, categories.length);
  const depthScale = clamp((chart.view3d?.depthPercent ?? 100) / 100, 0.5, 4);
  const halfDepth = (seriesCount <= 1 ? 0.82 : 1.02) * depthScale;
  const frontZ = halfDepth;
  const backZ = -halfDepth;
  const baseValue = isAreaChart ? Math.max(minValue, Math.min(maxValue, 0)) : minValue;
  const rotXRad = clamp(chart.view3d?.rotX ?? 18, -80, 80) * (Math.PI / 180);
  const rotYRad = clamp(chart.view3d?.rotY ?? 24, -80, 80) * (Math.PI / 180);
  const usePerspective = chart.view3d?.rAngAx === false;
  const perspectiveStrength = clamp((chart.view3d?.perspective ?? (usePerspective ? 26 : 0)) / 100, 0, 1);

  const normalizeX = (categoryIndex: number) => (
    categoryCount <= 1 ? 0 : ((categoryIndex / (categoryCount - 1)) - 0.5) * 2
  );
  const normalizeY = (value: number) => (
    -((((value - minValue) / valueSpan) - 0.5) * 2)
  );
  const stackedAreaDepthSpan = isAreaChart && isStackedSeries
    ? Math.max(0.18, (frontZ - backZ) * 0.56)
    : 0;
  const stackedAreaBackZ = frontZ - stackedAreaDepthSpan;
  const normalizeZ = (seriesIndex: number) => (
    seriesCount <= 1 || (isAreaChart && isStackedSeries)
      ? frontZ * 0.94
      : (((seriesIndex / (seriesCount - 1)) - 0.5) * (frontZ - backZ))
  );
  const projectSeriesPoint = (x: number, value: number, z: number) => projectCartesian3dPoint(
    x,
    normalizeY(value),
    z,
    rotXRad,
    rotYRad,
    usePerspective,
    perspectiveStrength
  );

  const cubeCorners = [
    { x: -1, y: -1, z: backZ },
    { x: 1, y: -1, z: backZ },
    { x: 1, y: 1, z: backZ },
    { x: -1, y: 1, z: backZ },
    { x: -1, y: -1, z: frontZ },
    { x: 1, y: -1, z: frontZ },
    { x: 1, y: 1, z: frontZ },
    { x: -1, y: 1, z: frontZ }
  ].map((corner) => projectCartesian3dPoint(
    corner.x,
    corner.y,
    corner.z,
    rotXRad,
    rotYRad,
    usePerspective,
    perspectiveStrength
  ));

  const projectedSeries = stackedPointsBySeries.map((seriesPoints, seriesIndex) => {
    const z = normalizeZ(seriesIndex);
    return seriesPoints.map((point, categoryIndex) => {
      const x = normalizeX(categoryIndex);
      const topValue = isStackedSeries ? (point.y1 ?? point.y ?? baseValue) : (point.y ?? baseValue);
      const bottomValue = isAreaChart
        ? (isStackedSeries ? (point.y0 ?? baseValue) : baseValue)
        : baseValue;
      const top = projectSeriesPoint(x, topValue, z);
      const bottom = projectSeriesPoint(x, bottomValue, z);
      const topBack = isAreaChart && isStackedSeries
        ? projectSeriesPoint(x, topValue, stackedAreaBackZ)
        : null;
      const bottomBack = isAreaChart && isStackedSeries
        ? projectSeriesPoint(x, bottomValue, stackedAreaBackZ)
        : null;
      return {
        bottom,
        bottomBack,
        categoryIndex,
        defined: point.defined,
        depth: top.depth,
        depthBack: topBack?.depth ?? top.depth,
        top,
        topBack
      };
    });
  });

  const bounds = [
    ...cubeCorners,
    ...projectedSeries.flatMap((series) => series.flatMap((point) => point.defined
      ? [point.top, point.bottom, ...(point.topBack ? [point.topBack] : []), ...(point.bottomBack ? [point.bottomBack] : [])]
      : []))
  ];
  const minX = Math.min(...bounds.map((point) => point.x));
  const maxX = Math.max(...bounds.map((point) => point.x));
  const minY = Math.min(...bounds.map((point) => point.y));
  const maxY = Math.max(...bounds.map((point) => point.y));
  const scale = Math.min(
    plot.width / Math.max(0.4, maxX - minX),
    plot.height / Math.max(0.4, maxY - minY)
  ) * 0.82;
  const centerRawX = (minX + maxX) / 2;
  const centerRawY = (minY + maxY) / 2;
  const centerX = plot.left + plot.width / 2;
  const centerY = plot.top + plot.height / 2;
  const toScreenPoint = (point: { depth: number; x: number; y: number }) => ({
    depth: point.depth,
    x: centerX + (point.x - centerRawX) * scale,
    y: centerY + (point.y - centerRawY) * scale
  });

  const screenCorners = cubeCorners.map(toScreenPoint);
  const seriesGeometry = projectedSeries.map((series, seriesIndex) => ({
    averageDepth: series.reduce((sum, point) => sum + ((point.depth + point.depthBack) / 2), 0) / Math.max(1, series.length),
    points: series.map((point) => ({
      bottom: toScreenPoint(point.bottom),
      bottomBack: point.bottomBack ? toScreenPoint(point.bottomBack) : null,
      categoryIndex: point.categoryIndex,
      defined: point.defined,
      top: toScreenPoint(point.top),
      topBack: point.topBack ? toScreenPoint(point.topBack) : null
    })),
    seriesIndex
  })).sort((left, right) => left.averageDepth - right.averageDepth);

  const edgeColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const gridColor = lightenColor(edgeColor, 0.56);
  const labelColor = resolveChartAxisTextColor(chart);
  const boxEdges: Array<[number, number, boolean]> = [
    [0, 1, false], [1, 2, false], [2, 3, false], [3, 0, false],
    [4, 5, true], [5, 6, true], [6, 7, true], [7, 4, true],
    [0, 4, false], [1, 5, true], [2, 6, true], [3, 7, false]
  ];
  const yTicks = valueDomain.ticks;
  const xLabelPoints = categories.map((category, categoryIndex) => ({
    label: category,
    point: toScreenPoint(projectCartesian3dPoint(
      normalizeX(categoryIndex),
      -1.06,
      frontZ,
      rotXRad,
      rotYRad,
      usePerspective,
      perspectiveStrength
    ))
  }));
  const yLabelPoints = yTicks.map((tick) => ({
    point: toScreenPoint(projectCartesian3dPoint(
      -1.04,
      normalizeY(tick),
      frontZ,
      rotXRad,
      rotYRad,
      usePerspective,
      perspectiveStrength
    )),
    tick
  }));

  return (
    <g>
      {yTicks.map((tick) => {
        const yNorm = normalizeY(tick);
        const leftBack = toScreenPoint(projectCartesian3dPoint(-1, yNorm, backZ, rotXRad, rotYRad, usePerspective, perspectiveStrength));
        const rightBack = toScreenPoint(projectCartesian3dPoint(1, yNorm, backZ, rotXRad, rotYRad, usePerspective, perspectiveStrength));
        const leftFront = toScreenPoint(projectCartesian3dPoint(-1, yNorm, frontZ, rotXRad, rotYRad, usePerspective, perspectiveStrength));
        return (
          <g key={`line3d-grid-${tick}`}>
            <line stroke={gridColor} stroke-width={1} x1={leftBack.x} x2={rightBack.x} y1={leftBack.y} y2={rightBack.y} />
            <line stroke={gridColor} stroke-width={0.9} x1={leftBack.x} x2={leftFront.x} y1={leftBack.y} y2={leftFront.y} />
          </g>
        );
      })}
      {boxEdges.map(([startIndex, endIndex, emphasized], index) => {
        const start = screenCorners[startIndex];
        const end = screenCorners[endIndex];
        return (
          <line
            key={`line3d-box-${index}`}
            stroke={emphasized ? edgeColor : lightenColor(edgeColor, 0.34)}
            stroke-width={emphasized ? 1.35 : 1}
            x1={start.x}
            x2={end.x}
            y1={start.y}
            y2={end.y}
          />
        );
      })}
      {seriesGeometry.map(({ points, seriesIndex }) => {
        const definedPoints = points.filter((point) => point.defined);
        if (definedPoints.length === 0) {
          return null;
        }
        const linePoints = definedPoints.map((point) => point.top);
        const areaPoints = isAreaChart
          ? [
              ...definedPoints.map((point) => point.top),
              ...definedPoints.slice().reverse().map((point) => point.bottom)
            ]
          : [];
        const areaBackPoints = isAreaChart && isStackedSeries
          ? [
              ...definedPoints.map((point) => point.topBack ?? point.top),
              ...definedPoints.slice().reverse().map((point) => point.bottomBack ?? point.bottom)
            ]
          : [];
        const strokeColor = chartSeriesStrokeColor(chart, seriesIndex);
        const fillColor = chartSeriesColor(chart, seriesIndex);
        const markerSymbol = normalizeChartMarkerSymbol(chart.series[seriesIndex]?.markerSymbol);
        const markerPath = markerSymbolPath(markerSymbol, Math.max(4, chart.series[seriesIndex]?.markerSize ?? 6) * 0.52);
        const slabFaces = isAreaChart && isStackedSeries
          ? definedPoints.slice(1).map((point, pointIndex) => {
              const previous = definedPoints[pointIndex];
              if (!previous?.topBack || !point.topBack || !previous.bottomBack || !point.bottomBack) {
                return null;
              }
              const topFace = buildLinearSvgPath([previous.top, point.top, point.topBack, previous.topBack], true);
              const bottomFace = buildLinearSvgPath([previous.bottom, point.bottom, point.bottomBack, previous.bottomBack], true);
              return (
                <>
                  <path
                    {...chartElementDataProps(seriesIndex)}
                    d={topFace}
                    fill={lightenColor(fillColor, 0.08)}
                    fill-opacity={0.8}
                    stroke={darkenColor(fillColor, 0.16)}
                    stroke-width={0.8}
                  />
                  <path
                    {...chartElementDataProps(seriesIndex)}
                    d={bottomFace}
                    fill={darkenColor(fillColor, 0.2)}
                    fill-opacity={0.34}
                    stroke={darkenColor(fillColor, 0.24)}
                    stroke-width={0.6}
                  />
                </>
              );
            })
          : [];
        const firstDefinedPoint = definedPoints[0] ?? null;
        const lastDefinedPoint = definedPoints[definedPoints.length - 1] ?? null;
        const startCap = isAreaChart && isStackedSeries && firstDefinedPoint?.topBack && firstDefinedPoint?.bottomBack
          ? buildLinearSvgPath([
              firstDefinedPoint.top,
              firstDefinedPoint.bottom,
              firstDefinedPoint.bottomBack,
              firstDefinedPoint.topBack
            ], true)
          : "";
        const endCap = isAreaChart && isStackedSeries && lastDefinedPoint?.topBack && lastDefinedPoint?.bottomBack
          ? buildLinearSvgPath([
              lastDefinedPoint.top,
              lastDefinedPoint.bottom,
              lastDefinedPoint.bottomBack,
              lastDefinedPoint.topBack
            ], true)
          : "";

        return (
          <g key={`line3d-series-${seriesIndex}`}>
            {isAreaChart && isStackedSeries && areaBackPoints.length >= 3 ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={buildLinearSvgPath(areaBackPoints, true)}
                fill={darkenColor(fillColor, 0.18)}
                fill-opacity={0.44}
                stroke={darkenColor(fillColor, 0.26)}
                stroke-width={0.8}
              />
            ) : null}
            {slabFaces}
            {startCap ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={startCap}
                fill={darkenColor(fillColor, 0.24)}
                fill-opacity={0.54}
                stroke={darkenColor(fillColor, 0.3)}
                stroke-width={0.7}
              />
            ) : null}
            {endCap ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={endCap}
                fill={darkenColor(fillColor, 0.14)}
                fill-opacity={0.6}
                stroke={darkenColor(fillColor, 0.24)}
                stroke-width={0.7}
              />
            ) : null}
            {isAreaChart && areaPoints.length >= 3 ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={buildLinearSvgPath(areaPoints, true)}
                fill={fillColor}
                fill-opacity={0.74}
                stroke={darkenColor(fillColor, 0.12)}
                stroke-width={0.9}
              />
            ) : null}
            {!isAreaChart
              ? definedPoints.map((point, pointIndex) => (
                  <line
                    key={`line3d-drop-${seriesIndex}-${pointIndex}`}
                    stroke={lightenColor(strokeColor, 0.44)}
                    stroke-dasharray="2 2"
                    stroke-width={0.9}
                    x1={point.top.x}
                    x2={point.bottom.x}
                    y1={point.top.y}
                    y2={point.bottom.y}
                  />
                ))
              : null}
            <path
              {...chartElementDataProps(seriesIndex)}
              d={buildLinearSvgPath(linePoints)}
              fill="none"
              stroke={strokeColor}
              stroke-linejoin="round"
              stroke-width={Math.max(1.8, chart.series[seriesIndex]?.lineWidthPx ?? 2)}
            />
            {markerPath.length > 0
              ? definedPoints.map((point, pointIndex) => (
                  <g
                    key={`line3d-marker-${seriesIndex}-${pointIndex}`}
                    transform={`translate(${toSvgNumber(point.top.x)}, ${toSvgNumber(point.top.y)})`}
                  >
                    <path
                      {...chartElementDataProps(seriesIndex, point.categoryIndex)}
                      d={markerPath}
                      fill={chart.series[seriesIndex]?.markerColor ?? fillColor}
                      stroke={chart.series[seriesIndex]?.markerLineColor ?? chart.chartAreaFillColor ?? palette.surface}
                      stroke-width={1}
                    />
                  </g>
                ))
              : null}
            {definedPoints.map((point) => (
              <circle
                {...chartElementDataProps(seriesIndex, point.categoryIndex)}
                cx={point.top.x}
                cy={point.top.y}
                fill="transparent"
                key={`line3d-hit-${seriesIndex}-${point.categoryIndex}`}
                r={7}
                stroke="none"
              />
            ))}
            {definedPoints.map((point) => (
              isSelectedChartPointOrSeries(selectedChartElement, chart.id, seriesIndex, point.categoryIndex)
                ? renderSelectionPointHandles(
                    `line3d-selection-${seriesIndex}-${point.categoryIndex}`,
                    [{ x: point.top.x, y: point.top.y }]
                  )
                : null
            ))}
          </g>
        );
      })}
      {yLabelPoints.map(({ point, tick }) => (
        <text
          key={`line3d-y-label-${tick}`}
          fill={labelColor}
          font-size={10}
          text-anchor="end"
          x={point.x - 6}
          y={point.y + 3}
        >
          {formatTickValue(tick)}
        </text>
      ))}
      {xLabelPoints.map(({ label, point }, index) => (
        <text
          key={`line3d-x-label-${index}`}
          fill={labelColor}
          font-size={10}
          text-anchor="middle"
          x={point.x}
          y={point.y + 14}
        >
          {label}
        </text>
      ))}
    </g>
  );
}

export function indexByName(series: XlsxChart["series"], matcher: RegExp) {
  const index = series.findIndex((entry) => matcher.test((entry.name ?? "").toLowerCase()));
  return index >= 0 ? index : null;
}

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

export function renderChartPlot(
  chart: XlsxChart,
  palette: ChartRendererPalette,
  layout: ChartLayout,
  chartType: string,
  selectedChartElement?: XlsxChartElementSelection | null
) {
  if (isComboChart(chart)) {
    return renderComboChart(chart, palette, layout, selectedChartElement) ?? renderUnsupported(chart, palette, layout, "Combo");
  }
  if (
    chartType === "ColumnClustered"
    || chartType === "ColumnStacked"
    || chartType === "ColumnPercentStacked"
    || chartType === "BarClustered"
    || chartType === "BarStacked"
    || chartType === "BarPercentStacked"
  ) {
    return renderBarChart(chart, palette, layout, chartType, selectedChartElement);
  }
  if (
    chartType === "Line"
    || chartType === "LineStacked"
    || chartType === "LinePercentStacked"
    || chartType === "Area"
    || chartType === "AreaStacked"
    || chartType === "AreaPercentStacked"
  ) {
    return renderLineOrAreaChart(chart, palette, layout, chartType, selectedChartElement);
  }
  if (chartType === "Scatter") {
    return renderScatterChart(chart, palette, layout, false, selectedChartElement);
  }
  if (chartType === "ScatterLines") {
    return renderScatterChart(chart, palette, layout, false, selectedChartElement);
  }
  if (chartType === "ScatterSmooth") {
    return renderScatterChart(chart, palette, layout, true, selectedChartElement);
  }
  if (chartType === "Bubble") {
    return renderBubbleChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Radar") {
    return renderRadarChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Pie" || chartType === "Pie3D" || chartType === "PieExploded" || chartType === "Doughnut") {
    return renderPieChart(chart, palette, layout, chartType, selectedChartElement);
  }
  if (chartType === "BarOfPie") {
    return renderBarOfPieChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Surface") {
    return renderSurfaceChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Stock") {
    return renderStockChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Waterfall") {
    return renderWaterfallChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Funnel") {
    return renderFunnelChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "BoxWhisker") {
    return renderBoxWhiskerChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Sunburst") {
    return renderSunburstChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Treemap") {
    return renderTreemapChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "RegionMap") {
    return renderRegionMapChart(chart, palette, layout, selectedChartElement);
  }
  return renderUnsupported(chart, palette, layout, chartType);
}

export const MemoChartSvg = /* @__PURE__ */ defineComponent({
  name: "MemoChartSvg",
  props: {
    chart: { type: Object as () => XlsxChart, required: true },
    onChartElementDoubleClick: { type: Function as unknown as () => ((selection: XlsxChartElementSelection, event: MouseEvent) => void) | undefined, default: undefined },
    onChartElementPointerDown: { type: Function as unknown as () => ((selection: XlsxChartElementSelection, event: PointerEvent) => void) | undefined, default: undefined },
    palette: { type: Object as () => ChartRendererPalette, required: true },
    rect: { type: Object as () => XlsxImageRect, required: true },
    selectedChartElement: { type: Object as () => XlsxChartElementSelection | null | undefined, default: null },
  },
  setup(props) {
    return () => {
      const { chart, onChartElementDoubleClick, onChartElementPointerDown, palette, rect, selectedChartElement } = props;
      const renderChartType = normalizeRenderableChartType(chart);
      const legendItems = getLegendItems(chart, renderChartType, palette);
      const layout = buildLayout(chart, rect, legendItems);
      const chartRaw = chart.raw && typeof chart.raw === "object" ? chart.raw as Record<string, unknown> : null;
      const explicitNoFill = chartRaw?.chartAreaNoFill === true || chartRaw?.plotAreaNoFill === true;
      const background = chart.chartAreaFillColor ?? (explicitNoFill ? "transparent" : "#ffffff");
      const borderColor = chart.chartAreaBorderColor ?? "transparent";
      const normalizedBackground = background.trim().toLowerCase();
      const normalizedBorderColor = borderColor.trim().toLowerCase();
      const hideBackgroundRect = normalizedBackground === "transparent" && normalizedBorderColor === "transparent";
      const fontFamily = buildChartFontFamily(chart.fontFamily);

      const handlePointerDown = (event: PointerEvent) => {
        const selection = resolveChartSelectionFromTarget(chart, event.target, selectedChartElement);
        if (!selection) {
          return;
        }
        if (resolveChartElementTarget(event.target)) {
          event.stopPropagation();
        }
        onChartElementPointerDown?.(selection, event);
      };

      const handleDoubleClick = (event: MouseEvent) => {
        const selection = resolveChartSelectionFromTarget(chart, event.target, selectedChartElement);
        if (!selection) {
          return;
        }
        if (resolveChartElementTarget(event.target)) {
          event.stopPropagation();
        }
        onChartElementDoubleClick?.(selection, event);
      };

      if (renderChartType === "Surface") {
        return (
          <MemoSurfaceChartComposite
            background={background}
            borderColor={borderColor}
            chart={chart}
            fallback={renderSurfaceChart(chart, palette, layout, selectedChartElement)}
            fontFamily={fontFamily}
            layout={layout}
            onDoubleClick={handleDoubleClick}
            onPointerDown={handlePointerDown}
            overlay={
              <>
                {renderSurfaceHitOverlay(chart, layout, selectedChartElement)}
                {renderSurfaceAxes(chart, layout)}
                {renderTitle(chart, layout, palette)}
                {renderLegend(chart, layout, palette)}
              </>
            }
            palette={palette}
          />
        );
      }

      return (
        <svg
          aria-label={chart.title ?? chart.name ?? "Chart"}
          onDblclick={handleDoubleClick}
          onPointerdown={handlePointerDown}
          role="img"
          style={{ display: "block", fontFamily, height: "100%", pointerEvents: "auto", width: "100%" }}
          viewBox={`0 0 ${layout.width} ${layout.height}`}
        >
          {hideBackgroundRect
            ? null
            : <rect fill={background} height={layout.height} stroke={borderColor} stroke-width={1} width={layout.width} x={0} y={0} />}
          {renderTitle(chart, layout, palette)}
          {renderLegend(chart, layout, palette)}
          {renderChartPlot(chart, palette, layout, renderChartType, selectedChartElement)}
        </svg>
      );
    };
  },
});
