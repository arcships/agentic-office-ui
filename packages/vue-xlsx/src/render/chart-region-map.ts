/** @jsxImportSource vue */
import { feature as topojsonFeature } from "topojson-client";
import countiesAlbers10m from "us-atlas/counties-albers-10m.json";
import countries50m from "world-atlas/countries-50m.json";
import type { FeatureCollection, Geometry } from "geojson";
import type { XlsxChart, XlsxChartSeries } from "@extend-ai/xlsx-core";
import { clamp, mixRgbColor, lightenColor, darkenColor, normalizeRendererHexColor, safeNumber } from "./chart-shared";
import { chartPointColor } from "./chart-data";
import { formatTickValue } from "./chart-axis";
import { buildNumericTickValues } from "./chart-axis";
import type { RegionMapFeature, LegendItem } from "./chart-types";

export type { RegionMapFeature } from "./chart-types";

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
