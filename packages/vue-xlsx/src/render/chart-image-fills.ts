import type { XlsxChart } from "@arcships/xlsx-core";

/**
 * Relationship between the chart model's picture fills and the SVG pipeline.
 *
 * The chart model stores picture fills as data URLs (`a:blipFill` resolved by
 * `readChartEmbeddedMedia`). An SVG cannot paint a `<rect>` with a data URL
 * directly, so every image becomes one `<pattern>` in the chart's `<defs>` and
 * the fill is a `url(#…)` reference. Ids are derived from the chart id plus a
 * stable slot key, which keeps two charts on one sheet from sharing a pattern.
 */

/** Only self-contained raster data URLs are accepted as an image source. */
const CHART_IMAGE_DATA_URL_PATTERN = /^data:image\/(?:png|jpe?g|gif|webp|bmp|tiff);base64,[A-Za-z0-9+/]+={0,2}$/;

export const CHART_AREA_IMAGE_FILL_KEY = "chart-area";
export const CHART_PLOT_AREA_IMAGE_FILL_KEY = "plot-area";

/**
 * Returns the data URL when it is a raster image this renderer may paint.
 *
 * The chart parser already sniffs every embedded part, but the model is public:
 * a host that sets `chart.chartAreaFillImage` by hand could otherwise smuggle a
 * `javascript:`/`data:text/html` URL into `<image href>`/`<pattern>`.
 */
export function sanitizeChartImageSource(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  if (!trimmed || trimmed !== value) {
    return undefined;
  }
  return CHART_IMAGE_DATA_URL_PATTERN.test(trimmed) ? trimmed : undefined;
}

function sanitizeIdPart(value: string) {
  return value.replace(/[^A-Za-z0-9_-]+/g, "-").replace(/^-+|-+$/g, "") || "chart";
}

export function chartSeriesImageFillKey(seriesIndex: number) {
  return `series-${seriesIndex}`;
}

export function chartPointImageFillKey(seriesIndex: number, pointIndex: number) {
  return `series-${seriesIndex}-point-${pointIndex}`;
}

/** Stable SVG id for one picture fill of a chart. */
export function chartImageFillId(chart: Pick<XlsxChart, "id">, key: string) {
  return `xlsx-chart-image-${sanitizeIdPart(chart.id)}-${sanitizeIdPart(key)}`;
}

/** `fill` attribute value (`url(#…)`) for one picture fill. */
export function chartImageFillUrl(chart: Pick<XlsxChart, "id">, key: string) {
  return `url(#${chartImageFillId(chart, key)})`;
}

export interface ChartImageFill {
  key: string;
  url: string;
}

/**
 * Every picture fill the chart currently carries, in paint order.
 *
 * Chart area first, then plot area, then series and their data points — the same
 * order the SVG renders them, so the first definition wins if a document names
 * the same slot twice.
 */
export function collectChartImageFills(chart: XlsxChart): ChartImageFill[] {
  const fills: ChartImageFill[] = [];
  const seen = new Set<string>();
  const push = (key: string, source: unknown) => {
    const url = sanitizeChartImageSource(source);
    if (!url || seen.has(key)) {
      return;
    }
    seen.add(key);
    fills.push({ key, url });
  };

  push(CHART_AREA_IMAGE_FILL_KEY, chart.chartAreaFillImage);
  push(CHART_PLOT_AREA_IMAGE_FILL_KEY, chart.plotAreaFillImage);
  for (const [seriesIndex, series] of (chart.series ?? []).entries()) {
    push(chartSeriesImageFillKey(seriesIndex), series.fillImage);
    for (const pointStyle of series.dataPointStyles ?? []) {
      push(chartPointImageFillKey(seriesIndex, pointStyle.index), pointStyle.fillImage);
    }
  }
  return fills;
}
