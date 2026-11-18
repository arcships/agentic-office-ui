// Placeholder geometry and preview rendering for drawing objects.
//
// The document model carries shapes, charts, SmartArt, embedded OLE objects and
// 3D models, but the library has no renderer for them. What it *can* show is the
// content it parsed (text box text, chart data, the poster frame of a 3D model)
// and a labelled frame at the object's anchored position. These helpers keep
// that decision framework-free so the renderer and the tests share it.

import type {
  ChartData,
  DrawingAnchor,
  DrawingRunNode,
  TextStyle,
} from "../../engine/types";
import { escapeSvgText, svgDataUri } from "./synthetic-textbox";

const DEFAULT_DRAWING_WIDTH_PX = 320;
const DEFAULT_DRAWING_HEIGHT_PX = 180;
const CHART_PADDING_PX = 12;
const CHART_PALETTE = [
  "#4472C4",
  "#ED7D31",
  "#A5A5A5",
  "#FFC000",
  "#5B9BD5",
  "#70AD47",
];

function finitePositive(value: number | undefined, fallback: number): number {
  return Number.isFinite(value) && (value as number) > 0
    ? Math.round(value as number)
    : fallback;
}

/** Frame size of a drawing placeholder, in CSS pixels. */
export function drawingObjectSizePx(drawing: DrawingRunNode): {
  widthPx: number;
  heightPx: number;
} {
  return {
    widthPx: finitePositive(drawing.widthPx, DEFAULT_DRAWING_WIDTH_PX),
    heightPx: finitePositive(drawing.heightPx, DEFAULT_DRAWING_HEIGHT_PX),
  };
}

/** True when the object is painted out of the text flow (`wp:anchor`). */
export function drawingObjectIsFloating(drawing: DrawingRunNode): boolean {
  return Boolean(drawing.floating) && drawing.floating?.wrapType !== "topAndBottom";
}

/** True when the object must paint behind the paragraph text. */
export function drawingObjectIsBehindText(drawing: DrawingRunNode): boolean {
  return drawing.floating?.behindDocument === true;
}

/**
 * CSS that places a drawing placeholder. A text-wrapped object floats like an
 * image (`float: left/right`); a `wrapNone`/behind-text object is absolutely
 * positioned at its anchor offset. An inline object stays in the text flow.
 */
export function drawingObjectPlaceholderStyle(
  drawing: DrawingRunNode,
): Record<string, string | number | undefined> {
  const { widthPx, heightPx } = drawingObjectSizePx(drawing);
  const base: Record<string, string | number | undefined> = {
    width: `${widthPx}px`,
    height: `${heightPx}px`,
    boxSizing: "border-box",
    maxWidth: "100%",
  };
  const floating: DrawingAnchor | undefined = drawing.floating;
  if (!floating) {
    return { ...base, display: "inline-block", verticalAlign: "bottom" };
  }

  const wrapType = floating.wrapType;
  const zIndex = floating.behindDocument
    ? -1
    : Number.isFinite(floating.zIndex)
      ? Math.max(1, Math.min(65535, Math.round(floating.zIndex as number)))
      : 1;

  if (wrapType && wrapType !== "none" && wrapType !== "topAndBottom") {
    return {
      ...base,
      display: "inline-block",
      verticalAlign: "bottom",
      float: floating.horizontalAlign === "right" ? "right" : "left",
      position: "relative",
      zIndex,
      marginLeft: `${floating.distLPx ?? 0}px`,
      marginRight: `${floating.distRPx ?? 0}px`,
      marginTop: `${floating.distTPx ?? 0}px`,
      marginBottom: `${floating.distBPx ?? 0}px`,
    };
  }

  if (Number.isFinite(floating.xPx) || Number.isFinite(floating.yPx) || wrapType === "none") {
    return {
      ...base,
      display: "inline-block",
      position: "absolute",
      left: `${floating.xPx ?? 0}px`,
      top: `${floating.yPx ?? 0}px`,
      zIndex,
      pointerEvents: floating.behindDocument ? "none" : undefined,
    };
  }

  return {
    ...base,
    display: "inline-block",
    verticalAlign: "bottom",
    zIndex,
  };
}

/** Fill/stroke/border of the placeholder frame for one drawing object. */
export function drawingObjectFrameStyle(
  drawing: DrawingRunNode,
): Record<string, string | number | undefined> {
  const shape = drawing.shape;
  return {
    backgroundColor: shape?.fillColor ?? "rgba(241, 245, 249, 0.9)",
    borderColor: shape?.strokeColor ?? "#94a3b8",
    borderWidth: `${shape?.strokeWidthPx ?? 1}px`,
    borderStyle: "dashed",
    borderRadius: shape?.preset === "roundRect" || shape?.preset === "ellipse" ? "8px" : "2px",
    color: readTextColor(drawing),
    overflow: "hidden",
  };
}

function readTextColor(drawing: DrawingRunNode): string {
  const segmentStyle: TextStyle | undefined = drawing.textBox?.paragraphs[0]?.segments[0]?.style;
  return segmentStyle?.color ?? "#0f172a";
}

/** Wrapping style of the placeholder frame itself (flex centring). */
export function drawingObjectContentStyle(): Record<string, string> {
  return {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: "100%",
    height: "100%",
    textAlign: "center",
    fontSize: "12px",
    lineHeight: "1.3",
    padding: "4px 6px",
  };
}

/** Text of a chart placeholder: title, type and series count. */
export function chartPlaceholderText(chart: ChartData | undefined): string | undefined {
  if (!chart) return undefined;
  const parts: string[] = [];
  if (chart.title) parts.push(chart.title);
  if (chart.chartType) parts.push(chart.chartType.replace(/3D$/, ""));
  const seriesCount = chart.series?.length ?? 0;
  if (seriesCount > 0) parts.push(`${seriesCount} series`);
  return parts.length > 0 ? parts.join(" · ") : undefined;
}

/**
 * A small bar preview of the chart's first cached series, as an SVG data URI.
 * Returns `undefined` when the chart carries no cached values.
 */
export function chartPreviewDataUri(
  chart: ChartData | undefined,
  widthPx?: number,
  heightPx?: number,
): string | undefined {
  const values = chart?.series?.find((series) => (series.values?.length ?? 0) > 0)?.values;
  if (!chart || !values || values.length === 0) return undefined;

  const width = finitePositive(widthPx, DEFAULT_DRAWING_WIDTH_PX);
  const height = finitePositive(heightPx, DEFAULT_DRAWING_HEIGHT_PX);
  const titleHeight = chart.title ? 18 : 4;
  const plotTop = CHART_PADDING_PX + titleHeight;
  const plotBottom = height - CHART_PADDING_PX;
  const plotHeight = Math.max(10, plotBottom - plotTop);
  const plotLeft = CHART_PADDING_PX;
  const plotWidth = Math.max(10, width - CHART_PADDING_PX * 2);

  const numeric = values.filter((value): value is number => Number.isFinite(value));
  const max = numeric.length > 0 ? Math.max(...numeric) : 0;
  const min = numeric.length > 0 ? Math.min(0, ...numeric) : 0;
  const span = max - min || 1;
  const barWidth = Math.max(2, plotWidth / values.length - 2);

  const bars = values
    .map((value, index) => {
      if (!Number.isFinite(value)) return "";
      const ratio = ((value as number) - min) / span;
      const barHeight = Math.max(1, Math.round(plotHeight * ratio));
      const x = plotLeft + (plotWidth / values.length) * index + 1;
      const y = plotBottom - barHeight;
      const color = CHART_PALETTE[index % CHART_PALETTE.length];
      return `<rect x="${x.toFixed(1)}" y="${y}" width="${barWidth.toFixed(
        1,
      )}" height="${barHeight}" fill="${color}"/>`;
    })
    .join("");

  const title = chart.title
    ? `<text x="${CHART_PADDING_PX}" y="${CHART_PADDING_PX + 12}" font-size="12" font-weight="600" fill="#0f172a">${escapeSvgText(
        chart.title,
      )}</text>`
    : "";

  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">` +
    `<rect x="0" y="0" width="${width}" height="${height}" fill="#ffffff"/>` +
    title +
    `<line x1="${plotLeft}" y1="${plotBottom}" x2="${plotLeft + plotWidth}" y2="${plotBottom}" stroke="#cbd5e1"/>` +
    bars +
    `</svg>`;
  return svgDataUri(svg);
}
