/**
 * Advance-width cache for DOCX text measurement.
 *
 * Widths are keyed by the full canvas font string — family, weight, style and
 * size — plus the kerning mode, the OpenType feature string and the run text.
 * The canvas is the primary source; a hidden DOM span is the fallback for hosts
 * that expose a document but no canvas, and a deterministic glyph estimate
 * keeps headless layout usable when neither exists.
 *
 * `registerFontMetricCache()` clears the cache when `document.fonts` loads a
 * face or the rendering profile changes.
 */

import { registerFontMetricCache } from "./font-metrics";
import type { FontKerningMode } from "./font-kerning";
import { setCanvasFontFeatureSettings } from "./font-features";
import { createDocxTextMeasurementContext } from "./font-features";
import { textMeasurementGeneration } from "./text-measurement-generation";

type CanvasAdvanceContext = Pick<
  CanvasRenderingContext2D,
  "font" | "measureText"
> & {
  letterSpacing?: string;
  wordSpacing?: string;
  fontKerning?: string;
  direction?: string;
  textRendering?: string;
  canvas?: HTMLCanvasElement | OffscreenCanvas;
};

const MAX_FONT_ADVANCE_SCALES = 8192;
const MAX_FONT_ADVANCE_WIDTHS = 16384;
const advanceScaleByFont = new Map<string, number>();
const advanceWidthByKey = new Map<string, number>();
const FONT_SIZE_RE =
  /(?:^|\s)([+]?(?:\d+(?:\.\d*)?|\.\d+))(px|pt)(?=\s|\/|$)/i;
const ADVANCE_PRECISION_PROBE = "Hamburgefontsiv 0123456789";

registerFontMetricCache(() => {
  advanceScaleByFont.clear();
  advanceWidthByKey.clear();
});

function isDesktopChromium(): boolean {
  if (typeof navigator === "undefined") return false;
  const userAgent = navigator.userAgent;
  return (
    /(?:Chrome|Chromium|Edg)\/\d/.test(userAgent) &&
    !/Android|Mobile|iPhone|iPad|iPod|CriOS/i.test(userAgent)
  );
}

/**
 * Restores nominal advances only when the current font measurement truncates
 * its size. Desktop Chromium measures a fractional font size at a truncated
 * integer advance; without this correction a 10.5 pt run is 1-2 % narrow.
 */
export function canvasFontAdvanceScale(
  context: CanvasAdvanceContext,
  font: string
): number {
  if (!isDesktopChromium()) return 1;
  const sizeMatch = FONT_SIZE_RE.exec(font);
  if (!sizeMatch) return 1;
  const requestedSizePx =
    Number(sizeMatch[1]) * (sizeMatch[2]?.toLowerCase() === "pt" ? 96 / 72 : 1);
  const effectiveSizePx =
    Math.floor(Math.fround(Math.fround(requestedSizePx) * 100)) / 100;
  if (
    !Number.isFinite(requestedSizePx) ||
    effectiveSizePx <= 0 ||
    effectiveSizePx >= requestedSizePx
  ) {
    return 1;
  }

  const cacheKey = `${font}\u0000${context.fontKerning ?? ""}\u0000${
    context.direction ?? ""
  }\u0000${context.textRendering ?? ""}\u0000${
    context.canvas && "style" in context.canvas
      ? context.canvas.style.fontFeatureSettings
      : ""
  }`;
  const cached = advanceScaleByFont.get(cacheKey);
  if (cached !== undefined) return cached;

  const sizeStart = sizeMatch.index + sizeMatch[0].indexOf(sizeMatch[1]!);
  const sizeEnd = sizeStart + sizeMatch[1]!.length + sizeMatch[2]!.length;
  const effectiveFont = `${font.slice(0, sizeStart)}${effectiveSizePx}px${font.slice(
    sizeEnd
  )}`;
  const originalFont = context.font;
  const originalLetterSpacing = context.letterSpacing;
  const originalWordSpacing = context.wordSpacing;
  let scale = 1;

  try {
    if (typeof originalLetterSpacing === "string") context.letterSpacing = "0px";
    if (typeof originalWordSpacing === "string") context.wordSpacing = "0px";
    context.font = font;
    const requestedWidth = context.measureText(ADVANCE_PRECISION_PROBE).width;
    context.font = effectiveFont;
    const effectiveWidth = context.measureText(ADVANCE_PRECISION_PROBE).width;
    const tolerance =
      Number.EPSILON * Math.max(1, requestedWidth, effectiveWidth) * 8;
    if (
      requestedWidth > 0 &&
      Number.isFinite(requestedWidth) &&
      Number.isFinite(effectiveWidth) &&
      Math.abs(requestedWidth - effectiveWidth) <= tolerance
    ) {
      scale = requestedSizePx / effectiveSizePx;
    }
  } catch {
    scale = 1;
  } finally {
    context.font = originalFont;
    if (typeof originalLetterSpacing === "string") {
      context.letterSpacing = originalLetterSpacing;
    }
    if (typeof originalWordSpacing === "string") {
      context.wordSpacing = originalWordSpacing;
    }
  }

  advanceScaleByFont.set(cacheKey, scale);
  if (advanceScaleByFont.size > MAX_FONT_ADVANCE_SCALES) {
    const firstKey = advanceScaleByFont.keys().next().value;
    if (firstKey !== undefined) advanceScaleByFont.delete(firstKey);
  }
  return scale;
}

/**
 * Measure one string with explicit kerning/feature settings. The caller's
 * letter/word spacing is restored so shared measurement contexts do not leak
 * settings between runs.
 */
export function measureCanvasGlyphAdvance(
  context: CanvasAdvanceContext,
  font: string,
  text: string,
  fontKerning: FontKerningMode = "auto",
  fontFeatureSettings?: string
): number {
  setCanvasFontFeatureSettings(context, fontFeatureSettings);
  context.font = font;
  if (typeof context.fontKerning === "string") {
    context.fontKerning = fontKerning;
  }
  const scale = canvasFontAdvanceScale(context, font);
  const originalLetterSpacing = context.letterSpacing;
  const originalWordSpacing = context.wordSpacing;
  try {
    if (typeof originalLetterSpacing === "string") context.letterSpacing = "0px";
    if (typeof originalWordSpacing === "string") context.wordSpacing = "0px";
    return context.measureText(text).width * scale;
  } finally {
    if (typeof originalLetterSpacing === "string") {
      context.letterSpacing = originalLetterSpacing;
    }
    if (typeof originalWordSpacing === "string") {
      context.wordSpacing = originalWordSpacing;
    }
  }
}

export interface FontAdvanceOptions {
  fontKerning?: FontKerningMode;
  fontFeatureSettings?: string;
}

/** Cache key for one measured advance; exposed so tests can assert the key. */
export function fontAdvanceCacheKey(
  font: string,
  text: string,
  options?: FontAdvanceOptions
): string {
  return `${textMeasurementGeneration()}\u0000${font}\u0000${
    options?.fontKerning ?? ""
  }\u0000${options?.fontFeatureSettings ?? ""}\u0000${text}`;
}

function fontSizePxFromFont(font: string): number {
  const match = /(\d+(?:\.\d+)?)px/.exec(font ?? "");
  const value = match ? Number.parseFloat(match[1] ?? "") : Number.NaN;
  return Number.isFinite(value) && value > 0 ? value : 12;
}

/**
 * Deterministic fallback for hosts without canvas or DOM. The ratios match
 * `estimatedGlyphWidthPx`, so a canvas-less layout keeps the same line breaks
 * as the selection/thumbnail fallback paths.
 */
function fallbackAdvancePx(font: string, text: string): number {
  const fontSizePx = fontSizePxFromFont(font);
  let widthPx = 0;
  for (const character of text) {
    if (/\s/.test(character)) {
      widthPx += fontSizePx * 0.34;
    } else if (/[A-Z]/.test(character)) {
      widthPx += fontSizePx * 0.66;
    } else if (/[a-z]/.test(character)) {
      widthPx += fontSizePx * 0.55;
    } else if (/[0-9]/.test(character)) {
      widthPx += fontSizePx * 0.57;
    } else if (/[\u2e80-\u9fff\u3040-\u30ff\uac00-\ud7af]/.test(character)) {
      widthPx += fontSizePx;
    } else {
      widthPx += fontSizePx * 0.48;
    }
  }
  return widthPx;
}

/**
 * DOM fallback used when a host has a document but no working canvas. The span
 * is inline-block and absolutely positioned so it never affects layout.
 */
export function measureTextWidthWithDom(
  font: string,
  text: string,
  options?: FontAdvanceOptions
): number | undefined {
  if (typeof document === "undefined" || !text) {
    return undefined;
  }
  const body = document.body;
  if (!body || typeof document.createElement !== "function") {
    return undefined;
  }

  try {
    const span = document.createElement("span");
    span.textContent = text;
    span.setAttribute("aria-hidden", "true");
    const style = span.style;
    style.position = "absolute";
    style.left = "-10000px";
    style.top = "0";
    style.visibility = "hidden";
    style.pointerEvents = "none";
    style.whiteSpace = "pre";
    style.font = font;
    if (options?.fontKerning) {
      style.fontKerning = options.fontKerning;
    }
    if (options?.fontFeatureSettings) {
      style.fontFeatureSettings = options.fontFeatureSettings;
    }
    body.appendChild(span);
    const width = span.getBoundingClientRect().width;
    body.removeChild(span);
    return Number.isFinite(width) && width >= 0 ? width : undefined;
  } catch {
    return undefined;
  }
}

/**
 * Measured advance of `text` in `font`. The canvas is preferred, then the DOM,
 * then the deterministic estimate. Results are cached until the font-metric
 * revision changes.
 */
export function fontAdvanceWidthPx(
  font: string,
  text: string,
  options?: FontAdvanceOptions
): number {
  if (!text) {
    return 0;
  }

  const key = fontAdvanceCacheKey(font, text, options);
  const cached = advanceWidthByKey.get(key);
  if (cached !== undefined) {
    return cached;
  }

  let width: number | undefined;
  const context = createDocxTextMeasurementContext();
  if (context) {
    try {
      width = measureCanvasGlyphAdvance(
        context,
        font,
        text,
        options?.fontKerning,
        options?.fontFeatureSettings
      );
    } catch {
      width = undefined;
    }
  }
  if (width === undefined || !Number.isFinite(width) || width < 0) {
    width = measureTextWidthWithDom(font, text, options);
  }
  if (width === undefined || !Number.isFinite(width) || width < 0) {
    width = fallbackAdvancePx(font, text);
  }

  advanceWidthByKey.set(key, width);
  if (advanceWidthByKey.size > MAX_FONT_ADVANCE_WIDTHS) {
    const firstKey = advanceWidthByKey.keys().next().value;
    if (firstKey !== undefined) advanceWidthByKey.delete(firstKey);
  }
  return width;
}

export interface FontMetricSpec {
  family: string;
  weight: number;
  style: "normal" | "italic";
  sizePx: number;
}

/** Decompose a CSS canvas font string into the family/weight/style it names. */
export function parseCanvasFontSpec(font: string): FontMetricSpec {
  const sizeMatch = /(\d+(?:\.\d+)?)px/.exec(font ?? "");
  const sizePx = sizeMatch ? Number.parseFloat(sizeMatch[1] ?? "") : 12;
  const weightMatch = /(?:^|\s)([1-9]00)(?=\s|$)/.exec(font ?? "");
  const weight = weightMatch ? Number(weightMatch[1]) : 400;
  const style = /(?:^|\s)italic(?=\s|$)/i.test(font ?? "") ? "italic" : "normal";
  const familyStart = sizeMatch
    ? (sizeMatch.index ?? 0) + sizeMatch[0].length
    : 0;
  const family = (font ?? "").slice(familyStart).trim() || "sans-serif";
  return { family, weight, style, sizePx };
}

export interface FontMetrics {
  font: string;
  family: string;
  weight: number;
  style: "normal" | "italic";
  sizePx: number;
  kerning: FontKerningMode;
  featureSettings: string | undefined;
  advancePx(text: string): number;
}

/**
 * Metrics object for one resolved font. The issue's `fontMetrics` cache is
 * this factory plus `fontAdvanceWidthPx`'s keyed map: same family, weight,
 * style, size, kerning mode and feature string ⇒ same cached advances.
 */
export function resolveFontMetrics(
  font: string,
  options?: FontAdvanceOptions
): FontMetrics {
  const spec = parseCanvasFontSpec(font);
  return {
    font,
    ...spec,
    kerning: options?.fontKerning ?? "auto",
    featureSettings: options?.fontFeatureSettings,
    advancePx: (text: string): number =>
      fontAdvanceWidthPx(font, text, options),
  };
}
