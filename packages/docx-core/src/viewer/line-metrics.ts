// Inline vertical line metrics: ascent, descent, unsuppressible descent and
// suppressible bottom space.
//
// DOCX lays a line out from the glyph boxes of its runs, not from a fake
// `fontSizePt * factor` line box: a 28 pt inline image and a 12 pt run share
// one baseline, and the line box grows to the tallest ascent plus the deepest
// descent. Everything below the baseline that is not glyph ink (half-leading,
// the paragraph's extra line height, an inline object's blank bottom) is
// "suppressible": Word trims it when a line has to be rescued at a page
// break.
//
// Upstream: `packages/layout-engine/src/line-metrics.ts:43-135`
// (`aggregateLineMetrics`, `checkLineVerticalFit`, `constrainParagraphLineSplit`).

import { textMeasurementGeneration } from "./text-measurement-generation";

/**
 * Vertical metrics of one laid-out line box.
 *
 * `ascentPx + descentPx === heightPx`; `unsuppressibleDescentPx` is the part of
 * the descent that a page break may not trim, and `suppressibleBottomSpacePx`
 * is the bottom part that it may.
 */
export interface LineVerticalMetrics {
  /** Distance from the line box top to the shared baseline. */
  ascentPx: number;
  /** Distance from the shared baseline to the line box bottom. */
  descentPx: number;
  /** Descent that carries ink or an inline object and cannot be trimmed. */
  unsuppressibleDescentPx: number;
  /** Bottom space that a page break may trim to rescue the line. */
  suppressibleBottomSpacePx: number;
  /** Total line box height. */
  heightPx: number;
}

/**
 * One vertical contributor to a line box: a text run, an inline object or a
 * run with an explicit baseline shift.
 */
export interface LineMetricElement {
  /** Distance from the shared baseline up to the element's top. */
  ascentPx: number;
  /** Distance from the shared baseline down to the element's bottom. */
  descentPx?: number;
  /**
   * Part of `descentPx` that must stay below the baseline. Defaults to the
   * whole descent: glyph ink and inline objects cannot be trimmed.
   */
  unsuppressibleDescentPx?: number;
  /** Extra trimmable bottom space the element already knows about. */
  suppressibleBottomSpacePx?: number;
}

/** Font-level ascent/descent, measured from the font's bounding box. */
export interface FontVerticalMetrics {
  ascentPx: number;
  descentPx: number;
}

/** Result of testing a line box against the space left on a page. */
export interface LineVerticalFit {
  /** The line fits at its full height. */
  fits: boolean;
  /** Full line box height. */
  heightPx: number;
  /** Height left after trimming the suppressible bottom space. */
  suppressedHeightPx: number;
  /** Bottom space trimmed to make the line fit (0 when it fits at full height). */
  suppressedBottomSpacePx: number;
  /** The line only fits after trimming its suppressible bottom space. */
  rescued: boolean;
}

// Font metrics that a headless host (SSR, Worker without canvas) can still
// report deterministically. The pair matches the single-line-spacing heuristic
// used by `estimateParagraphLineHeightPx`, so a text-only line keeps the same
// height with and without canvas measurement.
const FALLBACK_FONT_ASCENT_EM = 0.9;
const FALLBACK_FONT_DESCENT_EM = 0.25;

/** Word raises a superscript baseline by a third of the parent font size. */
export const SCRIPT_SUPERSCRIPT_RAISE_EM = 0.33;
/** Word lowers a subscript baseline by a fifth of the parent font size. */
export const SCRIPT_SUBSCRIPT_LOWER_EM = 0.2;

const FONT_METRICS_LIMIT = 512;

let metricsCanvas: OffscreenCanvas | HTMLCanvasElement | undefined;
let metricsCanvasContext:
  | OffscreenCanvasRenderingContext2D
  | CanvasRenderingContext2D
  | null
  | undefined;

const fontMetricsByFont = new Map<
  string,
  { generation: number; metrics: FontVerticalMetrics }
>();

function positivePixelValue(value: number | undefined): number {
  return Number.isFinite(value) && (value as number) > 0 ? (value as number) : 0;
}

function roundMetric(value: number): number {
  return Math.round(value * 100) / 100;
}

function getMetricsContext():
  | OffscreenCanvasRenderingContext2D
  | CanvasRenderingContext2D
  | undefined {
  if (metricsCanvasContext !== undefined) {
    return metricsCanvasContext ?? undefined;
  }

  if (typeof OffscreenCanvas !== "undefined") {
    metricsCanvas = new OffscreenCanvas(1, 1);
    metricsCanvasContext = metricsCanvas.getContext("2d");
    return metricsCanvasContext ?? undefined;
  }

  if (typeof document !== "undefined") {
    metricsCanvas = document.createElement("canvas");
    metricsCanvasContext = metricsCanvas.getContext("2d");
    return metricsCanvasContext ?? undefined;
  }

  metricsCanvasContext = null;
  return undefined;
}

/** Font size in CSS pixels of a canvas font string (`... 16px "Calibri"`). */
export function fontSizePxFromFont(font: string): number {
  const match = /(\d+(?:\.\d+)?)px/.exec(font ?? "");
  const fontSizePx = match ? Number.parseFloat(match[1] ?? "") : Number.NaN;
  return Number.isFinite(fontSizePx) && fontSizePx > 0 ? fontSizePx : 12;
}

function measureFontVerticalMetrics(font: string): FontVerticalMetrics {
  const fontSizePx = fontSizePxFromFont(font);
  const fallback: FontVerticalMetrics = {
    ascentPx: roundMetric(fontSizePx * FALLBACK_FONT_ASCENT_EM),
    descentPx: roundMetric(fontSizePx * FALLBACK_FONT_DESCENT_EM),
  };

  const context = getMetricsContext();
  if (!context) {
    return fallback;
  }

  try {
    context.font = font;
    const measured = context.measureText("Hxg");
    const fontBoundingBoxAscent = measured.fontBoundingBoxAscent;
    const fontBoundingBoxDescent = measured.fontBoundingBoxDescent;
    const ascentPx =
      Number.isFinite(fontBoundingBoxAscent) && fontBoundingBoxAscent > 0
        ? fontBoundingBoxAscent
        : Number.isFinite(measured.actualBoundingBoxAscent) &&
          measured.actualBoundingBoxAscent > 0
        ? measured.actualBoundingBoxAscent
        : fallback.ascentPx;
    const descentPx =
      Number.isFinite(fontBoundingBoxDescent) && fontBoundingBoxDescent > 0
        ? fontBoundingBoxDescent
        : Number.isFinite(measured.actualBoundingBoxDescent) &&
          measured.actualBoundingBoxDescent > 0
        ? measured.actualBoundingBoxDescent
        : fallback.descentPx;
    return {
      ascentPx: roundMetric(Math.max(0, ascentPx)),
      descentPx: roundMetric(Math.max(0, descentPx)),
    };
  } catch {
    return fallback;
  }
}

/**
 * Ascent/descent of a canvas font string, measured with
 * `fontBoundingBoxAscent/Descent` and cached until the measurement generation
 * changes (web-font load, font-set change). Falls back to deterministic
 * font-size ratios when the host has no canvas.
 */
export function fontVerticalMetrics(font: string): FontVerticalMetrics {
  const safeFont = font ?? "";
  const generation = textMeasurementGeneration();
  const cached = fontMetricsByFont.get(safeFont);
  if (cached && cached.generation === generation) {
    return cached.metrics;
  }

  const metrics = measureFontVerticalMetrics(safeFont);
  fontMetricsByFont.delete(safeFont);
  fontMetricsByFont.set(safeFont, { generation, metrics });
  while (fontMetricsByFont.size > FONT_METRICS_LIMIT) {
    const oldestKey = fontMetricsByFont.keys().next().value as
      | string
      | undefined;
    if (oldestKey === undefined) {
      break;
    }
    fontMetricsByFont.delete(oldestKey);
  }
  return metrics;
}

/**
 * Vertical metrics contribution of a text run. The contribution is measured
 * from the run's own font, so a 65% sub/superscript run contributes its (small)
 * font box; its baseline shift is recorded separately and never grows the line.
 */
export function textLineMetricElement(font: string): LineMetricElement {
  const { ascentPx, descentPx } = fontVerticalMetrics(font);
  return {
    ascentPx,
    descentPx,
    unsuppressibleDescentPx: descentPx,
  };
}

/** Baseline shift of a superscript/subscript run at `parentFontSizePx`. */
export function scriptBaselineShiftPx(
  verticalAlign: "superscript" | "subscript" | undefined,
  parentFontSizePx: number
): number {
  const parentPx = Number.isFinite(parentFontSizePx) && parentFontSizePx > 0
    ? parentFontSizePx
    : 12;
  if (verticalAlign === "superscript") {
    return roundMetric(parentPx * SCRIPT_SUPERSCRIPT_RAISE_EM);
  }
  if (verticalAlign === "subscript") {
    return -roundMetric(parentPx * SCRIPT_SUBSCRIPT_LOWER_EM);
  }
  return 0;
}

/**
 * Vertical metrics contribution of an inline object. The object sits on the
 * shared baseline: its whole height is ascent and it has no descent, so it can
 * never be trimmed away at a page break.
 */
export function inlineObjectLineMetricElement(
  heightPx: number
): LineMetricElement {
  const safeHeightPx = Math.max(
    0,
    Math.round(Number.isFinite(heightPx) ? heightPx : 0)
  );
  return {
    ascentPx: safeHeightPx,
    descentPx: 0,
    unsuppressibleDescentPx: 0,
  };
}

export function emptyLineMetrics(heightPx = 0): LineVerticalMetrics {
  const safeHeightPx = Math.max(0, Math.round(heightPx));
  return {
    ascentPx: safeHeightPx,
    descentPx: 0,
    unsuppressibleDescentPx: 0,
    suppressibleBottomSpacePx: 0,
    heightPx: safeHeightPx,
  };
}

/**
 * Aggregate the vertical metrics of every element on one line.
 *
 * The line's ascent is the tallest ascent above the shared baseline and its
 * descent the deepest descent below it; the surplus descent that carries no
 * ink becomes suppressible bottom space.
 */
export function aggregateLineMetrics(
  elements: readonly LineMetricElement[]
): LineVerticalMetrics {
  let ascentPx = 0;
  let descentPx = 0;
  let unsuppressibleDescentPx = 0;
  let declaredSuppressiblePx = 0;

  for (const element of elements) {
    if (!element) {
      continue;
    }
    const elementAscentPx = positivePixelValue(element.ascentPx);
    const elementDescentPx = positivePixelValue(element.descentPx);
    const elementUnsuppressiblePx =
      element.unsuppressibleDescentPx === undefined
        ? elementDescentPx
        : Math.min(
            positivePixelValue(element.unsuppressibleDescentPx),
            elementDescentPx
          );

    ascentPx = Math.max(ascentPx, elementAscentPx);
    descentPx = Math.max(descentPx, elementDescentPx);
    unsuppressibleDescentPx = Math.max(
      unsuppressibleDescentPx,
      elementUnsuppressiblePx
    );
    declaredSuppressiblePx = Math.max(
      declaredSuppressiblePx,
      positivePixelValue(element.suppressibleBottomSpacePx)
    );
  }

  const safeUnsuppressibleDescentPx = Math.min(
    unsuppressibleDescentPx,
    descentPx
  );
  const suppressibleBottomSpacePx = Math.max(
    descentPx - safeUnsuppressibleDescentPx,
    declaredSuppressiblePx
  );

  // Declared trimmable space lives inside the descent, so the box always
  // contains the bottom space it advertises.
  const nextDescentPx = Math.max(
    descentPx,
    safeUnsuppressibleDescentPx + suppressibleBottomSpacePx
  );

  return {
    ascentPx,
    descentPx: nextDescentPx,
    unsuppressibleDescentPx: safeUnsuppressibleDescentPx,
    suppressibleBottomSpacePx,
    heightPx: ascentPx + nextDescentPx,
  };
}

/**
 * Grow a line box to at least `minimumHeightPx` by adding the difference to the
 * bottom. The added space is leading, so it stays suppressible.
 */
export function padLineMetricsToHeight(
  metrics: LineVerticalMetrics,
  minimumHeightPx: number
): LineVerticalMetrics {
  const safeMinimumHeightPx = Number.isFinite(minimumHeightPx)
    ? Math.max(0, minimumHeightPx)
    : 0;
  const ascentPx = Math.max(0, metrics.ascentPx);
  const descentPx = Math.max(metrics.descentPx, 0);
  const heightPx = Math.max(ascentPx + descentPx, safeMinimumHeightPx);
  const nextDescentPx = heightPx - ascentPx;
  const unsuppressibleDescentPx = Math.min(
    Math.max(0, metrics.unsuppressibleDescentPx),
    nextDescentPx
  );

  return {
    ascentPx,
    descentPx: nextDescentPx,
    unsuppressibleDescentPx,
    suppressibleBottomSpacePx: nextDescentPx - unsuppressibleDescentPx,
    heightPx,
  };
}

/**
 * Re-seat a line inside a fixed line box (`w:lineRule="exact"`): every line of
 * the paragraph has the same height, so the baseline moves with the largest
 * ascent while the rest of the box becomes suppressible bottom space.
 */
export function withFixedLineBaseline(
  metrics: LineVerticalMetrics,
  fixedLineHeightPx: number
): LineVerticalMetrics {
  const heightPx = Math.max(
    1,
    Math.round(
      Number.isFinite(fixedLineHeightPx) ? fixedLineHeightPx : metrics.heightPx
    )
  );
  const ascentPx = Math.min(
    Math.max(0, Math.round(metrics.ascentPx)),
    heightPx
  );
  const descentPx = heightPx - ascentPx;
  const unsuppressibleDescentPx = Math.min(
    Math.max(0, metrics.unsuppressibleDescentPx),
    descentPx
  );

  return {
    ascentPx,
    descentPx,
    unsuppressibleDescentPx,
    suppressibleBottomSpacePx: descentPx - unsuppressibleDescentPx,
    heightPx,
  };
}

/**
 * Snap a line height up to the document grid (`w:docGrid w:linePitch`). Without
 * a grid pitch the height is only rounded to whole pixels.
 */
export function resolveLineGridHeightPx(
  heightPx: number,
  docGridLinePitchPx?: number
): number {
  const safeHeightPx = Math.max(
    1,
    Math.round(Number.isFinite(heightPx) ? heightPx : 1)
  );
  const pitchPx =
    Number.isFinite(docGridLinePitchPx) && (docGridLinePitchPx as number) > 0
      ? Math.round(docGridLinePitchPx as number)
      : 0;
  if (pitchPx <= 0) {
    return safeHeightPx;
  }

  return Math.max(pitchPx, Math.ceil(safeHeightPx / pitchPx) * pitchPx);
}

/**
 * Test a line box against the space left on the current page. A line that does
 * not fit at full height is rescued when trimming its suppressible bottom space
 * is enough (Word's behaviour for the last line on a page).
 */
export function checkLineVerticalFit(
  metrics: LineVerticalMetrics,
  availableHeightPx: number
): LineVerticalFit {
  const heightPx = Math.max(0, metrics.heightPx);
  const suppressibleBottomSpacePx = Math.max(
    0,
    metrics.suppressibleBottomSpacePx
  );
  const suppressedHeightPx = Math.max(
    0,
    heightPx - suppressibleBottomSpacePx
  );
  const safeAvailableHeightPx = Number.isFinite(availableHeightPx)
    ? Math.max(0, availableHeightPx)
    : 0;

  if (heightPx <= safeAvailableHeightPx) {
    return {
      fits: true,
      heightPx,
      suppressedHeightPx,
      suppressedBottomSpacePx: 0,
      rescued: false,
    };
  }

  if (suppressedHeightPx <= safeAvailableHeightPx) {
    return {
      fits: true,
      heightPx,
      suppressedHeightPx,
      suppressedBottomSpacePx: heightPx - suppressedHeightPx,
      rescued: true,
    };
  }

  return {
    fits: false,
    heightPx,
    suppressedHeightPx,
    suppressedBottomSpacePx: 0,
    rescued: false,
  };
}

/** Sum of the line box heights in `[startLineIndex, endLineIndex)`. */
export function sumLineHeightsPx(
  lineMetrics: readonly LineVerticalMetrics[] | undefined,
  startLineIndex: number,
  endLineIndex: number
): number {
  if (!lineMetrics || lineMetrics.length === 0) {
    return 0;
  }

  const safeStart = Math.max(
    0,
    Math.min(Math.round(startLineIndex), lineMetrics.length)
  );
  const safeEnd = Math.max(
    safeStart,
    Math.min(Math.round(endLineIndex), lineMetrics.length)
  );
  let heightPx = 0;
  for (let index = safeStart; index < safeEnd; index += 1) {
    heightPx += Math.max(0, lineMetrics[index]?.heightPx ?? 0);
  }
  return heightPx;
}

/**
 * Largest line index (exclusive) of a paragraph that still fits in
 * `availableHeightPx`, given per-line metrics.
 *
 * Every line is placed at full height except the last one, which may be rescued
 * by its suppressible bottom space — that is the only place a page break can
 * legitimately trim a DOCX line box.
 */
export function constrainParagraphLineSplit(
  lineMetrics: readonly LineVerticalMetrics[],
  startLineIndex: number,
  maxEndLineIndex: number,
  availableHeightPx: number
): number {
  if (!lineMetrics || lineMetrics.length === 0) {
    return Math.max(0, Math.round(startLineIndex));
  }

  const safeStart = Math.max(
    0,
    Math.min(Math.round(startLineIndex), lineMetrics.length)
  );
  const safeMaxEnd = Math.max(
    safeStart,
    Math.min(Math.round(maxEndLineIndex), lineMetrics.length)
  );
  const safeAvailableHeightPx = Number.isFinite(availableHeightPx)
    ? Math.max(0, availableHeightPx)
    : 0;

  let consumedHeightPx = 0;
  let endLineIndex = safeStart;
  while (endLineIndex < safeMaxEnd) {
    const lineHeightPx = Math.max(0, lineMetrics[endLineIndex]?.heightPx ?? 0);
    if (consumedHeightPx + lineHeightPx <= safeAvailableHeightPx) {
      consumedHeightPx += lineHeightPx;
      endLineIndex += 1;
      continue;
    }

    const fit = checkLineVerticalFit(
      lineMetrics[endLineIndex] as LineVerticalMetrics,
      safeAvailableHeightPx - consumedHeightPx
    );
    if (fit.fits) {
      endLineIndex += 1;
    }
    break;
  }

  return endLineIndex;
}
