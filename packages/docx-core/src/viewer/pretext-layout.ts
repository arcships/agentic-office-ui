import {
  clearCache,
  layoutNextLine,
  measureLineStats,
  prepareWithSegments,
  type LayoutCursor,
  type LayoutLine,
  type PreparedTextWithSegments,
} from "@chenglou/pretext";
import { textMeasurementGeneration } from "./text-measurement-generation";
import { registerFontMetricCache } from "./font-metrics";
import { measureCanvasGlyphAdvance } from "./font-advances";
import type { FontKerningMode } from "./font-kerning";
import { createDocxTextMeasurementContext } from "./font-features";
import {
  aggregateLineMetrics,
  emptyLineMetrics,
  padLineMetricsToHeight,
  resolveLineGridHeightPx,
  textLineMetricElement,
  withFixedLineBaseline,
  type LineMetricElement,
  type LineVerticalMetrics,
} from "./line-metrics";

const fragmentOffsetAdvancesByFragment = new WeakMap<
  PretextLineFragment,
  { generation: number; advances: number[] }
>();

// Pretext caches measured segment widths itself. A web font that arrives after
// the first pagination pass changes those widths, so drop its cache (and ours)
// on the same font-metric revision.
registerFontMetricCache(() => {
  clearCache();
});

type PretextWordBreak = "normal" | "keep-all";

export interface PretextLayoutItem {
  text: string;
  font: string;
  startOffset: number;
  endOffset: number;
  break?: "normal" | "never";
  wordBreak?: PretextWordBreak;
  /** Kerning/OpenType settings used for this item's advance widths. */
  fontKerning?: FontKerningMode;
  fontFeatureSettings?: string;
  /**
   * Ascent above the shared line baseline. Set for atomic inline objects
   * (images) and for runs in a different font size than the paragraph's.
   */
  ascentPx?: number;
  /** Descent below the shared line baseline. */
  descentPx?: number;
  /** Descent that must stay below the baseline (glyph ink, object bottom). */
  unsuppressibleDescentPx?: number;
  /**
   * True when the item can make the line box taller than the paragraph line
   * height: an inline object, or a run in a different font size. A
   * sub/superscript run only moves its baseline inside the existing line box.
   */
  hasCustomVerticalMetrics?: boolean;
  /**
   * Distance the run's own baseline sits above the line baseline (negative for
   * a subscript). Recorded per fragment so runs align by baseline.
   */
  baselineShiftPx?: number;
}

export interface PretextExclusionRect {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

export interface PretextLineFragment {
  text: string;
  width: number;
  x: number;
  intervalX: number;
  intervalWidth: number;
  startOffset: number;
  endOffset: number;
  font?: string;
  fontKerning?: FontKerningMode;
  fontFeatureSettings?: string;
  /** Ascent of this fragment above the shared line baseline. */
  ascentPx?: number;
  /** Descent of this fragment below the shared line baseline. */
  descentPx?: number;
  /** Descent of this fragment that cannot be trimmed at a page break. */
  unsuppressibleDescentPx?: number;
  /** Bottom space contributed by this fragment that a page break may trim. */
  suppressibleBottomSpacePx?: number;
  /** This fragment's metrics did not come from the font alone. */
  hasCustomVerticalMetrics?: boolean;
  /**
   * Distance this fragment's own baseline sits above the line baseline;
   * negative for a subscript. `lineBaselineYPx - baselineShiftPx` is where the
   * fragment's glyphs sit.
   */
  baselineShiftPx?: number;
}

export interface PretextLineLayout {
  y: number;
  fragments: PretextLineFragment[];
  /** Vertical metrics of the line box (ascent/descent/baseline). */
  metrics?: LineVerticalMetrics;
  /** Absolute baseline offset of the line's fragments, from the layout top. */
  baselineYPx?: number;
}

export interface PretextVariableWidthLayout {
  lineCount: number;
  height: number;
  lines: PretextLineLayout[];
  text?: string;
  font?: string;
  containerWidthPx?: number;
  lineHeightPx?: number;
  exclusions?: PretextExclusionRect[];
  /** Per-line line-box metrics, in visual order. */
  lineMetrics?: LineVerticalMetrics[];
  /**
   * True when at least one line grew past `lineHeightPx` because of an inline
   * object or a run in another font size.
   */
  hasCustomVerticalMetrics?: boolean;
}

export interface PretextSelectionRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

class NonRetainingMap<K, V> extends Map<K, V> {
  set(_key: K, _value: V): this {
    return this;
  }
}

// Kept as a Map-compatible internal bridge for pretext-items-layout. Values are
// deliberately not retained across calls; a viewer-owned cache can replace it later.
export const layoutByKey = new NonRetainingMap<string, PretextVariableWidthLayout>();

let graphemeSegmenter: Intl.Segmenter | undefined;


export function canUsePretext(): boolean {
  return (
    typeof OffscreenCanvas !== "undefined" || typeof document !== "undefined"
  );
}

export function getCachedValue<K, V>(cache: Map<K, V>, key: K): V | undefined {
  const cached = cache.get(key);
  if (cached === undefined) {
    return undefined;
  }

  cache.delete(key);
  cache.set(key, cached);
  return cached;
}

export function trimCache<K, V>(cache: Map<K, V>, maxEntries: number): void {
  while (cache.size > maxEntries) {
    const firstKey = cache.keys().next().value as K | undefined;
    if (firstKey === undefined) {
      break;
    }
    cache.delete(firstKey);
  }
}

function getMeasureContext():
  | OffscreenCanvasRenderingContext2D
  | CanvasRenderingContext2D
  | undefined {
  if (!canUsePretext()) {
    return undefined;
  }
  return createDocxTextMeasurementContext();
}

function measureTextWidthPx(
  font: string,
  text: string,
  fontKerning?: FontKerningMode,
  fontFeatureSettings?: string
): number {
  if (!text) {
    return 0;
  }

  const context = getMeasureContext();
  if (!context) {
    // No canvas (SSR, worker without OffscreenCanvas): keep selection offsets
    // usable with a deterministic font-size-based estimate instead of zero.
    const fontSizePx = Number.parseFloat(
      /(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? ""
    );
    const safeFontSizePx = Number.isFinite(fontSizePx) && fontSizePx > 0
      ? fontSizePx
      : 12;
    return Math.max(0, Math.round(text.length * safeFontSizePx * 0.5));
  }

  try {
    return Math.max(
      0,
      Math.round(
        measureCanvasGlyphAdvance(
          context,
          font,
          text,
          fontKerning,
          fontFeatureSettings
        )
      )
    );
  } catch {
    return Math.max(0, Math.round(text.length * 8));
  }
}

function measureOffsetWidthPx(
  font: string,
  text: string,
  offset: number,
  fontKerning?: FontKerningMode,
  fontFeatureSettings?: string
): number {
  if (offset <= 0 || !text) {
    return 0;
  }

  return measureTextWidthPx(
    font,
    text.slice(0, Math.max(0, Math.min(offset, text.length))),
    fontKerning,
    fontFeatureSettings
  );
}

function getGraphemeSegmenter(): Intl.Segmenter | undefined {
  if (graphemeSegmenter) {
    return graphemeSegmenter;
  }

  if (typeof Intl === "undefined" || typeof Intl.Segmenter === "undefined") {
    return undefined;
  }

  graphemeSegmenter = new Intl.Segmenter(undefined, {
    granularity: "grapheme",
  });
  return graphemeSegmenter;
}

function graphemeCodeUnitOffsets(text: string): number[] {
  if (!text) {
    return [0];
  }

  const segmenter = getGraphemeSegmenter();
  const offsets = [0];
  if (segmenter) {
    for (const grapheme of segmenter.segment(text)) {
      offsets.push(grapheme.index + grapheme.segment.length);
    }
  } else {
    let nextOffset = 0;
    for (const codePoint of text) {
      nextOffset += codePoint.length;
      offsets.push(nextOffset);
    }
  }
  if (offsets[offsets.length - 1] !== text.length) {
    offsets[offsets.length - 1] = text.length;
  }
  return offsets;
}

function countGraphemes(text: string): number {
  return Math.max(0, graphemeCodeUnitOffsets(text).length - 1);
}

function codeUnitOffsetAtGrapheme(text: string, graphemeIndex: number): number {
  const offsets = graphemeCodeUnitOffsets(text);
  const safeIndex = Math.max(
    0,
    Math.min(Math.round(graphemeIndex), offsets.length - 1)
  );
  return offsets[safeIndex] ?? text.length;
}

export function cursorAdvanceCodeUnits(
  prepared: PreparedTextWithSegments,
  start: LayoutCursor,
  end: LayoutCursor
): number {
  if (
    end.segmentIndex < start.segmentIndex ||
    (end.segmentIndex === start.segmentIndex &&
      end.graphemeIndex <= start.graphemeIndex)
  ) {
    return 0;
  }

  let consumedCodeUnits = 0;
  const lastSegmentIndex = Math.min(end.segmentIndex, prepared.segments.length);
  for (
    let segmentIndex = start.segmentIndex;
    segmentIndex <= lastSegmentIndex;
    segmentIndex += 1
  ) {
    const segmentText = prepared.segments[segmentIndex] ?? "";
    if (
      segmentIndex === start.segmentIndex &&
      segmentIndex === end.segmentIndex
    ) {
      consumedCodeUnits += Math.max(
        0,
        codeUnitOffsetAtGrapheme(segmentText, end.graphemeIndex) -
          codeUnitOffsetAtGrapheme(segmentText, start.graphemeIndex)
      );
      break;
    }
    if (segmentIndex === start.segmentIndex) {
      consumedCodeUnits += Math.max(
        0,
        segmentText.length -
          codeUnitOffsetAtGrapheme(segmentText, start.graphemeIndex)
      );
      continue;
    }
    if (segmentIndex === end.segmentIndex) {
      consumedCodeUnits += codeUnitOffsetAtGrapheme(
        segmentText,
        end.graphemeIndex
      );
      break;
    }
    consumedCodeUnits += segmentText.length;
  }

  return consumedCodeUnits;
}

export function layoutCacheKey(
  layoutSignature: string,
  containerWidthPx: number,
  lineHeightPx: number,
  exclusions: PretextExclusionRect[],
  docGridLinePitchPx?: number
): string {
  const exclusionsKey = exclusions
    .map(
      (exclusion) =>
        `${exclusion.left},${exclusion.right},${exclusion.top},${exclusion.bottom}`
    )
    .join(";");
  return `${layoutSignature}\u0000${containerWidthPx}\u0000${lineHeightPx}\u0000${docGridLinePitchPx ?? 0}\u0000${exclusionsKey}`;
}

/**
 * Vertical metrics contribution of a fragment. Fragments that carry explicit
 * metrics (inline objects, sub/superscript, mixed font sizes) keep them; plain
 * text fragments fall back to the metrics of their own font.
 */
export function fragmentLineMetricElement(
  fragment: PretextLineFragment,
  defaultFont?: string
): LineMetricElement {
  const hasExplicitMetrics =
    fragment.ascentPx !== undefined ||
    fragment.descentPx !== undefined ||
    fragment.unsuppressibleDescentPx !== undefined;
  if (hasExplicitMetrics) {
    const descentPx = fragment.descentPx ?? 0;
    return {
      ascentPx: fragment.ascentPx ?? 0,
      descentPx,
      unsuppressibleDescentPx: fragment.unsuppressibleDescentPx ?? descentPx,
      suppressibleBottomSpacePx: fragment.suppressibleBottomSpacePx,
    };
  }

  return textLineMetricElement(fragment.font ?? defaultFont ?? "");
}

/** Aggregated ascent/descent of every fragment on one line. */
export function aggregateFragmentLineMetrics(
  fragments: readonly PretextLineFragment[],
  defaultFont?: string
): LineVerticalMetrics {
  return aggregateLineMetrics(
    fragments.map((fragment) => fragmentLineMetricElement(fragment, defaultFont))
  );
}

export interface PretextLineBoxOptions {
  /** Paragraph line height: a line box never gets shorter than this. */
  minimumHeightPx: number;
  /** Snap the line box up to the document grid line pitch. */
  docGridLinePitchPx?: number;
  /**
   * Let content metrics grow the box past `minimumHeightPx`. Only content whose
   * metrics a font-only line box cannot express (inline objects, other font
   * sizes) may grow; everything else keeps the paragraph line height so page
   * breaks stay where the paragraph model puts them.
   */
  allowGrowth?: boolean;
}

/**
 * Resolve the line box of one laid-out line: content metrics plus the
 * paragraph's line height as the floor, snapped to the document grid.
 */
export function resolvePretextLineBoxMetrics(
  metrics: LineVerticalMetrics,
  options: PretextLineBoxOptions
): LineVerticalMetrics {
  const minimumHeightPx = Math.max(
    1,
    Math.round(
      Number.isFinite(options.minimumHeightPx) ? options.minimumHeightPx : 1
    )
  );
  if (!options.allowGrowth) {
    // The paragraph line height already models Word's single-space line box
    // (including its wrap compensation); keep it as the box height and let the
    // baseline sit inside it.
    return withFixedLineBaseline(metrics, minimumHeightPx);
  }

  return padLineMetricsToHeight(
    metrics,
    resolveLineGridHeightPx(
      Math.max(minimumHeightPx, Math.ceil(metrics.heightPx)),
      options.docGridLinePitchPx
    )
  );
}

/** Height of one line box, falling back to the layout line height. */
export function lineBoxHeightPx(
  line: PretextLineLayout | undefined,
  fallbackLineHeightPx = 0
): number {
  if (line?.metrics) {
    return Math.max(0, line.metrics.heightPx);
  }

  return Math.max(0, Math.round(fallbackLineHeightPx));
}

/** Absolute baseline offset of one line box, from the layout top. */
export function lineBaselineYPx(
  line: PretextLineLayout | undefined,
  fallbackLineHeightPx = 0
): number {
  if (!line) {
    return 0;
  }

  const ascentPx = line.metrics
    ? line.metrics.ascentPx
    : lineBoxHeightPx(line, fallbackLineHeightPx);
  return line.y + Math.max(0, ascentPx);
}

/**
 * Absolute baseline of one fragment: the line baseline shifted by the
 * fragment's own sub/superscript offset.
 */
export function fragmentBaselineYPx(
  line: PretextLineLayout | undefined,
  fragment: PretextLineFragment | undefined,
  fallbackLineHeightPx = 0
): number {
  const shiftPx = fragment?.baselineShiftPx ?? 0;
  return lineBaselineYPx(line, fallbackLineHeightPx) - shiftPx;
}

export function cachedFragmentOffsetAdvances(
  defaultFont: string,
  fragment: PretextLineFragment
): number[] {
  const measurementGeneration = textMeasurementGeneration();
  const cached = fragmentOffsetAdvancesByFragment.get(fragment);
  if (cached && cached.generation === measurementGeneration) {
    return cached.advances;
  }

  const advances = new Array<number>(fragment.text.length + 1);
  for (
    let localOffset = 0;
    localOffset <= fragment.text.length;
    localOffset += 1
  ) {
    advances[localOffset] = measureOffsetWidthPx(
      fragment.font ?? defaultFont,
      fragment.text,
      localOffset,
      fragment.fontKerning,
      fragment.fontFeatureSettings
    );
  }
  fragmentOffsetAdvancesByFragment.set(fragment, {
    generation: measurementGeneration,
    advances,
  });
  return advances;
}

export function prepareCached(
  text: string,
  font: string,
  wordBreak: PretextWordBreak = "normal"
): PreparedTextWithSegments | undefined {
  if (!canUsePretext()) {
    return undefined;
  }

  try {
    return prepareWithSegments(text, font, {
      whiteSpace: "pre-wrap",
      wordBreak,
    });
  } catch {
    return undefined;
  }
}

/**
 * Fast line-count-only path for plain single-font paragraphs with no
 * exclusions. Uses pretext's `measureLineStats` (added in 0.0.5) so we can
 * wrap text and count lines in pure arithmetic without allocating any line
 * text strings. Intended for hot pagination loops that only read
 * `lineCount` from the result and discard the rest.
 *
 * Returns `undefined` when pretext is not available in the host environment
 * (e.g. SSR without Canvas); callers should fall back to the general layout
 * path in that case.
 */
export function measurePretextPlainTextLineCount(
  text: string,
  font: string,
  containerWidthPx: number,
  options?: {
    wordBreak?: PretextWordBreak;
  }
): number | undefined {
  if (!text) {
    return 0;
  }

  const wordBreak = options?.wordBreak ?? "normal";
  const safeWidth = Math.max(1, Math.round(containerWidthPx));
  const prepared = prepareCached(text, font, wordBreak);
  if (!prepared) {
    return undefined;
  }

  try {
    return measureLineStats(prepared, safeWidth).lineCount;
  } catch {
    return undefined;
  }
}


function cursorIsDone(
  prepared: PreparedTextWithSegments,
  cursor: LayoutCursor
): boolean {
  return cursor.segmentIndex >= prepared.segments.length;
}

function cursorEndedAtHardBreak(
  prepared: PreparedTextWithSegments,
  cursor: LayoutCursor
): boolean {
  if (cursor.graphemeIndex > 0 || cursor.segmentIndex <= 0) {
    return false;
  }

  return prepared.kinds[cursor.segmentIndex - 1] === "hard-break";
}

function cursorSplitsLeadingBreakableSegment(
  prepared: PreparedTextWithSegments,
  start: LayoutCursor,
  end: LayoutCursor
): boolean {
  if (start.graphemeIndex !== 0) {
    return false;
  }

  const segmentText = prepared.segments[start.segmentIndex];
  const segmentGraphemeCount = segmentText ? countGraphemes(segmentText) : 0;
  if (segmentGraphemeCount <= 1) {
    return false;
  }

  return (
    end.segmentIndex === start.segmentIndex &&
    end.graphemeIndex > start.graphemeIndex &&
    end.graphemeIndex < segmentGraphemeCount
  );
}

function lineSplitsLeadingBreakableSegment(
  prepared: PreparedTextWithSegments,
  start: LayoutCursor,
  line: LayoutLine
): boolean {
  return cursorSplitsLeadingBreakableSegment(prepared, start, line.end);
}

function laterIntervalFitsLeadingSegmentWithoutSplit(
  prepared: PreparedTextWithSegments,
  start: LayoutCursor,
  laterIntervals: Array<{
    x: number;
    width: number;
  }>
): boolean {
  for (const interval of laterIntervals) {
    const candidate = layoutNextLine(prepared, start, interval.width);
    if (!candidate) {
      continue;
    }

    if (!lineSplitsLeadingBreakableSegment(prepared, start, candidate)) {
      return true;
    }
  }

  return false;
}

export function rowWidthsAtY(
  containerWidthPx: number,
  lineHeightPx: number,
  rowTopPx: number,
  exclusions: PretextExclusionRect[]
): Array<{
  x: number;
  width: number;
}> {
  const safeContainerWidthPx = Math.max(0, Math.round(containerWidthPx));
  let intervals = [
    {
      x: 0,
      width: safeContainerWidthPx,
    },
  ];

  const rowBottomPx = rowTopPx + Math.max(1, Math.round(lineHeightPx));
  for (const exclusion of exclusions) {
    const overlapsExclusion =
      rowBottomPx > exclusion.top && rowTopPx < exclusion.bottom;
    if (!overlapsExclusion) {
      continue;
    }

    const exclusionLeftPx = Math.max(
      0,
      Math.min(safeContainerWidthPx, Math.round(exclusion.left))
    );
    const exclusionRightPx = Math.max(
      exclusionLeftPx,
      Math.min(safeContainerWidthPx, Math.round(exclusion.right))
    );

    intervals = intervals.flatMap((interval) => {
      const intervalLeftPx = interval.x;
      const intervalRightPx = interval.x + interval.width;
      if (
        exclusionRightPx <= intervalLeftPx ||
        exclusionLeftPx >= intervalRightPx
      ) {
        return [interval];
      }

      const nextIntervals: Array<{ x: number; width: number }> = [];
      if (exclusionLeftPx > intervalLeftPx) {
        nextIntervals.push({
          x: intervalLeftPx,
          width: exclusionLeftPx - intervalLeftPx,
        });
      }
      if (exclusionRightPx < intervalRightPx) {
        nextIntervals.push({
          x: exclusionRightPx,
          width: intervalRightPx - exclusionRightPx,
        });
      }
      return nextIntervals;
    });
  }

  return intervals.filter((interval) => interval.width > 0.5);
}

export function layoutTextWithPretextAroundExclusions(
  text: string,
  font: string,
  containerWidthPx: number,
  lineHeightPx: number,
  exclusions?: PretextExclusionRect[],
  options?: {
    wordBreak?: PretextWordBreak;
  }
): PretextVariableWidthLayout | undefined {
  if (!text) {
    return {
      lineCount: 0,
      height: Math.max(
        0,
        ...(exclusions ?? []).map((exclusion) => Math.round(exclusion.bottom))
      ),
      lines: [],
      text,
      font,
      containerWidthPx: Math.max(1, Math.round(containerWidthPx)),
      lineHeightPx: Math.max(1, Math.round(lineHeightPx)),
      exclusions: (exclusions ?? []).map((exclusion) => ({
        left: Math.round(exclusion.left),
        right: Math.round(exclusion.right),
        top: Math.round(exclusion.top),
        bottom: Math.round(exclusion.bottom),
      })),
      lineMetrics: [],
      hasCustomVerticalMetrics: false,
    };
  }

  const wordBreak = options?.wordBreak ?? "normal";
  const prepared = prepareCached(text, font, wordBreak);
  if (!prepared) {
    return undefined;
  }

  const safeContainerWidthPx = Math.max(1, Math.round(containerWidthPx));
  const safeLineHeightPx = Math.max(1, Math.round(lineHeightPx));
  const normalizedExclusions = (exclusions ?? []).map((exclusion) => ({
    left: Math.round(exclusion.left),
    right: Math.round(exclusion.right),
    top: Math.round(exclusion.top),
    bottom: Math.round(exclusion.bottom),
  }));
  const lines: PretextLineLayout[] = [];

  let cursor: LayoutCursor = { segmentIndex: 0, graphemeIndex: 0 };
  let consumedOffset = 0;
  let rowTopPx = 0;

  while (!cursorIsDone(prepared, cursor)) {
    const rowIntervals = rowWidthsAtY(
      safeContainerWidthPx,
      safeLineHeightPx,
      rowTopPx,
      normalizedExclusions
    );
    const fragments: PretextLineFragment[] = [];

    if (rowIntervals.length === 0) {
      rowTopPx += safeLineHeightPx;
      continue;
    }

    for (
      let intervalIndex = 0;
      intervalIndex < rowIntervals.length;
      intervalIndex += 1
    ) {
      const interval = rowIntervals[intervalIndex]!;
      if (
        cursorIsDone(prepared, cursor) ||
        cursorEndedAtHardBreak(prepared, cursor)
      ) {
        break;
      }

      const line = layoutNextLine(prepared, cursor, interval.width);
      if (line) {
        if (
          lineSplitsLeadingBreakableSegment(prepared, cursor, line) &&
          laterIntervalFitsLeadingSegmentWithoutSplit(
            prepared,
            cursor,
            rowIntervals.slice(intervalIndex + 1)
          )
        ) {
          continue;
        }

        fragments.push({
          text: line.text,
          width: line.width,
          x: interval.x,
          intervalX: interval.x,
          intervalWidth: interval.width,
          startOffset: consumedOffset,
          endOffset:
            consumedOffset + cursorAdvanceCodeUnits(prepared, cursor, line.end),
          font,
        });
        consumedOffset += cursorAdvanceCodeUnits(prepared, cursor, line.end);
        cursor = line.end;
      }
    }

    if (fragments.length === 0) {
      break;
    }

    // A single-font paragraph has no custom vertical metrics: the line box keeps
    // the paragraph line height, but it still records the shared baseline and
    // the suppressible bottom space so page breaks can trim it.
    const lineBoxMetrics = resolvePretextLineBoxMetrics(
      aggregateFragmentLineMetrics(fragments, font),
      { minimumHeightPx: safeLineHeightPx }
    );
    lines.push({
      y: rowTopPx,
      fragments,
      metrics: lineBoxMetrics,
      baselineYPx: rowTopPx + lineBoxMetrics.ascentPx,
    });
    rowTopPx += lineBoxMetrics.heightPx;
  }

  const lineCount = lines.length;
  const lastLine = lines[lines.length - 1];
  const contentBottomPx = lastLine
    ? lastLine.y + lineBoxHeightPx(lastLine, safeLineHeightPx)
    : 0;
  const nextLayout: PretextVariableWidthLayout = {
    lineCount,
    height: Math.max(
      contentBottomPx,
      ...normalizedExclusions.map((exclusion) => Math.round(exclusion.bottom)),
      0
    ),
    lines,
    text,
    font,
    containerWidthPx: safeContainerWidthPx,
    lineHeightPx: safeLineHeightPx,
    exclusions: normalizedExclusions,
    lineMetrics: lines.map((line) => line.metrics ?? emptyLineMetrics()),
    hasCustomVerticalMetrics: false,
  };
  return nextLayout;
}
