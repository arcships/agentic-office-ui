import {
  layoutNextLine,
  measureLineStats,
  prepareWithSegments,
  type LayoutCursor,
  type LayoutLine,
  type PreparedTextWithSegments,
} from "@chenglou/pretext";

const PREPARED_TEXT_CACHE_MAX_ENTRIES = 8192;
const LAYOUT_CACHE_MAX_ENTRIES = 4096;
const LINE_COUNT_CACHE_MAX_ENTRIES = 16384;

const preparedTextByKey = new Map<string, PreparedTextWithSegments>();
export const layoutByKey = new Map<string, PretextVariableWidthLayout>();
const lineCountByKey = new Map<string, number>();
const fragmentOffsetAdvancesByFragment = new WeakMap<
  PretextLineFragment,
  number[]
>();
const graphemeOffsetsByText = new Map<string, number[]>();

type PretextWordBreak = "normal" | "keep-all";

export interface PretextLayoutItem {
  text: string;
  font: string;
  startOffset: number;
  endOffset: number;
  break?: "normal" | "never";
  wordBreak?: PretextWordBreak;
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
}

export interface PretextLineLayout {
  y: number;
  fragments: PretextLineFragment[];
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
}

export interface PretextSelectionRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

let measureCanvas: OffscreenCanvas | HTMLCanvasElement | undefined;
let measureCanvasContext:
  | OffscreenCanvasRenderingContext2D
  | CanvasRenderingContext2D
  | null
  | undefined;
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

  if (measureCanvasContext) {
    return measureCanvasContext ?? undefined;
  }

  if (typeof OffscreenCanvas !== "undefined") {
    measureCanvas = new OffscreenCanvas(1, 1);
    measureCanvasContext = measureCanvas.getContext("2d");
    return measureCanvasContext ?? undefined;
  }

  if (typeof document !== "undefined") {
    measureCanvas = document.createElement("canvas");
    measureCanvasContext = measureCanvas.getContext("2d");
    return measureCanvasContext ?? undefined;
  }

  return undefined;
}

function measureTextWidthPx(font: string, text: string): number {
  if (!text) {
    return 0;
  }

  const context = getMeasureContext();
  if (!context) {
    return 0;
  }

  context.font = font;
  return Math.max(0, Math.round(context.measureText(text).width));
}

function measureOffsetWidthPx(
  font: string,
  text: string,
  offset: number
): number {
  if (offset <= 0 || !text) {
    return 0;
  }

  return measureTextWidthPx(
    font,
    text.slice(0, Math.max(0, Math.min(offset, text.length)))
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

  const cached = graphemeOffsetsByText.get(text);
  if (cached) {
    return cached;
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
  graphemeOffsetsByText.set(text, offsets);
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
  exclusions: PretextExclusionRect[]
): string {
  const exclusionsKey = exclusions
    .map(
      (exclusion) =>
        `${exclusion.left},${exclusion.right},${exclusion.top},${exclusion.bottom}`
    )
    .join(";");
  return `${layoutSignature}\u0000${containerWidthPx}\u0000${lineHeightPx}\u0000${exclusionsKey}`;
}

export function cachedFragmentOffsetAdvances(
  defaultFont: string,
  fragment: PretextLineFragment
): number[] {
  const cached = fragmentOffsetAdvancesByFragment.get(fragment);
  if (cached) {
    return cached;
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
      localOffset
    );
  }
  fragmentOffsetAdvancesByFragment.set(fragment, advances);
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

  const cacheKey = `${font}\u0000${wordBreak}\u0000${text}`;
  const cached = getCachedValue(preparedTextByKey, cacheKey);
  if (cached) {
    return cached;
  }

  try {
    const prepared = prepareWithSegments(text, font, {
      whiteSpace: "pre-wrap",
      wordBreak,
    });
    preparedTextByKey.set(cacheKey, prepared);
    trimCache(preparedTextByKey, PREPARED_TEXT_CACHE_MAX_ENTRIES);
    return prepared;
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
  const cacheKey =
    `line-count\u0000${font}\u0000${wordBreak}` +
    `\u0000${safeWidth}\u0000${text}`;
  const cached = getCachedValue(lineCountByKey, cacheKey);
  if (cached !== undefined) {
    return cached;
  }

  const prepared = prepareCached(text, font, wordBreak);
  if (!prepared) {
    return undefined;
  }

  try {
    const lineCount = measureLineStats(prepared, safeWidth).lineCount;
    lineCountByKey.set(cacheKey, lineCount);
    trimCache(lineCountByKey, LINE_COUNT_CACHE_MAX_ENTRIES);
    return lineCount;
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
  const cacheKey = layoutCacheKey(
    `plain\u0000${font}\u0000${wordBreak}\u0000${text}`,
    safeContainerWidthPx,
    safeLineHeightPx,
    normalizedExclusions
  );
  const cachedLayout = getCachedValue(layoutByKey, cacheKey);
  if (cachedLayout) {
    return cachedLayout;
  }

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

    lines.push({
      y: rowTopPx,
      fragments,
    });
    rowTopPx += safeLineHeightPx;
  }

  const lineCount = lines.length;
  const contentBottomPx =
    lines.length > 0 ? (lines[lines.length - 1]?.y ?? 0) + safeLineHeightPx : 0;
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
  };
  layoutByKey.set(cacheKey, nextLayout);
  trimCache(layoutByKey, LAYOUT_CACHE_MAX_ENTRIES);
  return nextLayout;
}

