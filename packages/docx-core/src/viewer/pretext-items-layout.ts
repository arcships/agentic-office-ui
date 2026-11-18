import {
  layoutNextLine,
  type LayoutCursor,
  type LayoutLine,
  type PreparedTextWithSegments,
} from "@chenglou/pretext";
import {
  canUsePretext,
  prepareCached,
  cursorAdvanceCodeUnits,
  aggregateFragmentLineMetrics,
  layoutCacheKey,
  getCachedValue,
  lineBoxHeightPx,
  resolvePretextLineBoxMetrics,
  trimCache,
  rowWidthsAtY,
  layoutByKey,
  type PretextLayoutItem,
  type PretextExclusionRect,
  type PretextLineFragment,
  type PretextLineLayout,
  type PretextVariableWidthLayout,
} from "./pretext-layout";
import { emptyLineMetrics } from "./line-metrics";

const LAYOUT_CACHE_MAX_ENTRIES = 4096;

export interface PretextItemCursor {
  itemIndex: number;
  segmentIndex: number;
  graphemeIndex: number;
}

interface InternalPretextItemLineFragment {
  itemIndex: number;
  text: string;
  width: number;
  font: string;
  start: LayoutCursor;
  end: LayoutCursor;
}

interface InternalPretextItemLine {
  end: PretextItemCursor;
  fragments: InternalPretextItemLineFragment[];
}

function cloneItemCursor(cursor: PretextItemCursor): PretextItemCursor {
  return {
    itemIndex: cursor.itemIndex,
    segmentIndex: cursor.segmentIndex,
    graphemeIndex: cursor.graphemeIndex,
  };
}

function itemCursorAtStart(cursor: PretextItemCursor): boolean {
  return cursor.segmentIndex === 0 && cursor.graphemeIndex === 0;
}

function normalizeItemCursor(
  preparedItems: Array<PreparedTextWithSegments | undefined>,
  cursor: PretextItemCursor
): PretextItemCursor {
  const nextCursor = cloneItemCursor(cursor);
  while (nextCursor.itemIndex < preparedItems.length) {
    const prepared = preparedItems[nextCursor.itemIndex];
    if (!prepared || nextCursor.segmentIndex >= prepared.segments.length) {
      nextCursor.itemIndex += 1;
      nextCursor.segmentIndex = 0;
      nextCursor.graphemeIndex = 0;
      continue;
    }
    break;
  }
  return nextCursor;
}

function itemCursorIsDone(
  preparedItems: Array<PreparedTextWithSegments | undefined>,
  cursor: PretextItemCursor
): boolean {
  return (
    normalizeItemCursor(preparedItems, cursor).itemIndex >= preparedItems.length
  );
}

function wholeRemainingItemLine(
  prepared: PreparedTextWithSegments,
  cursor: LayoutCursor
): LayoutLine | null {
  return layoutNextLine(prepared, cursor, Number.POSITIVE_INFINITY);
}

function layoutNextItemLine(
  items: PretextLayoutItem[],
  preparedItems: Array<PreparedTextWithSegments | undefined>,
  start: PretextItemCursor,
  maxWidth: number
): InternalPretextItemLine | null {
  const cursor = normalizeItemCursor(preparedItems, start);
  if (cursor.itemIndex >= items.length) {
    return null;
  }

  const safeMaxWidth = Math.max(1, maxWidth);
  const fragments: InternalPretextItemLineFragment[] = [];
  let remainingWidth = safeMaxWidth;
  let current = cloneItemCursor(cursor);

  while (current.itemIndex < items.length) {
    const item = items[current.itemIndex];
    const prepared = preparedItems[current.itemIndex];
    if (!item || !prepared) {
      current.itemIndex += 1;
      current.segmentIndex = 0;
      current.graphemeIndex = 0;
      continue;
    }

    const itemCursor: LayoutCursor = {
      segmentIndex: current.segmentIndex,
      graphemeIndex: current.graphemeIndex,
    };
    const atItemStart = itemCursorAtStart(current);
    const remainingItemLine =
      item.break === "never"
        ? wholeRemainingItemLine(prepared, itemCursor)
        : layoutNextLine(prepared, itemCursor, Math.max(1, remainingWidth));
    if (!remainingItemLine) {
      current.itemIndex += 1;
      current.segmentIndex = 0;
      current.graphemeIndex = 0;
      continue;
    }

    const noProgress =
      remainingItemLine.end.segmentIndex === itemCursor.segmentIndex &&
      remainingItemLine.end.graphemeIndex === itemCursor.graphemeIndex &&
      remainingItemLine.text.length === 0;
    if (noProgress) {
      current.itemIndex += 1;
      current.segmentIndex = 0;
      current.graphemeIndex = 0;
      continue;
    }

    const overflowsCurrentLine =
      fragments.length > 0 &&
      atItemStart &&
      remainingItemLine.width > remainingWidth + 0.5;
    if (overflowsCurrentLine) {
      break;
    }

    fragments.push({
      itemIndex: current.itemIndex,
      text: remainingItemLine.text,
      width: remainingItemLine.width,
      font: item.font,
      start: itemCursor,
      end: remainingItemLine.end,
    });

    remainingWidth = Math.max(0, remainingWidth - remainingItemLine.width);

    if (remainingItemLine.end.segmentIndex >= prepared.segments.length) {
      current.itemIndex += 1;
      current.segmentIndex = 0;
      current.graphemeIndex = 0;
      if (remainingWidth <= 0.5) {
        break;
      }
      continue;
    }

    current.segmentIndex = remainingItemLine.end.segmentIndex;
    current.graphemeIndex = remainingItemLine.end.graphemeIndex;
    break;
  }

  if (fragments.length === 0) {
    return null;
  }

  return {
    end: normalizeItemCursor(preparedItems, current),
    fragments,
  };
}

function lineSplitsLeadingItem(
  items: PretextLayoutItem[],
  preparedItems: Array<PreparedTextWithSegments | undefined>,
  start: PretextItemCursor,
  line: InternalPretextItemLine
): boolean {
  if (!itemCursorAtStart(start)) {
    return false;
  }

  const firstFragment = line.fragments[0];
  const item = firstFragment ? items[firstFragment.itemIndex] : undefined;
  const prepared = firstFragment
    ? preparedItems[firstFragment.itemIndex]
    : undefined;
  if (
    !firstFragment ||
    !item ||
    !prepared ||
    item.break === "never" ||
    firstFragment.itemIndex !== start.itemIndex
  ) {
    return false;
  }

  const wholeLine = wholeRemainingItemLine(prepared, {
    segmentIndex: start.segmentIndex,
    graphemeIndex: start.graphemeIndex,
  });
  if (!wholeLine) {
    return false;
  }

  return firstFragment.text.length < wholeLine.text.length;
}

function laterIntervalFitsLeadingItemWithoutSplit(
  items: PretextLayoutItem[],
  preparedItems: Array<PreparedTextWithSegments | undefined>,
  start: PretextItemCursor,
  line: InternalPretextItemLine,
  laterIntervals: Array<{
    x: number;
    width: number;
  }>
): boolean {
  const item = items[start.itemIndex];
  const prepared = preparedItems[start.itemIndex];
  if (
    !item ||
    !prepared ||
    !itemCursorAtStart(start) ||
    item.break === "never"
  ) {
    return false;
  }

  const startCursor: LayoutCursor = {
    segmentIndex: start.segmentIndex,
    graphemeIndex: start.graphemeIndex,
  };
  const wholeLine = wholeRemainingItemLine(prepared, startCursor);
  if (!wholeLine) {
    return false;
  }

  if (
    laterIntervals.some((interval) => wholeLine.width <= interval.width + 0.5)
  ) {
    return true;
  }
  return false;
}

/**
 * Vertical metrics an item contributes to its fragment: explicit ascent/
 * descent for inline objects and runs in another font size, the baseline shift
 * of a sub/superscript run, and nothing for plain text (its font is measured
 * when the line box is aggregated).
 */
function itemVerticalMetrics(item: PretextLayoutItem): {
  ascentPx?: number;
  descentPx?: number;
  unsuppressibleDescentPx?: number;
  hasCustomVerticalMetrics?: boolean;
  baselineShiftPx?: number;
} {
  const shiftMetrics =
    item.baselineShiftPx !== undefined && item.baselineShiftPx !== 0
      ? { baselineShiftPx: item.baselineShiftPx }
      : {};
  if (
    item.ascentPx === undefined &&
    item.descentPx === undefined &&
    item.unsuppressibleDescentPx === undefined
  ) {
    return {
      ...shiftMetrics,
      ...(item.hasCustomVerticalMetrics === true
        ? { hasCustomVerticalMetrics: true }
        : {}),
    };
  }

  return {
    ascentPx: item.ascentPx ?? 0,
    descentPx: item.descentPx ?? 0,
    unsuppressibleDescentPx: item.unsuppressibleDescentPx ?? item.descentPx ?? 0,
    ...(item.hasCustomVerticalMetrics === true
      ? { hasCustomVerticalMetrics: true }
      : {}),
    ...shiftMetrics,
  };
}

export function layoutItemsWithPretextAroundExclusions(
  text: string,
  items: PretextLayoutItem[],
  containerWidthPx: number,
  lineHeightPx: number,
  exclusions?: PretextExclusionRect[],
  fallbackFont?: string,
  options?: {
    /** Snap grown line boxes up to the document grid line pitch. */
    docGridLinePitchPx?: number;
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
      font: fallbackFont,
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

  if (!canUsePretext()) {
    return undefined;
  }

  const preparedItems = items.map((item) =>
    prepareCached(item.text, item.font, item.wordBreak ?? "normal")
  );
  if (
    preparedItems.some(
      (prepared, index) => !prepared && items[index]?.text.length
    )
  ) {
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
  const layoutSignature = items
    .map(
      (item) =>
        `${item.font}\u0001${item.break ?? "normal"}\u0001${
          item.wordBreak ?? "normal"
        }\u0001${item.startOffset}\u0001${item.endOffset}\u0001${
          item.ascentPx ?? ""
        }\u0001${item.descentPx ?? ""}\u0001${
          item.unsuppressibleDescentPx ?? ""
        }\u0001${item.hasCustomVerticalMetrics === true ? 1 : 0}\u0001${
          item.baselineShiftPx ?? ""
        }\u0001${item.text}`
    )
    .join("\u0002");
  const cacheKey = layoutCacheKey(
    `items\u0000${layoutSignature}`,
    safeContainerWidthPx,
    safeLineHeightPx,
    normalizedExclusions,
    options?.docGridLinePitchPx
  );
  const cachedLayout = getCachedValue(layoutByKey, cacheKey);
  if (cachedLayout) {
    return cachedLayout;
  }

  // A line may only grow past the paragraph line height when the content on
  // that line cannot be described by a font-only box (an inline object or a run
  // in another font size). Lines without such content keep the paragraph line
  // height, so existing pagination stays unchanged and a single inline image
  // does not inflate every other line of its paragraph.
  const lines: PretextLineLayout[] = [];
  const consumedOffsetsByItemIndex = items.map(() => 0);
  let cursor: PretextItemCursor = {
    itemIndex: 0,
    segmentIndex: 0,
    graphemeIndex: 0,
  };
  let rowTopPx = 0;
  let grewAnyLine = false;
  // Band height used to test exclusions. It tracks the previous line box so a
  // tall inline image keeps wrapped text clear of its band.
  let rowHeightPx = safeLineHeightPx;

  while (!itemCursorIsDone(preparedItems, cursor)) {
    cursor = normalizeItemCursor(preparedItems, cursor);
    const rowIntervals = rowWidthsAtY(
      safeContainerWidthPx,
      rowHeightPx,
      rowTopPx,
      normalizedExclusions
    );
    const fragments: PretextLineFragment[] = [];

    if (rowIntervals.length === 0) {
      rowTopPx += rowHeightPx;
      continue;
    }

    for (
      let intervalIndex = 0;
      intervalIndex < rowIntervals.length;
      intervalIndex += 1
    ) {
      cursor = normalizeItemCursor(preparedItems, cursor);
      if (itemCursorIsDone(preparedItems, cursor)) {
        break;
      }

      const interval = rowIntervals[intervalIndex]!;
      const line = layoutNextItemLine(
        items,
        preparedItems,
        cursor,
        interval.width
      );
      if (!line) {
        continue;
      }

      if (
        lineSplitsLeadingItem(items, preparedItems, cursor, line) &&
        laterIntervalFitsLeadingItemWithoutSplit(
          items,
          preparedItems,
          cursor,
          line,
          rowIntervals.slice(intervalIndex + 1)
        )
      ) {
        continue;
      }

      let nextFragmentX = interval.x;
      let nextCursor = cloneItemCursor(cursor);
      for (const lineFragment of line.fragments) {
        const item = items[lineFragment.itemIndex];
        const prepared = preparedItems[lineFragment.itemIndex];
        if (!item || !prepared) {
          continue;
        }

        const consumedCodeUnits = cursorAdvanceCodeUnits(
          prepared,
          lineFragment.start,
          lineFragment.end
        );
        const startOffset =
          item.startOffset +
          (consumedOffsetsByItemIndex[lineFragment.itemIndex] ?? 0);
        const endOffset = startOffset + consumedCodeUnits;

        fragments.push({
          text: lineFragment.text,
          width: lineFragment.width,
          x: nextFragmentX,
          intervalX: interval.x,
          intervalWidth: interval.width,
          startOffset,
          endOffset,
          font: lineFragment.font,
          ...itemVerticalMetrics(item),
        });

        consumedOffsetsByItemIndex[lineFragment.itemIndex] =
          (consumedOffsetsByItemIndex[lineFragment.itemIndex] ?? 0) +
          consumedCodeUnits;
        nextFragmentX += lineFragment.width;

        if (lineFragment.itemIndex === nextCursor.itemIndex) {
          if (lineFragment.end.segmentIndex >= prepared.segments.length) {
            nextCursor.itemIndex += 1;
            nextCursor.segmentIndex = 0;
            nextCursor.graphemeIndex = 0;
          } else {
            nextCursor.segmentIndex = lineFragment.end.segmentIndex;
            nextCursor.graphemeIndex = lineFragment.end.graphemeIndex;
          }
        }
      }

      cursor = line.end;
    }

    if (fragments.length === 0) {
      break;
    }

    const lineHasGrowthContent = fragments.some(
      (fragment) => fragment.hasCustomVerticalMetrics === true
    );
    const lineBoxMetrics = resolvePretextLineBoxMetrics(
      aggregateFragmentLineMetrics(fragments, fallbackFont),
      {
        minimumHeightPx: safeLineHeightPx,
        docGridLinePitchPx: options?.docGridLinePitchPx,
        allowGrowth: lineHasGrowthContent
      }
    );
    if (lineBoxMetrics.heightPx > safeLineHeightPx) {
      grewAnyLine = true;
    }
    lines.push({
      y: rowTopPx,
      fragments,
      metrics: lineBoxMetrics,
      baselineYPx: rowTopPx + lineBoxMetrics.ascentPx,
    });
    rowHeightPx = lineBoxMetrics.heightPx;
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
    font: fallbackFont ?? items[0]?.font,
    containerWidthPx: safeContainerWidthPx,
    lineHeightPx: safeLineHeightPx,
    exclusions: normalizedExclusions,
    lineMetrics: lines.map((line) => line.metrics ?? emptyLineMetrics()),
    hasCustomVerticalMetrics: grewAnyLine,
  };
  layoutByKey.set(cacheKey, nextLayout);
  trimCache(layoutByKey, LAYOUT_CACHE_MAX_ENTRIES);
  return nextLayout;
}