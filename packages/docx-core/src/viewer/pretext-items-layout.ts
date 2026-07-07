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
  layoutCacheKey,
  getCachedValue,
  trimCache,
  rowWidthsAtY,
  layoutByKey,
  type PretextLayoutItem,
  type PretextExclusionRect,
  type PretextLineFragment,
  type PretextLineLayout,
  type PretextVariableWidthLayout,
} from "./pretext-layout";

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

export function layoutItemsWithPretextAroundExclusions(
  text: string,
  items: PretextLayoutItem[],
  containerWidthPx: number,
  lineHeightPx: number,
  exclusions?: PretextExclusionRect[],
  fallbackFont?: string
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
        }\u0001${item.startOffset}\u0001${item.endOffset}\u0001${item.text}`
    )
    .join("\u0002");
  const cacheKey = layoutCacheKey(
    `items\u0000${layoutSignature}`,
    safeContainerWidthPx,
    safeLineHeightPx,
    normalizedExclusions
  );
  const cachedLayout = getCachedValue(layoutByKey, cacheKey);
  if (cachedLayout) {
    return cachedLayout;
  }

  const lines: PretextLineLayout[] = [];
  const consumedOffsetsByItemIndex = items.map(() => 0);
  let cursor: PretextItemCursor = {
    itemIndex: 0,
    segmentIndex: 0,
    graphemeIndex: 0,
  };
  let rowTopPx = 0;

  while (!itemCursorIsDone(preparedItems, cursor)) {
    cursor = normalizeItemCursor(preparedItems, cursor);
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
    font: fallbackFont ?? items[0]?.font,
    containerWidthPx: safeContainerWidthPx,
    lineHeightPx: safeLineHeightPx,
    exclusions: normalizedExclusions,
  };
  layoutByKey.set(cacheKey, nextLayout);
  trimCache(layoutByKey, LAYOUT_CACHE_MAX_ENTRIES);
  return nextLayout;
}