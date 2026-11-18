/**
 * Scroll anchoring for the measure→repaginate loop (#99).
 *
 * Repagination can change how many pages precede the viewport. The viewer
 * captures the page under the scroll position plus the offset inside it before
 * applying measurement feedback, then restores that same visual position once
 * the new page offsets exist.
 */
export interface PaginationScrollAnchor {
  /** 0-based page index containing the scroll position. */
  pageIndex: number;
  /** Scroll offset inside (or, in a page gap, relative to) that page. */
  offsetPx: number;
}

/**
 * Resolve the page under `scrollTop` from the page top offsets. Returns
 * `undefined` for an empty document; a scroll position above the first page
 * anchors to page 0 with a negative offset so it can be restored exactly.
 */
export function resolvePaginationScrollAnchor(
  pageOffsets: readonly number[],
  scrollTop: number
): PaginationScrollAnchor | undefined {
  if (pageOffsets.length === 0 || !Number.isFinite(scrollTop)) {
    return undefined;
  }

  const safeScrollTop = Math.max(0, scrollTop);
  let pageIndex = 0;
  for (let index = 1; index < pageOffsets.length; index += 1) {
    if ((pageOffsets[index] ?? 0) <= safeScrollTop) {
      pageIndex = index;
    } else {
      break;
    }
  }

  return {
    pageIndex,
    offsetPx: safeScrollTop - (pageOffsets[pageIndex] ?? 0),
  };
}

/**
 * Resolve the scroll position that puts the anchored page back where it was.
 * Returns `undefined` when the anchored page no longer exists.
 */
export function resolvePaginationScrollTop(
  pageOffsets: readonly number[],
  anchor: PaginationScrollAnchor
): number | undefined {
  if (
    pageOffsets.length === 0 ||
    !Number.isFinite(anchor.pageIndex) ||
    anchor.pageIndex < 0 ||
    anchor.pageIndex >= pageOffsets.length ||
    !Number.isFinite(anchor.offsetPx)
  ) {
    return undefined;
  }
  const pageTop = pageOffsets[anchor.pageIndex] ?? 0;
  return Math.max(0, pageTop + anchor.offsetPx);
}
