/**
 * Pure vertical layout math for `PdfSurface`'s lazily rasterized page stack.
 *
 * The surface stacks page slots with an explicit width/height per slot plus a
 * fixed flex `gap` and container `padding`, so page offsets inside the scroll
 * container are fully determined by the page display sizes. Deriving them here
 * keeps scrolling and hit-testing free of per-frame DOM queries: the caller
 * caches the offsets per document / zoom / rotation and only recomputes them
 * when that geometry changes.
 *
 * Every function is dependency free so it can be exercised without a browser.
 */

/** Inclusive zero-based index range of pages kept rasterized. */
export interface PdfPageRange {
  /** First rendered page index; `-1` when nothing is rendered. */
  first: number
  /** Last rendered page index; `-1` when nothing is rendered. */
  last: number
}

/** Pages rasterized above and below the viewport by default. */
export const DEFAULT_PAGE_OVERSCAN = 1

export interface PdfPageOffsetOptions {
  /** Vertical gap between stacked pages in CSS px. */
  gap?: number
  /** Padding above the first page in CSS px. */
  paddingTop?: number
}

export interface PdfPageVisibilityOptions {
  /** Current scroll offset of the page container in CSS px. */
  scrollTop: number
  /** Visible height of the page container in CSS px. `0` means not measured yet. */
  viewportHeight: number
  /** Extra pages kept rasterized above and below the viewport. */
  overscanPages?: number
}

function toNonNegativeNumber(value: unknown): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? value : 0
}

function pageCountOf(offsets: readonly number[], heights: readonly number[]): number {
  return Math.min(offsets.length, heights.length)
}

function pageBottom(offsets: readonly number[], heights: readonly number[], pageIndex: number): number {
  return offsets[pageIndex] + heights[pageIndex]
}

/**
 * Offsets of every page top inside the scroll container, in CSS px.
 *
 * `heights` must be ordered by page index. Non-finite or negative heights are
 * treated as `0` so a broken measurement degrades to a degenerate stack instead
 * of `NaN` offsets.
 */
export function computePageOffsets(
  heights: readonly number[],
  options: PdfPageOffsetOptions = {},
): number[] {
  const gap = toNonNegativeNumber(options.gap)
  const paddingTop = toNonNegativeNumber(options.paddingTop)
  const offsets: number[] = []
  let top = paddingTop
  for (let pageIndex = 0; pageIndex < heights.length; pageIndex++) {
    offsets.push(top)
    top += toNonNegativeNumber(heights[pageIndex]) + gap
  }
  return offsets
}

/** First page whose visible bottom reaches `position`, or the page count. */
function firstPageReaching(
  offsets: readonly number[],
  heights: readonly number[],
  position: number,
): number {
  const count = pageCountOf(offsets, heights)
  let low = 0
  let high = count - 1
  let found = count
  while (low <= high) {
    const middle = (low + high) >> 1
    if (pageBottom(offsets, heights, middle) >= position) {
      found = middle
      high = middle - 1
    } else {
      low = middle + 1
    }
  }
  return found
}

/** Last page whose top is at or above `position`, or `-1` when none is. */
function lastPageStartingBefore(
  offsets: readonly number[],
  heights: readonly number[],
  position: number,
): number {
  const count = pageCountOf(offsets, heights)
  let low = 0
  let high = count - 1
  let found = -1
  while (low <= high) {
    const middle = (low + high) >> 1
    if (offsets[middle] <= position) {
      found = middle
      low = middle + 1
    } else {
      high = middle - 1
    }
  }
  return found
}

/**
 * Page containing `position` inside the scroll content.
 *
 * Positions above the first page resolve to page `0` and positions past the
 * last page resolve to the last page, matching how `PdfSurface` picks the page
 * under the viewport centre and under a context-menu click.
 */
export function pageIndexAtPosition(
  offsets: readonly number[],
  heights: readonly number[],
  position: number,
): number {
  const count = pageCountOf(offsets, heights)
  if (count <= 0) return 0
  const candidate = lastPageStartingBefore(offsets, heights, Number.isFinite(position) ? position : 0)
  return Math.max(0, Math.min(count - 1, candidate))
}

/**
 * Pages that must be rasterized for the given scroll position.
 *
 * When `viewportHeight` is not measured yet (SSR, headless host, first paint)
 * the result is a bounded window from the top of the document: rasterizing
 * every page is exactly the behaviour this module exists to avoid.
 */
export function visiblePageRange(
  offsets: readonly number[],
  heights: readonly number[],
  options: PdfPageVisibilityOptions,
): PdfPageRange {
  const count = pageCountOf(offsets, heights)
  if (count <= 0) return { first: 0, last: -1 }
  const overscan = Math.max(0, Math.trunc(toNonNegativeNumber(options.overscanPages ?? DEFAULT_PAGE_OVERSCAN)))
  const viewportHeight = toNonNegativeNumber(options.viewportHeight)
  if (viewportHeight === 0) {
    return { first: 0, last: Math.min(count - 1, overscan * 2) }
  }
  const scrollTop = Math.max(0, Number.isFinite(options.scrollTop) ? options.scrollTop : 0)
  const viewportEnd = scrollTop + viewportHeight
  const firstRendered = Math.max(0, firstPageReaching(offsets, heights, scrollTop) - overscan)
  const lastRendered = Math.min(
    count - 1,
    lastPageStartingBefore(offsets, heights, viewportEnd) + overscan,
  )
  if (firstRendered > lastRendered) return { first: 0, last: -1 }
  return { first: firstRendered, last: lastRendered }
}

/** Whether `pageIndex` falls inside an inclusive page range. */
export function pageRangeContains(range: PdfPageRange, pageIndex: number): boolean {
  return range.last >= range.first && pageIndex >= range.first && pageIndex <= range.last
}

/** Structural comparison of two page ranges. */
export function pageRangeEquals(left: PdfPageRange, right: PdfPageRange): boolean {
  return left.first === right.first && left.last === right.last
}

/** Rasterization state of one page slot. */
export type PdfPageBitmapState = "pending" | "loading" | "ready" | "error"

export interface PdfPageSlotSnapshot {
  state: PdfPageBitmapState
  /** Whether the slot still holds an object URL for a rasterized bitmap. */
  hasBitmap: boolean
}

export interface PdfPageRenderPlan {
  /** Page indexes to rasterize, ascending. */
  render: number[]
  /** Page indexes whose bitmap must be revoked, ascending. */
  release: number[]
}

/**
 * Pages to rasterize and pages to release for a viewport window.
 *
 * In-range pages are rendered unless they are already `ready`, so a failed
 * page retries the next time it enters the window. In-range `loading` slots
 * are rendered again: callers plan after aborting the previous pass, so a
 * `loading` slot that is still visible belongs to work that will never finish.
 * Out-of-range slots must release their bitmap and return to `pending`, which
 * keeps the bitmap budget proportional to the viewport rather than to the
 * document.
 */
export function planPageRenders(
  slots: readonly PdfPageSlotSnapshot[],
  range: PdfPageRange,
): PdfPageRenderPlan {
  const plan: PdfPageRenderPlan = { render: [], release: [] }
  for (let pageIndex = 0; pageIndex < slots.length; pageIndex++) {
    const slot = slots[pageIndex]
    if (pageRangeContains(range, pageIndex)) {
      if (!slot || slot.state !== "ready") plan.render.push(pageIndex)
      continue
    }
    if ((slot && slot.state !== "pending") || slot?.hasBitmap === true) plan.release.push(pageIndex)
  }
  return plan
}
