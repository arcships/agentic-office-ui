// Page layout modes for the DOCX surface.
//
// The continuous viewer stacks one page per row inside a scroll spacer. The
// same scroll space is reused by every mode so switching modes can keep the
// reader on the same page: `single` keeps one page per row (and the surface
// snaps the scroll offset to a page boundary), `spread` pairs pages into rows
// that are rendered side by side. Page offsets therefore stay scroll-space
// values that `scrollToPage`, bookmarks and the pagination scroll anchor can
// keep using unchanged.

/** How the DOCX surface lays its pages out. */
export type DocxViewMode = "continuous" | "single" | "spread"

/** Fit-to-container modes for the DOCX paper. */
export type DocxFitMode = "width" | "page"

export const DOCX_VIEW_MODES: readonly DocxViewMode[] = [
  "continuous",
  "single",
  "spread",
]

/** Anything that is not a known mode falls back to continuous scrolling. */
export function normalizeDocxViewMode(value: unknown): DocxViewMode {
  return value === "single" || value === "spread" ? value : "continuous"
}

/** Anything that is not a known fit mode means "use the zoom factor". */
export function normalizeDocxFitMode(value: unknown): DocxFitMode | null {
  return value === "width" || value === "page" ? value : null
}

/** True when the mode renders two pages on one row. */
export function isDocxSpreadMode(value: unknown): boolean {
  return normalizeDocxViewMode(value) === "spread"
}

/** Page indexes grouped into rows: one page per row, or pairs for spreads. */
export function docxPageRows(pageCount: number, mode: DocxViewMode): number[][] {
  const count = Number.isFinite(pageCount) ? Math.max(0, Math.trunc(pageCount)) : 0
  const perRow = normalizeDocxViewMode(mode) === "spread" ? 2 : 1
  const rows: number[][] = []
  for (let start = 0; start < count; start += perRow) {
    const indexes: number[] = []
    for (let pageIndex = start; pageIndex < Math.min(start + perRow, count); pageIndex++) {
      indexes.push(pageIndex)
    }
    rows.push(indexes)
  }
  return rows
}

export interface DocxPageDimension {
  widthPx: number
  heightPx: number
}

export interface DocxPageRowLayout {
  rowIndex: number
  pageIndexes: readonly number[]
  /** Scroll-space top of the row. */
  topPx: number
  /** Tallest page in the row. */
  heightPx: number
  /** Sum of the page footprints plus the spread gap. */
  widthPx: number
  /** Left edge of each page, relative to the row's left edge. */
  pageLeftsPx: readonly number[]
  /** Footprint (page + gutter) width of each page in the row. */
  pageWidthsPx: readonly number[]
  /** On-screen height of each page (the row is as tall as the tallest page). */
  pageHeightsPx: readonly number[]
}

export interface DocxViewLayout {
  mode: DocxViewMode
  rows: DocxPageRowLayout[]
  /** Scroll-space top per page (the row top for both pages of a spread). */
  pageOffsets: number[]
  /** On-screen height per page, in scroll px. */
  pageHeights: number[]
  /** On-screen width per page, in scroll px. */
  pageWidths: number[]
  /** Page + gutter footprint per page, used for centering and min width. */
  pageFootprints: number[]
  /** Left edge per page relative to its row. */
  pageLefts: number[]
  /** Row index per page. */
  pageRows: number[]
  totalHeightPx: number
  /** Widest row footprint; drives the horizontal scroll width. */
  maxRowWidthPx: number
}

export interface DocxViewLayoutOptions {
  pageCount: number
  mode?: DocxViewMode | null
  pageSize: (pageIndex: number) => DocxPageDimension
  /** Zoom factor applied to every page size (1 = 100%). */
  zoom?: number
  /** Vertical gap between rows. */
  gapPx?: number
  /** Top padding inside the scroll container. */
  topPaddingPx?: number
  /** Horizontal gap between the two pages of a spread. */
  spreadGapPx?: number
  /** Screen px reserved beside the paper (annotation pane). */
  gutterWidthPx?: number
  /** Printing always stacks one page per row so a spread never lands on one sheet. */
  print?: boolean
}

const DEFAULT_ROW_GAP_PX = 28
const DEFAULT_SPREAD_GAP_PX = 24
const DEFAULT_TOP_PADDING_PX = 24

function positive(value: number, fallback: number): number {
  return Number.isFinite(value) && value > 0 ? value : fallback
}

/**
 * Resolve the complete row/page geometry for a view mode.
 *
 * The result is pure: the surface only needs page sizes and the zoom factor,
 * so the same numbers drive rendering, hit testing and page navigation.
 */
export function resolveDocxViewLayout(options: DocxViewLayoutOptions): DocxViewLayout {
  const mode = options.print ? "continuous" : normalizeDocxViewMode(options.mode)
  const zoom = Number.isFinite(options.zoom) && (options.zoom ?? 0) > 0 ? options.zoom! : 1
  const gapPx = Math.max(0, Number.isFinite(options.gapPx) ? options.gapPx! : DEFAULT_ROW_GAP_PX)
  const topPaddingPx = Math.max(
    0,
    Number.isFinite(options.topPaddingPx) ? options.topPaddingPx! : DEFAULT_TOP_PADDING_PX,
  )
  const spreadGapPx = Math.max(
    0,
    Number.isFinite(options.spreadGapPx) ? options.spreadGapPx! : DEFAULT_SPREAD_GAP_PX,
  )
  const gutterWidthPx = Math.max(0, Number.isFinite(options.gutterWidthPx) ? options.gutterWidthPx! : 0)
  const pageCount = Number.isFinite(options.pageCount) ? Math.max(0, Math.trunc(options.pageCount)) : 0

  const pageWidths: number[] = []
  const pageHeights: number[] = []
  const pageFootprints: number[] = []
  for (let pageIndex = 0; pageIndex < pageCount; pageIndex++) {
    const size = options.pageSize(pageIndex)
    const width = positive(size?.widthPx ?? 0, 1) * zoom
    const height = positive(size?.heightPx ?? 0, 1) * zoom
    pageWidths.push(width)
    pageHeights.push(height)
    pageFootprints.push(width + gutterWidthPx)
  }

  const rows: DocxPageRowLayout[] = []
  const pageOffsets: number[] = []
  const pageLefts: number[] = []
  const pageRows: number[] = []
  const rowPageIndexes = docxPageRows(pageCount, mode)
  const columnCount = mode === "spread" ? 2 : 1
  // A column is as wide as the widest page that lands in it, so a narrow page
  // beside a wide one stays centered instead of hugging the column edge.
  const columnWidths = Array.from({ length: columnCount }, () => 0)
  for (const indexes of rowPageIndexes) {
    for (const [position, pageIndex] of indexes.entries()) {
      columnWidths[position] = Math.max(columnWidths[position] ?? 0, pageFootprints[pageIndex])
    }
  }
  const rowWidth = columnWidths.reduce((total, width) => total + width, 0) +
    spreadGapPx * Math.max(0, columnWidths.length - 1)
  let offset = topPaddingPx
  let maxRowWidthPx = 0
  for (const [rowIndex, pageIndexes] of rowPageIndexes.entries()) {
    const pageLeftsPx: number[] = []
    const rowPageWidths: number[] = []
    // A trailing lone page keeps the spread width and is centered on it.
    const lonePage = pageIndexes.length === 1 && columnCount > 1
    let height = 0
    for (const [position, pageIndex] of pageIndexes.entries()) {
      const columnLeft = columnWidths
        .slice(0, position)
        .reduce((total, width) => total + width + spreadGapPx, 0)
      const left = lonePage
        ? (rowWidth - pageFootprints[pageIndex]) / 2
        : columnLeft + (columnWidths[position] - pageFootprints[pageIndex]) / 2
      pageLeftsPx.push(left)
      rowPageWidths.push(pageFootprints[pageIndex])
      height = Math.max(height, pageHeights[pageIndex])
      pageOffsets[pageIndex] = offset
      pageLefts[pageIndex] = left
      pageRows[pageIndex] = rowIndex
    }
    rows.push({
      rowIndex,
      pageIndexes,
      topPx: offset,
      heightPx: height,
      widthPx: rowWidth,
      pageLeftsPx,
      pageWidthsPx: rowPageWidths,
      pageHeightsPx: pageIndexes.map((pageIndex) => pageHeights[pageIndex]),
    })
    maxRowWidthPx = Math.max(maxRowWidthPx, rowWidth)
    offset += height + gapPx
  }

  return {
    mode,
    rows,
    pageOffsets,
    pageHeights,
    pageWidths,
    pageFootprints,
    pageLefts,
    pageRows,
    totalHeightPx: rows.length === 0 ? 0 : offset,
    maxRowWidthPx,
  }
}

/** Row index containing a page, or -1 when the page is out of range. */
export function docxRowIndexForPage(layout: DocxViewLayout, pageIndex: number): number {
  const rowIndex = layout.pageRows[pageIndex]
  return rowIndex === undefined ? -1 : rowIndex
}

/**
 * Row nearest to a scroll offset, measured against the viewport centre so a
 * page is "current" once it owns the middle of the viewport.
 */
export function docxRowIndexAtOffset(
  rows: readonly DocxPageRowLayout[],
  scrollTop: number,
  viewportHeight = 0,
): number {
  if (rows.length === 0) return -1
  const anchor = Math.max(0, scrollTop) + Math.max(0, viewportHeight) / 2
  let best = 0
  let bestDistance = Number.POSITIVE_INFINITY
  for (const [rowIndex, row] of rows.entries()) {
    const top = row.topPx
    const bottom = row.topPx + row.heightPx
    if (anchor >= top && anchor <= bottom) return rowIndex
    const distance = anchor < top ? top - anchor : anchor - bottom
    if (distance < bestDistance) {
      bestDistance = distance
      best = rowIndex
    }
  }
  return best
}

/** Scroll offset of the row nearest to `scrollTop` (single-page snapping). */
export function docxSnapTopForRow(
  rows: readonly DocxPageRowLayout[],
  scrollTop: number,
  viewportHeight = 0,
): number | undefined {
  const rowIndex = docxRowIndexAtOffset(rows, scrollTop, viewportHeight)
  return rowIndex < 0 ? undefined : rows[rowIndex].topPx
}

/**
 * Page under a point in scroll space. `x` is relative to the scroll content's
 * left edge (scrollLeft + pointer offset) and `contentWidth` is the scroll
 * content width, which is what centers every row.
 */
export function docxPageAtPoint(
  rows: readonly DocxPageRowLayout[],
  point: { x: number; y: number; contentWidth: number },
): number | undefined {
  if (rows.length === 0) return undefined
  let row: DocxPageRowLayout | undefined
  let nearestRow = rows[0]
  let nearestRowDistance = Number.POSITIVE_INFINITY
  for (const candidate of rows) {
    if (point.y >= candidate.topPx && point.y < candidate.topPx + candidate.heightPx) {
      row = candidate
      break
    }
    const distance = point.y < candidate.topPx
      ? candidate.topPx - point.y
      : point.y - (candidate.topPx + candidate.heightPx)
    if (distance < nearestRowDistance) {
      nearestRowDistance = distance
      nearestRow = candidate
    }
  }
  const target = row ?? nearestRow
  const rowLeft = (Math.max(0, point.contentWidth) - target.widthPx) / 2
  let best = target.pageIndexes[0]
  let bestDistance = Number.POSITIVE_INFINITY
  for (const [position, pageIndex] of target.pageIndexes.entries()) {
    const left = rowLeft + target.pageLeftsPx[position]
    if (point.x >= left && point.x < left + target.pageWidthsPx[position]) return pageIndex
    // The pointer can sit in the spread gap or beside the paper; then the page
    // whose centre is closest wins.
    const distance = Math.abs(point.x - (left + target.pageWidthsPx[position] / 2))
    if (distance < bestDistance) {
      bestDistance = distance
      best = pageIndex
    }
  }
  return best
}

/** Rows that intersect a viewport window in scroll space. */
export function docxRowsInViewport(
  rows: readonly DocxPageRowLayout[],
  window: { startPx: number; endPx: number },
): DocxPageRowLayout[] {
  return rows.filter(
    (row) => row.topPx + row.heightPx >= window.startPx && row.topPx <= window.endPx,
  )
}

/** Advances a page index by whole spreads (used by spread page navigation). */
export function docxSpreadStepPage(
  pageCount: number,
  pageIndex: number,
  direction: 1 | -1,
): number {
  if (pageCount <= 0) return 0
  const clamped = Math.min(Math.max(0, pageIndex), pageCount - 1)
  const rowStart = clamped - (clamped % 2)
  const next = rowStart + direction * 2
  return Math.min(Math.max(0, next), Math.max(0, pageCount - 1))
}
