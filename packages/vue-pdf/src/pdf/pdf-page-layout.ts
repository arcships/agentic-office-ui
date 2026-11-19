/**
 * Page layout math for the PDF viewer page modes.
 *
 * `continuous` stacks one page per row, `single` keeps one page per row and the
 * surface snaps the scroll offset to a page boundary, and `spread` places two
 * facing pages on one row. Page offsets stay scroll-space values so page
 * navigation, hit testing and the rasterization window keep working for every
 * mode. Rows are laid out as a CSS grid with one or two `max-content` columns,
 * so the width of a column is the widest page that lands in it; the helpers
 * below resolve the same geometry the grid paints so hit testing cannot drift
 * from the DOM.
 *
 * Every function is dependency free so it can be exercised without a browser.
 */

/** How the PDF surface lays its pages out. */
export type PdfPageLayout = "continuous" | "single" | "spread"

export const PDF_PAGE_LAYOUTS: readonly PdfPageLayout[] = [
  "continuous",
  "single",
  "spread",
]

/** Anything that is not a known mode falls back to continuous scrolling. */
export function normalizePdfPageLayout(value: unknown): PdfPageLayout {
  return value === "single" || value === "spread" ? value : "continuous"
}

/** True when the mode renders two facing pages on one row. */
export function isPdfSpreadLayout(value: unknown): boolean {
  return normalizePdfPageLayout(value) === "spread"
}

/** Page indexes grouped into rows: one page per row, or facing pairs. */
export function pdfPageLayoutRows(pageCount: number, layout: PdfPageLayout): number[][] {
  const count = Number.isFinite(pageCount) ? Math.max(0, Math.trunc(pageCount)) : 0
  const perRow = normalizePdfPageLayout(layout) === "spread" ? 2 : 1
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

/** The pages a spread shows around `pageIndex`, ascending. */
export function pdfSpreadPagePair(pageCount: number, pageIndex: number): number[] {
  const count = Number.isFinite(pageCount) ? Math.max(0, Math.trunc(pageCount)) : 0
  if (count <= 0) return []
  const clamped = Math.min(Math.max(0, Math.trunc(pageIndex)), count - 1)
  const start = clamped - (clamped % 2)
  return start + 1 < count ? [start, start + 1] : [start]
}

/** Advances a page index by whole spreads. */
export function pdfSpreadStepPage(
  pageCount: number,
  pageIndex: number,
  direction: 1 | -1,
): number {
  const count = Number.isFinite(pageCount) ? Math.max(0, Math.trunc(pageCount)) : 0
  if (count <= 0) return 0
  const clamped = Math.min(Math.max(0, Math.trunc(pageIndex)), count - 1)
  const start = clamped - (clamped % 2)
  const next = start + direction * 2
  return Math.min(Math.max(0, next), count - 1)
}

export interface PdfPageDisplaySize {
  width: number
  height: number
}

export interface PdfPageLayoutRow {
  rowIndex: number
  pageIndexes: readonly number[]
  /** Scroll-space top of the row. */
  topPx: number
  /** Tallest page in the row. */
  heightPx: number
  /** Width of the whole row, including the spread gap. */
  widthPx: number
  /** Left edge of each page, relative to the row's left edge. */
  pageLeftsPx: readonly number[]
  /** Width of each page in the row. */
  pageWidthsPx: readonly number[]
  /** Height of each page in the row. */
  pageHeightsPx: readonly number[]
}

export interface PdfPageLayoutMetrics {
  layout: PdfPageLayout
  rows: PdfPageLayoutRow[]
  /** Scroll-space top per page (both pages of a spread share the row top). */
  pageOffsets: number[]
  pageHeights: number[]
  pageWidths: number[]
  /** Left edge per page relative to its row. */
  pageLefts: number[]
  /** Grid row per page (1-based, matches `grid-row`). */
  pageGridRows: number[]
  /** Grid column per page (1-based) or `0` when the page spans every column. */
  pageGridColumns: number[]
  /** Width of every grid column, in px. */
  columnWidths: number[]
  rowGapPx: number
  columnGapPx: number
  paddingTopPx: number
  totalHeightPx: number
  /** Widest row; drives the horizontal scroll width. */
  maxRowWidthPx: number
}

export interface PdfPageLayoutOptions {
  pageCount: number
  layout?: PdfPageLayout | null
  pageSize: (pageIndex: number) => PdfPageDisplaySize
  /** Vertical gap between rows. */
  rowGap?: number
  /** Horizontal gap between the two pages of a spread. */
  columnGap?: number
  /** Padding above the first row. */
  paddingTop?: number
  /** Printing keeps one page per row so a spread never lands on one sheet. */
  print?: boolean
}

const DEFAULT_ROW_GAP_PX = 8
const DEFAULT_COLUMN_GAP_PX = 8
const DEFAULT_PADDING_TOP_PX = 16

function nonNegative(value: number | undefined, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value >= 0 ? value : fallback
}

function positive(value: number, fallback: number): number {
  return Number.isFinite(value) && value > 0 ? value : fallback
}

function spreadColumns(layout: PdfPageLayout): number {
  return layout === "spread" ? 2 : 1
}

/**
 * Resolve every row and page rectangle for a page layout.
 *
 * Column widths follow CSS grid `max-content` sizing: a column is as wide as
 * its widest page, and narrower pages are centred inside it. A spread row with
 * a single page (odd page count) spans both columns so it stays centered.
 */
export function resolvePdfPageLayoutMetrics(options: PdfPageLayoutOptions): PdfPageLayoutMetrics {
  const layout = options.print ? "continuous" : normalizePdfPageLayout(options.layout)
  const rowGapPx = nonNegative(options.rowGap, DEFAULT_ROW_GAP_PX)
  const columnGapPx = nonNegative(options.columnGap, DEFAULT_COLUMN_GAP_PX)
  const paddingTopPx = nonNegative(options.paddingTop, DEFAULT_PADDING_TOP_PX)
  const pageCount = Number.isFinite(options.pageCount) ? Math.max(0, Math.trunc(options.pageCount)) : 0
  const columnCount = spreadColumns(layout)

  const pageWidths: number[] = []
  const pageHeights: number[] = []
  for (let pageIndex = 0; pageIndex < pageCount; pageIndex++) {
    const size = options.pageSize(pageIndex)
    pageWidths.push(positive(size?.width ?? 0, 1))
    pageHeights.push(positive(size?.height ?? 0, 1))
  }

  const pageRows = pdfPageLayoutRows(pageCount, layout)
  // A column is as wide as the widest page that lands in it.
  const columnWidths = Array.from({ length: columnCount }, () => 0)
  for (const indexes of pageRows) {
    for (const [position, pageIndex] of indexes.entries()) {
      const column = indexes.length === 1 && columnCount > 1 ? 0 : position
      columnWidths[column] = Math.max(columnWidths[column] ?? 0, pageWidths[pageIndex])
    }
  }
  const columnsTotal = columnWidths.reduce((total, width) => total + width, 0) +
    columnGapPx * Math.max(0, columnWidths.length - 1)

  const rows: PdfPageLayoutRow[] = []
  const pageOffsets: number[] = []
  const pageLefts: number[] = []
  const pageGridRows: number[] = []
  const pageGridColumns: number[] = []
  let offset = paddingTopPx
  let maxRowWidthPx = 0
  for (const [rowIndex, indexes] of pageRows.entries()) {
    const pageLeftsPx: number[] = []
    const rowWidth = columnsTotal
    let height = 0
    const spansRow = indexes.length === 1 && columnCount > 1
    for (const [position, pageIndex] of indexes.entries()) {
      const column = spansRow ? 0 : position
      const columnLeft = columnWidths
        .slice(0, column)
        .reduce((total, width) => total + width + columnGapPx, 0)
      const pageLeft = spansRow
        ? (rowWidth - pageWidths[pageIndex]) / 2
        : columnLeft + (columnWidths[column] - pageWidths[pageIndex]) / 2
      pageLeftsPx.push(pageLeft)
      height = Math.max(height, pageHeights[pageIndex])
      pageOffsets[pageIndex] = offset
      pageLefts[pageIndex] = pageLeft
      pageGridRows[pageIndex] = rowIndex + 1
      pageGridColumns[pageIndex] = spansRow ? 0 : column + 1
    }
    rows.push({
      rowIndex,
      pageIndexes: indexes,
      topPx: offset,
      heightPx: height,
      widthPx: rowWidth,
      pageLeftsPx,
      pageWidthsPx: indexes.map((pageIndex) => pageWidths[pageIndex]),
      pageHeightsPx: indexes.map((pageIndex) => pageHeights[pageIndex]),
    })
    maxRowWidthPx = Math.max(maxRowWidthPx, rowWidth)
    offset += height + rowGapPx
  }

  return {
    layout,
    rows,
    pageOffsets,
    pageHeights,
    pageWidths,
    pageLefts,
    pageGridRows,
    pageGridColumns,
    columnWidths,
    rowGapPx,
    columnGapPx,
    paddingTopPx,
    totalHeightPx: rows.length === 0 ? 0 : offset,
    maxRowWidthPx,
  }
}

/** Row nearest to a scroll offset, measured against the viewport centre. */
export function pdfRowIndexAtOffset(
  rows: readonly PdfPageLayoutRow[],
  scrollTop: number,
  viewportHeight = 0,
): number {
  if (rows.length === 0) return -1
  const anchor = Math.max(0, scrollTop) + Math.max(0, viewportHeight) / 2
  let best = 0
  let bestDistance = Number.POSITIVE_INFINITY
  for (const [rowIndex, row] of rows.entries()) {
    if (anchor >= row.topPx && anchor <= row.topPx + row.heightPx) return rowIndex
    const distance = anchor < row.topPx
      ? row.topPx - anchor
      : anchor - (row.topPx + row.heightPx)
    if (distance < bestDistance) {
      bestDistance = distance
      best = rowIndex
    }
  }
  return best
}

/** First page of the row nearest to `scrollTop` (the spread's left page). */
export function pdfPageIndexAtOffset(
  rows: readonly PdfPageLayoutRow[],
  scrollTop: number,
  viewportHeight = 0,
): number {
  const rowIndex = pdfRowIndexAtOffset(rows, scrollTop, viewportHeight)
  return rowIndex < 0 ? 0 : rows[rowIndex].pageIndexes[0]
}

/** Scroll offset of the row nearest to `scrollTop` (single-page snapping). */
export function pdfSnapTopForRow(
  rows: readonly PdfPageLayoutRow[],
  scrollTop: number,
  viewportHeight = 0,
): number | undefined {
  const rowIndex = pdfRowIndexAtOffset(rows, scrollTop, viewportHeight)
  return rowIndex < 0 ? undefined : rows[rowIndex].topPx
}

/**
 * Page under a point in scroll space. `x` is relative to the scroll content's
 * left edge and `contentWidth` is the scroll content width, which is what
 * centers every row.
 */
export function pdfPageAtPoint(
  rows: readonly PdfPageLayoutRow[],
  point: { x: number; y: number; contentWidth: number },
): number | undefined {
  if (rows.length === 0) return undefined
  let row: PdfPageLayoutRow | undefined
  let nearest = rows[0]
  let nearestDistance = Number.POSITIVE_INFINITY
  for (const candidate of rows) {
    if (point.y >= candidate.topPx && point.y < candidate.topPx + candidate.heightPx) {
      row = candidate
      break
    }
    const distance = point.y < candidate.topPx
      ? candidate.topPx - point.y
      : point.y - (candidate.topPx + candidate.heightPx)
    if (distance < nearestDistance) {
      nearestDistance = distance
      nearest = candidate
    }
  }
  const target = row ?? nearest
  const rowLeft = (Math.max(0, point.contentWidth) - target.widthPx) / 2
  let best = target.pageIndexes[0]
  let bestDistance = Number.POSITIVE_INFINITY
  for (const [position, pageIndex] of target.pageIndexes.entries()) {
    const left = rowLeft + target.pageLeftsPx[position]
    if (point.x >= left && point.x < left + target.pageWidthsPx[position]) return pageIndex
    const distance = Math.abs(point.x - (left + target.pageWidthsPx[position] / 2))
    if (distance < bestDistance) {
      bestDistance = distance
      best = pageIndex
    }
  }
  return best
}

/** Rows that intersect a viewport window in scroll space. */
export function pdfRowsInViewport(
  rows: readonly PdfPageLayoutRow[],
  window: { startPx: number; endPx: number },
): PdfPageLayoutRow[] {
  return rows.filter(
    (row) => row.topPx + row.heightPx >= window.startPx && row.topPx <= window.endPx,
  )
}
