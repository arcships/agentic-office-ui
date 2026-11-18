import type { PageTextSlice, PdfPageGeometry } from "@embedpdf/models"

/**
 * A resolved PDF character position. `charIndex` is a PDFium character index,
 * not a UTF-16 offset inside the extracted text.
 */
export interface PdfSelectionPoint {
  /** Zero-based page index. */
  pageIndex: number
  /** Zero-based PDFium character index on that page. */
  charIndex: number
}

/**
 * Number of PDFium character indices on a page.
 *
 * Mirrors `getTotalCharCount()` from `@embedpdf/plugin-selection` so ranges
 * built here and `rectsWithinSlice()` agree on the last character index.
 */
export function pdfPageCharCount(geometry: PdfPageGeometry | undefined): number {
  if (!geometry || geometry.runs.length === 0) return 0
  const lastRun = geometry.runs[geometry.runs.length - 1]
  return lastRun.charStart + lastRun.glyphs.length
}

/**
 * Order two selection endpoints into document order. Endpoints on the same page
 * follow character order, endpoints on different pages follow page order.
 */
export function orderSelectionPoints(
  a: PdfSelectionPoint,
  b: PdfSelectionPoint,
): { from: PdfSelectionPoint; to: PdfSelectionPoint } {
  if (a.pageIndex !== b.pageIndex) {
    return a.pageIndex < b.pageIndex ? { from: a, to: b } : { from: b, to: a }
  }
  return a.charIndex <= b.charIndex ? { from: a, to: b } : { from: b, to: a }
}

/** Every page index touched by the selection, in document order. */
export function selectionPageIndexes(
  a: PdfSelectionPoint,
  b: PdfSelectionPoint,
): number[] {
  const { from, to } = orderSelectionPoints(a, b)
  const indexes: number[] = []
  for (let pageIndex = from.pageIndex; pageIndex <= to.pageIndex; pageIndex++) indexes.push(pageIndex)
  return indexes
}

/**
 * Split a selection into one `PageTextSlice` per touched page, in document
 * order. Pages without glyphs are skipped, and character indexes are clamped to
 * the page's character count so an empty page cannot produce an invalid slice.
 */
export function selectionPageRanges(
  a: PdfSelectionPoint,
  b: PdfSelectionPoint,
  charCountOf: (pageIndex: number) => number,
): PageTextSlice[] {
  const { from, to } = orderSelectionPoints(a, b)
  const ranges: PageTextSlice[] = []
  for (let pageIndex = from.pageIndex; pageIndex <= to.pageIndex; pageIndex++) {
    const total = Math.floor(charCountOf(pageIndex))
    if (!(total > 0)) continue
    const first = pageIndex === from.pageIndex ? from.charIndex : 0
    const last = pageIndex === to.pageIndex ? to.charIndex : total - 1
    const charIndex = Math.max(0, Math.min(first, total - 1))
    const lastIndex = Math.max(0, Math.min(last, total - 1))
    if (lastIndex < charIndex) continue
    ranges.push({ pageIndex, charIndex, charCount: lastIndex - charIndex + 1 })
  }
  return ranges
}

/**
 * Join per-page extracted text into clipboard text. Empty pages contribute no
 * separator, so a one-page selection keeps its text unchanged.
 */
export function joinSelectionTexts(texts: readonly string[]): string {
  return texts.filter((text) => text.length > 0).join("\n")
}

/**
 * True when extracted text contains replacement characters or NULs, which means
 * the document's Unicode mapping cannot be trusted for copy or references.
 */
export function hasUnmappedPdfText(text: string): boolean {
  return /[\u0000\ufffd\ufffe]/u.test(text)
}
