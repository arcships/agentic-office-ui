import { PdfActionType, type PdfBookmarkObject, type PdfLinkTarget } from "@embedpdf/models"

/** One entry of the PDF bookmark/outline tree. */
export interface PdfOutlineItem {
  /** Bookmark title as stored in the document. May be empty for structural entries. */
  title: string
  /** Zero-based target page index inside this document. */
  pageIndex?: number
  /** Absolute URI for bookmarks that open an external link instead of a page. */
  url?: string
  /** Nested outline entries. */
  children: readonly PdfOutlineItem[]
}

interface PdfOutlineTarget {
  pageIndex?: number
  url?: string
}

function normalizeOutlineTitle(title: unknown): string {
  return typeof title === "string" ? title.trim() : ""
}

function pageIndexTarget(pageIndex: unknown, pageCount: number): PdfOutlineTarget {
  if (typeof pageIndex !== "number" || !Number.isInteger(pageIndex)) return {}
  if (pageIndex < 0 || pageIndex >= pageCount) return {}
  return { pageIndex }
}

function outlineTarget(target: PdfLinkTarget | undefined, pageCount: number): PdfOutlineTarget {
  if (!target || typeof target !== "object") return {}
  if (target.type === "destination") {
    return pageIndexTarget(target.destination?.pageIndex, pageCount)
  }
  if (target.type !== "action") return {}
  const action = target.action
  if (!action || typeof action !== "object") return {}
  switch (action.type) {
    case PdfActionType.Goto:
      return pageIndexTarget(action.destination?.pageIndex, pageCount)
    case PdfActionType.URI: {
      const url = typeof action.uri === "string" ? action.uri.trim() : ""
      return url ? { url } : {}
    }
    default:
      return {}
  }
}

function mapBookmark(bookmark: PdfBookmarkObject | undefined, pageCount: number): PdfOutlineItem | undefined {
  if (!bookmark || typeof bookmark !== "object") return undefined
  const children: PdfOutlineItem[] = []
  if (Array.isArray(bookmark.children)) {
    for (const child of bookmark.children) {
      const mapped = mapBookmark(child, pageCount)
      if (mapped) children.push(mapped)
    }
  }
  const target = outlineTarget(bookmark.target, pageCount)
  const title = normalizeOutlineTitle(bookmark.title)
  // A structural entry with neither a title, a target nor children cannot be
  // navigated to and would only add an empty row to the tree.
  if (!title && target.pageIndex === undefined && !target.url && children.length === 0) return undefined
  return {
    title,
    ...(target.pageIndex === undefined ? {} : { pageIndex: target.pageIndex }),
    ...(target.url === undefined ? {} : { url: target.url }),
    children,
  }
}

/**
 * Convert the engine bookmark tree into the navigation model rendered by the
 * PDF surface. Bookmarks pointing outside the document keep their title and
 * children but drop their page index instead of exposing an invalid target.
 */
export function mapPdfBookmarks(
  bookmarks: readonly PdfBookmarkObject[] | undefined,
  pageCount: number,
): readonly PdfOutlineItem[] {
  if (!Array.isArray(bookmarks) || !Number.isFinite(pageCount) || pageCount <= 0) return []
  const outline: PdfOutlineItem[] = []
  for (const bookmark of bookmarks) {
    const mapped = mapBookmark(bookmark, pageCount)
    if (mapped) outline.push(mapped)
  }
  return outline
}
