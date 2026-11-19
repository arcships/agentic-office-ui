import type { DocumentPageNotes } from "@arcships/docx-core"
import type { DocxSearchRegion } from "./useDocxSearch"

/** Search target whose page the viewer must reveal before highlighting. */
export interface DocxSearchPageTarget {
  region: DocxSearchRegion
  partName?: string
  noteId?: number
}

interface SearchPagePresentation {
  headerSection?: { partName: string }
  footerSection?: { partName: string }
}

/**
 * Resolve the page that renders a header/footer part or a footnote/endnote
 * body. Body matches are placed by `scrollToNode` instead.
 *
 * Headers/footers repeat on every page of a section, so the first page that
 * renders the part is the navigation target. Notes are printed on the page
 * that owns their first reference (`buildDocumentPageNotes`).
 */
export function findDocxSearchTargetPage(
  target: DocxSearchPageTarget,
  presentations: readonly SearchPagePresentation[],
  notes: readonly DocumentPageNotes[],
): number | undefined {
  if (target.region === "header" || target.region === "footer") {
    const sectionKey = target.region === "header" ? "headerSection" : "footerSection"
    const pageIndex = presentations.findIndex((presentation) => {
      const section = presentation[sectionKey]
      if (!section) return false
      return target.partName ? section.partName === target.partName : true
    })
    return pageIndex >= 0 ? pageIndex : undefined
  }

  if (target.region === "footnote" || target.region === "endnote") {
    if (target.noteId === undefined) return undefined
    const entriesKey = target.region === "footnote" ? "footnotes" : "endnotes"
    const pageIndex = notes.findIndex((pageNotes) =>
      pageNotes[entriesKey].some((entry) => entry.note.id === target.noteId),
    )
    return pageIndex >= 0 ? pageIndex : undefined
  }

  return undefined
}
