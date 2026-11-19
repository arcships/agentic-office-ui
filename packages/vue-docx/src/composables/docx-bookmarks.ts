// Internal-link navigation helpers shared by the DOCX render hosts.
//
// The renderer and the page hosts both need to turn an `#anchor` href into the
// bookmark name the viewer resolves. Keeping the parsing here means paragraph,
// table-cell, header/footer and note hosts agree on encoding, the leading `#`
// and case handling.

import { normalizeDocxBookmarkName } from "@arcships/docx-core"

/** Result of asking the viewer to reveal one bookmark. */
export type DocxBookmarkRevealResult =
  | {
      status: "revealed"
      /** Bookmark name as stored in the document. */
      name: string
      /** Body node that owns the bookmark. */
      nodeIndex: number
      /** Zero-based page currently laying out that node, when paginated. */
      pageIndex?: number
      /** Direct table-cell coordinates for a table bookmark. */
      tableRowIndex?: number
      tableCellIndex?: number
      /** Paragraph text at the bookmark, for affordances and debugging. */
      text: string
    }
  | {
      status: "missing"
      /** Name as requested (without the leading `#`). */
      name: string
    }

/** Transient highlight state published to the page surfaces. */
export interface DocxBookmarkTargetHighlight {
  nodeIndex: number
  tableRowIndex?: number
  tableCellIndex?: number
  /**
   * Changes on every reveal so clicking the same link twice restarts the
   * highlight animation instead of reusing an unchanged prop.
   */
  nonce: number
}

function anchorElementFromTarget(target: EventTarget | null | undefined): Element | undefined {
  if (!target) return undefined
  const element = target as Element
  if (typeof element.closest !== "function") {
    return element.tagName === "A" ? element : undefined
  }
  return element.closest("a[href]") ?? undefined
}

/**
 * Returns the bookmark name of the internal link under `target`, or
 * `undefined` when the event did not land on an `#anchor` link.
 */
export function docxBookmarkNameFromClick(
  target: EventTarget | null | undefined
): string | undefined {
  const anchor = anchorElementFromTarget(target)
  if (!anchor) return undefined
  const href = anchor.getAttribute?.("href") ?? ""
  if (!href.startsWith("#")) return undefined
  return normalizeDocxBookmarkName(href) || undefined
}
