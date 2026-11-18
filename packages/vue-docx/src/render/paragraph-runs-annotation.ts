// Paragraph-mark tracked-change marker.
//
// A paragraph mark insertion/deletion lives in pPr/rPr and has no run text.
// The renderer appends this pilcrow marker to the paragraph so the revision is
// visible inline in addition to its gutter card.

import { h, type VNode } from "vue"
import type {
  DocxDocumentTheme,
  ParagraphTrackedInlineChange,
} from "@arcships/docx-core"
import { trackedDeletedStyle, trackedInlineStyle } from "@arcships/docx-core"

export function renderParagraphMarkChange(options: {
  keyPrefix: string
  change: ParagraphTrackedInlineChange | undefined
  documentTheme: DocxDocumentTheme
}): VNode | undefined {
  const { keyPrefix, change, documentTheme } = options
  if (!change) return undefined

  const isDeletion = change.kind === "deletion" || change.kind === "move-from"
  const markerStyle = isDeletion
    ? trackedDeletedStyle(documentTheme)
    : trackedInlineStyle({}, change)

  return h(
    "span",
    {
      key: `${keyPrefix}-paragraph-mark`,
      "data-docx-tracked-change": change.kind,
      "data-docx-tracked-change-id": change.id,
      "data-docx-paragraph-mark": "true",
      title: change.author
        ? `${change.author} ${isDeletion ? "deleted" : "inserted"} paragraph mark`
        : isDeletion
        ? "Deleted paragraph mark"
        : "Inserted paragraph mark",
      style: {
        ...markerStyle,
        display: "inline-block",
        minWidth: "0.65em",
        textAlign: "center",
        userSelect: "none",
      },
    },
    "\u00b6"
  )
}
