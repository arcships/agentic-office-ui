// editor-paragraph-range.ts — paragraph-level command targeting
//
// Paragraph-level commands (alignment, line spacing, heading level, paragraph
// style, list toggles) must touch every paragraph spanned by the active
// selection, not only the paragraph that owns the selection anchor.

import type {
  DocModel,
  DocxEditorSelection,
  DocxTextRange,
  ParagraphLocation,
} from "@arcships/docx-core"
import {
  cloneTextRangeLocation,
  paragraphRangeForMutate,
  resolveSelectedParagraphLocation,
} from "@arcships/docx-core"

/**
 * Paragraph locations a paragraph-level command has to mutate.
 *
 * A collapsed caret yields exactly one location; an expanded selection yields
 * every paragraph between its two boundaries (inclusive), including paragraphs
 * that live in table cells.
 */
export function selectedParagraphLocations(
  model: DocModel,
  selection: DocxEditorSelection,
  activeTextRange?: DocxTextRange
): ParagraphLocation[] {
  const anchor = resolveSelectedParagraphLocation(selection, activeTextRange)
  if (!activeTextRange) return [anchor]

  const entries = paragraphRangeForMutate(
    model,
    activeTextRange.start.location,
    activeTextRange.end.location
  )
  if (entries.length === 0) return [anchor]
  return entries.map((entry) => cloneTextRangeLocation(entry.location))
}
