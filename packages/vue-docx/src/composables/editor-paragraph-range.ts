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
  tableCellParagraphs,
  tableCellSelectionRangeCells,
  tableCellSelectionRangeHasMultipleCells,
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
  // A multi-cell table range wins over any caret/text range that happens to sit
  // inside it; the rectangle is the explicit selection the user made.
  if (
    selection.kind === "table-cell" &&
    !selection.tableRuntimeKey &&
    selection.cellRange &&
    tableCellSelectionRangeHasMultipleCells(selection.cellRange)
  ) {
    const table = model.nodes[selection.tableIndex]
    if (table?.type === "table") {
      const locations: ParagraphLocation[] = []
      for (const { rowIndex, cellIndex } of tableCellSelectionRangeCells(
        table,
        selection.cellRange
      )) {
        const cell = table.rows[rowIndex]?.cells[cellIndex]
        if (!cell) continue
        tableCellParagraphs(cell.nodes).forEach((_paragraph, paragraphIndex) => {
          locations.push({
            kind: "table-cell",
            tableIndex: selection.tableIndex,
            rowIndex,
            cellIndex,
            paragraphIndex,
          })
        })
      }
      if (locations.length > 0) return locations
    }
  }
  if (!activeTextRange) return [anchor]

  const entries = paragraphRangeForMutate(
    model,
    activeTextRange.start.location,
    activeTextRange.end.location
  )
  if (entries.length === 0) return [anchor]
  return entries.map((entry) => cloneTextRangeLocation(entry.location))
}
