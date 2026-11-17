// Revision-markup loss detection.
//
// Tracked changes are derived from a paragraph's `sourceXml` (the imported
// OOXML the serializer passes through). Every text mutation clears that field,
// which silently removes the paragraph's revisions from the inline view and the
// gutter: there is nothing left to accept or reject. `countParagraphsLosingRevisionMarkup`
// lets the editor report that loss instead of dropping revisions without a word.
//
// Deliberately cheap: it only compares paragraphs whose `sourceXml` disappeared
// in one transaction and runs a single tag regex on the previous value, so it
// never costs a full markup scan.

import type {
  DocModel,
  ParagraphNode,
  TableCellContentNode
} from "../../engine/types";

/** Revision markup that lives in a paragraph's `sourceXml`. */
const PARAGRAPH_REVISION_MARKUP_PATTERN =
  /<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|tblPrChange|trPrChange|tcPrChange)\b/u;

function paragraphLostRevisionMarkup(
  before: ParagraphNode,
  after: ParagraphNode
): boolean {
  if (!before.sourceXml || after.sourceXml) {
    return false;
  }
  return PARAGRAPH_REVISION_MARKUP_PATTERN.test(before.sourceXml);
}

function countLostCellParagraphs(
  beforeNodes: readonly TableCellContentNode[],
  afterNodes: readonly TableCellContentNode[]
): number {
  const beforeParagraphs = beforeNodes.filter(
    (node): node is ParagraphNode => node.type === "paragraph"
  );
  const afterParagraphs = afterNodes.filter(
    (node): node is ParagraphNode => node.type === "paragraph"
  );

  let lost = 0;
  const shared = Math.min(beforeParagraphs.length, afterParagraphs.length);
  for (let index = 0; index < shared; index += 1) {
    const previous = beforeParagraphs[index];
    const next = afterParagraphs[index];
    if (previous && next && paragraphLostRevisionMarkup(previous, next)) {
      lost += 1;
    }
  }
  return lost;
}

/**
 * Counts paragraphs whose revision markup an edit dropped between two models.
 * Only direct paragraph children and direct table-cell paragraphs are compared;
 * revisions inside nested tables are not counted.
 */
export function countParagraphsLosingRevisionMarkup(
  before: DocModel,
  after: DocModel
): number {
  let lost = 0;

  const shared = Math.min(before.nodes.length, after.nodes.length);
  for (let index = 0; index < shared; index += 1) {
    const previous = before.nodes[index];
    const next = after.nodes[index];
    if (!previous || !next || previous.type !== next.type) {
      continue;
    }

    if (previous.type === "paragraph" && next.type === "paragraph") {
      if (paragraphLostRevisionMarkup(previous, next)) {
        lost += 1;
      }
      continue;
    }

    if (previous.type !== "table" || next.type !== "table") {
      continue;
    }

    const rowCount = Math.min(previous.rows.length, next.rows.length);
    for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
      const previousRow = previous.rows[rowIndex];
      const nextRow = next.rows[rowIndex];
      if (!previousRow || !nextRow) {
        continue;
      }
      const cellCount = Math.min(previousRow.cells.length, nextRow.cells.length);
      for (let cellIndex = 0; cellIndex < cellCount; cellIndex += 1) {
        const previousCell = previousRow.cells[cellIndex];
        const nextCell = nextRow.cells[cellIndex];
        if (!previousCell || !nextCell) {
          continue;
        }
        lost += countLostCellParagraphs(previousCell.nodes, nextCell.nodes);
      }
    }
  }

  return lost;
}
