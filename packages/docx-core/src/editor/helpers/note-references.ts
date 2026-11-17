// Model-based footnote/endnote reference extraction.
//
// The DocModel is the source of truth for note references: the importer
// populates TextRunNode.noteReference for every w:footnoteReference /
// w:endnoteReference run. Reading references from paragraph.sourceXml instead
// breaks as soon as a paragraph is edited (the editor drops sourceXml), which
// made notes vanish from the page while their in-text marker stayed behind.

import type {
  DocNode,
  NoteReference,
  NoteReferenceKind,
  ParagraphNode,
} from "../../engine/types";
import type {
  ParagraphLineRange,
  TableRowRange,
} from "../../layout/page-segmentation-core";
import { tableCellParagraphsRecursively } from "./paragraph-inspect";

function validReference(reference: NoteReference | undefined): NoteReference | undefined {
  if (!reference) {
    return undefined;
  }
  const id = Number(reference.id);
  if (!Number.isFinite(id) || id < 0) {
    return undefined;
  }
  return { kind: reference.kind, id: Math.round(id) };
}

/** Note references carried by a paragraph's text runs, in document order. */
export function paragraphNoteReferences(paragraph: ParagraphNode): NoteReference[] {
  const references: NoteReference[] = [];
  for (const child of paragraph.children) {
    if (child.type !== "text") {
      continue;
    }
    const reference = validReference(child.noteReference);
    if (reference) {
      references.push(reference);
    }
  }
  return references;
}

/**
 * Note references carried by one document node (paragraph or table), optionally
 * narrowed to a single note kind and to the part of the node placed on a page.
 */
export function nodeNoteReferences(
  node: DocNode,
  kind?: NoteReferenceKind,
  tableRowRange?: TableRowRange,
  paragraphLineRange?: ParagraphLineRange
): NoteReference[] {
  const matchesKind = (reference: NoteReference): boolean =>
    kind === undefined || reference.kind === kind;

  if (node.type === "paragraph") {
    // A paragraph split across pages only reports its notes on the first
    // segment, which is the page Word keeps the marker on.
    if (paragraphLineRange && paragraphLineRange.startLineIndex > 0) {
      return [];
    }
    return paragraphNoteReferences(node).filter(matchesKind);
  }

  const references: NoteReference[] = [];
  const startRowIndex = Math.max(0, tableRowRange?.startRowIndex ?? 0);
  const endRowIndex = Math.min(
    node.rows.length,
    tableRowRange?.endRowIndex ?? node.rows.length
  );
  for (let rowIndex = startRowIndex; rowIndex < endRowIndex; rowIndex += 1) {
    const row = node.rows[rowIndex];
    row?.cells.forEach((cell) => {
      tableCellParagraphsRecursively(cell.nodes).forEach((paragraph) => {
        references.push(...paragraphNoteReferences(paragraph).filter(matchesKind));
      });
    });
  }
  return references;
}
