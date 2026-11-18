// Stable source ids for the block/word model.
//
// The id is the contract between extraction and rendering: the viewer writes the
// same value to `data-docx-source-id`, so geometry resolution can find the DOM
// that represents a block without guessing from text.

import type { DocModel, DocNode, TableCellContentNode } from "../engine/types";

export const DOCX_BODY_PART_NAME = "word/document.xml";
export const DOCX_FOOTNOTES_PART_NAME = "word/footnotes.xml";
export const DOCX_ENDNOTES_PART_NAME = "word/endnotes.xml";

/** Source id of a top-level body node (paragraph or table). */
export function docxNodeSourceId(nodeIndex: number): string {
  return `docx:node:${nodeIndex}`;
}

/** Source id of a header/footer container. */
export function docxStoryContainerSourceId(
  story: "header" | "footer",
  partName: string,
): string {
  return `docx:${story}:${partName}`;
}

/** Source id of a paragraph/table inside a header or footer part. */
export function docxStoryNodeSourceId(
  containerId: string,
  nodeIndex: number,
): string {
  return `${containerId}:node:${nodeIndex}`;
}

/** Source id of a paragraph/table inside a footnote or endnote definition. */
export function docxNoteNodeSourceId(
  kind: "footnote" | "endnote",
  noteId: number,
  nodeIndex: number,
): string {
  return `docx:${kind}:${noteId}:node:${nodeIndex}`;
}

/** Source id of a table row. */
export function docxTableRowSourceId(tableId: string, rowIndex: number): string {
  return `${tableId}/r${rowIndex}`;
}

/** Source id of a table cell. */
export function docxTableCellSourceId(
  tableId: string,
  rowIndex: number,
  cellIndex: number,
): string {
  return `${tableId}/r${rowIndex}/c${cellIndex}`;
}

/** Source id of a paragraph inside a table cell. */
export function docxCellParagraphSourceId(
  cellId: string,
  paragraphIndex: number,
): string {
  return `${cellId}/p${paragraphIndex}`;
}

/** Source id of a nested table inside a table cell. */
export function docxNestedTableSourceId(
  cellId: string,
  contentIndex: number,
): string {
  return `${cellId}/n${contentIndex}`;
}

/** Source id of a figure/embedded image inside a paragraph. */
export function docxImageSourceId(
  paragraphId: string,
  childIndex: number,
): string {
  return `${paragraphId}/child/${childIndex}`;
}

/** Word id for a UTF-16 range in a paragraph block. */
export function docxWordId(
  blockId: string,
  start: number,
  end: number,
): string {
  return `${blockId}/word/${start}-${end}`;
}

interface BlockIdCarrier {
  blockId?: string;
}

function nodeBlockId(node: DocNode | TableCellContentNode): string | undefined {
  return (node as unknown as BlockIdCarrier).blockId;
}

function collectNodeBlockIds(
  node: DocNode | TableCellContentNode,
  seen: Set<string>,
  duplicates: Set<string>,
): void {
  const blockId = nodeBlockId(node);
  if (blockId) {
    if (seen.has(blockId)) duplicates.add(blockId);
    seen.add(blockId);
  }

  if (node.type !== "table") return;
  for (const row of node.rows) {
    for (const cell of row.cells) {
      for (const child of cell.nodes) {
        collectNodeBlockIds(child, seen, duplicates);
      }
    }
  }
}

/**
 * Returns block ids that appear on more than one node within the same container
 * (body, a header/footer part or a note). Models without stored `blockId` values
 * produce an empty list; the derived ids used by `extractDocumentBlocks` are
 * unique by construction.
 */
export function collectDuplicateDocModelBlockIds(model: DocModel): string[] {
  const duplicates = new Set<string>();

  const visitContainer = (nodes: DocNode[] | undefined): void => {
    if (!nodes) return;
    const seen = new Set<string>();
    nodes.forEach((node) => collectNodeBlockIds(node, seen, duplicates));
  };

  visitContainer(model.nodes);
  for (const section of model.metadata.headerSections ?? []) {
    visitContainer(section.nodes);
  }
  for (const section of model.metadata.footerSections ?? []) {
    visitContainer(section.nodes);
  }
  for (const note of model.metadata.footnotes ?? []) visitContainer(note.nodes);
  for (const note of model.metadata.endnotes ?? []) visitContainer(note.nodes);

  return [...duplicates];
}
