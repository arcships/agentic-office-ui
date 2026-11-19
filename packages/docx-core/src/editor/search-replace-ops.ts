// search-replace-ops.ts — find & replace over the editable document model.
//
// A search match is a UTF-16 offset pair inside one paragraph's plain text.
// Replacing it must not rewrite the paragraph with a flat string: the runs that
// carry the document's formatting have to survive, and the inserted text takes
// the style of the character it replaced ("preserve run styles where
// possible"). These helpers therefore splice the text runs, keep every other
// child (images, form fields) in place, remap comment anchors and drop the
// stale `sourceXml` so export re-emits the paragraph.

import type { DocModel, ParagraphNode, TextRunNode } from "../engine/types";
import { cloneDocModelShallow, cloneDocNode } from "../engine/clone";
import { invalidateParagraphSourceXml } from "../engine/section-break";
import { remapParagraphCommentRanges } from "../engine/comment-ranges";
import { paragraphTextContent } from "../engine/paragraph-text";
import type { DocxTextRangeLocation } from "./helpers/editor-types";
import {
  cloneTextStyle,
  firstTextStyleAtOffset,
  linkAtOffset,
} from "./helpers/text-mutation";

/** One step from a table into a table nested inside one of its cells. */
export interface DocxReplaceNestedTableStep {
  rowIndex: number;
  cellIndex: number;
  /** Content-node index of the nested table inside the cell's `nodes`. */
  contentIndex: number;
}

/** Paragraph inside a nested table, addressed by its full nesting path. */
export interface DocxReplaceParagraphLocationInNestedTableCell {
  kind: "nested-table-cell";
  /** Top-level table node index within the document body. */
  tableIndex: number;
  /** Nesting path from the top-level table into the innermost table. */
  path: readonly DocxReplaceNestedTableStep[];
  rowIndex: number;
  cellIndex: number;
  paragraphIndex: number;
}

/**
 * Paragraph address a replace command can reach. Body paragraphs and table-cell
 * paragraphs (including nested tables) are editable; headers, footers and note
 * bodies are not, so they never appear here.
 */
export type DocxReplaceParagraphLocation =
  | DocxTextRangeLocation
  | DocxReplaceParagraphLocationInNestedTableCell;

/**
 * One requested replacement: the paragraph to edit plus the offsets of the
 * matched text inside that paragraph's plain text.
 */
export interface DocxParagraphTextReplacement {
  location: DocxReplaceParagraphLocation;
  start: number;
  end: number;
  replacement: string;
}

export interface DocxReplaceParagraphResult {
  model: DocModel;
  /** False when the offsets did not change the text (for example an identical replacement). */
  changed: boolean;
}

function clampOffset(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return Math.max(min, Math.min(Math.round(value), max));
}

function sameStyle(left?: TextRunNode["style"], right?: TextRunNode["style"]): boolean {
  if (left === right) return true;
  if (!left || !right) return false;
  return JSON.stringify(left) === JSON.stringify(right);
}

function mergeAdjacentRuns(children: ParagraphNode["children"]): ParagraphNode["children"] {
  const merged: ParagraphNode["children"] = [];
  for (const child of children) {
    const previous = merged[merged.length - 1];
    // Note references and links are run-scoped metadata, so only plain text
    // runs with identical formatting are safe to join.
    if (
      child.type === "text" &&
      child.text.length > 0 &&
      previous?.type === "text" &&
      previous.text.length > 0 &&
      !child.noteReference &&
      !previous.noteReference &&
      child.link === previous.link &&
      sameStyle(child.style, previous.style)
    ) {
      merged[merged.length - 1] = { ...previous, text: previous.text + child.text };
      continue;
    }
    merged.push(child);
  }
  return merged;
}

function textRunWithLength(run: TextRunNode, text: string): TextRunNode {
  return { ...run, text };
}

/**
 * Replaces the plain-text range `[start, end)` of one paragraph in place.
 *
 * The paragraph must already be a private clone; siblings keep their identity.
 * Returns true when the text changed. The inserted text inherits the style and
 * link of the run it replaces, and runs left and right of the match are only
 * split where the range starts and ends.
 */
export function replaceTextRangeInParagraph(
  paragraph: ParagraphNode,
  start: number,
  end: number,
  replacement: string,
): boolean {
  const previousText = paragraphTextContent(paragraph);
  const safeStart = clampOffset(start, 0, previousText.length);
  const safeEnd = clampOffset(end, safeStart, previousText.length);
  const removedText = previousText.slice(safeStart, safeEnd);
  if (removedText === replacement) return false;

  const insertedStyle = firstTextStyleAtOffset(paragraph, safeStart, false);
  const insertedLink = linkAtOffset(paragraph, safeStart, false);

  const children: ParagraphNode["children"] = [];
  let cursor = 0;
  let inserted = false;
  const insertReplacement = (): void => {
    if (inserted) return;
    inserted = true;
    if (!replacement) return;
    children.push({
      type: "text",
      text: replacement,
      style: cloneTextStyle(insertedStyle),
      ...(insertedLink ? { link: insertedLink } : undefined),
    });
  };

  for (const child of paragraph.children) {
    if (child.type !== "text") {
      // A zero-width anchor (image, form field) that sits exactly at the edit
      // point keeps its place: the replacement goes before it.
      if (cursor >= safeStart && cursor <= safeEnd) insertReplacement();
      children.push(child);
      continue;
    }

    const runStart = cursor;
    const runEnd = runStart + child.text.length;
    cursor = runEnd;

    if (runEnd <= safeStart) {
      children.push(child);
      continue;
    }
    if (runStart >= safeEnd) {
      insertReplacement();
      children.push(child);
      continue;
    }

    if (runStart < safeStart) {
      children.push(textRunWithLength(child, child.text.slice(0, safeStart - runStart)));
    }
    insertReplacement();
    if (runEnd > safeEnd) {
      children.push(textRunWithLength(child, child.text.slice(safeEnd - runStart)));
    }
  }
  insertReplacement();

  if (children.length === 0) {
    children.push({ type: "text", text: "", style: cloneTextStyle(insertedStyle) });
  }

  paragraph.children = mergeAdjacentRuns(children);
  const nextText = paragraphTextContent(paragraph);
  paragraph.commentRanges = remapParagraphCommentRanges(
    previousText,
    nextText,
    paragraph.commentRanges,
  );
  invalidateParagraphSourceXml(paragraph);
  return true;
}

function paragraphInCell(
  table: Extract<DocModel["nodes"][number], { type: "table" }>,
  rowIndex: number,
  cellIndex: number,
  paragraphIndex: number,
): ParagraphNode | undefined {
  const cell = table.rows[rowIndex]?.cells[cellIndex];
  if (!cell) return undefined;
  let seen = 0;
  for (const node of cell.nodes) {
    if (node.type !== "paragraph") continue;
    if (seen === paragraphIndex) return node;
    seen += 1;
  }
  return undefined;
}

function tableAtNestingPath(
  table: Extract<DocModel["nodes"][number], { type: "table" }>,
  path: readonly DocxReplaceNestedTableStep[],
): Extract<DocModel["nodes"][number], { type: "table" }> | undefined {
  let current = table;
  for (const step of path) {
    const cell = current.rows[step.rowIndex]?.cells[step.cellIndex];
    const next = cell?.nodes[step.contentIndex];
    if (!next || next.type !== "table") return undefined;
    current = next;
  }
  return current;
}

/**
 * Applies one replacement to a copy-on-write revision of the model.
 *
 * Only the path to the paragraph is copied: the top-level node (body paragraph
 * or the table that owns the cell) plus the paragraph itself when it is nested
 * inside a deep-cloned table. Returns `undefined` when the target does not
 * exist or the text would not change, so callers can skip the undo entry.
 */
export function replaceTextInParagraphLocation(
  model: DocModel,
  edit: DocxParagraphTextReplacement,
): DocxReplaceParagraphResult | undefined {
  const { location } = edit;
  const next = cloneDocModelShallow(model);

  if (location.kind === "paragraph") {
    const node = next.nodes[location.nodeIndex];
    if (!node || node.type !== "paragraph") return undefined;
    const paragraph = cloneDocNode(node);
    next.nodes[location.nodeIndex] = paragraph;
    if (!replaceTextRangeInParagraph(paragraph, edit.start, edit.end, edit.replacement)) {
      return undefined;
    }
    return { model: next, changed: true };
  }

  const node = next.nodes[location.tableIndex];
  if (!node || node.type !== "table") return undefined;
  const table = cloneDocNode(node);
  next.nodes[location.tableIndex] = table;

  const hostingTable = location.kind === "nested-table-cell"
    ? tableAtNestingPath(table, location.path)
    : table;
  if (!hostingTable) return undefined;
  const paragraph = paragraphInCell(
    hostingTable,
    location.rowIndex,
    location.cellIndex,
    location.paragraphIndex,
  );
  if (!paragraph) return undefined;
  if (!replaceTextRangeInParagraph(paragraph, edit.start, edit.end, edit.replacement)) {
    return undefined;
  }
  // The table's XML wraps every cell, so the whole table has to be re-emitted
  // once one of its paragraphs changed.
  table.sourceXml = undefined;
  return { model: next, changed: true };
}

/**
 * Applies replacements grouped per paragraph.
 *
 * Offsets are resolved against the revision the caller computed them on, so the
 * edits for one paragraph are applied from the end of the paragraph backwards.
 * Edits in different paragraphs never shift each other's offsets.
 */
export function replaceTextInDocxParagraphs(
  model: DocModel,
  edits: readonly DocxParagraphTextReplacement[],
): { model: DocModel; applied: number } {
  const byLocation = new Map<string, DocxParagraphTextReplacement[]>();
  const order: string[] = [];
  for (const edit of edits) {
    if (edit.start < 0 || edit.end <= edit.start) continue;
    const key = JSON.stringify(edit.location);
    const bucket = byLocation.get(key);
    if (bucket) {
      bucket.push(edit);
      continue;
    }
    byLocation.set(key, [edit]);
    order.push(key);
  }

  let current = model;
  let applied = 0;
  for (const key of order) {
    const bucket = byLocation.get(key) ?? [];
    const descending = [...bucket].sort((left, right) => right.start - left.start);
    for (const edit of descending) {
      const result = replaceTextInParagraphLocation(current, edit);
      if (!result) continue;
      current = result.model;
      applied += 1;
    }
  }
  return { model: current, applied };
}
