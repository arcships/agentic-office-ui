// Vertical-merge markup for serialized tables.
//
// The bundled engine build writes `w:gridSpan` from `cell.style.gridSpan`, but
// it does not turn `cell.style.rowSpan` / `cell.style.vMergeContinuation` back
// into `w:vMerge`. A table that keeps its stored markup survives because the
// engine emits that markup verbatim; the moment an edit clears the table's
// `sourceXml` (which every cell edit does) a vertically merged table is rebuilt
// without its merges and re-opens as separate cells.
//
// This patch puts the merge markers back into the generated parts. Cells are
// aligned in document order — which is the same depth-first order as the model
// (`table → row → cell → nested content`) — and only the cell's own `w:tcPr` is
// rewritten, so nested tables keep their own edits. The patch is idempotent: a
// cell whose marker already matches the model is left untouched.

import type { DocModel, DocNode, DocumentNoteDefinition, TableCellNode } from "./types";
import type { OoxmlPackage, OoxmlPart } from "./ooxml-core";
import {
  applyDocumentXmlEdits,
  type DocumentXmlEdit
} from "./document-xml-patch";
import { extractBalancedTagRanges } from "./text-typography";

const DOCUMENT_PART = "word/document.xml";
const FOOTNOTES_PART = "word/footnotes.xml";
const ENDNOTES_PART = "word/endnotes.xml";
const HEADER_FOOTER_PART = /^word\/(?:header|footer)\d+\.xml$/i;

type TableRowMergeKind = "none" | "restart" | "continue";

interface SerializedCellMerge {
  /** `unchanged` cells belong to a table whose stored markup is written verbatim. */
  kind: TableRowMergeKind | "unchanged";
}

interface CellContentSpan {
  contentStart: number;
  contentEnd: number;
  selfClosing: boolean;
}

interface TcPrSpan {
  start: number;
  end: number;
  contentStart: number;
  contentEnd: number;
  selfClosing: boolean;
}

interface ExistingMerge {
  start: number;
  end: number;
  kind: Exclude<TableRowMergeKind, "none">;
}

/**
 * OOXML `w:tcPr` child order. `w:vMerge` is inserted before the first property
 * that must follow it so the rebuilt cell stays schema-valid.
 */
const TC_PR_CHILDREN = [
  "w:cnfStyle",
  "w:tcW",
  "w:gridSpan",
  "w:hMerge",
  "w:vMerge",
  "w:tcBorders",
  "w:shd",
  "w:noWrap",
  "w:tcMar",
  "w:textDirection",
  "w:tcFitText",
  "w:vAlign",
  "w:hideMark"
];
const V_MERGE_ORDER = TC_PR_CHILDREN.indexOf("w:vMerge");

function cellMergeKind(cell: TableCellNode): TableRowMergeKind {
  if (cell.style?.vMergeContinuation) return "continue";
  const rowSpan = cell.style?.rowSpan;
  if (rowSpan !== undefined && Number.isFinite(rowSpan) && rowSpan > 1) {
    return "restart";
  }
  return "none";
}

/**
 * Model cells in serialized document order: a cell, then the nested tables
 * inside it, then the next sibling cell.
 */
function collectCellMerges(
  nodes: readonly DocNode[],
  merges: SerializedCellMerge[]
): void {
  for (const node of nodes) {
    if (node.type !== "table") continue;
    // A table with stored markup is emitted verbatim, merges included; its
    // cells must stay untouched. Nested tables are still visited because an
    // edit below the stored table clears the nested table only.
    const stored = node.sourceXml !== undefined;
    for (const row of node.rows) {
      for (const cell of row.cells) {
        merges.push({ kind: stored ? "unchanged" : cellMergeKind(cell) });
        collectCellMerges(cell.nodes, merges);
      }
    }
  }
}

function nodesNeedTableRowMerges(nodes: readonly DocNode[]): boolean {
  for (const node of nodes) {
    if (node.type !== "table") continue;
    const stored = node.sourceXml !== undefined;
    for (const row of node.rows) {
      for (const cell of row.cells) {
        if (!stored && cellMergeKind(cell) !== "none") return true;
        if (nodesNeedTableRowMerges(cell.nodes)) return true;
      }
    }
  }
  return false;
}

function sectionStoryNodes(model: DocModel): Map<string, DocNode[]> {
  const stories = new Map<string, DocNode[]>();
  const add = (
    sections: ReadonlyArray<{ partName: string; nodes: DocNode[] }> | undefined
  ): void => {
    for (const section of sections ?? []) {
      const key = section.partName.toLowerCase();
      if (!stories.has(key)) stories.set(key, section.nodes);
    }
  };
  add(model.metadata.headerSections);
  add(model.metadata.footerSections);
  for (const section of model.metadata.sections ?? []) {
    add(section.headerSections);
    add(section.footerSections);
  }
  return stories;
}

/** True when any story holds a table with a vertical merge. */
export function modelNeedsTableRowMergeSerialization(model: DocModel): boolean {
  if (nodesNeedTableRowMerges(model.nodes)) return true;
  for (const story of sectionStoryNodes(model).values()) {
    if (nodesNeedTableRowMerges(story)) return true;
  }
  const noteNeedsMerge = (notes: DocumentNoteDefinition[] | undefined): boolean =>
    (notes ?? []).some((note) => note.nodes && nodesNeedTableRowMerges(note.nodes));
  return noteNeedsMerge(model.metadata.footnotes) ||
    noteNeedsMerge(model.metadata.endnotes);
}

/**
 * Content spans of every `<w:tc>` in a part, in document order (an outer cell
 * precedes the cells of the tables nested inside it).
 */
function locateCellContentSpans(xml: string): CellContentSpan[] {
  const pattern = /<(\/)?w:tc\b[^>]*?(\/)?>/gi;
  const spans: CellContentSpan[] = [];
  const open: number[] = [];

  for (const match of xml.matchAll(pattern)) {
    const start = match.index ?? 0;
    if (match[1] === "/") {
      const spanIndex = open.pop();
      if (spanIndex === undefined) continue;
      spans[spanIndex].contentEnd = start;
      continue;
    }
    const contentStart = start + match[0].length;
    open.push(spans.length);
    spans.push({
      contentStart,
      contentEnd: contentStart,
      selfClosing: match[2] === "/"
    });
  }

  return spans;
}

function findTcPr(content: string): TcPrSpan | undefined {
  const open = /^\s*<w:tcPr\b[^>]*?(\/)?>/i.exec(content);
  if (!open || open.index === undefined) return undefined;
  const start = open.index + (open[0].length - open[0].trimStart().length);
  const end = open.index + open[0].length;
  if (open[1] === "/") {
    return { start, end, contentStart: end, contentEnd: end, selfClosing: true };
  }
  const close = content.indexOf("</w:tcPr>", end);
  if (close < 0) return undefined;
  return {
    start,
    end: close + "</w:tcPr>".length,
    contentStart: end,
    contentEnd: close,
    selfClosing: false
  };
}

function findExistingMerge(
  content: string,
  tcPr: TcPrSpan
): ExistingMerge | undefined {
  if (tcPr.selfClosing) return undefined;
  const tcPrContent = content.slice(tcPr.contentStart, tcPr.contentEnd);
  const match = /<w:vMerge\b([^>]*?)\/?>/i.exec(tcPrContent);
  if (!match || match.index === undefined) return undefined;
  const kind = /w:val\s*=\s*"restart"/i.test(match[1]) ? "restart" : "continue";
  return {
    start: tcPr.contentStart + match.index,
    end: tcPr.contentStart + match.index + match[0].length,
    kind
  };
}

function tcPrInsertionPoint(
  content: string,
  contentStart: number,
  contentEnd: number
): number {
  const tcPrContent = content.slice(contentStart, contentEnd);
  const pattern = /<(\/?)([A-Za-z_][\w.:-]*)\b[^>]*?(\/)?>/g;
  for (const match of tcPrContent.matchAll(pattern)) {
    if (match[1] === "/") continue;
    const order = TC_PR_CHILDREN.indexOf(match[2].toLowerCase());
    if (order > V_MERGE_ORDER) {
      return contentStart + (match.index ?? 0);
    }
  }
  return contentEnd;
}

function vMergeElement(kind: Exclude<TableRowMergeKind, "none">): string {
  return kind === "restart" ? '<w:vMerge w:val="restart"/>' : "<w:vMerge/>";
}

/**
 * Edits that make one serialized cell's merge marker match the model. Offsets
 * are relative to the cell content.
 */
function cellMergeEdits(
  content: string,
  merge: SerializedCellMerge
): DocumentXmlEdit[] {
  if (merge.kind === "unchanged") return [];
  const tcPr = findTcPr(content);
  const existing = tcPr ? findExistingMerge(content, tcPr) : undefined;
  if (existing?.kind === merge.kind) return [];

  if (existing) {
    return [
      {
        start: existing.start,
        end: existing.end,
        replacement: merge.kind === "none" ? "" : vMergeElement(merge.kind)
      }
    ];
  }
  if (merge.kind === "none") return [];

  const element = vMergeElement(merge.kind);
  if (!tcPr) {
    return [{ start: 0, end: 0, replacement: `<w:tcPr>${element}</w:tcPr>` }];
  }
  if (tcPr.selfClosing) {
    return [
      {
        start: tcPr.start,
        end: tcPr.end,
        replacement: `<w:tcPr>${element}</w:tcPr>`
      }
    ];
  }
  const insertAt = tcPrInsertionPoint(content, tcPr.contentStart, tcPr.contentEnd);
  return [{ start: insertAt, end: insertAt, replacement: element }];
}

/**
 * Rewrites the merge markers of every cell in a story's XML.
 *
 * Cell alignment is all-or-nothing: when the part has a different number of
 * `<w:tc>` elements than the model (text-box content, for example, is not part
 * of the model), the story is returned unchanged rather than patched at the
 * wrong offset.
 */
export function applyTableRowMergesToStoryXml(
  xml: string,
  nodes: readonly DocNode[]
): string {
  const merges: SerializedCellMerge[] = [];
  collectCellMerges(nodes, merges);
  if (!merges.some((merge) => merge.kind === "restart" || merge.kind === "continue")) {
    return xml;
  }

  const cells = locateCellContentSpans(xml);
  if (cells.length !== merges.length) return xml;

  const edits: DocumentXmlEdit[] = [];
  for (let index = 0; index < cells.length; index += 1) {
    const cell = cells[index];
    if (cell.selfClosing) continue;
    const content = xml.slice(cell.contentStart, cell.contentEnd);
    for (const edit of cellMergeEdits(content, merges[index])) {
      edits.push({
        start: cell.contentStart + edit.start,
        end: cell.contentStart + edit.end,
        replacement: edit.replacement
      });
    }
  }
  return applyDocumentXmlEdits(xml, edits);
}

/**
 * Footnote/endnote bodies are separate elements inside one part; each is
 * patched with its own definition so the order of `metadata.footnotes` cannot
 * scramble the alignment.
 */
function applyNoteMergesToXml(
  xml: string,
  tagName: "w:footnote" | "w:endnote",
  notes: DocumentNoteDefinition[] | undefined
): string {
  if (!notes || notes.length === 0) return xml;
  const byId = new Map<number, DocumentNoteDefinition>();
  for (const note of notes) {
    if (note.nodes) byId.set(note.id, note);
  }
  if (byId.size === 0) return xml;

  const edits: DocumentXmlEdit[] = [];
  for (const range of extractBalancedTagRanges(xml, tagName)) {
    const idMatch = /\bw:id\s*=\s*"(-?\d+)"/.exec(range.openTag);
    if (!idMatch) continue;
    const note = byId.get(Number(idMatch[1]));
    if (!note?.nodes) continue;
    const body = xml.slice(range.start, range.end);
    const patched = applyTableRowMergesToStoryXml(body, note.nodes);
    if (patched !== body) {
      edits.push({ start: range.start, end: range.end, replacement: patched });
    }
  }
  return applyDocumentXmlEdits(xml, edits);
}

/**
 * Patches one package part. Parts that hold no story (styles, numbering,
 * settings, media) are returned unchanged.
 */
export function applyTableRowMergesToPartContent(
  partName: string,
  content: string,
  model: DocModel
): string {
  const name = partName.toLowerCase();
  if (name === DOCUMENT_PART) {
    return applyTableRowMergesToStoryXml(content, model.nodes);
  }
  if (HEADER_FOOTER_PART.test(name)) {
    const nodes = sectionStoryNodes(model).get(name);
    return nodes ? applyTableRowMergesToStoryXml(content, nodes) : content;
  }
  if (name === FOOTNOTES_PART) {
    return applyNoteMergesToXml(content, "w:footnote", model.metadata.footnotes);
  }
  if (name === ENDNOTES_PART) {
    return applyNoteMergesToXml(content, "w:endnote", model.metadata.endnotes);
  }
  return content;
}

/** True for header/footer/footnote/endnote parts that carry story nodes. */
export function isTableMergeStoryPart(partName: string): boolean {
  const name = partName.toLowerCase();
  return (
    name === FOOTNOTES_PART ||
    name === ENDNOTES_PART ||
    HEADER_FOOTER_PART.test(name)
  );
}

/** Body `word/document.xml` only, for `modelToDocumentXml`. */
export function applyTableRowMergesToDocumentXml(
  documentXml: string,
  model: DocModel
): string {
  return applyTableRowMergesToStoryXml(documentXml, model.nodes);
}

/**
 * Patches every story part of a parsed package. Used by `serializeDocModel`,
 * which returns a package instead of bytes.
 */
export function applyTableRowMergesToPackage(
  pkg: OoxmlPackage,
  model: DocModel
): OoxmlPackage {
  if (!modelNeedsTableRowMergeSerialization(model)) return pkg;
  let parts: Map<string, OoxmlPart> | undefined;
  for (const [partName, part] of pkg.parts) {
    const content = applyTableRowMergesToPartContent(partName, part.content, model);
    if (content === part.content) continue;
    parts ??= new Map(pkg.parts);
    parts.set(partName, { name: part.name, content });
  }
  return parts ? { parts, binaryAssets: pkg.binaryAssets } : pkg;
}
