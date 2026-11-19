// Floating-table markup for serialized tables.
//
// Word stores a floating table's position and text-wrap distances in
// `w:tblPr/w:tblpPr`. The bundled engine build writes `w:tblW`, `w:tblLayout`
// and the borders from `TableStyle` but ignores `TableStyle.floating`, so the
// moment an edit clears the table's `sourceXml` — which every cell edit does —
// a table that was floating in the imported file comes back inline and its
// drop position and wrap distances are lost on export.
//
// This patch puts `w:tblpPr` back into the generated parts. Tables are aligned
// in document order, which is the same depth-first order as the model
// (`table → row → cell → nested content`), and only the table's own
// `w:tblPr` is rewritten, so nested tables keep their own edits. The patch is
// idempotent: a table whose `w:tblpPr` already matches the model is left
// untouched.
//
// Header and footer parts are rebuilt from the model by the same engine path as
// the body, so they get the same patch. Footnote and endnote parts are copied
// from the base package by this engine build, so they keep their stored markup
// (and their own `w:tblpPr`) without help.

import type { DocModel, DocNode, DocumentNoteDefinition, TableStyle } from "./types";
import type { OoxmlPackage, OoxmlPart } from "./ooxml-core";
import {
  applyDocumentXmlEdits,
  type DocumentXmlEdit,
} from "./document-xml-patch";

const DOCUMENT_PART = "word/document.xml";
const HEADER_FOOTER_PART = /^word\/(?:header|footer)\d+\.xml$/i;

type TableFloatingStyle = NonNullable<TableStyle["floating"]>;

interface SerializedTableFloating {
  /** `verbatim` tables are written from their stored markup, `tblpPr` included. */
  verbatim: boolean;
  floating?: TableFloatingStyle;
}

interface TableOpenTag {
  start: number;
  end: number;
}

interface TblPrSpan {
  start: number;
  end: number;
  contentStart: number;
  contentEnd: number;
  selfClosing: boolean;
}

/**
 * OOXML `w:tblPr` child order. `w:tblpPr` is written after `w:tblStyle` and
 * before everything else so the rebuilt table stays schema-valid.
 */
const TBL_PR_CHILDREN = [
  "w:tblStyle",
  "w:tblpPr",
  "w:tblOverlap",
  "w:bidiVisual",
  "w:tblStyleRowBandSize",
  "w:tblStyleColBandSize",
  "w:tblW",
  "w:jc",
  "w:tblCellSpacing",
  "w:tblInd",
  "w:tblBorders",
  "w:shd",
  "w:tblLayout",
  "w:tblCellMar",
  "w:tblLook",
];
const TBLP_PR_ORDER = TBL_PR_CHILDREN.indexOf("w:tblpPr");
/** Lower-case names for the case-insensitive scan of a serialized `w:tblPr`. */
const TBL_PR_CHILDREN_LOWER = TBL_PR_CHILDREN.map((name) => name.toLowerCase());

/** `w:tblpPr` attribute order Word writes. */
const TBLP_PR_ATTRIBUTES: Array<[keyof TableFloatingStyle, string]> = [
  ["leftFromTextTwips", "w:leftFromText"],
  ["rightFromTextTwips", "w:rightFromText"],
  ["topFromTextTwips", "w:topFromText"],
  ["bottomFromTextTwips", "w:bottomFromText"],
  ["verticalAnchor", "w:vertAnchor"],
  ["horizontalAnchor", "w:horzAnchor"],
  ["xTwips", "w:tblpX"],
  ["horizontalAlign", "w:tblpXSpec"],
  ["yTwips", "w:tblpY"],
  ["verticalAlign", "w:tblpYSpec"],
];

const VERTICAL_ANCHORS = new Set(["text", "margin", "page"]);
const HORIZONTAL_ANCHORS = new Set(["text", "margin", "page"]);
const TABLE_ALIGNMENTS = new Set(["left", "center", "right", "inside", "outside"]);
const TABLE_VERTICAL_ALIGNMENTS = new Set([
  "top",
  "center",
  "bottom",
  "inside",
  "outside",
]);

/**
 * Model tables in serialized document order: a table, then the tables nested in
 * its cells, then the next sibling table.
 */
function collectTableFloating(
  nodes: readonly DocNode[],
  tables: SerializedTableFloating[]
): void {
  for (const node of nodes) {
    if (node.type !== "table") continue;
    tables.push({
      verbatim: node.sourceXml !== undefined,
      floating: node.style?.floating,
    });
    for (const row of node.rows) {
      for (const cell of row.cells) {
        collectTableFloating(cell.nodes, tables);
      }
    }
  }
}

function nodesNeedTableFloating(nodes: readonly DocNode[]): boolean {
  for (const node of nodes) {
    if (node.type !== "table") continue;
    if (node.sourceXml === undefined && hasFloatingAttributes(node.style?.floating)) {
      return true;
    }
    for (const row of node.rows) {
      for (const cell of row.cells) {
        if (nodesNeedTableFloating(cell.nodes)) return true;
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

function hasFloatingAttributes(floating: TableFloatingStyle | undefined): boolean {
  if (!floating) return false;
  return tblpPrElement(floating) !== undefined;
}

/** True when any rebuilt table in any story has to write `w:tblpPr`. */
export function modelNeedsTableFloatingSerialization(model: DocModel): boolean {
  if (nodesNeedTableFloating(model.nodes)) return true;
  for (const story of sectionStoryNodes(model).values()) {
    if (nodesNeedTableFloating(story)) return true;
  }
  const noteNeedsFloating = (notes: DocumentNoteDefinition[] | undefined): boolean =>
    (notes ?? []).some((note) => note.nodes && nodesNeedTableFloating(note.nodes));
  return noteNeedsFloating(model.metadata.footnotes) ||
    noteNeedsFloating(model.metadata.endnotes);
}

/** True for header/footer/footnote/endnote parts that carry story nodes. */
export function isTableFloatingStoryPart(partName: string): boolean {
  return HEADER_FOOTER_PART.test(partName.toLowerCase());
}

/** `<w:tblpPr …/>` for a model style, or `undefined` when nothing is set. */
export function tblpPrElement(
  floating: TableFloatingStyle | undefined
): string | undefined {
  if (!floating) return undefined;
  const attributes: string[] = [];
  for (const [key, attribute] of TBLP_PR_ATTRIBUTES) {
    const raw = floating[key];
    if (key === "horizontalAnchor") {
      const value = normalizeAnchor(raw, HORIZONTAL_ANCHORS);
      if (value) attributes.push(`${attribute}="${value}"`);
      continue;
    }
    if (key === "verticalAnchor") {
      const value = normalizeAnchor(raw, VERTICAL_ANCHORS);
      if (value) attributes.push(`${attribute}="${value}"`);
      continue;
    }
    if (key === "horizontalAlign" || key === "verticalAlign") {
      const value = normalizeAlignment(
        raw,
        key === "verticalAlign" ? TABLE_VERTICAL_ALIGNMENTS : TABLE_ALIGNMENTS
      );
      if (value) attributes.push(`${attribute}="${value}"`);
      continue;
    }
    if (typeof raw === "number" && Number.isFinite(raw)) {
      attributes.push(`${attribute}="${Math.round(raw)}"`);
    }
  }
  if (attributes.length === 0) return undefined;
  return `<w:tblpPr ${attributes.join(" ")}/>`;
}

function normalizeAnchor(
  value: unknown,
  allowed: ReadonlySet<string>
): string | undefined {
  if (typeof value !== "string") return undefined;
  const normalized = value.trim().toLowerCase();
  return allowed.has(normalized) ? normalized : undefined;
}

function normalizeAlignment(
  value: unknown,
  allowed: ReadonlySet<string>
): string | undefined {
  if (typeof value !== "string") return undefined;
  const normalized = value.trim().toLowerCase();
  return allowed.has(normalized) ? normalized : undefined;
}

/** Open tags of every `<w:tbl>` in serialized document order. */
function locateTableOpenTags(xml: string): TableOpenTag[] {
  const pattern = /<w:tbl(?=[\s>])[^>]*>/gi;
  const tags: TableOpenTag[] = [];
  for (const match of xml.matchAll(pattern)) {
    const start = match.index ?? 0;
    tags.push({ start, end: start + match[0].length });
  }
  return tags;
}

/**
 * The table's own `w:tblPr`. It has to be the first child of `w:tbl`, so the
 * scan starts right after the table's open tag and only accepts whitespace in
 * front of it — a `w:tblPr` of a nested table can never be matched here.
 */
function findTblPr(xml: string, openTagEnd: number): TblPrSpan | undefined {
  const rest = xml.slice(openTagEnd, openTagEnd + 4096);
  const open = /^\s*<w:tblPr\b[^>]*?(\/)?>/i.exec(rest);
  if (!open) return undefined;
  const leading = open[0].length - open[0].trimStart().length;
  const start = openTagEnd + leading;
  const end = openTagEnd + open[0].length;
  if (open[1] === "/") {
    return { start, end, contentStart: end, contentEnd: end, selfClosing: true };
  }
  const close = xml.indexOf("</w:tblPr>", end);
  if (close < 0) return undefined;
  return {
    start,
    end: close + "</w:tblPr>".length,
    contentStart: end,
    contentEnd: close,
    selfClosing: false,
  };
}

function findExistingTblpPr(
  content: string,
  tblPr: TblPrSpan
): { start: number; end: number; element: string } | undefined {
  if (tblPr.selfClosing) return undefined;
  const inner = content.slice(tblPr.contentStart, tblPr.contentEnd);
  const match = /<w:tblpPr\b[^>]*?(\/)?>/i.exec(inner);
  if (!match || match.index === undefined) return undefined;
  return {
    start: tblPr.contentStart + match.index,
    end: tblPr.contentStart + match.index + match[0].length,
    element: match[0],
  };
}

function tblPrInsertionPoint(xml: string, tblPr: TblPrSpan): number {
  const inner = xml.slice(tblPr.contentStart, tblPr.contentEnd);
  const pattern = /<(\/?)([A-Za-z_][\w.:-]*)\b[^>]*?(\/)?>/g;
  for (const match of inner.matchAll(pattern)) {
    if (match[1] === "/") continue;
    const order = TBL_PR_CHILDREN_LOWER.indexOf(match[2].toLowerCase());
    if (order > TBLP_PR_ORDER) {
      return tblPr.contentStart + (match.index ?? 0);
    }
  }
  return tblPr.contentEnd;
}

/**
 * Edits that make one serialized table carry (or drop) the model's `w:tblpPr`.
 * Offsets are absolute into `xml`.
 */
function tableFloatingEdits(
  xml: string,
  openTag: TableOpenTag,
  expected: SerializedTableFloating
): DocumentXmlEdit[] {
  if (expected.verbatim) return [];
  const element = tblpPrElement(expected.floating);
  const tblPr = findTblPr(xml, openTag.end);

  if (!tblPr) {
    if (!element) return [];
    return [
      {
        start: openTag.end,
        end: openTag.end,
        replacement: `<w:tblPr>${element}</w:tblPr>`,
      },
    ];
  }

  const existing = findExistingTblpPr(xml, tblPr);
  if (!element) {
    return existing ? [{ start: existing.start, end: existing.end, replacement: "" }] : [];
  }
  if (existing) {
    // Word also writes the attributes in a fixed order, so a string compare is
    // enough to keep the patch idempotent.
    return existing.element === element
      ? []
      : [{ start: existing.start, end: existing.end, replacement: element }];
  }
  if (tblPr.selfClosing) {
    return [
      {
        start: tblPr.start,
        end: tblPr.end,
        replacement: `<w:tblPr>${element}</w:tblPr>`,
      },
    ];
  }
  const insertAt = tblPrInsertionPoint(xml, tblPr);
  return [{ start: insertAt, end: insertAt, replacement: element }];
}

/**
 * Rewrites the floating position of every table in a story's XML.
 *
 * Table alignment is all-or-nothing: when the part has a different number of
 * `<w:tbl>` elements than the model (text-box content, for example, is not part
 * of the model), the story is returned unchanged rather than patched at the
 * wrong offset.
 */
export function applyTableFloatingToStoryXml(
  xml: string,
  nodes: readonly DocNode[]
): string {
  const tables: SerializedTableFloating[] = [];
  collectTableFloating(nodes, tables);
  if (!tables.some((table) => !table.verbatim && hasFloatingAttributes(table.floating))) {
    return xml;
  }

  const openTags = locateTableOpenTags(xml);
  if (openTags.length !== tables.length) return xml;

  const edits: DocumentXmlEdit[] = [];
  for (let index = 0; index < openTags.length; index += 1) {
    edits.push(...tableFloatingEdits(xml, openTags[index], tables[index]));
  }
  return applyDocumentXmlEdits(xml, edits);
}

/**
 * Patches one package part. Parts that hold no story (styles, numbering,
 * settings, media) are returned unchanged.
 */
export function applyTableFloatingToPartContent(
  partName: string,
  content: string,
  model: DocModel
): string {
  const name = partName.toLowerCase();
  if (name === DOCUMENT_PART) {
    return applyTableFloatingToStoryXml(content, model.nodes);
  }
  if (HEADER_FOOTER_PART.test(name)) {
    const nodes = sectionStoryNodes(model).get(name);
    return nodes ? applyTableFloatingToStoryXml(content, nodes) : content;
  }
  return content;
}

/** Body `word/document.xml` only, for `modelToDocumentXml`. */
export function applyTableFloatingToDocumentXml(
  documentXml: string,
  model: DocModel
): string {
  return applyTableFloatingToStoryXml(documentXml, model.nodes);
}

/**
 * Patches every story part of a parsed package. Used by `serializeDocModel`,
 * which returns a package instead of bytes.
 */
export function applyTableFloatingToPackage(
  pkg: OoxmlPackage,
  model: DocModel
): OoxmlPackage {
  if (!modelNeedsTableFloatingSerialization(model)) return pkg;
  let parts: Map<string, OoxmlPart> | undefined;
  for (const [partName, part] of pkg.parts) {
    const content = applyTableFloatingToPartContent(partName, part.content, model);
    if (content === part.content) continue;
    parts ??= new Map(pkg.parts);
    parts.set(partName, { name: part.name, content });
  }
  return parts ? { parts, binaryAssets: pkg.binaryAssets } : pkg;
}
