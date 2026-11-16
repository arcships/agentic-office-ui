// XML identifier hygiene for duplicated nodes.
//
// `DocNode.sourceXml` holds the original `<w:p>` / `<w:tbl>` markup and the
// exporter re-emits it verbatim, so anything stored in it survives a round-trip
// untouched — including identifiers that must stay document-unique:
//
//   * `w:bookmarkStart` / `w:bookmarkEnd` (bookmark name + id)
//   * `w14:paraId` / `w14:textId` on the paragraph mark
//   * `wp:docPr@id` of every inline/anchored drawing
//   * `w:id` of tracked-change containers
//   * `w:commentRangeStart` / `w:commentRangeEnd` / `w:commentReference`
//
// Copying a paragraph without touching them yields two nodes that share ids:
// Word reports the exported file as damaged, and bookmarks/comments bind to the
// wrong paragraph. `cloneNodeForInsertion()` clones a node and gives it a fresh,
// collision-free identity; markers that cannot be duplicated (bookmarks,
// comment ranges) are dropped from the copy instead of duplicated.

import type { DocModel, DocNode, HeaderSection, FooterSection } from "../../engine/types";
import { cloneDocNode } from "../../engine/clone";

/** First value handed out for a fresh `w14:paraId` (kept inside the valid range). */
const FIRST_PARA_ID = 0x10000000;

const BOOKMARK_MARKER_PATTERN = /<w:bookmark(?:Start|End)\b[^>]*?\/?>/gi;
const COMMENT_RANGE_MARKER_PATTERN = /<w:comment(?:RangeStart|RangeEnd|Reference)\b[^>]*?\/?>/gi;
const PARA_ID_ATTRIBUTE_PATTERN = /(\bw14:(?:paraId|textId)=")([^"]*)(")/g;
const DRAWING_ID_ATTRIBUTE_PATTERN = /(<wp:docPr\b[^>]*?\bid=")(\d+)(")/g;
const REVISION_ID_ATTRIBUTE_PATTERN =
  /(<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange)\b[^>]*?\bw:id=")(\d+)(")/g;

export interface SourceXmlIdentifierRegistry {
  /** Every `w14:paraId` / `w14:textId` value already used in the model (uppercase hex). */
  paraIds: Set<string>;
  /** Every `wp:docPr@id` already used in the model. */
  drawingIds: Set<number>;
  /** Every tracked-change `w:id` already used in the model. */
  revisionIds: Set<number>;
  nextParaId: number;
  nextDrawingId: number;
  nextRevisionId: number;
}

function formatParaId(value: number): string {
  return value.toString(16).toUpperCase().padStart(8, "0");
}

function collectIdentifiersFromSourceXml(
  sourceXml: string | undefined,
  registry: SourceXmlIdentifierRegistry
): void {
  if (!sourceXml) {
    return;
  }

  for (const match of sourceXml.matchAll(PARA_ID_ATTRIBUTE_PATTERN)) {
    registry.paraIds.add(match[2].toUpperCase());
  }

  for (const match of sourceXml.matchAll(DRAWING_ID_ATTRIBUTE_PATTERN)) {
    const id = Number.parseInt(match[2], 10);
    if (Number.isFinite(id)) {
      registry.drawingIds.add(id);
    }
  }

  for (const match of sourceXml.matchAll(REVISION_ID_ATTRIBUTE_PATTERN)) {
    const id = Number.parseInt(match[2], 10);
    if (Number.isFinite(id)) {
      registry.revisionIds.add(id);
    }
  }
}

function collectIdentifiersFromNode(
  node: DocNode,
  registry: SourceXmlIdentifierRegistry
): void {
  collectIdentifiersFromSourceXml(node.sourceXml, registry);

  if (node.type !== "table") {
    return;
  }

  for (const row of node.rows) {
    for (const cell of row.cells) {
      for (const child of cell.nodes) {
        collectIdentifiersFromNode(child, registry);
      }
    }
  }
}

function collectIdentifiersFromSections(
  sections: HeaderSection[] | FooterSection[],
  registry: SourceXmlIdentifierRegistry
): void {
  for (const section of sections) {
    for (const node of section.nodes) {
      collectIdentifiersFromNode(node, registry);
    }
  }
}

/**
 * Scan a model and record every XML identifier that a duplicate of one of its
 * nodes must not reuse. Build it once per insertion batch so several inserted
 * nodes cannot collide with each other either.
 */
export function createSourceXmlIdentifierRegistry(
  model: DocModel
): SourceXmlIdentifierRegistry {
  const registry: SourceXmlIdentifierRegistry = {
    paraIds: new Set<string>(),
    drawingIds: new Set<number>(),
    revisionIds: new Set<number>(),
    nextParaId: FIRST_PARA_ID,
    nextDrawingId: 1,
    nextRevisionId: 1,
  };

  for (const node of model.nodes) {
    collectIdentifiersFromNode(node, registry);
  }

  collectIdentifiersFromSections(model.metadata.headerSections ?? [], registry);
  collectIdentifiersFromSections(model.metadata.footerSections ?? [], registry);

  for (const section of model.metadata.sections ?? []) {
    collectIdentifiersFromSections(section.headerSections ?? [], registry);
    collectIdentifiersFromSections(section.footerSections ?? [], registry);
  }

  for (const note of [...(model.metadata.footnotes ?? []), ...(model.metadata.endnotes ?? [])]) {
    for (const node of note.nodes ?? []) {
      collectIdentifiersFromNode(node, registry);
    }
  }

  registry.nextDrawingId = highestIdentifier(registry.drawingIds) + 1;
  registry.nextRevisionId = highestIdentifier(registry.revisionIds) + 1;

  return registry;
}

function highestIdentifier(ids: Set<number>): number {
  let highest = 0;
  for (const id of ids) {
    if (id > highest) {
      highest = id;
    }
  }

  return highest;
}

function allocateParaId(registry: SourceXmlIdentifierRegistry): string {
  let candidate = formatParaId(registry.nextParaId);
  while (registry.paraIds.has(candidate)) {
    registry.nextParaId += 1;
    candidate = formatParaId(registry.nextParaId);
  }

  registry.nextParaId += 1;
  registry.paraIds.add(candidate);
  return candidate;
}

function allocateDrawingId(registry: SourceXmlIdentifierRegistry): number {
  const candidate = registry.nextDrawingId;
  registry.nextDrawingId += 1;
  registry.drawingIds.add(candidate);
  return candidate;
}

function allocateRevisionId(registry: SourceXmlIdentifierRegistry): number {
  const candidate = registry.nextRevisionId;
  registry.nextRevisionId += 1;
  registry.revisionIds.add(candidate);
  return candidate;
}

/**
 * Rewrite `sourceXml` so the node can be inserted next to the node it was
 * copied from: bookmarks and comment ranges are dropped, and paragraph ids,
 * drawing ids and tracked-change ids are renumbered when they are already taken.
 */
export function sanitizeDuplicatedSourceXml(
  sourceXml: string,
  registry: SourceXmlIdentifierRegistry
): string {
  let sanitized = sourceXml
    .replace(BOOKMARK_MARKER_PATTERN, "")
    .replace(COMMENT_RANGE_MARKER_PATTERN, "");

  sanitized = sanitized.replace(
    PARA_ID_ATTRIBUTE_PATTERN,
    (match: string, prefix: string, value: string, suffix: string) =>
      registry.paraIds.has(value.toUpperCase())
        ? `${prefix}${allocateParaId(registry)}${suffix}`
        : match
  );

  sanitized = sanitized.replace(
    DRAWING_ID_ATTRIBUTE_PATTERN,
    (match: string, prefix: string, idText: string, suffix: string) =>
      registry.drawingIds.has(Number.parseInt(idText, 10))
        ? `${prefix}${allocateDrawingId(registry)}${suffix}`
        : match
  );

  sanitized = sanitized.replace(
    REVISION_ID_ATTRIBUTE_PATTERN,
    (match: string, prefix: string, idText: string, suffix: string) =>
      registry.revisionIds.has(Number.parseInt(idText, 10))
        ? `${prefix}${allocateRevisionId(registry)}${suffix}`
        : match
  );

  return sanitized;
}

function sanitizeNodeSourceXml(
  node: DocNode,
  registry: SourceXmlIdentifierRegistry
): void {
  if (node.sourceXml) {
    node.sourceXml = sanitizeDuplicatedSourceXml(node.sourceXml, registry);
  }

  if (node.type !== "table") {
    return;
  }

  for (const row of node.rows) {
    for (const cell of row.cells) {
      for (const child of cell.nodes) {
        sanitizeNodeSourceXml(child, registry);
      }
    }
  }
}

/**
 * Clone a node for insertion into `registry`'s model: the clone keeps every
 * model field but no XML identifier is shared with the node it came from.
 */
export function cloneNodeForInsertion<T extends DocNode>(
  node: T,
  registry: SourceXmlIdentifierRegistry
): T {
  const clone = cloneDocNode(node) as T;
  sanitizeNodeSourceXml(clone, registry);
  return clone;
}
