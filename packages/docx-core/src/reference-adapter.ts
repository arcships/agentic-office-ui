import type {
  DocxObjectLocator,
  DocxPartLocator,
  DocxPathSegment,
  DocxReferenceKind,
  ManualRegionLocator,
  OfficeDocumentRevision,
  OfficeObjectDescriptor,
  OfficeObjectFingerprint,
  OfficeObjectKind,
  OfficeObjectReference,
  OfficeObjectReferenceDraft,
  OfficeObjectReliability,
  ResolveReferenceResult,
} from "@arcships/office-interaction";
import type { DocModel, DocNode, ImageRunNode, ParagraphNode, TableNode } from "./engine/types";
import type {
  DocxBlockPart,
  DocxBlockPathSegment,
  DocxWordRange,
  DocumentBlock,
  DocumentBlockDetails,
  DocumentWord,
} from "./blocks/types";
import type {
  DocxTextRange,
  DocxTextRangeBoundary,
  DocxTextRangeLocation,
} from "./editor/helpers/editor-types";
import { normalizeTextRange } from "./editor/helpers/editor-types";
import {
  getParagraphAtLocation,
  firstParagraphLocationInDocument,
  normalizeRangeBoundaryParagraphOffset,
  paragraphRangeForMutate,
  sameParagraphLocation,
} from "./editor/helpers/selection-helpers";
import { lastParagraphLocationInDocument } from "./editor/helpers/selection-helpers-range";
import { paragraphText } from "./editor/helpers/paragraph-inspect";

export type DocxOfficeReference = Extract<
  OfficeObjectReference,
  { document: { format: "docx" } }
>;

export type DocxOfficeReferenceDraft = Extract<
  OfficeObjectReferenceDraft,
  { document: { format: "docx" } }
>;

export interface DocxReferenceContext {
  revision: OfficeDocumentRevision & { format: "docx" };
  model: DocModel;
  /** Current rendered page count, when pagination is available. */
  pageCount?: number;
}

interface TextIndexEntry {
  location: DocxTextRangeLocation;
  paragraph: ParagraphNode;
  text: string;
  start: number;
  end: number;
}

interface TextIndex {
  text: string;
  entries: readonly TextIndexEntry[];
}

const TEXT_QUOTE_LIMIT = 256;
const TEXT_CONTEXT_LIMIT = 64;

const DOCX_TEXT_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["docx.native-text"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["docx.model-range"] },
  hierarchy: { level: "likely", reasonCodes: ["docx.model-path"] },
  relocation: { level: "likely", reasonCodes: ["docx.text-quote"] },
};

const DOCX_PAGE_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["docx.rendered-page"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["docx.page-boundary"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["docx.page-index"] },
  relocation: { level: "uncertain", reasonCodes: ["docx.reflow-sensitive"] },
};

const DOCX_REGION_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "unknown", reasonCodes: ["manual-region"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["normalized-region"] },
  hierarchy: { level: "likely", reasonCodes: ["docx.page-index"] },
  relocation: { level: "uncertain", reasonCodes: ["docx.reflow-sensitive"] },
};

const DOCX_IMAGE_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["docx.native-image"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["docx.model-child-index"] },
  hierarchy: { level: "likely", reasonCodes: ["docx.model-path"] },
  relocation: { level: "likely", reasonCodes: ["docx.structural-path"] },
};

const DOCX_TABLE_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["docx.native-table"] },
  boundary: { level: "likely", reasonCodes: ["docx.model-table-index"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["docx.model-path"] },
  relocation: { level: "likely", reasonCodes: ["docx.structural-path"] },
};

const DOCX_COMMENT_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["docx.native-comment"] },
  boundary: { level: "likely", reasonCodes: ["docx.comment-anchor"] },
  hierarchy: { level: "likely", reasonCodes: ["docx.model-path"] },
  relocation: { level: "uncertain", reasonCodes: ["docx.comment-anchor"] },
};

function copyReliability(value: OfficeObjectReliability): OfficeObjectReliability {
  return {
    semantic: { ...value.semantic, reasonCodes: [...value.semantic.reasonCodes] },
    boundary: { ...value.boundary, reasonCodes: [...value.boundary.reasonCodes] },
    hierarchy: { ...value.hierarchy, reasonCodes: [...value.hierarchy.reasonCodes] },
    relocation: { ...value.relocation, reasonCodes: [...value.relocation.reasonCodes] },
  };
}

function docxTextReliability(text: string): OfficeObjectReliability {
  if (text.length <= TEXT_QUOTE_LIMIT) return copyReliability(DOCX_TEXT_RELIABILITY);
  return {
    ...copyReliability(DOCX_TEXT_RELIABILITY),
    relocation: { level: "uncertain", reasonCodes: ["docx.text-hash-only"] },
  };
}

function textContentHash(text: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    hash ^= code & 0xff;
    hash = Math.imul(hash, 0x01000193);
    hash ^= code >>> 8;
    hash = Math.imul(hash, 0x01000193);
  }
  return `fnv1a32:${text.length}:${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

function hasMatchingTextFingerprint(text: string, fingerprint?: OfficeObjectFingerprint): boolean {
  if (!fingerprint) return false;
  if (fingerprint.exactText !== undefined && text !== fingerprint.exactText) return false;
  if (fingerprint.contentHash?.startsWith("fnv1a32:") && textContentHash(text) !== fingerprint.contentHash) {
    return false;
  }
  return fingerprint.exactText !== undefined || fingerprint.contentHash?.startsWith("fnv1a32:") === true;
}

function hasVerifiableTextFingerprint(fingerprint?: OfficeObjectFingerprint): boolean {
  return fingerprint?.exactText !== undefined || fingerprint?.contentHash?.startsWith("fnv1a32:") === true;
}

function assertDocxContext(context: DocxReferenceContext): void {
  if (context.revision.format !== "docx") throw new TypeError("DOCX reference context requires format=docx");
  if (!context.revision.documentId.trim()) throw new TypeError("DOCX documentId must not be empty");
  if (!context.revision.revision.trim()) throw new TypeError("DOCX revision must not be empty");
  if (context.pageCount !== undefined && (!Number.isSafeInteger(context.pageCount) || context.pageCount < 0)) {
    throw new TypeError("DOCX pageCount must be a non-negative safe integer");
  }
}

function sameRevision(left: OfficeDocumentRevision, right: OfficeDocumentRevision): boolean {
  return left.format === right.format
    && left.documentId === right.documentId
    && left.revision === right.revision;
}

function assertPageIndex(pageIndex: number, pageCount?: number): void {
  if (!Number.isSafeInteger(pageIndex) || pageIndex < 0) {
    throw new TypeError("DOCX pageIndex must be a non-negative safe integer");
  }
  if (pageCount !== undefined && pageIndex >= pageCount) {
    throw new RangeError(`DOCX pageIndex ${pageIndex} is outside the current ${pageCount}-page document`);
  }
}

function assertPageRegion(region: Extract<ManualRegionLocator, { space: "page" }>, pageCount?: number): void {
  assertPageIndex(region.pageIndex, pageCount);
  const { x, y, width, height } = region.rect;
  if (![x, y, width, height].every(Number.isFinite)) throw new TypeError("DOCX region coordinates must be finite");
  if (x < 0 || y < 0 || width <= 0 || height <= 0 || x + width > 1 || y + height > 1) {
    throw new RangeError("DOCX region must be a positive rectangle inside normalized page space");
  }
}

function pathFromLocation(location: DocxTextRangeLocation): readonly DocxPathSegment[] {
  if (location.kind === "paragraph") {
    return [{ kind: "paragraph", index: location.nodeIndex }];
  }
  return [
    { kind: "table", index: location.tableIndex },
    { kind: "row", index: location.rowIndex },
    { kind: "cell", index: location.cellIndex },
    { kind: "paragraph", index: location.paragraphIndex },
  ];
}

function locationFromPath(path: readonly DocxPathSegment[]): DocxTextRangeLocation | undefined {
  if (path.length === 1 && (path[0]?.kind === "paragraph" || path[0]?.kind === "node")) {
    return { kind: "paragraph", nodeIndex: path[0].index };
  }
  const table = path.find((segment) => segment.kind === "table" || segment.kind === "node");
  const row = path.find((segment) => segment.kind === "row");
  const cell = path.find((segment) => segment.kind === "cell");
  const paragraph = path.find((segment) => segment.kind === "paragraph");
  if (!table || !row || !cell || !paragraph) return undefined;
  return {
    kind: "table-cell",
    tableIndex: table.index,
    rowIndex: row.index,
    cellIndex: cell.index,
    paragraphIndex: paragraph.index,
  };
}

function locatorFromRange(range: DocxTextRange): DocxObjectLocator {
  const normalized = normalizeTextRange(range);
  return {
    kind: "text-range",
    part: { kind: "body" },
    start: { path: pathFromLocation(normalized.start.location), offset: normalized.start.offset },
    end: { path: pathFromLocation(normalized.end.location), offset: normalized.end.offset },
  };
}

function rangeFromLocator(locator: DocxObjectLocator): DocxTextRange | undefined {
  if (locator.kind !== "text-range" || locator.part.kind !== "body") return undefined;
  const start = locationFromPath(locator.start.path);
  const end = locationFromPath(locator.end.path);
  if (!start || !end) return undefined;
  return normalizeTextRange({
    start: { location: start, offset: locator.start.offset },
    end: { location: end, offset: locator.end.offset },
  });
}

function buildTextIndex(model: DocModel): TextIndex {
  const first = firstParagraphLocationInDocument(model);
  const last = lastParagraphLocationInDocument(model);
  if (!first || !last) return { text: "", entries: [] };
  const locations = paragraphRangeForMutate(model, first, last);
  const entries: TextIndexEntry[] = [];
  let text = "";
  for (const item of locations) {
    const paragraph = getParagraphAtLocation(model, item.location).paragraph;
    if (!paragraph) continue;
    if (entries.length > 0) text += "\n";
    const start = text.length;
    const value = paragraphText(paragraph);
    text += value;
    entries.push({
      location: item.location,
      paragraph,
      text: value,
      start,
      end: text.length,
    });
  }
  return { text, entries };
}

function globalOffsetForBoundary(index: TextIndex, boundary: DocxTextRangeBoundary): number | undefined {
  const entry = index.entries.find((candidate) => sameParagraphLocation(candidate.location, boundary.location));
  if (!entry) return undefined;
  return entry.start + normalizeRangeBoundaryParagraphOffset(entry.paragraph, boundary.offset);
}

function boundaryForGlobalOffset(
  index: TextIndex,
  offset: number,
  edge: "start" | "end",
): DocxTextRangeBoundary | undefined {
  if (index.entries.length === 0 || offset < 0 || offset > index.text.length) return undefined;
  for (let entryIndex = 0; entryIndex < index.entries.length; entryIndex += 1) {
    const entry = index.entries[entryIndex]!;
    if (offset < entry.end || offset === entry.end && (edge === "end" || entryIndex === index.entries.length - 1)) {
      return { location: entry.location, offset: Math.max(0, Math.min(entry.text.length, offset - entry.start)) };
    }
    if (offset === entry.end + 1 && edge === "start") {
      const next = index.entries[entryIndex + 1];
      if (next) return { location: next.location, offset: 0 };
    }
  }
  const last = index.entries[index.entries.length - 1]!;
  return { location: last.location, offset: last.text.length };
}

function indexedRange(index: TextIndex, range: DocxTextRange): {
  range: DocxTextRange;
  text: string;
  start: number;
  end: number;
} | undefined {
  const normalized = normalizeTextRange(range);
  const start = globalOffsetForBoundary(index, normalized.start);
  const end = globalOffsetForBoundary(index, normalized.end);
  if (start === undefined || end === undefined || end <= start) return undefined;
  return { range: normalized, text: index.text.slice(start, end), start, end };
}

function fingerprintForIndexedText(index: TextIndex, start: number, end: number): OfficeObjectFingerprint {
  const selectedText = index.text.slice(start, end);
  const prefixText = index.text.slice(Math.max(0, start - TEXT_CONTEXT_LIMIT), start);
  const suffixText = index.text.slice(end, Math.min(index.text.length, end + TEXT_CONTEXT_LIMIT));
  return {
    ...(selectedText.length <= TEXT_QUOTE_LIMIT ? { exactText: selectedText } : {}),
    ...(prefixText ? { prefixText } : {}),
    ...(suffixText ? { suffixText } : {}),
    contentHash: textContentHash(selectedText),
  };
}

/** Return the canonical model text represented by a DOCX editor range. */
export function docxTextRangeText(model: DocModel, range: DocxTextRange): string | undefined {
  return indexedRange(buildTextIndex(model), range)?.text;
}

export function createDocxTextReferenceDraft(
  context: DocxReferenceContext,
  range: DocxTextRange,
  options: { fallbackRegion?: Extract<ManualRegionLocator, { space: "page" }> } = {},
): DocxOfficeReferenceDraft {
  assertDocxContext(context);
  if (options.fallbackRegion) assertPageRegion(options.fallbackRegion, context.pageCount);
  const index = buildTextIndex(context.model);
  const selected = indexedRange(index, range);
  if (!selected?.text) throw new RangeError("DOCX text reference requires a non-empty range in the current model");
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "text-range",
    source: "native",
    locator: { type: "format", format: "docx", value: locatorFromRange(selected.range) },
    fingerprint: fingerprintForIndexedText(index, selected.start, selected.end),
    ...(options.fallbackRegion ? { fallbackRegion: options.fallbackRegion } : {}),
    reliability: docxTextReliability(selected.text),
  };
}

export function createDocxPageReferenceDraft(
  context: DocxReferenceContext,
  pageIndex: number,
): DocxOfficeReferenceDraft {
  assertDocxContext(context);
  assertPageIndex(pageIndex, context.pageCount);
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "page",
    source: "native",
    locator: { type: "format", format: "docx", value: { kind: "page", pageIndex } },
    reliability: copyReliability(DOCX_PAGE_RELIABILITY),
  };
}

export function createDocxRegionReferenceDraft(
  context: DocxReferenceContext,
  region: Extract<ManualRegionLocator, { space: "page" }>,
): DocxOfficeReferenceDraft {
  assertDocxContext(context);
  assertPageRegion(region, context.pageCount);
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "region",
    source: "manual",
    locator: { type: "manual-region", format: "docx", value: region },
    fallbackRegion: region,
    reliability: copyReliability(DOCX_REGION_RELIABILITY),
  };
}

export interface DocxImageReferenceTarget {
  paragraph: DocxTextRangeLocation;
  childIndex: number;
}

function draftParagraph(
  context: DocxReferenceContext,
  location: DocxTextRangeLocation,
): ParagraphNode | undefined {
  return getParagraphAtLocation(context.model, location).paragraph;
}

export function createDocxImageReferenceDraft(
  context: DocxReferenceContext,
  target: DocxImageReferenceTarget,
): DocxOfficeReferenceDraft {
  assertDocxContext(context);
  if (!Number.isSafeInteger(target.childIndex) || target.childIndex < 0) {
    throw new TypeError("DOCX image childIndex must be a non-negative safe integer");
  }
  const paragraph = draftParagraph(context, target.paragraph);
  const child = paragraph?.children[target.childIndex];
  if (!child || child.type !== "image") {
    throw new RangeError("DOCX image reference requires an image child in the current model");
  }
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "image",
    source: "native",
    locator: {
      type: "format",
      format: "docx",
      value: {
        kind: "image",
        part: { kind: "body" },
        paragraphPath: pathFromLocation(target.paragraph),
        childIndex: target.childIndex,
      },
    },
    reliability: copyReliability(DOCX_IMAGE_RELIABILITY),
  };
}

export function createDocxTableReferenceDraft(
  context: DocxReferenceContext,
  tableIndex: number,
): DocxOfficeReferenceDraft {
  assertDocxContext(context);
  if (!Number.isSafeInteger(tableIndex) || tableIndex < 0) {
    throw new TypeError("DOCX tableIndex must be a non-negative safe integer");
  }
  if (context.model.nodes[tableIndex]?.type !== "table") {
    throw new RangeError(`DOCX table ${tableIndex} is not present in the current model`);
  }
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "table",
    source: "native",
    locator: {
      type: "format",
      format: "docx",
      value: {
        kind: "structure",
        part: { kind: "body" },
        path: [{ kind: "table", index: tableIndex }],
      },
    },
    reliability: copyReliability(DOCX_TABLE_RELIABILITY),
  };
}

export interface DocxCommentReferenceTarget {
  commentId: string;
  anchor?: DocxTextRangeLocation;
}

export function createDocxCommentReferenceDraft(
  context: DocxReferenceContext,
  target: DocxCommentReferenceTarget,
): DocxOfficeReferenceDraft {
  assertDocxContext(context);
  const commentId = target.commentId.trim();
  if (!commentId) throw new TypeError("DOCX commentId must not be empty");
  if (!context.model.metadata.comments?.some((comment) => String(comment.id) === commentId)) {
    throw new RangeError(`DOCX comment ${commentId} is not present in the current model`);
  }
  if (target.anchor && !draftParagraph(context, target.anchor)) {
    throw new RangeError("DOCX comment anchor is not present in the current model");
  }
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "comment",
    source: "native",
    locator: {
      type: "format",
      format: "docx",
      value: {
        kind: "comment",
        part: { kind: "body" },
        commentId,
        ...(target.anchor ? { anchorPath: pathFromLocation(target.anchor) } : {}),
      },
    },
    reliability: copyReliability(DOCX_COMMENT_RELIABILITY),
  };
}

// ────────────────────────────────────────────────────────────────────
// Document block / word references
// ────────────────────────────────────────────────────────────────────

const DOCX_BLOCK_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["docx.block-identity"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["docx.block-source-ref"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["docx.model-path"] },
  relocation: { level: "likely", reasonCodes: ["docx.structural-path"] },
};

const DOCX_BLOCK_CONTAINER_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "likely", score: 1, reasonCodes: ["docx.story-container"] },
  boundary: { level: "likely", score: 1, reasonCodes: ["docx.story-part"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["docx.part-locator"] },
  relocation: { level: "uncertain", reasonCodes: ["docx.story-part"] },
};

function blockPart(block: DocumentBlock): DocxBlockPart {
  const part = block.details.part;
  if (!part) return { kind: "body" };
  return part;
}

function blockModelPath(block: DocumentBlock): readonly DocxPathSegment[] {
  const path = block.details.modelPath as readonly DocxBlockPathSegment[] | undefined;
  return (path ?? []) as readonly DocxPathSegment[];
}

function blockReferenceKind(block: DocumentBlock): DocxReferenceKind {
  if (block.type === "table") return "table";
  if (block.type === "table_head") return "table-row";
  if (block.type === "table_cell") return "table-cell";
  if (block.type === "figure") return "image";
  if (block.type === "header" || block.type === "footer") return "section";
  if (block.type === "heading" || block.type === "section_heading") return "heading";
  if (block.type === "legend") return "paragraph";
  if (block.type === "page_number") return "paragraph";
  if (block.details.list) return "list-item";
  return "paragraph";
}

function textFingerprintAt(
  text: string,
  start: number,
  end: number,
): OfficeObjectFingerprint {
  const selected = text.slice(start, end);
  const prefix = text.slice(Math.max(0, start - TEXT_CONTEXT_LIMIT), start);
  const suffix = text.slice(end, Math.min(text.length, end + TEXT_CONTEXT_LIMIT));
  return {
    ...(selected.length <= TEXT_QUOTE_LIMIT ? { exactText: selected } : {}),
    ...(prefix ? { prefixText: prefix } : {}),
    ...(suffix ? { suffixText: suffix } : {}),
    contentHash: textContentHash(selected),
  };
}

function blockStructureLocator(block: DocumentBlock): DocxObjectLocator {
  const path = blockModelPath(block);
  return {
    kind: "structure",
    part: blockPart(block) as DocxPartLocator,
    path: path.length > 0 ? path : [{ kind: "node", index: 0 }],
  };
}

function blockImageLocator(block: DocumentBlock): DocxObjectLocator | undefined {
  const childIndex = block.details.childIndex;
  if (!Number.isSafeInteger(childIndex) || (childIndex as number) < 0) return undefined;
  return {
    kind: "image",
    part: blockPart(block) as DocxPartLocator,
    paragraphPath: blockModelPath(block),
    childIndex: childIndex as number,
  };
}

/**
 * Creates a unified reference draft for an extracted document block.
 *
 * Paragraph-derived blocks keep their semantic kind (paragraph/heading/
 * list-item) with a structural model path, tables/cells/rows use structural
 * paths, figures use the image locator and header/footer containers use their
 * part locator. The draft can be passed to `confirmOfficeReferenceDraft` so
 * `referenceConfirm` carries block identity.
 */
export function createDocxBlockReferenceDraft(
  context: DocxReferenceContext,
  block: DocumentBlock,
  options: { fallbackRegion?: Extract<ManualRegionLocator, { space: "page" }> } = {},
): DocxOfficeReferenceDraft {
  assertDocxContext(context);
  if (options.fallbackRegion) assertPageRegion(options.fallbackRegion, context.pageCount);
  if (!block.sourceRefs.length) {
    throw new TypeError("DOCX block reference requires at least one source reference");
  }

  let locator: DocxObjectLocator;
  if (block.type === "figure") {
    const image = blockImageLocator(block);
    if (!image) throw new RangeError("DOCX figure block has no image child index");
    locator = image;
  } else {
    // Paragraph-derived blocks keep their semantic kind (paragraph/heading/
    // list-item); the unified contract only pairs `text-range` locators with
    // the `text-range` kind, so a structural path is the correct locator here.
    locator = blockStructureLocator(block);
  }

  const kind = blockReferenceKind(block);
  const container = block.type === "header" || block.type === "footer";
  return {
    schemaVersion: 1,
    document: context.revision,
    kind,
    source: "native",
    locator: { type: "format", format: "docx", value: locator },
    ...(options.fallbackRegion ? { fallbackRegion: options.fallbackRegion } : {}),
    reliability: copyReliability(
      container ? DOCX_BLOCK_CONTAINER_RELIABILITY : DOCX_BLOCK_RELIABILITY,
    ),
  } as DocxOfficeReferenceDraft;
}

function wordRangeOffsets(
  word: DocumentWord | DocxWordRange,
): { start: number; end: number; content?: string } | undefined {
  if ("start" in word && typeof word.start === "number") {
    return {
      start: word.start,
      end: word.end,
      ...(word.content ? { content: word.content } : {}),
    };
  }
  if ("sourceRefs" in word) {
    const ref = word.sourceRefs[0];
    if (!ref || ref.start === undefined || ref.end === undefined) return undefined;
    return { start: ref.start, end: ref.end, content: word.content };
  }
  return undefined;
}

/**
 * Creates a unified reference draft for a word span inside a document block.
 * The locator is a text-range with UTF-16 offsets, so `referenceConfirm` can
 * carry word-level citations and resolution can relocate them by quote.
 */
export function createDocxWordReferenceDraft(
  context: DocxReferenceContext,
  block: DocumentBlock,
  word: DocumentWord | DocxWordRange,
  options: { fallbackRegion?: Extract<ManualRegionLocator, { space: "page" }> } = {},
): DocxOfficeReferenceDraft {
  assertDocxContext(context);
  if (options.fallbackRegion) assertPageRegion(options.fallbackRegion, context.pageCount);
  if (block.details.kind !== "paragraph") {
    throw new TypeError("DOCX word reference requires a paragraph-derived block");
  }
  const offsets = wordRangeOffsets(word);
  if (!offsets) {
    throw new TypeError("DOCX word reference requires UTF-16 start/end offsets");
  }
  const { start, end, content: expected } = offsets;
  if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end <= start) {
    throw new RangeError("DOCX word reference offsets must describe a non-empty integer range");
  }
  if (end > block.content.length) {
    throw new RangeError("DOCX word reference offsets exceed the block text");
  }
  const selected = block.content.slice(start, end);
  if (expected !== undefined && expected !== selected) {
    throw new RangeError("DOCX word reference text does not match the block text at those offsets");
  }
  const path = blockModelPath(block);
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "text-range",
    source: "native",
    locator: {
      type: "format",
      format: "docx",
      value: {
        kind: "text-range",
        part: blockPart(block) as DocxPartLocator,
        start: { path, offset: start },
        end: { path, offset: end },
      },
    },
    fingerprint: textFingerprintAt(block.content, start, end),
    ...(options.fallbackRegion ? { fallbackRegion: options.fallbackRegion } : {}),
    reliability: docxTextReliability(selected),
  } as DocxOfficeReferenceDraft;
}

function draftFromReference(reference: DocxOfficeReference): DocxOfficeReferenceDraft {
  const { referenceId: _referenceId, ...draft } = reference;
  return draft;
}

function locationPath(location: DocxTextRangeLocation): OfficeObjectDescriptor["path"] {
  if (location.kind === "paragraph") {
    return [
      { kind: "document", label: "Document" },
      { kind: "paragraph", label: `Paragraph ${location.nodeIndex + 1}` },
    ];
  }
  return [
    { kind: "document", label: "Document" },
    { kind: "table", label: `Table at node ${location.tableIndex + 1}` },
    { kind: "table-row", label: `Row ${location.rowIndex + 1}` },
    { kind: "table-cell", label: `Cell ${location.cellIndex + 1}` },
    { kind: "paragraph", label: `Paragraph ${location.paragraphIndex + 1}` },
  ];
}

function textDescriptor(reference: DocxOfficeReference, range: DocxTextRange, text: string): OfficeObjectDescriptor {
  const visual = reference.fallbackRegion
    ? {
        fragments: [{
          container: { space: "page" as const, pageIndex: reference.fallbackRegion.pageIndex },
          rect: reference.fallbackRegion.rect,
        }],
        layoutVersion: reference.document.revision,
      }
    : undefined;
  const labelText = text.replace(/\s+/gu, " ").trim();
  return {
    objectId: `docx:text:${JSON.stringify((reference.locator as { value: DocxObjectLocator }).value)}`,
    draft: draftFromReference(reference),
    label: labelText ? `Text: ${labelText.slice(0, 80)}` : "Text range",
    path: locationPath(range.start.location),
    childrenState: "none",
    content: { text },
    ...(visual ? { visual } : {}),
  };
}

function simpleDescriptor(reference: DocxOfficeReference): OfficeObjectDescriptor {
  if (reference.locator.type === "manual-region") {
    const region = reference.locator.value;
    return {
      objectId: `docx:region:${region.pageIndex}:${JSON.stringify(region.rect)}`,
      draft: draftFromReference(reference),
      label: `Region on page ${region.pageIndex + 1}`,
      path: [
        { kind: "document", label: "Document" },
        { kind: "page", label: `Page ${region.pageIndex + 1}` },
        { kind: "region", label: "Selected region" },
      ],
      childrenState: "none",
      visual: {
        fragments: [{ container: { space: "page", pageIndex: region.pageIndex }, rect: region.rect }],
        layoutVersion: reference.document.revision,
      },
    };
  }
  const locator = reference.locator.value;
  const pageIndex = locator.kind === "page" ? locator.pageIndex : 0;
  return {
    objectId: `docx:page:${pageIndex}`,
    draft: draftFromReference(reference),
    label: `Page ${pageIndex + 1}`,
    path: [
      { kind: "document", label: "Document" },
      { kind: "page", label: `Page ${pageIndex + 1}` },
    ],
    childrenState: "unknown",
  };
}

function visualForFallbackRegion(reference: DocxOfficeReference): OfficeObjectDescriptor["visual"] {
  const region = reference.fallbackRegion;
  return region
    ? {
        fragments: [{ container: { space: "page", pageIndex: region.pageIndex }, rect: region.rect }],
        layoutVersion: reference.document.revision,
      }
    : undefined;
}

function pathKey(path: readonly DocxPathSegment[]): string {
  return path.map((segment) => `${segment.kind}:${segment.index}`).join("/");
}

function imageDescriptor(reference: DocxOfficeReference, image: ImageRunNode): OfficeObjectDescriptor {
  const locator = reference.locator.type === "format" && reference.locator.value.kind === "image"
    ? reference.locator.value
    : undefined;
  const paragraphLocation = locator ? locationFromPath(locator.paragraphPath) : undefined;
  const label = image.alt?.replace(/\s+/gu, " ").trim() || "Image";
  const visual = visualForFallbackRegion(reference);
  return {
    objectId: `docx:image:${locator ? pathKey(locator.paragraphPath) : "?"}:${locator?.childIndex ?? -1}`,
    draft: draftFromReference(reference),
    label: label === "Image" ? "Image" : `Image: ${label.slice(0, 80)}`,
    path: [
      { kind: "document", label: "Document" },
      ...(paragraphLocation ? locationPath(paragraphLocation).slice(1) : []),
      { kind: "image", label },
    ],
    childrenState: "none",
    ...(image.alt ? { content: { text: image.alt } } : {}),
    ...(visual ? { visual } : {}),
  };
}

function commentDescriptor(
  reference: DocxOfficeReference,
  commentId: string,
  comment: { author?: string; text: string },
): OfficeObjectDescriptor {
  const text = comment.text.replace(/\s+/gu, " ").trim();
  const anchorLocation = reference.locator.type === "format" && reference.locator.value.kind === "comment"
    ? locationFromPath(reference.locator.value.anchorPath ?? [])
    : undefined;
  return {
    objectId: `docx:comment:${commentId}`,
    draft: draftFromReference(reference),
    label: text ? `Comment: ${text.slice(0, 80)}` : `Comment ${commentId}`,
    path: [
      { kind: "document", label: "Document" },
      ...(anchorLocation ? locationPath(anchorLocation).slice(1) : []),
      { kind: "comment", label: comment.author ? `Comment by ${comment.author}` : "Comment" },
    ],
    childrenState: "none",
    content: { text: comment.text },
  };
}

function currentReference(
  context: DocxReferenceContext,
  reference: DocxOfficeReference,
  locator: DocxObjectLocator,
  fallbackRegion?: Extract<ManualRegionLocator, { space: "page" }>,
  fingerprint: OfficeObjectFingerprint | undefined = reference.fingerprint,
): DocxOfficeReference {
  const { fallbackRegion: _oldFallback, ...base } = reference;
  return {
    ...base,
    document: context.revision,
    locator: { type: "format", format: "docx", value: locator },
    ...(fingerprint ? { fingerprint } : {}),
    ...(fallbackRegion ? { fallbackRegion } : {}),
  } as DocxOfficeReference;
}

function textCandidates(
  context: DocxReferenceContext,
  reference: DocxOfficeReference,
  index: TextIndex,
): readonly { reference: DocxOfficeReference; range: DocxTextRange; text: string; score: number }[] {
  const quote = reference.fingerprint?.exactText;
  if (!quote) return [];
  const candidates: { reference: DocxOfficeReference; range: DocxTextRange; text: string; score: number }[] = [];
  let from = 0;
  while (from <= index.text.length - quote.length) {
    const start = index.text.indexOf(quote, from);
    if (start < 0) break;
    const end = start + quote.length;
    const startBoundary = boundaryForGlobalOffset(index, start, "start");
    const endBoundary = boundaryForGlobalOffset(index, end, "end");
    if (startBoundary && endBoundary) {
      const prefixMatches = !reference.fingerprint?.prefixText
        || index.text.slice(0, start).endsWith(reference.fingerprint.prefixText);
      const suffixMatches = !reference.fingerprint?.suffixText
        || index.text.slice(end).startsWith(reference.fingerprint.suffixText);
      const range = normalizeTextRange({ start: startBoundary, end: endBoundary });
      const locator = locatorFromRange(range);
      candidates.push({
        reference: currentReference(
          context,
          reference,
          locator,
          undefined,
          fingerprintForIndexedText(index, start, end),
        ),
        range,
        text: quote,
        score: 4 + (prefixMatches ? 2 : 0) + (suffixMatches ? 2 : 0),
      });
    }
    from = start + 1;
  }
  const topScore = Math.max(0, ...candidates.map((candidate) => candidate.score));
  return candidates.filter((candidate) => candidate.score === topScore);
}

function docxReferenceKind(
  reference: DocxOfficeReference,
): "text" | "page" | "region" | "image" | "structure" | "comment" | "other" {
  if (reference.locator.type === "manual-region") return "region";
  const value = reference.locator.value;
  if (value.kind === "text-range") return "text";
  if (value.kind === "page") return "page";
  if (value.kind === "image") return "image";
  if (value.kind === "comment") return "comment";
  if (value.kind === "structure") return "structure";
  return "other";
}

function resolveDocxImageReference(
  context: DocxReferenceContext,
  reference: DocxOfficeReference,
): ResolveReferenceResult {
  if (reference.locator.type !== "format" || reference.locator.value.kind !== "image") {
    return { status: "unsupported", reasonCode: "docx.image-locator-unsupported" };
  }
  const locator = reference.locator.value;
  const paragraphLocation = locationFromPath(locator.paragraphPath);
  const image = paragraphLocation
    ? getParagraphAtLocation(context.model, paragraphLocation).paragraph?.children[locator.childIndex]
    : undefined;
  if (!image || image.type !== "image") {
    return { status: "not-found", reasonCode: "docx.image-not-found" };
  }
  if (sameRevision(reference.document, context.revision)) {
    return { status: "exact", reference, descriptor: imageDescriptor(reference, image) };
  }
  const relocated = currentReference(context, reference, locator);
  return {
    status: "relocated",
    reference: relocated,
    descriptor: imageDescriptor(relocated, image),
    reasonCodes: ["docx.revision-changed", "docx.locator-stable"],
  };
}

function resolveDocxCommentReference(
  context: DocxReferenceContext,
  reference: DocxOfficeReference,
): ResolveReferenceResult {
  if (reference.locator.type !== "format" || reference.locator.value.kind !== "comment") {
    return { status: "unsupported", reasonCode: "docx.comment-locator-unsupported" };
  }
  const locator = reference.locator.value;
  const comment = context.model.metadata.comments?.find(
    (entry) => String(entry.id) === locator.commentId,
  );
  if (!comment) {
    return { status: "not-found", reasonCode: "docx.comment-not-found" };
  }
  if (sameRevision(reference.document, context.revision)) {
    return { status: "exact", reference, descriptor: commentDescriptor(reference, locator.commentId, comment) };
  }
  const relocated = currentReference(context, reference, locator);
  return {
    status: "relocated",
    reference: relocated,
    descriptor: commentDescriptor(relocated, locator.commentId, comment),
    reasonCodes: ["docx.revision-changed", "docx.locator-stable"],
  };
}

type ResolvedBlockPath =
  | { kind: "paragraph"; paragraph: ParagraphNode }
  | { kind: "table"; table: TableNode }
  | { kind: "table-row"; table: TableNode; rowIndex: number }
  | { kind: "table-cell"; table: TableNode; rowIndex: number; cellIndex: number; cell: TableNode["rows"][number]["cells"][number] }
  | { kind: "container"; nodes: readonly DocNode[] };

function findStoryNodes(
  model: DocModel,
  part: DocxPartLocator,
): readonly DocNode[] | undefined {
  if (part.kind === "body") return model.nodes;
  if (part.kind === "header" || part.kind === "footer") {
    const flat = part.kind === "header"
      ? model.metadata.headerSections
      : model.metadata.footerSections;
    const direct = flat?.find((section) => section.partName === part.partName);
    if (direct) return direct.nodes;
    for (const section of model.metadata.sections ?? []) {
      const nested = (part.kind === "header"
        ? section.headerSections
        : section.footerSections)?.find((entry) => entry.partName === part.partName);
      if (nested) return nested.nodes;
    }
    return undefined;
  }
  if (part.kind === "footnote" || part.kind === "endnote") {
    const notes = part.kind === "footnote"
      ? model.metadata.footnotes
      : model.metadata.endnotes;
    return notes?.find((note) => String(note.id) === part.noteId)?.nodes;
  }
  return undefined;
}

function resolveTablePath(
  table: TableNode,
  path: readonly DocxPathSegment[],
  tableIndex: number,
): ResolvedBlockPath | undefined {
  if (tableIndex === path.length - 1) return { kind: "table", table };
  const rowSegment = path[tableIndex + 1];
  if (!rowSegment || rowSegment.kind !== "row") return undefined;
  const row = table.rows[rowSegment.index];
  if (!row) return undefined;
  if (tableIndex + 1 === path.length - 1) {
    return { kind: "table-row", table, rowIndex: rowSegment.index };
  }
  const cellSegment = path[tableIndex + 2];
  if (!cellSegment || cellSegment.kind !== "cell") return undefined;
  const cell = row.cells[cellSegment.index];
  if (!cell) return undefined;
  if (tableIndex + 2 === path.length - 1) {
    return {
      kind: "table-cell",
      table,
      rowIndex: rowSegment.index,
      cellIndex: cellSegment.index,
      cell,
    };
  }
  const child = path[tableIndex + 3];
  if (!child) return undefined;
  if (child.kind === "paragraph") {
    if (tableIndex + 3 !== path.length - 1) return undefined;
    const paragraph = cell.nodes[child.index];
    return paragraph?.type === "paragraph"
      ? { kind: "paragraph", paragraph }
      : undefined;
  }
  if (child.kind === "table") {
    const nested = cell.nodes[child.index];
    if (nested?.type !== "table") return undefined;
    return resolveTablePath(nested, path, tableIndex + 3);
  }
  return undefined;
}

function resolvePathInNodes(
  nodes: readonly DocNode[],
  path: readonly DocxPathSegment[],
): ResolvedBlockPath | undefined {
  const first = path[0];
  if (!first) return { kind: "container", nodes };
  if (first.kind === "paragraph" || first.kind === "node") {
    if (path.length !== 1) return undefined;
    const node = nodes[first.index];
    return node?.type === "paragraph"
      ? { kind: "paragraph", paragraph: node }
      : undefined;
  }
  if (first.kind === "table") {
    const node = nodes[first.index];
    if (node?.type !== "table") return undefined;
    return resolveTablePath(node, path, 0);
  }
  return undefined;
}

function partDescriptorPath(
  part: DocxPartLocator,
): { kind: OfficeObjectKind; label: string } | undefined {
  if (part.kind === "body") return undefined;
  if (part.kind === "header") {
    return { kind: "section", label: `Header ${part.partName}` };
  }
  if (part.kind === "footer") {
    return { kind: "section", label: `Footer ${part.partName}` };
  }
  if (part.kind === "footnote" || part.kind === "endnote") {
    return {
      kind: "note",
      label: `${part.kind === "footnote" ? "Footnote" : "Endnote"} ${part.noteId}`,
    };
  }
  return undefined;
}

function blockDescriptor(
  reference: DocxOfficeReference,
  part: DocxPartLocator,
  resolved: ResolvedBlockPath,
  text: string,
): OfficeObjectDescriptor {
  const partPath = partDescriptorPath(part);
  const basePath: Array<{ kind: OfficeObjectKind; label: string }> = [
    { kind: "document", label: "Document" },
    ...(partPath ? [partPath] : []),
  ];
  const common = {
    objectId: `docx:block:${JSON.stringify(reference.locator)}`,
    draft: draftFromReference(reference),
    ...(text.trim() ? { content: { text } } : {}),
    ...(reference.fallbackRegion
      ? {
          visual: {
            fragments: [{
              container: { space: "page" as const, pageIndex: reference.fallbackRegion.pageIndex },
              rect: reference.fallbackRegion.rect,
            }],
            layoutVersion: reference.document.revision,
          },
        }
      : {}),
  };
  if (resolved.kind === "container") {
    return {
      ...common,
      label: partPath?.label ?? "Document part",
      path: [...basePath, { kind: "section", label: "Part container" }],
      childrenState: "available",
    };
  }
  if (resolved.kind === "table") {
    const tableIndex =
      reference.locator.type === "format" &&
      reference.locator.value.kind === "structure" &&
      reference.locator.value.path[0]?.kind === "table"
        ? reference.locator.value.path[0].index
        : undefined;
    return {
      ...common,
      label: tableIndex === undefined ? "Table" : `Table ${tableIndex + 1}`,
      path: [...basePath, { kind: "table", label: "Table" }],
      childrenState: "available",
    };
  }
  if (resolved.kind === "table-row") {
    return {
      ...common,
      label: `Row ${resolved.rowIndex + 1}`,
      path: [
        ...basePath,
        { kind: "table", label: "Table" },
        { kind: "table-row", label: `Row ${resolved.rowIndex + 1}` },
      ],
      childrenState: "available",
    };
  }
  if (resolved.kind === "table-cell") {
    return {
      ...common,
      label: `Cell ${resolved.rowIndex + 1}, ${resolved.cellIndex + 1}`,
      path: [
        ...basePath,
        { kind: "table", label: "Table" },
        { kind: "table-row", label: `Row ${resolved.rowIndex + 1}` },
        { kind: "table-cell", label: `Cell ${resolved.cellIndex + 1}` },
      ],
      childrenState: "available",
    };
  }
  const labelText = text.replace(/\s+/gu, " ").trim();
  const heading = reference.kind === "heading";
  const listItem = reference.kind === "list-item";
  return {
    ...common,
    label: labelText
      ? `${heading ? "Heading" : listItem ? "List item" : "Text"}: ${labelText.slice(0, 80)}`
      : heading
        ? "Heading"
        : listItem
          ? "List item"
          : "Paragraph",
    path: [
      ...basePath,
      { kind: "paragraph", label: heading ? "Heading" : listItem ? "List item" : "Paragraph" },
    ],
    childrenState: "none",
  };
}

function resolveDocxStructureReference(
  context: DocxReferenceContext,
  reference: DocxOfficeReference,
): ResolveReferenceResult {
  if (reference.locator.type !== "format" || reference.locator.value.kind !== "structure") {
    return { status: "unsupported", reasonCode: "docx.structure-locator-unsupported" };
  }
  const locator = reference.locator.value;
  const nodes = findStoryNodes(context.model, locator.part);
  if (!nodes) return { status: "not-found", reasonCode: "docx.part-not-found" };
  const container =
    reference.kind === "section" &&
    locator.part.kind !== "body" &&
    locator.path.length === 1 &&
    locator.path[0]?.kind === "node" &&
    locator.path[0]?.index === 0;
  const resolved: ResolvedBlockPath | undefined = container
    ? { kind: "container", nodes }
    : resolvePathInNodes(nodes, locator.path);
  if (!resolved) {
    return {
      status: "not-found",
      reasonCode: locator.path[0]?.kind === "table" && locator.path.length === 1
        ? "docx.table-not-found"
        : "docx.structure-not-found",
    };
  }
  const text = resolved.kind === "paragraph"
    ? paragraphText(resolved.paragraph)
    : resolved.kind === "table-cell"
      ? resolved.cell.nodes
          .map((node) => (node.type === "paragraph" ? paragraphText(node) : ""))
          .filter(Boolean)
          .join("\n")
      : "";
  if (sameRevision(reference.document, context.revision)) {
    return {
      status: "exact",
      reference,
      descriptor: blockDescriptor(reference, locator.part, resolved, text),
    };
  }
  const relocated = currentReference(context, reference, locator);
  return {
    status: "relocated",
    reference: relocated,
    descriptor: blockDescriptor(relocated, locator.part, resolved, text),
    reasonCodes: ["docx.revision-changed", "docx.locator-stable"],
  };
}

function resolveStoryTextReference(
  context: DocxReferenceContext,
  reference: DocxOfficeReference,
  locator: Extract<DocxObjectLocator, { kind: "text-range" }>,
): ResolveReferenceResult {
  const nodes = findStoryNodes(context.model, locator.part);
  if (!nodes) return { status: "not-found", reasonCode: "docx.part-not-found" };
  if (pathKey(locator.start.path) !== pathKey(locator.end.path)) {
    return { status: "unsupported", reasonCode: "docx.cross-paragraph-text-unsupported" };
  }
  const resolved = resolvePathInNodes(nodes, locator.start.path);
  if (!resolved || resolved.kind !== "paragraph") {
    return { status: "not-found", reasonCode: "docx.text-range-not-found" };
  }
  const text = paragraphText(resolved.paragraph);
  const startOffset = locator.start.offset;
  const endOffset = locator.end.offset;
  if (
    !Number.isInteger(startOffset) ||
    !Number.isInteger(endOffset) ||
    startOffset < 0 ||
    endOffset <= startOffset ||
    endOffset > text.length
  ) {
    return { status: "not-found", reasonCode: "docx.text-range-invalid" };
  }
  const isCurrentRevision = sameRevision(reference.document, context.revision);
  const selected = text.slice(startOffset, endOffset);
  const locatorAt = (start: number, end: number): DocxObjectLocator => ({
    ...locator,
    start: { path: locator.start.path, offset: start },
    end: { path: locator.end.path, offset: end },
  });
  const directMatches =
    hasMatchingTextFingerprint(selected, reference.fingerprint) ||
    (isCurrentRevision && !hasVerifiableTextFingerprint(reference.fingerprint));
  if (directMatches) {
    if (isCurrentRevision) {
      return {
        status: "exact",
        reference,
        descriptor: blockDescriptor(reference, locator.part, resolved, selected),
      };
    }
    const relocated = currentReference(
      context,
      reference,
      locatorAt(startOffset, endOffset),
      undefined,
      textFingerprintAt(text, startOffset, endOffset),
    );
    return {
      status: "relocated",
      reference: relocated,
      descriptor: blockDescriptor(relocated, locator.part, resolved, selected),
      reasonCodes: ["docx.revision-changed", "docx.locator-stable"],
    };
  }
  if (isCurrentRevision) {
    return { status: "not-found", reasonCode: "docx.text-range-invalid" };
  }
  const quote = reference.fingerprint?.exactText;
  if (!quote) return { status: "unsupported", reasonCode: "docx.text-quote-unavailable" };
  const matches: Array<{ start: number; end: number }> = [];
  for (let from = 0; from <= text.length - quote.length; ) {
    const start = text.indexOf(quote, from);
    if (start < 0) break;
    matches.push({ start, end: start + quote.length });
    from = start + 1;
  }
  if (matches.length === 0) {
    return { status: "not-found", reasonCode: "docx.text-not-found" };
  }
  if (matches.length > 1) {
    return {
      status: "ambiguous",
      candidates: matches.map((match) =>
        blockDescriptor(
          currentReference(context, reference, locatorAt(match.start, match.end)),
          locator.part,
          resolved,
          quote,
        ),
      ),
      reasonCodes: ["docx.multiple-text-matches"],
    };
  }
  const match = matches[0]!;
  const relocated = currentReference(
    context,
    reference,
    locatorAt(match.start, match.end),
    undefined,
    textFingerprintAt(text, match.start, match.end),
  );
  return {
    status: "relocated",
    reference: relocated,
    descriptor: blockDescriptor(relocated, locator.part, resolved, quote),
    reasonCodes: ["docx.revision-changed", "docx.text-quote-match"],
  };
}

export function resolveDocxReference(
  context: DocxReferenceContext,
  reference: OfficeObjectReference,
): ResolveReferenceResult {
  assertDocxContext(context);
  if (reference.document.format !== "docx") return { status: "unsupported", reasonCode: "docx.format-mismatch" };
  const docxReference = reference as DocxOfficeReference;
  if (docxReference.document.documentId !== context.revision.documentId) {
    return { status: "not-found", reasonCode: "docx.document-id-mismatch" };
  }
  const kind = docxReferenceKind(docxReference);
  if (kind === "other") return { status: "unsupported", reasonCode: "docx.reference-kind-unsupported" };

  if (kind === "image") return resolveDocxImageReference(context, docxReference);
  if (kind === "structure") return resolveDocxStructureReference(context, docxReference);
  if (kind === "comment") return resolveDocxCommentReference(context, docxReference);

  if (kind === "page" || kind === "region") {
    const pageIndex = docxReference.locator.type === "manual-region"
      ? docxReference.locator.value.pageIndex
      : docxReference.locator.value.kind === "page" ? docxReference.locator.value.pageIndex : -1;
    if (pageIndex < 0 || context.pageCount !== undefined && pageIndex >= context.pageCount) {
      return { status: "not-found", reasonCode: "docx.page-not-found" };
    }
    if (!sameRevision(docxReference.document, context.revision)) {
      return { status: "unsupported", reasonCode: "docx.reflow-requires-reselection" };
    }
    return {
      status: "exact",
      reference: docxReference,
      descriptor: simpleDescriptor(docxReference),
    };
  }

  if (docxReference.locator.type !== "format") {
    return { status: "unsupported", reasonCode: "docx.text-locator-unsupported" };
  }
  if (
    docxReference.locator.value.kind === "text-range" &&
    docxReference.locator.value.part.kind !== "body"
  ) {
    return resolveStoryTextReference(context, docxReference, docxReference.locator.value);
  }
  const range = rangeFromLocator(docxReference.locator.value);
  if (!range) return { status: "unsupported", reasonCode: "docx.text-part-unsupported" };
  const index = buildTextIndex(context.model);
  const selected = indexedRange(index, range);
  const isCurrentRevision = sameRevision(docxReference.document, context.revision);
  const directTextMatches = selected && (
    hasMatchingTextFingerprint(selected.text, docxReference.fingerprint)
    || isCurrentRevision && !hasVerifiableTextFingerprint(docxReference.fingerprint)
  );
  if (selected && directTextMatches) {
    if (isCurrentRevision) {
      return {
        status: "exact",
        reference: docxReference,
        descriptor: textDescriptor(docxReference, selected.range, selected.text),
      };
    }
    const relocated = currentReference(
      context,
      docxReference,
      locatorFromRange(selected.range),
      undefined,
      fingerprintForIndexedText(index, selected.start, selected.end),
    );
    return {
      status: "relocated",
      reference: relocated,
      descriptor: textDescriptor(relocated, selected.range, selected.text),
      reasonCodes: ["docx.revision-changed", "docx.locator-stable"],
    };
  }
  if (isCurrentRevision) {
    return { status: "not-found", reasonCode: "docx.text-range-invalid" };
  }
  if (!docxReference.fingerprint?.exactText) {
    return { status: "unsupported", reasonCode: "docx.text-quote-unavailable" };
  }
  const candidates = textCandidates(context, docxReference, index);
  if (candidates.length === 0) return { status: "not-found", reasonCode: "docx.text-not-found" };
  if (candidates.length > 1) {
    return {
      status: "ambiguous",
      candidates: candidates.map((candidate) => textDescriptor(candidate.reference, candidate.range, candidate.text)),
      reasonCodes: ["docx.multiple-text-matches"],
    };
  }
  const candidate = candidates[0]!;
  return {
    status: "relocated",
    reference: candidate.reference,
    descriptor: textDescriptor(candidate.reference, candidate.range, candidate.text),
    reasonCodes: ["docx.revision-changed", "docx.text-quote-match"],
  };
}

export function describeDocxReference(
  context: DocxReferenceContext,
  reference: OfficeObjectReference,
): OfficeObjectDescriptor | undefined {
  const result = resolveDocxReference(context, reference);
  return result.status === "exact" || result.status === "relocated" ? result.descriptor : undefined;
}
