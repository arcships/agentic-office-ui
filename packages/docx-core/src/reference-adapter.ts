import type {
  DocxObjectLocator,
  DocxPartLocator,
  DocxPathSegment,
  ManualRegionLocator,
  OfficeDocumentRevision,
  OfficeObjectDescriptor,
  OfficeObjectFingerprint,
  OfficeObjectReference,
  OfficeObjectReferenceDraft,
  OfficeObjectReliability,
  ResolveReferenceResult,
} from "@arcships/office-interaction";
import type { DocModel, ParagraphNode } from "./engine/types";
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

function docxReferenceKind(reference: DocxOfficeReference): "text" | "page" | "region" | "other" {
  if (reference.locator.type === "manual-region") return "region";
  if (reference.locator.value.kind === "text-range") return "text";
  if (reference.locator.value.kind === "page") return "page";
  return "other";
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
