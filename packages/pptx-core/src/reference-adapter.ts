import type {
  ManualRegionLocator,
  OfficeDocumentRevision,
  OfficeObjectDescriptor,
  OfficeObjectFingerprint,
  OfficeObjectReference,
  OfficeObjectReferenceDraft,
  OfficeObjectReliability,
  PptxObjectKeyLocator,
  PptxObjectLocator,
  PptxReferenceKind,
  PptxSlideLocator,
  PptxSubObjectPathSegment,
  ResolveReferenceResult,
} from "@arcships/office-interaction";
import type {
  PptxObjectIdentity,
  PptxPlaybackDocument,
  PptxPlaybackSlide,
} from "./playback/types";
import type { PptxBehaviorRecord } from "./pptx-behaviors";
import { enumeratePptxSlideBehaviors } from "./pptx-behaviors";
import type { PptxPreviewDocument } from "./types";

export type PptxOfficeReference = Extract<
  OfficeObjectReference,
  { document: { format: "pptx" } }
>;

export type PptxOfficeReferenceDraft = Extract<
  OfficeObjectReferenceDraft,
  { document: { format: "pptx" } }
>;

export interface PptxReferenceContext {
  revision: OfficeDocumentRevision & { format: "pptx" };
  document: PptxPreviewDocument;
  playbackDocument: PptxPlaybackDocument;
  /** Optional native slide ids/part paths, especially useful for empty slides. */
  slideIds?: readonly (string | undefined)[];
  /** Current text paragraphs for a rendered/native object, used by precise text references. */
  getObjectParagraphs?: (slideIndex: number, objectKey: string) => readonly string[] | undefined;
  /**
   * Rendered table grid of an object, when the host can read it. PowerPoint
   * stores cell text inside the table shape; the adapter accepts it from the
   * Surface so `pptx-core` does not need a DOM.
   */
  getObjectTable?: (slideIndex: number, objectKey: string) => PptxTableSnapshot | undefined;
  /**
   * Chart series/point labels of a chart object, when the host can read them.
   * Chart internals are a stage 3 object family; without this provider the
   * chart itself stays selectable and its internals are simply absent.
   */
  getObjectChart?: (slideIndex: number, objectKey: string) => PptxChartSnapshot | undefined;
}

/** One rendered table: rows of cell text, in visual order. */
export interface PptxTableSnapshot {
  rows: readonly (readonly string[])[];
}

/** One chart's series names and per-series point labels/categories. */
export interface PptxChartSnapshot {
  series: readonly {
    name?: string;
    points?: readonly string[];
  }[];
}

export interface PptxObjectReferenceOptions {
  kind?: Extract<PptxReferenceKind, "text-box" | "shape" | "image" | "table" | "chart" | "group">;
  fallbackRegion?: Extract<ManualRegionLocator, { space: "slide" }>;
}

const PPTX_SLIDE_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["pptx.native-slide"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["pptx.slide-boundary"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["pptx.presentation-slide"] },
  relocation: { level: "likely", reasonCodes: ["pptx.slide-id"] },
};

const PPTX_OBJECT_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["pptx.native-object-type"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["pptx.shape-id"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["pptx.group-path"] },
  relocation: { level: "likely", reasonCodes: ["pptx.object-key-or-name"] },
};

const PPTX_TEXT_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["pptx.native-text"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["pptx.paragraph-offsets"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["pptx.object-key-and-paragraph"] },
  relocation: { level: "likely", reasonCodes: ["pptx.object-and-text-quote"] },
};

const PPTX_SUB_OBJECT_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["pptx.native-sub-object"] },
  boundary: { level: "likely", reasonCodes: ["pptx.rendered-grid-or-plot"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["pptx.object-and-index"] },
  relocation: { level: "likely", reasonCodes: ["pptx.object-key-and-index"] },
};

const PPTX_BEHAVIOR_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["pptx.native-behavior"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["pptx.behavior-native-id"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["pptx.slide-id"] },
  relocation: { level: "likely", reasonCodes: ["pptx.behavior-native-id"] },
};

const PPTX_REGION_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "unknown", reasonCodes: ["manual-region"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["normalized-region"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["pptx.slide-id"] },
  relocation: { level: "likely", reasonCodes: ["pptx.slide-id-and-size"] },
};

function copyReliability(value: OfficeObjectReliability): OfficeObjectReliability {
  return {
    semantic: { ...value.semantic, reasonCodes: [...value.semantic.reasonCodes] },
    boundary: { ...value.boundary, reasonCodes: [...value.boundary.reasonCodes] },
    hierarchy: { ...value.hierarchy, reasonCodes: [...value.hierarchy.reasonCodes] },
    relocation: { ...value.relocation, reasonCodes: [...value.relocation.reasonCodes] },
  };
}

function sameRevision(left: OfficeDocumentRevision, right: OfficeDocumentRevision): boolean {
  return left.format === right.format
    && left.documentId === right.documentId
    && left.revision === right.revision;
}

function assertContext(context: PptxReferenceContext): void {
  if (context.revision.format !== "pptx") throw new TypeError("PPTX reference context requires format=pptx");
  if (!context.revision.documentId.trim()) throw new TypeError("PPTX documentId must not be empty");
  if (!context.revision.revision.trim()) throw new TypeError("PPTX revision must not be empty");
  if (!Number.isFinite(context.document.width) || context.document.width <= 0
    || !Number.isFinite(context.document.height) || context.document.height <= 0) {
    throw new RangeError("PPTX document dimensions must be positive finite numbers");
  }
  if (context.document.slides.length !== context.playbackDocument.slides.length) {
    throw new RangeError("PPTX preview and playback documents must contain the same number of slides");
  }
  if (context.slideIds && context.slideIds.length !== context.document.slides.length) {
    throw new RangeError("PPTX slideIds must align with the document slide list");
  }
  context.document.slides.forEach((slide, index) => {
    if (slide.index !== index || slide.number !== index + 1 || context.playbackDocument.slides[index]?.index !== index) {
      throw new RangeError("PPTX slides must use ordered zero-based indices and one-based numbers");
    }
  });
}

function slideAt(context: PptxReferenceContext, slideIndex: number): PptxPlaybackSlide | undefined {
  return Number.isSafeInteger(slideIndex) && slideIndex >= 0
    ? context.playbackDocument.slides[slideIndex]
    : undefined;
}

function slideId(context: PptxReferenceContext, slideIndex: number): string | undefined {
  const explicit = context.slideIds?.[slideIndex]?.trim();
  if (explicit) return explicit;
  return slideAt(context, slideIndex)?.objects.find((object) => object.source === "slide")?.slidePath;
}

function slideLocator(context: PptxReferenceContext, slideIndex: number): PptxSlideLocator {
  const id = slideId(context, slideIndex);
  return { ...(id ? { slideId: id } : {}), index: slideIndex };
}

function slideSizeSignature(context: PptxReferenceContext): string {
  return `pptx.slide-size:${context.document.width}x${context.document.height}`;
}

function slideFingerprint(context: PptxReferenceContext, slideIndex: number): OfficeObjectFingerprint {
  const id = slideId(context, slideIndex);
  return {
    ancestorKeys: [
      ...(id ? [`pptx.slide-id:${id}`] : []),
      slideSizeSignature(context),
    ],
  };
}

function objectLocator(identity: PptxObjectIdentity): PptxObjectKeyLocator {
  return {
    objectKey: identity.key,
    shapeId: identity.shapeId,
    source: identity.source,
    groupPath: [...identity.groupPath],
  };
}

/**
 * Fingerprints use the same slide id the locator and resolver use, so a host
 * that supplies documented `slideIds` keeps matching its own references. The
 * parser part path stays as the fallback when no slide id is available.
 */
function objectFingerprint(identity: PptxObjectIdentity, resolvedSlideId?: string): OfficeObjectFingerprint {
  return {
    ...(identity.name ? { objectName: identity.name } : {}),
    ancestorKeys: [
      `pptx.slide-id:${resolvedSlideId ?? identity.slidePath}`,
      `pptx.node-type:${identity.nodeType}`,
      `pptx.shape-id:${identity.shapeId}`,
      ...identity.groupPath.map((id) => `pptx.group:${id}`),
      ...(identity.creationId ? [`pptx.creation-id:${identity.creationId}`] : []),
    ],
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

type PptxTextBoundary = { paragraphIndex: number; offset: number };

function normalizeTextBoundaries(
  start: PptxTextBoundary,
  end: PptxTextBoundary,
): { start: PptxTextBoundary; end: PptxTextBoundary } {
  for (const boundary of [start, end]) {
    if (!Number.isSafeInteger(boundary.paragraphIndex) || boundary.paragraphIndex < 0
      || !Number.isSafeInteger(boundary.offset) || boundary.offset < 0) {
      throw new TypeError("PPTX text boundaries require non-negative paragraph indices and offsets");
    }
  }
  const ordered = start.paragraphIndex < end.paragraphIndex
    || start.paragraphIndex === end.paragraphIndex && start.offset <= end.offset;
  return ordered ? { start: { ...start }, end: { ...end } } : { start: { ...end }, end: { ...start } };
}

function textForBoundaries(
  paragraphs: readonly string[],
  start: PptxTextBoundary,
  end: PptxTextBoundary,
): string | undefined {
  if (start.paragraphIndex >= paragraphs.length || end.paragraphIndex >= paragraphs.length) return undefined;
  const startText = paragraphs[start.paragraphIndex]!;
  const endText = paragraphs[end.paragraphIndex]!;
  if (start.offset > startText.length || end.offset > endText.length) return undefined;
  if (start.paragraphIndex === end.paragraphIndex) return startText.slice(start.offset, end.offset);
  return [
    startText.slice(start.offset),
    ...paragraphs.slice(start.paragraphIndex + 1, end.paragraphIndex),
    endText.slice(0, end.offset),
  ].join("\n");
}

function textFingerprint(
  identity: PptxObjectIdentity,
  text: string,
  resolvedSlideId?: string,
): OfficeObjectFingerprint {
  const object = objectFingerprint(identity, resolvedSlideId);
  return {
    ...object,
    ...(text.length <= 256 ? { exactText: text } : {}),
    contentHash: textContentHash(text),
  };
}

/** Map the parser's native node type to the public Office object vocabulary. */
export function pptxReferenceKindForObject(identity: PptxObjectIdentity): PptxObjectReferenceOptions["kind"] {
  switch (identity.nodeType.toLowerCase()) {
    case "picture":
    case "image":
    case "pic":
      return "image";
    case "table":
      return "table";
    case "chart":
      return "chart";
    case "group":
    case "groupshape":
    case "grpsp":
      return "group";
    case "textbox":
    case "text-box":
      return "text-box";
    default:
      return "shape";
  }
}

function assertSlideIndex(context: PptxReferenceContext, slideIndex: number): PptxPlaybackSlide {
  const slide = slideAt(context, slideIndex);
  if (!slide) throw new RangeError("PPTX slideIndex is outside the current document");
  return slide;
}

function assertSlideRegion(
  context: PptxReferenceContext,
  region: Extract<ManualRegionLocator, { space: "slide" }>,
): void {
  assertSlideIndex(context, region.slideIndex);
  const { x, y, width, height } = region.rect;
  if (![x, y, width, height].every(Number.isFinite)) throw new TypeError("PPTX region coordinates must be finite");
  if (x < 0 || y < 0 || width <= 0 || height <= 0 || x + width > 1 || y + height > 1) {
    throw new RangeError("PPTX region must be a positive rectangle inside normalized slide space");
  }
}

export function createPptxSlideReferenceDraft(
  context: PptxReferenceContext,
  slideIndex: number,
): PptxOfficeReferenceDraft {
  assertContext(context);
  assertSlideIndex(context, slideIndex);
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "slide",
    source: "native",
    locator: { type: "format", format: "pptx", value: { kind: "slide", slide: slideLocator(context, slideIndex) } },
    fingerprint: slideFingerprint(context, slideIndex),
    reliability: copyReliability(PPTX_SLIDE_RELIABILITY),
  };
}

export function createPptxObjectReferenceDraft(
  context: PptxReferenceContext,
  slideIndex: number,
  objectKey: string,
  options: PptxObjectReferenceOptions = {},
): PptxOfficeReferenceDraft {
  assertContext(context);
  const slide = assertSlideIndex(context, slideIndex);
  const identity = slide.objects.find((object) => object.key === objectKey);
  if (!identity) throw new RangeError("PPTX objectKey is not present on the selected slide");
  if (options.fallbackRegion) {
    assertSlideRegion(context, options.fallbackRegion);
    if (options.fallbackRegion.slideIndex !== slideIndex) {
      throw new RangeError("PPTX object fallback region must use the selected slideIndex");
    }
  }
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: options.kind ?? pptxReferenceKindForObject(identity) ?? "shape",
    source: "native",
    locator: {
      type: "format",
      format: "pptx",
      value: {
        kind: "object",
        slide: slideLocator(context, slideIndex),
        object: objectLocator(identity),
      },
    },
    fingerprint: objectFingerprint(identity, slideId(context, slideIndex)),
    ...(options.fallbackRegion ? { fallbackRegion: options.fallbackRegion } : {}),
    reliability: copyReliability(PPTX_OBJECT_RELIABILITY),
  };
}

/** Create an exclusive-end text range inside a PPTX object's paragraph list. */
export function createPptxTextReferenceDraft(
  context: PptxReferenceContext,
  slideIndex: number,
  objectKey: string,
  start: PptxTextBoundary,
  end: PptxTextBoundary,
): PptxOfficeReferenceDraft {
  assertContext(context);
  const slide = assertSlideIndex(context, slideIndex);
  const identity = slide.objects.find((object) => object.key === objectKey);
  if (!identity) throw new RangeError("PPTX objectKey is not present on the selected slide");
  const range = normalizeTextBoundaries(start, end);
  const paragraphs = context.getObjectParagraphs?.(slideIndex, objectKey);
  if (!paragraphs) throw new TypeError("PPTX text references require getObjectParagraphs evidence");
  const text = textForBoundaries(paragraphs, range.start, range.end);
  if (text === undefined || text.length === 0) throw new RangeError("PPTX text range must select existing non-empty text");
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "text-range",
    source: "native",
    locator: {
      type: "format",
      format: "pptx",
      value: {
        kind: "text-range",
        slide: slideLocator(context, slideIndex),
        object: objectLocator(identity),
        start: range.start,
        end: range.end,
      },
    },
    fingerprint: textFingerprint(identity, text, slideId(context, slideIndex)),
    reliability: copyReliability(PPTX_TEXT_RELIABILITY),
  };
}

export type PptxSubObjectTarget = PptxSubObjectPathSegment;

function subObjectSegment(target: PptxSubObjectTarget): PptxSubObjectPathSegment {
  return { ...target };
}

function tableSnapshot(
  context: PptxReferenceContext,
  slideIndex: number,
  identity: PptxObjectIdentity,
): PptxTableSnapshot {
  const snapshot = context.getObjectTable?.(slideIndex, identity.key);
  if (!snapshot) {
    throw new RangeError("PPTX table cells require getObjectTable evidence");
  }
  return snapshot;
}

function assertSubObjectTarget(
  context: PptxReferenceContext,
  slideIndex: number,
  identity: PptxObjectIdentity,
  target: PptxSubObjectTarget,
): void {
  const nodeType = identity.nodeType.toLowerCase();
  if (target.kind === "table-cell") {
    if (nodeType !== "table") {
      throw new RangeError("PPTX table-cell references require a table object");
    }
    if (!Number.isSafeInteger(target.rowIndex) || target.rowIndex < 0
      || !Number.isSafeInteger(target.cellIndex) || target.cellIndex < 0) {
      throw new TypeError("PPTX table cell indices must be non-negative safe integers");
    }
    const snapshot = tableSnapshot(context, slideIndex, identity);
    const row = snapshot.rows[target.rowIndex];
    if (!row || !row[target.cellIndex]) {
      throw new RangeError("PPTX table cell is outside the rendered table grid");
    }
    return;
  }
  if (nodeType !== "chart") {
    throw new RangeError("PPTX chart element references require a chart object");
  }
  const chart = context.getObjectChart?.(slideIndex, identity.key);
  if (!chart) throw new RangeError("PPTX chart elements require getObjectChart evidence");
  if (target.kind === "chart-series") {
    if (!Number.isSafeInteger(target.seriesIndex) || target.seriesIndex < 0 || !chart.series[target.seriesIndex]) {
      throw new RangeError("PPTX chart series is outside the current chart");
    }
    return;
  }
  if (target.kind === "chart-point") {
    const series = chart.series[target.seriesIndex];
    if (!series) throw new RangeError("PPTX chart series is outside the current chart");
    if (!Number.isSafeInteger(target.pointIndex) || target.pointIndex < 0
      || (series.points !== undefined && target.pointIndex >= series.points.length)) {
      throw new RangeError("PPTX chart point is outside the current series");
    }
    return;
  }
  if (!Number.isSafeInteger(target.legendIndex) || target.legendIndex < 0 || !chart.series[target.legendIndex]) {
    throw new RangeError("PPTX chart legend entry is outside the current chart");
  }
  if (target.seriesIndex !== undefined
    && (!Number.isSafeInteger(target.seriesIndex) || target.seriesIndex < 0 || !chart.series[target.seriesIndex])) {
    throw new RangeError("PPTX chart legend series is outside the current chart");
  }
}

function subObjectFingerprint(
  identity: PptxObjectIdentity,
  target: PptxSubObjectTarget,
  resolvedSlideId: string | undefined,
  text: string | undefined,
): OfficeObjectFingerprint {
  const base = objectFingerprint(identity, resolvedSlideId);
  return {
    ...base,
    ancestorKeys: [
      ...(base.ancestorKeys ?? []),
      `pptx.sub-object:${target.kind}`,
    ],
    ...(text ? { contentHash: textContentHash(text) } : {}),
  };
}

function subObjectText(
  context: PptxReferenceContext,
  slideIndex: number,
  identity: PptxObjectIdentity,
  target: PptxSubObjectTarget,
): string | undefined {
  if (target.kind === "table-cell") {
    return context.getObjectTable?.(slideIndex, identity.key)?.rows[target.rowIndex]?.[target.cellIndex];
  }
  const chart = context.getObjectChart?.(slideIndex, identity.key);
  if (!chart) return undefined;
  if (target.kind === "chart-series") return chart.series[target.seriesIndex]?.name;
  if (target.kind === "chart-point") {
    return chart.series[target.seriesIndex]?.points?.[target.pointIndex];
  }
  const seriesIndex = target.seriesIndex ?? target.legendIndex;
  return chart.series[seriesIndex]?.name;
}

/**
 * Create a reference to a table cell or chart internal element (stage 3).
 *
 * The target path is validated against the rendered table grid or chart
 * snapshot the Surface provides; without that evidence the reference is
 * rejected instead of guessing.
 */
export function createPptxSubObjectReferenceDraft(
  context: PptxReferenceContext,
  slideIndex: number,
  objectKey: string,
  target: PptxSubObjectTarget,
): PptxOfficeReferenceDraft {
  assertContext(context);
  const slide = assertSlideIndex(context, slideIndex);
  const identity = slide.objects.find((object) => object.key === objectKey);
  if (!identity) throw new RangeError("PPTX objectKey is not present on the selected slide");
  assertSubObjectTarget(context, slideIndex, identity, target);
  const kind = target.kind === "table-cell"
    ? "table-cell"
    : target.kind === "chart-series"
      ? "chart-series"
      : target.kind === "chart-point" ? "chart-point" : "chart-legend-entry";
  return {
    schemaVersion: 1,
    document: context.revision,
    kind,
    source: "native",
    locator: {
      type: "format",
      format: "pptx",
      value: {
        kind: "sub-object",
        slide: slideLocator(context, slideIndex),
        object: objectLocator(identity),
        path: [subObjectSegment(target)],
      },
    },
    fingerprint: subObjectFingerprint(
      identity,
      target,
      slideId(context, slideIndex),
      subObjectText(context, slideIndex, identity, target),
    ),
    reliability: copyReliability(PPTX_SUB_OBJECT_RELIABILITY),
  };
}

function subObjectLabel(
  target: PptxSubObjectTarget,
  text: string | undefined,
): string {
  if (target.kind === "table-cell") {
    return text?.trim()
      ? `Cell ${target.rowIndex + 1}, ${target.cellIndex + 1}: ${text.replace(/\s+/gu, " ").trim().slice(0, 60)}`
      : `Cell ${target.rowIndex + 1}, ${target.cellIndex + 1}`;
  }
  if (target.kind === "chart-series") return `Series: ${text?.trim() || `Series ${target.seriesIndex + 1}`}`;
  if (target.kind === "chart-point") {
    return text?.trim() ? `Point ${target.pointIndex + 1}: ${text.slice(0, 60)}` : `Point ${target.pointIndex + 1}`;
  }
  return `Legend entry ${target.legendIndex + 1}`;
}

function subObjectDescriptor(
  reference: PptxOfficeReference,
  slideIndex: number,
  identity: PptxObjectIdentity,
  target: PptxSubObjectTarget,
  text: string | undefined,
): OfficeObjectDescriptor {
  const kind = reference.kind === "region" || reference.kind === "slide" ? "shape" : reference.kind;
  const objectLabel = identity.name?.trim() || `Object ${identity.shapeId}`;
  return {
    objectId: `pptx:sub-object:${identity.key}:${target.kind}:${JSON.stringify(target)}`,
    draft: draftFromReference(reference),
    label: subObjectLabel(target, text),
    path: [
      { kind: "document", label: "Presentation" },
      { kind: "slide", label: `Slide ${slideIndex + 1}` },
      ...identity.groupPath.map((id) => ({ kind: "group" as const, label: `Group ${id}` })),
      { kind: identity.nodeType === "table" ? "table" : "chart", label: objectLabel },
      { kind, label: subObjectLabel(target, text) },
    ],
    childrenState: "none",
    ...(text ? { content: { text } } : {}),
    facets: {
      slideIndex,
      objectKey: identity.key,
      nodeType: identity.nodeType,
      path: [subObjectSegment(target)],
    },
  };
}

function resolvePptxSubObjectReference(
  context: PptxReferenceContext,
  reference: PptxOfficeReference,
  slideIndex: number,
  identity: PptxObjectIdentity,
  locator: Extract<PptxObjectLocator, { kind: "sub-object" }>,
  currentRevision: boolean,
): ResolveReferenceResult {
  const segment = locator.path[locator.path.length - 1];
  if (!segment) return { status: "unsupported", reasonCode: "pptx.sub-object-path-empty" };
  try {
    assertSubObjectTarget(context, slideIndex, identity, segment);
  } catch (reason) {
    const evidenceMissing = reason instanceof RangeError
      && /evidence/u.test(reason.message);
    return evidenceMissing
      ? { status: "unsupported", reasonCode: "pptx.sub-object-evidence-unavailable" }
      : { status: "not-found", reasonCode: "pptx.sub-object-not-found" };
  }
  const text = subObjectText(context, slideIndex, identity, segment);
  if (currentRevision) {
    return {
      status: "exact",
      reference,
      descriptor: subObjectDescriptor(reference, slideIndex, identity, segment, text),
    };
  }
  const relocated = currentFormatReference(context, reference, {
    kind: "sub-object",
    slide: slideLocator(context, slideIndex),
    object: objectLocator(identity),
    path: [segment],
  }, subObjectFingerprint(identity, segment, slideId(context, slideIndex), text));
  return {
    status: "relocated",
    reference: relocated,
    descriptor: subObjectDescriptor(relocated, slideIndex, identity, segment, text),
    reasonCodes: ["pptx.revision-changed", "pptx.object-key-and-index"],
  };
}

export interface PptxBehaviorTarget {
  behavior: PptxBehaviorRecord["behavior"];
  /** Native id from `enumeratePptxBehaviors()`; deterministic per slide. */
  instanceId: string;
}

/** Enumerate stage 4 behaviors for one slide. */
export function enumeratePptxBehaviors(
  context: PptxReferenceContext,
  slideIndex: number,
): readonly PptxBehaviorRecord[] {
  return enumeratePptxSlideBehaviors(context, slideIndex);
}

/** Create an invisible-behavior reference draft (stage 4). */
export function createPptxBehaviorReferenceDraft(
  context: PptxReferenceContext,
  slideIndex: number,
  target: PptxBehaviorTarget,
): PptxOfficeReferenceDraft {
  assertContext(context);
  assertSlideIndex(context, slideIndex);
  const record = enumeratePptxBehaviors(context, slideIndex)
    .find((candidate) => candidate.instanceId === target.instanceId);
  if (!record || record.behavior !== target.behavior) {
    throw new RangeError(`PPTX behavior ${target.instanceId} is not present on the selected slide`);
  }
  return createPptxBehaviorDraftFromRecord(context, slideIndex, record);
}

/** Draft for an already-enumerated behavior record. */
export function createPptxBehaviorDraftFromRecord(
  context: PptxReferenceContext,
  slideIndex: number,
  record: PptxBehaviorRecord,
): PptxOfficeReferenceDraft {
  assertContext(context);
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: record.behavior,
    source: "native",
    locator: {
      type: "format",
      format: "pptx",
      value: {
        kind: "behavior",
        slide: slideLocator(context, slideIndex),
        behavior: record.behavior,
        nativeId: record.instanceId,
        ...(record.ownerObjectKey === undefined ? {} : { ownerObjectKey: record.ownerObjectKey }),
      },
    },
    fingerprint: {
      ancestorKeys: [
        `pptx.slide-id:${slideId(context, slideIndex) ?? slideIndex}`,
        `pptx.behavior:${record.behavior}`,
        `pptx.behavior-native-id:${record.instanceId}`,
      ],
    },
    reliability: copyReliability(PPTX_BEHAVIOR_RELIABILITY),
  } as PptxOfficeReferenceDraft;
}

function behaviorDescriptor(
  reference: PptxOfficeReference,
  record: PptxBehaviorRecord,
): OfficeObjectDescriptor {
  return {
    objectId: `pptx:behavior:${record.instanceId}`,
    draft: draftFromReference(reference),
    label: record.label,
    path: record.path.map((entry) => ({ ...entry })),
    childrenState: "none",
    ...(record.content === undefined ? {} : { content: { ...record.content } }),
    facets: { ...record.facets, behavior: record.behavior, instanceId: record.instanceId },
  };
}

export function createPptxRegionReferenceDraft(
  context: PptxReferenceContext,
  region: Extract<ManualRegionLocator, { space: "slide" }>,
): PptxOfficeReferenceDraft {
  assertContext(context);
  assertSlideRegion(context, region);
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "region",
    source: "manual",
    locator: { type: "manual-region", format: "pptx", value: region },
    fingerprint: slideFingerprint(context, region.slideIndex),
    fallbackRegion: region,
    reliability: copyReliability(PPTX_REGION_RELIABILITY),
  };
}

function draftFromReference(reference: PptxOfficeReference): PptxOfficeReferenceDraft {
  const { referenceId: _referenceId, ...draft } = reference;
  return draft;
}

function currentFormatReference(
  context: PptxReferenceContext,
  reference: PptxOfficeReference,
  locator: PptxObjectLocator,
  fingerprint: OfficeObjectFingerprint | undefined = reference.fingerprint,
  fallbackRegion: Extract<ManualRegionLocator, { space: "slide" }> | undefined = reference.fallbackRegion,
): PptxOfficeReference {
  const { fallbackRegion: _oldFallback, ...base } = reference;
  return {
    ...base,
    document: context.revision,
    locator: { type: "format", format: "pptx", value: locator },
    ...(fingerprint ? { fingerprint } : {}),
    ...(fallbackRegion ? { fallbackRegion } : {}),
  } as PptxOfficeReference;
}

function currentRegionReference(
  context: PptxReferenceContext,
  reference: PptxOfficeReference,
  slideIndex: number,
): PptxOfficeReference {
  if (reference.locator.type !== "manual-region") return reference;
  const value = { ...reference.locator.value, slideIndex };
  return {
    ...reference,
    document: context.revision,
    locator: { type: "manual-region", format: "pptx", value },
    fallbackRegion: value,
    fingerprint: slideFingerprint(context, slideIndex),
  } as PptxOfficeReference;
}

function idFromFingerprint(fingerprint?: OfficeObjectFingerprint): string | undefined {
  return fingerprint?.ancestorKeys
    ?.find((key) => key.startsWith("pptx.slide-id:"))
    ?.slice("pptx.slide-id:".length);
}

function nodeTypeFromFingerprint(fingerprint?: OfficeObjectFingerprint): string | undefined {
  return fingerprint?.ancestorKeys
    ?.find((key) => key.startsWith("pptx.node-type:"))
    ?.slice("pptx.node-type:".length);
}

function slideCandidates(
  context: PptxReferenceContext,
  locator: PptxSlideLocator,
  currentRevision: boolean,
  fingerprint?: OfficeObjectFingerprint,
): { indices: readonly number[]; reasonCode: string; hasStableId: boolean } {
  const id = locator.slideId ?? idFromFingerprint(fingerprint);
  if (id) {
    const indices = context.document.slides
      .map((slide) => slide.index)
      .filter((index) => slideId(context, index) === id);
    if (indices.length) return { indices, reasonCode: "pptx.slide-id-match", hasStableId: true };
    return { indices: [], reasonCode: "pptx.slide-not-found", hasStableId: true };
  }
  if (currentRevision && slideAt(context, locator.index)) {
    return { indices: [locator.index], reasonCode: "pptx.slide-index-match", hasStableId: false };
  }
  return { indices: [], reasonCode: "pptx.slide-id-unavailable", hasStableId: false };
}

function slideDescriptor(reference: PptxOfficeReference, slideIndex: number): OfficeObjectDescriptor {
  return {
    objectId: `pptx:slide:${slideIndex}`,
    draft: draftFromReference(reference),
    label: `Slide ${slideIndex + 1}`,
    path: [
      { kind: "document", label: "Presentation" },
      { kind: "slide", label: `Slide ${slideIndex + 1}` },
    ],
    childrenState: "available",
    facets: { slideIndex },
  };
}

function objectDescriptor(
  reference: PptxOfficeReference,
  slideIndex: number,
  identity: PptxObjectIdentity,
): OfficeObjectDescriptor {
  const kind = reference.kind === "region" || reference.kind === "slide" ? "shape" : reference.kind;
  const label = identity.name?.trim() || `${kind} ${identity.shapeId}`;
  const visual = reference.fallbackRegion
    ? {
        fragments: [{
          container: { space: "slide" as const, slideIndex },
          rect: reference.fallbackRegion.rect,
        }],
        layoutVersion: reference.document.revision,
      }
    : undefined;
  return {
    objectId: `pptx:object:${identity.key}`,
    draft: draftFromReference(reference),
    label,
    path: [
      { kind: "document", label: "Presentation" },
      { kind: "slide", label: `Slide ${slideIndex + 1}` },
      ...identity.groupPath.map((id) => ({ kind: "group" as const, label: `Group ${id}` })),
      { kind, label },
    ],
    childrenState: kind === "group" || kind === "table" || kind === "chart" ? "available" : "none",
    ...(visual ? { visual } : {}),
    facets: {
      slideIndex,
      objectKey: identity.key,
      shapeId: identity.shapeId,
      source: identity.source,
      nodeType: identity.nodeType,
      groupPath: [...identity.groupPath],
    },
  };
}

function textDescriptor(
  reference: PptxOfficeReference,
  slideIndex: number,
  identity: PptxObjectIdentity,
  text: string,
): OfficeObjectDescriptor {
  const objectLabel = identity.name?.trim() || `Object ${identity.shapeId}`;
  const locator = reference.locator.type === "format" && reference.locator.value.kind === "text-range"
    ? reference.locator.value
    : undefined;
  return {
    objectId: `pptx:text:${identity.key}:${locator?.start.paragraphIndex ?? 0}:${locator?.start.offset ?? 0}:${locator?.end.paragraphIndex ?? 0}:${locator?.end.offset ?? 0}`,
    draft: draftFromReference(reference),
    label: text.length > 80 ? `${text.slice(0, 77)}…` : text,
    path: [
      { kind: "document", label: "Presentation" },
      { kind: "slide", label: `Slide ${slideIndex + 1}` },
      ...identity.groupPath.map((id) => ({ kind: "group" as const, label: `Group ${id}` })),
      { kind: "text-box", label: objectLabel },
      { kind: "text-range", label: "Selected text" },
    ],
    childrenState: "none",
    content: { text },
    facets: {
      slideIndex,
      objectKey: identity.key,
      startParagraphIndex: locator?.start.paragraphIndex ?? 0,
      startOffset: locator?.start.offset ?? 0,
      endParagraphIndex: locator?.end.paragraphIndex ?? 0,
      endOffset: locator?.end.offset ?? 0,
    },
  };
}

function regionDescriptor(reference: PptxOfficeReference, slideIndex: number): OfficeObjectDescriptor {
  const region = reference.locator.type === "manual-region" ? reference.locator.value : reference.fallbackRegion!;
  return {
    objectId: `pptx:region:${slideIndex}:${JSON.stringify(region.rect)}`,
    draft: draftFromReference(reference),
    label: `Region on slide ${slideIndex + 1}`,
    path: [
      { kind: "document", label: "Presentation" },
      { kind: "slide", label: `Slide ${slideIndex + 1}` },
      { kind: "region", label: "Selected region" },
    ],
    childrenState: "none",
    visual: {
      fragments: [{ container: { space: "slide", slideIndex }, rect: region.rect }],
      layoutVersion: reference.document.revision,
    },
  };
}

function groupPathMatches(left: readonly string[], right: readonly string[]): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

function referencesForObjects(
  context: PptxReferenceContext,
  reference: PptxOfficeReference,
  slideIndex: number,
  identities: readonly PptxObjectIdentity[],
): readonly PptxOfficeReference[] {
  return identities.map((identity) => currentFormatReference(context, reference, {
    kind: "object",
    slide: slideLocator(context, slideIndex),
    object: objectLocator(identity),
  }, objectFingerprint(identity, slideId(context, slideIndex)), reference.fallbackRegion
    ? { ...reference.fallbackRegion, slideIndex }
    : undefined));
}

function textReferenceAt(
  context: PptxReferenceContext,
  reference: PptxOfficeReference,
  slideIndex: number,
  identity: PptxObjectIdentity,
  start: PptxTextBoundary,
  end: PptxTextBoundary,
  text: string,
): PptxOfficeReference {
  return currentFormatReference(context, reference, {
    kind: "text-range",
    slide: slideLocator(context, slideIndex),
    object: objectLocator(identity),
    start,
    end,
  }, textFingerprint(identity, text, slideId(context, slideIndex)));
}

function flattenedParagraphText(paragraphs: readonly string[]): { text: string; starts: number[] } {
  const starts: number[] = [];
  let text = "";
  paragraphs.forEach((paragraph, index) => {
    if (index > 0) text += "\n";
    starts.push(text.length);
    text += paragraph;
  });
  return { text, starts };
}

function boundaryAtGlobalOffset(
  paragraphs: readonly string[],
  starts: readonly number[],
  offset: number,
  edge: "start" | "end",
): PptxTextBoundary | undefined {
  for (let index = 0; index < paragraphs.length; index += 1) {
    const start = starts[index]!;
    const end = start + paragraphs[index]!.length;
    if (offset < end || offset === end && (edge === "end" || index === paragraphs.length - 1)) {
      return { paragraphIndex: index, offset: Math.max(0, offset - start) };
    }
    if (offset === end + 1 && edge === "start" && index + 1 < paragraphs.length) {
      return { paragraphIndex: index + 1, offset: 0 };
    }
  }
  return undefined;
}

function textMatchesForIdentity(
  context: PptxReferenceContext,
  reference: PptxOfficeReference,
  slideIndex: number,
  identity: PptxObjectIdentity,
): readonly { reference: PptxOfficeReference; text: string }[] {
  if (reference.locator.type !== "format" || reference.locator.value.kind !== "text-range") return [];
  const locator = reference.locator.value;
  const paragraphs = context.getObjectParagraphs?.(slideIndex, identity.key);
  if (!paragraphs) return [];
  const atOffsets = textForBoundaries(paragraphs, locator.start, locator.end);
  const expected = reference.fingerprint?.exactText;
  const hash = reference.fingerprint?.contentHash;
  const matchesExpected = (text: string) => expected !== undefined
    ? text === expected
    : hash?.startsWith("fnv1a32:") === true && textContentHash(text) === hash;
  if (!expected) {
    return atOffsets !== undefined && matchesExpected(atOffsets)
      ? [{ reference: textReferenceAt(context, reference, slideIndex, identity, locator.start, locator.end, atOffsets), text: atOffsets }]
      : [];
  }
  const flattened = flattenedParagraphText(paragraphs);
  const matches: { reference: PptxOfficeReference; text: string }[] = [];
  let from = 0;
  while (from <= flattened.text.length) {
    const index = flattened.text.indexOf(expected, from);
    if (index < 0) break;
    const start = boundaryAtGlobalOffset(paragraphs, flattened.starts, index, "start");
    const end = boundaryAtGlobalOffset(paragraphs, flattened.starts, index + expected.length, "end");
    if (start && end) matches.push({
      reference: textReferenceAt(context, reference, slideIndex, identity, start, end, expected),
      text: expected,
    });
    from = index + Math.max(1, expected.length);
  }
  return matches;
}

export function resolvePptxReference(
  context: PptxReferenceContext,
  reference: OfficeObjectReference,
): ResolveReferenceResult {
  assertContext(context);
  if (reference.document.format !== "pptx") return { status: "unsupported", reasonCode: "pptx.format-mismatch" };
  const pptxReference = reference as PptxOfficeReference;
  if (pptxReference.document.documentId !== context.revision.documentId) {
    return { status: "not-found", reasonCode: "pptx.document-id-mismatch" };
  }
  const currentRevision = sameRevision(pptxReference.document, context.revision);

  if (pptxReference.locator.type === "manual-region") {
    const locator = { index: pptxReference.locator.value.slideIndex };
    const candidates = slideCandidates(context, locator, currentRevision, pptxReference.fingerprint);
    if (candidates.indices.length === 0) {
      return candidates.hasStableId
        ? { status: "not-found", reasonCode: candidates.reasonCode }
        : { status: "unsupported", reasonCode: candidates.reasonCode };
    }
    const sizeMatches = pptxReference.fingerprint?.ancestorKeys?.includes(slideSizeSignature(context)) === true;
    if (!currentRevision && !sizeMatches) {
      return pptxReference.fingerprint?.ancestorKeys?.some((key) => key.startsWith("pptx.slide-size:"))
        ? { status: "not-found", reasonCode: "pptx.slide-size-changed" }
        : { status: "unsupported", reasonCode: "pptx.slide-size-unavailable" };
    }
    if (candidates.indices.length > 1) {
      return {
        status: "ambiguous",
        candidates: candidates.indices.map((index) => {
          const current = currentRegionReference(context, pptxReference, index);
          return regionDescriptor(current, index);
        }),
        reasonCodes: ["pptx.multiple-slide-matches"],
      };
    }
    const index = candidates.indices[0]!;
    if (currentRevision) {
      return { status: "exact", reference: pptxReference, descriptor: regionDescriptor(pptxReference, index) };
    }
    const relocated = currentRegionReference(context, pptxReference, index);
    return {
      status: "relocated",
      reference: relocated,
      descriptor: regionDescriptor(relocated, index),
      reasonCodes: ["pptx.revision-changed", candidates.reasonCode, "pptx.slide-size-stable"],
    };
  }

  const locator = pptxReference.locator.value;
  if (locator.kind !== "slide" && locator.kind !== "object" && locator.kind !== "text-range"
    && locator.kind !== "sub-object" && locator.kind !== "behavior") {
    return { status: "unsupported", reasonCode: "pptx.reference-kind-unsupported" };
  }
  const candidates = slideCandidates(context, locator.slide, currentRevision, pptxReference.fingerprint);
  if (candidates.indices.length === 0) {
    return candidates.hasStableId
      ? { status: "not-found", reasonCode: candidates.reasonCode }
      : { status: "unsupported", reasonCode: candidates.reasonCode };
  }
  if (candidates.indices.length > 1) {
    const descriptors = candidates.indices.flatMap((index) => {
      if (locator.kind === "slide") {
        const current = currentFormatReference(context, pptxReference, {
          kind: "slide",
          slide: slideLocator(context, index),
        }, slideFingerprint(context, index));
        return [slideDescriptor(current, index)];
      }
      const slide = slideAt(context, index)!;
      if (locator.kind === "behavior") {
        return enumeratePptxBehaviors(context, index)
          .filter((record) => record.instanceId === locator.nativeId)
          .map((record) => behaviorDescriptor(
            currentFormatReference(context, pptxReference, {
              kind: "behavior",
              slide: slideLocator(context, index),
              behavior: record.behavior,
              nativeId: record.instanceId,
              ...(record.ownerObjectKey === undefined ? {} : { ownerObjectKey: record.ownerObjectKey }),
            }),
            record,
          ));
      }
      if (locator.kind === "sub-object") {
        return slide.objects
          .filter((identity) => identity.key === locator.object.objectKey)
          .flatMap((identity) => {
            const segment = locator.path[locator.path.length - 1];
            if (!segment) return [];
            try {
              assertSubObjectTarget(context, index, identity, segment);
            } catch {
              return [];
            }
            const current = currentFormatReference(context, pptxReference, {
              kind: "sub-object",
              slide: slideLocator(context, index),
              object: objectLocator(identity),
              path: [segment],
            });
            return [subObjectDescriptor(
              current,
              index,
              identity,
              segment,
              subObjectText(context, index, identity, segment),
            )];
          });
      }
      if (locator.kind === "text-range") {
        return slide.objects.flatMap((identity) => {
          const objectMatches = identity.key === locator.object.objectKey
            || identity.shapeId === locator.object.shapeId && identity.source === locator.object.source
              && groupPathMatches(identity.groupPath, locator.object.groupPath);
          if (!objectMatches) return [];
          return textMatchesForIdentity(context, pptxReference, index, identity)
            .map((match) => textDescriptor(match.reference, index, identity, match.text));
        });
      }
      return slide.objects
        .filter((identity) => identity.key === locator.object.objectKey)
        .map((identity) => objectDescriptor(referencesForObjects(context, pptxReference, index, [identity])[0]!, index, identity));
    });
    return { status: "ambiguous", candidates: descriptors, reasonCodes: ["pptx.multiple-slide-matches"] };
  }
  const index = candidates.indices[0]!;

  if (locator.kind === "slide") {
    const current = currentFormatReference(context, pptxReference, {
      kind: "slide",
      slide: slideLocator(context, index),
    }, slideFingerprint(context, index));
    if (currentRevision) return { status: "exact", reference: pptxReference, descriptor: slideDescriptor(pptxReference, index) };
    return {
      status: "relocated",
      reference: current,
      descriptor: slideDescriptor(current, index),
      reasonCodes: ["pptx.revision-changed", candidates.reasonCode],
    };
  }

  const slide = slideAt(context, index)!;

  if (locator.kind === "behavior") {
    const record = enumeratePptxBehaviors(context, index)
      .find((candidate) => candidate.instanceId === locator.nativeId);
    if (!record || record.behavior !== locator.behavior) {
      return { status: "not-found", reasonCode: "pptx.behavior-not-found" };
    }
    if (currentRevision) {
      return { status: "exact", reference: pptxReference, descriptor: behaviorDescriptor(pptxReference, record) };
    }
    const relocated = currentFormatReference(context, pptxReference, {
      kind: "behavior",
      slide: slideLocator(context, index),
      behavior: record.behavior,
      nativeId: record.instanceId,
      ...(record.ownerObjectKey === undefined ? {} : { ownerObjectKey: record.ownerObjectKey }),
    });
    return {
      status: "relocated",
      reference: relocated,
      descriptor: behaviorDescriptor(relocated, record),
      reasonCodes: ["pptx.revision-changed", "pptx.behavior-native-id"],
    };
  }

  let identities = slide.objects.filter((identity) => identity.key === locator.object.objectKey);
  let objectReason = "pptx.object-key-match";
  if (identities.length === 0) {
    identities = slide.objects.filter((identity) => identity.shapeId === locator.object.shapeId
      && identity.source === locator.object.source
      && groupPathMatches(identity.groupPath, locator.object.groupPath));
    objectReason = "pptx.shape-group-path-match";
  }
  if (identities.length === 0 && !currentRevision && pptxReference.fingerprint?.objectName) {
    const nodeType = nodeTypeFromFingerprint(pptxReference.fingerprint);
    identities = slide.objects.filter((identity) => identity.name === pptxReference.fingerprint?.objectName
      && (!nodeType || identity.nodeType === nodeType));
    objectReason = "pptx.object-name-type-match";
  }
  if (identities.length === 0) return { status: "not-found", reasonCode: "pptx.object-not-found" };
  if (locator.kind === "sub-object") {
    if (identities.length === 0) return { status: "not-found", reasonCode: "pptx.object-not-found" };
    if (identities.length > 1) {
      const segment = locator.path[locator.path.length - 1];
      return {
        status: "ambiguous",
        candidates: segment === undefined ? [] : identities.flatMap((identity) => {
          try {
            assertSubObjectTarget(context, index, identity, segment);
          } catch {
            return [];
          }
          return [subObjectDescriptor(pptxReference, index, identity, segment, subObjectText(context, index, identity, segment))];
        }),
        reasonCodes: ["pptx.multiple-object-matches"],
      };
    }
    return resolvePptxSubObjectReference(context, pptxReference, index, identities[0]!, locator, currentRevision);
  }

  if (locator.kind === "text-range") {
    if (!context.getObjectParagraphs) return { status: "unsupported", reasonCode: "pptx.text-evidence-unavailable" };
    const matches = identities.flatMap((identity) => textMatchesForIdentity(context, pptxReference, index, identity)
      .map((match) => ({ ...match, identity })));
    if (matches.length === 0) return { status: "not-found", reasonCode: "pptx.text-not-found" };
    if (matches.length > 1) {
      return {
        status: "ambiguous",
        candidates: matches.map((match) => textDescriptor(match.reference, index, match.identity, match.text)),
        reasonCodes: ["pptx.multiple-text-matches"],
      };
    }
    const match = matches[0]!;
    if (currentRevision && match.reference.locator.type === "format"
      && match.reference.locator.value.kind === "text-range"
      && match.reference.locator.value.start.paragraphIndex === locator.start.paragraphIndex
      && match.reference.locator.value.start.offset === locator.start.offset
      && match.reference.locator.value.end.paragraphIndex === locator.end.paragraphIndex
      && match.reference.locator.value.end.offset === locator.end.offset) {
      return { status: "exact", reference: pptxReference, descriptor: textDescriptor(pptxReference, index, match.identity, match.text) };
    }
    return {
      status: "relocated",
      reference: match.reference,
      descriptor: textDescriptor(match.reference, index, match.identity, match.text),
      reasonCodes: ["pptx.revision-changed", candidates.reasonCode, objectReason, "pptx.text-quote-match"],
    };
  }
  const currentReferences = referencesForObjects(context, pptxReference, index, identities);
  if (identities.length > 1) {
    return {
      status: "ambiguous",
      candidates: identities.map((identity, identityIndex) => objectDescriptor(
        currentReferences[identityIndex]!,
        index,
        identity,
      )),
      reasonCodes: ["pptx.multiple-object-matches"],
    };
  }
  if (currentRevision) {
    return { status: "exact", reference: pptxReference, descriptor: objectDescriptor(pptxReference, index, identities[0]!) };
  }
  return {
    status: "relocated",
    reference: currentReferences[0]!,
    descriptor: objectDescriptor(currentReferences[0]!, index, identities[0]!),
    reasonCodes: ["pptx.revision-changed", candidates.reasonCode, objectReason],
  };
}

export function describePptxReference(
  context: PptxReferenceContext,
  reference: OfficeObjectReference,
): OfficeObjectDescriptor | undefined {
  const result = resolvePptxReference(context, reference);
  return result.status === "exact" || result.status === "relocated" ? result.descriptor : undefined;
}
