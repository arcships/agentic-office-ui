import type {
  ManualRegionLocator,
  NormalizedRect,
  OfficeDocumentRevision,
  OfficeObjectDescriptor,
  OfficeObjectFingerprint,
  OfficeObjectReference,
  OfficeObjectReferenceDraft,
  OfficeObjectReliability,
  PdfObjectLocator,
  ResolveReferenceResult,
} from "@arcships/office-interaction"
import { formatOfficeMessage, officeMessagesEn } from "@arcships/office-interaction"
import type { PageTextSlice } from "@embedpdf/models"
import type {
  PdfRenderAnnotation,
  PdfRenderDocument,
  PdfRenderPageInfo,
  PdfRenderRect,
  PdfRenderRuntime,
  PdfSearchHit,
} from "./pdf-render-runtime"
import type {
  PdfBehaviorRecord,
  PdfVisualObjectDetection,
} from "./pdf-native-objects"
import {
  PDF_BEHAVIOR_REFERENCE_KIND,
  pdfBehaviorLocatorValue,
  pdfNativeObjectType,
} from "./pdf-native-objects"
import {
  fnv1a32,
  isPdfPageContentSignature,
  isPdfPageSignature,
  pdfPageAncestorKeys,
  pdfPageBaseSignature,
  pdfPageContentSignature,
} from "./pdf-page-signature"

export type PdfOfficeReference = Extract<
  OfficeObjectReference,
  { document: { format: "pdf" } }
>

export type PdfOfficeReferenceDraft = Extract<
  OfficeObjectReferenceDraft,
  { document: { format: "pdf" } }
>

export interface PdfReferenceDescriptorLabels {
  document: string
  page: (pageIndex: number) => string
  selectedText: string
  textRange: string
  selectedRegion: string
  textLabel: (text: string) => string
  regionOnPage: (pageIndex: number) => string
}

export interface PdfReferenceContext {
  revision: OfficeDocumentRevision & { format: "pdf" }
  document: PdfRenderDocument
  runtime: Pick<PdfRenderRuntime, "getTextSlices" | "search"> & {
    getPageContentHash?: PdfRenderRuntime["getPageContentHash"]
  }
  /** Localized labels for descriptors emitted to hosts. */
  descriptors?: PdfReferenceDescriptorLabels
  /** Native link/widget/annotation objects of the current revision. */
  annotations?: readonly PdfRenderAnnotation[]
  /** Visual recognizer results of the current revision, one page at a time. */
  visualObjects?: readonly (PdfVisualObjectDetection & { pageIndex: number })[]
  /** Stage 4 behaviors: link actions, form rules, annotation state, … */
  behaviors?: readonly PdfBehaviorRecord[]
}

export interface PdfTextReferenceSelection {
  range: PageTextSlice
  text: string
  rects?: readonly PdfRenderRect[]
}

const TEXT_QUOTE_LIMIT = 256

const PDF_TEXT_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["pdf.native-text"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["pdf.character-range"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["pdf.page-index"] },
  relocation: { level: "likely", reasonCodes: ["pdf.text-quote"] },
}

const PDF_PAGE_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["pdf.native-page"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["pdf.page-boundary"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["pdf.page-index"] },
  relocation: { level: "likely", reasonCodes: ["pdf.page-signature"] },
}

const PDF_LINK_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["pdf.native-link"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["pdf.annotation-rect"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["pdf.page-index"] },
  relocation: { level: "likely", reasonCodes: ["pdf.annotation-id"] },
}

const PDF_FORM_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["pdf.native-form-widget"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["pdf.widget-rect"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["pdf.page-index"] },
  relocation: { level: "likely", reasonCodes: ["pdf.widget-id"] },
}

const PDF_ANNOTATION_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["pdf.native-annotation"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["pdf.annotation-rect"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["pdf.page-index"] },
  relocation: { level: "uncertain", reasonCodes: ["pdf.annotation-id"] },
}

const PDF_BEHAVIOR_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["pdf.native-behavior"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["pdf.behavior-native-id"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["pdf.page-index"] },
  relocation: { level: "likely", reasonCodes: ["pdf.behavior-native-id"] },
}

const PDF_REGION_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "unknown", reasonCodes: ["manual-region"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["normalized-region"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["pdf.page-index"] },
  relocation: { level: "likely", reasonCodes: ["pdf.page-signature"] },
}

function copyReliability(value: OfficeObjectReliability): OfficeObjectReliability {
  return {
    semantic: { ...value.semantic, reasonCodes: [...value.semantic.reasonCodes] },
    boundary: { ...value.boundary, reasonCodes: [...value.boundary.reasonCodes] },
    hierarchy: { ...value.hierarchy, reasonCodes: [...value.hierarchy.reasonCodes] },
    relocation: { ...value.relocation, reasonCodes: [...value.relocation.reasonCodes] },
  }
}

function pdfTextReliability(text: string): OfficeObjectReliability {
  if (text.length <= TEXT_QUOTE_LIMIT) return copyReliability(PDF_TEXT_RELIABILITY)
  return {
    ...copyReliability(PDF_TEXT_RELIABILITY),
    relocation: { level: "uncertain", reasonCodes: ["pdf.text-hash-only"] },
  }
}

function textContentHash(text: string): string {
  return fnv1a32(text)
}

function hasMatchingTextFingerprint(text: string, fingerprint?: OfficeObjectFingerprint): boolean {
  if (!fingerprint) return false
  if (fingerprint.exactText !== undefined && text !== fingerprint.exactText) return false
  if (fingerprint.contentHash?.startsWith("fnv1a32:") && textContentHash(text) !== fingerprint.contentHash) return false
  return fingerprint.exactText !== undefined || fingerprint.contentHash?.startsWith("fnv1a32:") === true
}

function hasVerifiableTextFingerprint(fingerprint?: OfficeObjectFingerprint): boolean {
  return fingerprint?.exactText !== undefined || fingerprint?.contentHash?.startsWith("fnv1a32:") === true
}

function sameRevision(left: OfficeDocumentRevision, right: OfficeDocumentRevision): boolean {
  return left.format === right.format
    && left.documentId === right.documentId
    && left.revision === right.revision
}

function assertContext(context: PdfReferenceContext): void {
  if (context.revision.format !== "pdf") throw new TypeError("PDF reference context requires format=pdf")
  if (!context.revision.documentId.trim()) throw new TypeError("PDF documentId must not be empty")
  if (!context.revision.revision.trim()) throw new TypeError("PDF revision must not be empty")
  if (context.document.pageCount !== context.document.pages.length) {
    throw new RangeError("PDF render document pageCount must match its page metadata")
  }
  context.document.pages.forEach((page, pageIndex) => {
    if (page.index !== pageIndex || !Number.isFinite(page.width) || page.width <= 0
      || !Number.isFinite(page.height) || page.height <= 0) {
      throw new RangeError("PDF page metadata must use ordered indices and positive finite dimensions")
    }
  })
}

function pageAt(context: PdfReferenceContext, pageIndex: number): PdfRenderPageInfo | undefined {
  return Number.isSafeInteger(pageIndex) && pageIndex >= 0 ? context.document.pages[pageIndex] : undefined
}

function pageSignatureMatches(reference: PdfOfficeReference, page: PdfRenderPageInfo): boolean {
  const keys = reference.fingerprint?.ancestorKeys ?? []
  const pageKeys = keys.filter(isPdfPageSignature)
  if (!pageKeys.length) return false
  if (!pageKeys.filter((key) => !isPdfPageContentSignature(key)).every((key) => key === pdfPageBaseSignature(page))) {
    return false
  }
  const contentKeys = pageKeys.filter(isPdfPageContentSignature)
  if (!contentKeys.length) return true
  const current = pdfPageContentSignature(page)
  return !!current && contentKeys.includes(current)
}

function assertPageRegion(region: Extract<ManualRegionLocator, { space: "page" }>, pageCount: number): void {
  if (!Number.isSafeInteger(region.pageIndex) || region.pageIndex < 0 || region.pageIndex >= pageCount) {
    throw new RangeError("PDF region pageIndex is outside the current document")
  }
  const { x, y, width, height } = region.rect
  if (![x, y, width, height].every(Number.isFinite)) throw new TypeError("PDF region coordinates must be finite")
  if (x < 0 || y < 0 || width <= 0 || height <= 0 || x + width > 1 || y + height > 1) {
    throw new RangeError("PDF region must be a positive rectangle inside normalized page space")
  }
}

export function normalizePdfReferenceRect(page: PdfRenderPageInfo, rect: PdfRenderRect): NormalizedRect | undefined {
  if (page.width <= 0 || page.height <= 0) return undefined
  const left = Math.max(0, Math.min(page.width, rect.x))
  const top = Math.max(0, Math.min(page.height, rect.y))
  const right = Math.max(left, Math.min(page.width, rect.x + rect.width))
  const bottom = Math.max(top, Math.min(page.height, rect.y + rect.height))
  if (right <= left || bottom <= top) return undefined
  return {
    x: left / page.width,
    y: top / page.height,
    width: (right - left) / page.width,
    height: (bottom - top) / page.height,
  }
}

function boundingRegion(
  page: PdfRenderPageInfo,
  rects: readonly PdfRenderRect[] | undefined,
): Extract<ManualRegionLocator, { space: "page" }> | undefined {
  const normalized = rects?.map((rect) => normalizePdfReferenceRect(page, rect)).filter((rect): rect is NormalizedRect => !!rect)
  if (!normalized?.length) return undefined
  const left = Math.min(...normalized.map((rect) => rect.x))
  const top = Math.min(...normalized.map((rect) => rect.y))
  const right = Math.max(...normalized.map((rect) => rect.x + rect.width))
  const bottom = Math.max(...normalized.map((rect) => rect.y + rect.height))
  return {
    space: "page",
    pageIndex: page.index,
    rect: { x: left, y: top, width: right - left, height: bottom - top },
  }
}

function textFingerprint(text: string, page: PdfRenderPageInfo): OfficeObjectFingerprint {
  return {
    ...(text.length <= TEXT_QUOTE_LIMIT ? { exactText: text } : {}),
    contentHash: textContentHash(text),
    ancestorKeys: pdfPageAncestorKeys(page),
  }
}

export function createPdfTextReferenceDraft(
  context: PdfReferenceContext,
  selection: PdfTextReferenceSelection,
): PdfOfficeReferenceDraft {
  assertContext(context)
  const page = pageAt(context, selection.range.pageIndex)
  if (!page) throw new RangeError("PDF text range pageIndex is outside the current document")
  if (!Number.isSafeInteger(selection.range.charIndex) || selection.range.charIndex < 0) {
    throw new TypeError("PDF charIndex must be a non-negative safe integer")
  }
  if (!Number.isSafeInteger(selection.range.charCount) || selection.range.charCount <= 0) {
    throw new TypeError("PDF charCount must be a positive safe integer")
  }
  if (!selection.text) throw new RangeError("PDF text reference requires non-empty selected text")
  const fallbackRegion = boundingRegion(page, selection.rects)
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "text-range",
    source: "native",
    locator: {
      type: "format",
      format: "pdf",
      value: {
        kind: "text-range",
        pageIndex: selection.range.pageIndex,
        charIndex: selection.range.charIndex,
        charCount: selection.range.charCount,
      },
    },
    fingerprint: textFingerprint(selection.text, page),
    ...(fallbackRegion ? { fallbackRegion } : {}),
    reliability: pdfTextReliability(selection.text),
  }
}

function pageAncestorKeys(page: PdfRenderPageInfo): readonly string[] {
  return pdfPageAncestorKeys(page)
}

function nativeObjectReliability(
  objectType: "link" | "form" | "annotation",
): OfficeObjectReliability {
  if (objectType === "link") return copyReliability(PDF_LINK_RELIABILITY)
  if (objectType === "form") return copyReliability(PDF_FORM_RELIABILITY)
  return copyReliability(PDF_ANNOTATION_RELIABILITY)
}

function nativeObjectKind(objectType: "link" | "form" | "annotation"): PdfReferenceKindValue {
  if (objectType === "link") return "link"
  if (objectType === "form") return "form-control"
  return "annotation"
}

type PdfReferenceKindValue = NonNullable<PdfOfficeReference["kind"]>

/**
 * Create a reference to a native PDF link, form widget or annotation.
 *
 * The annotation rectangle is normalized against the current page, so a link
 * selected in object mode keeps a page region fallback even when the native id
 * cannot be resolved after a document update.
 */
export function createPdfNativeObjectReferenceDraft(
  context: PdfReferenceContext,
  annotation: PdfRenderAnnotation,
): PdfOfficeReferenceDraft {
  assertContext(context)
  const page = pageAt(context, annotation.pageIndex)
  if (!page) throw new RangeError("PDF annotation pageIndex is outside the current document")
  const id = annotation.id.trim()
  if (!id) throw new TypeError("PDF native objects require a non-empty annotation id")
  const rect = normalizePdfReferenceRect(page, annotation.rect)
  if (!rect) throw new RangeError("PDF annotation rectangle is outside the current page")
  const objectType = pdfNativeObjectType(annotation)
  const region = { space: "page" as const, pageIndex: annotation.pageIndex, rect }
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: nativeObjectKind(objectType),
    source: "native",
    locator: {
      type: "format",
      format: "pdf",
      value: {
        kind: "native-object",
        pageIndex: annotation.pageIndex,
        objectType,
        nativeId: id,
      },
    },
    fingerprint: {
      ...(annotation.url ? { contentHash: textContentHash(annotation.url) } : {}),
      ...(annotation.fieldName ? { objectName: annotation.fieldName } : {}),
      ...(annotation.contents ? { exactText: annotation.contents.slice(0, TEXT_QUOTE_LIMIT) } : {}),
      ancestorKeys: [
        ...pageAncestorKeys(page),
        `pdf.native-object:${objectType}:${id}`,
      ],
    },
    fallbackRegion: region,
    reliability: nativeObjectReliability(objectType),
  }
}

function visualReliability(detection: PdfVisualObjectDetection): OfficeObjectReliability {
  return {
    semantic: { ...detection.reliability.semantic, reasonCodes: [...detection.reliability.semantic.reasonCodes] },
    boundary: { ...detection.reliability.boundary, reasonCodes: [...detection.reliability.boundary.reasonCodes] },
    hierarchy: { ...detection.reliability.hierarchy, reasonCodes: [...detection.reliability.hierarchy.reasonCodes] },
    relocation: { ...detection.reliability.relocation, reasonCodes: [...detection.reliability.relocation.reasonCodes] },
  }
}

/**
 * Create a reference from a visual recognizer result.
 *
 * The visual locator carries the provider id, so two providers with different
 * models never shadow each other's objects, and the normalized region stays as
 * the page-area fallback.
 */
export function createPdfVisualObjectReferenceDraft(
  context: PdfReferenceContext,
  providerId: string,
  pageIndex: number,
  detection: PdfVisualObjectDetection,
): PdfOfficeReferenceDraft {
  assertContext(context)
  if (!providerId.trim()) throw new TypeError("PDF visual objects require a provider id")
  const page = pageAt(context, pageIndex)
  if (!page) throw new RangeError("PDF visual object pageIndex is outside the current document")
  const { x, y, width, height } = detection.region
  if (![x, y, width, height].every(Number.isFinite)) {
    throw new TypeError("PDF visual object coordinates must be finite")
  }
  if (x < 0 || y < 0 || width <= 0 || height <= 0 || x + width > 1 || y + height > 1) {
    throw new RangeError("PDF visual object must be a positive rectangle inside normalized page space")
  }
  const kind = detection.kind
  return {
    schemaVersion: 1,
    document: context.revision,
    kind,
    source: "visual",
    locator: {
      type: "format",
      format: "pdf",
      value: {
        kind: "visual-object",
        pageIndex,
        providerId,
        objectId: detection.objectId,
        region: detection.region,
      },
    },
    fingerprint: {
      ...(detection.content?.text
        ? { contentHash: textContentHash(detection.content.text) }
        : {}),
      ...(detection.label ? { objectName: detection.label } : {}),
      ancestorKeys: [
        ...pageAncestorKeys(page),
        `pdf.visual-object:${providerId}:${detection.objectId}`,
      ],
    },
    fallbackRegion: { space: "page", pageIndex, rect: detection.region },
    reliability: visualReliability(detection),
  } as PdfOfficeReferenceDraft
}

/** Create a stage 4 invisible-behavior reference (link action, form rule, …). */
export function createPdfBehaviorReferenceDraft(
  context: PdfReferenceContext,
  record: PdfBehaviorRecord,
): PdfOfficeReferenceDraft {
  assertContext(context)
  const current = (context.behaviors ?? []).find((candidate) => candidate.instanceId === record.instanceId)
  if (!current) throw new RangeError("PDF behavior is not present in the reference context")
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: PDF_BEHAVIOR_REFERENCE_KIND[current.behavior],
    source: "native",
    locator: {
      type: "format",
      format: "pdf",
      value: pdfBehaviorLocatorValue(current),
    },
    fingerprint: {
      ancestorKeys: [
        `pdf.behavior:${current.behavior}`,
        `pdf.behavior-instance:${current.instanceId}`,
      ],
    },
    reliability: copyReliability(PDF_BEHAVIOR_RELIABILITY),
  } as PdfOfficeReferenceDraft
}

export function createPdfPageReferenceDraft(
  context: PdfReferenceContext,
  pageIndex: number,
): PdfOfficeReferenceDraft {
  assertContext(context)
  const page = pageAt(context, pageIndex)
  if (!page) throw new RangeError("PDF pageIndex is outside the current document")
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "page",
    source: "native",
    locator: { type: "format", format: "pdf", value: { kind: "page", pageIndex } },
    fingerprint: { ancestorKeys: pdfPageAncestorKeys(page) },
    reliability: copyReliability(PDF_PAGE_RELIABILITY),
  }
}

export function createPdfRegionReferenceDraft(
  context: PdfReferenceContext,
  region: Extract<ManualRegionLocator, { space: "page" }>,
): PdfOfficeReferenceDraft {
  assertContext(context)
  assertPageRegion(region, context.document.pageCount)
  const page = context.document.pages[region.pageIndex]!
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "region",
    source: "manual",
    locator: { type: "manual-region", format: "pdf", value: region },
    fingerprint: { ancestorKeys: pdfPageAncestorKeys(page) },
    fallbackRegion: region,
    reliability: copyReliability(PDF_REGION_RELIABILITY),
  }
}

function draftFromReference(reference: PdfOfficeReference): PdfOfficeReferenceDraft {
  const { referenceId: _referenceId, ...draft } = reference
  return draft
}

function currentFormatReference(
  context: PdfReferenceContext,
  reference: PdfOfficeReference,
  locator: PdfObjectLocator,
  fallbackRegion?: Extract<ManualRegionLocator, { space: "page" }>,
  fingerprint: OfficeObjectFingerprint | undefined = reference.fingerprint,
): PdfOfficeReference {
  const { fallbackRegion: _oldFallback, ...base } = reference
  return {
    ...base,
    document: context.revision,
    locator: { type: "format", format: "pdf", value: locator },
    ...(fingerprint ? { fingerprint } : {}),
    ...(fallbackRegion ? { fallbackRegion } : {}),
  } as PdfOfficeReference
}

function currentRegionReference(
  context: PdfReferenceContext,
  reference: PdfOfficeReference,
): PdfOfficeReference {
  return { ...reference, document: context.revision } as PdfOfficeReference
}

function resolveDescriptorLabels(context: PdfReferenceContext): PdfReferenceDescriptorLabels {
  if (context.descriptors) return context.descriptors
  const messages = officeMessagesEn.common
  return {
    document: messages.descriptors.document,
    page: (pageIndex) => formatOfficeMessage(messages.descriptors.page, { number: pageIndex + 1 }),
    selectedText: messages.selectedText,
    textRange: messages.descriptors.textRange,
    selectedRegion: messages.descriptors.selectedRegion,
    textLabel: (text) => formatOfficeMessage(messages.descriptors.textLabel, { text }),
    regionOnPage: (pageIndex) => formatOfficeMessage(messages.descriptors.regionOnPage, { number: pageIndex + 1 }),
  }
}

function textDescriptor(
  context: PdfReferenceContext,
  reference: PdfOfficeReference,
  text: string,
  page: PdfRenderPageInfo,
  rects?: readonly PdfRenderRect[],
): OfficeObjectDescriptor {
  const locator = reference.locator.type === "format" ? reference.locator.value : undefined
  const pageIndex = locator?.kind === "text-range" ? locator.pageIndex : page.index
  const fragments = rects
    ?.map((rect) => normalizePdfReferenceRect(page, rect))
    .filter((rect): rect is NormalizedRect => !!rect)
    .map((rect) => ({ container: { space: "page" as const, pageIndex }, rect }))
  const fallbackFragments = !fragments?.length && reference.fallbackRegion
    ? [{
        container: { space: "page" as const, pageIndex: reference.fallbackRegion.pageIndex },
        rect: reference.fallbackRegion.rect,
      }]
    : undefined
  const labelText = text.replace(/\s+/gu, " ").trim()
  const labels = resolveDescriptorLabels(context)
  return {
    objectId: `pdf:text:${pageIndex}:${locator?.kind === "text-range" ? `${locator.charIndex}:${locator.charCount}` : "unknown"}`,
    draft: draftFromReference(reference),
    label: labelText
      ? labels.textLabel(labelText.slice(0, 80))
      : labels.textRange,
    path: [
      { kind: "document", label: labels.document },
      { kind: "page", label: labels.page(pageIndex) },
      { kind: "text-range", label: labels.selectedText },
    ],
    childrenState: "none",
    content: { text },
    ...((fragments?.length || fallbackFragments?.length) ? {
      visual: {
        fragments: fragments?.length ? fragments : fallbackFragments!,
        layoutVersion: reference.document.revision,
      },
    } : {}),
  }
}

function simpleDescriptor(context: PdfReferenceContext, reference: PdfOfficeReference): OfficeObjectDescriptor {
  const region = reference.locator.type === "manual-region" ? reference.locator.value : undefined
  const locator = reference.locator.type === "format" ? reference.locator.value : undefined
  const pageIndex = region?.pageIndex ?? (locator?.kind === "page" ? locator.pageIndex : 0)
  const labels = resolveDescriptorLabels(context)
  return {
    objectId: region ? `pdf:region:${pageIndex}:${JSON.stringify(region.rect)}` : `pdf:page:${pageIndex}`,
    draft: draftFromReference(reference),
    label: region ? labels.regionOnPage(pageIndex) : labels.page(pageIndex),
    path: [
      { kind: "document", label: labels.document },
      { kind: "page", label: labels.page(pageIndex) },
      ...(region ? [{ kind: "region" as const, label: labels.selectedRegion }] : []),
    ],
    childrenState: region ? "none" : "unknown",
    ...(region ? {
      visual: {
        fragments: [{ container: { space: "page" as const, pageIndex }, rect: region.rect }],
        layoutVersion: reference.document.revision,
      },
    } : {}),
  }
}

function throwIfAborted(signal?: AbortSignal): void {
  if (!signal?.aborted) return
  if (typeof DOMException !== "undefined") throw new DOMException("PDF reference resolution aborted", "AbortError")
  const error = new Error("PDF reference resolution aborted")
  error.name = "AbortError"
  throw error
}

async function ensureCurrentPageContentHash(
  context: PdfReferenceContext,
  page: PdfRenderPageInfo,
  reference: PdfOfficeReference,
  signal?: AbortSignal,
): Promise<void> {
  const keys = reference.fingerprint?.ancestorKeys ?? []
  if (!keys.some(isPdfPageContentSignature) || page.contentHash) return
  const getPageContentHash = context.runtime.getPageContentHash
  if (typeof getPageContentHash !== "function") return
  try {
    const hash = await getPageContentHash.call(context.runtime, context.document, page.index, { signal })
    if (typeof hash === "string" && hash) page.contentHash = hash
  } catch (error) {
    throwIfAborted(signal)
    if ((error as Error)?.name === "AbortError") throw error
    throw error
  }
}

async function readTextAtLocator(
  context: PdfReferenceContext,
  locator: Extract<PdfObjectLocator, { kind: "text-range" }>,
  signal?: AbortSignal,
): Promise<string | undefined> {
  if (!pageAt(context, locator.pageIndex)) return undefined
  try {
    const [text] = await context.runtime.getTextSlices(context.document, [{
      pageIndex: locator.pageIndex,
      charIndex: locator.charIndex,
      charCount: locator.charCount,
    }], { signal })
    return text
  } catch (error) {
    throwIfAborted(signal)
    if ((error as Error)?.name === "AbortError") throw error
    throw error
  }
}

function hitScore(hit: PdfSearchHit, fingerprint?: OfficeObjectFingerprint): number {
  let score = 4
  if (fingerprint?.prefixText && hit.before.endsWith(fingerprint.prefixText)) score += 2
  if (fingerprint?.suffixText && hit.after.startsWith(fingerprint.suffixText)) score += 2
  return score
}

function uniqueSearchHits(hits: readonly PdfSearchHit[]): PdfSearchHit[] {
  const seen = new Set<string>()
  return hits.filter((hit) => {
    const key = `${hit.pageIndex}:${hit.charIndex}:${hit.charCount}`
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}

function referenceFromHit(
  context: PdfReferenceContext,
  reference: PdfOfficeReference,
  hit: PdfSearchHit,
): PdfOfficeReference {
  const page = context.document.pages[hit.pageIndex]!
  const prefixText = hit.before.slice(-64)
  const suffixText = hit.after.slice(0, 64)
  const fingerprint = {
    ...textFingerprint(hit.match, page),
    ...(prefixText ? { prefixText } : {}),
    ...(suffixText ? { suffixText } : {}),
  }
  return currentFormatReference(
    context,
    reference,
    { kind: "text-range", pageIndex: hit.pageIndex, charIndex: hit.charIndex, charCount: hit.charCount },
    boundingRegion(page, hit.rects),
    fingerprint,
  )
}

function nativeObjectDescriptor(
  context: PdfReferenceContext,
  reference: PdfOfficeReference,
  annotation: PdfRenderAnnotation,
  page: PdfRenderPageInfo,
): OfficeObjectDescriptor {
  const rect = normalizePdfReferenceRect(page, annotation.rect) ?? { x: 0, y: 0, width: 1, height: 1 }
  const objectType = pdfNativeObjectType(annotation)
  const kind = nativeObjectKind(objectType)
  const label = objectType === "link"
    ? "Link"
    : objectType === "form"
      ? `Form field${annotation.fieldName ? `: ${annotation.fieldName}` : ""}`
      : annotation.contents?.trim()
        ? `Annotation: ${annotation.contents.replace(/\s+/gu, " ").trim().slice(0, 80)}`
        : `Annotation ${annotation.id}`
  return {
    objectId: `pdf:native:${annotation.pageIndex}:${objectType}:${annotation.id}`,
    draft: draftFromReference(reference),
    label,
    path: [
      { kind: "document", label: "PDF" },
      { kind: "page", label: `Page ${annotation.pageIndex + 1}` },
      { kind, label },
    ],
    childrenState: "none",
    ...((annotation.url || annotation.fieldValue || annotation.contents) ? {
      content: {
        ...(annotation.url ? { text: annotation.url } : annotation.contents ? { text: annotation.contents } : {}),
        ...(annotation.fieldValue ? { value: annotation.fieldValue } : {}),
      },
    } : {}),
    visual: {
      fragments: [{ container: { space: "page", pageIndex: annotation.pageIndex }, rect }],
      layoutVersion: reference.document.revision,
    },
    facets: {
      nativeId: annotation.id,
      objectType,
      ...(annotation.url ? { url: annotation.url } : {}),
      ...(annotation.targetPageIndex === undefined ? {} : { targetPageIndex: annotation.targetPageIndex }),
      ...(annotation.fieldName ? { fieldName: annotation.fieldName } : {}),
      ...(annotation.fieldType ? { fieldType: annotation.fieldType } : {}),
      ...(annotation.readOnly === undefined ? {} : { readOnly: annotation.readOnly }),
    },
  }
}

function visualObjectDescriptor(
  reference: PdfOfficeReference,
  detection: PdfVisualObjectDetection,
  pageIndex: number,
): OfficeObjectDescriptor {
  const label = detection.label?.trim()
    || detection.content?.text?.replace(/\s+/gu, " ").trim().slice(0, 80)
    || detection.kind
  return {
    objectId: `pdf:visual:${pageIndex}:${detection.objectId}`,
    draft: draftFromReference(reference),
    label,
    path: [
      { kind: "document", label: "PDF" },
      { kind: "page", label: `Page ${pageIndex + 1}` },
      { kind: detection.kind, label },
    ],
    childrenState: detection.kind === "table" ? "available" : "none",
    ...(detection.content?.text ? { content: { text: detection.content.text } } : {}),
    visual: {
      fragments: [{ container: { space: "page", pageIndex }, rect: detection.region }],
      layoutVersion: reference.document.revision,
    },
    facets: {
      objectId: detection.objectId,
      ...(detection.parentObjectId ? { parentObjectId: detection.parentObjectId } : {}),
    },
  }
}

function pdfBehaviorDescriptor(
  reference: PdfOfficeReference,
  record: PdfBehaviorRecord,
): OfficeObjectDescriptor {
  return {
    objectId: `pdf:behavior:${record.instanceId}`,
    draft: draftFromReference(reference),
    label: record.label,
    path: record.path.map((entry) => ({ ...entry })),
    childrenState: "none",
    ...(record.content === undefined ? {} : { content: { ...record.content } }),
    facets: {
      ...record.facets,
      behavior: record.behavior,
      instanceId: record.instanceId,
      referenceKind: PDF_BEHAVIOR_REFERENCE_KIND[record.behavior],
    },
  }
}

function referenceFromDraft(draft: PdfOfficeReferenceDraft): PdfOfficeReference {
  return { ...draft, referenceId: "descriptor" } as PdfOfficeReference
}

/** Describe a native link/form/annotation without committing a reference. */
export function describePdfNativeObject(
  context: PdfReferenceContext,
  annotation: PdfRenderAnnotation,
): OfficeObjectDescriptor {
  const draft = createPdfNativeObjectReferenceDraft(context, annotation)
  const page = pageAt(context, annotation.pageIndex)!
  return nativeObjectDescriptor(context, referenceFromDraft(draft), annotation, page)
}

/** Describe a visual recognizer result without committing a reference. */
export function describePdfVisualObject(
  context: PdfReferenceContext,
  providerId: string,
  pageIndex: number,
  detection: PdfVisualObjectDetection,
): OfficeObjectDescriptor {
  const draft = createPdfVisualObjectReferenceDraft(context, providerId, pageIndex, detection)
  return visualObjectDescriptor(referenceFromDraft(draft), detection, pageIndex)
}

/** Describe a stage 4 PDF behavior without committing a reference. */
export function describePdfBehavior(
  context: PdfReferenceContext,
  record: PdfBehaviorRecord,
): OfficeObjectDescriptor {
  return pdfBehaviorDescriptor(referenceFromDraft(createPdfBehaviorReferenceDraft(context, record)), record)
}

export async function resolvePdfReference(
  context: PdfReferenceContext,
  reference: OfficeObjectReference,
  signal?: AbortSignal,
): Promise<ResolveReferenceResult> {
  assertContext(context)
  throwIfAborted(signal)
  if (reference.document.format !== "pdf") return { status: "unsupported", reasonCode: "pdf.format-mismatch" }
  const pdfReference = reference as PdfOfficeReference
  if (pdfReference.document.documentId !== context.revision.documentId) {
    return { status: "not-found", reasonCode: "pdf.document-id-mismatch" }
  }

  const region = pdfReference.locator.type === "manual-region" ? pdfReference.locator.value : undefined
  const locator = pdfReference.locator.type === "format" ? pdfReference.locator.value : undefined
  const isCurrentRevision = sameRevision(pdfReference.document, context.revision)

  if (locator?.kind === "native-object") {
    if (context.annotations === undefined) {
      return { status: "unsupported", reasonCode: "pdf.annotation-evidence-unavailable" }
    }
    const page = pageAt(context, locator.pageIndex)
    if (!page) return { status: "not-found", reasonCode: "pdf.page-not-found" }
    const annotation = context.annotations.find((candidate) =>
      candidate.pageIndex === locator.pageIndex
      && candidate.id === locator.nativeId
      && pdfNativeObjectType(candidate) === locator.objectType)
    if (!annotation) return { status: "not-found", reasonCode: "pdf.native-object-not-found" }
    if (isCurrentRevision) {
      return { status: "exact", reference: pdfReference, descriptor: nativeObjectDescriptor(context, pdfReference, annotation, page) }
    }
    const relocated = currentFormatReference(context, pdfReference, locator)
    return {
      status: "relocated",
      reference: relocated,
      descriptor: nativeObjectDescriptor(context, relocated, annotation, page),
      reasonCodes: ["pdf.revision-changed", "pdf.annotation-id-match"],
    }
  }

  if (locator?.kind === "visual-object") {
    if (context.visualObjects === undefined) {
      return { status: "unsupported", reasonCode: "pdf.visual-provider-unavailable" }
    }
    const page = pageAt(context, locator.pageIndex)
    if (!page) return { status: "not-found", reasonCode: "pdf.page-not-found" }
    const detection = context.visualObjects.find((candidate) =>
      candidate.objectId === locator.objectId && candidate.pageIndex === locator.pageIndex)
    if (!detection) return { status: "not-found", reasonCode: "pdf.visual-object-not-found" }
    if (isCurrentRevision) {
      return { status: "exact", reference: pdfReference, descriptor: visualObjectDescriptor(pdfReference, detection, locator.pageIndex) }
    }
    const relocated = currentFormatReference(context, pdfReference, {
      ...locator,
      region: detection.region,
    }, { space: "page", pageIndex: locator.pageIndex, rect: detection.region })
    return {
      status: "relocated",
      reference: relocated,
      descriptor: visualObjectDescriptor(relocated, detection, locator.pageIndex),
      reasonCodes: ["pdf.revision-changed", "pdf.visual-object-id-match"],
    }
  }

  if (locator?.kind === "behavior") {
    if (context.behaviors === undefined) {
      return { status: "unsupported", reasonCode: "pdf.behavior-evidence-unavailable" }
    }
    const record = context.behaviors.find((candidate) => candidate.instanceId === locator.nativeId)
    if (!record || record.behavior !== locator.behavior) {
      return { status: "not-found", reasonCode: "pdf.behavior-not-found" }
    }
    if (isCurrentRevision) {
      return { status: "exact", reference: pdfReference, descriptor: pdfBehaviorDescriptor(pdfReference, record) }
    }
    const relocated = currentFormatReference(context, pdfReference, pdfBehaviorLocatorValue(record))
    return {
      status: "relocated",
      reference: relocated,
      descriptor: pdfBehaviorDescriptor(relocated, record),
      reasonCodes: ["pdf.revision-changed", "pdf.behavior-native-id"],
    }
  }
  if (region || locator?.kind === "page") {
    const pageIndex = region ? region.pageIndex : locator?.kind === "page" ? locator.pageIndex : -1
    const page = pageAt(context, pageIndex)
    if (!page) return { status: "not-found", reasonCode: "pdf.page-not-found" }
    if (sameRevision(pdfReference.document, context.revision)) {
      return { status: "exact", reference: pdfReference, descriptor: simpleDescriptor(context, pdfReference) }
    }
    if (!pdfReference.fingerprint?.ancestorKeys?.some(isPdfPageSignature)) {
      return { status: "unsupported", reasonCode: "pdf.page-signature-unavailable" }
    }
    await ensureCurrentPageContentHash(context, page, pdfReference, signal)
    if (!pageSignatureMatches(pdfReference, page)) {
      return { status: "not-found", reasonCode: "pdf.page-signature-changed" }
    }
    const relocated = region
      ? currentRegionReference(context, pdfReference)
      : currentFormatReference(context, pdfReference, locator!)
    return {
      status: "relocated",
      reference: relocated,
      descriptor: simpleDescriptor(context, relocated),
      reasonCodes: ["pdf.revision-changed", "pdf.page-signature-match"],
    }
  }

  if (!locator || locator.kind !== "text-range") {
    return { status: "unsupported", reasonCode: "pdf.reference-kind-unsupported" }
  }
  const currentText = await readTextAtLocator(context, locator, signal)
  if (currentText !== undefined && (
    hasMatchingTextFingerprint(currentText, pdfReference.fingerprint)
    || isCurrentRevision && !hasVerifiableTextFingerprint(pdfReference.fingerprint)
  )) {
    const page = context.document.pages[locator.pageIndex]!
    if (isCurrentRevision) {
      return {
        status: "exact",
        reference: pdfReference,
        descriptor: textDescriptor(context, pdfReference, currentText, page),
      }
    }
    const relocated = currentFormatReference(
      context,
      pdfReference,
      locator,
      undefined,
      textFingerprint(currentText, page),
    )
    return {
      status: "relocated",
      reference: relocated,
      descriptor: textDescriptor(context, relocated, currentText, page),
      reasonCodes: ["pdf.revision-changed", "pdf.locator-stable"],
    }
  }
  if (isCurrentRevision) {
    return { status: "not-found", reasonCode: "pdf.text-range-invalid" }
  }
  const quote = pdfReference.fingerprint?.exactText
  if (!quote) return { status: "unsupported", reasonCode: "pdf.text-quote-unavailable" }

  let searchHits: readonly PdfSearchHit[]
  try {
    searchHits = await context.runtime.search(context.document, quote, { signal })
  } catch (error) {
    throwIfAborted(signal)
    if ((error as Error)?.name === "AbortError") throw error
    throw error
  }
  const exactHits = uniqueSearchHits(searchHits.filter((hit) =>
    hit.match === quote && !!pageAt(context, hit.pageIndex),
  ))
  if (exactHits.length === 0) return { status: "not-found", reasonCode: "pdf.text-not-found" }
  const bestScore = Math.max(...exactHits.map((hit) => hitScore(hit, pdfReference.fingerprint)))
  const bestHits = exactHits.filter((hit) => hitScore(hit, pdfReference.fingerprint) === bestScore)
  if (bestHits.length > 1) {
    return {
      status: "ambiguous",
      candidates: bestHits.map((hit) => {
        const candidate = referenceFromHit(context, pdfReference, hit)
        return textDescriptor(context, candidate, hit.match, context.document.pages[hit.pageIndex]!, hit.rects)
      }),
      reasonCodes: ["pdf.multiple-text-matches"],
    }
  }
  const hit = bestHits[0]!
  const relocated = referenceFromHit(context, pdfReference, hit)
  return {
    status: "relocated",
    reference: relocated,
    descriptor: textDescriptor(context, relocated, hit.match, context.document.pages[hit.pageIndex]!, hit.rects),
    reasonCodes: ["pdf.revision-changed", "pdf.text-quote-match"],
  }
}

export async function describePdfReference(
  context: PdfReferenceContext,
  reference: OfficeObjectReference,
  signal?: AbortSignal,
): Promise<OfficeObjectDescriptor | undefined> {
  const result = await resolvePdfReference(context, reference, signal)
  return result.status === "exact" || result.status === "relocated" ? result.descriptor : undefined
}
