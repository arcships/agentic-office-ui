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
import type { PageTextSlice } from "@embedpdf/models"
import type {
  PdfRenderDocument,
  PdfRenderPageInfo,
  PdfRenderRect,
  PdfRenderRuntime,
  PdfSearchHit,
} from "./pdf-render-runtime"

export type PdfOfficeReference = Extract<
  OfficeObjectReference,
  { document: { format: "pdf" } }
>

export type PdfOfficeReferenceDraft = Extract<
  OfficeObjectReferenceDraft,
  { document: { format: "pdf" } }
>

export interface PdfReferenceContext {
  revision: OfficeDocumentRevision & { format: "pdf" }
  document: PdfRenderDocument
  runtime: Pick<PdfRenderRuntime, "getTextSlices" | "search">
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
  let hash = 0x811c9dc5
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index)
    hash ^= code & 0xff
    hash = Math.imul(hash, 0x01000193)
    hash ^= code >>> 8
    hash = Math.imul(hash, 0x01000193)
  }
  return `fnv1a32:${text.length}:${(hash >>> 0).toString(16).padStart(8, "0")}`
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

function pageSignature(page: PdfRenderPageInfo): string {
  return `pdf-page-v1:${page.width}:${page.height}:${page.rotation}`
}

function pageSignatureMatches(reference: PdfOfficeReference, page: PdfRenderPageInfo): boolean {
  return reference.fingerprint?.ancestorKeys?.includes(pageSignature(page)) === true
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
    ancestorKeys: [pageSignature(page)],
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
    fingerprint: { ancestorKeys: [pageSignature(page)] },
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
    fingerprint: { ancestorKeys: [pageSignature(page)] },
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

function textDescriptor(
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
  return {
    objectId: `pdf:text:${pageIndex}:${locator?.kind === "text-range" ? `${locator.charIndex}:${locator.charCount}` : "unknown"}`,
    draft: draftFromReference(reference),
    label: labelText ? `Text: ${labelText.slice(0, 80)}` : "Text range",
    path: [
      { kind: "document", label: "Document" },
      { kind: "page", label: `Page ${pageIndex + 1}` },
      { kind: "text-range", label: "Selected text" },
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

function simpleDescriptor(reference: PdfOfficeReference): OfficeObjectDescriptor {
  const region = reference.locator.type === "manual-region" ? reference.locator.value : undefined
  const locator = reference.locator.type === "format" ? reference.locator.value : undefined
  const pageIndex = region?.pageIndex ?? (locator?.kind === "page" ? locator.pageIndex : 0)
  return {
    objectId: region ? `pdf:region:${pageIndex}:${JSON.stringify(region.rect)}` : `pdf:page:${pageIndex}`,
    draft: draftFromReference(reference),
    label: region ? `Region on page ${pageIndex + 1}` : `Page ${pageIndex + 1}`,
    path: [
      { kind: "document", label: "Document" },
      { kind: "page", label: `Page ${pageIndex + 1}` },
      ...(region ? [{ kind: "region" as const, label: "Selected region" }] : []),
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
  if (region || locator?.kind === "page") {
    const pageIndex = region ? region.pageIndex : locator?.kind === "page" ? locator.pageIndex : -1
    const page = pageAt(context, pageIndex)
    if (!page) return { status: "not-found", reasonCode: "pdf.page-not-found" }
    if (sameRevision(pdfReference.document, context.revision)) {
      return { status: "exact", reference: pdfReference, descriptor: simpleDescriptor(pdfReference) }
    }
    if (!pdfReference.fingerprint?.ancestorKeys?.some((key) => key.startsWith("pdf-page-v1:"))) {
      return { status: "unsupported", reasonCode: "pdf.page-signature-unavailable" }
    }
    if (!pageSignatureMatches(pdfReference, page)) {
      return { status: "not-found", reasonCode: "pdf.page-signature-changed" }
    }
    const relocated = region
      ? currentRegionReference(context, pdfReference)
      : currentFormatReference(context, pdfReference, locator!)
    return {
      status: "relocated",
      reference: relocated,
      descriptor: simpleDescriptor(relocated),
      reasonCodes: ["pdf.revision-changed", "pdf.page-signature-match"],
    }
  }

  if (!locator || locator.kind !== "text-range") {
    return { status: "unsupported", reasonCode: "pdf.reference-kind-unsupported" }
  }
  const currentText = await readTextAtLocator(context, locator, signal)
  const isCurrentRevision = sameRevision(pdfReference.document, context.revision)
  if (currentText !== undefined && (
    hasMatchingTextFingerprint(currentText, pdfReference.fingerprint)
    || isCurrentRevision && !hasVerifiableTextFingerprint(pdfReference.fingerprint)
  )) {
    const page = context.document.pages[locator.pageIndex]!
    if (isCurrentRevision) {
      return {
        status: "exact",
        reference: pdfReference,
        descriptor: textDescriptor(pdfReference, currentText, page),
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
      descriptor: textDescriptor(relocated, currentText, page),
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
        return textDescriptor(candidate, hit.match, context.document.pages[hit.pageIndex]!, hit.rects)
      }),
      reasonCodes: ["pdf.multiple-text-matches"],
    }
  }
  const hit = bestHits[0]!
  const relocated = referenceFromHit(context, pdfReference, hit)
  return {
    status: "relocated",
    reference: relocated,
    descriptor: textDescriptor(relocated, hit.match, context.document.pages[hit.pageIndex]!, hit.rects),
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
