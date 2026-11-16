import {
  createPdfiumEngine as createPdfiumWorkerEngine,
} from "@embedpdf/engines/pdfium-worker-engine"
import {
  PdfErrorCode,
  Rotation,
  type PageTextSlice,
  type PdfDocumentObject,
  type PdfErrorReason,
  type PdfPageGeometry,
} from "@embedpdf/models"
import { pdfIntrinsicRectToCanonical } from "./pdf-coordinate-transform"
import { pdfPageTextHash } from "./pdf-page-signature"

export type PdfRotation = 0 | 90 | 180 | 270

export type PdfRenderErrorCode =
  | "ENCRYPTED"
  | "INVALID_RENDER_OPTIONS"
  | "PAGE_TOO_LARGE"

/** Stable error raised by the PDF render runtime before touching PDFium. */
export class PdfRenderError extends Error {
  readonly code: PdfRenderErrorCode

  constructor(code: PdfRenderErrorCode, message: string, options?: { cause?: unknown }) {
    super(message)
    if (options?.cause !== undefined) {
      (this as Error & { cause?: unknown }).cause = options.cause
    }
    this.name = "PdfRenderError"
    this.code = code
  }
}

export interface PdfPasswordRequest {
  /** 1-based password attempt number. */
  attempt: number
  /** Engine message, safe to surface to a password prompt. */
  message: string
}

export type PdfPasswordHandler = (
  request: PdfPasswordRequest,
) => string | undefined | Promise<string | undefined>

export interface PdfOpenDocumentOptions {
  signal?: AbortSignal
  /** Password for encrypted documents. */
  password?: string
  /** Invoked when the engine reports that a password is required. */
  onPassword?: PdfPasswordHandler
}

export interface PdfRenderPageInfo {
  /** Zero-based page index. */
  index: number
  width: number
  height: number
  rotation: PdfRotation
  /**
   * Digest of the page text, populated lazily by this runtime. Page and region
   * references include it in their signature when present so replacing a
   * page's content at the same size is detected as a changed page.
   */
  contentHash?: string
}

export interface PdfRenderDocument {
  id: string
  pageCount: number
  pages: readonly PdfRenderPageInfo[]
}

export interface PdfRenderRect {
  x: number
  y: number
  width: number
  height: number
}

export interface PdfSearchHit {
  kind: "pdf-text"
  pageIndex: number
  charIndex: number
  charCount: number
  before: string
  match: string
  after: string
  rects: readonly PdfRenderRect[]
}

export type PdfSearchStatus = "idle" | "searching" | "ready" | "error"

export interface PdfSearchState {
  status: PdfSearchStatus
  query: string
  matches: readonly PdfSearchHit[]
  activeIndex: number
  error?: {
    code: "SEARCH_FAILED" | "ACTIVATION_FAILED"
    message: string
  }
}

export type PdfSelectionState =
  | { kind: "none" }
  | {
      kind: "text"
      range: PageTextSlice
      text: string
      rects: readonly PdfRenderRect[]
    }
  | {
      kind: "unavailable"
      pageIndex: number
      reason: "empty" | "unicode-map"
    }

export interface PdfPageRenderOptions {
  scale: number
  rotation: PdfRotation
  dpr?: number
  signal?: AbortSignal
}

export interface PdfPageRenderParameters {
  scaleFactor: number
  rotation: PdfRotation
  dpr: number
  pixelWidth: number
  pixelHeight: number
}

/** Upper bounds applied before PDFium allocates a rendered page bitmap. */
export const DEFAULT_PDF_MAX_RENDER_PIXELS = 32 * 1024 * 1024
export const MAX_PDF_RENDER_SCALE = 16
export const MAX_PDF_RENDER_DPR = 4

function normalizePdfRotation(value: unknown, label: string): PdfRotation {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new PdfRenderError("INVALID_RENDER_OPTIONS", `${label} must be a finite number of degrees.`)
  }
  const normalized = ((Math.round(value / 90) * 90) % 360 + 360) % 360
  if (!Object.is(normalized, value)) {
    throw new PdfRenderError(
      "INVALID_RENDER_OPTIONS",
      `${label} must be one of 0, 90, 180 or 270 degrees.`,
    )
  }
  return normalized as PdfRotation
}

function normalizePixelBudget(value: number | undefined): number {
  if (value === undefined) return DEFAULT_PDF_MAX_RENDER_PIXELS
  return Number.isFinite(value) && value > 0
    ? Math.max(1, Math.floor(value))
    : DEFAULT_PDF_MAX_RENDER_PIXELS
}

/**
 * Validates public render input and enforces a pixel budget before the engine
 * allocates anything. Invalid values and requests beyond the budget are
 * rejected instead of silently producing an unusable canvas.
 */
export function resolvePdfPageRenderParameters(
  page: Pick<PdfRenderPageInfo, "width" | "height" | "rotation">,
  options: Pick<PdfPageRenderOptions, "scale" | "rotation" | "dpr">,
  maxPixels: number = DEFAULT_PDF_MAX_RENDER_PIXELS,
): PdfPageRenderParameters {
  const { width, height } = page
  if (!Number.isFinite(width) || width <= 0 || !Number.isFinite(height) || height <= 0) {
    throw new PdfRenderError("INVALID_RENDER_OPTIONS", "PDF page dimensions must be positive finite numbers.")
  }
  const scale = options.scale
  if (typeof scale !== "number" || !Number.isFinite(scale) || scale <= 0) {
    throw new PdfRenderError("INVALID_RENDER_OPTIONS", "PDF render scale must be a positive finite number.")
  }
  if (scale > MAX_PDF_RENDER_SCALE) {
    throw new PdfRenderError(
      "PAGE_TOO_LARGE",
      `PDF render scale ${scale} exceeds the maximum supported scale ${MAX_PDF_RENDER_SCALE}.`,
    )
  }
  const dpr = options.dpr ?? 1
  if (typeof dpr !== "number" || !Number.isFinite(dpr) || dpr < 1) {
    throw new PdfRenderError("INVALID_RENDER_OPTIONS", "PDF render dpr must be a finite number of at least 1.")
  }
  if (dpr > MAX_PDF_RENDER_DPR) {
    throw new PdfRenderError(
      "PAGE_TOO_LARGE",
      `PDF render dpr ${dpr} exceeds the maximum supported dpr ${MAX_PDF_RENDER_DPR}.`,
    )
  }
  const userRotation = normalizePdfRotation(options.rotation ?? 0, "PDF render rotation")
  const pageRotation = normalizePdfRotation(page.rotation, "PDF page rotation")
  const totalRotation = ((pageRotation + userRotation) % 360) as PdfRotation
  const quarterTurns = totalRotation / 90
  const displayWidth = quarterTurns % 2 ? height : width
  const displayHeight = quarterTurns % 2 ? width : height
  const pixelWidth = Math.ceil(displayWidth * scale * dpr)
  const pixelHeight = Math.ceil(displayHeight * scale * dpr)
  const pixels = pixelWidth * pixelHeight
  const budget = normalizePixelBudget(maxPixels)
  if (!Number.isFinite(pixels) || pixels > budget) {
    throw new PdfRenderError(
      "PAGE_TOO_LARGE",
      `PDF page render would allocate ${pixelWidth}x${pixelHeight} pixels, exceeding the ${budget} pixel budget.`,
    )
  }
  return { scaleFactor: scale, rotation: totalRotation, dpr, pixelWidth, pixelHeight }
}

export interface PdfThumbnailRenderOptions {
  scale?: number
  signal?: AbortSignal
}

/**
 * Public instance boundary used by PdfViewer. Each runtime owns one PDF Worker,
 * its open documents and its task queue. It never fetches the original PDF URL.
 */
export interface PdfRenderRuntime {
  readonly runtimeId: string
  openDocument(bytes: ArrayBuffer, options?: PdfOpenDocumentOptions): Promise<PdfRenderDocument>
  renderPage(
    document: PdfRenderDocument,
    pageIndex: number,
    options: PdfPageRenderOptions,
  ): Promise<Blob>
  renderThumbnail(
    document: PdfRenderDocument,
    pageIndex: number,
    options?: PdfThumbnailRenderOptions,
  ): Promise<Blob>
  search(
    document: PdfRenderDocument,
    query: string,
    options?: { signal?: AbortSignal },
  ): Promise<readonly PdfSearchHit[]>
  getPageTextGeometry(
    document: PdfRenderDocument,
    pageIndex: number,
    options?: { signal?: AbortSignal },
  ): Promise<PdfPageGeometry>
  getTextSlices(
    document: PdfRenderDocument,
    slices: readonly PageTextSlice[],
    options?: { signal?: AbortSignal },
  ): Promise<readonly string[]>
  closeDocument(document: PdfRenderDocument): Promise<void>
  /**
   * Returns (and caches) the page text digest. Optional so injected runtimes
   * that only render pages do not need to implement reference signatures.
   */
  getPageContentHash?(
    document: PdfRenderDocument,
    pageIndex: number,
    options?: { signal?: AbortSignal },
  ): Promise<string>
  dispose(): Promise<void>
}

export interface PdfRenderRuntimeConfig {
  /** Defaults to the PDFium WASM shipped by @arcships/vue-pdf. */
  wasmUrl?: string
  /** Bounds an engine operation, including initial Worker/WASM startup. */
  operationTimeoutMs?: number
  runtimeId?: string
  /** Pixel budget for one rendered page. Defaults to 32 megapixels. */
  maxRenderPixels?: number
  /** Default password for encrypted documents. */
  password?: string
  /** Default password hook for encrypted documents. */
  onPassword?: PdfPasswordHandler
  /** Maximum password prompts per open attempt. Defaults to 3. */
  maxPasswordAttempts?: number
}

/** URL of the PDFium binary emitted beside the public package entry. */
export const bundledPdfiumWasmUrl = new URL("./pdfium.wasm", import.meta.url).href

type PdfiumEngine = ReturnType<typeof createPdfiumWorkerEngine>

interface AbortablePdfTask<T> {
  toPromise(): Promise<T>
  abort(reason: PdfErrorReason): void
}

function createId(prefix: string): string {
  const suffix = typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`
  return `${prefix}-${suffix}`
}

function createAbortError(message = "PDF operation aborted."): Error {
  if (typeof DOMException !== "undefined") return new DOMException(message, "AbortError")
  const error = new Error(message)
  error.name = "AbortError"
  return error
}

function asRotation(degrees: PdfRotation): Rotation {
  switch (degrees) {
    case 90: return Rotation.Degree90
    case 180: return Rotation.Degree180
    case 270: return Rotation.Degree270
    default: return Rotation.Degree0
  }
}

function toDegrees(rotation: Rotation): PdfRotation {
  switch (rotation) {
    case Rotation.Degree90: return 90
    case Rotation.Degree180: return 180
    case Rotation.Degree270: return 270
    default: return 0
  }
}

function createOwnedPdfiumEngine(wasmUrl: string): PdfiumEngine {
  // The upstream factory creates its PDF and encoder Workers from private Blob
  // URLs but exposes neither URL for revocation. We deliberately leave
  // URL.createObjectURL untouched: patching the global can capture other
  // callers' URLs, breaks nested construction, and throws in read-only/CSP
  // environments. The two script URLs created here live until the document is
  // unloaded; Worker termination on dispose still releases their runtimes.
  return createPdfiumWorkerEngine(wasmUrl, {
    encoderPoolSize: 2,
    fontFallback: null,
  })
}

function readPdfErrorCode(error: unknown): unknown {
  if (!error || typeof error !== "object") return undefined
  const reason = (error as { reason?: unknown }).reason
  if (reason && typeof reason === "object" && "code" in reason) {
    return (reason as { code?: unknown }).code
  }
  return (error as { code?: unknown }).code
}

function isPdfPasswordError(error: unknown): boolean {
  if (readPdfErrorCode(error) === PdfErrorCode.Password) return true
  return error instanceof Error && /password/i.test(error.message)
}

async function runPdfTask<T>(
  task: AbortablePdfTask<T>,
  options: { signal?: AbortSignal; timeoutMs: number; label: string },
): Promise<T> {
  const { signal, timeoutMs, label } = options
  if (signal?.aborted) {
    task.abort({ code: PdfErrorCode.Cancelled, message: `${label} cancelled` })
    throw createAbortError()
  }

  let timeout: ReturnType<typeof setTimeout> | undefined
  let removeAbortListener: (() => void) | undefined
  const guard = new Promise<never>((_resolve, reject) => {
    const abort = () => {
      task.abort({ code: PdfErrorCode.Cancelled, message: `${label} cancelled` })
      reject(createAbortError())
    }
    signal?.addEventListener("abort", abort, { once: true })
    removeAbortListener = () => signal?.removeEventListener("abort", abort)
    timeout = setTimeout(() => {
      task.abort({ code: PdfErrorCode.Cancelled, message: `${label} timed out` })
      reject(new Error(`${label} timed out after ${timeoutMs} ms.`))
    }, timeoutMs)
  })

  try {
    return await Promise.race([task.toPromise(), guard])
  } finally {
    if (timeout) clearTimeout(timeout)
    removeAbortListener?.()
  }
}

class PdfiumRenderRuntime implements PdfRenderRuntime {
  readonly runtimeId: string
  private readonly wasmUrl: string
  private readonly operationTimeoutMs: number
  private readonly maxRenderPixels: number
  private readonly defaultPassword: string | undefined
  private readonly defaultOnPassword: PdfPasswordHandler | undefined
  private readonly maxPasswordAttempts: number
  private engine: PdfiumEngine | undefined
  private readonly documents = new Map<string, PdfDocumentObject>()
  private readonly contentHashRequests = new Map<string, Promise<string>>()
  private documentSequence = 0
  private disposed = false

  constructor(config: PdfRenderRuntimeConfig) {
    this.runtimeId = config.runtimeId?.trim() || createId("pdf-runtime")
    this.wasmUrl = config.wasmUrl?.trim() || bundledPdfiumWasmUrl
    this.operationTimeoutMs = Math.max(1_000, config.operationTimeoutMs ?? 30_000)
    this.maxRenderPixels = normalizePixelBudget(config.maxRenderPixels)
    this.defaultPassword = config.password
    this.defaultOnPassword = config.onPassword
    this.maxPasswordAttempts = config.maxPasswordAttempts !== undefined && Number.isFinite(config.maxPasswordAttempts)
      ? Math.max(1, Math.floor(config.maxPasswordAttempts))
      : 3
  }

  private getEngine(): PdfiumEngine {
    if (this.disposed) throw new Error("PDF runtime has been disposed.")
    if (!this.engine) this.engine = createOwnedPdfiumEngine(this.wasmUrl)
    return this.engine
  }

  private resetFailedEngine(engine: PdfiumEngine): void {
    if (this.documents.size !== 0 || this.engine !== engine) return
    this.engine = undefined
    void engine.destroy().toPromise().catch(() => undefined)
  }

  private getNativeDocument(document: PdfRenderDocument): PdfDocumentObject {
    if (this.disposed) throw new Error("PDF runtime has been disposed.")
    const nativeDocument = this.documents.get(document.id)
    if (!nativeDocument) throw new Error("PDF document is no longer open.")
    return nativeDocument
  }

  async openDocument(
    bytes: ArrayBuffer,
    options: PdfOpenDocumentOptions = {},
  ): Promise<PdfRenderDocument> {
    if (!(bytes instanceof ArrayBuffer) || bytes.byteLength === 0) {
      throw new TypeError("PDF document bytes must be a non-empty ArrayBuffer.")
    }
    const engine = this.getEngine()
    let password = options.password ?? this.defaultPassword
    const onPassword = options.onPassword ?? this.defaultOnPassword
    let passwordAttempt = 0
    let nativeDocument: PdfDocumentObject | undefined

    for (;;) {
      const id = `${this.runtimeId}-document-${++this.documentSequence}`
      const task = engine.openDocumentBuffer(
        { id, content: bytes },
        password === undefined ? undefined : { password },
      )
      try {
        nativeDocument = await runPdfTask(task, {
          signal: options.signal,
          timeoutMs: this.operationTimeoutMs,
          label: "PDF Worker startup and document open",
        })
        break
      } catch (error) {
        if (isPdfPasswordError(error)) {
          if (!onPassword || passwordAttempt >= this.maxPasswordAttempts) {
            throw new PdfRenderError(
              "ENCRYPTED",
              "PDF document is encrypted and requires a password.",
              { cause: error },
            )
          }
          passwordAttempt += 1
          const nextPassword = await onPassword({
            attempt: passwordAttempt,
            message: "PDF document is encrypted and requires a password.",
          })
          if (options.signal?.aborted || this.disposed) throw createAbortError()
          if (typeof nextPassword !== "string" || !nextPassword) {
            throw new PdfRenderError(
              "ENCRYPTED",
              "PDF document is encrypted and requires a password.",
              { cause: error },
            )
          }
          password = nextPassword
          continue
        }
        this.resetFailedEngine(engine)
        throw error
      }
    }

    if (options.signal?.aborted || this.disposed) {
      void engine.closeDocument(nativeDocument).toPromise().catch(() => undefined)
      throw createAbortError()
    }
    if (!nativeDocument.pageCount || nativeDocument.pages.length !== nativeDocument.pageCount) {
      void engine.closeDocument(nativeDocument).toPromise().catch(() => undefined)
      throw new Error("PDF document contains no renderable pages.")
    }
    this.documents.set(nativeDocument.id, nativeDocument)
    return {
      id: nativeDocument.id,
      pageCount: nativeDocument.pageCount,
      pages: nativeDocument.pages.map((page) => ({
        index: page.index,
        width: page.size.width,
        height: page.size.height,
        rotation: toDegrees(page.rotation),
      })),
    }
  }

  async renderPage(
    document: PdfRenderDocument,
    pageIndex: number,
    options: PdfPageRenderOptions,
  ): Promise<Blob> {
    const pageInfo = document.pages[pageIndex]
    if (!pageInfo) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    // Validate all public input before the engine or document map is touched so
    // invalid/huge requests can never reach PDFium or allocate a bitmap.
    const parameters = resolvePdfPageRenderParameters(pageInfo, options, this.maxRenderPixels)
    const nativeDocument = this.getNativeDocument(document)
    const page = nativeDocument.pages[pageIndex]
    if (!page) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    const task = this.getEngine().renderPage(nativeDocument, page, {
      scaleFactor: parameters.scaleFactor,
      rotation: asRotation(parameters.rotation),
      dpr: parameters.dpr,
      imageType: "image/png",
      withAnnotations: true,
      withForms: true,
    })
    return runPdfTask(task, {
      signal: options.signal,
      timeoutMs: this.operationTimeoutMs,
      label: `PDF page ${pageIndex + 1} render`,
    })
  }

  async renderThumbnail(
    document: PdfRenderDocument,
    pageIndex: number,
    options: PdfThumbnailRenderOptions = {},
  ): Promise<Blob> {
    const pageInfo = document.pages[pageIndex]
    if (!pageInfo) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    const parameters = resolvePdfPageRenderParameters(
      pageInfo,
      { scale: options.scale ?? 0.2, rotation: 0 },
      this.maxRenderPixels,
    )
    const nativeDocument = this.getNativeDocument(document)
    const page = nativeDocument.pages[pageIndex]
    if (!page) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    const task = this.getEngine().renderThumbnail(nativeDocument, page, {
      scaleFactor: parameters.scaleFactor,
      dpr: 1,
      imageType: "image/png",
      withAnnotations: true,
    })
    return runPdfTask(task, {
      signal: options.signal,
      timeoutMs: this.operationTimeoutMs,
      label: `PDF page ${pageIndex + 1} thumbnail`,
    })
  }

  async getPageContentHash(
    document: PdfRenderDocument,
    pageIndex: number,
    options: { signal?: AbortSignal } = {},
  ): Promise<string> {
    const pageInfo = document.pages[pageIndex]
    if (!pageInfo) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    if (pageInfo.contentHash) return pageInfo.contentHash
    const key = `${document.id}:${pageIndex}`
    const existing = this.contentHashRequests.get(key)
    if (existing) return existing
    const request = (async () => {
      const nativeDocument = this.getNativeDocument(document)
      const page = nativeDocument.pages[pageIndex]
      if (!page) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
      const result = await runPdfTask(this.getEngine().getPageTextRuns(nativeDocument, page), {
        signal: options.signal,
        timeoutMs: this.operationTimeoutMs,
        label: `PDF page ${pageIndex + 1} content hash`,
      })
      const hash = pdfPageTextHash(result.runs.map((run) => run.text).join("\n"))
      const current = document.pages[pageIndex]
      if (current && this.documents.has(document.id)) current.contentHash = hash
      return hash
    })()
    this.contentHashRequests.set(key, request)
    try {
      return await request
    } finally {
      if (this.contentHashRequests.get(key) === request) this.contentHashRequests.delete(key)
    }
  }

  async search(
    document: PdfRenderDocument,
    query: string,
    options: { signal?: AbortSignal } = {},
  ): Promise<readonly PdfSearchHit[]> {
    const nativeDocument = this.getNativeDocument(document)
    const normalizedQuery = query.trim()
    if (!normalizedQuery) return []
    const task = this.getEngine().searchAllPages(nativeDocument, normalizedQuery)
    const result = await runPdfTask(task, {
      signal: options.signal,
      timeoutMs: this.operationTimeoutMs,
      label: "PDF text search",
    })
    return result.results.map((hit) => {
      const page = document.pages[hit.pageIndex]
      return {
        kind: "pdf-text" as const,
        pageIndex: hit.pageIndex,
        charIndex: hit.charIndex,
        charCount: hit.charCount,
        before: hit.context.before,
        match: hit.context.match,
        after: hit.context.after,
        rects: hit.rects.map((rect) => {
          const normalized = {
            x: rect.origin.x,
            y: rect.origin.y,
            width: rect.size.width,
            height: rect.size.height,
          }
          return page ? pdfIntrinsicRectToCanonical(page, normalized) : normalized
        }),
      }
    })
  }

  async getPageTextGeometry(
    document: PdfRenderDocument,
    pageIndex: number,
    options: { signal?: AbortSignal } = {},
  ): Promise<PdfPageGeometry> {
    const nativeDocument = this.getNativeDocument(document)
    const page = nativeDocument.pages[pageIndex]
    if (!page) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    return runPdfTask(this.getEngine().getPageGeometry(nativeDocument, page), {
      signal: options.signal,
      timeoutMs: this.operationTimeoutMs,
      label: `PDF page ${pageIndex + 1} text geometry`,
    })
  }

  async getTextSlices(
    document: PdfRenderDocument,
    slices: readonly PageTextSlice[],
    options: { signal?: AbortSignal } = {},
  ): Promise<readonly string[]> {
    const nativeDocument = this.getNativeDocument(document)
    if (!slices.length) return []
    return runPdfTask(this.getEngine().getTextSlices(nativeDocument, [...slices]), {
      signal: options.signal,
      timeoutMs: this.operationTimeoutMs,
      label: "PDF text extraction",
    })
  }

  async closeDocument(document: PdfRenderDocument): Promise<void> {
    const nativeDocument = this.documents.get(document.id)
    if (!nativeDocument) return
    this.documents.delete(document.id)
    for (const key of this.contentHashRequests.keys()) {
      if (key.startsWith(`${document.id}:`)) this.contentHashRequests.delete(key)
    }
    const engine = this.engine
    if (!engine) return
    try {
      await runPdfTask(engine.closeDocument(nativeDocument), {
        timeoutMs: this.operationTimeoutMs,
        label: "PDF document close",
      })
    } catch {
      // dispose() still terminates the owning Worker if graceful close fails.
    }
  }

  async dispose(): Promise<void> {
    if (this.disposed) return
    this.disposed = true
    this.documents.clear()
    this.contentHashRequests.clear()
    const engine = this.engine
    this.engine = undefined
    if (!engine) return
    try {
      await runPdfTask(engine.destroy(), {
        timeoutMs: Math.min(this.operationTimeoutMs, 5_000),
        label: "PDF runtime destroy",
      })
    } catch {
      // The runtime is already unusable and all local references are released.
    }
  }
}

export function createPdfRenderRuntime(config: PdfRenderRuntimeConfig = {}): PdfRenderRuntime {
  return new PdfiumRenderRuntime(config)
}
