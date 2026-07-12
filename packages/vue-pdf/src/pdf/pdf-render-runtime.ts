import {
  createPdfiumEngine as createPdfiumWorkerEngine,
} from "@embedpdf/engines/pdfium-worker-engine"
import {
  PdfErrorCode,
  Rotation,
  type PdfDocumentObject,
  type PdfErrorReason,
} from "@embedpdf/models"

export type PdfRotation = 0 | 90 | 180 | 270

export interface PdfRenderPageInfo {
  /** Zero-based page index. */
  index: number
  width: number
  height: number
  rotation: PdfRotation
}

/** A text item with bounding box from a PDF page, for text layer overlay. */
export interface PdfTextRectItem {
  content: string
  x: number
  y: number
  width: number
  height: number
  fontFamily: string
  fontSize: number
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
  pageIndex: number
  before: string
  match: string
  after: string
  rects: readonly PdfRenderRect[]
}

export interface PdfPageRenderOptions {
  scale: number
  rotation: PdfRotation
  dpr?: number
  signal?: AbortSignal
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
  openDocument(bytes: ArrayBuffer, options?: { signal?: AbortSignal }): Promise<PdfRenderDocument>
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
  closeDocument(document: PdfRenderDocument): Promise<void>
  /** Returns text items with bounding boxes for overlay / selection. */
  getPageTextRects(document: PdfRenderDocument, pageIndex: number): Promise<PdfTextRectItem[]>
  dispose(): Promise<void>
}

export interface PdfRenderRuntimeConfig {
  /** Defaults to the PDFium WASM shipped by @arcships/vue-pdf. */
  wasmUrl?: string
  /** Bounds an engine operation, including initial Worker/WASM startup. */
  operationTimeoutMs?: number
  runtimeId?: string
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

function createOwnedPdfiumEngine(wasmUrl: string): {
  engine: PdfiumEngine
  workerScriptUrls: readonly string[]
} {
  const originalCreateObjectUrl = URL.createObjectURL
  const workerScriptUrls: string[] = []
  let engine: PdfiumEngine | undefined

  // The upstream factory creates its PDF and encoder Workers synchronously from
  // private Blob URLs but does not expose those URLs for cleanup. Capture only
  // this factory call, then revoke the script URLs after Worker construction.
  URL.createObjectURL = ((object: Blob | MediaSource) => {
    const url = originalCreateObjectUrl.call(URL, object)
    workerScriptUrls.push(url)
    return url
  }) as typeof URL.createObjectURL
  try {
    engine = createPdfiumWorkerEngine(wasmUrl, {
      encoderPoolSize: 2,
      fontFallback: null,
    })
  } finally {
    URL.createObjectURL = originalCreateObjectUrl
    if (!engine) {
      for (const url of workerScriptUrls) URL.revokeObjectURL(url)
    }
  }
  return { engine, workerScriptUrls }
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
  private engine: PdfiumEngine | undefined
  private readonly workerScriptUrls = new Map<PdfiumEngine, readonly string[]>()
  private readonly documents = new Map<string, PdfDocumentObject>()
  private documentSequence = 0
  private disposed = false

  constructor(config: PdfRenderRuntimeConfig) {
    this.runtimeId = config.runtimeId?.trim() || createId("pdf-runtime")
    this.wasmUrl = config.wasmUrl?.trim() || bundledPdfiumWasmUrl
    this.operationTimeoutMs = Math.max(1_000, config.operationTimeoutMs ?? 30_000)
  }

  private getEngine(): PdfiumEngine {
    if (this.disposed) throw new Error("PDF runtime has been disposed.")
    if (!this.engine) {
      const owned = createOwnedPdfiumEngine(this.wasmUrl)
      this.engine = owned.engine
      this.workerScriptUrls.set(owned.engine, owned.workerScriptUrls)
    }
    return this.engine
  }

  private revokeWorkerScriptUrls(engine: PdfiumEngine): void {
    const urls = this.workerScriptUrls.get(engine) ?? []
    this.workerScriptUrls.delete(engine)
    for (const url of urls) URL.revokeObjectURL(url)
  }

  private getNativeDocument(document: PdfRenderDocument): PdfDocumentObject {
    if (this.disposed) throw new Error("PDF runtime has been disposed.")
    const nativeDocument = this.documents.get(document.id)
    if (!nativeDocument) throw new Error("PDF document is no longer open.")
    return nativeDocument
  }

  async openDocument(
    bytes: ArrayBuffer,
    options: { signal?: AbortSignal } = {},
  ): Promise<PdfRenderDocument> {
    const engine = this.getEngine()
    const id = `${this.runtimeId}-document-${++this.documentSequence}`
    const task = engine.openDocumentBuffer({ id, content: bytes })
    let nativeDocument: PdfDocumentObject
    try {
      nativeDocument = await runPdfTask(task, {
        signal: options.signal,
        timeoutMs: this.operationTimeoutMs,
        label: "PDF Worker startup and document open",
      })
    } catch (error) {
      // The upstream Worker can otherwise remain pending when WASM startup fails.
      if (this.documents.size === 0 && this.engine === engine) {
        this.engine = undefined
        void engine.destroy().toPromise()
          .catch(() => undefined)
          .finally(() => this.revokeWorkerScriptUrls(engine))
      }
      throw error
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
    const nativeDocument = this.getNativeDocument(document)
    const page = nativeDocument.pages[pageIndex]
    if (!page) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    const task = this.getEngine().renderPage(nativeDocument, page, {
      scaleFactor: Math.max(0.1, options.scale),
      rotation: asRotation(options.rotation),
      dpr: Math.max(1, options.dpr ?? 1),
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
    const nativeDocument = this.getNativeDocument(document)
    const page = nativeDocument.pages[pageIndex]
    if (!page) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    const task = this.getEngine().renderThumbnail(nativeDocument, page, {
      scaleFactor: Math.max(0.1, options.scale ?? 0.2),
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
    return result.results.map((hit) => ({
      pageIndex: hit.pageIndex,
      before: hit.context.before,
      match: hit.context.match,
      after: hit.context.after,
      rects: hit.rects.map((rect) => ({
        x: rect.origin.x,
        y: rect.origin.y,
        width: rect.size.width,
        height: rect.size.height,
      })),
    }))
  }

  async getPageTextRects(document: PdfRenderDocument, pageIndex: number): Promise<PdfTextRectItem[]> {
    const nativeDocument = this.getNativeDocument(document)
    const page = nativeDocument.pages[pageIndex]
    if (!page) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    const task = this.getEngine().getPageTextRects(nativeDocument, page)
    const rects = await runPdfTask(task, {
      timeoutMs: this.operationTimeoutMs,
      label: `PDF page ${pageIndex + 1} text rects`,
    })
    // Transform from PDF coords (bottom-left origin, Y-up) to CSS coords (top-left origin, Y-down)
    const pageHeight = page.size.height
    return rects.map((r) => {
      const pdfY = r.rect.origin.y
      const pdfH = r.rect.size.height
      return {
        content: r.content,
        x: r.rect.origin.x,
        y: pageHeight - pdfY - pdfH,
        width: r.rect.size.width,
        height: pdfH,
        fontFamily: r.font.family,
        fontSize: r.font.size,
      }
    })
  }

  async closeDocument(document: PdfRenderDocument): Promise<void> {
    const nativeDocument = this.documents.get(document.id)
    if (!nativeDocument) return
    this.documents.delete(document.id)
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
    } finally {
      this.revokeWorkerScriptUrls(engine)
    }
  }
}

export function createPdfRenderRuntime(config: PdfRenderRuntimeConfig = {}): PdfRenderRuntime {
  return new PdfiumRenderRuntime(config)
}
