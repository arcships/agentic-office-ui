import {
  createPdfiumEngine as createPdfiumWorkerEngine,
} from "@embedpdf/engines/pdfium-worker-engine"
import { formatOfficeMessage, officeMessagesEn } from "@arcships/office-interaction"
import {
  PdfAnnotationSubtype,
  PdfErrorCode,
  Rotation,
  type FormFieldValue,
  type PageTextSlice,
  type PdfAnnotationObject,
  type PdfDocumentObject,
  type PdfErrorReason,
  type PdfPageGeometry,
} from "@embedpdf/models"
import { pdfIntrinsicRectToCanonical } from "./pdf-coordinate-transform"
import { mapPdfAnnotationObjects } from "./pdf-annotation"
import { pdfPageTextHash } from "./pdf-page-signature"

export type PdfRotation = 0 | 90 | 180 | 270

export type PdfRenderErrorCode =
  | "ENCRYPTED"
  | "INVALID_RENDER_OPTIONS"
  | "PAGE_TOO_LARGE"
  | "INVALID_FIELD_VALUE"
  | "ANNOTATION_NOT_FOUND"
  | "ANNOTATION_NOT_WRITABLE"

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
  /** Canonical page width in 0-degree space (the visible width at zero user rotation). */
  width: number
  /** Canonical page height in 0-degree space (the visible height at zero user rotation). */
  height: number
  /**
   * Intrinsic `/Rotate` value in degrees. Documents are opened with normalized
   * rotation, so this is metadata for display math and diagnostics; engine
   * coordinates already account for it and must not apply it again.
   */
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

/** One page of a text selection, in document order. */
export interface PdfSelectionSegment {
  range: PageTextSlice
  text: string
  rects: readonly PdfRenderRect[]
}

export type PdfSelectionState =
  | { kind: "none" }
  | {
      kind: "text"
      /** First page slice of the selection, kept for single-page compatibility. */
      range: PageTextSlice
      /** Complete selection text. Multi-page selections join pages with `\n`. */
      text: string
      /** Highlight rectangles for `range.pageIndex`. */
      rects: readonly PdfRenderRect[]
      /** Every selected page in document order, including `range`. */
      segments?: readonly PdfSelectionSegment[]
    }
  | {
      kind: "unavailable"
      pageIndex: number
      reason: "empty" | "unicode-map"
    }

/**
 * A link, form widget or annotation rectangle that the raster page cannot
 * expose on its own. Rectangles are in canonical page space.
 */
export interface PdfRenderAnnotation {
  id: string
  pageIndex: number
  kind: "link" | "widget" | "annotation"
  rect: PdfRenderRect
  /** Sanitized absolute link target when the annotation links to a URL. */
  url?: string
  /** Destination page index when the annotation links inside the document. */
  targetPageIndex?: number
  /** True when the link target was rejected before it reached the host. */
  blocked?: boolean
  /** Form field name for widget annotations. */
  fieldName?: string
  /** Current form field value for widget annotations. */
  fieldValue?: string
  /** Writable form field kind for widget annotations. */
  fieldType?: PdfFormFieldType
  /** Export value that identifies a checkbox or radio widget inside its group. */
  exportValue?: string
  /** True when the widget field is read-only or locked. */
  readOnly?: boolean
  /** Selectable options for combo, list and radio widgets. */
  options?: readonly PdfFormFieldOption[]
  /** Annotation contents or subject, used for tooltips and notes. */
  contents?: string
  /** Current fill or stroke color when the annotation exposes one. */
  color?: string
  /** Current opacity in the `0..1` range when the annotation exposes one. */
  opacity?: number
  /** Raw annotation flags that affect interaction (`readOnly`, `noZoom`, ...). */
  flags?: readonly string[]
}

/** Writable AcroForm field kinds surfaced by `PdfRenderAnnotation`. */
export type PdfFormFieldType =
  | "text"
  | "checkbox"
  | "radio"
  | "combobox"
  | "listbox"
  | "button"
  | "signature"
  | "unknown"

export interface PdfFormFieldOption {
  label: string
  selected: boolean
}

/**
 * Value accepted by `PdfRenderRuntime.setFormFieldValue`. Text fields take a
 * string, checkboxes take a boolean, and radio/choice fields take either an
 * option label or its zero-based index as a string.
 */
export type PdfFormFieldValue = string | boolean

export interface PdfFormFieldDescriptor {
  fieldType?: PdfFormFieldType
  fieldValue?: string
  options?: readonly PdfFormFieldOption[]
}

/** Appearance fields a host may update on an existing annotation. */
export interface PdfAnnotationAppearancePatch {
  /** Annotation contents / note text. */
  contents?: string
  /** Fill or stroke color as a `#rrggbb` string, depending on the annotation. */
  color?: string
  /** Opacity in the `0..1` range. */
  opacity?: number
  /** Border/stroke color for link, ink and shape annotations. */
  strokeColor?: string
  /** Border/stroke width in points. */
  strokeWidth?: number
}

/** One positioned character of the accessible PDF text layer. */
export interface PdfPageTextLayerGlyph {
  /** PDFium character index used for text extraction and hit-testing. */
  charIndex: number
  text: string
  /** Canonical page-space box of the glyph; the layer transform scales it. */
  rect: PdfRenderRect
  /** True when PDFium has no reliable Unicode mapping for the character. */
  unmapped: boolean
}

/**
 * A page of extractable text positioned in canonical page space. The Surface
 * renders one transparent DOM span per glyph so browser find, screen readers and
 * CJK extraction see real text, while the bitmap stays the visual source.
 */
export interface PdfPageTextLayer {
  pageIndex: number
  /** Canonical page width; the text layer's untransformed coordinate width. */
  width: number
  /** Canonical page height; the text layer's untransformed coordinate height. */
  height: number
  /** Glyphs in PDFium reading order (run order, then character order). */
  glyphs: readonly PdfPageTextLayerGlyph[]
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

function normalizeCheckboxValue(value: PdfFormFieldValue): boolean {
  if (typeof value === "boolean") return value
  const normalized = value.trim().toLowerCase()
  if (["true", "on", "1", "yes", "checked"].includes(normalized)) return true
  if (["false", "off", "0", "no", "unchecked", ""].includes(normalized)) return false
  throw new PdfRenderError("INVALID_FIELD_VALUE", `Checkbox value "${value}" is not a boolean.`)
}

function resolveChoiceIndex(
  field: PdfFormFieldDescriptor,
  value: string,
): number {
  const options = field.options ?? []
  const labelIndex = options.findIndex((option) => option.label === value)
  if (labelIndex >= 0) return labelIndex
  const numeric = Number(value)
  if (Number.isInteger(numeric) && numeric >= 0 && numeric < options.length) return numeric
  throw new PdfRenderError(
    "INVALID_FIELD_VALUE",
    `Choice value "${value}" does not match any option (${options.length} available).`,
  )
}

/**
 * Translate a public form value into the PDFium form-field value for a widget.
 * Invalid combinations fail before the engine is touched so callers get a
 * stable error instead of a silently ignored write.
 */
export function resolvePdfFormFieldValue(
  field: PdfFormFieldDescriptor,
  value: PdfFormFieldValue,
): FormFieldValue {
  switch (field.fieldType) {
    case "checkbox":
      return { kind: "checked", checked: normalizeCheckboxValue(value) }
    case "radio":
    case "combobox":
    case "listbox": {
      if (typeof value !== "string") {
        throw new PdfRenderError(
          "INVALID_FIELD_VALUE",
          `Field type ${field.fieldType} expects an option label or index.`,
        )
      }
      const index = resolveChoiceIndex(field, value)
      return { kind: "selection", index, isSelected: true }
    }
    case "text":
    case "signature":
    case "unknown":
    case "button":
    case undefined:
      return { kind: "text", text: typeof value === "string" ? value : String(value) }
    default:
      return { kind: "text", text: typeof value === "string" ? value : String(value) }
  }
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
  /**
   * Optional link/widget/annotation hit layer. Runtimes that cannot enumerate
   * annotations simply omit this method and the Surface skips the hit layer.
   */
  getPageAnnotations?(
    document: PdfRenderDocument,
    pageIndex: number,
    options?: { signal?: AbortSignal },
  ): Promise<readonly PdfRenderAnnotation[]>
  /**
   * Positioned text for the accessible text layer. Optional so injected
   * runtimes that only raster pages can omit it; the Surface then renders no
   * text layer instead of fabricating text.
   */
  getPageTextLayer?(
    document: PdfRenderDocument,
    pageIndex: number,
    options?: { signal?: AbortSignal },
  ): Promise<PdfPageTextLayer>
  /**
   * Write one AcroForm widget value and return the refreshed annotation.
   *
   * Runtime implementers that cannot write forms omit this method; the Surface
   * then keeps widgets read-only instead of pretending the write succeeded.
   */
  setFormFieldValue?(
    document: PdfRenderDocument,
    pageIndex: number,
    annotationId: string,
    value: PdfFormFieldValue,
    options?: { signal?: AbortSignal },
  ): Promise<PdfRenderAnnotation>
  /**
   * Update an existing annotation's appearance (contents, color, opacity or
   * stroke) and regenerate its appearance stream.
   */
  updateAnnotationAppearance?(
    document: PdfRenderDocument,
    pageIndex: number,
    annotationId: string,
    patch: PdfAnnotationAppearancePatch,
    options?: { signal?: AbortSignal },
  ): Promise<PdfRenderAnnotation>
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
    task.abort({ code: PdfErrorCode.Cancelled, message: formatOfficeMessage(officeMessagesEn.pdf.runtime.cancelled, { label }) })
    throw createAbortError()
  }

  let timeout: ReturnType<typeof setTimeout> | undefined
  let removeAbortListener: (() => void) | undefined
  const guard = new Promise<never>((_resolve, reject) => {
    const abort = () => {
      task.abort({ code: PdfErrorCode.Cancelled, message: formatOfficeMessage(officeMessagesEn.pdf.runtime.cancelled, { label }) })
      reject(createAbortError())
    }
    signal?.addEventListener("abort", abort, { once: true })
    removeAbortListener = () => signal?.removeEventListener("abort", abort)
    timeout = setTimeout(() => {
      task.abort({ code: PdfErrorCode.Cancelled, message: formatOfficeMessage(officeMessagesEn.pdf.runtime.timedOut, { label }) })
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
  private readonly rawPageAnnotations = new Map<string, readonly PdfAnnotationObject[]>()
  private readonly pageTextLayers = new Map<string, PdfPageTextLayer>()
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
        {
          ...(password === undefined ? {} : { password }),
          // Page sizes from PDFium are already rotated for /Rotate pages, while
          // this runtime keeps the original rotation in `page.rotation` and
          // applies it explicitly to rendering and canonical geometry. Opening
          // in normalized (0-degree) space makes both models agree; without it
          // a /Rotate 90 page is rotated twice and renders sideways.
          normalizeRotation: true,
        },
      )
      try {
        nativeDocument = await runPdfTask(task, {
          signal: options.signal,
          timeoutMs: this.operationTimeoutMs,
          label: officeMessagesEn.pdf.runtime.workerStartup,
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
            message: officeMessagesEn.pdf.errors.encrypted,
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
      // Thumbnails must show the page as it is displayed at zero user rotation,
      // so the intrinsic /Rotate value is applied here explicitly.
      rotation: asRotation(parameters.rotation),
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
      label: officeMessagesEn.pdf.runtime.textSearch,
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
      label: officeMessagesEn.pdf.runtime.textExtraction,
    })
  }

  private pageCacheKey(document: PdfRenderDocument, pageIndex: number): string {
    return `${document.id}:${pageIndex}`
  }

  private async readRawPageAnnotations(
    document: PdfRenderDocument,
    pageIndex: number,
    options: { signal?: AbortSignal } = {},
  ): Promise<readonly PdfAnnotationObject[]> {
    const key = this.pageCacheKey(document, pageIndex)
    const cached = this.rawPageAnnotations.get(key)
    if (cached) return cached
    const nativeDocument = this.getNativeDocument(document)
    const page = nativeDocument.pages[pageIndex]
    if (!page) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    const annotations = await runPdfTask(this.getEngine().getPageAnnotations(nativeDocument, page), {
      signal: options.signal,
      timeoutMs: this.operationTimeoutMs,
      label: `PDF page ${pageIndex + 1} annotation read`,
    })
    if (this.documents.has(document.id)) this.rawPageAnnotations.set(key, annotations)
    return annotations
  }

  private async refreshMappedPageAnnotations(
    document: PdfRenderDocument,
    pageIndex: number,
    options: { signal?: AbortSignal } = {},
  ): Promise<readonly PdfRenderAnnotation[]> {
    this.rawPageAnnotations.delete(this.pageCacheKey(document, pageIndex))
    const annotations = await this.readRawPageAnnotations(document, pageIndex, options)
    const renderPage = document.pages[pageIndex]
    if (!renderPage) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    return mapPdfAnnotationObjects(annotations, renderPage)
  }

  private async writableAnnotation(
    document: PdfRenderDocument,
    pageIndex: number,
    annotationId: string,
    options: { signal?: AbortSignal } = {},
  ): Promise<PdfAnnotationObject> {
    const annotations = await this.readRawPageAnnotations(document, pageIndex, options)
    const annotation = annotations.find((candidate) => candidate.id === annotationId)
    if (!annotation) {
      throw new PdfRenderError(
        "ANNOTATION_NOT_FOUND",
        `PDF annotation ${annotationId} was not found on page ${pageIndex + 1}.`,
      )
    }
    return annotation
  }

  async getPageAnnotations(
    document: PdfRenderDocument,
    pageIndex: number,
    options: { signal?: AbortSignal } = {},
  ): Promise<readonly PdfRenderAnnotation[]> {
    const renderPage = document.pages[pageIndex]
    if (!renderPage) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    const annotations = await this.readRawPageAnnotations(document, pageIndex, options)
    return mapPdfAnnotationObjects(annotations, renderPage)
  }

  async getPageTextLayer(
    document: PdfRenderDocument,
    pageIndex: number,
    options: { signal?: AbortSignal } = {},
  ): Promise<PdfPageTextLayer> {
    const key = this.pageCacheKey(document, pageIndex)
    const cached = this.pageTextLayers.get(key)
    if (cached) return cached
    const nativeDocument = this.getNativeDocument(document)
    const page = nativeDocument.pages[pageIndex]
    const renderPage = document.pages[pageIndex]
    if (!page || !renderPage) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    const geometry = await runPdfTask(this.getEngine().getPageGeometry(nativeDocument, page), {
      signal: options.signal,
      timeoutMs: this.operationTimeoutMs,
      label: `PDF page ${pageIndex + 1} text layer geometry`,
    })
    const boxes: Array<{ charIndex: number; rect: PdfRenderRect }> = []
    for (const run of geometry.runs) {
      for (let offset = 0; offset < run.glyphs.length; offset++) {
        const glyph = run.glyphs[offset]
        // flags 2 marks an empty glyph that carries no text or box.
        if (glyph.flags === 2) continue
        if (!(glyph.width > 0) || !(glyph.height > 0)) continue
        boxes.push({
          charIndex: run.charStart + offset,
          rect: { x: glyph.x, y: glyph.y, width: glyph.width, height: glyph.height },
        })
      }
    }
    const texts = boxes.length > 0
      ? await runPdfTask(
          this.getEngine().getTextSlices(
            nativeDocument,
            boxes.map((box) => ({ pageIndex, charIndex: box.charIndex, charCount: 1 })),
          ),
          {
            signal: options.signal,
            timeoutMs: this.operationTimeoutMs,
            label: `PDF page ${pageIndex + 1} text layer characters`,
          },
        )
      : []
    const glyphs: PdfPageTextLayerGlyph[] = []
    for (let index = 0; index < boxes.length; index++) {
      const text = texts[index] ?? ""
      if (!text) continue
      glyphs.push({
        charIndex: boxes[index].charIndex,
        rect: boxes[index].rect,
        text,
        unmapped: /[\u0000\ufffd\ufffe]/u.test(text),
      })
    }
    const layer: PdfPageTextLayer = {
      pageIndex,
      width: renderPage.width,
      height: renderPage.height,
      glyphs,
    }
    if (this.documents.has(document.id)) this.pageTextLayers.set(key, layer)
    return layer
  }

  async setFormFieldValue(
    document: PdfRenderDocument,
    pageIndex: number,
    annotationId: string,
    value: PdfFormFieldValue,
    options: { signal?: AbortSignal } = {},
  ): Promise<PdfRenderAnnotation> {
    const nativeDocument = this.getNativeDocument(document)
    const page = nativeDocument.pages[pageIndex]
    const renderPage = document.pages[pageIndex]
    if (!page || !renderPage) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    const annotation = await this.writableAnnotation(document, pageIndex, annotationId, options)
    if (annotation.type !== PdfAnnotationSubtype.WIDGET) {
      throw new PdfRenderError(
        "ANNOTATION_NOT_WRITABLE",
        `PDF annotation ${annotationId} is not a form widget.`,
      )
    }
    const mapped = mapPdfAnnotationObjects([annotation], renderPage)[0]
    if (!mapped) {
      throw new PdfRenderError("ANNOTATION_NOT_WRITABLE", `PDF annotation ${annotationId} is hidden.`)
    }
    const formValue = resolvePdfFormFieldValue(mapped, value)
    await runPdfTask(this.getEngine().setFormFieldValue(nativeDocument, page, annotation, formValue), {
      signal: options.signal,
      timeoutMs: this.operationTimeoutMs,
      label: `PDF page ${pageIndex + 1} form field write`,
    })
    const updated = await this.refreshMappedPageAnnotations(document, pageIndex, options)
    const next = updated.find((candidate) => candidate.id === annotationId)
    if (!next) throw new PdfRenderError("ANNOTATION_NOT_FOUND", `PDF annotation ${annotationId} disappeared after the write.`)
    // A form value can change the page's extractable text, so drop the text
    // layer for this page; appearance-only updates do not need it.
    this.pageTextLayers.delete(this.pageCacheKey(document, pageIndex))
    return next
  }

  async updateAnnotationAppearance(
    document: PdfRenderDocument,
    pageIndex: number,
    annotationId: string,
    patch: PdfAnnotationAppearancePatch,
    options: { signal?: AbortSignal } = {},
  ): Promise<PdfRenderAnnotation> {
    const nativeDocument = this.getNativeDocument(document)
    const page = nativeDocument.pages[pageIndex]
    const renderPage = document.pages[pageIndex]
    if (!page || !renderPage) throw new Error(`PDF page ${pageIndex + 1} does not exist.`)
    const annotation = await this.writableAnnotation(document, pageIndex, annotationId, options)
    const target = annotation as PdfAnnotationObject & Record<string, unknown>
    if (patch.opacity !== undefined) {
      if (!Number.isFinite(patch.opacity) || patch.opacity < 0 || patch.opacity > 1) {
        throw new PdfRenderError("INVALID_RENDER_OPTIONS", "Annotation opacity must be between 0 and 1.")
      }
      target.opacity = patch.opacity
    }
    if (patch.strokeWidth !== undefined) {
      if (!Number.isFinite(patch.strokeWidth) || patch.strokeWidth < 0) {
        throw new PdfRenderError("INVALID_RENDER_OPTIONS", "Annotation stroke width must be a non-negative number.")
      }
      if (!("strokeWidth" in target)) {
        throw new PdfRenderError("ANNOTATION_NOT_WRITABLE", `PDF annotation ${annotationId} has no stroke width.`)
      }
      target.strokeWidth = patch.strokeWidth
    }
    for (const [field, value] of [["contents", patch.contents], ["color", patch.color], ["strokeColor", patch.strokeColor]] as const) {
      if (value === undefined) continue
      if (typeof value !== "string") {
        throw new PdfRenderError("INVALID_RENDER_OPTIONS", `Annotation ${field} must be a string.`)
      }
      if (!(field in target) && !(field === "color" && "strokeColor" in target)) {
        throw new PdfRenderError("ANNOTATION_NOT_WRITABLE", `PDF annotation ${annotationId} has no ${field}.`)
      }
      if (field === "color" && !("color" in target) && "strokeColor" in target) target.strokeColor = value
      else target[field] = value
    }
    await runPdfTask(this.getEngine().updatePageAnnotation(nativeDocument, page, annotation, { regenerateAppearance: true }), {
      signal: options.signal,
      timeoutMs: this.operationTimeoutMs,
      label: `PDF page ${pageIndex + 1} annotation update`,
    })
    const updated = await this.refreshMappedPageAnnotations(document, pageIndex, options)
    const next = updated.find((candidate) => candidate.id === annotationId)
    if (!next) throw new PdfRenderError("ANNOTATION_NOT_FOUND", `PDF annotation ${annotationId} disappeared after the update.`)
    return next
  }

  async closeDocument(document: PdfRenderDocument): Promise<void> {
    const nativeDocument = this.documents.get(document.id)
    if (!nativeDocument) return
    this.documents.delete(document.id)
    for (const key of this.contentHashRequests.keys()) {
      if (key.startsWith(`${document.id}:`)) this.contentHashRequests.delete(key)
    }
    for (const key of this.rawPageAnnotations.keys()) {
      if (key.startsWith(`${document.id}:`)) this.rawPageAnnotations.delete(key)
    }
    for (const key of this.pageTextLayers.keys()) {
      if (key.startsWith(`${document.id}:`)) this.pageTextLayers.delete(key)
    }
    const engine = this.engine
    if (!engine) return
    try {
      await runPdfTask(engine.closeDocument(nativeDocument), {
        timeoutMs: this.operationTimeoutMs,
        label: officeMessagesEn.pdf.runtime.documentClose,
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
    this.rawPageAnnotations.clear()
    this.pageTextLayers.clear()
    const engine = this.engine
    this.engine = undefined
    if (!engine) return
    try {
      await runPdfTask(engine.destroy(), {
        timeoutMs: Math.min(this.operationTimeoutMs, 5_000),
        label: officeMessagesEn.pdf.runtime.runtimeDestroy,
      })
    } catch {
      // The runtime is already unusable and all local references are released.
    }
  }
}

export function createPdfRenderRuntime(config: PdfRenderRuntimeConfig = {}): PdfRenderRuntime {
  return new PdfiumRenderRuntime(config)
}
