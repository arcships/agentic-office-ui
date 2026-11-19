import type {
  ManualRegionLocator,
  NormalizedRect,
  OfficeFormat,
} from "@arcships/office-interaction"
import { rasterizeElementRegion } from "./dom-raster"

/** Pixel budget for one capture; larger requests reject before allocating. */
export const DEFAULT_REGION_CAPTURE_MAX_PIXELS = 24_000_000
/** Longest edge of a captured bitmap. Matches the preview limit used by the surfaces. */
export const DEFAULT_REGION_CAPTURE_MAX_DIMENSION = 16_384
/** Byte budget for the encoded bitmap. */
export const DEFAULT_REGION_CAPTURE_MAX_BYTES = 32 * 1024 * 1024
/** 8x is enough for print previews while keeping accidental allocations bounded. */
export const MAX_REGION_CAPTURE_SCALE = 8
/** Default bitmap type. PNG preserves text edges and transparency. */
export const DEFAULT_REGION_CAPTURE_MIME_TYPE = "image/png"

/**
 * A region may be expressed as the normalized rectangle emitted by
 * `OfficeRegionSelector`, or as a format-neutral `ManualRegionLocator`
 * (page, slide, or cell-anchored sheet region). Surfaces that can resolve a
 * locator into their own element geometry do so and return the resolved
 * normalized `region` with their capture source.
 */
export type OfficeRegionCaptureRegion = NormalizedRect | ManualRegionLocator

export interface OfficeRegionCaptureTarget {
  pageIndex?: number
  slideIndex?: number
  sheetId?: string
}

/** An already-rasterized representation of the complete surface element. */
export interface OfficeRegionCaptureRaster {
  source: CanvasImageSource
  /** Intrinsic pixel width of `source`. */
  width: number
  /** Intrinsic pixel height of `source`. */
  height: number
}

/** Serializable provenance copied next to the capture. */
export interface OfficeRegionCaptureSourceMetadata {
  format?: OfficeFormat
  documentId?: string
  revision?: string
  /** Which format-neutral container the region belongs to. */
  space?: "page" | "slide" | "sheet"
  pageIndex?: number
  slideIndex?: number
  sheetId?: string
}

/**
 * A surface describes the element that owns the normalized coordinate space.
 * `raster` is optional: canvas-backed surfaces (XLSX/PDF) should provide it so
 * capture crops the rendered pixels instead of re-rasterizing the DOM.
 */
export interface OfficeRegionCaptureSource extends OfficeRegionCaptureSourceMetadata {
  element: Element
  raster?: OfficeRegionCaptureRaster | null
  /** Provider-resolved normalized region, required when the request used a locator. */
  region?: NormalizedRect
}

export interface OfficeRegionCaptureRequest {
  region: OfficeRegionCaptureRegion
  target?: OfficeRegionCaptureTarget
  scale: number
}

/**
 * First-party surfaces implement `getCaptureSurface()` so hosts can pass the
 * exposed component instance directly:
 *
 * ```ts
 * const capture = await captureRegion({ surface: surfaceRef.value, region, scale })
 * ```
 */
export interface OfficeRegionCaptureProvider {
  getCaptureSurface?: (
    request: OfficeRegionCaptureRequest,
  ) => OfficeRegionCaptureSource | null | Promise<OfficeRegionCaptureSource | null>
  /** Optional direct source for simple host wrappers. */
  captureSource?: OfficeRegionCaptureSource | null
  /** Optional direct descriptors; used when no provider method is present. */
  element?: Element | null
  raster?: OfficeRegionCaptureRaster | null
  region?: NormalizedRect
}

export type OfficeRegionCaptureSurface = Element | OfficeRegionCaptureProvider

export interface CaptureRegionOptions {
  /** An element, a component expose object, or a direct capture provider. */
  surface: OfficeRegionCaptureSurface | null | undefined
  region: OfficeRegionCaptureRegion
  /**
   * Output pixels per CSS pixel for the selected element. Defaults to
   * `window.devicePixelRatio` (1 when unavailable), so Retina devices keep
   * their pixel density unless the host asks for an explicit scale.
   */
  scale?: number
  target?: OfficeRegionCaptureTarget
  maxPixels?: number
  maxDimension?: number
  maxBytes?: number
  mimeType?: string
  quality?: number
  signal?: AbortSignal
  /** Return an object URL. Defaults to `true`; call `release()` when done. */
  objectUrl?: boolean
  /** Background painted behind transparent DOM captures. `null` keeps transparency. */
  background?: string | null
  /** Test/worker hook; defaults to `<canvas>` or `OffscreenCanvas`. */
  createCanvas?: (width: number, height: number) => HTMLCanvasElement
  /** Deterministic timestamp hook for tests. */
  now?: () => Date
}

export interface OfficeRegionCaptureMetadata extends OfficeRegionCaptureSourceMetadata {
  /** How the pixels were produced. */
  kind: "dom" | "canvas" | "image" | "bitmap"
  /** Normalized region inside the source element that was captured. */
  region: NormalizedRect
  scale: number
  width: number
  height: number
  devicePixelRatio: number
  elementWidth: number
  elementHeight: number
  sourceWidth: number
  sourceHeight: number
  capturedAt: string
  mimeType: string
  byteLength: number
}

export interface OfficeRegionCapture {
  blob: Blob
  /** Object URL, or `null` when URL support is unavailable or disabled. */
  url: string | null
  width: number
  height: number
  scale: number
  region: NormalizedRect
  source: OfficeRegionCaptureMetadata
  /** Revokes `url`; safe to call more than once. */
  release(): void
}

export type OfficeRegionCaptureErrorCode =
  | "INVALID_REGION"
  | "CAPTURE_UNSUPPORTED"
  | "CAPTURE_LIMIT_EXCEEDED"
  | "CAPTURE_FAILED"
  | "ABORTED"

export class OfficeRegionCaptureError extends Error {
  readonly code: OfficeRegionCaptureErrorCode
  readonly operation = "capture" as const

  constructor(code: OfficeRegionCaptureErrorCode, message: string, options?: { cause?: unknown }) {
    super(message)
    this.name = "OfficeRegionCaptureError"
    this.code = code
    if (options?.cause !== undefined) (this as { cause?: unknown }).cause = options.cause
  }
}

export interface RegionCaptureGeometryInput {
  region: NormalizedRect
  elementWidth: number
  elementHeight: number
  scale?: number
  devicePixelRatio?: number
  maxPixels?: number
  maxDimension?: number
}

export interface RegionCaptureGeometry {
  region: NormalizedRect
  scale: number
  devicePixelRatio: number
  elementWidth: number
  elementHeight: number
  outputWidth: number
  outputHeight: number
}

/** Pixel rectangle cropped from an existing raster. */
export interface RegionCaptureCrop {
  x: number
  y: number
  width: number
  height: number
}

function clampUnit(value: number): number {
  if (!Number.isFinite(value)) return 0
  return Math.min(1, Math.max(0, value))
}

/** Clamp all four edges so a selection that touches the surface stays valid. */
export function clampNormalizedRegion(region: NormalizedRect): NormalizedRect {
  if (!region || typeof region !== "object") {
    throw new OfficeRegionCaptureError("INVALID_REGION", "region must be an object")
  }
  const x = clampUnit(region.x)
  const y = clampUnit(region.y)
  const right = clampUnit(region.x + region.width)
  const bottom = clampUnit(region.y + region.height)
  const width = right - x
  const height = bottom - y
  if (!(width > 0) || !(height > 0)) {
    throw new OfficeRegionCaptureError("INVALID_REGION", "region must have a positive area inside [0, 1]")
  }
  return { x, y, width, height }
}

export function isOfficeNormalizedRect(
  region: OfficeRegionCaptureRegion,
): region is NormalizedRect {
  return typeof region === "object"
    && region !== null
    && !("space" in region)
    && Number.isFinite((region as NormalizedRect).x)
    && Number.isFinite((region as NormalizedRect).y)
    && Number.isFinite((region as NormalizedRect).width)
    && Number.isFinite((region as NormalizedRect).height)
}

export function isOfficeManualRegionLocator(
  region: OfficeRegionCaptureRegion,
): region is ManualRegionLocator {
  return typeof region === "object"
    && region !== null
    && "space" in region
    && ((region as ManualRegionLocator).space === "page"
      || (region as ManualRegionLocator).space === "slide"
      || (region as ManualRegionLocator).space === "sheet")
}

/** Target inferred from a manual locator when the host did not pass one. */
export function targetForRegion(region: OfficeRegionCaptureRegion): OfficeRegionCaptureTarget | undefined {
  if (!isOfficeManualRegionLocator(region)) return undefined
  if (region.space === "page") return { pageIndex: region.pageIndex }
  if (region.space === "slide") return { slideIndex: region.slideIndex }
  return { sheetId: region.sheetId }
}

function positiveInteger(value: number | undefined, fallback: number, label: string): number {
  if (value === undefined) return fallback
  if (!Number.isFinite(value) || value <= 0) {
    throw new OfficeRegionCaptureError("CAPTURE_LIMIT_EXCEEDED", `${label} must be a positive number`)
  }
  return Math.floor(value)
}

function resolveDevicePixelRatio(explicit?: number): number {
  if (explicit !== undefined) {
    if (!Number.isFinite(explicit) || explicit <= 0) {
      throw new OfficeRegionCaptureError("CAPTURE_LIMIT_EXCEEDED", "devicePixelRatio must be a positive number")
    }
    return explicit
  }
  const detected = (globalThis as { devicePixelRatio?: unknown }).devicePixelRatio
  return typeof detected === "number" && Number.isFinite(detected) && detected > 0 ? detected : 1
}

function resolveScale(scale: number | undefined, devicePixelRatio: number): number {
  const resolved = scale ?? devicePixelRatio
  if (!Number.isFinite(resolved) || resolved <= 0) {
    throw new OfficeRegionCaptureError("CAPTURE_LIMIT_EXCEEDED", "scale must be a positive number")
  }
  if (resolved > MAX_REGION_CAPTURE_SCALE) {
    throw new OfficeRegionCaptureError(
      "CAPTURE_LIMIT_EXCEEDED",
      `scale must not exceed ${MAX_REGION_CAPTURE_SCALE}`,
    )
  }
  return resolved
}

/**
 * Pure output-size and limit math. Hosts and tests can use it to preflight a
 * request without touching the DOM.
 */
export function computeRegionCaptureGeometry(input: RegionCaptureGeometryInput): RegionCaptureGeometry {
  const region = clampNormalizedRegion(input.region)
  const elementWidth = input.elementWidth
  const elementHeight = input.elementHeight
  if (!Number.isFinite(elementWidth) || elementWidth <= 0 || !Number.isFinite(elementHeight) || elementHeight <= 0) {
    throw new OfficeRegionCaptureError("CAPTURE_UNSUPPORTED", "capture element must have a positive size")
  }
  const devicePixelRatio = resolveDevicePixelRatio(input.devicePixelRatio)
  const scale = resolveScale(input.scale, devicePixelRatio)
  const maxPixels = positiveInteger(input.maxPixels, DEFAULT_REGION_CAPTURE_MAX_PIXELS, "maxPixels")
  const maxDimension = positiveInteger(input.maxDimension, DEFAULT_REGION_CAPTURE_MAX_DIMENSION, "maxDimension")
  const outputWidth = Math.max(1, Math.round(region.width * elementWidth * scale))
  const outputHeight = Math.max(1, Math.round(region.height * elementHeight * scale))
  if (outputWidth > maxDimension || outputHeight > maxDimension) {
    throw new OfficeRegionCaptureError(
      "CAPTURE_LIMIT_EXCEEDED",
      `capture must not exceed ${maxDimension} pixels per edge (requested ${outputWidth}x${outputHeight})`,
    )
  }
  if (outputWidth * outputHeight > maxPixels) {
    throw new OfficeRegionCaptureError(
      "CAPTURE_LIMIT_EXCEEDED",
      `capture must not exceed ${maxPixels} pixels (requested ${outputWidth * outputHeight})`,
    )
  }
  return {
    region,
    scale,
    devicePixelRatio,
    elementWidth,
    elementHeight,
    outputWidth,
    outputHeight,
  }
}

/** Pixel-exact crop from an existing raster. Edges are rounded once, then clamped. */
export function computeRegionRasterCrop(
  region: NormalizedRect,
  sourceWidth: number,
  sourceHeight: number,
): RegionCaptureCrop {
  const normalized = clampNormalizedRegion(region)
  if (!Number.isFinite(sourceWidth) || sourceWidth <= 0 || !Number.isFinite(sourceHeight) || sourceHeight <= 0) {
    throw new OfficeRegionCaptureError("CAPTURE_UNSUPPORTED", "capture raster must have a positive size")
  }
  const left = Math.round(normalized.x * sourceWidth)
  const top = Math.round(normalized.y * sourceHeight)
  const right = Math.round((normalized.x + normalized.width) * sourceWidth)
  const bottom = Math.round((normalized.y + normalized.height) * sourceHeight)
  const x = Math.min(sourceWidth - 1, Math.max(0, left))
  const y = Math.min(sourceHeight - 1, Math.max(0, top))
  return {
    x,
    y,
    width: Math.max(1, Math.min(sourceWidth - x, right - x)),
    height: Math.max(1, Math.min(sourceHeight - y, bottom - y)),
  }
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return
  throw new OfficeRegionCaptureError("ABORTED", "region capture was aborted")
}

function isElementLike(value: unknown): value is Element {
  return !!value
    && typeof value === "object"
    && typeof (value as Element).getBoundingClientRect === "function"
    && typeof (value as Element).nodeType === "number"
}

function isCanvasLike(value: unknown): value is HTMLCanvasElement {
  return !!value
    && typeof value === "object"
    && typeof (value as HTMLCanvasElement).getContext === "function"
    && typeof (value as HTMLCanvasElement).width === "number"
    && typeof (value as HTMLCanvasElement).height === "number"
}

function isImageLike(value: unknown): value is HTMLImageElement {
  return !!value
    && typeof value === "object"
    && (value as HTMLImageElement).tagName?.toUpperCase() === "IMG"
    && typeof (value as HTMLImageElement).naturalWidth === "number"
    && (value as HTMLImageElement).naturalWidth > 0
}

async function resolveCaptureSource(
  surface: OfficeRegionCaptureSurface | null | undefined,
  request: OfficeRegionCaptureRequest,
): Promise<OfficeRegionCaptureSource> {
  if (isElementLike(surface)) return { element: surface }
  if (surface && typeof surface === "object") {
    const provider = surface as OfficeRegionCaptureProvider
    if (typeof provider.getCaptureSurface === "function") {
      const source = await provider.getCaptureSurface(request)
      if (source?.element) return source
    }
    if (provider.captureSource?.element) return provider.captureSource
    if (isElementLike(provider.element)) {
      return {
        element: provider.element,
        raster: provider.raster ?? null,
        region: provider.region,
      }
    }
  }
  throw new OfficeRegionCaptureError(
    "CAPTURE_UNSUPPORTED",
    "surface does not expose a capture element; pass an element or implement getCaptureSurface()",
  )
}

function defaultCreateCanvas(width: number, height: number): HTMLCanvasElement {
  if (typeof document !== "undefined" && typeof document.createElement === "function") {
    const canvas = document.createElement("canvas")
    canvas.width = width
    canvas.height = height
    return canvas
  }
  const OffscreenCanvasConstructor = (globalThis as { OffscreenCanvas?: typeof OffscreenCanvas }).OffscreenCanvas
  if (typeof OffscreenCanvasConstructor === "function") {
    return new OffscreenCanvasConstructor(width, height) as unknown as HTMLCanvasElement
  }
  throw new OfficeRegionCaptureError(
    "CAPTURE_UNSUPPORTED",
    "canvas is unavailable; region capture needs a DOM canvas or OffscreenCanvas",
  )
}

function context2d(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const context = canvas.getContext("2d") as CanvasRenderingContext2D | null
  if (!context || typeof context.drawImage !== "function") {
    throw new OfficeRegionCaptureError("CAPTURE_UNSUPPORTED", "2D canvas context is unavailable")
  }
  return context
}

async function canvasToBlob(
  canvas: HTMLCanvasElement,
  mimeType: string,
  quality: number | undefined,
): Promise<Blob> {
  const offscreen = canvas as unknown as { convertToBlob?: (options?: { type?: string; quality?: number }) => Promise<Blob> }
  if (typeof offscreen.convertToBlob === "function") {
    return offscreen.convertToBlob({ type: mimeType, quality })
  }
  if (typeof canvas.toBlob === "function") {
    return new Promise<Blob>((resolve, reject) => {
      canvas.toBlob((blob) => {
        if (blob) resolve(blob)
        else reject(new OfficeRegionCaptureError("CAPTURE_FAILED", "canvas encoder returned no bitmap"))
      }, mimeType, quality)
    })
  }
  throw new OfficeRegionCaptureError("CAPTURE_UNSUPPORTED", "canvas cannot encode a bitmap in this environment")
}

function sourceKind(source: OfficeRegionCaptureSource, element: Element): OfficeRegionCaptureMetadata["kind"] {
  const raster = source.raster
  if (raster) {
    if (isCanvasLike(raster.source)) return "canvas"
    if (isImageLike(raster.source)) return "image"
    return "bitmap"
  }
  if (isCanvasLike(element)) return "canvas"
  if (isImageLike(element)) return "image"
  return "dom"
}

function sourceRaster(source: OfficeRegionCaptureSource): OfficeRegionCaptureRaster | null {
  if (source.raster) return source.raster
  if (isCanvasLike(source.element) && source.element.width > 0 && source.element.height > 0) {
    return { source: source.element, width: source.element.width, height: source.element.height }
  }
  if (isImageLike(source.element)) {
    return { source: source.element, width: source.element.naturalWidth, height: source.element.naturalHeight }
  }
  return null
}

/**
 * Capture the visible pixels of `region` inside a surface element.
 *
 * The returned object owns an object URL (unless `objectUrl: false`) and must
 * be released with `release()` or {@link releaseRegionCapture}. Pixel
 * allocation is bounded by `maxPixels`, `maxDimension`, `maxBytes`, and
 * `MAX_REGION_CAPTURE_SCALE`; limits are checked before the canvas is created.
 *
 * DOM-backed surfaces are rasterized through an SVG `foreignObject`; canvas or
 * image-backed surfaces should provide a `raster` source for pixel-exact
 * cropping. Headless hosts that cannot create a canvas get a
 * `CAPTURE_UNSUPPORTED` error instead of touching `document`.
 */
export async function captureRegion(options: CaptureRegionOptions): Promise<OfficeRegionCapture> {
  throwIfAborted(options.signal)
  const devicePixelRatio = resolveDevicePixelRatio()
  const scale = resolveScale(options.scale, devicePixelRatio)
  const request: OfficeRegionCaptureRequest = {
    region: options.region,
    target: options.target ?? targetForRegion(options.region),
    scale,
  }
  const source = await resolveCaptureSource(options.surface, request)
  throwIfAborted(options.signal)
  const element = source.element
  const bounds = element.getBoundingClientRect()
  const elementWidth = bounds?.width ?? 0
  const elementHeight = bounds?.height ?? 0
  if (!(elementWidth > 0) || !(elementHeight > 0)) {
    throw new OfficeRegionCaptureError(
      "CAPTURE_UNSUPPORTED",
      "capture element is not laid out (zero width or height)",
    )
  }
  const region = source.region ?? (isOfficeNormalizedRect(request.region) ? request.region : null)
  if (!region) {
    throw new OfficeRegionCaptureError(
      "INVALID_REGION",
      "surface did not resolve the manual region locator to a normalized region",
    )
  }
  const geometry = computeRegionCaptureGeometry({
    region,
    elementWidth,
    elementHeight,
    scale,
    devicePixelRatio,
    maxPixels: options.maxPixels,
    maxDimension: options.maxDimension,
  })
  const raster = sourceRaster(source)
  const canvas = options.createCanvas
    ? options.createCanvas(geometry.outputWidth, geometry.outputHeight)
    : defaultCreateCanvas(geometry.outputWidth, geometry.outputHeight)
  if (!canvas) {
    throw new OfficeRegionCaptureError("CAPTURE_UNSUPPORTED", "canvas factory returned no canvas")
  }
  if (canvas.width !== geometry.outputWidth || canvas.height !== geometry.outputHeight) {
    canvas.width = geometry.outputWidth
    canvas.height = geometry.outputHeight
  }
  const context = context2d(canvas)
  let kind: OfficeRegionCaptureMetadata["kind"]
  try {
    if (raster) {
      throwIfAborted(options.signal)
      const crop = computeRegionRasterCrop(geometry.region, raster.width, raster.height)
      kind = sourceKind({ ...source, raster }, element)
      context.drawImage(
        raster.source,
        crop.x,
        crop.y,
        crop.width,
        crop.height,
        0,
        0,
        geometry.outputWidth,
        geometry.outputHeight,
      )
    } else {
      kind = sourceKind(source, element)
      await rasterizeElementRegion({
        element,
        region: geometry.region,
        outputWidth: geometry.outputWidth,
        outputHeight: geometry.outputHeight,
        background: options.background,
        signal: options.signal,
        createCanvas: options.createCanvas,
        drawInto: { canvas, context },
      })
    }
  } catch (reason) {
    if (reason instanceof OfficeRegionCaptureError) throw reason
    const detail = reason instanceof Error ? reason.message : String(reason)
    throw new OfficeRegionCaptureError(
      "CAPTURE_FAILED",
      detail,
      { cause: reason },
    )
  }
  throwIfAborted(options.signal)
  const mimeType = options.mimeType || DEFAULT_REGION_CAPTURE_MIME_TYPE
  const blob = await canvasToBlob(canvas, mimeType, options.quality)
  const maxBytes = positiveInteger(options.maxBytes, DEFAULT_REGION_CAPTURE_MAX_BYTES, "maxBytes")
  if (blob.size > maxBytes) {
    throw new OfficeRegionCaptureError(
      "CAPTURE_LIMIT_EXCEEDED",
      `capture bitmap is ${blob.size} bytes; limit is ${maxBytes}`,
    )
  }
  let url: string | null = null
  if (options.objectUrl !== false && typeof URL !== "undefined" && typeof URL.createObjectURL === "function") {
    try {
      url = URL.createObjectURL(blob)
    } catch {
      url = null
    }
  }
  let released = false
  const metadata: OfficeRegionCaptureMetadata = {
    kind,
    ...(source.format ? { format: source.format } : {}),
    ...(source.documentId ? { documentId: source.documentId } : {}),
    ...(source.revision ? { revision: source.revision } : {}),
    ...(source.space ? { space: source.space } : {}),
    ...(source.pageIndex !== undefined ? { pageIndex: source.pageIndex } : {}),
    ...(source.slideIndex !== undefined ? { slideIndex: source.slideIndex } : {}),
    ...(source.sheetId ? { sheetId: source.sheetId } : {}),
    region: geometry.region,
    scale: geometry.scale,
    width: geometry.outputWidth,
    height: geometry.outputHeight,
    devicePixelRatio: geometry.devicePixelRatio,
    elementWidth,
    elementHeight,
    sourceWidth: raster?.width ?? geometry.outputWidth,
    sourceHeight: raster?.height ?? geometry.outputHeight,
    capturedAt: (options.now?.() ?? new Date()).toISOString(),
    mimeType: blob.type || mimeType,
    byteLength: blob.size,
  }
  return {
    blob,
    url,
    width: geometry.outputWidth,
    height: geometry.outputHeight,
    scale: geometry.scale,
    region: geometry.region,
    source: metadata,
    release() {
      if (released) return
      released = true
      if (url && typeof URL !== "undefined" && typeof URL.revokeObjectURL === "function") URL.revokeObjectURL(url)
    },
  }
}

/** Release a capture's object URL. Safe with `null` and repeated calls. */
export function releaseRegionCapture(
  capture: Pick<OfficeRegionCapture, "release"> | null | undefined,
): void {
  capture?.release()
}
