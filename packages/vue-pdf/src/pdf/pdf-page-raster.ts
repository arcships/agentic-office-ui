/**
 * Per-page raster pixel budget for `PdfSurface`.
 *
 * A Letter page at 200 % zoom and DPR 2 asks the engine for roughly 2 000 ×
 * 2 600 px (~21 MB of RGBA) per page. A handful of those already exhausts a
 * mobile tab, so the surface renders each page at a capped pixel budget and
 * lets CSS downscale the bitmap to the slot size. The cap only ever *reduces*
 * the requested scale: sharp rendering below the budget is untouched.
 *
 * Pure math, no DOM and no engine dependency.
 */

/** Default cap per page (8 MP ≈ 32 MB of RGBA). */
export const DEFAULT_MAX_PAGE_RASTER_PIXELS = 8_000_000

/** Bytes one rasterized pixel occupies in the decoded bitmap (RGBA). */
export const PDF_RASTER_BYTES_PER_PIXEL = 4

export interface PdfPageRasterRequest {
  /** Display width in CSS px at the requested zoom, rotation applied. */
  displayWidth: number
  /** Display height in CSS px at the requested zoom, rotation applied. */
  displayHeight: number
  /** Zoom factor the surface would otherwise request from the runtime. */
  zoom: number
  /** Device pixel ratio the surface would otherwise request. */
  dpr: number
  /** Per-page pixel cap. Defaults to {@link DEFAULT_MAX_PAGE_RASTER_PIXELS}. */
  maxPixels?: number
}

export interface PdfPageRasterPlan {
  /** Scale factor to request from the runtime; never above the requested zoom. */
  scale: number
  /** Device pixel ratio to request from the runtime. */
  dpr: number
  /** Raster width in device px. */
  pixelWidth: number
  /** Raster height in device px. */
  pixelHeight: number
  /** `pixelWidth * pixelHeight`, always at or below the cap. */
  pixels: number
  /** Estimated decoded RGBA footprint of the bitmap in bytes. */
  bytes: number
  /** True when the cap reduced the requested scale. */
  capped: boolean
}

function positive(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? value : fallback
}

/**
 * Resolve the scale/dpr pair a page should rasterize at.
 *
 * `displayWidth`/`displayHeight` already include zoom and rotation, so the
 * bitmap the engine produces is `displayWidth * scale / zoom * dpr` wide. When
 * that exceeds `maxPixels` the scale shrinks by the square root of the
 * overshoot and the returned `scale` stays proportional to the requested zoom,
 * which keeps the rendered extent identical to the layout box.
 */
export function planPageRaster(request: PdfPageRasterRequest): PdfPageRasterPlan {
  const zoom = positive(request.zoom, 1)
  const dpr = positive(request.dpr, 1)
  const maxPixels = positive(request.maxPixels, DEFAULT_MAX_PAGE_RASTER_PIXELS)
  const displayWidth = positive(request.displayWidth, 0)
  const displayHeight = positive(request.displayHeight, 0)

  if (displayWidth === 0 || displayHeight === 0) {
    return { scale: zoom, dpr, pixelWidth: 0, pixelHeight: 0, pixels: 0, bytes: 0, capped: false }
  }

  const pixelWidth = Math.max(1, Math.round(displayWidth * dpr))
  const pixelHeight = Math.max(1, Math.round(displayHeight * dpr))
  const pixels = pixelWidth * pixelHeight
  if (pixels <= maxPixels) {
    return {
      scale: zoom,
      dpr,
      pixelWidth,
      pixelHeight,
      pixels,
      bytes: pixels * PDF_RASTER_BYTES_PER_PIXEL,
      capped: false,
    }
  }

  const shrink = Math.sqrt(maxPixels / pixels)
  let cappedWidth = Math.max(1, Math.floor(displayWidth * dpr * shrink))
  let cappedHeight = Math.max(1, Math.floor(displayHeight * dpr * shrink))
  if (cappedWidth > maxPixels) cappedWidth = Math.max(1, Math.floor(maxPixels))
  if (cappedHeight > maxPixels) cappedHeight = Math.max(1, Math.floor(maxPixels))
  if (cappedWidth * cappedHeight > maxPixels) {
    // Rounding can still overshoot a tiny cap; the budget wins and the other
    // axis absorbs the difference.
    cappedHeight = Math.max(1, Math.floor(maxPixels / cappedWidth))
  }
  const cappedPixels = cappedWidth * cappedHeight
  return {
    scale: zoom * shrink,
    dpr,
    pixelWidth: cappedWidth,
    pixelHeight: cappedHeight,
    pixels: cappedPixels,
    bytes: cappedPixels * PDF_RASTER_BYTES_PER_PIXEL,
    capped: true,
  }
}
