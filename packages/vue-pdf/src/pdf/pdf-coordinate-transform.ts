import type { Position } from "@embedpdf/models"
import type { PdfRenderPageInfo, PdfRenderRect, PdfRotation } from "./pdf-render-runtime"

export function pdfQuarterTurns(page: PdfRenderPageInfo, userRotation: PdfRotation): number {
  return ((page.rotation + userRotation) / 90) % 4
}

export function pdfDisplaySize(
  page: PdfRenderPageInfo,
  userRotation: PdfRotation,
  scale: number,
): { width: number; height: number } {
  const quarterTurns = pdfQuarterTurns(page, userRotation)
  return {
    width: (quarterTurns % 2 ? page.height : page.width) * scale,
    height: (quarterTurns % 2 ? page.width : page.height) * scale,
  }
}

/**
 * CSS transform mapping canonical page coordinates onto the displayed page.
 *
 * A text layer keeps spans in canonical units (PDFium glyph boxes) and applies
 * this matrix once per page, so zoom and rotation scale and rotate the text with
 * the bitmap instead of re-positioning every span. The matrix mirrors
 * `pdfCanonicalRectToDisplay` for all four quarter turns.
 */
export function pdfCanonicalToDisplayTransform(
  page: PdfRenderPageInfo,
  userRotation: PdfRotation,
  scale: number,
): string {
  const quarterTurns = pdfQuarterTurns(page, userRotation)
  if (quarterTurns === 1) {
    return `matrix(0, ${scale}, ${-scale}, 0, ${page.height * scale}, 0)`
  }
  if (quarterTurns === 2) {
    return `matrix(${-scale}, 0, 0, ${-scale}, ${page.width * scale}, ${page.height * scale})`
  }
  if (quarterTurns === 3) {
    return `matrix(0, ${-scale}, ${scale}, 0, 0, ${page.width * scale})`
  }
  return `matrix(${scale}, 0, 0, ${scale}, 0, 0)`
}

export function pdfCanonicalRectToDisplay(
  page: PdfRenderPageInfo,
  rect: PdfRenderRect,
  userRotation: PdfRotation,
  scale: number,
): PdfRenderRect {
  const quarterTurns = pdfQuarterTurns(page, userRotation)
  if (quarterTurns === 1) {
    return {
      x: (page.height - rect.y - rect.height) * scale,
      y: rect.x * scale,
      width: rect.height * scale,
      height: rect.width * scale,
    }
  }
  if (quarterTurns === 2) {
    return {
      x: (page.width - rect.x - rect.width) * scale,
      y: (page.height - rect.y - rect.height) * scale,
      width: rect.width * scale,
      height: rect.height * scale,
    }
  }
  if (quarterTurns === 3) {
    return {
      x: rect.y * scale,
      y: (page.width - rect.x - rect.width) * scale,
      width: rect.height * scale,
      height: rect.width * scale,
    }
  }
  return {
    x: rect.x * scale,
    y: rect.y * scale,
    width: rect.width * scale,
    height: rect.height * scale,
  }
}

export function pdfDisplayPointToCanonical(
  page: PdfRenderPageInfo,
  point: Position,
  userRotation: PdfRotation,
  scale: number,
): Position {
  const displayX = point.x / scale
  const displayY = point.y / scale
  const quarterTurns = pdfQuarterTurns(page, userRotation)
  let x = displayX
  let y = displayY
  if (quarterTurns === 1) {
    x = displayY
    y = page.height - displayX
  } else if (quarterTurns === 2) {
    x = page.width - displayX
    y = page.height - displayY
  } else if (quarterTurns === 3) {
    x = page.width - displayY
    y = displayX
  }
  return {
    x: Math.min(page.width, Math.max(0, x)),
    y: Math.min(page.height, Math.max(0, y)),
  }
}

/**
 * The runtime opens documents with `normalizeRotation: true`, so every engine
 * rect (search hits, annotations) is already in canonical 0-degree page space.
 * Intrinsic `/Rotate` values are exposed through `page.rotation` for display
 * math and diagnostics but must not be applied a second time here.
 */
export function pdfIntrinsicRectToCanonical(
  _page: PdfRenderPageInfo,
  rect: PdfRenderRect,
): PdfRenderRect {
  return rect
}
