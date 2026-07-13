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

/** Convert an engine search rect, which includes intrinsic page rotation, back to canonical space. */
export function pdfIntrinsicRectToCanonical(
  page: PdfRenderPageInfo,
  rect: PdfRenderRect,
): PdfRenderRect {
  switch (page.rotation) {
    case 90:
      return {
        x: rect.y,
        y: page.height - rect.x - rect.width,
        width: rect.height,
        height: rect.width,
      }
    case 180:
      return {
        x: page.width - rect.x - rect.width,
        y: page.height - rect.y - rect.height,
        width: rect.width,
        height: rect.height,
      }
    case 270:
      return {
        x: page.width - rect.y - rect.height,
        y: rect.x,
        width: rect.height,
        height: rect.width,
      }
    default:
      return rect
  }
}
