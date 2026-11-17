import {
  PdfActionType,
  PdfAnnotationSubtype,
  type PdfAnnotationObject,
  type PdfLinkTarget,
} from "@embedpdf/models"
import { pdfIntrinsicRectToCanonical } from "./pdf-coordinate-transform"
import type { PdfRenderAnnotation, PdfRenderPageInfo } from "./pdf-render-runtime"

/** Protocols a PDF link may target before the host applies its own URL policy. */
export const PDF_LINK_PROTOCOLS: readonly string[] = ["http:", "https:", "mailto:"]

/** Annotation flags that make a widget or link non-interactive. */
const HIDDEN_ANNOTATION_FLAGS = new Set(["hidden", "invisible", "noView"])

export interface PdfLinkUrlResolution {
  url?: string
  blocked: boolean
}

/**
 * Validate a PDF link URI. Only absolute `http`, `https` and `mailto` links are
 * accepted; embedded credentials are stripped. Everything else is reported as
 * blocked instead of silently producing a dead or dangerous target.
 */
export function resolvePdfLinkUrl(rawUrl: string | undefined): PdfLinkUrlResolution {
  const trimmed = rawUrl?.trim()
  if (!trimmed) return { blocked: true }
  let url: URL
  try {
    url = new URL(trimmed)
  } catch {
    return { blocked: true }
  }
  if (!PDF_LINK_PROTOCOLS.includes(url.protocol.toLowerCase())) return { blocked: true }
  url.username = ""
  url.password = ""
  return { url: url.href, blocked: false }
}

function readLinkTarget(target: PdfLinkTarget | undefined): {
  url?: string
  targetPageIndex?: number
  blocked?: boolean
} {
  if (!target) return { blocked: true }
  if (target.type === "destination") {
    const pageIndex = target.destination?.pageIndex
    return Number.isSafeInteger(pageIndex) && pageIndex >= 0 ? { targetPageIndex: pageIndex } : { blocked: true }
  }
  const action = target.action
  if (!action) return { blocked: true }
  switch (action.type) {
    case PdfActionType.URI: {
      const resolved = resolvePdfLinkUrl(action.uri)
      return resolved.url ? { url: resolved.url } : { blocked: true }
    }
    case PdfActionType.Goto: {
      const pageIndex = action.destination?.pageIndex
      return Number.isSafeInteger(pageIndex) && pageIndex >= 0 ? { targetPageIndex: pageIndex } : { blocked: true }
    }
    default:
      // Remote go-to, launch/file/open and unsupported actions never become
      // clickable targets in the current document.
      return { blocked: true }
  }
}

/**
 * Convert engine annotations into the hit-layer model: canonical rectangles plus
 * the link destination, form-field identity or annotation text a host needs.
 */
export function mapPdfAnnotationObjects(
  annotations: readonly PdfAnnotationObject[],
  page: PdfRenderPageInfo,
): PdfRenderAnnotation[] {
  const mapped: PdfRenderAnnotation[] = []
  for (const annotation of annotations) {
    if (!annotation || typeof annotation !== "object") continue
    const flags = annotation.flags ?? []
    if (flags.some((flag) => HIDDEN_ANNOTATION_FLAGS.has(flag))) continue
    const rect = annotation.rect
    if (!rect || !rect.origin || !rect.size) continue
    const canonical: PdfRenderAnnotation = {
      id: annotation.id,
      pageIndex: page.index,
      kind: "annotation",
      rect: pdfIntrinsicRectToCanonical(page, {
        x: rect.origin.x,
        y: rect.origin.y,
        width: Math.abs(rect.size.width),
        height: Math.abs(rect.size.height),
      }),
      ...(flags.length ? { flags: [...flags] } : {}),
    }
    if (annotation.type === PdfAnnotationSubtype.LINK) {
      const target = readLinkTarget(annotation.target)
      mapped.push({
        ...canonical,
        kind: "link",
        ...(target.url ? { url: target.url } : {}),
        ...(target.targetPageIndex !== undefined ? { targetPageIndex: target.targetPageIndex } : {}),
        ...(target.blocked ? { blocked: true } : {}),
      })
      continue
    }
    if (annotation.type === PdfAnnotationSubtype.WIDGET) {
      const field = annotation.field
      mapped.push({
        ...canonical,
        kind: "widget",
        ...(field?.name ? { fieldName: field.name } : {}),
        ...(field?.value ? { fieldValue: field.value } : {}),
      })
      continue
    }
    const contents = annotation.contents ?? annotation.subject
    mapped.push({
      ...canonical,
      ...(contents ? { contents } : {}),
    })
  }
  return mapped
}
