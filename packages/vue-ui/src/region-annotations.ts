import type {
  NormalizedPoint,
  NormalizedRect,
  RegionAnnotation,
} from "@arcships/office-interaction"
import type { OfficeRegionCaptureMetadata } from "./region-capture"

export const OFFICE_REGION_ANNOTATION_SCHEMA_VERSION = 1

export const OFFICE_REGION_ANNOTATION_KINDS = [
  "rectangle",
  "arrow",
  "text",
  "highlight",
] as const

export type OfficeRegionAnnotationKind = typeof OFFICE_REGION_ANNOTATION_KINDS[number]

/**
 * Annotation styling. `strokeWidth` and `fontSize` are normalized against the
 * shorter edge of the captured bitmap, so an annotation keeps its visual
 * weight when the same bitmap is displayed at 1x or 2x.
 */
export interface OfficeRegionAnnotationStyle {
  color?: string
  opacity?: number
  /** Fraction of the shorter image edge. Defaults to 0.006. */
  strokeWidth?: number
  /** Fraction of the shorter image edge. Defaults to 0.05. */
  fontSize?: number
}

interface OfficeRegionAnnotationBase {
  id: string
  style?: OfficeRegionAnnotationStyle
}

export interface OfficeRegionRectangleAnnotation extends OfficeRegionAnnotationBase {
  kind: "rectangle"
  rect: NormalizedRect
}

export interface OfficeRegionHighlightAnnotation extends OfficeRegionAnnotationBase {
  kind: "highlight"
  rect: NormalizedRect
}

export interface OfficeRegionArrowAnnotation extends OfficeRegionAnnotationBase {
  kind: "arrow"
  start: NormalizedPoint
  end: NormalizedPoint
}

export interface OfficeRegionTextAnnotation extends OfficeRegionAnnotationBase {
  kind: "text"
  point: NormalizedPoint
  text: string
}

export type OfficeRegionAnnotation =
  | OfficeRegionRectangleAnnotation
  | OfficeRegionHighlightAnnotation
  | OfficeRegionArrowAnnotation
  | OfficeRegionTextAnnotation

export type OfficeRegionAnnotationInput =
  | Omit<OfficeRegionRectangleAnnotation, "id">
  | Omit<OfficeRegionHighlightAnnotation, "id">
  | Omit<OfficeRegionArrowAnnotation, "id">
  | Omit<OfficeRegionTextAnnotation, "id">

/**
 * Serializable annotation document. Coordinates are normalized to the selected
 * region's own width/height (`[0, 1]`), independent of the capture scale or DPR.
 */
export interface OfficeRegionAnnotationDocument {
  schemaVersion: 1
  annotations: OfficeRegionAnnotation[]
  /** Capture provenance; never required for rendering. */
  source?: OfficeRegionCaptureMetadata
}

export type OfficeRegionAnnotationDocumentInput = {
  schemaVersion?: 1
  annotations?: readonly OfficeRegionAnnotation[]
  source?: OfficeRegionCaptureMetadata
}

export type OfficeRegionAnnotationValidationCode =
  | "INVALID_DOCUMENT"
  | "INVALID_ANNOTATION"
  | "INVALID_STYLE"
  | "INVALID_COORDINATE"
  | "DUPLICATE_ID"

export class OfficeRegionAnnotationValidationError extends Error {
  readonly code: OfficeRegionAnnotationValidationCode
  readonly path: string

  constructor(code: OfficeRegionAnnotationValidationCode, message: string, path = "document") {
    super(message)
    this.name = "OfficeRegionAnnotationValidationError"
    this.code = code
    this.path = path
  }
}

export const DEFAULT_REGION_ANNOTATION_MAX_COUNT = 500

function fail(code: OfficeRegionAnnotationValidationCode, path: string, message: string): never {
  throw new OfficeRegionAnnotationValidationError(code, message, path)
}

function assertUnit(value: unknown, path: string): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0 || value > 1) {
    fail("INVALID_COORDINATE", path, "coordinate must be a finite number in [0, 1]")
  }
  return value
}

function normalizePoint(value: unknown, path: string): NormalizedPoint {
  if (!value || typeof value !== "object") fail("INVALID_COORDINATE", path, "point must be an object")
  const point = value as NormalizedPoint
  return { x: assertUnit(point.x, `${path}.x`), y: assertUnit(point.y, `${path}.y`) }
}

function normalizeRect(value: unknown, path: string): NormalizedRect {
  if (!value || typeof value !== "object") fail("INVALID_COORDINATE", path, "rectangle must be an object")
  const rect = value as NormalizedRect
  const x = assertUnit(rect.x, `${path}.x`)
  const y = assertUnit(rect.y, `${path}.y`)
  const width = assertUnit(rect.width, `${path}.width`)
  const height = assertUnit(rect.height, `${path}.height`)
  if (width <= 0 || height <= 0) fail("INVALID_COORDINATE", path, "rectangle must have a positive area")
  if (x + width > 1 + Number.EPSILON || y + height > 1 + Number.EPSILON) {
    fail("INVALID_COORDINATE", path, "rectangle must stay inside [0, 1]")
  }
  return { x, y, width, height }
}

function normalizeStyle(value: unknown, path: string): OfficeRegionAnnotationStyle | undefined {
  if (value === undefined) return undefined
  if (!value || typeof value !== "object") fail("INVALID_STYLE", path, "style must be an object")
  const style = value as OfficeRegionAnnotationStyle & Record<string, unknown>
  const result: OfficeRegionAnnotationStyle = {}
  if (style.color !== undefined) {
    if (typeof style.color !== "string" || !style.color.trim() || style.color.length > 64) {
      fail("INVALID_STYLE", `${path}.color`, "color must be a non-empty CSS color string")
    }
    result.color = style.color.trim()
  }
  if (style.opacity !== undefined) {
    if (typeof style.opacity !== "number" || !Number.isFinite(style.opacity) || style.opacity < 0 || style.opacity > 1) {
      fail("INVALID_STYLE", `${path}.opacity`, "opacity must be in [0, 1]")
    }
    result.opacity = style.opacity
  }
  if (style.strokeWidth !== undefined) {
    if (typeof style.strokeWidth !== "number" || !Number.isFinite(style.strokeWidth) || style.strokeWidth <= 0 || style.strokeWidth > 0.25) {
      fail("INVALID_STYLE", `${path}.strokeWidth`, "strokeWidth must be in (0, 0.25]")
    }
    result.strokeWidth = style.strokeWidth
  }
  if (style.fontSize !== undefined) {
    if (typeof style.fontSize !== "number" || !Number.isFinite(style.fontSize) || style.fontSize <= 0 || style.fontSize > 0.5) {
      fail("INVALID_STYLE", `${path}.fontSize`, "fontSize must be in (0, 0.5]")
    }
    result.fontSize = style.fontSize
  }
  return Object.keys(result).length ? result : undefined
}

function normalizeAnnotation(value: unknown, index: number): OfficeRegionAnnotation {
  const path = `annotations[${index}]`
  if (!value || typeof value !== "object") fail("INVALID_ANNOTATION", path, "annotation must be an object")
  const input = value as Record<string, unknown>
  if (typeof input.id !== "string" || !input.id.trim()) {
    fail("INVALID_ANNOTATION", `${path}.id`, "id must be a non-empty string")
  }
  const id = input.id.trim()
  const style = normalizeStyle(input.style, `${path}.style`)
  const base = style ? { id, style } : { id }
  switch (input.kind) {
    case "rectangle":
      return { ...base, kind: "rectangle", rect: normalizeRect(input.rect, `${path}.rect`) }
    case "highlight":
      return { ...base, kind: "highlight", rect: normalizeRect(input.rect, `${path}.rect`) }
    case "arrow": {
      const start = normalizePoint(input.start, `${path}.start`)
      const end = normalizePoint(input.end, `${path}.end`)
      if (Math.abs(start.x - end.x) < Number.EPSILON && Math.abs(start.y - end.y) < Number.EPSILON) {
        fail("INVALID_COORDINATE", path, "arrow must have two distinct endpoints")
      }
      return { ...base, kind: "arrow", start, end }
    }
    case "text": {
      if (typeof input.text !== "string" || !input.text.trim()) {
        fail("INVALID_ANNOTATION", `${path}.text`, "text must be a non-empty string")
      }
      if (input.text.length > 4096) {
        fail("INVALID_ANNOTATION", `${path}.text`, "text must be at most 4096 characters")
      }
      return { ...base, kind: "text", point: normalizePoint(input.point, `${path}.point`), text: input.text }
    }
    default:
      fail("INVALID_ANNOTATION", `${path}.kind`, "kind must be rectangle, arrow, text, or highlight")
  }
}

function normalizeSource(value: unknown): OfficeRegionCaptureMetadata | undefined {
  if (value === undefined) return undefined
  if (!value || typeof value !== "object") fail("INVALID_DOCUMENT", "source", "source must be an object")
  const source = value as OfficeRegionCaptureMetadata
  if (typeof source.capturedAt !== "string" || !source.capturedAt) {
    fail("INVALID_DOCUMENT", "source.capturedAt", "source.capturedAt must be an ISO timestamp")
  }
  if (typeof source.width !== "number" || typeof source.height !== "number"
    || !Number.isFinite(source.width) || !Number.isFinite(source.height)
    || source.width <= 0 || source.height <= 0) {
    fail("INVALID_DOCUMENT", "source", "source must include positive width and height")
  }
  return { ...source, region: normalizeRect(source.region, "source.region") }
}

/**
 * Validate and canonicalize an annotation document. Unknown fields are
 * dropped, so the parsed value is the canonical form that `stringify` emits.
 */
export function parseOfficeRegionAnnotationDocument(
  input: unknown,
  options: { maxAnnotations?: number } = {},
): OfficeRegionAnnotationDocument {
  let value = input
  if (typeof value === "string") {
    try {
      value = JSON.parse(value) as unknown
    } catch (reason) {
      fail("INVALID_DOCUMENT", "document", reason instanceof Error ? `invalid JSON: ${reason.message}` : "invalid JSON")
    }
  }
  if (!value || typeof value !== "object") fail("INVALID_DOCUMENT", "document", "annotation document must be an object")
  const document = value as Record<string, unknown>
  if (document.schemaVersion !== OFFICE_REGION_ANNOTATION_SCHEMA_VERSION) {
    fail("INVALID_DOCUMENT", "schemaVersion", `schemaVersion must be ${OFFICE_REGION_ANNOTATION_SCHEMA_VERSION}`)
  }
  if (!Array.isArray(document.annotations)) {
    fail("INVALID_DOCUMENT", "annotations", "annotations must be an array")
  }
  const maxAnnotations = options.maxAnnotations ?? DEFAULT_REGION_ANNOTATION_MAX_COUNT
  if (document.annotations.length > maxAnnotations) {
    fail("INVALID_DOCUMENT", "annotations", `annotations must not exceed ${maxAnnotations} entries`)
  }
  const seen = new Set<string>()
  const annotations = document.annotations.map((annotation, index) => {
    const normalized = normalizeAnnotation(annotation, index)
    if (seen.has(normalized.id)) fail("DUPLICATE_ID", `annotations[${index}].id`, `duplicate annotation id: ${normalized.id}`)
    seen.add(normalized.id)
    return normalized
  })
  const source = normalizeSource(document.source)
  return {
    schemaVersion: OFFICE_REGION_ANNOTATION_SCHEMA_VERSION,
    annotations,
    ...(source ? { source } : {}),
  }
}

export function createOfficeRegionAnnotationDocument(
  input: OfficeRegionAnnotationDocumentInput = {},
): OfficeRegionAnnotationDocument {
  return parseOfficeRegionAnnotationDocument({
    schemaVersion: input.schemaVersion ?? OFFICE_REGION_ANNOTATION_SCHEMA_VERSION,
    annotations: input.annotations ?? [],
    ...(input.source ? { source: input.source } : {}),
  })
}

export function stringifyOfficeRegionAnnotationDocument(
  document: OfficeRegionAnnotationDocument,
): string {
  return JSON.stringify(parseOfficeRegionAnnotationDocument(document))
}

export function cloneOfficeRegionAnnotationDocument(
  document: OfficeRegionAnnotationDocument,
): OfficeRegionAnnotationDocument {
  return parseOfficeRegionAnnotationDocument(JSON.parse(JSON.stringify(document)))
}

let annotationCounter = 0

/** Stable-enough per-session id; hosts may also supply their own ids. */
export function createOfficeRegionAnnotationId(prefix = "annotation"): string {
  const cryptoObject = (globalThis as { crypto?: { randomUUID?: () => string } }).crypto
  if (typeof cryptoObject?.randomUUID === "function") return `${prefix}-${cryptoObject.randomUUID()}`
  annotationCounter += 1
  return `${prefix}-${Date.now().toString(36)}-${annotationCounter.toString(36)}`
}

/** Bounding box in normalized image coordinates; text bounds are a point. */
export function officeRegionAnnotationBounds(annotation: OfficeRegionAnnotation): NormalizedRect {
  if (annotation.kind === "rectangle" || annotation.kind === "highlight") return { ...annotation.rect }
  if (annotation.kind === "arrow") {
    const x = Math.min(annotation.start.x, annotation.end.x)
    const y = Math.min(annotation.start.y, annotation.end.y)
    return {
      x,
      y,
      width: Math.abs(annotation.end.x - annotation.start.x),
      height: Math.abs(annotation.end.y - annotation.start.y),
    }
  }
  return { x: annotation.point.x, y: annotation.point.y, width: 0, height: 0 }
}

/** Convert the format-neutral geometry type from `office-interaction` when possible. */
export function officeRegionAnnotationFromRegionAnnotation(
  annotation: RegionAnnotation,
): OfficeRegionAnnotation | null {
  const style = annotation.color ? { color: annotation.color } : undefined
  const base = style ? { id: annotation.id, style } : { id: annotation.id }
  if (annotation.kind === "rectangle") return { ...base, kind: "rectangle", rect: { ...annotation.rect } }
  if (annotation.kind === "arrow") {
    return { ...base, kind: "arrow", start: { ...annotation.start }, end: { ...annotation.end } }
  }
  if (annotation.kind === "text") {
    return { ...base, kind: "text", point: { ...annotation.point }, text: annotation.text }
  }
  // Ellipse, freehand, and pin are not part of the annotator's compact tool set.
  return null
}
