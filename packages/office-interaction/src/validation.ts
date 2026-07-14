import type {
  DocxBehaviorLocator,
  DocxObjectLocator,
  DocxPartLocator,
  DocxPathSegment,
  ManualRegionLocator,
  NormalizedRect,
  OfficeDocumentRevision,
  OfficeFormat,
  OfficeObjectKind,
  OfficeObjectReference,
  OfficeObjectReliability,
  OfficeReferenceConfirmEvent,
  OfficeReferenceSnapshot,
  PdfObjectLocator,
  PptxObjectKeyLocator,
  PptxObjectLocator,
  PptxSlideLocator,
  RecognitionSource,
  ReliabilityDimension,
  SheetRegionPoint,
  XlsxBehaviorLocator,
  XlsxObjectLocator,
  XlsxSheetLocator,
} from "./types"

export type OfficeValidationIssueCode =
  | "INVALID_JSON_VALUE"
  | "INVALID_TYPE"
  | "INVALID_VALUE"
  | "INVALID_SCHEMA_VERSION"
  | "FORMAT_MISMATCH"
  | "KIND_LOCATOR_MISMATCH"

export interface OfficeValidationIssue {
  code: OfficeValidationIssueCode
  path: string
  message: string
}

export class OfficeInteractionValidationError extends TypeError {
  readonly code = "INVALID_OFFICE_INTERACTION" as const
  readonly issues: readonly OfficeValidationIssue[]

  constructor(issue: OfficeValidationIssue | readonly OfficeValidationIssue[]) {
    const issues = Array.isArray(issue) ? issue : [issue]
    super(issues.map((entry) => `${entry.path}: ${entry.message}`).join("; "))
    this.name = "OfficeInteractionValidationError"
    this.issues = issues
  }
}

export type OfficeSafeParseResult<T> =
  | { success: true; data: T }
  | { success: false; error: OfficeInteractionValidationError }

type RecordValue = Record<string, unknown>

const formats = new Set<OfficeFormat>(["docx", "xlsx", "pptx", "pdf"])
const recognitionSources = new Set<RecognitionSource>(["native", "structural", "visual", "manual"])
const reliabilityLevels = new Set(["exact", "likely", "uncertain", "unknown"])
const selectionTriggers = new Set(["pointer", "keyboard", "touch", "programmatic"])
const objectKinds = new Set<OfficeObjectKind>([
  "document", "page", "section", "paragraph", "heading", "list-item", "text-range", "text-block",
  "table", "table-row", "table-cell", "image", "workbook", "worksheet", "cell", "cell-range", "row", "column",
  "chart", "chart-series", "chart-point", "chart-legend-entry", "shape", "text-box", "group", "slide", "layout",
  "master", "comment", "annotation", "link", "form-control", "region", "formula", "style", "conditional-format",
  "data-validation", "animation", "transition", "action", "note", "bookmark", "field", "named-range", "protection",
  "tracked-change", "cross-reference", "form-rule", "attachment", "layer", "signature", "permission", "unknown",
])
const reasonCodePattern = /^[a-z0-9][a-z0-9._-]*$/u
const sha256Pattern = /^sha256:[a-f0-9]{64}$/u
const a1Pattern = /^\$?[A-Z]{1,4}\$?[1-9][0-9]*(?::\$?[A-Z]{1,4}\$?[1-9][0-9]*)?$/iu
const fingerprintTextLimit = 256

function fail(code: OfficeValidationIssueCode, path: string, message: string): never {
  throw new OfficeInteractionValidationError({ code, path, message })
}

function isPlainRecord(value: unknown): value is RecordValue {
  if (value === null || typeof value !== "object" || Array.isArray(value)) return false
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function record(value: unknown, path: string): RecordValue {
  if (!isPlainRecord(value)) fail("INVALID_TYPE", path, "expected a plain object")
  return value
}

function onlyKeys(value: RecordValue, allowed: readonly string[], path: string): void {
  const allowedKeys = new Set(allowed)
  const unexpected = Object.keys(value).find((key) => !allowedKeys.has(key))
  if (unexpected !== undefined) fail("INVALID_VALUE", `${path}.${unexpected}`, "is not part of this contract")
}

function array(value: unknown, path: string): readonly unknown[] {
  if (!Array.isArray(value)) fail("INVALID_TYPE", path, "expected an array")
  return value
}

function stringValue(value: unknown, path: string, options: { max?: number; pattern?: RegExp } = {}): string {
  if (typeof value !== "string" || value.length === 0) fail("INVALID_TYPE", path, "expected a non-empty string")
  if (options.max !== undefined && value.length > options.max) {
    fail("INVALID_VALUE", path, `must contain at most ${options.max} characters`)
  }
  if (options.pattern && !options.pattern.test(value)) fail("INVALID_VALUE", path, "has an invalid format")
  return value
}

function optionalString(value: unknown, path: string, max = 512): string | undefined {
  if (value === undefined) return undefined
  return stringValue(value, path, { max })
}

function integer(value: unknown, path: string): number {
  if (typeof value !== "number" || !Number.isSafeInteger(value) || value < 0) {
    fail("INVALID_TYPE", path, "expected a non-negative safe integer")
  }
  return value
}

function finite(value: unknown, path: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) fail("INVALID_TYPE", path, "expected a finite number")
  return value
}

function unit(value: unknown, path: string): number {
  const result = finite(value, path)
  if (result < 0 || result > 1) fail("INVALID_VALUE", path, "must be between 0 and 1")
  return result
}

function booleanValue(value: unknown, path: string): boolean {
  if (typeof value !== "boolean") fail("INVALID_TYPE", path, "expected a boolean")
  return value
}

function enumValue<T extends string>(value: unknown, values: ReadonlySet<T>, path: string): T {
  if (typeof value !== "string" || !values.has(value as T)) fail("INVALID_VALUE", path, "contains an unsupported value")
  return value as T
}

function assertJsonTree(value: unknown, path = "$", seen = new Set<object>()): void {
  if (value === null || typeof value === "boolean") return
  if (typeof value === "string") {
    if (value.startsWith("blob:")) fail("INVALID_JSON_VALUE", path, "Blob URLs are not serializable references")
    return
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) fail("INVALID_JSON_VALUE", path, "non-finite numbers are not JSON values")
    return
  }
  if (typeof value !== "object") fail("INVALID_JSON_VALUE", path, "contains a non-JSON value")
  if (seen.has(value)) fail("INVALID_JSON_VALUE", path, "contains a circular reference")
  seen.add(value)
  if (Array.isArray(value)) {
    value.forEach((item, index) => assertJsonTree(item, `${path}[${index}]`, seen))
  } else {
    if (!isPlainRecord(value)) fail("INVALID_JSON_VALUE", path, "contains a non-plain object")
    for (const [key, item] of Object.entries(value)) assertJsonTree(item, `${path}.${key}`, seen)
  }
  seen.delete(value)
}

function validateRect(value: unknown, path: string): NormalizedRect {
  const input = record(value, path)
  const x = unit(input.x, `${path}.x`)
  const y = unit(input.y, `${path}.y`)
  const width = unit(input.width, `${path}.width`)
  const height = unit(input.height, `${path}.height`)
  if (width <= 0 || height <= 0) fail("INVALID_VALUE", path, "width and height must be greater than zero")
  if (x + width > 1 + Number.EPSILON || y + height > 1 + Number.EPSILON) {
    fail("INVALID_VALUE", path, "rectangle must fit inside its normalized coordinate space")
  }
  return input as unknown as NormalizedRect
}

function validateSheetPoint(value: unknown, path: string): SheetRegionPoint {
  const input = record(value, path)
  integer(input.row, `${path}.row`)
  integer(input.col, `${path}.col`)
  unit(input.xOffset, `${path}.xOffset`)
  unit(input.yOffset, `${path}.yOffset`)
  return input as unknown as SheetRegionPoint
}

function validateManualRegion(value: unknown, path: string): ManualRegionLocator {
  const input = record(value, path)
  if (input.space === "page") {
    integer(input.pageIndex, `${path}.pageIndex`)
    validateRect(input.rect, `${path}.rect`)
  } else if (input.space === "slide") {
    integer(input.slideIndex, `${path}.slideIndex`)
    validateRect(input.rect, `${path}.rect`)
  } else if (input.space === "sheet") {
    stringValue(input.sheetId, `${path}.sheetId`, { max: 512 })
    const start = validateSheetPoint(input.start, `${path}.start`)
    const end = validateSheetPoint(input.end, `${path}.end`)
    if (start.row > end.row || start.col > end.col) {
      fail("INVALID_VALUE", path, "sheet region start must not follow its end")
    }
    if (start.row === end.row && start.yOffset > end.yOffset) {
      fail("INVALID_VALUE", path, "sheet region y offsets are reversed")
    }
    if (start.col === end.col && start.xOffset > end.xOffset) {
      fail("INVALID_VALUE", path, "sheet region x offsets are reversed")
    }
    const rowSpan = end.row - start.row + end.yOffset - start.yOffset
    const colSpan = end.col - start.col + end.xOffset - start.xOffset
    if (rowSpan <= 0 || colSpan <= 0) {
      fail("INVALID_VALUE", path, "sheet region must have a positive area")
    }
  } else {
    fail("INVALID_VALUE", `${path}.space`, "contains an unsupported coordinate space")
  }
  return input as unknown as ManualRegionLocator
}

function validatePart(value: unknown, path: string): DocxPartLocator {
  const input = record(value, path)
  if (input.kind === "body") return input as unknown as DocxPartLocator
  if (input.kind === "header" || input.kind === "footer") {
    stringValue(input.partName, `${path}.partName`, { max: 512 })
    return input as unknown as DocxPartLocator
  }
  if (input.kind === "footnote" || input.kind === "endnote") {
    stringValue(input.noteId, `${path}.noteId`, { max: 256 })
    return input as unknown as DocxPartLocator
  }
  fail("INVALID_VALUE", `${path}.kind`, "contains an unsupported DOCX part")
}

function validateDocxPath(value: unknown, path: string, allowEmpty = false): readonly DocxPathSegment[] {
  const input = array(value, path)
  if (!allowEmpty && input.length === 0) fail("INVALID_VALUE", path, "must not be empty")
  input.forEach((item, index) => {
    const segment = record(item, `${path}[${index}]`)
    enumValue(segment.kind, new Set(["node", "table", "row", "cell", "paragraph", "run"]), `${path}[${index}].kind`)
    integer(segment.index, `${path}[${index}].index`)
    optionalString(segment.nativeId, `${path}[${index}].nativeId`, 256)
  })
  return input as readonly DocxPathSegment[]
}

function validateDocxBehavior(value: RecordValue, path: string): DocxBehaviorLocator {
  const behavior = enumValue(value.behavior, new Set([
    "style", "field", "bookmark", "section", "tracked-change-state", "comment-state", "cross-reference",
  ]), `${path}.behavior`)
  stringValue(value.instanceId, `${path}.instanceId`, { max: 512 })
  const owner = record(value.owner, `${path}.owner`)
  if (owner.scope === "document") {
    // instanceId keeps multiple document-wide behaviors unambiguous.
  } else if (owner.scope === "part") {
    validatePart(owner.part, `${path}.owner.part`)
    validateDocxPath(owner.path, `${path}.owner.path`)
  } else {
    fail("INVALID_VALUE", `${path}.owner.scope`, "contains an unsupported DOCX behavior scope")
  }
  return value as unknown as DocxBehaviorLocator
}

function validateDocxLocator(value: unknown, path: string): DocxObjectLocator {
  const input = record(value, path)
  switch (input.kind) {
    case "page":
      integer(input.pageIndex, `${path}.pageIndex`)
      break
    case "structure":
      validatePart(input.part, `${path}.part`)
      validateDocxPath(input.path, `${path}.path`)
      break
    case "text-range": {
      validatePart(input.part, `${path}.part`)
      const start = record(input.start, `${path}.start`)
      const end = record(input.end, `${path}.end`)
      validateDocxPath(start.path, `${path}.start.path`)
      validateDocxPath(end.path, `${path}.end.path`)
      integer(start.offset, `${path}.start.offset`)
      integer(end.offset, `${path}.end.offset`)
      break
    }
    case "image":
      validatePart(input.part, `${path}.part`)
      validateDocxPath(input.paragraphPath, `${path}.paragraphPath`)
      integer(input.childIndex, `${path}.childIndex`)
      optionalString(input.relationId, `${path}.relationId`, 256)
      break
    case "comment":
      validatePart(input.part, `${path}.part`)
      stringValue(input.commentId, `${path}.commentId`, { max: 256 })
      if (input.anchorPath !== undefined) validateDocxPath(input.anchorPath, `${path}.anchorPath`)
      break
    case "behavior":
      validateDocxBehavior(input, path)
      break
    default:
      fail("INVALID_VALUE", `${path}.kind`, "contains an unsupported DOCX locator")
  }
  return input as unknown as DocxObjectLocator
}

function validateSheet(value: unknown, path: string): XlsxSheetLocator {
  const input = record(value, path)
  optionalString(input.sheetId, `${path}.sheetId`, 512)
  stringValue(input.name, `${path}.name`, { max: 512 })
  integer(input.index, `${path}.index`)
  return input as unknown as XlsxSheetLocator
}

function validateA1(value: unknown, path: string): string {
  return stringValue(value, path, { max: 128, pattern: a1Pattern })
}

function isSingleA1(value: string): boolean {
  const [start, end] = value.toUpperCase().replace(/\$/g, "").split(":")
  return end === undefined || end === start
}

function validateXlsxBehavior(value: RecordValue, path: string): XlsxBehaviorLocator {
  stringValue(value.instanceId, `${path}.instanceId`, { max: 512 })
  const scope = record(value.scope, `${path}.scope`)
  if (["formula", "conditional-format", "data-validation"].includes(String(value.behavior))) {
    if (scope.kind !== "sheet") fail("INVALID_VALUE", `${path}.scope.kind`, "range behavior requires a sheet scope")
    validateSheet(scope.sheet, `${path}.scope.sheet`)
    validateA1(scope.appliesToA1, `${path}.scope.appliesToA1`)
  } else if (value.behavior === "named-range" || value.behavior === "protection") {
    if (scope.kind === "sheet") validateSheet(scope.sheet, `${path}.scope.sheet`)
    else if (scope.kind !== "workbook") fail("INVALID_VALUE", `${path}.scope.kind`, "contains an unsupported XLSX behavior scope")
  } else {
    fail("INVALID_VALUE", `${path}.behavior`, "contains an unsupported XLSX behavior")
  }
  return value as unknown as XlsxBehaviorLocator
}

function validateXlsxLocator(value: unknown, path: string): XlsxObjectLocator {
  const input = record(value, path)
  switch (input.kind) {
    case "worksheet":
      validateSheet(input.sheet, `${path}.sheet`)
      break
    case "range":
      validateSheet(input.sheet, `${path}.sheet`)
      validateA1(input.a1, `${path}.a1`)
      break
    case "row":
    case "column": {
      validateSheet(input.sheet, `${path}.sheet`)
      const start = integer(input.start, `${path}.start`)
      const end = integer(input.end, `${path}.end`)
      if (start > end) fail("INVALID_VALUE", path, "range start must not follow its end")
      break
    }
    case "table":
      validateSheet(input.sheet, `${path}.sheet`)
      stringValue(input.tableName, `${path}.tableName`, { max: 512 })
      optionalString(input.id, `${path}.id`, 512)
      break
    case "drawing":
      validateSheet(input.sheet, `${path}.sheet`)
      enumValue(input.objectType, new Set(["chart", "image", "shape", "form-control"]), `${path}.objectType`)
      stringValue(input.id, `${path}.id`, { max: 512 })
      optionalString(input.name, `${path}.name`, 512)
      break
    case "chart-element":
      validateSheet(input.sheet, `${path}.sheet`)
      stringValue(input.chartId, `${path}.chartId`, { max: 512 })
      if (input.element === "series") integer(input.seriesIndex, `${path}.seriesIndex`)
      else if (input.element === "point") {
        integer(input.seriesIndex, `${path}.seriesIndex`)
        integer(input.pointIndex, `${path}.pointIndex`)
      } else if (input.element === "legend-entry") {
        integer(input.legendIndex, `${path}.legendIndex`)
        if (input.seriesIndex !== undefined) integer(input.seriesIndex, `${path}.seriesIndex`)
      } else fail("INVALID_VALUE", `${path}.element`, "contains an unsupported chart element")
      break
    case "comment":
      validateSheet(input.sheet, `${path}.sheet`)
      validateA1(input.cellA1, `${path}.cellA1`)
      if (!isSingleA1(input.cellA1 as string)) fail("INVALID_VALUE", `${path}.cellA1`, "comment anchor must be one cell")
      optionalString(input.nativeId, `${path}.nativeId`, 512)
      break
    case "behavior":
      validateXlsxBehavior(input, path)
      break
    default:
      fail("INVALID_VALUE", `${path}.kind`, "contains an unsupported XLSX locator")
  }
  return input as unknown as XlsxObjectLocator
}

function validatePptxSlide(value: unknown, path: string): PptxSlideLocator {
  const input = record(value, path)
  optionalString(input.slideId, `${path}.slideId`, 512)
  integer(input.index, `${path}.index`)
  return input as unknown as PptxSlideLocator
}

function validatePptxObject(value: unknown, path: string): PptxObjectKeyLocator {
  const input = record(value, path)
  stringValue(input.objectKey, `${path}.objectKey`, { max: 1_024 })
  optionalString(input.shapeId, `${path}.shapeId`, 512)
  enumValue(input.source, new Set(["slide", "layout", "master"]), `${path}.source`)
  array(input.groupPath, `${path}.groupPath`).forEach((item, index) => stringValue(item, `${path}.groupPath[${index}]`, { max: 512 }))
  return input as unknown as PptxObjectKeyLocator
}

function validatePptxLocator(value: unknown, path: string): PptxObjectLocator {
  const input = record(value, path)
  switch (input.kind) {
    case "slide":
      validatePptxSlide(input.slide, `${path}.slide`)
      break
    case "object":
      validatePptxSlide(input.slide, `${path}.slide`)
      validatePptxObject(input.object, `${path}.object`)
      break
    case "text-range": {
      validatePptxSlide(input.slide, `${path}.slide`)
      validatePptxObject(input.object, `${path}.object`)
      const start = record(input.start, `${path}.start`)
      const end = record(input.end, `${path}.end`)
      integer(start.paragraphIndex, `${path}.start.paragraphIndex`)
      integer(start.offset, `${path}.start.offset`)
      integer(end.paragraphIndex, `${path}.end.paragraphIndex`)
      integer(end.offset, `${path}.end.offset`)
      break
    }
    case "sub-object": {
      validatePptxSlide(input.slide, `${path}.slide`)
      validatePptxObject(input.object, `${path}.object`)
      const segments = array(input.path, `${path}.path`)
      if (segments.length === 0) fail("INVALID_VALUE", `${path}.path`, "must not be empty")
      segments.forEach((value, index) => {
        const segment = record(value, `${path}.path[${index}]`)
        if (segment.kind === "table-cell") {
          integer(segment.rowIndex, `${path}.path[${index}].rowIndex`)
          integer(segment.cellIndex, `${path}.path[${index}].cellIndex`)
        } else if (segment.kind === "chart-series") {
          integer(segment.seriesIndex, `${path}.path[${index}].seriesIndex`)
        } else if (segment.kind === "chart-point") {
          integer(segment.seriesIndex, `${path}.path[${index}].seriesIndex`)
          integer(segment.pointIndex, `${path}.path[${index}].pointIndex`)
        } else if (segment.kind === "chart-legend-entry") {
          integer(segment.legendIndex, `${path}.path[${index}].legendIndex`)
          if (segment.seriesIndex !== undefined) integer(segment.seriesIndex, `${path}.path[${index}].seriesIndex`)
        } else fail("INVALID_VALUE", `${path}.path[${index}].kind`, "contains an unsupported PPTX sub-object")
      })
      break
    }
    case "behavior":
      validatePptxSlide(input.slide, `${path}.slide`)
      enumValue(input.behavior, new Set(["animation", "transition", "action", "master", "note"]), `${path}.behavior`)
      stringValue(input.nativeId, `${path}.nativeId`, { max: 512 })
      optionalString(input.ownerObjectKey, `${path}.ownerObjectKey`, 1_024)
      break
    default:
      fail("INVALID_VALUE", `${path}.kind`, "contains an unsupported PPTX locator")
  }
  return input as unknown as PptxObjectLocator
}

function validatePdfLocator(value: unknown, path: string): PdfObjectLocator {
  const input = record(value, path)
  switch (input.kind) {
    case "page":
      integer(input.pageIndex, `${path}.pageIndex`)
      break
    case "text-range":
      integer(input.pageIndex, `${path}.pageIndex`)
      integer(input.charIndex, `${path}.charIndex`)
      if (integer(input.charCount, `${path}.charCount`) === 0) fail("INVALID_VALUE", `${path}.charCount`, "must be greater than zero")
      break
    case "native-object":
      integer(input.pageIndex, `${path}.pageIndex`)
      enumValue(input.objectType, new Set(["link", "form", "annotation"]), `${path}.objectType`)
      stringValue(input.nativeId, `${path}.nativeId`, { max: 512 })
      break
    case "visual-object":
      integer(input.pageIndex, `${path}.pageIndex`)
      stringValue(input.providerId, `${path}.providerId`, { max: 512 })
      stringValue(input.objectId, `${path}.objectId`, { max: 512 })
      validateRect(input.region, `${path}.region`)
      break
    case "behavior":
      if (input.pageIndex !== undefined) integer(input.pageIndex, `${path}.pageIndex`)
      enumValue(input.behavior, new Set([
        "link-action", "form-rule", "annotation-state", "attachment", "layer", "signature", "permission",
      ]), `${path}.behavior`)
      stringValue(input.nativeId, `${path}.nativeId`, { max: 512 })
      break
    default:
      fail("INVALID_VALUE", `${path}.kind`, "contains an unsupported PDF locator")
  }
  return input as unknown as PdfObjectLocator
}

function validateDocument(value: unknown, path: string): OfficeDocumentRevision {
  const input = record(value, path)
  enumValue(input.format, formats, `${path}.format`)
  stringValue(input.documentId, `${path}.documentId`, { max: 512 })
  stringValue(input.revision, `${path}.revision`, { max: 512 })
  if (input.contentDigest !== undefined) stringValue(input.contentDigest, `${path}.contentDigest`, { pattern: sha256Pattern })
  return input as unknown as OfficeDocumentRevision
}

function validateReliabilityDimension(value: unknown, path: string): ReliabilityDimension {
  const input = record(value, path)
  enumValue(input.level, reliabilityLevels, `${path}.level`)
  if (input.score !== undefined) unit(input.score, `${path}.score`)
  array(input.reasonCodes, `${path}.reasonCodes`).forEach((code, index) => {
    stringValue(code, `${path}.reasonCodes[${index}]`, { max: 128, pattern: reasonCodePattern })
  })
  return input as unknown as ReliabilityDimension
}

function validateReliability(value: unknown, path: string): OfficeObjectReliability {
  const input = record(value, path)
  validateReliabilityDimension(input.semantic, `${path}.semantic`)
  validateReliabilityDimension(input.boundary, `${path}.boundary`)
  validateReliabilityDimension(input.hierarchy, `${path}.hierarchy`)
  validateReliabilityDimension(input.relocation, `${path}.relocation`)
  return input as unknown as OfficeObjectReliability
}

function validateFingerprint(value: unknown, path: string): void {
  const input = record(value, path)
  for (const key of ["exactText", "prefixText", "suffixText", "objectName"] as const) {
    optionalString(input[key], `${path}.${key}`, fingerprintTextLimit)
  }
  optionalString(input.contentHash, `${path}.contentHash`, 512)
  if (input.ancestorKeys !== undefined) {
    const keys = array(input.ancestorKeys, `${path}.ancestorKeys`)
    if (keys.length > 32) fail("INVALID_VALUE", `${path}.ancestorKeys`, "must contain at most 32 entries")
    keys.forEach((key, index) => stringValue(key, `${path}.ancestorKeys[${index}]`, { max: 256 }))
  }
}

function behaviorKind(format: OfficeFormat, locator: RecordValue): OfficeObjectKind | undefined {
  if (format === "docx") return ({
    style: "style",
    field: "field",
    bookmark: "bookmark",
    section: "section",
    "tracked-change-state": "tracked-change",
    "comment-state": "comment",
    "cross-reference": "cross-reference",
  } as Record<string, OfficeObjectKind>)[String(locator.behavior)]
  if (format === "xlsx") return ({
    formula: "formula",
    "conditional-format": "conditional-format",
    "data-validation": "data-validation",
    "named-range": "named-range",
    protection: "protection",
  } as Record<string, OfficeObjectKind>)[String(locator.behavior)]
  if (format === "pptx") return locator.behavior as OfficeObjectKind
  return ({
    "link-action": "action",
    "form-rule": "form-rule",
    "annotation-state": "annotation",
    attachment: "attachment",
    layer: "layer",
    signature: "signature",
    permission: "permission",
  } as Record<string, OfficeObjectKind>)[String(locator.behavior)]
}

function validateKindLocator(
  format: OfficeFormat,
  kind: OfficeObjectKind,
  source: RecognitionSource,
  locator: RecordValue,
  path: string,
): void {
  if (locator.type === "manual-region") {
    if (kind !== "region" || source !== "manual") {
      fail("KIND_LOCATOR_MISMATCH", path, "manual regions require kind=region and source=manual")
    }
    return
  }
  if (source === "manual") fail("KIND_LOCATOR_MISMATCH", path, "manual source requires a manual-region locator")
  const value = record(locator.value, `${path}.value`)
  let allowed: readonly OfficeObjectKind[] = []
  if (format === "docx") {
    if (value.kind === "page") allowed = ["page"]
    else if (value.kind === "text-range") allowed = ["text-range"]
    else if (value.kind === "image") allowed = ["image"]
    else if (value.kind === "comment") allowed = ["comment"]
    else if (value.kind === "structure") allowed = ["section", "paragraph", "heading", "list-item", "table", "table-row", "table-cell"]
    else if (value.kind === "behavior") allowed = [behaviorKind(format, value)!]
  } else if (format === "xlsx") {
    if (value.kind === "worksheet") allowed = ["worksheet"]
    else if (value.kind === "range") allowed = [isSingleA1(String(value.a1)) ? "cell" : "cell-range"]
    else if (value.kind === "row" || value.kind === "column" || value.kind === "table" || value.kind === "comment") allowed = [value.kind]
    else if (value.kind === "drawing") allowed = [value.objectType as OfficeObjectKind]
    else if (value.kind === "chart-element") allowed = [{ series: "chart-series", point: "chart-point", "legend-entry": "chart-legend-entry" }[String(value.element)] as OfficeObjectKind]
    else if (value.kind === "behavior") allowed = [behaviorKind(format, value)!]
  } else if (format === "pptx") {
    if (value.kind === "slide") allowed = ["slide"]
    else if (value.kind === "text-range") allowed = ["text-range"]
    else if (value.kind === "object") allowed = ["text-box", "shape", "image", "table", "chart", "group"]
    else if (value.kind === "sub-object") {
      const segments = value.path as readonly RecordValue[]
      allowed = [segments[segments.length - 1]?.kind as OfficeObjectKind]
    } else if (value.kind === "behavior") allowed = [behaviorKind(format, value)!]
  } else {
    if (value.kind === "page") allowed = ["page"]
    else if (value.kind === "text-range") allowed = ["text-range"]
    else if (value.kind === "native-object") allowed = [{ link: "link", form: "form-control", annotation: "annotation" }[String(value.objectType)] as OfficeObjectKind]
    else if (value.kind === "visual-object") {
      allowed = ["text-block", "image", "table", "table-row", "table-cell"]
      if (source !== "visual") fail("KIND_LOCATOR_MISMATCH", path, "PDF visual objects require source=visual")
    } else if (value.kind === "behavior") allowed = [behaviorKind(format, value)!]
  }
  if (!allowed.includes(kind)) {
    fail("KIND_LOCATOR_MISMATCH", path, `kind=${kind} is incompatible with the selected ${format} locator`)
  }
}

function validateLocator(value: unknown, document: OfficeDocumentRevision, kind: OfficeObjectKind, source: RecognitionSource, path: string): void {
  const input = record(value, path)
  if (input.format !== document.format) fail("FORMAT_MISMATCH", `${path}.format`, "must match document.format")
  if (input.type === "manual-region") {
    const region = validateManualRegion(input.value, `${path}.value`)
    const expectedSpace = { docx: "page", xlsx: "sheet", pptx: "slide", pdf: "page" }[document.format]
    if (region.space !== expectedSpace) fail("FORMAT_MISMATCH", `${path}.value.space`, `must be ${expectedSpace} for ${document.format}`)
  } else if (input.type === "format") {
    if (document.format === "docx") validateDocxLocator(input.value, `${path}.value`)
    else if (document.format === "xlsx") validateXlsxLocator(input.value, `${path}.value`)
    else if (document.format === "pptx") validatePptxLocator(input.value, `${path}.value`)
    else validatePdfLocator(input.value, `${path}.value`)
  } else {
    fail("INVALID_VALUE", `${path}.type`, "contains an unsupported locator type")
  }
  validateKindLocator(document.format, kind, source, input, path)
}

function validateReferenceInternal(value: unknown, path: string, assertJson = true): OfficeObjectReference {
  if (assertJson) assertJsonTree(value, path)
  const input = record(value, path)
  if (input.schemaVersion !== 1) fail("INVALID_SCHEMA_VERSION", `${path}.schemaVersion`, "only schemaVersion=1 is supported")
  stringValue(input.referenceId, `${path}.referenceId`, { max: 256 })
  const document = validateDocument(input.document, `${path}.document`)
  const kind = enumValue(input.kind, objectKinds, `${path}.kind`)
  const source = enumValue(input.source, recognitionSources, `${path}.source`)
  validateLocator(input.locator, document, kind, source, `${path}.locator`)
  if (input.fingerprint !== undefined) validateFingerprint(input.fingerprint, `${path}.fingerprint`)
  if (input.fallbackRegion !== undefined) {
    const region = validateManualRegion(input.fallbackRegion, `${path}.fallbackRegion`)
    const expectedSpace = { docx: "page", xlsx: "sheet", pptx: "slide", pdf: "page" }[document.format]
    if (region.space !== expectedSpace) fail("FORMAT_MISMATCH", `${path}.fallbackRegion.space`, `must be ${expectedSpace} for ${document.format}`)
  }
  validateReliability(input.reliability, `${path}.reliability`)
  return input as unknown as OfficeObjectReference
}

function validateReferenceSnapshot(value: unknown, path: string): OfficeReferenceSnapshot {
  const snapshot = record(value, path)
  onlyKeys(snapshot, ["label", "path", "content"], path)
  stringValue(snapshot.label, `${path}.label`, { max: 512 })
  array(snapshot.path, `${path}.path`).forEach((entry, index) => {
    const item = record(entry, `${path}.path[${index}]`)
    onlyKeys(item, ["kind", "label"], `${path}.path[${index}]`)
    enumValue(item.kind, objectKinds, `${path}.path[${index}].kind`)
    stringValue(item.label, `${path}.path[${index}].label`, { max: 512 })
  })
  if (snapshot.content !== undefined) {
    const content = record(snapshot.content, `${path}.content`)
    onlyKeys(content, ["text", "value", "formula", "truncated"], `${path}.content`)
    optionalString(content.text, `${path}.content.text`, 1_000_000)
    optionalString(content.value, `${path}.content.value`, 100_000)
    optionalString(content.formula, `${path}.content.formula`, 100_000)
    if (content.truncated !== undefined) booleanValue(content.truncated, `${path}.content.truncated`)
  }
  return snapshot as unknown as OfficeReferenceSnapshot
}

export function parseOfficeObjectReference(value: unknown): OfficeObjectReference {
  return validateReferenceInternal(value, "$reference")
}

export function safeParseOfficeObjectReference(value: unknown): OfficeSafeParseResult<OfficeObjectReference> {
  try {
    return { success: true, data: parseOfficeObjectReference(value) }
  } catch (error) {
    return { success: false, error: normalizeValidationError(error, "$reference") }
  }
}

export function parseOfficeReferenceConfirmEvent(value: unknown): OfficeReferenceConfirmEvent {
  assertJsonTree(value)
  const input = record(value, "$confirm")
  onlyKeys(input, ["reference", "snapshot", "trigger", "additiveRequested"], "$confirm")
  validateReferenceInternal(input.reference, "$confirm.reference", false)
  if (input.snapshot !== undefined) validateReferenceSnapshot(input.snapshot, "$confirm.snapshot")
  enumValue(input.trigger, selectionTriggers, "$confirm.trigger")
  booleanValue(input.additiveRequested, "$confirm.additiveRequested")
  return input as unknown as OfficeReferenceConfirmEvent
}

export function safeParseOfficeReferenceConfirmEvent(value: unknown): OfficeSafeParseResult<OfficeReferenceConfirmEvent> {
  try {
    return { success: true, data: parseOfficeReferenceConfirmEvent(value) }
  } catch (error) {
    return { success: false, error: normalizeValidationError(error, "$confirm") }
  }
}

export function assertOfficeJsonValue(value: unknown): asserts value is import("./types").JsonValue {
  assertJsonTree(value)
}

function normalizeValidationError(error: unknown, path: string): OfficeInteractionValidationError {
  if (error instanceof OfficeInteractionValidationError) return error
  return new OfficeInteractionValidationError({
    code: "INVALID_VALUE",
    path,
    message: error instanceof Error ? error.message : String(error),
  })
}

export function sameOfficeDocumentRevision(a: OfficeDocumentRevision, b: OfficeDocumentRevision): boolean {
  return a.format === b.format && a.documentId === b.documentId && a.revision === b.revision
}
