export type OfficeFormat = "docx" | "xlsx" | "pptx" | "pdf"

export type OfficeObjectKind =
  | "document" | "page" | "section" | "paragraph" | "heading" | "list-item"
  | "text-range" | "text-block" | "table" | "table-row" | "table-cell" | "image"
  | "workbook" | "worksheet" | "cell" | "cell-range" | "row" | "column"
  | "chart" | "chart-series" | "chart-point" | "chart-legend-entry" | "shape" | "text-box" | "group"
  | "slide" | "layout" | "master" | "comment" | "annotation" | "link" | "form-control"
  | "region" | "formula" | "style" | "conditional-format" | "data-validation"
  | "animation" | "transition" | "action" | "note" | "bookmark" | "field"
  | "named-range" | "protection" | "tracked-change" | "cross-reference" | "form-rule"
  | "attachment" | "layer" | "signature" | "permission" | "unknown"

export interface OfficeDocumentRevision {
  format: OfficeFormat
  documentId: string
  revision: string
  contentDigest?: `sha256:${string}`
}

export type ReliabilityLevel = "exact" | "likely" | "uncertain" | "unknown"

export interface ReliabilityDimension {
  level: ReliabilityLevel
  score?: number
  reasonCodes: readonly string[]
}

export interface OfficeObjectReliability {
  semantic: ReliabilityDimension
  boundary: ReliabilityDimension
  hierarchy: ReliabilityDimension
  relocation: ReliabilityDimension
}

export type RecognitionSource = "native" | "structural" | "visual" | "manual"

export interface NormalizedPoint {
  x: number
  y: number
}

export interface NormalizedRect extends NormalizedPoint {
  width: number
  height: number
}

export interface SheetRegionPoint {
  row: number
  col: number
  xOffset: number
  yOffset: number
}

export type ManualRegionLocator =
  | { space: "page"; pageIndex: number; rect: NormalizedRect }
  | { space: "slide"; slideIndex: number; rect: NormalizedRect }
  | { space: "sheet"; sheetId: string; start: SheetRegionPoint; end: SheetRegionPoint }

export type RegionAnnotation =
  | { id: string; kind: "rectangle" | "ellipse"; rect: NormalizedRect; color?: string }
  | { id: string; kind: "arrow"; start: NormalizedPoint; end: NormalizedPoint; color?: string }
  | { id: string; kind: "freehand"; points: readonly NormalizedPoint[]; color?: string }
  | { id: string; kind: "text" | "pin"; point: NormalizedPoint; text: string; color?: string }

export type DocxPartLocator =
  | { kind: "body" }
  | { kind: "header" | "footer"; partName: string }
  | { kind: "footnote" | "endnote"; noteId: string }

export type DocxPathSegment = {
  kind: "node" | "table" | "row" | "cell" | "paragraph" | "run"
  index: number
  nativeId?: string
}

export type DocxBehaviorOwner =
  | { scope: "document" }
  | { scope: "part"; part: DocxPartLocator; path: readonly DocxPathSegment[] }

export interface DocxBehaviorLocator {
  kind: "behavior"
  behavior: "style" | "field" | "bookmark" | "section" | "tracked-change-state" | "comment-state" | "cross-reference"
  owner: DocxBehaviorOwner
  instanceId: string
}

export type DocxObjectLocator =
  | { kind: "page"; pageIndex: number }
  | { kind: "structure"; part: DocxPartLocator; path: readonly DocxPathSegment[] }
  | {
      kind: "text-range"
      part: DocxPartLocator
      start: { path: readonly DocxPathSegment[]; offset: number }
      end: { path: readonly DocxPathSegment[]; offset: number }
    }
  | {
      kind: "image"
      part: DocxPartLocator
      paragraphPath: readonly DocxPathSegment[]
      childIndex: number
      relationId?: string
    }
  | { kind: "comment"; part: DocxPartLocator; commentId: string; anchorPath?: readonly DocxPathSegment[] }
  | DocxBehaviorLocator

export interface XlsxSheetLocator {
  sheetId?: string
  name: string
  index: number
}

export type XlsxBehaviorLocator =
  | {
      kind: "behavior"
      behavior: "formula" | "conditional-format" | "data-validation"
      scope: { kind: "sheet"; sheet: XlsxSheetLocator; appliesToA1: string }
      instanceId: string
    }
  | {
      kind: "behavior"
      behavior: "named-range" | "protection"
      scope: { kind: "workbook" } | { kind: "sheet"; sheet: XlsxSheetLocator }
      instanceId: string
    }

export type XlsxObjectLocator =
  | { kind: "worksheet"; sheet: XlsxSheetLocator }
  | { kind: "range"; sheet: XlsxSheetLocator; a1: string }
  | { kind: "row" | "column"; sheet: XlsxSheetLocator; start: number; end: number }
  | { kind: "table"; sheet: XlsxSheetLocator; tableName: string; id?: string }
  | {
      kind: "drawing"
      sheet: XlsxSheetLocator
      objectType: "chart" | "image" | "shape" | "form-control"
      id: string
      name?: string
    }
  | { kind: "chart-element"; sheet: XlsxSheetLocator; chartId: string; element: "series"; seriesIndex: number }
  | {
      kind: "chart-element"
      sheet: XlsxSheetLocator
      chartId: string
      element: "point"
      seriesIndex: number
      pointIndex: number
    }
  | {
      kind: "chart-element"
      sheet: XlsxSheetLocator
      chartId: string
      element: "legend-entry"
      legendIndex: number
      seriesIndex?: number
    }
  | { kind: "comment"; sheet: XlsxSheetLocator; cellA1: string; nativeId?: string }
  | XlsxBehaviorLocator

export interface PptxSlideLocator {
  slideId?: string
  index: number
}

export interface PptxObjectKeyLocator {
  objectKey: string
  shapeId?: string
  source: "slide" | "layout" | "master"
  groupPath: readonly string[]
}

export type PptxSubObjectPathSegment =
  | { kind: "table-cell"; rowIndex: number; cellIndex: number }
  | { kind: "chart-series"; seriesIndex: number }
  | { kind: "chart-point"; seriesIndex: number; pointIndex: number }
  | { kind: "chart-legend-entry"; legendIndex: number; seriesIndex?: number }

export type PptxObjectLocator =
  | { kind: "slide"; slide: PptxSlideLocator }
  | { kind: "object"; slide: PptxSlideLocator; object: PptxObjectKeyLocator }
  | {
      kind: "text-range"
      slide: PptxSlideLocator
      object: PptxObjectKeyLocator
      start: { paragraphIndex: number; offset: number }
      end: { paragraphIndex: number; offset: number }
    }
  | { kind: "sub-object"; slide: PptxSlideLocator; object: PptxObjectKeyLocator; path: readonly PptxSubObjectPathSegment[] }
  | {
      kind: "behavior"
      slide: PptxSlideLocator
      behavior: "animation" | "transition" | "action" | "master" | "note"
      nativeId: string
      ownerObjectKey?: string
    }

export type PdfObjectLocator =
  | { kind: "page"; pageIndex: number }
  | { kind: "text-range"; pageIndex: number; charIndex: number; charCount: number }
  | { kind: "native-object"; pageIndex: number; objectType: "link" | "form" | "annotation"; nativeId: string }
  | { kind: "visual-object"; pageIndex: number; providerId: string; objectId: string; region: NormalizedRect }
  | {
      kind: "behavior"
      pageIndex?: number
      behavior: "link-action" | "form-rule" | "annotation-state" | "attachment" | "layer" | "signature" | "permission"
      nativeId: string
    }

export type OfficeObjectLocator =
  | { type: "format"; format: "docx"; value: DocxObjectLocator }
  | { type: "format"; format: "xlsx"; value: XlsxObjectLocator }
  | { type: "format"; format: "pptx"; value: PptxObjectLocator }
  | { type: "format"; format: "pdf"; value: PdfObjectLocator }
  | { type: "manual-region"; format: OfficeFormat; value: ManualRegionLocator }

export interface OfficeObjectFingerprint {
  exactText?: string
  prefixText?: string
  suffixText?: string
  objectName?: string
  contentHash?: string
  ancestorKeys?: readonly string[]
}

export type JsonValue =
  | null | boolean | number | string
  | readonly JsonValue[]
  | { readonly [key: string]: JsonValue }

export type DocxReferenceKind =
  | "page" | "section" | "paragraph" | "heading" | "list-item" | "text-range"
  | "table" | "table-row" | "table-cell" | "image" | "comment" | "region"
  | "style" | "field" | "bookmark" | "tracked-change" | "cross-reference"

export type XlsxReferenceKind =
  | "worksheet" | "cell" | "cell-range" | "row" | "column" | "table" | "chart"
  | "chart-series" | "chart-point" | "chart-legend-entry" | "image" | "shape" | "form-control" | "comment" | "region"
  | "formula" | "conditional-format" | "data-validation" | "named-range" | "protection"

export type PptxReferenceKind =
  | "slide" | "text-range" | "text-box" | "shape" | "image" | "table" | "chart" | "group"
  | "table-cell" | "chart-series" | "chart-point" | "chart-legend-entry" | "region"
  | "animation" | "transition" | "action" | "master" | "note"

export type PdfReferenceKind =
  | "page" | "text-range" | "text-block" | "image" | "table" | "table-row" | "table-cell"
  | "link" | "form-control" | "annotation" | "region" | "action" | "form-rule"
  | "attachment" | "layer" | "signature" | "permission"

export type FormatLocator<F extends OfficeFormat, L> = { type: "format"; format: F; value: L }
export type RegionLocator<F extends OfficeFormat, S extends ManualRegionLocator["space"]> = {
  type: "manual-region"
  format: F
  value: Extract<ManualRegionLocator, { space: S }>
}

export interface OfficeObjectReferenceBase<
  F extends OfficeFormat,
  K extends OfficeObjectKind,
  L,
  R extends ManualRegionLocator,
> {
  schemaVersion: 1
  referenceId: string
  document: OfficeDocumentRevision & { format: F }
  kind: K
  source: RecognitionSource
  locator: L
  fingerprint?: OfficeObjectFingerprint
  fallbackRegion?: R
  reliability: OfficeObjectReliability
}

export type OfficeObjectReference =
  | OfficeObjectReferenceBase<
      "docx",
      DocxReferenceKind,
      FormatLocator<"docx", DocxObjectLocator> | RegionLocator<"docx", "page">,
      Extract<ManualRegionLocator, { space: "page" }>
    >
  | OfficeObjectReferenceBase<
      "xlsx",
      XlsxReferenceKind,
      FormatLocator<"xlsx", XlsxObjectLocator> | RegionLocator<"xlsx", "sheet">,
      Extract<ManualRegionLocator, { space: "sheet" }>
    >
  | OfficeObjectReferenceBase<
      "pptx",
      PptxReferenceKind,
      FormatLocator<"pptx", PptxObjectLocator> | RegionLocator<"pptx", "slide">,
      Extract<ManualRegionLocator, { space: "slide" }>
    >
  | OfficeObjectReferenceBase<
      "pdf",
      PdfReferenceKind,
      FormatLocator<"pdf", PdfObjectLocator> | RegionLocator<"pdf", "page">,
      Extract<ManualRegionLocator, { space: "page" }>
    >

export type DistributiveOmit<T, K extends PropertyKey> = T extends unknown ? Omit<T, K> : never
export type OfficeObjectReferenceDraft = DistributiveOmit<OfficeObjectReference, "referenceId">

export type OfficeVisualFragment =
  | {
      container: { space: "page"; pageIndex: number } | { space: "slide"; slideIndex: number }
      rect: NormalizedRect
      zIndex?: number
    }
  | {
      container: { space: "sheet"; sheetId: string }
      region: Extract<ManualRegionLocator, { space: "sheet" }>
      zIndex?: number
    }

export interface OfficeObjectDescriptor {
  objectId: string
  draft: OfficeObjectReferenceDraft
  label: string
  path: readonly { kind: OfficeObjectKind; label: string }[]
  parentObjectId?: string
  childrenState: "none" | "available" | "lazy" | "unknown"
  content?: { text?: string; value?: string; formula?: string }
  visual?: { fragments: readonly OfficeVisualFragment[]; layoutVersion: string }
  facets?: Readonly<Record<string, JsonValue>>
}

export interface OfficeHitCandidate<RuntimeTarget = unknown> {
  candidateId: string
  draft: OfficeObjectReferenceDraft
  preview?: Pick<OfficeObjectDescriptor, "label" | "path" | "visual">
  runtimeTarget?: RuntimeTarget
  hit: "direct" | "inside" | "ancestor" | "inferred"
  depth: number
  zIndex?: number
}

export interface OfficeReferenceSnapshot {
  label: string
  path: readonly { kind: OfficeObjectKind; label: string }[]
  content?: { text?: string; value?: string; formula?: string; truncated?: boolean }
}

export type OfficeSelectionTrigger = "pointer" | "keyboard" | "touch" | "programmatic"

export interface OfficeReferenceConfirmEvent {
  reference: OfficeObjectReference
  snapshot?: OfficeReferenceSnapshot
  trigger: OfficeSelectionTrigger
  additiveRequested: boolean
}

export type ResolveReferenceResult =
  | { status: "exact"; reference: OfficeObjectReference; descriptor: OfficeObjectDescriptor }
  | { status: "relocated"; reference: OfficeObjectReference; descriptor: OfficeObjectDescriptor; reasonCodes: readonly string[] }
  | { status: "ambiguous"; candidates: readonly OfficeObjectDescriptor[]; reasonCodes: readonly string[] }
  | { status: "not-found"; reasonCode: string }
  | { status: "unsupported"; reasonCode: string }

export type OfficeSelectionMode = "content" | "object" | "region"
export type OfficeSelectionPhase = "idle" | "pointing" | "choosing" | "drawing"

export interface OfficeCandidateNavigationState {
  candidates: readonly OfficeReferenceCandidatePreview[]
  activeCandidateId?: string
}

export interface OfficeSelectionSessionState {
  mode: OfficeSelectionMode
  phase: OfficeSelectionPhase
  candidates: readonly OfficeReferenceCandidatePreview[]
  activeCandidateId?: string
  regionDraft?: ManualRegionLocator
}

export interface OfficeSelectionKeyboardInput {
  key: string
  shiftKey?: boolean
}

export type OfficeSelectionKeyboardCommand =
  | "next-candidate"
  | "previous-candidate"
  | "confirm-candidate"
  | "dismiss-candidates"
  | "cancel-selection"
  | "enter-child"
  | "return-parent"

export interface OfficeSelectionKeyboardContext {
  canEnterChild?: boolean
  canReturnParent?: boolean
}

export interface OfficeSelectionKeyboardResult {
  state: OfficeSelectionSessionState
  handled: boolean
  command?: OfficeSelectionKeyboardCommand
  activeCandidate?: OfficeReferenceCandidatePreview
}

export interface OfficeReferenceCandidatePreview {
  candidateId: string
  draft: OfficeObjectReferenceDraft
  preview?: Pick<OfficeObjectDescriptor, "label" | "path" | "visual">
  hit: OfficeHitCandidate["hit"]
  depth: number
  zIndex?: number
}

export type OfficeReferenceOperation = "hit-test" | "describe" | "resolve" | "scroll" | "capture" | "provider"
export type OfficeReferenceErrorCode =
  | "INVALID_REFERENCE" | "REVISION_CHANGED" | "HIT_TEST_FAILED" | "DESCRIBE_FAILED"
  | "RESOLVE_FAILED" | "CAPTURE_UNSUPPORTED" | "CAPTURE_LIMIT_EXCEEDED"
  | "PROVIDER_FAILED" | "ABORTED"

export interface OfficeReferenceError {
  code: OfficeReferenceErrorCode
  operation: OfficeReferenceOperation
  format: OfficeFormat
  recoverable: boolean
  referenceId?: string
  message: string
}

export interface OfficeReferenceCandidateChange {
  candidates: readonly OfficeReferenceCandidatePreview[]
  activeCandidateId?: string
}

export interface OfficeRegionDraftChange {
  phase: "start" | "change"
  region: ManualRegionLocator
}

export interface OfficeSelectionCancelEvent {
  mode: OfficeSelectionMode
  reason: "escape" | "pointer-cancel" | "programmatic"
}

export interface OfficeReferenceResolveEvent {
  referenceId: string
  result: ResolveReferenceResult
}

/** Format-neutral props implemented by every first-party Office Surface. */
export interface OfficeReferenceSurfaceProps {
  /** Stable logical file identity. A per-Surface session id is used when omitted. */
  documentId?: string
  /** Controlled interpretation of the next pointer/keyboard selection gesture. */
  selectionMode?: OfficeSelectionMode
  /** Opt in to short-lived hover candidate events. */
  emitReferenceCandidates?: boolean
}

export interface OfficeSurfaceClientPoint {
  clientX: number
  clientY: number
}

export interface OfficeReferencePreviewOptions {
  maxWidth?: number
  maxHeight?: number
  signal?: AbortSignal
}

/** Shared imperative contract; format-specific methods remain available alongside it. */
export interface OfficeReferenceSurfaceExposed {
  getDocumentRevision(): OfficeDocumentRevision
  hitTest(point: OfficeSurfaceClientPoint): readonly OfficeReferenceCandidatePreview[]
  describeReference(reference: OfficeObjectReference, signal?: AbortSignal): Promise<OfficeObjectDescriptor>
  resolveReference(reference: OfficeObjectReference): Promise<ResolveReferenceResult>
  scrollToReference(reference: OfficeObjectReference): Promise<void>
  captureReferencePreview(reference: OfficeObjectReference, options?: OfficeReferencePreviewOptions): Promise<Blob>
}
