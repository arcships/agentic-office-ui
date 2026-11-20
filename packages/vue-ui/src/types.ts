import type { Component } from "vue"
import type {
  NormalizedRect,
  OfficeLocale,
  OfficeMessagesOverride,
  OfficeObjectKind,
  ReliabilityLevel,
} from "@arcships/office-interaction"
import type {
  OfficeRegionAnnotationDocument,
  OfficeRegionAnnotationKind,
} from "./region-annotations"

export interface SignaturePadProps {
  width?: number
  height?: number
  penColor?: string
  backgroundColor?: string
  className?: string
  /** Draw on a canvas or type a signature. Defaults to `draw`. */
  mode?: "draw" | "typed"
  /** Controlled value: data URL for `draw`, text for `typed`. */
  modelValue?: string | null
  /** Initial data URL for draw mode. */
  defaultValue?: string | null
  /** Typed-mode placeholder copy. Defaults to the localized message. */
  placeholder?: string
  /** Render the clear / cancel / confirm row. Defaults to true. */
  showControls?: boolean
  showClear?: boolean
  showCancel?: boolean
  showConfirm?: boolean
  /** Force the confirm action off (e.g. while a host dialog validates). */
  confirmDisabled?: boolean
  disabled?: boolean
  /** Draw the dashed signature guide around the canvas. */
  guide?: boolean
  loading?: boolean
  error?: string
  ariaLabel?: string
}

/** Built-in upload glyph or a host-supplied icon component. */
export type FileUploadAcceptedIcon = "image" | "pdf" | "sheet" | Component

/**
 * One accepted file-type glyph in the upload icon cluster.
 *
 * `label` is retained for API parity with the Extend prop shape; the glyphs
 * themselves are decorative. Defaults to Image / PDF / Sheet.
 */
export interface FileUploadAcceptedType {
  label?: string
  icon?: FileUploadAcceptedIcon
}

export interface FileUploadProps {
  accept?: string
  multiple?: boolean
  maxSize?: number
  disabled?: boolean
  className?: string
  /** Heading copy. Defaults to the localized upload prompt. */
  title?: string
  /** Supporting copy under the heading; hidden when omitted. */
  description?: string
  /** Drop-target chip copy when idle. */
  browseLabel?: string
  /** Drop-target chip copy while dragging. */
  draggingLabel?: string
  /** Glyphs shown in the icon cluster. Defaults to Image / PDF / Sheet. */
  acceptedFileTypes?: FileUploadAcceptedType[]
  /** Render the selected-file list under the drop zone. Defaults to false. */
  showFileList?: boolean
  /** Paint the optional animated border beam while dragging. */
  showBorderBeam?: boolean
  /** Border-beam palette hint, kept for parity with the Extend API. */
  borderBeamTheme?: "light" | "dark"
}

export type FileUploadRejectionCode = "FILE_TYPE_NOT_ACCEPTED" | "FILE_TOO_LARGE"

/** Machine-readable reason for a file omitted from `files-accepted`. */
export interface FileUploadRejection {
  code: FileUploadRejectionCode
  file: File
  message: string
}

export interface FileThumbnailFile {
  name: string
  type: string
  url?: string
  size?: number
}

export interface FileThumbnailProps {
  file: FileThumbnailFile
  /**
   * Preview width. The named sizes keep the pre-0.8 API compatible; a number
   * is interpreted as pixels and a string is passed through as a CSS length.
   */
  size?: "sm" | "md" | "lg" | number | string
  className?: string
  /** Explicit preview image URL. Falls back to `file.url` for image types. */
  previewImageUrl?: string | null
  /** Preview aspect ratio, e.g. `16 / 9`. Defaults to a square tile. */
  previewAspectRatio?: number
  /** Extra class merged on the preview (aspect-ratio) layer. */
  previewClassName?: string
  /** Show the loading overlay, independent of image loading. */
  isLoading?: boolean
  /** Force the muted error/empty fallback. */
  hasError?: boolean
  /** Render the file-type glyph in the fallback layer. */
  showTypeGlyph?: boolean
  /** Render the file name and type under the preview (inline row layout). */
  showMeta?: boolean
  /** Stretch the root to the parent box (Finder/thumbnail-strip layouts). */
  fill?: boolean
  /** Let the preview fill the parent height instead of a square tile. */
  previewFill?: boolean
  /** Image alt text. Defaults to the file name. */
  alt?: string
}

export interface BoundingBoxField {
  id: string
  label: string
  page: number
  rect: [number, number, number, number]
  value?: string
  confidence?: number
}

/** JSON value accepted by the Extend review-field contract. */
export type OfficeJsonPrimitive = string | number | boolean | null
export type OfficeJsonValue =
  | OfficeJsonPrimitive
  | OfficeJsonObject
  | OfficeJsonArray
export interface OfficeJsonObject {
  [key: string]: OfficeJsonValue
}
export type OfficeJsonArray = OfficeJsonValue[]

export type BoundingBoxReviewSchemaType =
  | "string"
  | "number"
  | "integer"
  | "boolean"
  | "object"
  | "array"

export interface BoundingBoxReviewFieldSchema {
  type: BoundingBoxReviewSchemaType
  title?: string
  description?: string
  enum?: Array<string | number>
  properties?: Record<string, BoundingBoxReviewFieldSchema>
  items?: BoundingBoxReviewFieldSchema
}

export interface BoundingBoxHighlightArea {
  left: number
  top: number
  width: number
  height: number
}

export interface BoundingBoxReviewLocation {
  page: number
  area: BoundingBoxHighlightArea
}

/** Extend `HumanReviewPanel` field shape (page highlight + actual/expected). */
export interface BoundingBoxReviewField {
  key: string
  schema: BoundingBoxReviewFieldSchema
  actual: OfficeJsonValue
  expected: OfficeJsonValue
  location?: BoundingBoxReviewLocation
  metadataPath?: string
}

export interface BoundingBoxCitationsProps {
  /** Page image or PDF URL. Required for the legacy page-card mode. */
  file?: string
  /** Legacy per-page field cards. */
  fields?: BoundingBoxField[]
  /** Extend review fields; enables the review-panel mode when present. */
  reviewFields?: BoundingBoxReviewField[]
  /** Page image used behind the highlight overlay in review mode. */
  pageImage?: string
  /** Intrinsic page size used to position `location` percentages. */
  pageWidth?: number
  pageHeight?: number
  /** Controlled active review field. */
  activeFieldKey?: string
  /** Show the editable Expected column. Defaults to true. */
  showExpected?: boolean
  /** Controlled expected-draft values keyed by field key. */
  expectedValues?: Record<string, OfficeJsonValue>
  /** Render the review values read-only. */
  readOnly?: boolean
  /** `split` renders the page + panel; `panel` renders the side list alone. */
  view?: "split" | "panel"
  loading?: boolean
  error?: string
  emptyMessage?: string
  className?: string
}

export type LayoutBlockType =
  | "heading"
  | "paragraph"
  | "list"
  | "table"
  | "figure"
  | "header"
  | "footer"
  | "page_number"

export interface LayoutBlock {
  id: string
  bbox: [number, number, number, number]
  /** Legacy free-form kind. Mapped to `type` when `type` is omitted. */
  kind: string
  /** Extend layout-block type. */
  type?: LayoutBlockType
  /** Optional badge label override. */
  label?: string
  text?: string
  /** 1-indexed page for the panel page chip. */
  page?: number
  confidence?: number
  parentId?: string
}

export interface ParsedOcrOutput {
  width: number
  height: number
  blocks: LayoutBlock[]
}

export interface LayoutBlocksProps {
  /** Page image or PDF URL. */
  file?: string
  /** Legacy parsed OCR payload. */
  output?: ParsedOcrOutput
  /** Extend block list; wins over `output.blocks`. */
  blocks?: LayoutBlock[]
  /** Image rendered behind the block overlay in split mode. */
  pageImage?: string
  /** Intrinsic page width used for overlay percentages when `blocks` is used. */
  pageWidth?: number
  /** Intrinsic page height used for overlay percentages when `blocks` is used. */
  pageHeight?: number
  activeBlockId?: string
  /** `split` renders page + panel; `panel` renders the block list alone. */
  view?: "split" | "panel"
  heading?: string
  emptyMessage?: string
  /** Render `text` as sanitized inline markdown. */
  renderMarkdown?: boolean
  loading?: boolean
  error?: string
  ariaLabel?: string
  className?: string
}

export interface ESignatureFieldBox {
  x: number
  y: number
  width: number
  height: number
}

export interface ESignatureField {
  id: string
  label: string
  page: number
  bbox: ESignatureFieldBox
  imageDataUrl?: string
}

export interface ESignatureFieldsProps {
  fields: ESignatureField[]
  /** Enable the export action (a signed field and a source file exist). */
  canExport?: boolean
  exporting?: boolean
  /** Render the signed image above the field actions. */
  showPreview?: boolean
  title?: string
  description?: string
  loading?: boolean
  error?: string
  className?: string
}

export type DocumentSplitPageId = `page-${number}`

export interface DocumentSplit {
  id: string
  title: string
  pages: DocumentSplitPageId[]
}
export interface DocumentSplitsProps {
  splits: DocumentSplit[]
  /** Page thumbnails keyed by page id (`page-1`). */
  thumbnailImages?: Record<string, string>
  /** Draw the left frame divider used by the Extend PDF shell. */
  withFrameDivider?: boolean
  /** Hide the add-split action. */
  hideAddSplit?: boolean
  /** Disable remove for every split (in addition to the single-split guard). */
  disableRemove?: boolean
  loading?: boolean
  error?: string
  emptyMessage?: string
  className?: string
}

export interface SpinnerProps {
  size?: "sm" | "md" | "lg"
  className?: string
}

export interface TooltipProps {
  content?: string
  side?: "top" | "bottom" | "left" | "right"
  align?: "start" | "center" | "end"
  delayMs?: number
  className?: string
}

export type OfficeObjectOutlineState = "available" | "active" | "selected" | "invalid"

/** One normalized visual fragment rendered inside a single Surface container. */
export interface OfficeObjectOutline {
  id: string
  referenceId?: string
  label: string
  kind: OfficeObjectKind
  rect: NormalizedRect
  state?: OfficeObjectOutlineState
  reliability?: ReliabilityLevel
  disabled?: boolean
}

export interface OfficeObjectOutlineLayerProps {
  items: readonly OfficeObjectOutline[]
  activeId?: string
  interactive?: boolean
  ariaLabel?: string
  className?: string
  locale?: OfficeLocale
  messages?: OfficeMessagesOverride
}

export interface OfficeOutlineConfirmOptions {
  additiveRequested: boolean
  penetrateRequested: boolean
}

export interface OfficeRegionSelectorProps {
  modelValue?: NormalizedRect | null
  disabled?: boolean
  minSize?: number
  keyboardStep?: number
  ariaLabel?: string
  className?: string
  locale?: OfficeLocale
  messages?: OfficeMessagesOverride
}

export interface OfficeRegionAnnotatorProps {
  /** Controlled annotation document. */
  modelValue?: OfficeRegionAnnotationDocument | null
  /** Captured bitmap URL, or a Blob that the component owns and revokes. */
  image?: string | Blob | null
  /** Intrinsic bitmap width; falls back to the decoded image size. */
  width?: number
  /** Intrinsic bitmap height; falls back to the decoded image size. */
  height?: number
  /** Active drawing tool. */
  mode?: OfficeRegionAnnotationKind
  color?: string
  /** Normalized stroke width (fraction of the shorter image edge). */
  strokeWidth?: number
  /** Normalized font size (fraction of the shorter image edge). */
  fontSize?: number
  disabled?: boolean
  readOnly?: boolean
  maxAnnotations?: number
  alt?: string
  ariaLabel?: string
  className?: string
  locale?: OfficeLocale
  messages?: OfficeMessagesOverride
}
