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
  /** Comma-separated accept list. Defaults to the Extend document/image list. */
  accept?: string
  /** Allow more than one file. Defaults to `true` like the Extend reference. */
  multiple?: boolean
  maxSize?: number
  disabled?: boolean
  className?: string
  /** Heading copy. Defaults to the localized upload title. */
  title?: string
  /** Supporting copy under the heading. Defaults to the localized description. */
  description?: string
  /** Drop-target chip copy when idle. */
  browseLabel?: string
  /** Drop-target chip copy while dragging. */
  draggingLabel?: string
  /** Glyphs shown in the icon cluster. Defaults to Image / PDF / Sheet. */
  acceptedFileTypes?: FileUploadAcceptedType[]
  /** Render the selected-file list under the drop zone. Defaults to `true`. */
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
  /** Render the file-type glyph in the fallback layer (Vue extension; Extend has no glyph). */
  showTypeGlyph?: boolean
  /** Render the file name and type under the preview (Vue extension, inline row layout). */
  showMeta?: boolean
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

export interface BoundingBoxCitationsProps {
  file: string
  fields: BoundingBoxField[]
  className?: string
}

export interface LayoutBlock {
  id: string
  bbox: [number, number, number, number]
  kind: string
  text?: string
  confidence?: number
  parentId?: string
}

export interface ParsedOcrOutput {
  width: number
  height: number
  blocks: LayoutBlock[]
}

export interface LayoutBlocksProps {
  file: string
  output: ParsedOcrOutput
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
