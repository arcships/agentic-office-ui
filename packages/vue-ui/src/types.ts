import type {
  NormalizedRect,
  OfficeObjectKind,
  ReliabilityLevel,
} from "@arcships/office-interaction"

export interface SignaturePadProps {
  width?: number
  height?: number
  penColor?: string
  backgroundColor?: string
  className?: string
}

export interface FileUploadProps {
  accept?: string
  multiple?: boolean
  maxSize?: number
  disabled?: boolean
  className?: string
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
}

export interface FileThumbnailProps {
  file: FileThumbnailFile
  size?: "sm" | "md" | "lg"
  className?: string
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
}
