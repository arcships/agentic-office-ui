import type { PdfSource, PdfUrlPolicy } from "./pdf-url-policy"
import type { PdfRenderRuntime } from "./pdf/pdf-render-runtime"

export type {
  PdfDiagnostic,
  PdfLoadError,
  PdfLoadErrorCode,
  PdfLoadOptions,
  PdfSource,
  PdfUrlPolicy,
  PdfVerifiedDocument,
} from "./pdf-url-policy"

export interface PdfViewerProps {
  src?: string
  /** Preferred source input. When both source and src are set, source wins. */
  source?: PdfSource
  /** Explicit URL policy required for URL sources. Blob/File sources do not fetch. */
  urlPolicy?: PdfUrlPolicy
  /** Optional caller-owned runtime. When omitted, this viewer owns a dedicated PDF Worker. */
  runtime?: PdfRenderRuntime
  /** Overrides the bundled PDFium WASM URL for the viewer-owned runtime. */
  pdfiumWasmUrl?: string
  /** Maximum accepted PDF source size in bytes. Defaults to 50 MiB. */
  maxFileSize?: number
  fileName?: string
  defaultZoom?: number
  className?: string
  showToolbar?: boolean
  showDownload?: boolean
  showRotateControls?: boolean
}

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
