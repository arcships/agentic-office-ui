import type { PdfSource, PdfUrlPolicy } from "./pdf-url-policy"
import type { PdfRenderRuntime } from "./pdf/pdf-render-runtime"

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
