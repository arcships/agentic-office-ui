/**
 * @arcships/vue-pdf
 *
 * Vue 3 PDF viewer component with PDFium-based rendering.
 */

// PDF viewer
export { default as PdfViewer } from "./components/PdfViewer.vue"
export {
  DEFAULT_PDF_MAX_FILE_SIZE,
  PdfSourceError,
  loadVerifiedPdfSource,
  toPdfLoadError,
} from "./pdf-url-policy"
export { bundledPdfiumWasmUrl, createPdfRenderRuntime } from "./pdf/pdf-render-runtime"
export type {
  PdfPageRenderOptions,
  PdfRenderDocument,
  PdfRenderPageInfo,
  PdfRenderRect,
  PdfRenderRuntime,
  PdfRenderRuntimeConfig,
  PdfRotation,
  PdfSearchHit,
  PdfThumbnailRenderOptions,
} from "./pdf/pdf-render-runtime"
export type {
  PdfDiagnostic,
  PdfLoadError,
  PdfLoadErrorCode,
  PdfLoadOptions,
  PdfSource,
  PdfUrlPolicy,
  PdfVerifiedDocument,
} from "./pdf-url-policy"

export type { PdfViewerProps } from "./types"
