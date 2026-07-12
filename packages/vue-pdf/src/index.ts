/**
 * @arcships/vue-pdf
 *
 * Vue 3 PDF viewer component with PDFium-based rendering.
 */

// PDF viewer
export { default as PdfViewer } from "./components/PdfViewer.vue"
/**
 * Minimal embeddable PDF renderer with vertical scroll through all pages.
 * No toolbar, thumbnails panel, or search bar — host owns all controls.
 * Renders every page as a stacked image in a scrollable container.
 *
 * @example
 * ```vue
 * <PdfSurface
 *   ref="surfaceRef"
 *   :source="{ kind: 'url', url: '/doc.pdf' }"
 *   @document-load-success="pages = $event"
 * />
 * // surfaceRef.value.scrollToPage(3)
 * ```
 */
export { default as PdfSurface } from "./components/PdfSurface.vue"
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
  PdfTextRectItem,
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
