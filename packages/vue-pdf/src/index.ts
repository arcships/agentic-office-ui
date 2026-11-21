/**
 * @arcships/vue-pdf
 *
 * Vue 3 PDF viewer component with PDFium-based rendering.
 */

export type {
  OfficeDocumentRevision,
  OfficeObjectReference,
  OfficeReferenceCandidateChange,
  OfficeReferenceConfirmEvent,
  OfficeReferenceError,
  OfficeReferenceResolveEvent,
  OfficeReferenceSurfaceExposed,
  OfficeReferenceSurfaceProps,
  OfficeRegionDraftChange,
  OfficeSelectionCancelEvent,
  OfficeSelectionMode,
} from "@arcships/office-interaction"

// PDF viewer
export { default as PdfViewer } from "./components/PdfViewer.vue"
/**
 * Minimal embeddable PDF renderer with vertical scroll through all pages.
 * Toolbar-free by default — hosts own all controls. Set `showToolbar` to render
 * the shared `PdfViewerToolbar` (page, zoom, rotate, sidebar, search, download,
 * print); layout stays host-controlled through `pageLayout`, so its menu items
 * are disabled. A bookmark outline is shown when the document has one, and page
 * thumbnails use the shared viewer strip; both can be disabled through
 * `showOutline` / `showThumbnails`. Renders every page as a stacked image in a
 * scrollable container.
 *
 * @example
 * ```vue
 * <PdfSurface
 *   ref="surfaceRef"
 *   :source="{ kind: 'url', url: '/doc.pdf' }"
 *   :show-toolbar="true"
 *   @document-load-success="pages = $event"
 * />
 * // surfaceRef.value.scrollToPage(3)
 * ```
 */
export { default as PdfSurface } from "./components/PdfSurface.vue"
export {
  PDF_FILE_ACCEPT,
  PDF_FILE_EXTENSIONS,
  PDF_FILE_MIME_TYPES,
  isSupportedPdfFileName,
  pdfFileExtension,
} from "./pdf-file-types"
export {
  DEFAULT_PDF_MAX_FILE_SIZE,
  PdfSourceError,
  loadVerifiedPdfSource,
  localizePdfErrorMessage,
  toPdfLoadError,
} from "./pdf-url-policy"
export type { PdfLoadErrorMessages } from "./pdf-url-policy"
export { bundledPdfiumWasmUrl, createPdfRenderRuntime } from "./pdf/pdf-render-runtime"
export type { PdfOutlineItem } from "./pdf/pdf-outline"
export {
  DEFAULT_PDF_MAX_RENDER_PIXELS,
  MAX_PDF_RENDER_DPR,
  MAX_PDF_RENDER_SCALE,
  PdfRenderError,
  resolvePdfFormFieldValue,
  resolvePdfPageRenderParameters,
} from "./pdf/pdf-render-runtime"
export {
  createPdfBehaviorReferenceDraft,
  createPdfNativeObjectReferenceDraft,
  createPdfPageReferenceDraft,
  createPdfRegionReferenceDraft,
  createPdfTextReferenceDraft,
  createPdfVisualObjectReferenceDraft,
  describePdfBehavior,
  describePdfNativeObject,
  describePdfReference,
  describePdfVisualObject,
  normalizePdfReferenceRect,
  resolvePdfReference,
} from "./pdf/pdf-reference-adapter"
export type {
  PdfOfficeReference,
  PdfOfficeReferenceDraft,
  PdfReferenceContext,
  PdfReferenceDescriptorLabels,
  PdfTextReferenceSelection,
} from "./pdf/pdf-reference-adapter"
export {
  createPdfBehaviorLabels,
  createPdfBehaviorRecords,
  PDF_BEHAVIOR_REFERENCE_KIND,
  pdfBehaviorLocatorValue,
  pdfNativeObjectType,
} from "./pdf/pdf-native-objects"
export type {
  PdfAttachmentSource,
  PdfBehaviorKind,
  PdfBehaviorLabels,
  PdfBehaviorRecord,
  PdfBehaviorSource,
  PdfLayerSource,
  PdfPermissionsSource,
  PdfSignatureSource,
  PdfVisualObjectDetection,
  PdfVisualObjectProvider,
} from "./pdf/pdf-native-objects"
export {
  createPdfObjectRegistry,
  pdfBehaviorObjectDescriptor,
  pdfBehaviorReferenceKind,
  pdfNativeObjectDescriptor,
  pdfVisualObjectDescriptor,
  withPdfVisualPage,
  PDF_BEHAVIOR_OBJECT_PROVIDER,
  PDF_NATIVE_OBJECT_PROVIDER,
  PDF_STAGE2_KINDS,
  PDF_STAGE4_KINDS,
  PDF_VISUAL_KINDS,
  PDF_VISUAL_OBJECT_PROVIDER,
} from "./pdf/pdf-object-index"
export type { PdfVisualObjectRecord } from "./pdf/pdf-object-index"
export type {
  PdfAnnotationAppearancePatch,
  PdfFormFieldDescriptor,
  PdfFormFieldOption,
  PdfFormFieldType,
  PdfFormFieldValue,
  PdfOpenDocumentOptions,
  PdfPageRenderOptions,
  PdfPageRenderParameters,
  PdfPageTextLayer,
  PdfPageTextLayerGlyph,
  PdfPasswordHandler,
  PdfPasswordRequest,
  PdfRenderAnnotation,
  PdfRenderDocument,
  PdfRenderErrorCode,
  PdfRenderPageInfo,
  PdfRenderRect,
  PdfRenderRuntime,
  PdfRenderRuntimeConfig,
  PdfRotation,
  PdfSearchHit,
  PdfSearchState,
  PdfSearchStatus,
  PdfSelectionSegment,
  PdfSelectionState,
  PdfThumbnailRenderOptions,
} from "./pdf/pdf-render-runtime"
export type { PageTextSlice, PdfPageGeometry } from "@embedpdf/models"
export type {
  PdfDiagnostic,
  PdfLoadError,
  PdfLoadErrorCode,
  PdfLoadOptions,
  PdfSource,
  PdfUrlPolicy,
  PdfVerifiedDocument,
} from "./pdf-url-policy"

/** Page layout modes (continuous / single / spread) and their row geometry. */
export {
  PDF_PAGE_LAYOUTS,
  isPdfSpreadLayout,
  normalizePdfPageLayout,
  pdfPageAtPoint,
  pdfPageIndexAtOffset,
  pdfPageLayoutRows,
  pdfRowIndexAtOffset,
  pdfRowsInViewport,
  pdfSnapTopForRow,
  pdfSpreadPagePair,
  pdfSpreadStepPage,
  resolvePdfPageLayoutMetrics,
} from "./pdf/pdf-page-layout"
export type {
  PdfPageDisplaySize,
  PdfPageLayout,
  PdfPageLayoutMetrics,
  PdfPageLayoutOptions,
  PdfPageLayoutRow,
} from "./pdf/pdf-page-layout"

export type { PdfViewerProps } from "./types"
