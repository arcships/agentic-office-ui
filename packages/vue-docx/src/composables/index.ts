// vue-docx composables — barrel export
// All Vue composables for the DOCX editor, organized by responsibility.

export { useDocxEditor } from "./useDocxEditor"

// Thin wrappers around editor controller properties
export { useDocxDocumentTheme } from "./useDocxDocumentTheme"
export { useDocxParagraphStyles } from "./useDocxParagraphStyles"
export { useDocxImageWrapMenu } from "./useDocxImageWrapMenu"
export { useDocxLineSpacing } from "./useDocxLineSpacing"
export { useDocxBorders } from "./useDocxBorders"
export { useDocxFormFields } from "./useDocxFormFields"
export { useDocxTrackChanges } from "./useDocxTrackChanges"
export { useDocxComments } from "./useDocxComments"
export { useDocxPageLayout } from "./useDocxPageLayout"
export { useDocxPagination } from "./useDocxPagination"

// Page surface registry (shared by viewer and thumbnail consumers)
export {
  createDocxViewerPageSurfaceRegistry,
  ensureDocxViewerPageSurfaceRegistry,
  subscribeDocxViewerPageSurfaces,
  notifyDocxViewerPageSurfaceListeners,
  docxViewerPageSurfaceRegistryOwner,
} from "./page-surface-registry"
export type {
  DocxViewerPageSurfaceRegistry,
  DocxViewerPageSurfaceSize,
  DocxPageThumbnailRenderSnapshot,
  DocxPageThumbnailSnapshotElement,
  DocxPageThumbnailTextRunSnapshot,
  DocxPageThumbnailTableCellSnapshot,
  DocxViewerPageThumbnailSnapshotEntry,
} from "./page-surface-registry"

// Thumbnails
export {
  useDocxPageThumbnails,
  useDocxViewerThumbnails,
} from "./useDocxPageThumbnails"
export type {
  DocxPageThumbnailStatus,
  DocxPageThumbnailItem,
  DocxViewerThumbnails,
  UseDocxPageThumbnailsOptions,
  UseDocxPageThumbnailsResult,
} from "./useDocxPageThumbnails"

// useDocxModel (standalone docx parser)
export { useDocxModel } from "./useDocxModel"
export type { UseDocxModelState } from "./useDocxModel"
