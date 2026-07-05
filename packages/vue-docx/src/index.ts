/**
 * @extend-ai/vue-docx
 *
 * Vue 3 components and composables for DOCX viewing and editing.
 * Vue API aligned with the public @extend-ai/react-docx package.
 */

// Components
export { default as DocxViewer } from "./DocxViewer.vue"
export { default as DocxEditorViewer } from "./DocxEditor.vue"
export type { DocxViewerProps, DocxEditorViewerProps } from "./types"

// Composables
export {
  useDocxModel,
  useDocxEditor,
  useDocxDocumentTheme,
  useDocxTrackChanges,
  useDocxComments,
  useDocxPageLayout,
  useDocxPagination,
  useDocxParagraphStyles,
  useDocxLineSpacing,
  useDocxBorders,
  useDocxViewerThumbnails,
  useDocxPageThumbnails,
} from "./composables"

export type {
  UseDocxModelState,
  UseDocxEditorOptions,
  UseDocxDocumentThemeResult,
  UseDocxTrackChangesResult,
  UseDocxCommentsResult,
  UseDocxPageLayoutResult,
  UseDocxPaginationResult,
  UseDocxParagraphStylesResult,
  UseDocxLineSpacingResult,
  UseDocxBordersResult,
  UseDocxPageThumbnailsResult,
} from "./composables"

// Re-export core types for convenience
export type * from "@extend-ai/docx-core"

