// ---- Components ----

export { default as DocxViewer } from "./components/DocxViewer.vue"
export { default as DocxEditor } from "./components/DocxEditor.vue"
export { default as DocxEditorViewer } from "./components/DocxEditor.vue"
export { default as DocxViewerRoot } from "./components/DocxViewerRoot.vue"
export { default as DocxPageWrapper } from "./components/DocxPageWrapper.vue"
export { default as DocxPageSurface } from "./components/DocxPageSurface.vue"
export { default as DocxPageHeader } from "./components/DocxPageHeader.vue"
export { default as DocxPageFooter } from "./components/DocxPageFooter.vue"
export { default as DocxPageBody } from "./components/DocxPageBody"
export { default as DocxParagraphHost } from "./components/DocxParagraphHost.vue"
export { default as DocxTableHost } from "./components/DocxTableHost.vue"
export { default as DocxImageLayer } from "./components/DocxImageLayer.vue"
export { default as DocxFormFieldLayer } from "./components/DocxFormFieldLayer.vue"
export { default as DocxTrackedChangeGutter } from "./components/DocxTrackedChangeGutter.vue"
export { default as DocxContextMenu } from "./components/DocxContextMenu.vue"
export { default as DocxToolbar } from "./components/DocxToolbar.vue"
export { default as DocxThumbnailPanel } from "./components/DocxThumbnailPanel.vue"
export { default as DocxDragOverlay } from "./components/DocxDragOverlay.vue"

// ---- Render ----

export {
  renderParagraphRuns,
  renderStaticHtml,
  type ParagraphRunRenderOptions,
} from "./render"

// ---- Composables ----

export { useDocxEditor } from "./composables/useDocxEditor"
export {
  useDocxDocumentTheme,
  useDocxParagraphStyles,
  useDocxImageWrapMenu,
  useDocxLineSpacing,
  useDocxBorders,
  useDocxFormFields,
  useDocxTrackChanges,
  useDocxComments,
  useDocxPageLayout,
  useDocxPagination,
} from "./composables"

export { useDocxModel } from "./composables/useDocxModel"

export {
  useDocxPageThumbnails,
  useDocxViewerThumbnails,
} from "./composables/useDocxPageThumbnails"

export {
  createDocxViewerPageSurfaceRegistry,
  ensureDocxViewerPageSurfaceRegistry,
  subscribeDocxViewerPageSurfaces,
  notifyDocxViewerPageSurfaceListeners,
} from "./composables/page-surface-registry"

// ---- Type re-exports ----

export type { DocModel } from "@extend-ai/docx-core"
export type {
  UseDocxEditorOptions,
  DocxEditorController,
  DocxEditorViewerProps,
  DocxEditorViewerMode,
  DocxEditorSelection,
  DocxTextRange,
  DocxTextRangeLocation,
  DocxDocumentTheme,
  DocxFormFieldLocation,
  DocxSelectedFormField,
  DocxImageLocation,
  DocxImageDropTarget,
  DocxHeadingStyleMap,
  DocxPageVirtualizationOptions,
  DocxVisiblePageRange,
  DocxContextMenuAction,
  DocxContextMenuActionId,
  DocxContextMenuContext,
  DocxContextMenuRenderProps,
  DocxImageWrapMenuOption,
  DocxImageWrapMode,
  DocxImageWrapState,
  DocxTableContextMenuAction,
  DocxTableContextMenuActionId,
  DocxTableContextMenuContext,
  DocxTableContextMenuRenderProps,
  DocxTrackedChangeCardRenderProps,
  DocxComment,
  DocxCommentCardRenderProps,
  UseDocxCommentsResult,
  DocxPageLayoutInfo,
  DocxPaginationInfo,
  DocxLineSpacingInfo,
  DocxBorderContext,
  DocxBorderPreset,
  DocxBorderPresetState,
  DocxSectionColumnLayout,
  DocxListType,
  DocxTrackedChange,
  DocxTrackedChangeKind,
  UseDocxDocumentThemeResult,
  UseDocxImageWrapMenuResult,
  UseDocxBordersResult,
  UseDocxLineSpacingResult,
  UseDocxFormFieldsResult,
  UseDocxViewerThumbnailsOptions,
  DocxViewerThumbnails,
  UseDocxPageThumbnailsOptions,
  UseDocxPageThumbnailsResult,
  UseDocxPageLayoutResult,
  DocxPageThumbnailItem,
  DocxPageThumbnailBounds,
  DocxPageThumbnailRenderWindow,
  DocxPageThumbnailResolution,
  DocxPageThumbnailResolutionOptions,
  DocxPageThumbnailStatus,
  UseDocxPaginationResult,
  UseDocxParagraphStylesResult,
  UseDocxTrackChangesResult,
} from "@extend-ai/docx-core"

export type {
  UseDocxModelState,
} from "./composables/useDocxModel"
