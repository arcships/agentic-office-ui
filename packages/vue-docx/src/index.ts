// ---- Components ----

export { default as DocxViewer } from "./components/DocxViewer.vue"
export { default as DocxEditor } from "./components/DocxEditor.vue"
/**
 * @deprecated Since 0.2.0. Use `DocxEditor`. Kept throughout 0.x; earliest
 * removal is 1.0.0.
 */
export { default as DocxEditorViewer } from "./components/DocxEditor.vue"
/**
 * @deprecated Since 0.2.0. Use `DocxViewer` or `DocxEditor`; page rendering is
 * an implementation detail. Kept throughout 0.x; earliest removal is 1.0.0.
 */
export { default as DocxViewerRoot } from "./components/DocxViewerRoot.vue"
/** @deprecated Since 0.2.0. Use `DocxViewer` or `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxPageWrapper } from "./components/DocxPageWrapper.vue"
/** @deprecated Since 0.2.0. Use `DocxViewer` or `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxPageSurface } from "./components/DocxPageSurface.vue"
/** @deprecated Since 0.2.0. Use `DocxViewer` or `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxPageHeader } from "./components/DocxPageHeader.vue"
/** @deprecated Since 0.2.0. Use `DocxViewer` or `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxPageFooter } from "./components/DocxPageFooter.vue"
/** @deprecated Since 0.2.0. Use `DocxViewer` or `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxPageBody } from "./components/DocxPageBody"
/** @deprecated Since 0.2.0. Use `DocxViewer` or `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxParagraphHost } from "./components/DocxParagraphHost.vue"
/** @deprecated Since 0.2.0. Use `DocxViewer` or `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxTableHost } from "./components/DocxTableHost.vue"
/** @deprecated Since 0.2.0. Use `DocxViewer` or `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxImageLayer } from "./components/DocxImageLayer.vue"
/** @deprecated Since 0.2.0. Use `DocxViewer` or `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxFormFieldLayer } from "./components/DocxFormFieldLayer.vue"
/** @deprecated Since 0.2.0. Use `DocxViewer` or `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxTrackedChangeGutter } from "./components/DocxTrackedChangeGutter.vue"
/** @deprecated Since 0.2.0. Use `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxContextMenu } from "./components/DocxContextMenu.vue"
/** @deprecated Since 0.2.0. Use `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxToolbar } from "./components/DocxToolbar.vue"
/** @deprecated Since 0.2.0. Use `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxThumbnailPanel } from "./components/DocxThumbnailPanel.vue"
/** @deprecated Since 0.2.0. Use `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxDragOverlay } from "./components/DocxDragOverlay.vue"

// ---- Render ----

/**
 * @deprecated Since 0.2.0. Static rendering helpers are not a stable document
 * surface; use `DocxViewer` or `DocxEditor`. Earliest removal: 1.0.0.
 */
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

export { useDocxPageThumbnails } from "./composables/useDocxPageThumbnails"
/**
 * @deprecated Since 0.2.0. Use `useDocxPageThumbnails`. Kept throughout 0.x;
 * earliest removal is 1.0.0.
 */
export {
  useDocxViewerThumbnails,
} from "./composables/useDocxPageThumbnails"

/**
 * @deprecated Since 0.2.0. Page-surface registration is owned by the high-level
 * viewer/editor. Earliest removal: 1.0.0.
 */
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

/**
 * @deprecated Since 0.2.0. Use `UseDocxPageThumbnailsOptions` and
 * `UseDocxPageThumbnailsResult`. Kept throughout 0.x; earliest removal is
 * 1.0.0.
 */
export type {
  UseDocxViewerThumbnailsOptions,
  DocxViewerThumbnails,
} from "@extend-ai/docx-core"

export type {
  UseDocxModelState,
} from "./composables/useDocxModel"
