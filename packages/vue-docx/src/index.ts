// ---- Components ----

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

export { default as DocxViewer } from "./components/DocxViewer.vue"
export { default as DocxEditor } from "./components/DocxEditor.vue"
/**
 * Minimal embeddable DOCX renderer for third-party consumers.
 *
 * Renders the paged document surface with built-in annotation gutter
 * (comments + tracked changes). No toolbar, thumbnails, file loading,
 * or outer chrome — host owns all controls.
 *
 * @example
 * ```vue
 * <DocxDocumentSurface
 *   :model="docModel"
 *   :zoom-scale="120"
 *   :show-comments="true"
 *   @page-count-change="total = $event"
 *   ref="surfaceRef"
 * />
 * ```
 * // surfaceRef.value.scrollToPage(3)
 */
export { default as DocxDocumentSurface } from "./components/DocxDocumentSurface.vue"
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
/** Page-setup dialog: margins, paper size and orientation for a section. */
export { default as DocxPageSetupDialog } from "./components/DocxPageSetupDialog.vue"
export type {
  DocxPageSetupApplyPatch,
  DocxPageSetupDialogProps,
  DocxPageSetupSectionOption,
} from "./components/page-setup-dialog-types"
/** @deprecated Since 0.2.0. Use `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxThumbnailPanel } from "./components/DocxThumbnailPanel.vue"
/** @deprecated Since 0.2.0. Use `DocxEditor`. Earliest removal: 1.0.0. */
export { default as DocxDragOverlay } from "./components/DocxDragOverlay.vue"
/**
 * Form-field properties dialog: edits a field's value and configuration
 * (type, name, placeholder, required, dropdown options, maximum length) through
 * one `controller.updateFormField()` call, so a save is one undo entry.
 */
export { default as DocxFormFieldDialog } from "./components/DocxFormFieldDialog.vue"
/**
 * Link hover card: open/edit/remove actions for the link under the pointer.
 * Edit and remove select the linked characters before calling
 * `controller.setLink()`. Used by `DocxEditor`; exported for custom surfaces.
 */
export { default as DocxLinkHoverCard } from "./components/DocxLinkHoverCard.vue"
/**
 * Paragraph-style preview popover: renders a sample paragraph with the hovered
 * style's own run/paragraph properties.
 */
export { default as DocxStylePreviewPopover } from "./components/DocxStylePreviewPopover.vue"
/** Absolute highlight boxes for document blocks on the mounted pages. */
export { default as DocxBlockHighlights } from "./components/DocxBlockHighlights.vue"
/** Absolute highlight boxes for word ranges on the mounted pages. */
export { default as DocxWordHighlights } from "./components/DocxWordHighlights.vue"
export type {
  DocxBlockHighlightsProps,
  DocxWordHighlightsProps,
} from "./components/block-highlight-types"

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

// ---- Annotation gutter ----

/**
 * Screen-space metrics for the DOCX annotation gutter. Exposed so hosts that
 * lay out their own scroll surface can reserve the gutter's footprint.
 */
export {
  DOCX_ANNOTATION_GUTTER_WIDTH_PX,
  DOCX_ANNOTATION_GUTTER_GAP_PX,
  DOCX_ANNOTATION_GUTTER_DRAWER_MAX_WIDTH_PX,
  DOCX_ANNOTATION_GUTTER_DRAWER_WIDTH_PX,
  DOCX_ANNOTATION_GUTTER_DRAWER_TAB_WIDTH_PX,
  annotationGutterReservedWidthPx,
  annotationGutterPaneStyle,
  annotationGutterDrawerPaneStyle,
  annotationGutterUsesDrawer,
  annotationGutterCardTopPx,
  annotationGutterMarkerStyle,
  annotationGutterLeaderStyle,
  normalizeAnnotationGutterZoom,
} from "./components/gutter-metrics"
export {
  useDocxGutterLayout,
  type DocxGutterLayout,
  type DocxGutterLayoutOptions,
} from "./composables/gutter-layout"

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
  useDocxNodeReorder,
  docxReorderNodeIndexFromElement,
  resolveDocxReorderInsertionIndex,
  resolveDocxReorderTarget,
  DOCX_REORDER_NODE_SELECTOR,
  type DocxReorderFrame,
  type DocxReorderHover,
  type DocxReorderRect,
  type UseDocxNodeReorderOptions,
  type UseDocxNodeReorderResult,
} from "./composables/useDocxNodeReorder"
export {
  docxStylePreviewCss,
  type DocxStylePreview,
} from "./composables/docx-style-preview"
export {
  parseDocxFormFieldLocationKey,
  resolveFormFieldAt,
} from "./composables/docx-form-fields"
export {
  findDocxSearchMatches,
  type DocxNestedTableStep,
  type DocxSearchBoundary,
  type DocxSearchLocation,
  type DocxSearchMatch,
  type DocxSearchOptions,
  type DocxSearchRange,
  type DocxSearchRegion,
  type DocxSearchState,
  type DocxSearchStatus,
  type ParagraphLocationInNestedTableCell,
  type ParagraphLocationInNoteText,
} from "./composables/useDocxSearch"
export { findDocxSearchTargetPage } from "./composables/docx-search-page"
export type { DocxSearchPageTarget } from "./composables/docx-search-page"

export { useDocxPageSetup } from "./composables/useDocxPageSetup"
export type {
  DocxPageSetupPatch,
  DocxPageSetupSection,
  UseDocxPageSetupResult,
} from "./composables/useDocxPageSetup"

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

export type { DocModel } from "@arcships/docx-core"
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
  DocxFormFieldUpdate,
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
  DocxPageMargins,
  DocxPageMarginsUpdate,
  DocxPageOrientation,
  DocxPageSize,
  DocxPageSizePreset,
  DocxPageSizeUpdate,
  DocxSectionPageSetup,
  DocxPageSetupCommandOptions,
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
} from "@arcships/docx-core"

/**
 * @deprecated Since 0.2.0. Use `UseDocxPageThumbnailsOptions` and
 * `UseDocxPageThumbnailsResult`. Kept throughout 0.x; earliest removal is
 * 1.0.0.
 */
export type {
  UseDocxViewerThumbnailsOptions,
  DocxViewerThumbnails,
} from "@arcships/docx-core"

export type {
  UseDocxModelState,
} from "./composables/useDocxModel"
