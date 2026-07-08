import { defineComponent, h } from "vue"

// ---- Stub components (full implementations land in components/) ----

export const DocxViewer = defineComponent({
  name: "DocxViewer",
  setup() {
    return () => h("div", { class: "docx-viewer-stub" }, "DOCX Viewer (pending)")
  },
})

export const DocxEditorViewer = defineComponent({
  name: "DocxEditorViewer",
  setup() {
    return () => h("div", { class: "docx-editor-stub" }, "DOCX Editor (pending)")
  },
})

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
