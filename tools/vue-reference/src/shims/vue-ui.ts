/**
 * @arcships/vue-ui is consumed through its published entry, which mixes every
 * component (and the signature_pad dependency) into the bundle. The reference
 * harness only needs the composables used by the PDF surface and the #157
 * primitives, so the package is aliased to this narrow source re-export.
 */
export {
  isEditableViewerTarget,
  resolveViewerKeyCommand,
  VIEWER_KEYBOARD_SHORTCUTS,
  useViewerKeys,
} from "../../../../packages/vue-ui/src/composables/viewer-keys";
export type {
  ViewerKeyCommand,
  ViewerKeyEvent,
  ViewerKeyMode,
} from "../../../../packages/vue-ui/src/composables/viewer-keys";
export { FILE_DROP_OUTLINE, useFileDrop } from "../../../../packages/vue-ui/src/composables/file-drop";

export { default as OfficeButton } from "../../../../packages/vue-ui/src/components/ui/OfficeButton.vue";
export { default as OfficeIconButton } from "../../../../packages/vue-ui/src/components/ui/OfficeIconButton.vue";
export { default as OfficeInput } from "../../../../packages/vue-ui/src/components/ui/OfficeInput.vue";
export { default as OfficeSelect } from "../../../../packages/vue-ui/src/components/ui/OfficeSelect.vue";
export { default as OfficeDialog } from "../../../../packages/vue-ui/src/components/ui/OfficeDialog.vue";
export { default as OfficeTooltip } from "../../../../packages/vue-ui/src/components/ui/OfficeTooltip.vue";
export { default as OfficeThemeProvider } from "../../../../packages/vue-ui/src/components/OfficeThemeProvider.vue";
export { default as OfficeSearchPopover } from "../../../../packages/vue-ui/src/components/chrome/OfficeSearchPopover.vue";
export { default as OfficePageControl } from "../../../../packages/vue-ui/src/components/chrome/OfficePageControl.vue";
export { default as OfficeZoomControl } from "../../../../packages/vue-ui/src/components/chrome/OfficeZoomControl.vue";
export { default as OfficeToolbar } from "../../../../packages/vue-ui/src/components/chrome/OfficeToolbar.vue";
export { default as OfficeToolbarGroup } from "../../../../packages/vue-ui/src/components/chrome/OfficeToolbarGroup.vue";
export { default as OfficeToolbarSeparator } from "../../../../packages/vue-ui/src/components/chrome/OfficeToolbarSeparator.vue";
export { default as OfficeBadge } from "../../../../packages/vue-ui/src/components/ui/OfficeBadge.vue";
export { default as OfficeEmptyState } from "../../../../packages/vue-ui/src/components/chrome/OfficeEmptyState.vue";
export { default as OfficeErrorState } from "../../../../packages/vue-ui/src/components/chrome/OfficeErrorState.vue";
export { default as OfficeLoadingState } from "../../../../packages/vue-ui/src/components/chrome/OfficeLoadingState.vue";
export { default as OfficeStatusBar } from "../../../../packages/vue-ui/src/components/chrome/OfficeStatusBar.vue";
export { default as DocumentViewerSidebar } from "../../../../packages/vue-ui/src/components/chrome/DocumentViewerSidebar.vue";
export { default as DocumentViewerSidebarSkeleton } from "../../../../packages/vue-ui/src/components/chrome/DocumentViewerSidebarSkeleton.vue";
export { default as FileThumbnail } from "../../../../packages/vue-ui/src/components/FileThumbnail.vue";
export { default as OfficeIcon } from "../../../../packages/vue-ui/src/components/OfficeIcon.vue";
export { HugeIcon } from "../../../../packages/vue-ui/src/components/HugeIcon";
export { default as OfficePopover } from "../../../../packages/vue-ui/src/components/ui/OfficePopover.vue";
export { default as OfficeFindBar } from "../../../../packages/vue-ui/src/components/OfficeFindBar.vue";
export { useElementWidth } from "../../../../packages/vue-ui/src/composables/element-width";
