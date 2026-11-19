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
