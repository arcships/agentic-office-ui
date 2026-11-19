/**
 * @arcships/vue-ui is consumed through its published entry, which mixes every
 * component (and the signature_pad dependency) into the bundle. The PDF viewer
 * only needs the shared viewer-key and file-drop composables, so the harness
 * aliases the package to this narrow re-export.
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
