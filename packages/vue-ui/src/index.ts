/**
 * @arcships/vue-ui
 *
 * Vue 3 UI components for document-centric applications.
 */

import "./print.css"
import "./styles/ui.css"

// Signature pad
export { default as SignaturePad } from "./components/SignaturePad.vue"

// File upload
export { default as FileUpload } from "./components/FileUpload.vue"

// File thumbnail
export { default as FileThumbnail } from "./components/FileThumbnail.vue"

// Bounding box citations
export { default as BoundingBoxCitations } from "./components/BoundingBoxCitations.vue"

// Layout blocks
export { default as LayoutBlocks } from "./components/LayoutBlocks.vue"

// Base UI
export { default as Spinner } from "./components/Spinner.vue"
export { default as Tooltip } from "./components/Tooltip.vue"
export { default as OfficeFindBar } from "./components/OfficeFindBar.vue"
export { default as OfficeLiveStatus } from "./components/OfficeLiveStatus.vue"

// Primitive layer (issue #157): token-driven Reka UI components
export * from "./components/ui"

// Theme provider and icon adapter
export { default as OfficeThemeProvider } from "./components/OfficeThemeProvider.vue"
export { default as OfficeIcon } from "./components/OfficeIcon.vue"
export {
  OFFICE_DARK_MEDIA_QUERY,
  OFFICE_DENSITY_ATTRIBUTE,
  OFFICE_THEME_ATTRIBUTE,
  OFFICE_THEME_INJECTION_KEY,
  OFFICE_THEME_RESOLVED_ATTRIBUTE,
  officeSystemPrefersDark,
  provideOfficeTheme,
  resolveOfficeTheme,
  useOfficeTheme,
} from "./composables/office-theme"
export type {
  OfficeDensity,
  OfficeResolvedTheme,
  OfficeTheme,
  OfficeThemeContext,
} from "./composables/office-theme"
export {
  OFFICE_ICONS,
  OFFICE_ICON_NAMES,
  OFFICE_ICON_REGISTRY_INJECTION_KEY,
  normalizeOfficeIconName,
  provideOfficeIcons,
  resolveOfficeIcon,
  useOfficeIcons,
} from "./icons"
export type {
  OfficeIconComponent,
  OfficeIconName,
  OfficeIconRegistry,
} from "./icons"

// Shared keyboard semantics for Office context menus
export {
  OFFICE_MENU_ITEM_SELECTOR,
  resolveOfficeMenuKeyCommand,
  useOfficeMenuKeyboard,
} from "./composables/office-menu-keyboard"
export type {
  OfficeMenuKeyboard,
  OfficeMenuKeyboardCommand,
  OfficeMenuKeyboardOptions,
} from "./composables/office-menu-keyboard"

// Shared viewer keyboard map (DOCX / PDF / PPTX)
export {
  VIEWER_KEYBOARD_SHORTCUTS,
  isEditableViewerTarget,
  isInteractiveViewerTarget,
  resolveViewerKeyCommand,
  useViewerKeys,
} from "./composables/viewer-keys"
export type {
  ViewerKeyCommand,
  ViewerKeyEvent,
  ViewerKeyboard,
  ViewerKeyMapOptions,
  ViewerKeyMode,
  ViewerKeyOptions,
  ViewerKeySwitch,
} from "./composables/viewer-keys"

// Shared drag-and-drop file import (DOCX / XLSX / PDF / PPTX)
export {
  FILE_DROP_OUTLINE,
  fileDropExtension,
  isSupportedDropFile,
  useFileDrop,
} from "./composables/file-drop"
export type {
  FileDropAccept,
  FileDropCandidate,
  UseFileDropOptions,
  UseFileDropReturn,
} from "./composables/file-drop"

// Office selection primitives
export { default as OfficeObjectOutlineLayer } from "./components/OfficeObjectOutlineLayer.vue"
export { default as OfficeRegionSelector } from "./components/OfficeRegionSelector.vue"
export { default as OfficeRegionAnnotator } from "./components/OfficeRegionAnnotator.vue"

// Region capture and annotation documents
export * from "./region-capture"
export * from "./region-annotations"

export type * from "./types"
