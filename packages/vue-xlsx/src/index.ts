import XlsxViewerComponent from "./components/XlsxViewer.vue"

export { useXlsxViewerController } from "./composables"
export { useXlsxViewerThumbnails } from "./composables/useXlsxViewerThumbnails"
export { XlsxFileSizeLimitExceededError } from "./composables"
export type {
  XlsxDiagnostic,
  XlsxLoadError,
  XlsxLoadErrorCode,
  XlsxSourceKind,
  XlsxSourceState,
  XlsxUrlPolicy,
} from "@arcships/xlsx-core"

// Optional renderers keep their historical root names, but load the heavy
// implementation only when a chart actually mounts.
/**
 * @deprecated Since 0.2.0. Rendering modules are owned by `XlsxViewer`; use
 * the high-level component and controller. Earliest removal: 1.0.0.
 */
export { MemoChartSvg, MemoSurfaceChartComposite } from "./optional/lazy-renderers"
/**
 * @deprecated Since 0.2.0. Rendering types are implementation details of
 * `XlsxViewer`. Kept throughout 0.x; earliest removal is 1.0.0.
 */
export type {
  ChartRendererPalette,
  ChartSvgProps,
  ChartLayout,
  LegendItem,
} from "./render/chart-types"

// Component exports
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxGrid } from "./components/XlsxGrid.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxToolbar } from "./components/XlsxToolbar.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxRibbon } from "./components/XlsxRibbon.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxFormulaBar } from "./components/XlsxFormulaBar.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxSheetTabs } from "./components/XlsxSheetTabs.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxChartOverlay } from "./components/XlsxChartOverlay.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxImageLayer } from "./components/XlsxImageLayer.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxSelectionOverlay } from "./components/XlsxSelectionOverlay.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxContextMenu } from "./components/XlsxContextMenu.vue"

export const XlsxViewer = XlsxViewerComponent
