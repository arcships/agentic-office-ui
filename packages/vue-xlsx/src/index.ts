import XlsxViewerComponent from "./components/XlsxViewer.vue"

export { useXlsxViewerController } from "./composables"
export { useXlsxViewerThumbnails } from "./composables/useXlsxViewerThumbnails"
export { XlsxFileSizeLimitExceededError } from "./composables"

// Render module exports (chart SVG + surface WebGL rendering)
export { MemoChartSvg, MemoSurfaceChartComposite } from "./render"
export type {
  ChartRendererPalette,
  ChartSvgProps,
  ChartLayout,
  LegendItem,
} from "./render"

// Component exports
export { default as XlsxGrid } from "./components/XlsxGrid.vue"
export { default as XlsxToolbar } from "./components/XlsxToolbar.vue"
export { default as XlsxRibbon } from "./components/XlsxRibbon.vue"
export { default as XlsxFormulaBar } from "./components/XlsxFormulaBar.vue"
export { default as XlsxSheetTabs } from "./components/XlsxSheetTabs.vue"
export { default as XlsxChartOverlay } from "./components/XlsxChartOverlay.vue"
export { default as XlsxImageLayer } from "./components/XlsxImageLayer.vue"
export { default as XlsxSelectionOverlay } from "./components/XlsxSelectionOverlay.vue"
export { default as XlsxContextMenu } from "./components/XlsxContextMenu.vue"

export const XlsxViewer = XlsxViewerComponent
