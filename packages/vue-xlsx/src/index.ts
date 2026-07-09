import { defineComponent, h } from "vue"
import { useXlsxViewerController } from "./composables"

export { useXlsxViewerController } from "./composables"
export { XlsxFileSizeLimitExceededError } from "./composables"

// Render module exports (chart SVG + surface WebGL rendering)
export { MemoChartSvg, MemoSurfaceChartComposite } from "./render"
export type {
  ChartRendererPalette,
  ChartSvgProps,
  ChartLayout,
  LegendItem,
} from "./render"

export const XlsxViewer = defineComponent({
  name: "XlsxViewer",
  setup() {
    return () => h("div", { class: "xlsx-viewer-stub" }, "XLSX Viewer (pending)")
  },
})
