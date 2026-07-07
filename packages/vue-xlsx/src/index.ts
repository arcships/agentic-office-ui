import { defineComponent, h } from "vue"
import { useXlsxViewerController } from "./composables"

export { useXlsxViewerController } from "./composables"
export { XlsxFileSizeLimitExceededError } from "./composables"

export const XlsxViewer = defineComponent({
  name: "XlsxViewer",
  setup() {
    return () => h("div", { class: "xlsx-viewer-stub" }, "XLSX Viewer (pending)")
  },
})
