import { defineComponent, h } from "vue"

export const XlsxViewer = defineComponent({
  name: "XlsxViewer",
  setup() {
    return () => h("div", { class: "xlsx-viewer-stub" }, "XLSX Viewer (pending)")
  },
})
