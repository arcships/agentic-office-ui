import { defineComponent, h } from "vue"

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

export { useDocxEditor } from "./composables"

export type { DocModel } from "@extend-ai/docx-core"
