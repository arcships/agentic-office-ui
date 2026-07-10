<template>
  <div
    v-if="section && section.nodes.length > 0"
    ref="footerEl"
    data-docx-page-footer="true"
    class="docx-page-footer docx-page-section--footer"
    :style="footerStyle"
  >
    <template v-for="(node, ni) of section.nodes" :key="`footer-node-${pageIndex}-${ni}`">
      <DocxParagraphHost
        v-if="node.type === 'paragraph'"
        :paragraph="node"
        :paragraph-index="-(pageIndex * 1000 + ni + 1)"
        :editable="false"
        :document-theme="theme"
        :controller="controller"
        :numbering-definitions="model.metadata.numberingDefinitions"
        :show-tracked-changes="trackedChangesEnabled"
        :show-comment-highlights="commentsEnabled"
        :page-number="pageNumber"
        :total-pages="totalPages"
        :page-number-format="pageNumberFormat"
        :within-header-footer="true"
        header-footer-region="footer"
      />
      <DocxTableHost
        v-else
        :table="node"
        :table-index="-(pageIndex * 1000 + ni + 1)"
        :editable="false"
        :document-theme="theme"
        :controller="controller"
        :numbering-definitions="model.metadata.numberingDefinitions"
        :show-tracked-changes="trackedChangesEnabled"
        :show-comment-highlights="commentsEnabled"
        :page-number="pageNumber"
        :total-pages="totalPages"
        :page-number-format="pageNumberFormat"
        :within-header-footer="true"
        header-footer-region="footer"
      />
    </template>
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type {
  DocModel,
  DocxEditorController,
  FooterSection,
} from "@extend-ai/docx-core"
import DocxParagraphHost from "./DocxParagraphHost.vue"
import DocxTableHost from "./DocxTableHost.vue"

// ── Props ──────────────────────────────────────────────────────────
const props = defineProps<{
  pageIndex: number
  section?: FooterSection
  model: DocModel
  pageLayout: { pageWidthPx: number; pageHeightPx: number; marginsPx?: { top: number; bottom: number; left: number; right: number } }
  pageContentWidthPx: number
  theme: "light" | "dark"
  controller?: DocxEditorController
  pageNumber: number
  totalPages: number
  pageNumberFormat?: string
  trackedChangesEnabled?: boolean
  commentsEnabled?: boolean
}>()

const footerStyle = computed(() => ({
  display: "grid",
  gap: "0px",
  width: "100%",
  boxSizing: "border-box" as const,
  paddingTop: "8px",
  borderTop: `1px solid ${props.theme === "dark" ? "#374151" : "#e5e7eb"}`,
  color: props.theme === "dark" ? "#9ca3af" : "#6b7280",
  marginTop: "auto",
}))

</script>

<style scoped>
.docx-page-footer {
  font-size: 10pt;
}
.docx-page-footer-node {
  width: 100%;
}
.docx-page-footer-paragraph {
  margin: 0;
  white-space: pre-wrap;
}
.docx-page-footer-unknown {
  color: #9ca3af;
  font-size: 10px;
}
</style>
