<template>
  <div
    v-if="section && section.nodes.length > 0"
    ref="headerEl"
    data-docx-page-header="true"
    class="docx-page-header docx-page-section--header"
    :style="headerStyle"
  >
    <template v-for="(node, ni) of section.nodes" :key="`header-node-${pageIndex}-${ni}`">
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
        header-footer-region="header"
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
        header-footer-region="header"
      />
    </template>
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type {
  DocModel,
  DocxEditorController,
  HeaderSection,
} from "@extend-ai/docx-core"
import DocxParagraphHost from "./DocxParagraphHost.vue"
import DocxTableHost from "./DocxTableHost.vue"

// ── Props ──────────────────────────────────────────────────────────
const props = defineProps<{
  pageIndex: number
  section?: HeaderSection
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

const headerStyle = computed(() => ({
  display: "grid",
  gap: "0px",
  width: "100%",
  boxSizing: "border-box" as const,
  paddingBottom: "8px",
  borderBottom: `1px solid ${props.theme === "dark" ? "#374151" : "#e5e7eb"}`,
  color: props.theme === "dark" ? "#9ca3af" : "#6b7280",
}))

</script>

<style scoped>
.docx-page-header {
  font-size: 10pt;
}
.docx-page-header-node {
  width: 100%;
}
.docx-page-header-paragraph {
  margin: 0;
  white-space: pre-wrap;
}
.docx-page-header-unknown {
  color: #9ca3af;
  font-size: 10px;
}
</style>
