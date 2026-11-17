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
        :data-docx-footer-node-index="ni"
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
        :data-docx-footer-node-index="ni"
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
        :max-available-width-px="pageContentWidthPx"
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
} from "@arcships/docx-core"
import DocxParagraphHost from "./DocxParagraphHost.vue"
import DocxTableHost from "./DocxTableHost.vue"

// ── Props ──────────────────────────────────────────────────────────
const props = defineProps<{
  pageIndex: number
  section?: FooterSection
  model: DocModel
  pageLayout: {
    pageWidthPx: number
    pageHeightPx: number
    marginsPx?: { top: number; bottom: number; left: number; right: number }
    headerDistancePx?: number
    footerDistancePx?: number
  }
  pageContentWidthPx: number
  theme: "light" | "dark"
  controller?: DocxEditorController
  pageNumber: number
  totalPages: number
  pageNumberFormat?: string
  trackedChangesEnabled?: boolean
  commentsEnabled?: boolean
  /** Distance from the paper's bottom edge to the bottom of the footer box. */
  footerDistancePx?: number
}>()

const resolvedFooterDistancePx = computed(() => {
  const explicit = props.footerDistancePx
  if (Number.isFinite(explicit) && (explicit as number) >= 0) return Math.round(explicit as number)
  const fromLayout = props.pageLayout.footerDistancePx
  if (Number.isFinite(fromLayout) && (fromLayout as number) >= 0) return Math.round(fromLayout as number)
  return Math.round(props.pageLayout.marginsPx?.bottom ?? 72)
})

const footerStyle = computed(() => ({
  position: "absolute" as const,
  bottom: `${resolvedFooterDistancePx.value}px`,
  left: `${props.pageLayout.marginsPx?.left ?? 72}px`,
  right: `${props.pageLayout.marginsPx?.right ?? 72}px`,
  display: "grid",
  gap: "0px",
  width: "auto",
  boxSizing: "border-box" as const,
}))

</script>

<style scoped>
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
