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
        :data-docx-header-node-index="ni"
        :editable="false"
        :document-theme="theme"
        :controller="controller"
        :numbering-definitions="model.metadata.numberingDefinitions"
        :show-tracked-changes="trackedChangesEnabled"
        :show-comment-highlights="commentsEnabled"
        :page-number="pageNumber"
        :total-pages="totalPages"
        :page-number-format="pageNumberFormat"
        :floating-page-origin-px="floatingPageOriginPx"
        :within-header-footer="true"
        header-footer-region="header"
      />
      <DocxTableHost
        v-else
        :table="node"
        :table-index="-(pageIndex * 1000 + ni + 1)"
        :data-docx-header-node-index="ni"
        :editable="false"
        :document-theme="theme"
        :controller="controller"
        :numbering-definitions="model.metadata.numberingDefinitions"
        :show-tracked-changes="trackedChangesEnabled"
        :show-comment-highlights="commentsEnabled"
        :page-number="pageNumber"
        :total-pages="totalPages"
        :page-number-format="pageNumberFormat"
        :floating-page-origin-px="floatingPageOriginPx"
        :within-header-footer="true"
        header-footer-region="header"
        :max-available-width-px="pageContentWidthPx"
      />
    </template>
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type {
  DocModel,
  DocxFloatingPageOriginPx,
  DocxEditorController,
  HeaderSection,
} from "@arcships/docx-core"
import DocxParagraphHost from "./DocxParagraphHost.vue"
import DocxTableHost from "./DocxTableHost.vue"

// ── Props ──────────────────────────────────────────────────────────
const props = defineProps<{
  pageIndex: number
  section?: HeaderSection
  model: DocModel
  pageLayout: {
    pageWidthPx: number
    pageHeightPx: number
    marginsPx?: { top: number; bottom: number; left: number; right: number }
    headerDistancePx?: number
    footerDistancePx?: number
  }
  pageContentWidthPx: number
  floatingPageOriginPx?: DocxFloatingPageOriginPx
  theme: "light" | "dark"
  controller?: DocxEditorController
  pageNumber: number
  totalPages: number
  pageNumberFormat?: string
  trackedChangesEnabled?: boolean
  commentsEnabled?: boolean
  /** Distance from the paper's top edge to the top of the header box. */
  headerDistancePx?: number
}>()

const resolvedHeaderDistancePx = computed(() => {
  const explicit = props.headerDistancePx
  if (Number.isFinite(explicit) && (explicit as number) >= 0) return Math.round(explicit as number)
  const fromLayout = props.pageLayout.headerDistancePx
  if (Number.isFinite(fromLayout) && (fromLayout as number) >= 0) return Math.round(fromLayout as number)
  return Math.round(props.pageLayout.marginsPx?.top ?? 72)
})

const headerStyle = computed(() => ({
  position: "absolute" as const,
  top: `${resolvedHeaderDistancePx.value}px`,
  left: `${props.pageLayout.marginsPx?.left ?? 72}px`,
  right: `${props.pageLayout.marginsPx?.right ?? 72}px`,
  display: "grid",
  gap: "0px",
  width: "auto",
  boxSizing: "border-box" as const,
}))

</script>

<style scoped>
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
