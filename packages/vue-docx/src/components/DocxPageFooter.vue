<template>
  <div
    v-if="section && section.nodes.length > 0"
    ref="footerEl"
    data-docx-page-footer="true"
    :data-docx-footer-part-name="section.partName"
    class="docx-page-footer docx-page-section--footer"
    :class="{ 'docx-page-section--editing': editable }"
    :style="footerStyle"
  >
    <template v-for="(node, ni) of section.nodes" :key="`footer-node-${pageIndex}-${ni}`">
      <DocxParagraphHost
        v-if="node.type === 'paragraph'"
        :paragraph="node"
        :paragraph-index="-(pageIndex * 1000 + ni + 1)"
        :data-docx-footer-node-index="ni"
        :editable="editable"
        :document-theme="theme"
        :controller="controller"
        :header-footer-part-name="section.partName"
        :section-story-node-index="ni"
        :numbering-definitions="model.metadata.numberingDefinitions"
        :auto-spacing-before-px="autoSpacingBeforePx(ni)"
        :auto-spacing-after-px="autoSpacingAfterPx(ni)"
        :default-tab-stop-px="defaultTabStopPx"
        :show-tracked-changes="trackedChangesEnabled"
        :show-comment-highlights="commentsEnabled"
        :page-number="pageNumber"
        :total-pages="totalPages"
        :page-number-format="pageNumberFormat"
        :section-pages="resolvedSectionPages"
        :resolve-style-ref-field-value="effectiveStyleRefResolver"
        :resolve-bookmark-field-value="effectiveBookmarkResolver"
        :resolve-bookmark-page-number="effectiveBookmarkPageResolver"
        :field-resolution-node-index="pageStartNodeIndex"
        :floating-page-origin-px="floatingPageOriginPx"
        :within-header-footer="true"
        header-footer-region="footer"
        @text-input="(_nodeIndex: number, text: string) => emit('textInput', ni, text)"
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
        :section-pages="resolvedSectionPages"
        :resolve-style-ref-field-value="effectiveStyleRefResolver"
        :resolve-bookmark-field-value="effectiveBookmarkResolver"
        :resolve-bookmark-page-number="effectiveBookmarkPageResolver"
        :model="model"
        :default-tab-stop-px="defaultTabStopPx"
        :floating-page-origin-px="floatingPageOriginPx"
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
  DocxFloatingPageOriginPx,
  DocxEditorController,
  FooterSection,
} from "@arcships/docx-core"
import {
  resolveDefaultTabStopPx,
  resolveDocumentLayoutDefaults,
  resolveParagraphAutoSpacingInStory,
} from "@arcships/docx-core"
import DocxParagraphHost from "./DocxParagraphHost.vue"
import DocxTableHost from "./DocxTableHost.vue"
import { createDocxFieldResolvers } from "../composables/docx-field-resolvers"

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
  floatingPageOriginPx?: DocxFloatingPageOriginPx
  theme: "light" | "dark"
  controller?: DocxEditorController
  pageNumber: number
  totalPages: number
  pageNumberFormat?: string
  sectionPages?: number
  pageStartNodeIndex?: number
  resolveStyleRefFieldValue?: (
    target: string,
    beforeNodeIndex?: number,
  ) => string | undefined
  resolveBookmarkFieldValue?: (target: string) => string | undefined
  resolveBookmarkPageNumber?: (target: string) => number | undefined
  trackedChangesEnabled?: boolean
  commentsEnabled?: boolean
  /** Distance from the paper's bottom edge to the bottom of the footer box. */
  footerDistancePx?: number
  /** Render footer paragraphs as content-editable (active footer editing). */
  editable?: boolean
}>()

const emit = defineEmits<{
  textInput: [nodeIndex: number, text: string]
}>()

// Header/footer stories share the document layout defaults: HTML automatic
// paragraph spacing and the default tab grid.
const layoutDefaults = computed(() => resolveDocumentLayoutDefaults(props.model.metadata))
const defaultTabStopPx = computed(() =>
  resolveDefaultTabStopPx(layoutDefaults.value.defaultTabStopTwips)
)

function autoSpacingBeforePx(nodeIndex: number): number {
  const node = props.section?.nodes[nodeIndex]
  if (!node || node.type !== "paragraph") return 0
  return resolveParagraphAutoSpacingInStory(
    props.section?.nodes ?? [],
    nodeIndex,
    node,
    layoutDefaults.value
  ).beforePx
}

function autoSpacingAfterPx(nodeIndex: number): number {
  const node = props.section?.nodes[nodeIndex]
  if (!node || node.type !== "paragraph") return 0
  return resolveParagraphAutoSpacingInStory(
    props.section?.nodes ?? [],
    nodeIndex,
    node,
    layoutDefaults.value
  ).afterPx
}

const resolvedFooterDistancePx = computed(() => {
  const explicit = props.footerDistancePx
  if (Number.isFinite(explicit) && (explicit as number) >= 0) return Math.round(explicit as number)
  const fromLayout = props.pageLayout.footerDistancePx
  if (Number.isFinite(fromLayout) && (fromLayout as number) >= 0) return Math.round(fromLayout as number)
  return Math.round(props.pageLayout.marginsPx?.bottom ?? 72)
})

const resolvedSectionPages = computed(() => props.sectionPages ?? props.totalPages)
const modelFieldResolvers = computed(() => createDocxFieldResolvers(props.model))
const effectiveStyleRefResolver = computed(
  () =>
    props.resolveStyleRefFieldValue ??
    ((target: string, beforeNodeIndex?: number) =>
      modelFieldResolvers.value.resolveStyleRefFieldValue(target, beforeNodeIndex)),
)
const effectiveBookmarkResolver = computed(
  () =>
    props.resolveBookmarkFieldValue ??
    ((target: string) => modelFieldResolvers.value.resolveBookmarkFieldValue(target)),
)
const effectiveBookmarkPageResolver = computed(
  () =>
    props.resolveBookmarkPageNumber ??
    ((target: string) => modelFieldResolvers.value.resolveBookmarkPageNumber(target)),
)

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
.docx-page-section--editing {
  outline: 1px dashed rgb(37 99 235 / 55%);
  outline-offset: 2px;
  border-radius: 2px;
}
.docx-page-section--editing :deep(.docx-paragraph-host--editable) {
  min-height: 1.4em;
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
