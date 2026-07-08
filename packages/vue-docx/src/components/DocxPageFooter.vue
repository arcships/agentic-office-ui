<template>
  <div
    v-if="model && footerSection"
    ref="footerEl"
    data-docx-page-footer="true"
    class="docx-page-footer"
    :style="footerStyle"
    @dblclick="onFooterDoubleClick"
  >
    <div
      v-for="(node, ni) of footerSection.nodes"
      :key="`footer-node-${pageIndex}-${ni}`"
      class="docx-page-footer-node"
    >
      <template v-if="node.type === 'paragraph'">
        <p class="docx-page-footer-paragraph">
          {{ renderParagraphText(node) }}
        </p>
      </template>
      <div v-else class="docx-page-footer-unknown">
        (footer {{ node.type }})
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from "vue"
import type {
  DocxEditorController,
  ParagraphNode,
} from "@extend-ai/docx-core"
import { paragraphText } from "@extend-ai/docx-core"

// ── Props ──────────────────────────────────────────────────────────
const props = defineProps<{
  pageIndex: number
  footerNodes: any[]
  pageLayout: { pageWidthPx: number; pageHeightPx: number; marginsPx?: { top: number; bottom: number; left: number; right: number } }
  pageContentWidthPx: number
  theme: "light" | "dark"
  controller: DocxEditorController
}>()

// ── Computed ───────────────────────────────────────────────────────
const model = computed(() => props.controller.model)
const footerSection = computed(() => model.value?.metadata?.footerSections?.[0] ?? null)

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

// ── Helpers ────────────────────────────────────────────────────────
function renderParagraphText(node: any): string {
  if (node.type === "paragraph") {
    return paragraphText(node as ParagraphNode)
  }
  return ""
}

function onFooterDoubleClick(_event: MouseEvent): void {
  // Activate footer editing mode — simplified
}
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
