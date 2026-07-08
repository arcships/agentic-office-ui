<template>
  <div
    v-if="model && headerSection"
    ref="headerEl"
    data-docx-page-header="true"
    class="docx-page-header"
    :style="headerStyle"
    @dblclick="onHeaderDoubleClick"
  >
    <div
      v-for="(node, ni) of headerSection.nodes"
      :key="`header-node-${pageIndex}-${ni}`"
      class="docx-page-header-node"
    >
      <template v-if="node.type === 'paragraph'">
        <p class="docx-page-header-paragraph">
          {{ renderParagraphText(node) }}
        </p>
      </template>
      <div v-else class="docx-page-header-unknown">
        (header {{ node.type }})
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
  headerNodes: any[]
  pageLayout: { pageWidthPx: number; pageHeightPx: number; marginsPx?: { top: number; bottom: number; left: number; right: number } }
  pageContentWidthPx: number
  theme: "light" | "dark"
  controller: DocxEditorController
}>()

// ── Computed ───────────────────────────────────────────────────────
const model = computed(() => props.controller.model)
const headerSection = computed(() => model.value?.metadata?.headerSections?.[0] ?? null)

const headerStyle = computed(() => ({
  display: "grid",
  gap: "0px",
  width: "100%",
  boxSizing: "border-box" as const,
  paddingBottom: "8px",
  borderBottom: `1px solid ${props.theme === "dark" ? "#374151" : "#e5e7eb"}`,
  color: props.theme === "dark" ? "#9ca3af" : "#6b7280",
}))

// ── Helpers ────────────────────────────────────────────────────────
function renderParagraphText(node: any): string {
  if (node.type === "paragraph") {
    return paragraphText(node as ParagraphNode)
  }
  return ""
}

function onHeaderDoubleClick(_event: MouseEvent): void {
  // Activate header editing mode — simplified
}
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
