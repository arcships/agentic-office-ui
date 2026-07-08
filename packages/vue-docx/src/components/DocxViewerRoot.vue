<template>
  <div ref="scrollContainer" class="docx-viewer-root" @scroll="onScroll">
    <div :style="spacerStyle">
      <div
        v-for="page of visiblePages"
        :key="page.number"
        :style="pageStyle(page)"
      >
        <DocxPageSurface
          :page="page.data"
          :editable="editable"
          :controller="controller"
        />
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from "vue"
import type {
  DocxEditorController,
  LayoutPage,
} from "@extend-ai/docx-core"
import { layoutDocument } from "@extend-ai/docx-core"
import DocxPageSurface from "./DocxPageSurface.vue"

const props = withDefaults(
  defineProps<{
    controller: DocxEditorController
    editable?: boolean
  }>(),
  { editable: true }
)

const scrollContainer = ref<HTMLElement | undefined>()

const pages = computed<LayoutPage[]>(() => {
  if (!props.controller.model) return []
  return layoutDocument(props.controller.model, {})
})

const spacerStyle = computed(() => ({
  height: `${totalHeightPx.value}px`,
  position: "relative" as const,
}))

const PAGE_GAP_PX = 16
const pageHeights = computed(() =>
  pages.value.map(() => 200)
)

const totalHeightPx = computed(() =>
  pageHeights.value.reduce((sum, h) => sum + h + PAGE_GAP_PX, 0)
)

interface VisiblePageEntry {
  number: number
  offsetTopPx: number
  data: LayoutPage
}

const visiblePages = computed<VisiblePageEntry[]>(() => {
  let offset = 0
  return pages.value.map((page, i) => {
    const entry: VisiblePageEntry = {
      number: page.number,
      offsetTopPx: offset,
      data: page,
    }
    offset += pageHeights.value[i] + PAGE_GAP_PX
    return entry
  })
})

function pageStyle(page: VisiblePageEntry) {
  return {
    position: "absolute" as const,
    top: `${page.offsetTopPx}px`,
    left: "50%",
    transform: "translateX(-50%)",
  }
}

function onScroll() {
  // Virtualization scroll handler
}
</script>

<style scoped>
.docx-viewer-root {
  flex: 1;
  overflow-y: auto;
  overflow-x: hidden;
  background: #f3f4f6;
  padding: 24px 0;
}
</style>
