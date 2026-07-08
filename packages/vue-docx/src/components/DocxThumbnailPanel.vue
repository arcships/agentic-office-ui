<template>
  <aside class="docx-thumbnail-panel">
    <div class="docx-thumbnail-header">
      <span class="docx-thumbnail-title">Pages</span>
      <span class="docx-thumbnail-count">{{ thumbnailItems.length }}</span>
    </div>
    <div class="docx-thumbnail-list">
      <button
        v-for="item of thumbnailItems"
        :key="item.pageNumber"
        class="docx-thumbnail-item"
        :class="{ 'docx-thumbnail-active': item.pageNumber === currentPage }"
        @click="onSelectPage(item.pageNumber)"
      >
        <canvas
          v-if="item.canvas"
          ref="thumbnailCanvases"
          class="docx-thumbnail-canvas"
          :width="item.canvas.width"
          :height="item.canvas.height"
        />
        <div v-else class="docx-thumbnail-placeholder">
          {{ item.pageNumber }}
        </div>
        <span class="docx-thumbnail-label">Page {{ item.pageNumber }}</span>
      </button>
    </div>
  </aside>
</template>

<script setup lang="ts">
import { ref, computed, onMounted } from "vue"
import type { DocxEditorController } from "@extend-ai/docx-core"

const props = defineProps<{
  controller?: DocxEditorController
}>()

const currentPage = computed(() => props.controller?.currentPage ?? 1)

interface ThumbnailItem {
  pageNumber: number
  canvas?: HTMLCanvasElement
}

const thumbnailItems = ref<ThumbnailItem[]>([])

function onSelectPage(pageNumber: number) {
  // Scroll to page (full implementation uses page surface registry)
}

onMounted(() => {
  const totalPages = props.controller?.totalPages ?? 1
  thumbnailItems.value = Array.from({ length: totalPages }, (_, i) => ({
    pageNumber: i + 1,
  }))
})
</script>

<style scoped>
.docx-thumbnail-panel {
  width: 180px;
  background: #f9fafb;
  border-right: 1px solid #e5e7eb;
  display: flex;
  flex-direction: column;
  overflow: hidden;
  flex-shrink: 0;
}
.docx-thumbnail-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 12px;
  border-bottom: 1px solid #e5e7eb;
}
.docx-thumbnail-title {
  font-size: 12px;
  font-weight: 600;
  color: #6b7280;
  text-transform: uppercase;
}
.docx-thumbnail-count {
  font-size: 11px;
  color: #9ca3af;
}
.docx-thumbnail-list {
  flex: 1;
  overflow-y: auto;
  padding: 8px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.docx-thumbnail-item {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 4px;
  padding: 8px;
  border: 2px solid transparent;
  border-radius: 4px;
  background: none;
  cursor: pointer;
  width: 100%;
}
.docx-thumbnail-item:hover {
  background: #f3f4f6;
}
.docx-thumbnail-active {
  border-color: #2563eb;
  background: #eff6ff;
}
.docx-thumbnail-canvas {
  width: 100%;
  border: 1px solid #e5e7eb;
  background: #fff;
}
.docx-thumbnail-placeholder {
  width: 100%;
  aspect-ratio: 0.75;
  background: #fff;
  border: 1px solid #e5e7eb;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 14px;
  color: #9ca3af;
}
.docx-thumbnail-label {
  font-size: 11px;
  color: #6b7280;
}
</style>
