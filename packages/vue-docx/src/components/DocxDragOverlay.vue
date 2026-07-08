<template>
  <Teleport to="body">
    <div
      v-if="visible"
      class="docx-drag-overlay"
      :style="overlayStyle"
    >
      <div class="docx-drag-cursor" />
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { ref, computed } from "vue"
import type { DocxEditorController } from "@extend-ai/docx-core"

const props = defineProps<{
  controller?: DocxEditorController
  visible?: boolean
  x?: number
  y?: number
}>()

const overlayStyle = computed(() => ({
  left: `${props.x ?? 0}px`,
  top: `${props.y ?? 0}px`,
}))
</script>

<style scoped>
.docx-drag-overlay {
  position: fixed;
  z-index: 9999;
  pointer-events: none;
}
.docx-drag-cursor {
  width: 2px;
  height: 20px;
  background: #2563eb;
  animation: docx-drag-pulse 1s infinite;
}
@keyframes docx-drag-pulse {
  0%, 100% { opacity: 1; }
  50% { opacity: 0.4; }
}
</style>
