<template>
  <Teleport to="body">
    <div
      v-if="visible"
      class="docx-drag-overlay"
      :style="overlayStyle"
    >
      <!-- Import file overlay -->
      <div v-if="importMode" class="docx-drag-overlay-import">
        <div class="docx-drag-overlay-import-icon">📂</div>
        <div class="docx-drag-overlay-import-text">
          Drop DOCX file to import
        </div>
      </div>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type { DocxEditorController } from "@arcships/docx-core"

// ── Props ──────────────────────────────────────────────────────────
defineProps<{
  controller: DocxEditorController
}>()

// ── State (managed externally via provide/inject or parent) ─────────
const visible = computed(() => false)
const importMode = computed(() => false)

// ── Styles ─────────────────────────────────────────────────────────
const overlayStyle = computed(() => ({
  position: "fixed" as const,
  inset: "0",
  zIndex: "9998",
  pointerEvents: "none" as const,
}))
</script>

<style scoped>
.docx-drag-overlay {
  position: fixed;
  inset: 0;
  z-index: 9998;
  pointer-events: none;
}
.docx-drag-overlay-import {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  width: 100%;
  height: 100%;
  background: rgba(59, 130, 246, 0.08);
  border: 3px dashed #3b82f6;
  border-radius: 12px;
}
.docx-drag-overlay-import-icon {
  font-size: 48px;
  margin-bottom: 12px;
}
.docx-drag-overlay-import-text {
  font-size: 16px;
  font-weight: 600;
  color: #2563eb;
}
</style>
