<template>
  <div class="xlsx-selection-overlay" />
</template>

<script setup lang="ts">
import type { XlsxViewerController, XlsxCellAddress, XlsxResolvedCellStyle } from "@extend-ai/xlsx-core";
import type { CSSProperties } from "vue";

defineProps<{
  controller: XlsxViewerController;
  getCellStyle?: ((cell: XlsxCellAddress) => Partial<CSSProperties> | undefined) | null;
  selectionColor?: string;
  selectionFillColor?: string;
}>();

// Selection rendering is handled by XlsxGrid canvas painting.
// This component is a placeholder for future DOM-based selection overlays
// (e.g., fill handle, resize cursors for images/charts).
</script>

<style scoped>
.xlsx-selection-overlay {
  height: 100%;
  left: 0;
  pointer-events: none;
  position: absolute;
  top: 0;
  width: 100%;
}
</style>
