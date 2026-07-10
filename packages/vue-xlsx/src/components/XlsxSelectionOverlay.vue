<template>
  <div
    class="xlsx-selection-overlay"
    role="status"
    aria-live="polite"
    aria-atomic="true"
    :data-selection-address="selectionAddress ?? undefined"
  >
    <span class="xlsx-selection-overlay__announcement">
      {{ selectionAnnouncement }}
    </span>
  </div>
</template>

<script setup lang="ts">
import { computed, type CSSProperties } from "vue";
import type { XlsxViewerController, XlsxCellAddress } from "@extend-ai/xlsx-core";

const props = defineProps<{
  controller: XlsxViewerController;
  getCellStyle?: ((cell: XlsxCellAddress) => Partial<CSSProperties> | undefined) | null;
  selectionColor?: string;
  selectionFillColor?: string;
}>();

// The grid owns the visual selection. This companion layer gives the same
// selection a stable accessibility announcement without intercepting input.
const selectionAddress = computed(
  () => props.controller.selectedRangeAddress ?? props.controller.activeCellAddress
);
const selectionAnnouncement = computed(() =>
  selectionAddress.value
    ? `当前选择 ${selectionAddress.value}`
    : "当前没有选择单元格"
);
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
.xlsx-selection-overlay__announcement {
  clip: rect(0 0 0 0);
  clip-path: inset(50%);
  height: 1px;
  overflow: hidden;
  position: absolute;
  white-space: nowrap;
  width: 1px;
}
</style>
