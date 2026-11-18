<template>
  <div
    class="xlsx-selection-overlay"
    role="status"
    aria-live="polite"
    aria-atomic="true"
    :data-selection-address="selectionAddress ?? undefined"
    :data-cut-address="cutAddress ?? undefined"
  >
    <span class="xlsx-selection-overlay__announcement">
      {{ selectionAnnouncement }}
    </span>
  </div>
</template>

<script setup lang="ts">
import { computed, type CSSProperties } from "vue";
import type { XlsxViewerController, XlsxCellAddress } from "@arcships/xlsx-core";
import { formatOfficeMessage } from "@arcships/office-interaction";
import { useXlsxMessages } from "../i18n";
import { rangeToA1 } from "../composables/selection";

const props = defineProps<{
  controller: XlsxViewerController;
  getCellStyle?: ((cell: XlsxCellAddress) => Partial<CSSProperties> | undefined) | null;
  selectionColor?: string;
  selectionFillColor?: string;
}>();
const m = useXlsxMessages();

// The grid owns the visual selection. This companion layer gives the same
// selection a stable accessibility announcement without intercepting input.
const selectionAddress = computed(
  () => props.controller.selectedRangeAddress ?? props.controller.activeCellAddress
);
// A pending cut belongs to the sheet it was made on; switching sheets must not
// announce (or highlight) it there.
const cutAddress = computed(() => {
  const cut = props.controller.cutState;
  const sheet = props.controller.activeSheet;
  if (!cut || !sheet || cut.sheetIndex !== sheet.workbookSheetIndex) return null;
  return rangeToA1(cut.range);
});
const selectionAnnouncement = computed(() => {
  if (cutAddress.value) {
    return formatOfficeMessage(m.value.xlsx.selection.cutPending, { address: cutAddress.value });
  }
  return selectionAddress.value
    ? formatOfficeMessage(m.value.xlsx.selection.current, { address: selectionAddress.value })
    : m.value.xlsx.selection.none;
});
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
