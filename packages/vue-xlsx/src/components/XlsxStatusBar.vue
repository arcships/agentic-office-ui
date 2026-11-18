<template>
  <div
    class="xlsx-status-bar"
    :class="{ 'xlsx-status-bar--dark': isDark }"
    data-testid="xlsx-status-bar"
    role="status"
  >
    <span
      v-if="address"
      class="xlsx-status-bar__address"
      data-testid="xlsx-status-range"
      :title="m.xlsx.statusBar.range"
    >{{ address }}</span>
    <span
      v-else
      class="xlsx-status-bar__address xlsx-status-bar__address--empty"
      data-testid="xlsx-status-ready"
    >{{ m.xlsx.statusBar.ready }}</span>

    <div class="xlsx-status-bar__stats">
      <span
        v-if="summary && summary.count > 0"
        class="xlsx-status-bar__stat"
        data-testid="xlsx-status-count"
        :data-stat-value="summary.count"
      >{{ m.xlsx.statusBar.count }}: {{ summary.count }}</span>
      <span
        v-if="summary && summary.numericCount > 0 && summary.numericCount !== summary.count"
        class="xlsx-status-bar__stat"
        data-testid="xlsx-status-numeric-count"
        :data-stat-value="summary.numericCount"
      >{{ m.xlsx.statusBar.numericalCount }}: {{ summary.numericCount }}</span>
      <span
        v-if="sumText"
        class="xlsx-status-bar__stat"
        data-testid="xlsx-status-sum"
        :data-stat-value="sumText"
      >{{ m.xlsx.statusBar.sum }}: {{ sumText }}</span>
      <span
        v-if="averageText"
        class="xlsx-status-bar__stat"
        data-testid="xlsx-status-average"
        :data-stat-value="averageText"
      >{{ m.xlsx.statusBar.average }}: {{ averageText }}</span>
      <span
        v-if="minText"
        class="xlsx-status-bar__stat"
        data-testid="xlsx-status-min"
        :data-stat-value="minText"
      >{{ m.xlsx.statusBar.min }}: {{ minText }}</span>
      <span
        v-if="maxText"
        class="xlsx-status-bar__stat"
        data-testid="xlsx-status-max"
        :data-stat-value="maxText"
      >{{ m.xlsx.statusBar.max }}: {{ maxText }}</span>
      <span
        v-if="summary?.truncated"
        class="xlsx-status-bar__stat xlsx-status-bar__truncated"
        data-testid="xlsx-status-truncated"
        :title="truncatedTitle"
        :aria-label="truncatedTitle"
      >…</span>
      <span
        v-if="note"
        class="xlsx-status-bar__note"
        data-testid="xlsx-status-note"
      >{{ note }}</span>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue";
import { formatOfficeMessage } from "@arcships/office-interaction";
import type { XlsxViewerController } from "@arcships/xlsx-core";
import { useXlsxMessages } from "../i18n";
import { normalizeRange } from "../composables/selection";
import { useXlsxSelectionStats } from "../composables/selection-stats";
import { formatSelectionStatValue } from "./selection-status-format";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
}>();

const m = useXlsxMessages();
const stats = useXlsxSelectionStats(() => props.controller);
const summary = computed(() => stats.summary.value);

const address = computed(() => props.controller.selectedRangeAddress ?? props.controller.activeCellAddress ?? "");

const sumText = computed(() => formatSelectionStatValue(
  summary.value?.sum,
  summary.value?.numberFormat,
  props.controller.locale,
));
const averageText = computed(() => formatSelectionStatValue(
  summary.value?.average,
  summary.value?.numberFormat,
  props.controller.locale,
));
const minText = computed(() => formatSelectionStatValue(
  summary.value?.min,
  summary.value?.numberFormat,
  props.controller.locale,
));
const maxText = computed(() => formatSelectionStatValue(
  summary.value?.max,
  summary.value?.numberFormat,
  props.controller.locale,
));

const truncatedTitle = computed(() => formatOfficeMessage(
  m.value.xlsx.statusBar.truncated,
  { count: stats.request.value?.maxCells ?? summary.value?.cellCount ?? 0 },
));

const note = computed(() => {
  const selection = props.controller.selection;
  if (!selection) return "";
  if (stats.filtered.value) return m.value.xlsx.statusBar.filtered;
  const range = normalizeRange(selection);
  const multiCell = range.start.row !== range.end.row || range.start.col !== range.end.col;
  return multiCell ? m.value.xlsx.statusBar.selected : "";
});
</script>

<style scoped>
.xlsx-status-bar {
  align-items: center;
  background: #f8f9fa;
  border-top: 1px solid #e4e4e7;
  color: #52525b;
  display: flex;
  flex: 0 0 auto;
  font-size: 11px;
  gap: 12px;
  justify-content: space-between;
  min-height: 24px;
  padding: 2px 10px;
  user-select: none;
}

.xlsx-status-bar--dark {
  background: #27272a;
  border-color: #3f3f46;
  color: #d4d4d8;
}

.xlsx-status-bar__address {
  flex: 0 0 auto;
  font-family: "SF Mono", "Menlo", "Monaco", monospace;
  max-width: 40%;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.xlsx-status-bar__address--empty {
  color: #71717a;
}

.xlsx-status-bar--dark .xlsx-status-bar__address--empty {
  color: #a1a1aa;
}

.xlsx-status-bar__stats {
  align-items: center;
  display: flex;
  gap: 12px;
  margin-inline-start: auto;
  min-width: 0;
  overflow: hidden;
}

.xlsx-status-bar__stat,
.xlsx-status-bar__note {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.xlsx-status-bar__note {
  color: #2563eb;
  font-weight: 600;
}

.xlsx-status-bar--dark .xlsx-status-bar__note {
  color: #93c5fd;
}

.xlsx-status-bar__truncated {
  cursor: help;
}

@container xlsx-viewer (max-width: 620px) {
  .xlsx-status-bar__stats {
    overflow-x: auto;
  }
}
</style>
