<template>
  <OfficeThemeProvider :theme="isDark ? 'dark' : 'light'">
    <OfficeStatusBar class="xlsx-status-bar" data-testid="xlsx-status-bar">
      <template #start>
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
      </template>
      <template #end>
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
      </template>
    </OfficeStatusBar>
  </OfficeThemeProvider>
</template>

<script setup lang="ts">
import { computed } from "vue";
import { formatOfficeMessage } from "@arcships/office-interaction";
import type { XlsxViewerController } from "@arcships/xlsx-core";
import { OfficeStatusBar, OfficeThemeProvider } from "@arcships/vue-ui";
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
/* Geometry comes from `OfficeStatusBar`; these rules keep the spreadsheet
 * address/stats typography and the shared token palette. */
.xlsx-status-bar {
  flex: 0 0 auto;
  gap: var(--office-ui-space-3, 0.75rem);
}

.xlsx-status-bar__address {
  flex: 0 0 auto;
  font-family: var(--office-font-mono);
  max-inline-size: 40%;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.xlsx-status-bar__address--empty {
  color: var(--office-muted-foreground);
}

.xlsx-status-bar__stat,
.xlsx-status-bar__note {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.xlsx-status-bar__note {
  color: var(--office-info-foreground);
  font-weight: 600;
}

.xlsx-status-bar__truncated {
  cursor: help;
}

@container xlsx-viewer (max-width: 620px) {
  .xlsx-status-bar :deep(.office-status-bar__group) {
    overflow-x: auto;
  }
}
</style>
