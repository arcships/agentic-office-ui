<template>
  <OfficeThemeProvider :theme="isDark ? 'dark' : 'light'">
    <!--
      Layout source: Extend's status strip (`file-system.tsx`): a 28px
      `aria-live` footer with `border-t bg-muted/40 px-3 text-xs` and centred
      `gap-1` segments. The pinned XLSX viewer has no status bar of its own, so
      the spreadsheet keeps its address/stat semantics inside that strip.
    -->
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
/*
 * Geometry and colour come from Extend's status strip; `OfficeStatusBar`
 * supplies the slots, `role` and aria wiring, and these rules map its shell
 * onto that strip (`h-7 px-3 justify-center gap-1 text-xs`).
 */
.xlsx-status-bar {
  flex: 0 0 auto;
  justify-content: center;
  min-block-size: 1.75rem;
  padding: 0 12px;
  border-block-start: var(--office-border-width, 1px) solid var(--office-border);
  background: color-mix(in oklab, var(--office-muted) 40%, transparent);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}

/* One centred line: the address, the statistics and the note are segments of
 * the same strip instead of a left/right split. */
.xlsx-status-bar :deep(.office-status-bar__start) {
  flex: 0 1 auto;
  flex-wrap: nowrap;
  justify-content: center;
  gap: 0 var(--office-ui-space-1, 0.25rem);
  overflow: hidden;
}

.xlsx-status-bar :deep(.office-status-bar__end) {
  display: none;
}

/* The segment separator is decoration: it renders like the `· ` prefix the
 * reference puts inside its spans without polluting the announced text. */
.xlsx-status-bar :deep(.office-status-bar__start) > span + span::before {
  content: "· ";
}

.xlsx-status-bar__address,
.xlsx-status-bar__stat,
.xlsx-status-bar__note {
  flex: 0 1 auto;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.xlsx-status-bar__address {
  max-inline-size: 40%;
}

.xlsx-status-bar__address--empty {
  color: var(--office-muted-foreground);
}

.xlsx-status-bar__note {
  font-weight: 600;
}

.xlsx-status-bar__truncated {
  cursor: help;
}

@container xlsx-viewer (max-width: 620px) {
  .xlsx-status-bar :deep(.office-status-bar__start) {
    justify-content: flex-start;
    overflow-x: auto;
    scrollbar-width: none;
  }
}
</style>
