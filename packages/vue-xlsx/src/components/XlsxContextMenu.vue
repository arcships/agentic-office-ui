<template>
  <Teleport to="body">
    <div
      v-if="visible"
      class="xlsx-contextmenu-backdrop"
      @pointerdown="close"
      @contextmenu.prevent="close"
    />
    <div
      v-if="visible"
      ref="menuRef"
      class="xlsx-contextmenu"
      data-testid="xlsx-context-menu"
      :style="menuStyle"
      role="menu"
      :aria-label="m.xlsx.contextMenu.label"
      @keydown="onMenuKeydown"
    >
      <template v-if="chartTarget">
        <button
          class="xlsx-contextmenu__item xlsx-contextmenu__item--submenu"
          data-testid="xlsx-context-chart-type"
          role="menuitem"
          aria-haspopup="menu"
          :aria-expanded="chartTypeMenuOpen"
          :disabled="isReadOnly"
          @click="toggleChartTypeMenu"
          @mouseenter="openChartTypeMenu"
        >
          <span>{{ m.xlsx.contextMenu.changeChartType }}</span>
          <span class="xlsx-contextmenu__chevron" aria-hidden="true">›</span>
        </button>
        <div
          v-if="chartTypeMenuOpen"
          class="xlsx-contextmenu__submenu"
          data-testid="xlsx-chart-type-menu"
          role="menu"
          :aria-label="m.xlsx.contextMenu.changeChartType"
        >
          <div v-for="family in chartFamilies" :key="family.family" class="xlsx-contextmenu__group">
            <button
              class="xlsx-contextmenu__item xlsx-contextmenu__item--submenu"
              role="menuitem"
              aria-haspopup="menu"
              :aria-expanded="openChartFamily === family.family"
              :data-testid="`xlsx-chart-family-${family.family}`"
              @click="openChartFamilyMenu(family.family)"
              @mouseenter="openChartFamilyMenu(family.family)"
            >
              <span>{{ family.family }}</span>
              <span class="xlsx-contextmenu__chevron" aria-hidden="true">›</span>
            </button>
            <div
              v-if="openChartFamily === family.family"
              class="xlsx-contextmenu__submenu xlsx-contextmenu__submenu--nested"
              role="menu"
              :aria-label="family.family"
            >
              <button
                v-for="option in family.options"
                :key="option.id"
                class="xlsx-contextmenu__item"
                role="menuitem"
                :data-testid="`xlsx-chart-type-${option.id}`"
                :disabled="isChartTypeUnavailable(option)"
                @click="onChangeChartType(option.id)"
              >
                <span class="xlsx-contextmenu__check" aria-hidden="true">{{ isCurrentChartType(option) ? "✓" : "" }}</span>
                <span>{{ option.label }}</span>
              </button>
            </div>
          </div>
        </div>
        <div class="xlsx-contextmenu__separator" role="separator" />
        <button class="xlsx-contextmenu__item" role="menuitem" @click="onUndo">{{ m.xlsx.contextMenu.undo }}</button>
        <button class="xlsx-contextmenu__item" role="menuitem" @click="onRedo">{{ m.xlsx.contextMenu.redo }}</button>
      </template>
      <template v-else>
      <button class="xlsx-contextmenu__item" role="menuitem" @click="onCopy">{{ m.xlsx.contextMenu.copy }}</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        role="menuitem"
        @click="onPaste"
      >{{ m.xlsx.contextMenu.paste }}</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        role="menuitem"
        @click="onClear"
      >{{ m.xlsx.contextMenu.clear }}</button>
      <template v-if="sortTarget && !isReadOnly">
        <div class="xlsx-contextmenu__separator" role="separator" />
        <button
          class="xlsx-contextmenu__item"
          data-testid="xlsx-sort-ascending"
          role="menuitem"
          @click="onSort('ascending')"
        >{{ m.xlsx.contextMenu.sortAscending }}</button>
        <button
          class="xlsx-contextmenu__item"
          data-testid="xlsx-sort-descending"
          role="menuitem"
          @click="onSort('descending')"
        >{{ m.xlsx.contextMenu.sortDescending }}</button>
      </template>
      <div class="xlsx-contextmenu__separator" role="separator" />
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-comment"
        role="menuitem"
        @click="onAuthoring('comment')"
      >{{ commentAtCell ? m.xlsx.authoring.context.editComment : m.xlsx.authoring.context.insertComment }}</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-hyperlink"
        role="menuitem"
        @click="onAuthoring('hyperlink')"
      >{{ hyperlinkAtCell ? m.xlsx.authoring.context.editHyperlink : m.xlsx.authoring.context.insertHyperlink }}</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-chart"
        role="menuitem"
        @click="onAuthoring('chart')"
      >{{ m.xlsx.authoring.context.insertChart }}</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-conditional"
        role="menuitem"
        @click="onAuthoring('conditional')"
      >{{ m.xlsx.authoring.context.conditional }}</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-validation"
        role="menuitem"
        @click="onAuthoring('validation')"
      >{{ m.xlsx.authoring.context.validation }}</button>
      <button
        v-if="!isReadOnly && !tableAtCell"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-table"
        role="menuitem"
        @click="onAuthoring('table')"
      >{{ m.xlsx.authoring.context.createTable }}</button>
      <button
        v-if="!isReadOnly && tableAtCell"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-table-remove"
        role="menuitem"
        @click="onRemoveTable"
      >{{ m.xlsx.authoring.context.removeTable }}</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-freeze"
        role="menuitem"
        @click="onToggleFreeze"
      >{{ isFrozen ? m.xlsx.authoring.context.unfreeze : m.xlsx.authoring.context.freezeHere }}</button>
      <div class="xlsx-contextmenu__separator" role="separator" />
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        role="menuitem"
        @click="onUndo"
      >{{ m.xlsx.contextMenu.undo }}</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        role="menuitem"
        @click="onRedo"
      >{{ m.xlsx.contextMenu.redo }}</button>
      </template>
    </div>
  </Teleport>
  <XlsxClipboardNotice :message="clipboardNotice" />
</template>

<script setup lang="ts">
import { ref, computed, nextTick, onUnmounted, watch, type CSSProperties } from "vue";
import { useOfficeMenuKeyboard } from "@arcships/vue-ui";
import type { ChartTypeOption, XlsxCellAddress, XlsxSheetData, XlsxViewerController } from "@arcships/xlsx-core";
import { chartTypeFamilies, chartTypeUnavailable, matchesChartType } from "@arcships/xlsx-core";
import XlsxClipboardNotice from "./XlsxClipboardNotice.vue";
import { useClipboardCopyNotice } from "../composables/useClipboardCopyNotice";
import { frozenExtent } from "./grid-metrics";
import { paneAxisContentOffset } from "./grid-pane-geometry";
import { useXlsxMessages } from "../i18n";
import type { XlsxAuthoringDialogKind } from "./authoring-dialog";

const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const HEADER_HEIGHT = 24;
const ROW_HEADER_WIDTH = 48;
const MIN_DISPLAY_ROWS = 200;
const MIN_DISPLAY_COLS = 50;

const props = defineProps<{
  controller: XlsxViewerController;
  targetElement?: HTMLElement | null;
  /** Native contextmenu is ignored right after a synthetic long press. */
  suppressContextMenu?: ((clientX: number, clientY: number) => boolean) | null;
}>();

const emit = defineEmits<{
  authoring: [kind: XlsxAuthoringDialogKind];
}>();

const visible = ref(false);
const menuRef = ref<HTMLElement | null>(null);
const m = useXlsxMessages();
const position = ref({ x: 0, y: 0 });
const sortTarget = ref<{ columnIndex: number; tableName: string } | null>(null);
/** Chart the menu was opened on, or `null` for the cell menu. */
const chartTargetId = ref<string | null>(null);
const chartTypeMenuOpen = ref(false);
const openChartFamily = ref<string | null>(null);
const chartFamilies = chartTypeFamilies();
const chartTarget = computed(() => {
  props.controller.revision;
  return chartTargetId.value ? props.controller.getChartById?.(chartTargetId.value) ?? null : null;
});

const isReadOnly = computed(() => props.controller.isEditingLocked ?? props.controller.readOnly);
const commentAtCell = computed(() => {
  props.controller.revision;
  return Boolean(props.controller.getCellComment?.());
});
const hyperlinkAtCell = computed(() => {
  props.controller.revision;
  return Boolean(props.controller.getCellHyperlink?.());
});
const tableAtCell = computed(() => {
  props.controller.revision;
  return props.controller.getTableAtCell?.() ?? null;
});
const isFrozen = computed(() => Boolean(props.controller.freezePanes));
const { notice: clipboardNotice, copySelection } = useClipboardCopyNotice(() => props.controller, {
  failureMessage: () => m.value.xlsx.clipboard.copyFailed,
});

// Escape/Tab and arrow navigation come from the shared Office menu keyboard, so
// the cell menu closes exactly like the DOCX context menu and returns focus to
// the surface that opened it.
const menuKeyboard = useOfficeMenuKeyboard({
  menu: () => menuRef.value,
  close: () => close(),
  restoreTarget: () => props.targetElement ?? null,
});

function onMenuKeydown(event: KeyboardEvent): void {
  menuKeyboard.onKeydown(event);
}

const menuStyle = computed<CSSProperties>(() => ({
  left: `${position.value.x}px`,
  position: "fixed",
  top: `${position.value.y}px`,
  zIndex: "10000",
}));

function resolveDisplayIndices(indices: number[] | undefined, count: number): number[] {
  if (!indices?.length) return Array.from({ length: count }, (_, index) => index);
  if (indices.length >= count) return indices.slice(0, count);
  const resolved = [...indices];
  let nextIndex = Math.max(-1, ...resolved) + 1;
  while (resolved.length < count) resolved.push(nextIndex++);
  return resolved;
}

function findDisplayIndexAtOffset(sizes: number[], offset: number): number {
  if (offset < 0) return -1;
  let start = 0;
  for (let index = 0; index < sizes.length; index += 1) {
    const end = start + Math.max(0, sizes[index] || 0);
    if (offset < end) return index;
    start = end;
  }
  return -1;
}

function resolveGridCell(
  target: HTMLElement,
  sheet: XlsxSheetData,
  zoomScale: number,
  clientX: number,
  clientY: number,
): XlsxCellAddress | null {
  const rect = target.getBoundingClientRect();
  const localX = clientX - rect.left - ROW_HEADER_WIDTH;
  const localY = clientY - rect.top - HEADER_HEIGHT;
  if (localX < 0 || localY < 0) return null;
  // A frozen row/column keeps content coordinates: only the scrolling pane
  // adds the scroll offset, otherwise a right-click on a pinned cell would
  // select whatever scrolled underneath it.
  const x = paneAxisContentOffset(localX, target.scrollLeft, frozenExtent(sheet, "column", zoomScale));
  const y = paneAxisContentOffset(localY, target.scrollTop, frozenExtent(sheet, "row", zoomScale));

  const zoom = Math.max(0.01, zoomScale / 100);
  const colCount = Math.max(sheet.colWidths.length, MIN_DISPLAY_COLS);
  const rowCount = Math.max(sheet.rowHeights.length, MIN_DISPLAY_ROWS);
  const colWidths = Array.from({ length: colCount }, (_, index) =>
    (sheet.colWidths[index] || sheet.defaultColWidthPx || DEFAULT_COL_WIDTH) * zoom,
  );
  const rowHeights = Array.from({ length: rowCount }, (_, index) =>
    (sheet.rowHeights[index] || sheet.defaultRowHeightPx || DEFAULT_ROW_HEIGHT) * zoom,
  );
  const displayCol = findDisplayIndexAtOffset(colWidths, x);
  const displayRow = findDisplayIndexAtOffset(rowHeights, y);
  if (displayCol < 0 || displayRow < 0) return null;

  const col = resolveDisplayIndices(sheet.visibleCols, colCount)[displayCol];
  const row = resolveDisplayIndices(sheet.visibleRows, rowCount)[displayRow];
  return col === undefined || row === undefined ? null : { col, row };
}

function onContextMenu(event: MouseEvent) {
  if (props.suppressContextMenu?.(event.clientX, event.clientY)) {
    event.preventDefault();
    return;
  }
  if (openAt(event.clientX, event.clientY, event.target)) event.preventDefault();
}

function openAt(clientX: number, clientY: number, eventTarget?: unknown): boolean {
  const target = props.targetElement;
  const sheet = props.controller.activeSheet;
  if (!target || !sheet) return false;
  if (
    typeof HTMLElement !== "undefined" &&
    eventTarget instanceof HTMLElement &&
    eventTarget.closest("input, textarea, [contenteditable='true']")
  ) return false;

  const cell = resolveGridCell(target, sheet, props.controller.zoomScale, clientX, clientY);
  if (!cell) return false;
  chartTargetId.value = null;
  closeChartTypeMenu();
  props.controller.selectCell(cell);
  const table = props.controller.tables.find((candidate) => {
    const headerRows = Math.max(1, candidate.headerRowCount);
    return cell.row >= candidate.start.row &&
      cell.row < candidate.start.row + headerRows &&
      cell.col >= candidate.start.col &&
      cell.col <= candidate.end.col;
  });
  sortTarget.value = table
    ? { columnIndex: cell.col - table.start.col, tableName: table.name }
    : null;
  position.value = { x: clientX, y: clientY };
  menuKeyboard.captureInvoker();
  visible.value = true;
  void nextTick(() => menuKeyboard.focusFirst());
  return true;
}

/**
 * Opens the chart menu for `chartId`: the chart's own actions replace the cell
 * actions, starting with the family/type submenu of "Change Chart Type".
 */
function openForChart(chartId: string, clientX: number, clientY: number): boolean {
  const chart = props.controller.getChartById?.(chartId) ?? null;
  if (!chart) return false;
  chartTargetId.value = chartId;
  sortTarget.value = null;
  closeChartTypeMenu();
  position.value = { x: clientX, y: clientY };
  menuKeyboard.captureInvoker();
  visible.value = true;
  void nextTick(() => menuKeyboard.focusFirst());
  return true;
}

function closeChartTypeMenu() {
  chartTypeMenuOpen.value = false;
  openChartFamily.value = null;
}

function toggleChartTypeMenu() {
  if (chartTypeMenuOpen.value) closeChartTypeMenu();
  else chartTypeMenuOpen.value = true;
}

function openChartTypeMenu() {
  chartTypeMenuOpen.value = true;
}

function openChartFamilyMenu(family: string) {
  openChartFamily.value = family;
}

function isCurrentChartType(option: ChartTypeOption): boolean {
  const chart = chartTarget.value;
  return chart ? matchesChartType(chart, option) : false;
}

function isChartTypeUnavailable(option: ChartTypeOption): boolean {
  const chart = chartTarget.value;
  return Boolean(isReadOnly.value || !chart || chartTypeUnavailable(chart, option));
}

function onChangeChartType(type: string) {
  const chart = chartTarget.value;
  if (!chart || isReadOnly.value) return;
  props.controller.changeChartType?.(chart.id, type);
  close();
}

function close() {
  visible.value = false;
  sortTarget.value = null;
  chartTargetId.value = null;
  closeChartTypeMenu();
}
function onCopy() { void copySelection(); close(); }
function onPaste() { props.controller.pasteFromClipboard(); close(); }
function onClear() { props.controller.clearSelectedCells(); close(); }
function onUndo() { props.controller.undo(); close(); }
function onRedo() { props.controller.redo(); close(); }
function onAuthoring(kind: XlsxAuthoringDialogKind) {
  close();
  emit("authoring", kind);
}

function onRemoveTable() {
  const table = tableAtCell.value;
  if (!table || isReadOnly.value) return;
  props.controller.removeTable?.(table.name);
  close();
}

function onToggleFreeze() {
  if (isReadOnly.value) return;
  props.controller.toggleFreezePanesAtActiveCell?.();
  close();
}

function onSort(direction: "ascending" | "descending") {
  const target = sortTarget.value;
  if (!target || isReadOnly.value) return;
  props.controller.sortTable(target.tableName, target.columnIndex, direction);
  close();
}

let boundTarget: HTMLElement | null = null;
const boundHandler = (event: MouseEvent) => onContextMenu(event);

watch(
  () => props.targetElement,
  (target) => {
    boundTarget?.removeEventListener("contextmenu", boundHandler);
    boundTarget = target ?? null;
    boundTarget?.addEventListener("contextmenu", boundHandler);
  },
  { immediate: true, flush: "post" },
);

onUnmounted(() => {
  boundTarget?.removeEventListener("contextmenu", boundHandler);
  boundTarget = null;
});

defineExpose({ openAt, close, openForChart });
</script>

<style scoped>
.xlsx-contextmenu-backdrop { bottom: 0; left: 0; position: fixed; right: 0; top: 0; z-index: 9999; }
.xlsx-contextmenu {
  background: #ffffff; border: 1px solid #d4d4d8; border-radius: 6px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.12); min-width: 160px; padding: 4px;
}
.xlsx-contextmenu__item {
  align-items: center; background: transparent; border: none; border-radius: 4px;
  color: #18181b; cursor: pointer; display: flex; font-family: inherit;
  font-size: 12px; height: 28px; padding: 0 8px; width: 100%;
}
.xlsx-contextmenu__item:hover:not(:disabled) { background: #f4f4f5; }
.xlsx-contextmenu__item:disabled { cursor: default; opacity: 0.45; }
.xlsx-contextmenu__item--submenu { justify-content: space-between; }
.xlsx-contextmenu__chevron { color: #71717a; }
.xlsx-contextmenu__check { display: inline-block; min-width: 12px; }
.xlsx-contextmenu__submenu {
  background: #ffffff; border: 1px solid #d4d4d8; border-radius: 6px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.12); left: calc(100% - 4px);
  max-height: 70vh; min-width: 190px; overflow-y: auto; padding: 4px;
  position: absolute; top: 0;
}
.xlsx-contextmenu__submenu--nested { z-index: 1; }
.xlsx-contextmenu__group { position: relative; }
.xlsx-contextmenu__separator { background: #e4e4e7; height: 1px; margin: 4px 0; }
@media (pointer: coarse) {
  .xlsx-contextmenu__item { height: 44px; }
}
</style>
