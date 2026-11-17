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
        role="menuitem"
        @click="onUndo"
      >{{ m.xlsx.contextMenu.undo }}</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        role="menuitem"
        @click="onRedo"
      >{{ m.xlsx.contextMenu.redo }}</button>
    </div>
  </Teleport>
  <XlsxClipboardNotice :message="clipboardNotice" />
</template>

<script setup lang="ts">
import { ref, computed, nextTick, onUnmounted, watch, type CSSProperties } from "vue";
import { useOfficeMenuKeyboard } from "@arcships/vue-ui";
import type { XlsxCellAddress, XlsxSheetData, XlsxViewerController } from "@arcships/xlsx-core";
import XlsxClipboardNotice from "./XlsxClipboardNotice.vue";
import { useClipboardCopyNotice } from "../composables/useClipboardCopyNotice";
import { frozenExtent } from "./grid-metrics";
import { paneAxisContentOffset } from "./grid-pane-geometry";
import { useXlsxMessages } from "../i18n";

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

const visible = ref(false);
const menuRef = ref<HTMLElement | null>(null);
const m = useXlsxMessages();
const position = ref({ x: 0, y: 0 });
const sortTarget = ref<{ columnIndex: number; tableName: string } | null>(null);

const isReadOnly = computed(() => props.controller.isEditingLocked ?? props.controller.readOnly);
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

function close() { visible.value = false; sortTarget.value = null; }
function onCopy() { void copySelection(); close(); }
function onPaste() { props.controller.pasteFromClipboard(); close(); }
function onClear() { props.controller.clearSelectedCells(); close(); }
function onUndo() { props.controller.undo(); close(); }
function onRedo() { props.controller.redo(); close(); }
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

defineExpose({ openAt, close });
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
.xlsx-contextmenu__item:hover { background: #f4f4f5; }
.xlsx-contextmenu__separator { background: #e4e4e7; height: 1px; margin: 4px 0; }
@media (pointer: coarse) {
  .xlsx-contextmenu__item { height: 44px; }
}
</style>
