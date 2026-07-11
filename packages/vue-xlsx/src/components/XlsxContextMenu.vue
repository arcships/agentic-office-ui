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
      class="xlsx-contextmenu"
      data-testid="xlsx-context-menu"
      :style="menuStyle"
      role="menu"
    >
      <button class="xlsx-contextmenu__item" role="menuitem" @click="onCopy">复制</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        role="menuitem"
        @click="onPaste"
      >粘贴</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        role="menuitem"
        @click="onClear"
      >清除内容</button>
      <template v-if="sortTarget && !isReadOnly">
        <div class="xlsx-contextmenu__separator" />
        <button
          class="xlsx-contextmenu__item"
          data-testid="xlsx-sort-ascending"
          role="menuitem"
          @click="onSort('ascending')"
        >按此列升序</button>
        <button
          class="xlsx-contextmenu__item"
          data-testid="xlsx-sort-descending"
          role="menuitem"
          @click="onSort('descending')"
        >按此列降序</button>
      </template>
      <div class="xlsx-contextmenu__separator" />
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        role="menuitem"
        @click="onUndo"
      >撤销</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        role="menuitem"
        @click="onRedo"
      >重做</button>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { ref, computed, onUnmounted, watch, type CSSProperties } from "vue";
import type { XlsxCellAddress, XlsxSheetData, XlsxViewerController } from "@arcships/xlsx-core";

const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const HEADER_HEIGHT = 24;
const ROW_HEADER_WIDTH = 48;
const MIN_DISPLAY_ROWS = 200;
const MIN_DISPLAY_COLS = 50;

const props = defineProps<{
  controller: XlsxViewerController;
  targetElement?: HTMLElement | null;
}>();

const visible = ref(false);
const position = ref({ x: 0, y: 0 });
const sortTarget = ref<{ columnIndex: number; tableName: string } | null>(null);

const isReadOnly = computed(() => props.controller.readOnly);

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
  const x = clientX - rect.left + target.scrollLeft - ROW_HEADER_WIDTH;
  const y = clientY - rect.top + target.scrollTop - HEADER_HEIGHT;
  if (x < 0 || y < 0) return null;

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
  const target = props.targetElement;
  const sheet = props.controller.activeSheet;
  if (!target || !sheet) return;
  const eventTarget = event.target;
  if (
    typeof HTMLElement !== "undefined" &&
    eventTarget instanceof HTMLElement &&
    eventTarget.closest("input, textarea, [contenteditable='true']")
  ) return;

  const cell = resolveGridCell(target, sheet, props.controller.zoomScale, event.clientX, event.clientY);
  if (!cell) return;
  event.preventDefault();
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
  position.value = { x: event.clientX, y: event.clientY };
  visible.value = true;
}

function close() { visible.value = false; sortTarget.value = null; }
function onCopy() { props.controller.copySelectionToClipboard(); close(); }
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
</style>
