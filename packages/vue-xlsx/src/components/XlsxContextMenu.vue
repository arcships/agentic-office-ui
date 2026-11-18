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
      :data-menu-kind="menuKind"
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
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-drawing-copy" role="menuitem" @click="onDrawingCopy">
          <span>{{ m.xlsx.contextMenu.copy }}</span><kbd class="xlsx-contextmenu__shortcut">{{ SHORTCUT_COPY }}</kbd>
        </button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-drawing-cut" role="menuitem" :disabled="isReadOnly" @click="onDrawingCut">
          <span>{{ m.xlsx.contextMenu.cut }}</span><kbd class="xlsx-contextmenu__shortcut">{{ SHORTCUT_CUT }}</kbd>
        </button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-drawing-paste" role="menuitem" :disabled="isReadOnly" @click="onDrawingPaste">
          <span>{{ m.xlsx.contextMenu.paste }}</span><kbd class="xlsx-contextmenu__shortcut">{{ SHORTCUT_PASTE }}</kbd>
        </button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-drawing-duplicate" role="menuitem" :disabled="isReadOnly" @click="onDrawingDuplicate">{{ m.xlsx.contextMenu.duplicate }}</button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-drawing-delete" role="menuitem" :disabled="isReadOnly" @click="onDrawingDelete">{{ m.xlsx.contextMenu.delete }}</button>
        <div class="xlsx-contextmenu__separator" role="separator" />
        <button class="xlsx-contextmenu__item" role="menuitem" @click="onUndo">{{ m.xlsx.contextMenu.undo }}</button>
        <button class="xlsx-contextmenu__item" role="menuitem" @click="onRedo">{{ m.xlsx.contextMenu.redo }}</button>
      </template>
      <template v-else-if="drawingTarget">
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-drawing-copy" role="menuitem" @click="onDrawingCopy">
          <span>{{ m.xlsx.contextMenu.copy }}</span><kbd class="xlsx-contextmenu__shortcut">{{ SHORTCUT_COPY }}</kbd>
        </button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-drawing-cut" role="menuitem" :disabled="isReadOnly" @click="onDrawingCut">
          <span>{{ m.xlsx.contextMenu.cut }}</span><kbd class="xlsx-contextmenu__shortcut">{{ SHORTCUT_CUT }}</kbd>
        </button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-drawing-paste" role="menuitem" :disabled="isReadOnly" @click="onDrawingPaste">
          <span>{{ m.xlsx.contextMenu.paste }}</span><kbd class="xlsx-contextmenu__shortcut">{{ SHORTCUT_PASTE }}</kbd>
        </button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-drawing-duplicate" role="menuitem" :disabled="isReadOnly" @click="onDrawingDuplicate">{{ m.xlsx.contextMenu.duplicate }}</button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-drawing-delete" role="menuitem" :disabled="isReadOnly" @click="onDrawingDelete">{{ m.xlsx.contextMenu.delete }}</button>
        <div class="xlsx-contextmenu__separator" role="separator" />
        <button class="xlsx-contextmenu__item" role="menuitem" @click="onUndo">{{ m.xlsx.contextMenu.undo }}</button>
        <button class="xlsx-contextmenu__item" role="menuitem" @click="onRedo">{{ m.xlsx.contextMenu.redo }}</button>
      </template>
      <template v-else>
      <button class="xlsx-contextmenu__item" data-testid="xlsx-context-cut" role="menuitem" :disabled="isReadOnly" @click="onCut">
        <span>{{ m.xlsx.contextMenu.cut }}</span><kbd class="xlsx-contextmenu__shortcut">{{ SHORTCUT_CUT }}</kbd>
      </button>
      <button class="xlsx-contextmenu__item" data-testid="xlsx-context-copy" role="menuitem" @click="onCopy">
        <span>{{ m.xlsx.contextMenu.copy }}</span><kbd class="xlsx-contextmenu__shortcut">{{ SHORTCUT_COPY }}</kbd>
      </button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-paste"
        role="menuitem"
        @click="onPaste"
      >
        <span>{{ m.xlsx.contextMenu.paste }}</span><kbd class="xlsx-contextmenu__shortcut">{{ SHORTCUT_PASTE }}</kbd>
      </button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item xlsx-contextmenu__item--submenu"
        data-testid="xlsx-context-paste-special"
        data-submenu="pasteSpecial"
        role="menuitem"
        aria-haspopup="menu"
        :aria-expanded="submenu === 'pasteSpecial'"
        @click="togglePasteSpecial"
      >
        <span>{{ m.xlsx.contextMenu.pasteSpecial }}</span><span class="xlsx-contextmenu__chevron">▸</span>
      </button>
      <div
        v-if="submenu === 'pasteSpecial'"
        class="xlsx-contextmenu__submenu"
        data-testid="xlsx-context-paste-special-menu"
        role="menu"
      >
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-paste-all" role="menuitem" @click="onPasteSpecial({ mode: 'all' })">{{ m.xlsx.contextMenu.pasteAll }}</button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-paste-values" role="menuitem" @click="onPasteSpecial({ mode: 'values' })">{{ m.xlsx.contextMenu.pasteValues }}</button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-paste-formats" role="menuitem" @click="onPasteSpecial({ mode: 'formats' })">{{ m.xlsx.contextMenu.pasteFormats }}</button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-paste-transpose" role="menuitem" @click="onPasteSpecial({ mode: 'all', transpose: true })">{{ m.xlsx.contextMenu.pasteTranspose }}</button>
      </div>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-edit-cell"
        role="menuitem"
        @click="onEditCell"
      >
        <span>{{ m.xlsx.contextMenu.editCell }}</span><kbd class="xlsx-contextmenu__shortcut">F2</kbd>
      </button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-clear"
        role="menuitem"
        @click="onClear"
      >
        <span>{{ m.xlsx.contextMenu.clear }}</span><kbd class="xlsx-contextmenu__shortcut">Del</kbd>
      </button>
      <button
        v-if="!isReadOnly && menuKind === 'cell'"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-number-format"
        role="menuitem"
        @click="onAuthoring('numberFormat')"
      >{{ m.xlsx.authoring.numberFormat.title }}</button>
      <button
        v-if="!isReadOnly && menuKind === 'cell'"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-names"
        role="menuitem"
        @click="onAuthoring('names')"
      >{{ m.xlsx.authoring.names.title }}</button>
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
      <template v-if="!isReadOnly">
        <div class="xlsx-contextmenu__separator" role="separator" />
        <button
          class="xlsx-contextmenu__item"
          data-testid="xlsx-context-insert-row-above"
          role="menuitem"
          @click="onInsertRow('above')"
        >{{ m.xlsx.contextMenu.insertRowAbove }}</button>
        <button
          class="xlsx-contextmenu__item"
          data-testid="xlsx-context-insert-row-below"
          role="menuitem"
          @click="onInsertRow('below')"
        >{{ m.xlsx.contextMenu.insertRowBelow }}</button>
        <button
          class="xlsx-contextmenu__item"
          data-testid="xlsx-context-delete-row"
          role="menuitem"
          @click="onDeleteRow"
        >{{ m.xlsx.contextMenu.deleteRow }}</button>
        <button
          class="xlsx-contextmenu__item"
          data-testid="xlsx-context-insert-column-left"
          role="menuitem"
          @click="onInsertColumn('left')"
        >{{ m.xlsx.contextMenu.insertColumnLeft }}</button>
        <button
          class="xlsx-contextmenu__item"
          data-testid="xlsx-context-insert-column-right"
          role="menuitem"
          @click="onInsertColumn('right')"
        >{{ m.xlsx.contextMenu.insertColumnRight }}</button>
        <button
          class="xlsx-contextmenu__item"
          data-testid="xlsx-context-delete-column"
          role="menuitem"
          @click="onDeleteColumn"
        >{{ m.xlsx.contextMenu.deleteColumn }}</button>
      </template>
      <template v-if="menuKind === 'cell' && !isReadOnly">
        <div class="xlsx-contextmenu__separator" role="separator" />
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-merge-center" role="menuitem" @click="onMergeCenter">{{ m.xlsx.contextMenu.mergeAndCenter }}</button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-merge-across" role="menuitem" @click="onMergeAcross">{{ m.xlsx.contextMenu.mergeAcross }}</button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-merge" role="menuitem" @click="onMergeCells">{{ m.xlsx.contextMenu.mergeCells }}</button>
        <button class="xlsx-contextmenu__item" data-testid="xlsx-context-unmerge" role="menuitem" @click="onUnmerge">{{ m.xlsx.contextMenu.unmerge }}</button>
      </template>
      <div class="xlsx-contextmenu__separator" role="separator" />
      <button
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-select-row"
        role="menuitemradio"
        :aria-checked="selectionIsRow"
        @click="onSelectRow"
      >
        <span>{{ m.xlsx.contextMenu.selectRow }}</span><span v-if="selectionIsRow" class="xlsx-contextmenu__checked">✓</span>
      </button>
      <button
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-select-column"
        role="menuitemradio"
        :aria-checked="selectionIsColumn"
        @click="onSelectColumn"
      >
        <span>{{ m.xlsx.contextMenu.selectColumn }}</span><span v-if="selectionIsColumn" class="xlsx-contextmenu__checked">✓</span>
      </button>
      <button
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-select-all"
        role="menuitemradio"
        :aria-checked="selectionIsAll"
        @click="onSelectAll"
      >
        <span>{{ m.xlsx.contextMenu.selectAll }}</span><span v-if="selectionIsAll" class="xlsx-contextmenu__checked">✓</span>
      </button>
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
        data-testid="xlsx-context-undo"
        role="menuitem"
        @click="onUndo"
      >
        <span>{{ m.xlsx.contextMenu.undo }}</span><kbd class="xlsx-contextmenu__shortcut">Ctrl+Z</kbd>
      </button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        data-testid="xlsx-context-redo"
        role="menuitem"
        @click="onRedo"
      >
        <span>{{ m.xlsx.contextMenu.redo }}</span><kbd class="xlsx-contextmenu__shortcut">Ctrl+Y</kbd>
      </button>
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
const SHORTCUT_COPY = "Ctrl+C";
const SHORTCUT_CUT = "Ctrl+X";
const SHORTCUT_PASTE = "Ctrl+V";
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
/** Image or shape the menu was opened on, or `null`. */
const drawingTargetId = ref<string | null>(null);
const drawingTargetKind = ref<"image" | "shape" | null>(null);
const chartTypeMenuOpen = ref(false);
const openChartFamily = ref<string | null>(null);
const chartFamilies = chartTypeFamilies();
const chartTarget = computed(() => {
  props.controller.revision;
  return chartTargetId.value ? props.controller.getChartById?.(chartTargetId.value) ?? null : null;
});
const drawingTarget = computed(() => {
  props.controller.revision;
  if (!drawingTargetId.value) return null;
  return drawingTargetKind.value === "image"
    ? props.controller.getImageById?.(drawingTargetId.value) ?? null
    : props.controller.getShapeById?.(drawingTargetId.value) ?? null;
});

const menuKind = ref<"cell" | "row" | "column" | "corner">("cell");
const headerIndex = ref(0);
const submenu = ref<"pasteSpecial" | null>(null);
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
const selectionIsRow = computed(() => {
  const selection = props.controller.selection;
  const sheet = props.controller.activeSheet;
  if (!selection || !sheet) return false;
  return selection.start.row === selection.end.row
    && selection.start.col === 0
    && selection.end.col >= Math.max(sheet.maxUsedCol, sheet.visibleCols.at(-1) ?? 0);
});
const selectionIsColumn = computed(() => {
  const selection = props.controller.selection;
  const sheet = props.controller.activeSheet;
  if (!selection || !sheet) return false;
  return selection.start.col === selection.end.col
    && selection.start.row === 0
    && selection.end.row >= Math.max(sheet.maxUsedRow, sheet.visibleRows.at(-1) ?? 0);
});
const selectionIsAll = computed(() => {
  const selection = props.controller.selection;
  const sheet = props.controller.activeSheet;
  if (!selection || !sheet) return false;
  return selection.start.row === 0
    && selection.start.col === 0
    && selection.end.row >= Math.max(sheet.maxUsedRow, sheet.visibleRows.at(-1) ?? 0)
    && selection.end.col >= Math.max(sheet.maxUsedCol, sheet.visibleCols.at(-1) ?? 0);
});
const { notice: clipboardNotice, copySelection } = useClipboardCopyNotice(() => props.controller, {
  failureMessage: () => m.value.xlsx.clipboard.copyFailed,
});

// Escape/Tab and arrow navigation come from the shared Office menu keyboard, so
// the cell menu closes exactly like the DOCX context menu and returns focus to
// the surface that opened it.
const menuKeyboard = useOfficeMenuKeyboard({
  menu: () => menuRef.value,
  close: () => close(),
  itemSelector: '[role="menuitem"]:not([disabled]):not([aria-disabled="true"]), [role="menuitemradio"]:not([disabled]):not([aria-disabled="true"])',
  restoreTarget: () => props.targetElement ?? null,
});

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

function resolveAxisSizes(sheet: XlsxSheetData, zoomScale: number) {
  const zoom = Math.max(0.01, zoomScale / 100);
  const colCount = Math.max(sheet.colWidths.length, MIN_DISPLAY_COLS);
  const rowCount = Math.max(sheet.rowHeights.length, MIN_DISPLAY_ROWS);
  return {
    colCount,
    colWidths: Array.from({ length: colCount }, (_, index) =>
      (sheet.colWidths[index] || sheet.defaultColWidthPx || DEFAULT_COL_WIDTH) * zoom,
    ),
    rowCount,
    rowHeights: Array.from({ length: rowCount }, (_, index) =>
      (sheet.rowHeights[index] || sheet.defaultRowHeightPx || DEFAULT_ROW_HEIGHT) * zoom,
    ),
  };
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
  // select whatever scrolled underneath it. RTL mirrors the column axis, so
  // the pointer offset is measured from the frozen right-hand strip.
  const mirroredX = sheet.rightToLeft === true
    ? Math.max(0, rect.width - ROW_HEADER_WIDTH) - localX
    : localX;
  const x = paneAxisContentOffset(mirroredX, target.scrollLeft, frozenExtent(sheet, "column", zoomScale));
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

/** Resolves a right-click on the header strips into a row/column/corner target. */
function resolveHeaderHit(
  target: HTMLElement,
  sheet: XlsxSheetData,
  zoomScale: number,
  clientX: number,
  clientY: number,
): { index: number; kind: "row" | "column" | "corner" } | null {
  const rect = target.getBoundingClientRect();
  const localX = clientX - rect.left;
  const localY = clientY - rect.top;
  if (localX < 0 || localY < 0) return null;
  if (localX <= ROW_HEADER_WIDTH && localY <= HEADER_HEIGHT) {
    return { index: 0, kind: "corner" };
  }
  const { colCount, colWidths, rowCount, rowHeights } = resolveAxisSizes(sheet, zoomScale);
  if (localY <= HEADER_HEIGHT) {
    const mirroredX = sheet.rightToLeft === true
      ? Math.max(0, rect.width - ROW_HEADER_WIDTH) - (localX - ROW_HEADER_WIDTH)
      : localX - ROW_HEADER_WIDTH;
    const x = paneAxisContentOffset(mirroredX, target.scrollLeft, frozenExtent(sheet, "column", zoomScale));
    const displayCol = findDisplayIndexAtOffset(colWidths, x);
    const col = resolveDisplayIndices(sheet.visibleCols, colCount)[displayCol];
    return col === undefined ? null : { index: col, kind: "column" };
  }
  if (localX <= ROW_HEADER_WIDTH) {
    const y = paneAxisContentOffset(localY - HEADER_HEIGHT, target.scrollTop, frozenExtent(sheet, "row", zoomScale));
    const displayRow = findDisplayIndexAtOffset(rowHeights, y);
    const row = resolveDisplayIndices(sheet.visibleRows, rowCount)[displayRow];
    return row === undefined ? null : { index: row, kind: "row" };
  }
  return null;
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
  chartTargetId.value = null;
  drawingTargetId.value = null;
  drawingTargetKind.value = null;
  closeChartTypeMenu();
  if (cell) {
    menuKind.value = "cell";
    headerIndex.value = cell.row;
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
    submenu.value = null;
    menuKeyboard.captureInvoker();
    visible.value = true;
    void nextTick(() => menuKeyboard.focusFirst());
    return true;
  }

  const header = resolveHeaderHit(target, sheet, props.controller.zoomScale, clientX, clientY);
  if (!header) return false;
  menuKind.value = header.kind;
  headerIndex.value = header.index;
  sortTarget.value = null;
  if (header.kind === "row") {
    props.controller.selectRow?.(header.index);
  } else if (header.kind === "column") {
    props.controller.selectColumn?.(header.index);
    const table = props.controller.tables.find((candidate) =>
      header.index >= candidate.start.col && header.index <= candidate.end.col
    );
    sortTarget.value = table
      ? { columnIndex: header.index - table.start.col, tableName: table.name }
      : null;
  } else {
    props.controller.selectAll?.();
  }
  position.value = { x: clientX, y: clientY };
  submenu.value = null;
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
  props.controller.selectChart?.(chartId);
  chartTargetId.value = chartId;
  drawingTargetId.value = null;
  drawingTargetKind.value = null;
  sortTarget.value = null;
  closeChartTypeMenu();
  position.value = { x: clientX, y: clientY };
  menuKeyboard.captureInvoker();
  visible.value = true;
  void nextTick(() => menuKeyboard.focusFirst());
  return true;
}

/** Opens the drawing menu for an image or shape. */
function openForDrawing(
  kind: "image" | "shape",
  drawingId: string,
  clientX: number,
  clientY: number,
): boolean {
  const drawing = kind === "image"
    ? props.controller.getImageById?.(drawingId) ?? null
    : props.controller.getShapeById?.(drawingId) ?? null;
  if (!drawing) return false;
  if (kind === "image") props.controller.selectImage?.(drawingId);
  else props.controller.selectShape?.(drawingId);
  chartTargetId.value = null;
  drawingTargetId.value = drawingId;
  drawingTargetKind.value = kind;
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
  drawingTargetId.value = null;
  drawingTargetKind.value = null;
  closeChartTypeMenu();
  submenu.value = null;
}
function onDrawingCopy() { void copySelection(); close(); }
function onDrawingCut() { props.controller.cutSelection?.(); close(); }
function onDrawingPaste() { void props.controller.pasteFromClipboard?.(); close(); }
function onDrawingDuplicate() { props.controller.duplicateDrawing?.(); close(); }
function onDrawingDelete() { props.controller.deleteSelectedDrawing?.(); close(); }
function onCopy() { void copySelection(); close(); }
function onCut() { props.controller.cutSelection?.(); close(); }
function onPaste() { props.controller.pasteFromClipboard(); close(); }
function onPasteSpecial(options: { mode: "all" | "values" | "formats"; transpose?: boolean }) {
  void props.controller.pasteSpecialFromClipboard?.(options);
  close();
}
function onEditCell() {
  props.controller.editCell?.(props.controller.activeCell);
  close();
}
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

function targetRow() {
  return menuKind.value === "row"
    ? headerIndex.value
    : props.controller.activeCell?.row ?? 0;
}

function targetColumn() {
  return menuKind.value === "column"
    ? headerIndex.value
    : props.controller.activeCell?.col ?? 0;
}

function onInsertRow(position: "above" | "below") {
  if (isReadOnly.value) return;
  props.controller.insertRows?.(targetRow() + (position === "below" ? 1 : 0));
  close();
}

function onDeleteRow() {
  if (isReadOnly.value) return;
  props.controller.deleteRows?.(targetRow());
  close();
}

function onInsertColumn(position: "left" | "right") {
  if (isReadOnly.value) return;
  props.controller.insertColumns?.(targetColumn() + (position === "right" ? 1 : 0));
  close();
}

function onDeleteColumn() {
  if (isReadOnly.value) return;
  props.controller.deleteColumns?.(targetColumn());
  close();
}

function onMergeCenter() { if (isReadOnly.value) return; props.controller.mergeAndCenterSelection?.(); close(); }
function onMergeAcross() { if (isReadOnly.value) return; props.controller.mergeAcrossSelection?.(); close(); }
function onMergeCells() { if (isReadOnly.value) return; props.controller.mergeSelection?.(); close(); }
function onUnmerge() { if (isReadOnly.value) return; props.controller.unmergeSelection?.(); close(); }
function onSelectRow() { props.controller.selectRow?.(targetRow()); close(); }
function onSelectColumn() { props.controller.selectColumn?.(targetColumn()); close(); }
function onSelectAll() { props.controller.selectAll?.(); close(); }

function togglePasteSpecial() {
  submenu.value = submenu.value === "pasteSpecial" ? null : "pasteSpecial";
  if (submenu.value) {
    void nextTick(() => {
      const first = menuRef.value?.querySelector?.("[data-testid='xlsx-context-paste-all']") as HTMLElement | null;
      first?.focus?.();
    });
  }
}

function closeSubmenu() {
  if (!submenu.value) return false;
  submenu.value = null;
  void nextTick(() => {
    const parent = menuRef.value?.querySelector?.("[data-submenu='pasteSpecial']") as HTMLElement | null;
    parent?.focus?.();
  });
  return true;
}

function menuItemElements() {
  if (!menuRef.value || typeof menuRef.value.querySelectorAll !== "function") {
    return [] as HTMLElement[];
  }
  return Array.from(
    menuRef.value.querySelectorAll('[role="menuitem"]:not([disabled]), [role="menuitemradio"]:not([disabled])')
  ) as HTMLElement[];
}

function onMenuKeydown(event: KeyboardEvent): void {
  const focused = typeof document !== "undefined" ? document.activeElement as HTMLElement | null : null;
  if (event.key === "ArrowRight") {
    if (focused?.dataset?.submenu === "pasteSpecial") {
      event.preventDefault();
      event.stopPropagation();
      submenu.value = "pasteSpecial";
      void nextTick(() => {
        const first = menuRef.value?.querySelector?.("[data-testid='xlsx-context-paste-all']") as HTMLElement | null;
        first?.focus?.();
      });
    }
    return;
  }
  if (event.key === "ArrowLeft") {
    event.preventDefault();
    event.stopPropagation();
    if (!closeSubmenu()) {
      close();
    }
    return;
  }
  if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
    // Type-ahead: cycle through the visible items whose label starts with the
    // typed character, matching native menu behavior.
    const items = menuItemElements();
    if (items.length > 0) {
      const currentIndex = focused ? items.indexOf(focused) : -1;
      const key = event.key.toLowerCase();
      for (let offset = 1; offset <= items.length; offset += 1) {
        const candidate = items[(currentIndex + offset + items.length) % items.length];
        const label = (candidate?.textContent ?? "").trim().toLowerCase();
        if (label.startsWith(key)) {
          event.preventDefault();
          event.stopPropagation();
          candidate.focus?.();
          break;
        }
      }
    }
    return;
  }
  menuKeyboard.onKeydown(event);
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

defineExpose({ openAt, close, openForChart, openForDrawing });
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
  font-size: 12px; gap: 16px; height: 28px; justify-content: space-between;
  padding: 0 8px; text-align: left; width: 100%;
}
.xlsx-contextmenu__item:hover:not(:disabled) { background: #f4f4f5; }
.xlsx-contextmenu__item:disabled { cursor: default; opacity: 0.45; }
.xlsx-contextmenu__item--submenu { justify-content: space-between; position: relative; }
.xlsx-contextmenu__shortcut {
  color: #a1a1aa; font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  font-size: 10px;
}
.xlsx-contextmenu__chevron { color: #71717a; }
.xlsx-contextmenu__checked { color: #2563eb; font-weight: 700; }
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
