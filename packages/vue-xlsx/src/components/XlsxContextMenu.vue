<template>
  <!--
    The grid menu is a real `OfficeDropdownMenu` (Reka): the invisible anchor
    below is positioned at the pointer, so the imperative `openAt(x, y)` API is
    unchanged while content, submenus, typeahead, focus and dismissal come from
    the shared primitive layer.
  -->
  <OfficeDropdownMenu
    v-if="mounted"
    :open="visible"
    :modal="true"
    :items="menuItems"
    :content-test-id="'xlsx-context-menu'"
    :content-data="{ 'data-menu-kind': menuKind }"
    :content-aria-label="m.xlsx.contextMenu.label"
    :content-class="'xlsx-contextmenu'"
    :side-offset="-4"
    :restore-focus="restoreFocusTarget"
    @update:open="onMenuOpenChange"
    @select="onMenuSelect"
  >
    <template #trigger>
      <span
        ref="anchorRef"
        class="xlsx-contextmenu__anchor"
        :style="anchorStyle"
        aria-hidden="true"
        tabindex="-1"
      />
    </template>
  </OfficeDropdownMenu>
  <XlsxClipboardNotice :message="clipboardNotice" />
</template>

<script setup lang="ts">
import { computed, onUnmounted, ref, watch, type CSSProperties } from "vue";
import {
  OfficeDropdownMenu,
  type OfficeMenuCheckboxItem,
  type OfficeMenuEntry,
  type OfficeMenuItem,
} from "@arcships/vue-ui";
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
/**
 * Reka menus need a real DOM, so the primitive mounts on the first open. A
 * closed menu therefore costs nothing and the component stays inert until the
 * grid actually asks for it.
 */
const mounted = ref(false);
const anchorRef = ref<HTMLElement | null>(null);
const m = useXlsxMessages();
const position = ref({ x: 0, y: 0 });
const sortTarget = ref<{ columnIndex: number; tableName: string } | null>(null);
/** Chart the menu was opened on, or `null` for the cell menu. */
const chartTargetId = ref<string | null>(null);
/** Image or shape the menu was opened on, or `null`. */
const drawingTargetId = ref<string | null>(null);
const drawingTargetKind = ref<"image" | "shape" | null>(null);
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
// `isEditingLocked: false` must not mask a read-only workbook, so both flags
// are checked explicitly instead of relying on nullish coalescing.
const isReadOnly = computed(
  () => props.controller.isEditingLocked === true || props.controller.readOnly === true,
);
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

const anchorStyle = computed<CSSProperties>(() => ({
  blockSize: "0px",
  inlineSize: "0px",
  insetBlockStart: `${position.value.y}px`,
  insetInlineStart: `${position.value.x}px`,
  pointerEvents: "none",
  position: "fixed",
}));

/** Focus returns to the surface that opened the menu (the grid element). */
function restoreFocusTarget(): HTMLElement | null {
  return props.targetElement ?? null;
}

// ── Menu model ────────────────────────────────────────────────────────────
function commandItem(
  label: string,
  testId: string,
  run: () => void,
  extra: Partial<OfficeMenuItem> = {},
): OfficeMenuItem {
  return { type: "item", id: testId, label, testId, onSelect: run, ...extra };
}

function checkItem(
  label: string,
  testId: string,
  checked: boolean,
  run: () => void,
): OfficeMenuCheckboxItem {
  return { type: "checkbox", id: testId, label, testId, checked, onSelect: run };
}

const separator = (): OfficeMenuEntry => ({ type: "separator" });

function chartMenuItems(): OfficeMenuEntry[] {
  const families: OfficeMenuEntry[] = chartFamilies.map((family) => ({
    type: "sub",
    id: `chart-family-${family.family}`,
    testId: `xlsx-chart-family-${family.family}`,
    contentTestId: `xlsx-chart-family-${family.family}-menu`,
    label: family.family,
    items: family.options.map((option) => checkItem(
      option.label,
      `xlsx-chart-type-${option.id}`,
      isCurrentChartType(option),
      () => onChangeChartType(option.id),
    )).map((entry, index) => ({
      ...entry,
      disabled: isChartTypeUnavailable(family.options[index]),
    })),
  }));
  return [
    {
      type: "sub",
      id: "change-chart-type",
      testId: "xlsx-context-chart-type",
      contentTestId: "xlsx-chart-type-menu",
      label: m.value.xlsx.contextMenu.changeChartType,
      disabled: isReadOnly.value,
      items: families,
    },
    separator(),
    commandItem(m.value.xlsx.contextMenu.copy, "xlsx-context-drawing-copy", onDrawingCopy, { shortcut: SHORTCUT_COPY }),
    commandItem(m.value.xlsx.contextMenu.cut, "xlsx-context-drawing-cut", onDrawingCut, {
      disabled: isReadOnly.value,
      shortcut: SHORTCUT_CUT,
    }),
    commandItem(m.value.xlsx.contextMenu.paste, "xlsx-context-drawing-paste", onDrawingPaste, {
      disabled: isReadOnly.value,
      shortcut: SHORTCUT_PASTE,
    }),
    commandItem(m.value.xlsx.contextMenu.duplicate, "xlsx-context-drawing-duplicate", onDrawingDuplicate, {
      disabled: isReadOnly.value,
    }),
    commandItem(m.value.xlsx.contextMenu.delete, "xlsx-context-drawing-delete", onDrawingDelete, {
      disabled: isReadOnly.value,
    }),
    separator(),
    commandItem(m.value.xlsx.contextMenu.undo, "xlsx-context-undo", onUndo),
    commandItem(m.value.xlsx.contextMenu.redo, "xlsx-context-redo", onRedo),
  ];
}

function drawingMenuItems(): OfficeMenuEntry[] {
  return [
    commandItem(m.value.xlsx.contextMenu.copy, "xlsx-context-drawing-copy", onDrawingCopy, { shortcut: SHORTCUT_COPY }),
    commandItem(m.value.xlsx.contextMenu.cut, "xlsx-context-drawing-cut", onDrawingCut, {
      disabled: isReadOnly.value,
      shortcut: SHORTCUT_CUT,
    }),
    commandItem(m.value.xlsx.contextMenu.paste, "xlsx-context-drawing-paste", onDrawingPaste, {
      disabled: isReadOnly.value,
      shortcut: SHORTCUT_PASTE,
    }),
    commandItem(m.value.xlsx.contextMenu.duplicate, "xlsx-context-drawing-duplicate", onDrawingDuplicate, {
      disabled: isReadOnly.value,
    }),
    commandItem(m.value.xlsx.contextMenu.delete, "xlsx-context-drawing-delete", onDrawingDelete, {
      disabled: isReadOnly.value,
    }),
    separator(),
    commandItem(m.value.xlsx.contextMenu.undo, "xlsx-context-undo", onUndo),
    commandItem(m.value.xlsx.contextMenu.redo, "xlsx-context-redo", onRedo),
  ];
}

function cellMenuItems(): OfficeMenuEntry[] {
  const entries: OfficeMenuEntry[] = [
    commandItem(m.value.xlsx.contextMenu.cut, "xlsx-context-cut", onCut, {
      disabled: isReadOnly.value,
      shortcut: SHORTCUT_CUT,
    }),
    commandItem(m.value.xlsx.contextMenu.copy, "xlsx-context-copy", onCopy, { shortcut: SHORTCUT_COPY }),
  ];
  if (isReadOnly.value) {
    entries.push(separator());
  } else {
    entries.push(
      commandItem(m.value.xlsx.contextMenu.paste, "xlsx-context-paste", onPaste, { shortcut: SHORTCUT_PASTE }),
      {
        type: "sub",
        id: "paste-special",
        testId: "xlsx-context-paste-special",
        contentTestId: "xlsx-context-paste-special-menu",
        label: m.value.xlsx.contextMenu.pasteSpecial,
        items: [
          commandItem(m.value.xlsx.contextMenu.pasteAll, "xlsx-context-paste-all", () => onPasteSpecial({ mode: "all" })),
          commandItem(m.value.xlsx.contextMenu.pasteValues, "xlsx-context-paste-values", () => onPasteSpecial({ mode: "values" })),
          commandItem(m.value.xlsx.contextMenu.pasteFormats, "xlsx-context-paste-formats", () => onPasteSpecial({ mode: "formats" })),
          commandItem(m.value.xlsx.contextMenu.pasteTranspose, "xlsx-context-paste-transpose", () => onPasteSpecial({ mode: "all", transpose: true })),
        ],
      },
      commandItem(m.value.xlsx.contextMenu.editCell, "xlsx-context-edit-cell", onEditCell, { shortcut: "F2" }),
      commandItem(m.value.xlsx.contextMenu.clear, "xlsx-context-clear", onClear, { shortcut: "Del" }),
    );
    if (menuKind.value === "cell") {
      entries.push(
        commandItem(m.value.xlsx.authoring.numberFormat.title, "xlsx-context-number-format", () => onAuthoring("numberFormat")),
        commandItem(m.value.xlsx.authoring.names.title, "xlsx-context-names", () => onAuthoring("names")),
      );
    }
  }
  if (sortTarget.value && !isReadOnly.value) {
    entries.push(
      separator(),
      commandItem(m.value.xlsx.contextMenu.sortAscending, "xlsx-sort-ascending", () => onSort("ascending")),
      commandItem(m.value.xlsx.contextMenu.sortDescending, "xlsx-sort-descending", () => onSort("descending")),
    );
  }
  if (!isReadOnly.value) {
    entries.push(
      separator(),
      commandItem(m.value.xlsx.contextMenu.insertRowAbove, "xlsx-context-insert-row-above", () => onInsertRow("above")),
      commandItem(m.value.xlsx.contextMenu.insertRowBelow, "xlsx-context-insert-row-below", () => onInsertRow("below")),
      commandItem(m.value.xlsx.contextMenu.deleteRow, "xlsx-context-delete-row", onDeleteRow),
      commandItem(m.value.xlsx.contextMenu.insertColumnLeft, "xlsx-context-insert-column-left", () => onInsertColumn("left")),
      commandItem(m.value.xlsx.contextMenu.insertColumnRight, "xlsx-context-insert-column-right", () => onInsertColumn("right")),
      commandItem(m.value.xlsx.contextMenu.deleteColumn, "xlsx-context-delete-column", onDeleteColumn),
    );
    if (menuKind.value === "cell") {
      entries.push(
        separator(),
        commandItem(m.value.xlsx.contextMenu.mergeAndCenter, "xlsx-context-merge-center", onMergeCenter),
        commandItem(m.value.xlsx.contextMenu.mergeAcross, "xlsx-context-merge-across", onMergeAcross),
        commandItem(m.value.xlsx.contextMenu.mergeCells, "xlsx-context-merge", onMergeCells),
        commandItem(m.value.xlsx.contextMenu.unmerge, "xlsx-context-unmerge", onUnmerge),
      );
    }
  }
  entries.push(
    separator(),
    checkItem(m.value.xlsx.contextMenu.selectRow, "xlsx-context-select-row", selectionIsRow.value, onSelectRow),
    checkItem(m.value.xlsx.contextMenu.selectColumn, "xlsx-context-select-column", selectionIsColumn.value, onSelectColumn),
    checkItem(m.value.xlsx.contextMenu.selectAll, "xlsx-context-select-all", selectionIsAll.value, onSelectAll),
    separator(),
  );
  if (!isReadOnly.value) {
    entries.push(
      commandItem(
        commentAtCell.value ? m.value.xlsx.authoring.context.editComment : m.value.xlsx.authoring.context.insertComment,
        "xlsx-context-comment",
        () => onAuthoring("comment"),
      ),
      commandItem(
        hyperlinkAtCell.value ? m.value.xlsx.authoring.context.editHyperlink : m.value.xlsx.authoring.context.insertHyperlink,
        "xlsx-context-hyperlink",
        () => onAuthoring("hyperlink"),
      ),
      commandItem(m.value.xlsx.authoring.context.insertChart, "xlsx-context-chart", () => onAuthoring("chart")),
      commandItem(m.value.xlsx.authoring.context.conditional, "xlsx-context-conditional", () => onAuthoring("conditional")),
      commandItem(m.value.xlsx.authoring.context.validation, "xlsx-context-validation", () => onAuthoring("validation")),
    );
    if (tableAtCell.value) {
      entries.push(commandItem(m.value.xlsx.authoring.context.removeTable, "xlsx-context-table-remove", onRemoveTable));
    } else {
      entries.push(commandItem(m.value.xlsx.authoring.context.createTable, "xlsx-context-table", () => onAuthoring("table")));
    }
    entries.push(commandItem(
      isFrozen.value ? m.value.xlsx.authoring.context.unfreeze : m.value.xlsx.authoring.context.freezeHere,
      "xlsx-context-freeze",
      onToggleFreeze,
    ));
  } else {
    // Read-only workbooks keep the discovery hooks for the remaining actions.
    entries.push(
      commandItem(m.value.xlsx.authoring.context.insertChart, "xlsx-context-chart", () => onAuthoring("chart"), { disabled: true }),
      commandItem(m.value.xlsx.authoring.context.conditional, "xlsx-context-conditional", () => onAuthoring("conditional"), { disabled: true }),
      commandItem(m.value.xlsx.authoring.context.validation, "xlsx-context-validation", () => onAuthoring("validation"), { disabled: true }),
    );
  }
  if (!isReadOnly.value) {
    entries.push(
      separator(),
      commandItem(m.value.xlsx.contextMenu.undo, "xlsx-context-undo", onUndo, { shortcut: "Ctrl+Z" }),
      commandItem(m.value.xlsx.contextMenu.redo, "xlsx-context-redo", onRedo, { shortcut: "Ctrl+Y" }),
    );
  }
  return entries;
}

const menuItems = computed<OfficeMenuEntry[]>(() => {
  if (chartTarget.value) return chartMenuItems();
  if (drawingTarget.value) return drawingMenuItems();
  return cellMenuItems();
});

function onMenuOpenChange(value: boolean): void {
  if (!value) close();
}

function onMenuSelect(_entry: OfficeMenuEntry, _event: Event): void {
  // Reka closes the menu after a select; the commands already call close().
}

// ── Hit testing ───────────────────────────────────────────────────────────
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

// ── Open / close ──────────────────────────────────────────────────────────
function showMenuAt(clientX: number, clientY: number): void {
  position.value = { x: clientX, y: clientY };
  mounted.value = true;
  visible.value = true;
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
    showMenuAt(clientX, clientY);
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
  showMenuAt(clientX, clientY);
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
  showMenuAt(clientX, clientY);
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
  showMenuAt(clientX, clientY);
  return true;
}


// ── Chart type helpers ────────────────────────────────────────────────────
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

function close(): void {
  visible.value = false;
  sortTarget.value = null;
  chartTargetId.value = null;
  drawingTargetId.value = null;
  drawingTargetKind.value = null;
}

// ── Commands ──────────────────────────────────────────────────────────────
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
/*
 * The menu surface, items, separators, shortcuts and submenus come from
 * `OfficeDropdownMenu` (Reka + the shared `.office-menu__*` contract). Only
 * the invisible anchor that places the menu at the pointer is local.
 */
.xlsx-contextmenu__anchor {
  border: 0;
  margin: 0;
  padding: 0;
}
</style>
