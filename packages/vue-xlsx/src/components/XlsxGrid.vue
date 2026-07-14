<template>
  <div
    ref="containerRef"
    class="xlsx-grid"
    data-testid="xlsx-grid"
    :data-display-row-count="effectiveRowHeights.length"
    :data-display-column-count="effectiveColWidths.length"
    :data-first-column-width="effectiveColWidths[0] ?? DEFAULT_COL_WIDTH"
    :data-first-row-height="effectiveRowHeights[0] ?? DEFAULT_ROW_HEIGHT"
    :data-merged-region-count="activeSheet?.mergedRegions?.length ?? 0"
    :data-worker-backed="String(Boolean(controller.isWorkerBacked))"
    tabindex="0"
    :style="containerStyle"
    @keydown="onGridKeydown"
    @scroll="onScroll"
    @pointerdown="onPointerDown"
    @pointermove="onPointerMove"
    @pointerup="onPointerUp"
    @pointercancel="onPointerUp"
    @pointerleave="onPointerLeave"
    @dblclick="onDblClick"
  >
    <div
      class="xlsx-grid__scroll-space"
      data-testid="xlsx-grid-scroll-space"
      :style="scrollSpaceStyle"
    >
      <div class="xlsx-grid__viewport" :style="viewportStyle">
        <!-- Column header canvas -->
        <canvas
          ref="colHeaderCanvasRef"
          class="xlsx-grid__col-header"
          :style="colHeaderCanvasStyle"
        />
        <!-- Row header canvas -->
        <canvas
          ref="rowHeaderCanvasRef"
          class="xlsx-grid__row-header"
          :style="rowHeaderCanvasStyle"
        />
        <!-- Corner canvas -->
        <canvas
          ref="cornerCanvasRef"
          class="xlsx-grid__corner"
          :style="cornerCanvasStyle"
        />
        <!-- Main body canvas -->
        <canvas
          ref="bodyCanvasRef"
          class="xlsx-grid__body"
          :style="bodyCanvasStyle"
        />
        <!-- Editing input (hidden unless editing) -->
        <input
          v-if="editingCell"
          ref="editInputRef"
          v-model="editingValue"
          class="xlsx-grid__edit-input"
          data-testid="xlsx-cell-editor"
          :style="editInputStyle"
          @keydown.enter.stop.prevent="commitEdit"
          @keydown.escape.stop.prevent="cancelEdit"
          @keydown.tab.stop.prevent="commitAndTab"
          @blur="commitEdit"
        />
        <div
          v-if="hoveredComment"
          class="xlsx-grid__comment-popover"
          data-testid="xlsx-comment-popover"
          :style="hoveredComment.style"
        >
          <strong v-if="hoveredComment.author">{{ hoveredComment.author }}</strong>
          <span>{{ hoveredComment.text }}</span>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, onMounted, onUnmounted, nextTick, type CSSProperties } from "vue";
import {
  resolveWorkbookColor,
  resolveWorkbookFillStyle,
  type XlsxViewerController,
  type XlsxCellAddress,
  type XlsxCellRange,
  type XlsxCellStyleContext,
  type XlsxCellHyperlink,
  type XlsxCellComment,
  type XlsxConditionalFormatRule,
  type XlsxSparkline,
  type XlsxResolvedCellStyle,
  type XlsxSheetData,
} from "@arcships/xlsx-core";

const props = defineProps<{
  controller: XlsxViewerController;
  getCellStyle?: ((cell: XlsxCellAddress, context?: XlsxCellStyleContext) => Partial<CSSProperties> | undefined) | null;
  isDark?: boolean;
  readOnly?: boolean;
  selectionColor?: string;
  selectionFillColor?: string;
}>();

const emit = defineEmits<{
  cellDoubleClick: [cell: XlsxCellAddress];
  viewportChange: [viewport: { scrollLeft: number; scrollTop: number }];
}>();

// ── Constants ──────────────────────────────────────────────────────────
const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const HEADER_HEIGHT = 24;
const ROW_HEADER_WIDTH = 48;
const OVERSCAN = 240;
const WORKER_ROW_BATCH_SIZE = 64;
const MIN_DISPLAY_ROWS = 200;
const MIN_DISPLAY_COLS = 50;
const AXIS_RESIZE_DRAG_THRESHOLD = 3;
const DPR = typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1;

// ── Refs ───────────────────────────────────────────────────────────────
const containerRef = ref<HTMLDivElement | null>(null);
const bodyCanvasRef = ref<HTMLCanvasElement | null>(null);
const colHeaderCanvasRef = ref<HTMLCanvasElement | null>(null);
const rowHeaderCanvasRef = ref<HTMLCanvasElement | null>(null);
const cornerCanvasRef = ref<HTMLCanvasElement | null>(null);
const editInputRef = ref<HTMLInputElement | null>(null);

const scrollTop = ref(0);
const scrollLeft = ref(0);
const containerWidth = ref(0);
const containerHeight = ref(0);

const editingCell = ref<XlsxCellAddress | null>(null);
const editingValue = ref("");
const hoveredCell = ref<XlsxCellAddress | null>(null);
const resizePreview = ref<{
  actualIndex: number;
  displayIndex: number;
  kind: "column" | "row";
  moved: boolean;
  pointerId: number;
  size: number;
  startPosition: number;
  startSize: number;
} | null>(null);

const isSelecting = ref(false);
const selectionAnchor = ref<XlsxCellAddress | null>(null);
const workerRowsRevision = ref(0);
interface GridCellSnapshot {
  style?: XlsxResolvedCellStyle;
  value: string;
}

interface GridMergeRegion {
  end: XlsxCellAddress;
  start: XlsxCellAddress;
}

const workerCellCache = new Map<string, GridCellSnapshot>();
const loadedWorkerBatches = new Set<string>();
const pendingWorkerBatches = new Map<string, AbortController>();
let workerCacheSheetIndex: number | null = null;
let paintedMergeRegions: GridMergeRegion[] = [];
let paintedMergeAnchors = new Map<string, XlsxCellAddress>();
let conditionalRuleStats = new Map<XlsxConditionalFormatRule, { min: number; max: number }>();

// ── Computed from controller ──────────────────────────────────────────
const activeSheet = computed(() => props.controller.activeSheet as XlsxSheetData | null);
const hyperlinksByCell = computed(() => new Map((activeSheet.value?.hyperlinks ?? []).map((link) => [`${link.cell.row}:${link.cell.col}`, link])));
const commentsByCell = computed(() => new Map((activeSheet.value?.comments ?? []).map((comment) => [`${comment.cell.row}:${comment.cell.col}`, comment])));
const sparklinesByCell = computed(() => new Map((activeSheet.value?.sparklines ?? []).map((sparkline) => [`${sparkline.target.row}:${sparkline.target.col}`, sparkline])));
const hoveredComment = computed(() => {
  const cell = hoveredCell.value;
  if (!cell) return null;
  const comment = commentsByCell.value.get(`${cell.row}:${cell.col}`);
  const display = getDisplayCell(cell);
  if (!comment || !display) return null;
  return {
    ...comment,
    style: {
      left: `${ROW_HEADER_WIDTH + getColOffsetSum(display.col + 1) - scrollLeft.value + 6}px`,
      top: `${HEADER_HEIGHT + getRowOffsetSum(display.row) - scrollTop.value + 4}px`,
    } as CSSProperties,
  };
});
const zoomScale = computed(() => props.controller.zoomScale);
const activeCell = computed(() => props.controller.activeCell);
const selection = computed(() => props.controller.selection);
const revision = computed(() => props.controller.revision);

const effectiveColWidths = computed(() => {
  const sheet = activeSheet.value;
  if (!sheet) return [];
  const z = zoomScale.value / 100;
  const count = Math.max(sheet.colWidths?.length ?? 0, MIN_DISPLAY_COLS);
  return Array.from({ length: count }, (_, index) => {
    const actualIndex = sheet.visibleCols?.[index] ?? index;
    const preview = resizePreview.value;
    const size = preview?.kind === "column" && preview.actualIndex === actualIndex
      ? preview.size
      : sheet.colWidths?.[index] || sheet.colWidthOverridesPx?.[actualIndex] || sheet.defaultColWidthPx || DEFAULT_COL_WIDTH;
    return size * z;
  });
});

const effectiveRowHeights = computed(() => {
  const sheet = activeSheet.value;
  if (!sheet) return [];
  const z = zoomScale.value / 100;
  const count = Math.max(sheet.rowHeights?.length ?? 0, MIN_DISPLAY_ROWS);
  return Array.from({ length: count }, (_, index) => {
    const actualIndex = sheet.visibleRows?.[index] ?? index;
    const preview = resizePreview.value;
    const size = preview?.kind === "row" && preview.actualIndex === actualIndex
      ? preview.size
      : sheet.rowHeights?.[index] || sheet.rowHeightOverridesPx?.[actualIndex] || sheet.defaultRowHeightPx || DEFAULT_ROW_HEIGHT;
    return size * z;
  });
});

function createOffsetIndex(sizes: number[]): number[] {
  const offsets = new Array<number>(sizes.length + 1);
  offsets[0] = 0;
  for (let index = 0; index < sizes.length; index += 1) {
    offsets[index + 1] = offsets[index] + Math.max(0, sizes[index] || 0);
  }
  return offsets;
}

function findDisplayIndexAtOffset(offsets: number[], offset: number): number {
  const itemCount = offsets.length - 1;
  const total = offsets[itemCount] ?? 0;
  if (itemCount <= 0 || offset < 0 || offset >= total) return -1;

  let low = 1;
  let high = itemCount;
  while (low < high) {
    const middle = Math.floor((low + high) / 2);
    if ((offsets[middle] ?? 0) > offset) high = middle;
    else low = middle + 1;
  }
  return low - 1;
}

function resolveVisibleRange(offsets: number[], startOffset: number, endOffset: number) {
  const itemCount = offsets.length - 1;
  if (itemCount <= 0) return { start: 0, end: -1 };

  const total = offsets[itemCount] ?? 0;
  const start = findDisplayIndexAtOffset(offsets, Math.max(0, startOffset));
  const end = findDisplayIndexAtOffset(offsets, Math.max(0, endOffset));
  return {
    start: start >= 0 ? start : itemCount - 1,
    end: end >= 0 && endOffset < total ? end : itemCount - 1,
  };
}

function resolveDisplayIndices(indices: number[] | undefined, count: number): number[] {
  if (!indices?.length) return Array.from({ length: count }, (_, index) => index);
  if (indices.length >= count) return indices.slice(0, count);
  const resolved = [...indices];
  let nextIndex = Math.max(-1, ...resolved) + 1;
  while (resolved.length < count) resolved.push(nextIndex++);
  return resolved;
}

function createSheetToDisplayIndex(indices: number[]): Map<number, number> {
  return new Map(indices.map((sheetIndex, displayIndex) => [sheetIndex, displayIndex]));
}

function resolveAdjacentSheetCell(
  cell: XlsxCellAddress,
  rowDelta: number,
  colDelta: number,
  rows: number[],
  cols: number[],
  rowIndex: Map<number, number>,
  colIndex: Map<number, number>,
): XlsxCellAddress | null {
  const displayRow = rowIndex.get(cell.row);
  const displayCol = colIndex.get(cell.col);
  if (displayRow === undefined || displayCol === undefined || rows.length === 0 || cols.length === 0) return null;
  const nextDisplayRow = Math.max(0, Math.min(displayRow + rowDelta, rows.length - 1));
  const nextDisplayCol = Math.max(0, Math.min(displayCol + colDelta, cols.length - 1));
  return { row: rows[nextDisplayRow], col: cols[nextDisplayCol] };
}

const colOffsets = computed(() => createOffsetIndex(effectiveColWidths.value));
const rowOffsets = computed(() => createOffsetIndex(effectiveRowHeights.value));
const displayCols = computed(() => resolveDisplayIndices(activeSheet.value?.visibleCols, effectiveColWidths.value.length));
const displayRows = computed(() => resolveDisplayIndices(activeSheet.value?.visibleRows, effectiveRowHeights.value.length));
const sheetColToDisplayIndex = computed(() => createSheetToDisplayIndex(displayCols.value));
const sheetRowToDisplayIndex = computed(() => createSheetToDisplayIndex(displayRows.value));

const visibleRowRange = computed(() => {
  const bodyHeight = Math.max(0, containerHeight.value - HEADER_HEIGHT);
  return resolveVisibleRange(
    rowOffsets.value,
    scrollTop.value - OVERSCAN,
    scrollTop.value + bodyHeight + OVERSCAN,
  );
});

const visibleColRange = computed(() => {
  const bodyWidth = Math.max(0, containerWidth.value - ROW_HEADER_WIDTH);
  return resolveVisibleRange(
    colOffsets.value,
    scrollLeft.value - OVERSCAN,
    scrollLeft.value + bodyWidth + OVERSCAN,
  );
});

// ── Helpers ───────────────────────────────────────────────────────────
function getColOffsetSum(col: number): number {
  return colOffsets.value[Math.max(0, Math.min(col, effectiveColWidths.value.length))] ?? 0;
}

function getRowOffsetSum(row: number): number {
  return rowOffsets.value[Math.max(0, Math.min(row, effectiveRowHeights.value.length))] ?? 0;
}

function getSheetCell(displayRow: number, displayCol: number): XlsxCellAddress | null {
  const row = displayRows.value[displayRow];
  const col = displayCols.value[displayCol];
  if (row === undefined || col === undefined) return null;
  return { row, col };
}

function getDisplayCell(cell: XlsxCellAddress): XlsxCellAddress | null {
  const row = sheetRowToDisplayIndex.value.get(cell.row);
  const col = sheetColToDisplayIndex.value.get(cell.col);
  if (row === undefined || col === undefined) return null;
  return { row, col };
}

type XlsxZoomAnchor = {
  row?: number;
  col?: number;
  rowRatio: number;
  colRatio: number;
  clientX: number;
  clientY: number;
  scrollX: boolean;
  scrollY: boolean;
};

function frozenPaneExtents(): { width: number; height: number } {
  const sheet = activeSheet.value;
  const freezeRow = sheet?.freezePanes?.row ?? 0;
  const freezeCol = sheet?.freezePanes?.col ?? 0;
  let width = 0;
  let height = 0;
  for (let index = 0; index < displayCols.value.length; index += 1) {
    if ((displayCols.value[index] ?? 0) >= freezeCol) break;
    width += effectiveColWidths.value[index] || 0;
  }
  for (let index = 0; index < displayRows.value.length; index += 1) {
    if ((displayRows.value[index] ?? 0) >= freezeRow) break;
    height += effectiveRowHeights.value[index] || 0;
  }
  return { width, height };
}

function captureZoomAnchor(clientX: number, clientY: number): XlsxZoomAnchor | null {
  const container = containerRef.value;
  if (!container) return null;
  const rect = container.getBoundingClientRect();
  const localX = clientX - rect.left;
  const localY = clientY - rect.top;
  const bodyX = localX - ROW_HEADER_WIDTH;
  const bodyY = localY - HEADER_HEIGHT;
  const frozen = frozenPaneExtents();
  const scrollX = localX > ROW_HEADER_WIDTH && bodyX >= frozen.width;
  const scrollY = localY > HEADER_HEIGHT && bodyY >= frozen.height;
  if (!scrollX && !scrollY) return null;

  const anchor: XlsxZoomAnchor = {
    rowRatio: 0,
    colRatio: 0,
    clientX,
    clientY,
    scrollX,
    scrollY,
  };
  if (scrollX) {
    const contentX = Math.max(0, bodyX + container.scrollLeft);
    const displayCol = findDisplayIndexAtOffset(colOffsets.value, contentX);
    const col = displayCols.value[displayCol];
    const width = effectiveColWidths.value[displayCol] || 1;
    if (displayCol < 0 || col === undefined) return null;
    anchor.col = col;
    anchor.colRatio = (contentX - getColOffsetSum(displayCol)) / width;
  }
  if (scrollY) {
    const contentY = Math.max(0, bodyY + container.scrollTop);
    const displayRow = findDisplayIndexAtOffset(rowOffsets.value, contentY);
    const row = displayRows.value[displayRow];
    const height = effectiveRowHeights.value[displayRow] || 1;
    if (displayRow < 0 || row === undefined) return null;
    anchor.row = row;
    anchor.rowRatio = (contentY - getRowOffsetSum(displayRow)) / height;
  }
  return anchor;
}

function restoreZoomAnchor(anchor: XlsxZoomAnchor): void {
  const container = containerRef.value;
  if (!container) return;
  const rect = container.getBoundingClientRect();
  if (anchor.scrollX && anchor.col !== undefined) {
    const displayCol = sheetColToDisplayIndex.value.get(anchor.col);
    if (displayCol !== undefined) {
      const contentX = getColOffsetSum(displayCol)
        + (effectiveColWidths.value[displayCol] || 0) * anchor.colRatio;
      container.scrollLeft = contentX - (anchor.clientX - rect.left - ROW_HEADER_WIDTH);
    }
  }
  if (anchor.scrollY && anchor.row !== undefined) {
    const displayRow = sheetRowToDisplayIndex.value.get(anchor.row);
    if (displayRow !== undefined) {
      const contentY = getRowOffsetSum(displayRow)
        + (effectiveRowHeights.value[displayRow] || 0) * anchor.rowRatio;
      container.scrollTop = contentY - (anchor.clientY - rect.top - HEADER_HEIGHT);
    }
  }
  schedulePaint();
}

function workerCellKey(sheetIndex: number, row: number, col: number) {
  return `${sheetIndex}:${row}:${col}`;
}

function resetWorkerRows(sheetIndex: number | null) {
  for (const controller of pendingWorkerBatches.values()) controller.abort();
  pendingWorkerBatches.clear();
  loadedWorkerBatches.clear();
  workerCellCache.clear();
  workerCacheSheetIndex = sheetIndex;
  workerRowsRevision.value += 1;
}

function getGridCellDisplayValue(cell: XlsxCellAddress) {
  workerRowsRevision.value;
  const sheetIndex = activeSheet.value?.workbookSheetIndex;
  if (sheetIndex !== undefined && props.controller.isWorkerBacked) {
    const cached = workerCellCache.get(workerCellKey(sheetIndex, cell.row, cell.col));
    if (cached !== undefined) return cached.value;
  }
  return props.controller.getCellDisplayValue(cell);
}

function getWorkerCellStyle(cell: XlsxCellAddress): XlsxResolvedCellStyle | undefined {
  workerRowsRevision.value;
  const sheetIndex = activeSheet.value?.workbookSheetIndex;
  if (sheetIndex === undefined || !props.controller.isWorkerBacked) return undefined;
  return workerCellCache.get(workerCellKey(sheetIndex, cell.row, cell.col))?.style;
}

function asWorkerCellValue(value: unknown) {
  if (value === null || value === undefined) return "";
  return typeof value === "string" ? value : String(value);
}

function hydrateVisibleWorkerRows() {
  const sheet = activeSheet.value;
  const loadRows = props.controller.getRowsBatchAsync;
  if (!sheet || !props.controller.isWorkerBacked || !loadRows) return;

  const sheetIndex = sheet.workbookSheetIndex;
  if (workerCacheSheetIndex !== sheetIndex) resetWorkerRows(sheetIndex);
  const visibleRows = displayRows.value.slice(
    visibleRowRange.value.start,
    visibleRowRange.value.end + 1,
  );
  if (visibleRows.length === 0) return;
  const firstRow = Math.min(...visibleRows);
  const lastRow = Math.max(...visibleRows);
  const firstBatch = Math.floor(firstRow / WORKER_ROW_BATCH_SIZE) * WORKER_ROW_BATCH_SIZE;
  const lastBatch = Math.floor(lastRow / WORKER_ROW_BATCH_SIZE) * WORKER_ROW_BATCH_SIZE;

  for (let startRow = firstBatch; startRow <= lastBatch; startRow += WORKER_ROW_BATCH_SIZE) {
    const batchKey = `${sheetIndex}:${startRow}`;
    if (loadedWorkerBatches.has(batchKey) || pendingWorkerBatches.has(batchKey)) continue;
    const controller = new AbortController();
    pendingWorkerBatches.set(batchKey, controller);
    void loadRows(sheetIndex, startRow, WORKER_ROW_BATCH_SIZE, controller.signal)
      .then((rows) => {
        if (controller.signal.aborted || workerCacheSheetIndex !== sheetIndex) return;
        if (Array.isArray(rows)) {
          for (const rowEntry of rows) {
            if (!rowEntry || typeof rowEntry !== "object") continue;
            const row = Reflect.get(rowEntry, "index");
            const cells = Reflect.get(rowEntry, "cells");
            if (!Number.isInteger(row) || !Array.isArray(cells)) continue;
            for (const cellEntry of cells) {
              if (!cellEntry || typeof cellEntry !== "object") continue;
              const col = Reflect.get(cellEntry, "col");
              if (!Number.isInteger(col)) continue;
              const style = Reflect.get(cellEntry, "style");
              workerCellCache.set(
                workerCellKey(sheetIndex, row, col),
                {
                  value: asWorkerCellValue(Reflect.get(cellEntry, "value")),
                  ...(style && typeof style === "object" ? { style: style as XlsxResolvedCellStyle } : {}),
                },
              );
            }
          }
        }
        loadedWorkerBatches.add(batchKey);
        workerRowsRevision.value += 1;
        schedulePaint();
      })
      .finally(() => {
        pendingWorkerBatches.delete(batchKey);
      });
  }
}

function columnLabel(col: number): string {
  let label = "";
  let n = col;
  while (n >= 0) { label = String.fromCharCode(65 + (n % 26)) + label; n = Math.floor(n / 26) - 1; }
  return label;
}

function truncateText(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string {
  if (ctx.measureText(text).width <= maxWidth) return text;
  let lo = 0, hi = text.length;
  while (lo < hi) {
    const mid = Math.ceil((lo + hi) / 2);
    if (ctx.measureText(text.slice(0, mid) + "\u2026").width <= maxWidth) lo = mid;
    else hi = mid - 1;
  }
  return text.slice(0, lo) + "\u2026";
}

function isCellInRange(cell: XlsxCellAddress, range: XlsxCellRange | null): boolean {
  if (!range) return false;
  return cell.row >= Math.min(range.start.row, range.end.row) &&
    cell.row <= Math.max(range.start.row, range.end.row) &&
    cell.col >= Math.min(range.start.col, range.end.col) &&
    cell.col <= Math.max(range.start.col, range.end.col);
}

function rangesIntersect(left: XlsxCellRange, right: XlsxCellRange): boolean {
  const leftStartRow = Math.min(left.start.row, left.end.row);
  const leftEndRow = Math.max(left.start.row, left.end.row);
  const leftStartCol = Math.min(left.start.col, left.end.col);
  const leftEndCol = Math.max(left.start.col, left.end.col);
  const rightStartRow = Math.min(right.start.row, right.end.row);
  const rightEndRow = Math.max(right.start.row, right.end.row);
  const rightStartCol = Math.min(right.start.col, right.end.col);
  const rightEndCol = Math.max(right.start.col, right.end.col);
  return leftStartRow <= rightEndRow && leftEndRow >= rightStartRow && leftStartCol <= rightEndCol && leftEndCol >= rightStartCol;
}

function drawSelectionBorder(ctx: CanvasRenderingContext2D, range: XlsxCellRange, color: string): void {
  const normalized = {
    start: { row: Math.min(range.start.row, range.end.row), col: Math.min(range.start.col, range.end.col) },
    end: { row: Math.max(range.start.row, range.end.row), col: Math.max(range.start.col, range.end.col) },
  };
  const start = getDisplayCell(normalized.start);
  const end = getDisplayCell(normalized.end);
  if (!start || !end) return;
  const x = getColOffsetSum(start.col);
  const y = getRowOffsetSum(start.row);
  const width = getColOffsetSum(end.col + 1) - x;
  const height = getRowOffsetSum(end.row + 1) - y;
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = 2;
  ctx.strokeRect(x + 1, y + 1, Math.max(0, width - 2), Math.max(0, height - 2));
  ctx.restore();
}

function normalizeMergeRegion(value: unknown): GridMergeRegion | null {
  if (!value || typeof value !== "object") return null;
  const record = value as Record<string, unknown>;
  const start = record.start && typeof record.start === "object" ? record.start as Record<string, unknown> : undefined;
  const end = record.end && typeof record.end === "object" ? record.end as Record<string, unknown> : undefined;
  const startRow = Number(start?.row ?? record.startRow);
  const startCol = Number(start?.col ?? record.startCol);
  const endRow = Number(end?.row ?? record.endRow);
  const endCol = Number(end?.col ?? record.endCol);
  if (![startRow, startCol, endRow, endCol].every(Number.isFinite)) return null;
  return {
    start: { row: Math.min(startRow, endRow), col: Math.min(startCol, endCol) },
    end: { row: Math.max(startRow, endRow), col: Math.max(startCol, endCol) },
  };
}

function updatePaintedMergeRegions(rawRegions: unknown): void {
  const regions = Array.isArray(rawRegions)
    ? rawRegions.map(normalizeMergeRegion).filter((region): region is GridMergeRegion => region !== null)
    : [];
  const anchors = new Map<string, XlsxCellAddress>();
  for (const region of regions) {
    for (let row = region.start.row; row <= region.end.row; row += 1) {
      for (let col = region.start.col; col <= region.end.col; col += 1) {
        if (row === region.start.row && col === region.start.col) continue;
        anchors.set(`${row}:${col}`, region.start);
      }
    }
  }
  paintedMergeRegions = regions;
  paintedMergeAnchors = anchors;
}

function resolveMergeAnchor(cell: XlsxCellAddress): XlsxCellAddress {
  return paintedMergeAnchors.get(`${cell.row}:${cell.col}`) ?? cell;
}

function mergeRegionAtAnchor(cell: XlsxCellAddress): GridMergeRegion | undefined {
  return paintedMergeRegions.find((region) =>
    region.start.row === cell.row && region.start.col === cell.col,
  );
}

function hitTestCell(clientX: number, clientY: number): XlsxCellAddress | null {
  const container = containerRef.value;
  if (!container) return null;
  const rect = container.getBoundingClientRect();
  const x = clientX - rect.left + container.scrollLeft - ROW_HEADER_WIDTH;
  const y = clientY - rect.top + container.scrollTop - HEADER_HEIGHT;
  if (x < 0 || y < 0) return null;

  const displayCol = findDisplayIndexAtOffset(colOffsets.value, x);
  const displayRow = findDisplayIndexAtOffset(rowOffsets.value, y);
  if (displayCol < 0 || displayRow < 0) return null;
  const cell = getSheetCell(displayRow, displayCol);
  return cell ? resolveMergeAnchor(cell) : null;
}

function hitTestSheetPoint(clientX: number, clientY: number): { row: number; col: number; xOffset: number; yOffset: number } | null {
  const container = containerRef.value;
  if (!container) return null;
  const rect = container.getBoundingClientRect();
  const x = clientX - rect.left + container.scrollLeft - ROW_HEADER_WIDTH;
  const y = clientY - rect.top + container.scrollTop - HEADER_HEIGHT;
  if (x < 0 || y < 0) return null;
  const displayCol = findDisplayIndexAtOffset(colOffsets.value, x);
  const displayRow = findDisplayIndexAtOffset(rowOffsets.value, y);
  if (displayCol < 0 || displayRow < 0) return null;
  const cell = getSheetCell(displayRow, displayCol);
  if (!cell) return null;
  const left = getColOffsetSum(displayCol);
  const top = getRowOffsetSum(displayRow);
  const width = Math.max(1, effectiveColWidths.value[displayCol] || DEFAULT_COL_WIDTH);
  const height = Math.max(1, effectiveRowHeights.value[displayRow] || DEFAULT_ROW_HEIGHT);
  return {
    row: cell.row,
    col: cell.col,
    xOffset: Math.max(0, Math.min(1, (x - left) / width)),
    yOffset: Math.max(0, Math.min(1, (y - top) / height)),
  };
}

// ── Canvas painting ───────────────────────────────────────────────────
function resizeCanvas(canvas: HTMLCanvasElement | null, width: number, height: number) {
  if (!canvas) return;
  const safeWidth = Math.max(0, width);
  const safeHeight = Math.max(0, height);
  const pixelWidth = Math.ceil(safeWidth * DPR);
  const pixelHeight = Math.ceil(safeHeight * DPR);
  const cssWidth = `${safeWidth}px`;
  const cssHeight = `${safeHeight}px`;
  if (canvas.width !== pixelWidth) canvas.width = pixelWidth;
  if (canvas.height !== pixelHeight) canvas.height = pixelHeight;
  if (canvas.style.width !== cssWidth) canvas.style.width = cssWidth;
  if (canvas.style.height !== cssHeight) canvas.style.height = cssHeight;
  const ctx = canvas.getContext("2d");
  if (ctx) ctx.setTransform(DPR, 0, 0, DPR, 0, 0);
}

interface CanvasCellStyle {
  alignment: "left" | "center" | "right";
  backgroundColor: string;
  border?: Record<string, Record<string, unknown>>;
  color: string;
  fontFamily: string;
  fontSizePx: number;
  fontStyle: string;
  fontWeight: string;
  strike: boolean;
  underline: boolean;
  vertical: "top" | "middle" | "bottom";
  wrap: boolean;
}

function asStyleRecord(value: unknown): Record<string, unknown> | undefined {
  return value && typeof value === "object" ? value as Record<string, unknown> : undefined;
}

function cssText(value: unknown): string | undefined {
  return typeof value === "string" || typeof value === "number" ? String(value) : undefined;
}

function cellInRange(cell: XlsxCellAddress, range: XlsxCellRange): boolean {
  return cell.row >= Math.min(range.start.row, range.end.row)
    && cell.row <= Math.max(range.start.row, range.end.row)
    && cell.col >= Math.min(range.start.col, range.end.col)
    && cell.col <= Math.max(range.start.col, range.end.col);
}

function conditionalRulesForCell(cell: XlsxCellAddress): XlsxConditionalFormatRule[] {
  return (activeSheet.value?.conditionalFormatRules ?? []).filter((rule) => rule.ranges.some((range) => cellInRange(cell, range)));
}

function ruleStats(rule: XlsxConditionalFormatRule): { min: number; max: number } {
  const cached = conditionalRuleStats.get(rule);
  if (cached) return cached;
  const sheet = activeSheet.value;
  if (!sheet) return { min: 0, max: 0 };
  const values: number[] = [];
  for (const range of rule.ranges) {
    const endRow = Math.min(range.end.row, Math.max(sheet.maxUsedRow, range.start.row));
    const endCol = Math.min(range.end.col, Math.max(sheet.maxUsedCol, range.start.col));
    for (let row = range.start.row; row <= endRow; row++) {
      for (let col = range.start.col; col <= endCol; col++) {
        const value = Number(getGridCellDisplayValue({ row, col }).replace(/[,％%]/g, ""));
        if (Number.isFinite(value)) values.push(value);
      }
    }
  }
  const stats = values.length ? { min: Math.min(...values), max: Math.max(...values) } : { min: 0, max: 0 };
  conditionalRuleStats.set(rule, stats);
  return stats;
}

function colorRgb(value: Record<string, unknown> | undefined, fallback: string): [number, number, number] {
  const resolved = resolveWorkbookColor(value, activeSheet.value?.themePalette) || fallback;
  const hex = resolved.match(/^#?([0-9a-f]{6})$/i)?.[1];
  if (!hex) return [37, 99, 235];
  return [Number.parseInt(hex.slice(0, 2), 16), Number.parseInt(hex.slice(2, 4), 16), Number.parseInt(hex.slice(4, 6), 16)];
}

function mixColor(start: [number, number, number], end: [number, number, number], ratio: number): string {
  const channel = (index: number) => Math.round(start[index] + (end[index] - start[index]) * ratio);
  return `rgb(${channel(0)}, ${channel(1)}, ${channel(2)})`;
}

function conditionalVisual(cell: XlsxCellAddress, valueText: string) {
  const rules = conditionalRulesForCell(cell);
  if (!rules.length) return null;
  const value = Number(valueText.replace(/[,％%]/g, ""));
  if (!Number.isFinite(value)) return null;
  const rule = rules[0];
  const stats = ruleStats(rule);
  const min = stats.min;
  const max = stats.max;
  const ratio = max === min ? 1 : Math.max(0, Math.min(1, (value - min) / (max - min)));
  if (rule.kind === "dataBar") {
    return { kind: rule.kind, ratio, color: resolveWorkbookColor(rule.color, activeSheet.value?.themePalette) || "#5b9bd5", showValue: rule.showValue !== false } as const;
  }
  if (rule.kind === "colorScale") {
    const colors = rule.colors;
    const low = colorRgb(colors[0], "#f8696b");
    const middle = colorRgb(colors[Math.floor((colors.length - 1) / 2)], "#ffeb84");
    const high = colorRgb(colors[colors.length - 1], "#63be7b");
    const background = ratio <= .5 ? mixColor(low, middle, ratio * 2) : mixColor(middle, high, (ratio - .5) * 2);
    return { kind: rule.kind, ratio, background, showValue: true } as const;
  }
  const symbols = rule.reverse ? ["▲", "●", "▼"] : ["▼", "●", "▲"];
  return { kind: rule.kind, ratio, icon: symbols[Math.min(2, Math.floor(ratio * 3))], showValue: rule.showValue !== false } as const;
}

function drawSparkline(ctx: CanvasRenderingContext2D, sparkline: XlsxSparkline, x: number, y: number, width: number, height: number): void {
  const values: number[] = [];
  for (let row = sparkline.range.start.row; row <= sparkline.range.end.row; row++) {
    for (let col = sparkline.range.start.col; col <= sparkline.range.end.col; col++) {
      const value = Number(getGridCellDisplayValue({ row, col }).replace(/,/g, ""));
      if (Number.isFinite(value)) values.push(value);
    }
  }
  if (!values.length) return;
  const min = Math.min(...values, 0);
  const max = Math.max(...values, 0);
  const span = Math.max(1e-9, max - min);
  const px = (index: number) => x + 4 + index * Math.max(1, (width - 8) / Math.max(1, values.length - 1));
  const py = (value: number) => y + height - 4 - ((value - min) / span) * Math.max(1, height - 8);
  ctx.save();
  ctx.strokeStyle = sparkline.color || "#2563eb";
  ctx.fillStyle = sparkline.color || "#2563eb";
  ctx.lineWidth = 1.5;
  if (sparkline.type === "column" || sparkline.type === "winLoss") {
    const barWidth = Math.max(2, (width - 8) / values.length - 2);
    values.forEach((value, index) => ctx.fillRect(px(index) - barWidth / 2, Math.min(py(value), py(0)), barWidth, Math.max(1, Math.abs(py(value) - py(0)))));
  } else {
    ctx.beginPath();
    values.forEach((value, index) => index ? ctx.lineTo(px(index), py(value)) : ctx.moveTo(px(index), py(value)));
    ctx.stroke();
    if (sparkline.markers) values.forEach((value, index) => { ctx.beginPath(); ctx.arc(px(index), py(value), 2, 0, Math.PI * 2); ctx.fill(); });
  }
  ctx.restore();
}

function resolveCanvasCellStyle(
  cell: XlsxCellAddress,
  value: string,
  workbookStyle?: XlsxResolvedCellStyle,
): CanvasCellStyle {
  const sheet = activeSheet.value;
  const font = asStyleRecord(workbookStyle?.font);
  const fill = asStyleRecord(workbookStyle?.fill);
  const alignment = asStyleRecord(workbookStyle?.alignment);
  const fillStyle = resolveWorkbookFillStyle(fill, sheet?.themePalette);
  const fontColor = resolveWorkbookColor(asStyleRecord(font?.color), sheet?.themePalette);
  const merged = Boolean(mergeRegionAtAnchor(cell));
  const tableHeader = props.controller.tables.some((table) =>
    cell.row === table.start.row && cell.col >= table.start.col && cell.col <= table.end.col,
  );
  const base: CanvasCellStyle = {
    alignment: alignment?.horizontal === "center" || alignment?.horizontal === "right"
      ? alignment.horizontal
      : "left",
    backgroundColor: fillStyle.backgroundColor || (props.isDark ? "#18181b" : "#ffffff"),
    border: asStyleRecord(workbookStyle?.border) as Record<string, Record<string, unknown>> | undefined,
    color: fontColor || (props.isDark ? "#e4e4e7" : "#18181b"),
    fontFamily: typeof font?.name === "string" && font.name.trim() ? font.name : "Segoe UI",
    fontSizePx: Math.max(8, (typeof font?.size === "number" ? font.size * 4 / 3 : 12) * zoomScale.value / 100),
    fontStyle: font?.italic ? "italic" : "normal",
    fontWeight: font?.bold ? "700" : "400",
    strike: Boolean(font?.strikethrough),
    underline: Boolean(font?.underline && font.underline !== "none"),
    vertical: alignment?.vertical === "top" || alignment?.vertical === "center"
      ? alignment.vertical === "center" ? "middle" : "top"
      : "bottom",
    wrap: Boolean(alignment?.wrapText),
  };
  const resolvedStyle: Record<string, string | number | undefined> = {
    backgroundColor: base.backgroundColor,
    color: base.color,
    fontFamily: base.fontFamily,
    fontSize: base.fontSizePx,
    fontStyle: base.fontStyle,
    fontWeight: base.fontWeight,
    textAlign: base.alignment,
    whiteSpace: base.wrap ? "pre-wrap" : "nowrap",
  };
  const context: XlsxCellStyleContext = {
    cell,
    hasChartHighlight: false,
    hasConditionalFormat: conditionalRulesForCell(cell).length > 0,
    hasHyperlink: hyperlinksByCell.value.has(`${cell.row}:${cell.col}`),
    hasValidation: false,
    isMerged: merged,
    isTableHeader: tableHeader,
    resolvedStyle,
    sheetName: sheet?.name ?? "",
    value,
    workbookSheetIndex: sheet?.workbookSheetIndex ?? 0,
  };
  const override = props.getCellStyle?.(cell, context) ?? undefined;
  if (!override) return base;
  const background = cssText(override.backgroundColor ?? override.background);
  const color = cssText(override.color);
  const fontFamily = cssText(override.fontFamily);
  const fontSize = Number.parseFloat(cssText(override.fontSize) ?? "");
  const textAlign = cssText(override.textAlign);
  const verticalAlign = cssText(override.verticalAlign);
  return {
    ...base,
    ...(background ? { backgroundColor: background } : {}),
    ...(color ? { color } : {}),
    ...(fontFamily ? { fontFamily } : {}),
    ...(Number.isFinite(fontSize) ? { fontSizePx: fontSize } : {}),
    ...(override.fontStyle ? { fontStyle: String(override.fontStyle) } : {}),
    ...(override.fontWeight ? { fontWeight: String(override.fontWeight) } : {}),
    ...(textAlign === "center" || textAlign === "right" || textAlign === "left" ? { alignment: textAlign } : {}),
    ...(verticalAlign === "top" || verticalAlign === "middle" || verticalAlign === "bottom" ? { vertical: verticalAlign } : {}),
    ...(override.whiteSpace === "pre-wrap" || override.whiteSpace === "normal" ? { wrap: true } : {}),
  };
}

function borderWidth(style: unknown): number {
  if (style === "medium" || style === "mediumDashed" || style === "mediumDashDot" || style === "mediumDashDotDot") return 2;
  if (style === "thick" || style === "double") return 3;
  return 1;
}

function drawCellBorders(
  ctx: CanvasRenderingContext2D,
  style: CanvasCellStyle,
  x: number,
  y: number,
  width: number,
  height: number,
): void {
  if (!style.border) return;
  const edges = [
    ["top", x, y, x + width, y],
    ["right", x + width, y, x + width, y + height],
    ["bottom", x, y + height, x + width, y + height],
    ["left", x, y, x, y + height],
  ] as const;
  for (const [name, x1, y1, x2, y2] of edges) {
    const edge = style.border[name];
    if (!edge?.style || edge.style === "none") continue;
    ctx.save();
    ctx.strokeStyle = resolveWorkbookColor(asStyleRecord(edge.color), activeSheet.value?.themePalette) || (props.isDark ? "#71717a" : "#52525b");
    ctx.lineWidth = borderWidth(edge.style);
    ctx.setLineDash(edge.style === "dashed" || String(edge.style).includes("Dash") ? [5, 3] : edge.style === "dotted" ? [1, 2] : []);
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    ctx.restore();
  }
}

function wrapCellText(ctx: CanvasRenderingContext2D, value: string, width: number, maxLines: number): string[] {
  if (maxLines <= 1) return [truncateText(ctx, value, width)];
  const lines: string[] = [];
  let line = "";
  for (const character of value) {
    const next = line + character;
    if (line && ctx.measureText(next).width > width) {
      lines.push(line);
      line = character;
      if (lines.length >= maxLines - 1) break;
    } else {
      line = next;
    }
  }
  if (line && lines.length < maxLines) lines.push(line);
  return lines;
}

function paintGrid() {
  paintBody();
  paintColHeaders();
  paintRowHeaders();
  paintCorner();
}

function paintBody() {
  const canvas = bodyCanvasRef.value;
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  const bw = Math.max(0, containerWidth.value - ROW_HEADER_WIDTH);
  const bh = Math.max(0, containerHeight.value - HEADER_HEIGHT);
  resizeCanvas(canvas, bw, bh);

  ctx.clearRect(0, 0, bw, bh);
  const z = zoomScale.value / 100;
  const rowRange = visibleRowRange.value;
  const colRange = visibleColRange.value;
  const sheet = activeSheet.value;
  if (!sheet) return;
  conditionalRuleStats = new Map();

  const widths = effectiveColWidths.value;
  const heights = effectiveRowHeights.value;
  const worksheet = props.controller.isWorkerBacked ? null : props.controller.getActiveWorksheet();
  updatePaintedMergeRegions(props.controller.isWorkerBacked ? sheet.mergedRegions : worksheet?.mergedRegions);

  const bgColor = props.isDark ? "#18181b" : "#ffffff";
  const gridColor = props.isDark ? "#3f3f46" : "#d9d9d9";
  const textColor = props.isDark ? "#e4e4e7" : "#18181b";
  const selFill = props.selectionFillColor || (props.isDark ? "rgba(37, 99, 235, 0.15)" : "rgba(37, 99, 235, 0.08)");
  const selBorder = props.selectionColor || "#2563eb";

  ctx.fillStyle = bgColor;
  ctx.fillRect(0, 0, bw, bh);
  if (rowRange.end < rowRange.start || colRange.end < colRange.start) return;

  ctx.save();
  ctx.translate(-scrollLeft.value, -scrollTop.value);

  // Gridlines
  if (sheet.showGridLines !== false) {
    ctx.strokeStyle = gridColor;
    ctx.lineWidth = 1;
    const gridTop = getRowOffsetSum(rowRange.start);
    const gridBottom = getRowOffsetSum(rowRange.end + 1);
    const gridLeft = getColOffsetSum(colRange.start);
    const gridRight = getColOffsetSum(colRange.end + 1);
    for (let c = colRange.start; c <= colRange.end + 1 && c <= widths.length; c += 1) {
      const x = getColOffsetSum(c);
      ctx.beginPath(); ctx.moveTo(x, gridTop); ctx.lineTo(x, gridBottom); ctx.stroke();
    }
    for (let r = rowRange.start; r <= rowRange.end + 1 && r <= heights.length; r += 1) {
      const y = getRowOffsetSum(r);
      ctx.beginPath(); ctx.moveTo(gridLeft, y); ctx.lineTo(gridRight, y); ctx.stroke();
    }
  }

  const sel = selection.value;
  const active = activeCell.value;

  try {
    for (let r = rowRange.start; r <= rowRange.end && r < heights.length; r++) {
      const rowH = heights[r] || 0;
      const rowY = getRowOffsetSum(r);
      for (let c = colRange.start; c <= colRange.end && c < widths.length; c++) {
        const cell = getSheetCell(r, c);
        if (!cell || paintedMergeAnchors.has(`${cell.row}:${cell.col}`)) continue;
        const merge = mergeRegionAtAnchor(cell);
        const mergeEnd = merge ? getDisplayCell(merge.end) : null;
        const colW = mergeEnd
          ? getColOffsetSum(mergeEnd.col + 1) - getColOffsetSum(c)
          : widths[c] || 0;
        const rowHResolved = mergeEnd
          ? getRowOffsetSum(mergeEnd.row + 1) - getRowOffsetSum(r)
          : rowH;
        const colX = getColOffsetSum(c);
        const value = getGridCellDisplayValue(cell);
        let workbookStyle = getWorkerCellStyle(cell);
        if (!workbookStyle && worksheet) {
          try {
            workbookStyle = worksheet.getCellStyleAt(cell.row, cell.col) as XlsxResolvedCellStyle | undefined;
          } catch {
            workbookStyle = undefined;
          }
        }
        if (!workbookStyle) {
          const rowStyleId = sheet.rowStyleIds[cell.row];
          const colStyleId = sheet.colStyleIds[cell.col];
          workbookStyle = rowStyleId !== undefined
            ? sheet.styleById[rowStyleId]
            : colStyleId !== undefined ? sheet.styleById[colStyleId] : undefined;
        }
        const style = resolveCanvasCellStyle(cell, value, workbookStyle);
        const conditional = conditionalVisual(cell, value);
        const hyperlink = hyperlinksByCell.value.get(`${cell.row}:${cell.col}`);
        const comment = commentsByCell.value.get(`${cell.row}:${cell.col}`);
        const sparkline = sparklinesByCell.value.get(`${cell.row}:${cell.col}`);

        ctx.fillStyle = conditional?.kind === "colorScale" ? conditional.background : style.backgroundColor;
        ctx.fillRect(colX + 0.5, rowY + 0.5, Math.max(0, colW - 1), Math.max(0, rowHResolved - 1));
        if (conditional?.kind === "dataBar") {
          ctx.fillStyle = `${conditional.color}99`;
          ctx.fillRect(colX + 3, rowY + 4, Math.max(2, (colW - 6) * conditional.ratio), Math.max(2, rowHResolved - 8));
        }
        if (sparkline) drawSparkline(ctx, sparkline, colX, rowY, colW, rowHResolved);
        if (merge && sheet.showGridLines !== false) {
          ctx.strokeStyle = gridColor;
          ctx.lineWidth = 1;
          ctx.strokeRect(colX + 0.5, rowY + 0.5, Math.max(0, colW - 1), Math.max(0, rowHResolved - 1));
        }

        if (isCellInRange(cell, sel) || (merge && sel && rangesIntersect(merge, sel))) {
          ctx.fillStyle = selFill;
          ctx.fillRect(colX + 1, rowY + 1, Math.max(0, colW - 2), Math.max(0, rowHResolved - 2));
        }

        if (value && !sparkline && (conditional?.kind !== "dataBar" || conditional.showValue)) {
          ctx.save();
          ctx.beginPath();
          ctx.rect(colX + 2, rowY + 1, Math.max(0, colW - 4), Math.max(0, rowHResolved - 2));
          ctx.clip();
          ctx.fillStyle = hyperlink ? "#2563eb" : style.color || textColor;
          ctx.font = `${style.fontStyle} ${style.fontWeight} ${Math.round(style.fontSizePx)}px ${JSON.stringify(style.fontFamily)}, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif`;
          ctx.textAlign = style.alignment;
          const padding = 5;
          const textX = style.alignment === "right"
            ? colX + colW - padding
            : style.alignment === "center" ? colX + colW / 2 : colX + padding;
          const lineHeight = Math.max(12, style.fontSizePx * 1.25);
          const maxLines = style.wrap ? Math.max(1, Math.floor((rowHResolved - 4) / lineHeight)) : 1;
          const lines = wrapCellText(ctx, value, Math.max(0, colW - padding * 2), maxLines);
          const blockHeight = lines.length * lineHeight;
          const firstBaseline = style.vertical === "top"
            ? rowY + 2 + lineHeight * 0.8
            : style.vertical === "middle"
              ? rowY + (rowHResolved - blockHeight) / 2 + lineHeight * 0.8
              : rowY + rowHResolved - blockHeight + lineHeight * 0.8 - 2;
          lines.forEach((line, index) => ctx.fillText(line, textX, firstBaseline + index * lineHeight));
          if (style.underline || style.strike || hyperlink) {
            const measured = Math.min(ctx.measureText(lines[0] ?? "").width, Math.max(0, colW - padding * 2));
            const startX = style.alignment === "right" ? textX - measured : style.alignment === "center" ? textX - measured / 2 : textX;
            ctx.strokeStyle = style.color;
            ctx.lineWidth = 1;
            for (const offset of [style.underline || hyperlink ? 2 : null, style.strike ? -style.fontSizePx * 0.3 : null]) {
              if (offset === null) continue;
              ctx.beginPath();
              ctx.moveTo(startX, firstBaseline + offset);
              ctx.lineTo(startX + measured, firstBaseline + offset);
              ctx.stroke();
            }
          }
          ctx.restore();
        }
        if (conditional?.kind === "iconSet") {
          ctx.fillStyle = conditional.ratio < .34 ? "#dc2626" : conditional.ratio < .67 ? "#d97706" : "#16a34a";
          ctx.font = `${Math.max(10, Math.min(15, rowHResolved - 6))}px sans-serif`;
          ctx.textAlign = "right";
          ctx.fillText(conditional.icon, colX + colW - 4, rowY + rowHResolved / 2 + 5);
        }
        if (comment) {
          ctx.fillStyle = "#dc2626";
          ctx.beginPath(); ctx.moveTo(colX + colW - 9, rowY + 1); ctx.lineTo(colX + colW - 1, rowY + 1); ctx.lineTo(colX + colW - 1, rowY + 9); ctx.closePath(); ctx.fill();
        }

        drawCellBorders(ctx, style, colX, rowY, colW, rowHResolved);
      }
    }

    if (sel) drawSelectionBorder(ctx, sel, selBorder);
    if (active && (!sel || (sel.start.row === sel.end.row && sel.start.col === sel.end.col))) {
      drawSelectionBorder(ctx, { start: active, end: active }, selBorder);
    }
  } finally {
    worksheet?.free();
    ctx.restore();
  }
  paintFrozenPanes(ctx, bw, bh, gridColor, selFill, selBorder);
}

function paintFrozenPanes(
  ctx: CanvasRenderingContext2D,
  bodyWidth: number,
  bodyHeight: number,
  gridColor: string,
  selectionFill: string,
  selectionBorder: string,
): void {
  const sheet = activeSheet.value;
  const freezeRow = sheet?.freezePanes?.row ?? 0;
  const freezeCol = sheet?.freezePanes?.col ?? 0;
  if (!sheet || (freezeRow <= 0 && freezeCol <= 0)) return;
  const frozenRows = displayRows.value
    .map((actualIndex, displayIndex) => ({ actualIndex, displayIndex }))
    .filter((entry) => entry.actualIndex < freezeRow);
  const frozenCols = displayCols.value
    .map((actualIndex, displayIndex) => ({ actualIndex, displayIndex }))
    .filter((entry) => entry.actualIndex < freezeCol);
  const frozenHeight = frozenRows.reduce((sum, entry) => sum + (effectiveRowHeights.value[entry.displayIndex] || 0), 0);
  const frozenWidth = frozenCols.reduce((sum, entry) => sum + (effectiveColWidths.value[entry.displayIndex] || 0), 0);
  const visibleRows = Array.from(
    { length: Math.max(0, visibleRowRange.value.end - visibleRowRange.value.start + 1) },
    (_, offset) => visibleRowRange.value.start + offset,
  ).filter((index) => (displayRows.value[index] ?? 0) >= freezeRow);
  const visibleCols = Array.from(
    { length: Math.max(0, visibleColRange.value.end - visibleColRange.value.start + 1) },
    (_, offset) => visibleColRange.value.start + offset,
  ).filter((index) => (displayCols.value[index] ?? 0) >= freezeCol);
  const worksheet = props.controller.isWorkerBacked ? null : props.controller.getActiveWorksheet();

  const paintCell = (displayRow: number, displayCol: number, frozenY: boolean, frozenX: boolean) => {
    const cell = getSheetCell(displayRow, displayCol);
    if (!cell || paintedMergeAnchors.has(`${cell.row}:${cell.col}`)) return;
    const width = effectiveColWidths.value[displayCol] || 0;
    const height = effectiveRowHeights.value[displayRow] || 0;
    const x = getColOffsetSum(displayCol) - (frozenX ? 0 : scrollLeft.value);
    const y = getRowOffsetSum(displayRow) - (frozenY ? 0 : scrollTop.value);
    const value = getGridCellDisplayValue(cell);
    let workbookStyle = getWorkerCellStyle(cell);
    if (!workbookStyle && worksheet) {
      try { workbookStyle = worksheet.getCellStyleAt(cell.row, cell.col) as XlsxResolvedCellStyle | undefined; } catch { /* keep fallback */ }
    }
    const style = resolveCanvasCellStyle(cell, value, workbookStyle);
    ctx.fillStyle = style.backgroundColor;
    ctx.fillRect(x, y, width, height);
    if (sheet.showGridLines !== false) {
      ctx.strokeStyle = gridColor;
      ctx.lineWidth = 1;
      ctx.strokeRect(x + 0.5, y + 0.5, Math.max(0, width - 1), Math.max(0, height - 1));
    }
    if (isCellInRange(cell, selection.value)) {
      ctx.fillStyle = selectionFill;
      ctx.fillRect(x + 1, y + 1, Math.max(0, width - 2), Math.max(0, height - 2));
    }
    if (value) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(x + 2, y + 1, Math.max(0, width - 4), Math.max(0, height - 2));
      ctx.clip();
      ctx.fillStyle = style.color;
      ctx.font = `${style.fontStyle} ${style.fontWeight} ${Math.round(style.fontSizePx)}px ${JSON.stringify(style.fontFamily)}, "Segoe UI", sans-serif`;
      ctx.textBaseline = "middle";
      ctx.textAlign = style.alignment;
      const textX = style.alignment === "right" ? x + width - 5 : style.alignment === "center" ? x + width / 2 : x + 5;
      ctx.fillText(truncateText(ctx, value, Math.max(0, width - 10)), textX, y + height / 2);
      ctx.restore();
    }
    drawCellBorders(ctx, style, x, y, width, height);
    if (activeCell.value?.row === cell.row && activeCell.value?.col === cell.col) {
      ctx.strokeStyle = selectionBorder;
      ctx.lineWidth = 2;
      ctx.strokeRect(x + 1, y + 1, Math.max(0, width - 2), Math.max(0, height - 2));
    }
  };

  try {
    if (frozenRows.length && visibleCols.length) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(frozenWidth, 0, Math.max(0, bodyWidth - frozenWidth), frozenHeight);
      ctx.clip();
      for (const row of frozenRows) for (const col of visibleCols) paintCell(row.displayIndex, col, true, false);
      ctx.restore();
    }
    if (frozenCols.length && visibleRows.length) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(0, frozenHeight, frozenWidth, Math.max(0, bodyHeight - frozenHeight));
      ctx.clip();
      for (const row of visibleRows) for (const col of frozenCols) paintCell(row, col.displayIndex, false, true);
      ctx.restore();
    }
    if (frozenRows.length && frozenCols.length) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(0, 0, frozenWidth, frozenHeight);
      ctx.clip();
      for (const row of frozenRows) for (const col of frozenCols) paintCell(row.displayIndex, col.displayIndex, true, true);
      ctx.restore();
    }
    ctx.save();
    ctx.strokeStyle = props.isDark ? "#71717a" : "#a1a1aa";
    ctx.lineWidth = 2;
    if (frozenWidth > 0) { ctx.beginPath(); ctx.moveTo(frozenWidth, 0); ctx.lineTo(frozenWidth, bodyHeight); ctx.stroke(); }
    if (frozenHeight > 0) { ctx.beginPath(); ctx.moveTo(0, frozenHeight); ctx.lineTo(bodyWidth, frozenHeight); ctx.stroke(); }
    ctx.restore();
  } finally {
    worksheet?.free();
  }
}

function paintColHeaders() {
  const canvas = colHeaderCanvasRef.value;
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  const cw = Math.max(0, containerWidth.value - ROW_HEADER_WIDTH);
  resizeCanvas(canvas, cw, HEADER_HEIGHT);
  ctx.clearRect(0, 0, cw, HEADER_HEIGHT);

  const bgColor = props.isDark ? "#27272a" : "#f4f4f5";
  const textColor = props.isDark ? "#a1a1aa" : "#71717a";
  const borderColor = props.isDark ? "#3f3f46" : "#d4d4d8";

  ctx.fillStyle = bgColor;
  ctx.fillRect(0, 0, cw, HEADER_HEIGHT);
  ctx.save();
  ctx.translate(-scrollLeft.value, 0);
  ctx.font = "11px -apple-system, BlinkMacSystemFont, sans-serif";
  ctx.textBaseline = "middle";
  ctx.textAlign = "center";
  ctx.fillStyle = textColor;

  const widths = effectiveColWidths.value;
  const colRange = visibleColRange.value;
  const selected = selection.value;
  for (let c = colRange.start; c <= colRange.end && c < widths.length; c++) {
    const colW = widths[c] || 0;
    const x = getColOffsetSum(c);
    const sheetCol = displayCols.value[c] ?? c;
    const selectedCol = activeCell.value?.col === sheetCol || Boolean(selected &&
      sheetCol >= Math.min(selected.start.col, selected.end.col) &&
      sheetCol <= Math.max(selected.start.col, selected.end.col));
    if (selectedCol) {
      ctx.fillStyle = props.isDark ? "#1e3a5f" : "#dbeafe";
      ctx.fillRect(x, 0, colW, HEADER_HEIGHT);
    }
    ctx.fillStyle = selectedCol ? (props.isDark ? "#bfdbfe" : "#1d4ed8") : textColor;
    ctx.fillText(columnLabel(sheetCol), x + colW / 2, HEADER_HEIGHT / 2);
    ctx.strokeStyle = borderColor; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(x + colW, 0); ctx.lineTo(x + colW, HEADER_HEIGHT); ctx.stroke();
  }
  ctx.strokeStyle = borderColor;
  ctx.beginPath();
  ctx.moveTo(getColOffsetSum(colRange.start), HEADER_HEIGHT - 0.5);
  ctx.lineTo(getColOffsetSum(colRange.end + 1), HEADER_HEIGHT - 0.5);
  ctx.stroke();
  ctx.restore();
}

function paintRowHeaders() {
  const canvas = rowHeaderCanvasRef.value;
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  const rh = Math.max(0, containerHeight.value - HEADER_HEIGHT);
  resizeCanvas(canvas, ROW_HEADER_WIDTH, rh);
  ctx.clearRect(0, 0, ROW_HEADER_WIDTH, rh);

  const bgColor = props.isDark ? "#27272a" : "#f4f4f5";
  const textColor = props.isDark ? "#a1a1aa" : "#71717a";
  const borderColor = props.isDark ? "#3f3f46" : "#d4d4d8";

  ctx.fillStyle = bgColor;
  ctx.fillRect(0, 0, ROW_HEADER_WIDTH, rh);
  ctx.save();
  ctx.translate(0, -scrollTop.value);
  ctx.font = "11px -apple-system, BlinkMacSystemFont, sans-serif";
  ctx.textBaseline = "middle";
  ctx.textAlign = "center";
  ctx.fillStyle = textColor;

  const heights = effectiveRowHeights.value;
  const rowRange = visibleRowRange.value;
  const selected = selection.value;
  for (let r = rowRange.start; r <= rowRange.end && r < heights.length; r++) {
    const rowH = heights[r] || 0;
    const y = getRowOffsetSum(r);
    const sheetRow = displayRows.value[r] ?? r;
    const selectedRow = activeCell.value?.row === sheetRow || Boolean(selected &&
      sheetRow >= Math.min(selected.start.row, selected.end.row) &&
      sheetRow <= Math.max(selected.start.row, selected.end.row));
    if (selectedRow) {
      ctx.fillStyle = props.isDark ? "#1e3a5f" : "#dbeafe";
      ctx.fillRect(0, y, ROW_HEADER_WIDTH, rowH);
    }
    ctx.fillStyle = selectedRow ? (props.isDark ? "#bfdbfe" : "#1d4ed8") : textColor;
    ctx.fillText(String(sheetRow + 1), ROW_HEADER_WIDTH / 2, y + rowH / 2);
    ctx.strokeStyle = borderColor; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(0, y + rowH); ctx.lineTo(ROW_HEADER_WIDTH, y + rowH); ctx.stroke();
  }
  ctx.strokeStyle = borderColor;
  ctx.beginPath();
  ctx.moveTo(ROW_HEADER_WIDTH - 0.5, getRowOffsetSum(rowRange.start));
  ctx.lineTo(ROW_HEADER_WIDTH - 0.5, getRowOffsetSum(rowRange.end + 1));
  ctx.stroke();
  ctx.restore();
}

function paintCorner() {
  const canvas = cornerCanvasRef.value;
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  resizeCanvas(canvas, ROW_HEADER_WIDTH, HEADER_HEIGHT);
  const bgColor = props.isDark ? "#27272a" : "#f4f4f5";
  const borderColor = props.isDark ? "#3f3f46" : "#d4d4d8";

  ctx.fillStyle = bgColor;
  ctx.fillRect(0, 0, ROW_HEADER_WIDTH, HEADER_HEIGHT);
  ctx.strokeStyle = borderColor; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(ROW_HEADER_WIDTH - 0.5, 0); ctx.lineTo(ROW_HEADER_WIDTH - 0.5, HEADER_HEIGHT); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(0, HEADER_HEIGHT - 0.5); ctx.lineTo(ROW_HEADER_WIDTH, HEADER_HEIGHT - 0.5); ctx.stroke();
}

// ── Event handlers ────────────────────────────────────────────────────
let paintFrame: number | null = null;

function schedulePaint() {
  if (paintFrame !== null) return;
  paintFrame = window.requestAnimationFrame(() => {
    paintFrame = null;
    const container = containerRef.value;
    if (!container) return;
    scrollTop.value = container.scrollTop;
    scrollLeft.value = container.scrollLeft;
    emit("viewportChange", { scrollLeft: scrollLeft.value, scrollTop: scrollTop.value });
    paintGrid();
    hydrateVisibleWorkerRows();
  });
}

function onScroll() {
  schedulePaint();
}

type AxisHit = {
  actualIndex: number;
  displayIndex: number;
  kind: "column" | "row" | "corner";
  resize: boolean;
};

function hitTestAxis(clientX: number, clientY: number): AxisHit | null {
  const container = containerRef.value;
  if (!container) return null;
  const rect = container.getBoundingClientRect();
  const localX = clientX - rect.left;
  const localY = clientY - rect.top;
  if (localX < 0 || localY < 0 || localX > rect.width || localY > rect.height) return null;
  if (localX <= ROW_HEADER_WIDTH && localY <= HEADER_HEIGHT) {
    return { actualIndex: 0, displayIndex: 0, kind: "corner", resize: false };
  }
  if (localY <= HEADER_HEIGHT && localX > ROW_HEADER_WIDTH) {
    const contentX = localX + container.scrollLeft - ROW_HEADER_WIDTH;
    const hitDisplayIndex = findDisplayIndexAtOffset(colOffsets.value, contentX);
    if (hitDisplayIndex < 0) return null;
    const startBoundary = getColOffsetSum(hitDisplayIndex);
    const usePreviousBoundary = hitDisplayIndex > 0 && Math.abs(contentX - startBoundary) <= 5;
    const displayIndex = usePreviousBoundary ? hitDisplayIndex - 1 : hitDisplayIndex;
    const actualIndex = displayCols.value[displayIndex];
    if (actualIndex === undefined) return null;
    const boundary = usePreviousBoundary ? startBoundary : getColOffsetSum(displayIndex + 1);
    return {
      actualIndex,
      displayIndex,
      kind: "column",
      resize: Math.abs(contentX - boundary) <= 5,
    };
  }
  if (localX <= ROW_HEADER_WIDTH && localY > HEADER_HEIGHT) {
    const contentY = localY + container.scrollTop - HEADER_HEIGHT;
    const hitDisplayIndex = findDisplayIndexAtOffset(rowOffsets.value, contentY);
    if (hitDisplayIndex < 0) return null;
    const startBoundary = getRowOffsetSum(hitDisplayIndex);
    const usePreviousBoundary = hitDisplayIndex > 0 && Math.abs(contentY - startBoundary) <= 5;
    const displayIndex = usePreviousBoundary ? hitDisplayIndex - 1 : hitDisplayIndex;
    const actualIndex = displayRows.value[displayIndex];
    if (actualIndex === undefined) return null;
    const boundary = usePreviousBoundary ? startBoundary : getRowOffsetSum(displayIndex + 1);
    return {
      actualIndex,
      displayIndex,
      kind: "row",
      resize: Math.abs(contentY - boundary) <= 5,
    };
  }
  return null;
}

function selectAxisRange(axisHit: Pick<AxisHit, "actualIndex" | "kind">) {
  const firstRow = displayRows.value[0] ?? 0;
  const lastRow = displayRows.value[displayRows.value.length - 1] ?? firstRow;
  const firstCol = displayCols.value[0] ?? 0;
  const lastCol = displayCols.value[displayCols.value.length - 1] ?? firstCol;
  if (axisHit.kind === "column") {
    props.controller.selectRange({
      start: { row: firstRow, col: axisHit.actualIndex },
      end: { row: lastRow, col: axisHit.actualIndex },
    });
  } else if (axisHit.kind === "row") {
    props.controller.selectRange({
      start: { row: axisHit.actualIndex, col: firstCol },
      end: { row: axisHit.actualIndex, col: lastCol },
    });
  } else {
    props.controller.selectRange({
      start: { row: firstRow, col: firstCol },
      end: { row: lastRow, col: lastCol },
    });
  }
}

function onPointerDown(event: PointerEvent) {
  // Only left-click (button 0) should change selection.
  // Right-click (button 2) must preserve the existing selection
  // so the context menu operates on the intended range.
  if (event.button !== 0) return
  const axisHit = hitTestAxis(event.clientX, event.clientY);
  if (axisHit?.resize && axisHit.kind !== "corner") {
    const z = Math.max(0.5, zoomScale.value / 100);
    const renderedSize = axisHit.kind === "column"
      ? effectiveColWidths.value[axisHit.displayIndex] ?? DEFAULT_COL_WIDTH * z
      : effectiveRowHeights.value[axisHit.displayIndex] ?? DEFAULT_ROW_HEIGHT * z;
    resizePreview.value = {
      actualIndex: axisHit.actualIndex,
      displayIndex: axisHit.displayIndex,
      kind: axisHit.kind,
      moved: false,
      pointerId: event.pointerId,
      size: renderedSize / z,
      startPosition: axisHit.kind === "column" ? event.clientX : event.clientY,
      startSize: renderedSize / z,
    };
    containerRef.value?.focus();
    containerRef.value?.setPointerCapture(event.pointerId);
    event.preventDefault();
    return;
  }
  if (axisHit) {
    selectAxisRange(axisHit);
    containerRef.value?.focus();
    event.preventDefault();
    return;
  }
  const cell = hitTestCell(event.clientX, event.clientY);
  if (!cell) return;
  containerRef.value?.focus();

  if (event.shiftKey && activeCell.value) {
    props.controller.selectRange({ start: activeCell.value, end: cell });
  } else {
    props.controller.selectCell(cell, { extend: event.shiftKey });
  }
  isSelecting.value = true;
  selectionAnchor.value = cell;
  containerRef.value?.setPointerCapture(event.pointerId);
}

function onPointerMove(event: PointerEvent) {
  const resizing = resizePreview.value;
  if (resizing) {
    const z = Math.max(0.5, zoomScale.value / 100);
    const position = resizing.kind === "column" ? event.clientX : event.clientY;
    const delta = position - resizing.startPosition;
    if (!resizing.moved && Math.abs(delta) < AXIS_RESIZE_DRAG_THRESHOLD) return;
    const minimum = resizing.kind === "column" ? 32 : 18;
    resizing.moved = true;
    resizing.size = Math.max(minimum, resizing.startSize + delta / z);
    resizePreview.value = { ...resizing };
    schedulePaint();
    return;
  }
  const axisHit = hitTestAxis(event.clientX, event.clientY);
  hoveredCell.value = axisHit ? null : hitTestCell(event.clientX, event.clientY);
  if (containerRef.value) {
    containerRef.value.style.cursor = axisHit?.resize
      ? axisHit.kind === "column" ? "col-resize" : "row-resize"
      : axisHit ? "default" : hoveredCell.value && hyperlinksByCell.value.has(`${hoveredCell.value.row}:${hoveredCell.value.col}`) ? "pointer" : "cell";
  }
  if (!isSelecting.value || !selectionAnchor.value) return;
  const cell = hitTestCell(event.clientX, event.clientY);
  if (!cell) return;
  props.controller.selectRange({ start: selectionAnchor.value, end: cell });
}

function resolveAxisResizeRelease(
  moved: boolean,
  eventType?: string,
): "cancel" | "resize" | "select" {
  if (eventType === "pointercancel") return "cancel";
  return moved ? "resize" : "select";
}

function onPointerUp(event?: PointerEvent) {
  const clickedCell = event ? hitTestCell(event.clientX, event.clientY) : null;
  const pressCell = selectionAnchor.value;
  const resizing = resizePreview.value;
  if (resizing) {
    const releaseAction = resolveAxisResizeRelease(resizing.moved, event?.type);
    if (releaseAction === "resize") {
      if (resizing.kind === "column") props.controller.resizeColumn(resizing.actualIndex, resizing.size);
      else props.controller.resizeRow(resizing.actualIndex, resizing.size);
    } else if (releaseAction === "select") {
      selectAxisRange(resizing);
    }
    resizePreview.value = null;
  }
  isSelecting.value = false;
  selectionAnchor.value = null;
  if (event && containerRef.value?.hasPointerCapture(event.pointerId)) {
    containerRef.value.releasePointerCapture(event.pointerId);
  }
  schedulePaint();
  if (clickedCell && pressCell && clickedCell.row === pressCell.row && clickedCell.col === pressCell.col) {
    activateHyperlink(clickedCell);
  }
}

function onPointerLeave(): void {
  hoveredCell.value = null;
  if (!resizePreview.value && !isSelecting.value && containerRef.value) {
    containerRef.value.style.cursor = "cell";
  }
}

function activateHyperlink(cell: XlsxCellAddress): void {
  const hyperlink = hyperlinksByCell.value.get(`${cell.row}:${cell.col}`) as XlsxCellHyperlink | undefined;
  if (!hyperlink) return;
  if (hyperlink.target.startsWith("#")) {
    const reference = hyperlink.target.slice(1);
    const [sheetPart, cellPart = sheetPart] = reference.split("!");
    const sheetName = cellPart === sheetPart ? null : sheetPart.replace(/^'|'$/g, "").replace(/''/g, "'");
    const match = /\$?([A-Z]+)\$?(\d+)/i.exec(cellPart);
    if (!match) return;
    if (sheetName) {
      const sheetIndex = props.controller.sheets.findIndex((sheet) => sheet.name === sheetName);
      if (sheetIndex >= 0) props.controller.setActiveSheetIndex(sheetIndex);
    }
    let col = 0;
    for (const char of match[1].toUpperCase()) col = col * 26 + char.charCodeAt(0) - 64;
    props.controller.selectCell({ row: Number(match[2]) - 1, col: col - 1 });
    return;
  }
  try {
    const url = new URL(hyperlink.target, globalThis.location?.href);
    if (!["http:", "https:", "mailto:"].includes(url.protocol)) return;
    globalThis.open?.(url.href, "_blank", "noopener,noreferrer");
  } catch { /* invalid links stay inert */ }
}

function onDblClick(event: MouseEvent) {
  const cell = hitTestCell(event.clientX, event.clientY);
  if (!cell || props.readOnly) return;
  editingCell.value = cell;
  editingValue.value = props.controller.getCellDisplayValue(cell);
  emit("cellDoubleClick", cell);
  nextTick(() => { editInputRef.value?.focus(); editInputRef.value?.select(); });
}

function commitEdit() {
  if (!editingCell.value) return;
  props.controller.setCellValue(editingCell.value, editingValue.value);
  editingCell.value = null;
  editingValue.value = "";
  void nextTick(() => containerRef.value?.focus());
}

function cancelEdit() {
  editingCell.value = null;
  editingValue.value = "";
  void nextTick(() => containerRef.value?.focus());
}

function commitAndTab() {
  if (!editingCell.value) return;
  props.controller.setCellValue(editingCell.value, editingValue.value);
  const nextCell = getAdjacentCell(editingCell.value, 0, 1);
  if (!nextCell) return;
  props.controller.selectCell(nextCell);
  editingCell.value = nextCell;
  editingValue.value = props.controller.getCellDisplayValue(nextCell);
}

function getAdjacentCell(cell: XlsxCellAddress, rowDelta: number, colDelta: number): XlsxCellAddress | null {
  return resolveAdjacentSheetCell(
    cell,
    rowDelta,
    colDelta,
    displayRows.value,
    displayCols.value,
    sheetRowToDisplayIndex.value,
    sheetColToDisplayIndex.value,
  );
}

function moveCell(rowDelta: number, colDelta: number, extend: boolean) {
  const current = activeCell.value;
  if (!current) return;
  const cell = getAdjacentCell(current, rowDelta, colDelta);
  if (!cell) return;
  if (extend) {
    props.controller.selectRange({ start: current, end: cell });
  } else {
    props.controller.selectCell(cell);
  }
  scrollToCell(cell);
}

function startEdit(cell: XlsxCellAddress) {
  editingCell.value = cell;
  editingValue.value = props.controller.getCellDisplayValue(cell);
  nextTick(() => editInputRef.value?.focus());
}

function scrollToCell(cell: XlsxCellAddress) {
  const container = containerRef.value;
  if (!container) return;
  const displayCell = getDisplayCell(cell);
  if (!displayCell) return;
  const colX = getColOffsetSum(displayCell.col);
  const rowY = getRowOffsetSum(displayCell.row);
  const colW = effectiveColWidths.value[displayCell.col] || DEFAULT_COL_WIDTH;
  const rowH = effectiveRowHeights.value[displayCell.row] || DEFAULT_ROW_HEIGHT;
  const bodyWidth = Math.max(0, container.clientWidth - ROW_HEADER_WIDTH);
  const bodyHeight = Math.max(0, container.clientHeight - HEADER_HEIGHT);

  if (colX < container.scrollLeft) container.scrollLeft = Math.max(0, colX);
  else if (colX + colW > container.scrollLeft + bodyWidth) container.scrollLeft = colX + colW - bodyWidth;
  if (rowY < container.scrollTop) container.scrollTop = Math.max(0, rowY);
  else if (rowY + rowH > container.scrollTop + bodyHeight) container.scrollTop = rowY + rowH - bodyHeight;
  schedulePaint();
}

function onGridKeydown(event: KeyboardEvent) {
  // When the inline edit input is active, let it consume all keys —
  // do not re-enter edit mode and overwrite the editing value.
  if (editingCell.value) return

  const active = activeCell.value;
  if (!active) return;

  if ((event.ctrlKey || event.metaKey) && event.key === "c") { event.preventDefault(); props.controller.copySelectionToClipboard(); return; }
  if ((event.ctrlKey || event.metaKey) && event.key === "v") { event.preventDefault(); props.controller.pasteFromClipboard(); return; }

  const handlers: Record<string, () => void> = {
    ArrowUp: () => moveCell(-1, 0, event.shiftKey),
    ArrowDown: () => moveCell(1, 0, event.shiftKey),
    ArrowLeft: () => moveCell(0, -1, event.shiftKey),
    ArrowRight: () => moveCell(0, 1, event.shiftKey),
    Tab: () => { event.preventDefault(); moveCell(0, 1, false); },
    Enter: () => { if (!props.readOnly) startEdit(active); },
    Delete: () => { if (!props.readOnly) props.controller.clearSelectedCells(); },
    Backspace: () => { if (!props.readOnly) props.controller.setCellValue(active, ""); },
    F2: () => { if (!props.readOnly) startEdit(active); },
  };

  const handler = handlers[event.key];
  if (handler) { event.preventDefault(); handler(); return; }

  if (!props.readOnly && event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
    editingCell.value = active;
    editingValue.value = event.key;
    nextTick(() => editInputRef.value?.focus());
  }
}

// ── Styles ────────────────────────────────────────────────────────────
const containerStyle: CSSProperties = {
  flex: "1",
  minHeight: "0px",
  minWidth: "0px",
  outline: "none",
  overflow: "auto",
  position: "relative",
};
const scrollSpaceStyle = computed<CSSProperties>(() => ({
  height: `${Math.max(containerHeight.value, HEADER_HEIGHT + (rowOffsets.value.at(-1) ?? 0))}px`,
  width: `${Math.max(containerWidth.value, ROW_HEADER_WIDTH + (colOffsets.value.at(-1) ?? 0))}px`,
}));
const viewportStyle = computed<CSSProperties>(() => ({
  height: `${containerHeight.value}px`,
  width: `${containerWidth.value}px`,
}));
const bodyCanvasStyle: CSSProperties = { left: `${ROW_HEADER_WIDTH}px`, position: "absolute", top: `${HEADER_HEIGHT}px` };
const colHeaderCanvasStyle: CSSProperties = { left: `${ROW_HEADER_WIDTH}px`, position: "absolute", top: "0px" };
const rowHeaderCanvasStyle: CSSProperties = { left: "0px", position: "absolute", top: `${HEADER_HEIGHT}px` };
const cornerCanvasStyle: CSSProperties = { left: "0px", position: "absolute", top: "0px" };

const editInputStyle = computed<CSSProperties>(() => {
  if (!editingCell.value) return { display: "none" };
  const cell = editingCell.value;
  const displayCell = getDisplayCell(cell);
  if (!displayCell) return { display: "none" };
  const z = zoomScale.value / 100;
  return {
    left: `${ROW_HEADER_WIDTH + getColOffsetSum(displayCell.col) - scrollLeft.value}px`,
    top: `${HEADER_HEIGHT + getRowOffsetSum(displayCell.row) - scrollTop.value}px`,
    width: `${effectiveColWidths.value[displayCell.col] || DEFAULT_COL_WIDTH}px`,
    height: `${effectiveRowHeights.value[displayCell.row] || DEFAULT_ROW_HEIGHT}px`,
    position: "absolute", zIndex: "10",
    border: "2px solid #2563eb", padding: "2px 4px",
    fontSize: `${12 * z}px`, fontFamily: "inherit", outline: "none",
    background: props.isDark ? "#18181b" : "#ffffff",
    color: props.isDark ? "#e4e4e7" : "#18181b",
    boxSizing: "border-box",
  };
});

defineExpose({
  captureZoomAnchor,
  restoreZoomAnchor,
  scrollToCell,
  hitTestCell,
  hitTestAxis,
  hitTestSheetPoint,
  get scrollContainer() {
    return containerRef.value;
  },
});

// ── Watchers ──────────────────────────────────────────────────────────
watch(
  () => [
    revision.value,
    workerRowsRevision.value,
    zoomScale.value,
    containerWidth.value,
    containerHeight.value,
    activeSheet.value,
    activeCell.value?.row,
    activeCell.value?.col,
    selection.value?.start.row,
    selection.value?.start.col,
    selection.value?.end.row,
    selection.value?.end.col,
    props.isDark,
    props.selectionColor,
    props.selectionFillColor,
  ],
  schedulePaint,
  { flush: "post" },
);

let resizeObs: ResizeObserver | null = null;

onMounted(() => {
  const container = containerRef.value;
  if (!container) return;
  containerWidth.value = container.clientWidth;
  containerHeight.value = container.clientHeight;
  resizeObs = new ResizeObserver((entries) => {
    for (const entry of entries) { containerWidth.value = entry.contentRect.width; containerHeight.value = entry.contentRect.height; }
  });
  resizeObs.observe(container);
  schedulePaint();
});

onUnmounted(() => {
  resizeObs?.disconnect();
  resetWorkerRows(null);
  if (paintFrame !== null) window.cancelAnimationFrame(paintFrame);
  paintFrame = null;
});
</script>

<style scoped>
.xlsx-grid { outline: none; }
.xlsx-grid__scroll-space { position: relative; }
.xlsx-grid__viewport {
  left: 0;
  overflow: hidden;
  pointer-events: none;
  position: sticky;
  top: 0;
}
.xlsx-grid__col-header,
.xlsx-grid__row-header,
.xlsx-grid__corner,
.xlsx-grid__body { pointer-events: none; }
.xlsx-grid__edit-input { box-sizing: border-box; pointer-events: auto; }
.xlsx-grid__comment-popover { background: #fffbe6; border: 1px solid #eab308; border-radius: 6px; box-shadow: 0 8px 24px rgba(0,0,0,.18); color: #422006; display: flex; flex-direction: column; gap: 3px; max-width: 260px; min-width: 150px; padding: 8px 10px; pointer-events: none; position: absolute; white-space: pre-wrap; z-index: 30; }
.xlsx-grid__comment-popover strong { font-size: 11px; }
.xlsx-grid__comment-popover span { font-size: 12px; line-height: 1.4; }
</style>
