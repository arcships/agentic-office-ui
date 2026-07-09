<template>
  <div
    ref="containerRef"
    class="xlsx-grid"
    :style="containerStyle"
    @keydown="onGridKeydown"
    @scroll="onScroll"
    @pointerdown="onPointerDown"
    @pointermove="onPointerMove"
    @pointerup="onPointerUp"
    @dblclick="onDblClick"
  >
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
      :style="editInputStyle"
      @keydown.enter="commitEdit"
      @keydown.escape="cancelEdit"
      @keydown.tab.prevent="commitAndTab"
      @blur="commitEdit"
    />
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, onMounted, onUnmounted, nextTick, type CSSProperties } from "vue";
import type { XlsxViewerController, XlsxCellAddress, XlsxCellRange, XlsxSheetData } from "@extend-ai/xlsx-core";

const props = defineProps<{
  controller: XlsxViewerController;
  getCellStyle?: ((cell: XlsxCellAddress) => Partial<CSSProperties> | undefined) | null;
  isDark?: boolean;
  readOnly?: boolean;
  selectionColor?: string;
  selectionFillColor?: string;
}>();

const emit = defineEmits<{
  cellDoubleClick: [cell: XlsxCellAddress];
}>();

// ── Constants ──────────────────────────────────────────────────────────
const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const HEADER_HEIGHT = 24;
const ROW_HEADER_WIDTH = 48;
const OVERSCAN = 240;
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

const isSelecting = ref(false);
const selectionAnchor = ref<XlsxCellAddress | null>(null);

// ── Computed from controller ──────────────────────────────────────────
const activeSheet = computed(() => props.controller.activeSheet as XlsxSheetData | null);
const zoomScale = computed(() => props.controller.zoomScale);
const activeCell = computed(() => props.controller.activeCell);
const selection = computed(() => props.controller.selection);
const revision = computed(() => props.controller.revision);

const effectiveColWidths = computed(() => {
  const sheet = activeSheet.value;
  if (!sheet) return [];
  const z = zoomScale.value / 100;
  return (sheet.colWidths ?? []).map(w => (w || DEFAULT_COL_WIDTH) * z);
});

const effectiveRowHeights = computed(() => {
  const sheet = activeSheet.value;
  if (!sheet) return [];
  const z = zoomScale.value / 100;
  return (sheet.rowHeights ?? []).map(h => (h || DEFAULT_ROW_HEIGHT) * z);
});

const visibleRowRange = computed(() => {
  const heights = effectiveRowHeights.value;
  if (heights.length === 0) return { start: 0, end: 0 };
  const top = scrollTop.value - HEADER_HEIGHT - OVERSCAN;
  const bottom = scrollTop.value - HEADER_HEIGHT + containerHeight.value + OVERSCAN;
  let offset = 0, start = 0, end = 0, foundStart = false;
  for (let i = 0; i < heights.length; i++) {
    offset += heights[i] || 0;
    if (!foundStart && offset > Math.max(0, top)) { start = Math.max(0, i - 2); foundStart = true; }
    if (offset < bottom) end = i + 1;
    if (offset > bottom + OVERSCAN && foundStart) break;
  }
  return { start, end: Math.min(end, heights.length - 1) };
});

const visibleColRange = computed(() => {
  const widths = effectiveColWidths.value;
  if (widths.length === 0) return { start: 0, end: 0 };
  const left = scrollLeft.value - ROW_HEADER_WIDTH - OVERSCAN;
  const right = scrollLeft.value - ROW_HEADER_WIDTH + containerWidth.value + OVERSCAN;
  let offset = 0, start = 0, end = 0, foundStart = false;
  for (let i = 0; i < widths.length; i++) {
    offset += widths[i] || 0;
    if (!foundStart && offset > Math.max(0, left)) { start = Math.max(0, i - 2); foundStart = true; }
    if (offset < right) end = i + 1;
    if (offset > right + OVERSCAN && foundStart) break;
  }
  return { start, end: Math.min(end, widths.length - 1) };
});

// ── Helpers ───────────────────────────────────────────────────────────
function getColOffsetSum(col: number): number {
  const widths = effectiveColWidths.value;
  let sum = 0;
  for (let i = 0; i < col && i < widths.length; i++) sum += widths[i] || 0;
  return sum;
}

function getRowOffsetSum(row: number): number {
  const heights = effectiveRowHeights.value;
  let sum = 0;
  for (let i = 0; i < row && i < heights.length; i++) sum += heights[i] || 0;
  return sum;
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

function hitTestCell(clientX: number, clientY: number): XlsxCellAddress | null {
  const container = containerRef.value;
  if (!container) return null;
  const rect = container.getBoundingClientRect();
  const x = clientX - rect.left + scrollLeft.value - ROW_HEADER_WIDTH;
  const y = clientY - rect.top + scrollTop.value - HEADER_HEIGHT;
  if (x < 0 || y < 0) return null;

  const widths = effectiveColWidths.value;
  const heights = effectiveRowHeights.value;
  let col = -1, colX = 0;
  for (let c = 0; c < widths.length; c++) { colX += widths[c] || 0; if (x < colX) { col = c; break; } }
  let row = -1, rowY = 0;
  for (let r = 0; r < heights.length; r++) { rowY += heights[r] || 0; if (y < rowY) { row = r; break; } }
  if (col < 0 || row < 0) return null;
  return { row, col };
}

// ── Canvas painting ───────────────────────────────────────────────────
function resizeCanvas(canvas: HTMLCanvasElement | null, width: number, height: number) {
  if (!canvas) return;
  canvas.width = Math.ceil(width * DPR);
  canvas.height = Math.ceil(height * DPR);
  canvas.style.width = `${width}px`;
  canvas.style.height = `${height}px`;
  const ctx = canvas.getContext("2d");
  if (ctx) ctx.setTransform(DPR, 0, 0, DPR, 0, 0);
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

  const bw = containerWidth.value - ROW_HEADER_WIDTH;
  const bh = containerHeight.value - HEADER_HEIGHT;
  resizeCanvas(canvas, bw, bh);

  ctx.clearRect(0, 0, bw, bh);
  ctx.save();
  ctx.translate(-scrollLeft.value, -scrollTop.value + HEADER_HEIGHT);

  const z = zoomScale.value / 100;
  const rowRange = visibleRowRange.value;
  const colRange = visibleColRange.value;
  const sheet = activeSheet.value;
  if (!sheet) { ctx.restore(); return; }

  const widths = effectiveColWidths.value;
  const heights = effectiveRowHeights.value;

  const bgColor = props.isDark ? "#18181b" : "#ffffff";
  const gridColor = props.isDark ? "#3f3f46" : "#d9d9d9";
  const textColor = props.isDark ? "#e4e4e7" : "#18181b";
  const selFill = props.selectionFillColor || (props.isDark ? "rgba(37, 99, 235, 0.15)" : "rgba(37, 99, 235, 0.08)");
  const selBorder = props.selectionColor || "#2563eb";

  ctx.fillStyle = bgColor;
  ctx.fillRect(0, 0, bw + scrollLeft.value, bh + scrollTop.value);

  // Gridlines
  if (sheet.showGridLines !== false) {
    ctx.strokeStyle = gridColor;
    ctx.lineWidth = 1;
    let x = 0;
    for (let c = 0; c <= colRange.end && c < widths.length; c++) {
      if (c >= colRange.start) {
        ctx.beginPath(); ctx.moveTo(x, getRowOffsetSum(rowRange.start)); ctx.lineTo(x, getRowOffsetSum(rowRange.end + 1)); ctx.stroke();
      }
      x += widths[c] || 0;
    }
    ctx.beginPath(); ctx.moveTo(x, getRowOffsetSum(rowRange.start)); ctx.lineTo(x, getRowOffsetSum(rowRange.end + 1)); ctx.stroke();
    let y = 0;
    for (let r = 0; r <= rowRange.end && r < heights.length; r++) {
      if (r >= rowRange.start) {
        ctx.beginPath(); ctx.moveTo(getColOffsetSum(colRange.start), y); ctx.lineTo(getColOffsetSum(colRange.end + 1), y); ctx.stroke();
      }
      y += heights[r] || 0;
    }
    ctx.beginPath(); ctx.moveTo(getColOffsetSum(colRange.start), y); ctx.lineTo(getColOffsetSum(colRange.end + 1), y); ctx.stroke();
  }

  // Cells
  ctx.font = `${Math.round(12 * z)}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif`;
  ctx.textBaseline = "middle";

  const sel = selection.value;
  const active = activeCell.value;

  for (let r = rowRange.start; r <= rowRange.end && r < heights.length; r++) {
    const rowH = heights[r] || 0;
    const rowY = getRowOffsetSum(r);
    for (let c = colRange.start; c <= colRange.end && c < widths.length; c++) {
      const colW = widths[c] || 0;
      const colX = getColOffsetSum(c);
      const cell: XlsxCellAddress = { row: r, col: c };

      if (isCellInRange(cell, sel)) {
        ctx.fillStyle = selFill;
        ctx.fillRect(colX + 1, rowY + 1, colW - 2, rowH - 2);
      }

      const value = props.controller.getCellDisplayValue(cell);
      if (value) {
        ctx.fillStyle = textColor;
        const displayText = truncateText(ctx, value, colW - 8);
        ctx.fillText(displayText, colX + 4, rowY + rowH / 2);
      }

      if (active && active.row === r && active.col === c) {
        ctx.strokeStyle = selBorder;
        ctx.lineWidth = 2;
        ctx.strokeRect(colX + 1, rowY + 1, colW - 2, rowH - 2);
      }
    }
  }
  ctx.restore();
}

function paintColHeaders() {
  const canvas = colHeaderCanvasRef.value;
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  const cw = containerWidth.value - ROW_HEADER_WIDTH;
  resizeCanvas(canvas, cw, HEADER_HEIGHT);
  ctx.clearRect(0, 0, cw, HEADER_HEIGHT);
  ctx.save();
  ctx.translate(-scrollLeft.value, 0);

  const bgColor = props.isDark ? "#27272a" : "#f4f4f5";
  const textColor = props.isDark ? "#a1a1aa" : "#71717a";
  const borderColor = props.isDark ? "#3f3f46" : "#d4d4d8";

  ctx.fillStyle = bgColor;
  ctx.fillRect(0, 0, cw + scrollLeft.value, HEADER_HEIGHT);
  ctx.font = "11px -apple-system, BlinkMacSystemFont, sans-serif";
  ctx.textBaseline = "middle";
  ctx.textAlign = "center";
  ctx.fillStyle = textColor;

  const widths = effectiveColWidths.value;
  const colRange = visibleColRange.value;
  let x = 0;
  for (let c = colRange.start; c <= colRange.end && c < widths.length; c++) {
    const colW = widths[c] || 0;
    ctx.fillText(columnLabel(c), x + colW / 2, HEADER_HEIGHT / 2);
    ctx.strokeStyle = borderColor; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(x + colW, 0); ctx.lineTo(x + colW, HEADER_HEIGHT); ctx.stroke();
    x += colW;
  }
  ctx.strokeStyle = borderColor;
  ctx.beginPath(); ctx.moveTo(0, HEADER_HEIGHT - 0.5); ctx.lineTo(cw + scrollLeft.value, HEADER_HEIGHT - 0.5); ctx.stroke();
  ctx.restore();
}

function paintRowHeaders() {
  const canvas = rowHeaderCanvasRef.value;
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  const rh = containerHeight.value - HEADER_HEIGHT;
  resizeCanvas(canvas, ROW_HEADER_WIDTH, rh);
  ctx.clearRect(0, 0, ROW_HEADER_WIDTH, rh);
  ctx.save();
  ctx.translate(0, -scrollTop.value + HEADER_HEIGHT);

  const bgColor = props.isDark ? "#27272a" : "#f4f4f5";
  const textColor = props.isDark ? "#a1a1aa" : "#71717a";
  const borderColor = props.isDark ? "#3f3f46" : "#d4d4d8";

  ctx.fillStyle = bgColor;
  ctx.fillRect(0, 0, ROW_HEADER_WIDTH, rh + scrollTop.value);
  ctx.font = "11px -apple-system, BlinkMacSystemFont, sans-serif";
  ctx.textBaseline = "middle";
  ctx.textAlign = "center";
  ctx.fillStyle = textColor;

  const heights = effectiveRowHeights.value;
  const rowRange = visibleRowRange.value;
  let y = 0;
  for (let r = rowRange.start; r <= rowRange.end && r < heights.length; r++) {
    const rowH = heights[r] || 0;
    ctx.fillText(String(r + 1), ROW_HEADER_WIDTH / 2, y + rowH / 2);
    ctx.strokeStyle = borderColor; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(0, y + rowH); ctx.lineTo(ROW_HEADER_WIDTH, y + rowH); ctx.stroke();
    y += rowH;
  }
  ctx.strokeStyle = borderColor;
  ctx.beginPath(); ctx.moveTo(ROW_HEADER_WIDTH - 0.5, 0); ctx.lineTo(ROW_HEADER_WIDTH - 0.5, rh + scrollTop.value); ctx.stroke();
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
function onScroll() {
  const container = containerRef.value;
  if (!container) return;
  scrollTop.value = container.scrollTop;
  scrollLeft.value = container.scrollLeft;
}

function onPointerDown(event: PointerEvent) {
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
}

function onPointerMove(event: PointerEvent) {
  if (!isSelecting.value || !selectionAnchor.value) return;
  const cell = hitTestCell(event.clientX, event.clientY);
  if (!cell) return;
  props.controller.selectRange({ start: selectionAnchor.value, end: cell });
}

function onPointerUp() {
  isSelecting.value = false;
  selectionAnchor.value = null;
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
}

function cancelEdit() {
  editingCell.value = null;
  editingValue.value = "";
}

function commitAndTab() {
  if (!editingCell.value) return;
  props.controller.setCellValue(editingCell.value, editingValue.value);
  const nextCell: XlsxCellAddress = { row: editingCell.value.row, col: editingCell.value.col + 1 };
  props.controller.selectCell(nextCell);
  editingCell.value = nextCell;
  editingValue.value = props.controller.getCellDisplayValue(nextCell);
}

function moveCell(row: number, col: number, extend: boolean) {
  const cell: XlsxCellAddress = { row: Math.max(0, row), col: Math.max(0, col) };
  if (extend) {
    props.controller.selectRange({ start: activeCell.value ?? cell, end: cell });
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
  const colX = getColOffsetSum(cell.col);
  const rowY = getRowOffsetSum(cell.row);
  const colW = effectiveColWidths.value[cell.col] || DEFAULT_COL_WIDTH;
  const rowH = effectiveRowHeights.value[cell.row] || DEFAULT_ROW_HEIGHT;
  const cellRight = colX + colW + ROW_HEADER_WIDTH;
  const cellBottom = rowY + rowH + HEADER_HEIGHT;

  if (colX < container.scrollLeft) container.scrollLeft = Math.max(0, colX);
  else if (cellRight > container.scrollLeft + container.clientWidth) container.scrollLeft = cellRight - container.clientWidth;
  if (rowY < container.scrollTop) container.scrollTop = Math.max(0, rowY);
  else if (cellBottom > container.scrollTop + container.clientHeight) container.scrollTop = cellBottom - container.clientHeight;
}

function onGridKeydown(event: KeyboardEvent) {
  const active = activeCell.value;
  if (!active) return;

  if ((event.ctrlKey || event.metaKey) && event.key === "c") { event.preventDefault(); props.controller.copySelectionToClipboard(); return; }
  if ((event.ctrlKey || event.metaKey) && event.key === "v") { event.preventDefault(); props.controller.pasteFromClipboard(); return; }

  const handlers: Record<string, () => void> = {
    ArrowUp: () => moveCell(active.row - 1, active.col, event.shiftKey),
    ArrowDown: () => moveCell(active.row + 1, active.col, event.shiftKey),
    ArrowLeft: () => moveCell(active.row, active.col - 1, event.shiftKey),
    ArrowRight: () => moveCell(active.row, active.col + 1, event.shiftKey),
    Tab: () => { event.preventDefault(); moveCell(active.row, active.col + 1, false); },
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
const containerStyle: CSSProperties = { overflow: "auto", position: "relative", flex: "1", outline: "none" };
const bodyCanvasStyle: CSSProperties = { left: `${ROW_HEADER_WIDTH}px`, position: "absolute", top: `${HEADER_HEIGHT}px` };
const colHeaderCanvasStyle: CSSProperties = { left: `${ROW_HEADER_WIDTH}px`, position: "absolute", top: "0px" };
const rowHeaderCanvasStyle: CSSProperties = { left: "0px", position: "absolute", top: `${HEADER_HEIGHT}px` };
const cornerCanvasStyle: CSSProperties = { left: "0px", position: "absolute", top: "0px" };

const editInputStyle = computed<CSSProperties>(() => {
  if (!editingCell.value) return { display: "none" };
  const cell = editingCell.value;
  const z = zoomScale.value / 100;
  return {
    left: `${ROW_HEADER_WIDTH + getColOffsetSum(cell.col) - scrollLeft.value}px`,
    top: `${HEADER_HEIGHT + getRowOffsetSum(cell.row) - scrollTop.value}px`,
    width: `${effectiveColWidths.value[cell.col] || DEFAULT_COL_WIDTH}px`,
    height: `${effectiveRowHeights.value[cell.row] || DEFAULT_ROW_HEIGHT}px`,
    position: "absolute", zIndex: "10",
    border: "2px solid #2563eb", padding: "2px 4px",
    fontSize: `${12 * z}px`, fontFamily: "inherit", outline: "none",
    background: props.isDark ? "#18181b" : "#ffffff",
    color: props.isDark ? "#e4e4e7" : "#18181b",
    boxSizing: "border-box",
  };
});

// ── Watchers ──────────────────────────────────────────────────────────
watch(
  () => [revision.value, zoomScale.value, scrollTop.value, scrollLeft.value, containerWidth.value, containerHeight.value, activeSheet.value],
  () => { paintGrid(); },
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
  paintGrid();
});

onUnmounted(() => { resizeObs?.disconnect(); });
</script>

<style scoped>
.xlsx-grid { outline: none; }
.xlsx-grid__col-header,
.xlsx-grid__row-header,
.xlsx-grid__corner,
.xlsx-grid__body { pointer-events: none; }
.xlsx-grid__edit-input { box-sizing: border-box; }
</style>
