import type { XlsxCellRange } from "@arcships/xlsx-core";

/**
 * Pure geometry / paint helpers shared by `XlsxGrid.vue`.
 *
 * They live outside the component so the fidelity rules (device-pixel caps,
 * merged-region viewport hits, icon-set buckets, overflow width, reference
 * outline clamps and resize zoom basis) can be unit-tested without a Canvas
 * implementation.
 */

// ── Canvas backing store ───────────────────────────────────────────────────

/**
 * Browser canvas limits vary, but 16.7 M pixels (4096 × 4096) and a 32 767 px
 * edge are safe across Blink, Gecko and WebKit. Beyond them browsers either
 * refuse to allocate the backing store or paint nothing, which blanks the grid.
 */
export const MAX_CANVAS_AREA_PX = 16_777_216;
export const MAX_CANVAS_DIMENSION_PX = 32_767;
export const MIN_CANVAS_PIXEL_RATIO = 0.1;

/** Current device-pixel ratio, read lazily so display / browser zoom changes apply. */
export function readDevicePixelRatio(): number {
  if (typeof window === "undefined") return 1;
  const value = window.devicePixelRatio;
  return Number.isFinite(value) && value > 0 ? value : 1;
}

/**
 * Effective backing-store ratio for a CSS-sized canvas. Starts from the live
 * `devicePixelRatio` and lowers it when the requested pixel area / edge would
 * exceed the browser limit, so the grid degrades in sharpness instead of
 * disappearing.
 */
export function resolveCanvasPixelRatio(
  cssWidth: number,
  cssHeight: number,
  devicePixelRatio: number,
): number {
  const width = Number.isFinite(cssWidth) ? Math.max(1, cssWidth) : 1;
  const height = Number.isFinite(cssHeight) ? Math.max(1, cssHeight) : 1;
  const requested = Number.isFinite(devicePixelRatio) && devicePixelRatio > 0 ? devicePixelRatio : 1;
  const areaScale = Math.sqrt(MAX_CANVAS_AREA_PX / (width * height));
  const dimensionScale = Math.min(
    MAX_CANVAS_DIMENSION_PX / width,
    MAX_CANVAS_DIMENSION_PX / height,
  );
  return Math.max(MIN_CANVAS_PIXEL_RATIO, Math.min(requested, areaScale, dimensionScale));
}

// ── Zoom basis ─────────────────────────────────────────────────────────────

export function resolveZoomFactor(zoomScale: number): number {
  const scale = zoomScale / 100;
  return Number.isFinite(scale) ? Math.max(0.1, scale) : 1;
}

/**
 * Row / column resize uses the same zoom basis as painting. The controller
 * allows zooming down to 10 %, so clamping at 0.5 made the first drag jump by
 * the ratio between the two scales.
 */
export function resolveAxisResizeSize(
  startSize: number,
  deltaPx: number,
  zoomScale: number,
  minimum: number,
): number {
  return Math.max(minimum, startSize + deltaPx / resolveZoomFactor(zoomScale));
}

// ── Per-frame identity keys ────────────────────────────────────────────────

/** XLSX's maximum column count; every cell key packs (row, col) into one number. */
export const CELL_KEY_STRIDE = 16_384;
/** XLSX's maximum row count, so (sheet, row, col) also stays one number. */
export const ROW_KEY_STRIDE = 1_048_576;

/**
 * Numeric identity for a sheet cell. Repaints resolve merge anchors, comments,
 * hyperlinks, sparklines and worker snapshots per visible cell; string keys
 * (`` `${row}:${col}` ``) allocated several throwaway strings per cell per
 * frame, numbers do not.
 */
export function cellKey(row: number, col: number): number {
  return row * CELL_KEY_STRIDE + col;
}

/** Numeric identity for a cell that also carries its workbook sheet index. */
export function sheetCellKey(sheetIndex: number, row: number, col: number): number {
  return (sheetIndex * ROW_KEY_STRIDE + row) * CELL_KEY_STRIDE + col;
}

// ── Merged regions ─────────────────────────────────────────────────────────

export interface PixelRect {
  height: number;
  left: number;
  top: number;
  width: number;
}

export interface PixelViewport {
  bottom: number;
  left: number;
  right: number;
  top: number;
}

export function intersectsViewport(rect: PixelRect, viewport: PixelViewport): boolean {
  return rect.width > 0
    && rect.height > 0
    && rect.left < viewport.right
    && rect.left + rect.width > viewport.left
    && rect.top < viewport.bottom
    && rect.top + rect.height > viewport.top;
}

// ── Formula reference outlines ─────────────────────────────────────────────

/**
 * Maps a sheet-coordinate reference range onto display coordinates for the
 * grid's outline pass.
 *
 * Whole-column (`A:A`) and whole-row (`1:1`) references end at Excel's maximum
 * index, which is not in the display map. Clamping the missing end to the last
 * displayed row/column draws the visible portion of the reference instead of
 * dropping the outline. A single off-screen cell has both endpoints unmapped and
 * is skipped rather than stretched over the viewport.
 */
export function resolvePaintedReferenceRange(
  range: XlsxCellRange,
  rowToDisplayIndex: ReadonlyMap<number, number>,
  colToDisplayIndex: ReadonlyMap<number, number>,
  displayRowCount: number,
  displayColCount: number,
): XlsxCellRange | null {
  if (displayRowCount <= 0 || displayColCount <= 0) return null;
  const startRow = rowToDisplayIndex.get(range.start.row);
  const endRow = rowToDisplayIndex.get(range.end.row);
  const startCol = colToDisplayIndex.get(range.start.col);
  const endCol = colToDisplayIndex.get(range.end.col);
  const wholeColumn = range.start.row === 0 && range.end.row === ROW_KEY_STRIDE - 1;
  const wholeRow = range.start.col === 0 && range.end.col === CELL_KEY_STRIDE - 1;

  if (startRow === undefined && endRow === undefined) {
    if (!wholeColumn) return null;
  }
  if (startCol === undefined && endCol === undefined) {
    if (!wholeRow) return null;
  }

  return {
    start: {
      row: startRow ?? 0,
      col: startCol ?? 0,
    },
    end: {
      row: endRow ?? Math.max(0, displayRowCount - 1),
      col: endCol ?? Math.max(0, displayColCount - 1),
    },
  };
}

// ── Text overflow ──────────────────────────────────────────────────────────

/**
 * Excel lets left-aligned text spill into adjacent empty cells. `neighbourWidths`
 * must only contain the cells between this one and the first occupied cell.
 */
export function resolveOverflowTextWidth(
  cellWidth: number,
  neighbourWidths: readonly number[],
): number {
  let width = Number.isFinite(cellWidth) ? Math.max(0, cellWidth) : 0;
  for (const neighbour of neighbourWidths) {
    if (!Number.isFinite(neighbour) || neighbour <= 0) continue;
    width += neighbour;
  }
  return width;
}
