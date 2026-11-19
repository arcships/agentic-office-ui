import type { XlsxCellAddress, XlsxCellRange } from "@arcships/xlsx-core";

/**
 * Pure geometry and range rules for the selection's fill handle.
 *
 * They live outside `XlsxGrid.vue` so the drag semantics (which rows/columns a
 * fill extends, and when the handle sits at the selection's corner) can be
 * unit-tested without a canvas or a pointer device.
 */

/** Excel's fill handle: a 7 px square centred on the selection's corner. */
export const FILL_HANDLE_SIZE_PX = 7;

/** Copy of a range ordered top-left → bottom-right. */
export function normalizeFillRange(range: XlsxCellRange): XlsxCellRange {
  return {
    start: {
      row: Math.min(range.start.row, range.end.row),
      col: Math.min(range.start.col, range.end.col),
    },
    end: {
      row: Math.max(range.start.row, range.end.row),
      col: Math.max(range.start.col, range.end.col),
    },
  };
}

/** Whether two ranges cover exactly the same cells. */
export function fillRangesEqual(left: XlsxCellRange, right: XlsxCellRange): boolean {
  const a = normalizeFillRange(left);
  const b = normalizeFillRange(right);
  return a.start.row === b.start.row
    && a.start.col === b.start.col
    && a.end.row === b.end.row
    && a.end.col === b.end.col;
}

/**
 * Range a fill drag from `source`'s handle onto `cell` would cover.
 *
 * The source range is the pattern; the drag only grows it towards the cell
 * under the pointer, one axis at a time, exactly like Excel: dragging the
 * handle down from A1:B2 to row 4 fills A1:B4, dragging it right fills
 * A1:D2, and dragging diagonally grows both edges. A pointer inside the source
 * (or on its bottom-right cell) returns the source unchanged, which is what
 * suppresses the preview and the edit.
 */
export function resolveFillTargetRange(
  source: XlsxCellRange,
  cell: XlsxCellAddress,
): XlsxCellRange {
  const range = normalizeFillRange(source);
  const start = { ...range.start };
  const end = { ...range.end };
  const row = Math.trunc(cell.row);
  const col = Math.trunc(cell.col);
  if (Number.isFinite(row)) {
    if (row > end.row) end.row = row;
    else if (row < start.row) start.row = row;
  }
  if (Number.isFinite(col)) {
    if (col > end.col) end.col = col;
    else if (col < start.col) start.col = col;
  }
  return { start, end };
}

/** Top-left corner of the handle box centred on a selection corner. */
export function resolveFillHandleBox(
  cornerX: number,
  cornerY: number,
  size = FILL_HANDLE_SIZE_PX,
): { left: number; top: number } {
  const safeSize = Number.isFinite(size) && size > 0 ? size : FILL_HANDLE_SIZE_PX;
  return {
    left: (Number.isFinite(cornerX) ? cornerX : 0) - safeSize / 2,
    top: (Number.isFinite(cornerY) ? cornerY : 0) - safeSize / 2,
  };
}
