/**
 * Keyboard navigation rules for the sheet grid.
 *
 * `XlsxGrid` owns the event wiring; the arithmetic behind Excel's page, edge
 * and jump keys lives here so it can be exercised without a mounted grid.
 */

/** Display-axis direction: `1` towards higher sheet indices, `-1` back. */
export type GridAxisDirection = -1 | 1;

/** Size a display row or column occupies; hidden entries report zero. */
function displaySize(size: number): number {
  return Number.isFinite(size) && size > 0 ? size : 0;
}

/**
 * How many display rows `PageUp`/`PageDown` travels for a body viewport.
 *
 * Excel pages by one screen but keeps the row at the old edge in view, so the
 * step is the number of rows that fit minus one. An unmeasured viewport (or a
 * sheet without rows) still moves one row, never zero.
 */
export function resolvePageRowStep(
  heights: readonly number[],
  viewportHeight: number,
): number {
  if (!(viewportHeight > 0) || heights.length === 0) return 1;
  let accumulated = 0;
  let fitting = 0;
  for (const height of heights) {
    const size = displaySize(height);
    if (size <= 0) continue;
    if (accumulated + size > viewportHeight) break;
    accumulated += size;
    fitting += 1;
  }
  return Math.max(1, fitting - 1);
}

/**
 * Direction a horizontal `Ctrl`/`Cmd` + arrow jump takes on the display axis.
 *
 * Plain arrows already mirror the column axis on an RTL sheet: `←` walks
 * towards the higher display index there. The jump keeps the same mapping so
 * `Ctrl+←` lands on the data edge the user pointed at.
 */
export function resolveHorizontalJumpDirection(
  key: string,
  rightToLeft: boolean,
): GridAxisDirection {
  const towardsHigherIndex = (key === "ArrowLeft") === rightToLeft;
  return towardsHigherIndex ? 1 : -1;
}

/**
 * Landing display index for `Ctrl`/`Cmd` + an arrow key.
 *
 * Excel jumps to the edge of the data region under the cursor: a run of
 * non-empty cells stops at the last cell of the run, a cell with an empty
 * neighbour continues to the next non-empty cell, and an empty stretch with no
 * data at all lands on the outermost cell of the axis.
 */
export function resolveJumpTarget(
  startIndex: number,
  direction: GridAxisDirection,
  length: number,
  isOccupied: (index: number) => boolean,
): number {
  if (length <= 0) return startIndex;
  const edge = direction > 0 ? length - 1 : 0;
  const from = Math.max(0, Math.min(startIndex, length - 1));
  const next = from + direction;
  if (next < 0 || next >= length) return from;

  if (isOccupied(from) && isOccupied(next)) {
    let index = from;
    while (index + direction >= 0 && index + direction < length && isOccupied(index + direction)) {
      index += direction;
    }
    return index;
  }

  let index = next;
  while (index >= 0 && index < length && !isOccupied(index)) index += direction;
  return index < 0 || index >= length ? edge : index;
}

/**
 * Largest display index whose sheet index is still inside the used range.
 *
 * `End`, `Ctrl+End` and `Ctrl+A` resolve their target from the sheet's used
 * extent; a hidden trailing row or column is missing from the display list, so
 * the search walks backwards instead of trusting `indices[maxUsed]`.
 */
export function resolveLastUsedDisplayIndex(
  indices: readonly number[],
  maxUsed: number | undefined,
  fallback: number,
): number {
  if (indices.length === 0) return Math.max(0, fallback);
  if (typeof maxUsed !== "number" || !Number.isFinite(maxUsed) || maxUsed < 0) return fallback;
  for (let index = indices.length - 1; index >= 0; index -= 1) {
    if (indices[index] <= maxUsed) return index;
  }
  return 0;
}

/** Whether two sheet addresses name the same cell. */
export function isSameCell(
  left: { row: number; col: number } | null | undefined,
  right: { row: number; col: number } | null | undefined,
): boolean {
  if (!left || !right) return false;
  return left.row === right.row && left.col === right.col;
}

/** Whether two sheet ranges describe the same rectangle. */
export function isSameRange(
  left: { start: { row: number; col: number }; end: { row: number; col: number } } | null | undefined,
  right: { start: { row: number; col: number }; end: { row: number; col: number } } | null | undefined,
): boolean {
  if (!left || !right) return false;
  return isSameCell(left.start, right.start) && isSameCell(left.end, right.end);
}
