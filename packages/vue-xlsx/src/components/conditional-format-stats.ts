import type { XlsxCellAddress, XlsxCellRange } from "@arcships/xlsx-core";

/** Axis-aligned range with `start <= end` on both axes. */
export interface NormalizedRange {
  endCol: number;
  endRow: number;
  startCol: number;
  startRow: number;
}

/** Used-extent of the sheet the rule ranges are clamped against. */
export interface ConditionalFormatBounds {
  maxUsedCol: number;
  maxUsedRow: number;
}

/**
 * Cell reads allowed for one rule's value pass. Rules over huge ranges are
 * stride-sampled so that a scroll repaint never depends on how much data the
 * range covers; rules that fit the budget are still scanned cell by cell.
 */
export const CONDITIONAL_FORMAT_CELL_BUDGET = 4096;

const NUMERIC_NOISE = /[,％%]/g;
const EMPTY_VALUES: number[] = [];

export function normalizeConditionalFormatRanges(
  ranges: readonly XlsxCellRange[],
): NormalizedRange[] {
  const normalized: NormalizedRange[] = [];
  for (const range of ranges) {
    const normalizedRange = {
      endCol: Math.max(range.start.col, range.end.col),
      endRow: Math.max(range.start.row, range.end.row),
      startCol: Math.min(range.start.col, range.end.col),
      startRow: Math.min(range.start.row, range.end.row),
    };
    if (normalizedRange.endRow < normalizedRange.startRow || normalizedRange.endCol < normalizedRange.startCol) continue;
    normalized.push(normalizedRange);
  }
  return normalized;
}

export function cellInNormalizedRange(cell: XlsxCellAddress, range: NormalizedRange): boolean {
  return cell.row >= range.startRow && cell.row <= range.endRow
    && cell.col >= range.startCol && cell.col <= range.endCol;
}

function numericDisplayValue(text: string): number | null {
  // Excel leaves blank cells out of a rule's scale, and formatted values arrive
  // with display grouping ("1,234") and percent signs.
  if (!text.trim()) return null;
  const value = Number(text.replace(NUMERIC_NOISE, ""));
  return Number.isFinite(value) ? value : null;
}

/**
 * Numeric values of one rule, clamped to the sheet's used extent and capped at
 * `cellBudget` reads. Percentile bounds need the value list itself, so this
 * returns the sampled values instead of only their extremes; above the budget
 * the walk is strided, which keeps a rule over a million cells from allocating
 * (and sorting) a million entries on a scroll frame.
 */
export function collectConditionalFormatValues(
  ranges: readonly NormalizedRange[],
  bounds: ConditionalFormatBounds,
  readValue: (cell: XlsxCellAddress) => string,
  cellBudget: number = CONDITIONAL_FORMAT_CELL_BUDGET,
): number[] {
  const budget = Math.max(1, Math.floor(cellBudget) || 1);
  const windows: NormalizedRange[] = [];
  let totalCells = 0;
  for (const range of ranges) {
    const endRow = Math.min(range.endRow, Math.max(bounds.maxUsedRow, range.startRow));
    const endCol = Math.min(range.endCol, Math.max(bounds.maxUsedCol, range.startCol));
    if (endRow < range.startRow || endCol < range.startCol) continue;
    windows.push({ endCol, endRow, startCol: range.startCol, startRow: range.startRow });
    totalCells += (endRow - range.startRow + 1) * (endCol - range.startCol + 1);
  }
  if (totalCells === 0) return EMPTY_VALUES;

  const values: number[] = [];
  let remaining = budget;
  for (const window of windows) {
    const cols = window.endCol - window.startCol + 1;
    const cells = (window.endRow - window.startRow + 1) * cols;
    const share = Math.max(1, Math.min(remaining, Math.round((budget * cells) / totalCells)));
    remaining = Math.max(0, remaining - share);
    const stride = Math.max(1, Math.ceil(cells / share));
    for (let offset = 0; offset < cells; offset += stride) {
      const value = numericDisplayValue(readValue({
        col: window.startCol + (offset % cols),
        row: window.startRow + Math.floor(offset / cols),
      }));
      if (value !== null) values.push(value);
    }
  }
  return values;
}

export interface ScopedCache<T> {
  /** Number of entries held for the current scope. */
  readonly size: number;
  /**
   * Value cached for `(scope, key)`. The entries are dropped as soon as a
   * different scope object is passed, which is how per-revision invalidation
   * happens: callers hand in a scope object whose identity changes with the
   * sheet revision.
   */
  resolve(scope: object, key: object, compute: () => T): T;
}

export function createScopedCache<T>(): ScopedCache<T> {
  let activeScope: object | null = null;
  let entries = new Map<object, T>();
  return {
    get size() {
      return entries.size;
    },
    resolve(scope, key, compute) {
      if (scope !== activeScope) {
        activeScope = scope;
        entries = new Map();
      }
      const cached = entries.get(key);
      if (cached !== undefined) return cached;
      const value = compute();
      entries.set(key, value);
      return value;
    },
  };
}
