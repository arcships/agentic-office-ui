import type { XlsxCellNumberFormatInput } from "./types/worksheet-types";

/**
 * Largest number of cells one selection summary reads.
 *
 * The status bar follows the selection, so the scan has to stay independent of
 * how much a sheet holds: whole-column selections and very large ranges are
 * summarized from their first cells and marked as truncated.
 */
export const XLSX_SELECTION_STATS_CELL_LIMIT = 10_000;

/** Aggregate of the non-empty cells the viewer read from a selection. */
export interface XlsxSelectionSummary {
  /** Cells covered by the selection, including cells beyond the scan budget. */
  cellCount: number;
  /** Non-empty cells that were read; empty cells never reach the summary. */
  count: number;
  /** Cells whose value is a number, formula results included. */
  numericCount: number;
  /** Sum of the numeric cells, or `null` when the selection has none. */
  sum: number | null;
  /** `sum / numericCount`, or `null` when the selection has no numbers. */
  average: number | null;
  min: number | null;
  max: number | null;
  /**
   * Number format of the first numeric cell. Totals are rendered with it so a
   * currency or percentage selection reports its own units, like Excel.
   */
  numberFormat: XlsxCellNumberFormatInput | null;
  /** True when the selection was larger than the scan budget. */
  truncated: boolean;
}

export interface XlsxSelectionSummaryOptions {
  /** Total selected cells, used as `cellCount`; defaults to the cells read. */
  totalCells?: number;
  /** Cells read before the budget stopped the scan; defaults to the values read. */
  scannedCells?: number;
  /** Number format of the first numeric cell, when the caller resolved one. */
  numberFormat?: XlsxCellNumberFormatInput | null;
}

/**
 * Cells the viewer asks the Worker to summarize for a Worker-backed selection.
 *
 * `hiddenRows` is the AutoFilter view state, not workbook data: filtered rows
 * keep their cells but do not contribute to the status bar totals. The list is
 * already clipped to the rows the scan can reach.
 */
export interface XlsxSelectionStatsQuery {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  /** Maximum cells to read, normally {@link XLSX_SELECTION_STATS_CELL_LIMIT}. */
  maxCells: number;
  /** Visible cells in the whole selection, reported as `cellCount`. */
  cellCount: number;
  /** Filtered-out rows inside the scan window. */
  hiddenRows?: readonly number[];
}

/**
 * Summarize one selection without touching the workbook.
 *
 * `values` holds one entry per non-empty cell the caller read: a finite number
 * for numeric cells and `null` for text, booleans and errors. `scannedCells`
 * counts every cell read, empty ones included, so `truncated` can tell a
 * budget-limited summary from a complete one. Numeric totals use compensated
 * summation so long decimal selections do not drift.
 */
export function summarizeSelectionValues(
  values: Iterable<number | null>,
  options: XlsxSelectionSummaryOptions = {},
): XlsxSelectionSummary {
  let count = 0;
  let numericCount = 0;
  let sum = 0;
  let compensation = 0;
  let min: number | null = null;
  let max: number | null = null;

  for (const value of values) {
    count += 1;
    if (value === null || !Number.isFinite(value)) continue;
    numericCount += 1;
    // Kahan summation keeps the totals stable for cent-like values.
    const addend = value - compensation;
    const next = sum + addend;
    compensation = (next - sum) - addend;
    sum = next;
    if (min === null || value < min) min = value;
    if (max === null || value > max) max = value;
  }

  const scannedCells = normalizeCount(options.scannedCells) ?? count;
  const totalCells = normalizeCount(options.totalCells) ?? scannedCells;

  return {
    average: numericCount > 0 ? sum / numericCount : null,
    cellCount: totalCells,
    count,
    max,
    min,
    numberFormat: options.numberFormat ?? null,
    numericCount,
    sum: numericCount > 0 ? sum : null,
    truncated: totalCells > scannedCells,
  };
}

function normalizeCount(value: number | undefined): number | null {
  if (value === undefined || !Number.isFinite(value)) return null;
  return Math.max(0, Math.floor(value));
}
