import type { Workbook } from "@dukelib/sheets-wasm";
import type { XlsxTableSortDirection } from "@arcships/xlsx-core";
import { applyCellMutationState, resolveCellFallbackStyle, type CellMutationState, type RangeCellMutation } from "./internal";
import { decodeHtmlEntities } from "./formatting";

/** One captured data row of a sortable range, keyed by its original row. */
export interface SortableRow {
  cells: CellMutationState[];
  index: number;
  sortBoolean: boolean | undefined;
  sortEmpty: boolean;
  sortNumber: number | undefined;
  sortText: string;
}

export interface SortableRange {
  endCol: number;
  endRow: number;
  sortCol: number;
  startCol: number;
  startRow: number;
}

/**
 * Captures every cell of a range together with the sort key of one column.
 *
 * The key uses the calculated value, not the raw formula, so a column of
 * formulas sorts by what the user sees; `getFormattedValueAt` is the fallback
 * for text and for empty cells.
 */
export function captureSortableRows(
  worksheet: ReturnType<Workbook["getSheet"]>,
  range: SortableRange
): SortableRow[] {
  const rows: SortableRow[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const cells: CellMutationState[] = [];
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const cellValue = worksheet.getCellAt(row, col);
      try {
        cells.push({
          formula: worksheet.getFormulaAt(row, col) ?? null,
          style: worksheet.getCellStyleAt(row, col),
          value: cellValue.toJs()
        });
      } finally {
        cellValue.free();
      }
    }

    const calculated = worksheet.getCalculatedValueAt(row, range.sortCol);
    const formatted = decodeHtmlEntities(worksheet.getFormattedValueAt(row, range.sortCol) ?? "");
    try {
      rows.push({
        cells,
        index: row,
        sortBoolean: calculated.is_boolean ? calculated.asBoolean() : undefined,
        sortEmpty: calculated.is_empty || formatted.length === 0,
        sortNumber: calculated.is_number ? calculated.asNumber() : undefined,
        sortText: calculated.is_text ? (calculated.asText() ?? formatted) : formatted
      });
    } finally {
      calculated.free();
    }
  }
  return rows;
}

/**
 * Orders captured rows: empties sink, numbers compare numerically, booleans
 * before text and text uses a numeric-aware locale comparison.
 */
export function sortSortableRows(
  rows: readonly SortableRow[],
  direction: XlsxTableSortDirection
): SortableRow[] {
  return [...rows].sort((left, right) => {
    if (left.sortEmpty !== right.sortEmpty) {
      return left.sortEmpty ? 1 : -1;
    }

    if (left.sortNumber !== undefined && right.sortNumber !== undefined) {
      return direction === "ascending" ? left.sortNumber - right.sortNumber : right.sortNumber - left.sortNumber;
    }

    if (left.sortBoolean !== undefined && right.sortBoolean !== undefined) {
      const leftValue = left.sortBoolean ? 1 : 0;
      const rightValue = right.sortBoolean ? 1 : 0;
      return direction === "ascending" ? leftValue - rightValue : rightValue - leftValue;
    }

    const comparedText = left.sortText.localeCompare(right.sortText, undefined, { numeric: true, sensitivity: "base" });
    return direction === "ascending" ? comparedText : -comparedText;
  });
}

/** Whether the sorted order would move any row. */
export function sortableRowsMoved(
  rows: readonly SortableRow[],
  sortedRows: readonly SortableRow[]
) {
  return !sortedRows.every((row, index) => row.index === rows[index]?.index);
}

/**
 * Writes the sorted rows back from `startRow` and returns the mutations of the
 * one range-edit history entry that restores them.
 */
export function writeSortedRows(
  worksheet: ReturnType<Workbook["getSheet"]>,
  rows: readonly SortableRow[],
  sortedRows: readonly SortableRow[],
  range: Pick<SortableRange, "startCol" | "endCol">,
  startRow: number,
  fallbackStyle: ReturnType<typeof resolveCellFallbackStyle>
): RangeCellMutation[] {
  const mutations: RangeCellMutation[] = [];
  for (let rowOffset = 0; rowOffset < rows.length; rowOffset += 1) {
    const targetRow = startRow + rowOffset;
    const sourceRow = sortedRows[rowOffset];
    const beforeRow = rows[rowOffset];
    if (!sourceRow || !beforeRow) {
      continue;
    }

    for (let colOffset = 0; colOffset <= range.endCol - range.startCol; colOffset += 1) {
      const before = beforeRow.cells[colOffset];
      const after = sourceRow.cells[colOffset];
      if (!before || !after) {
        continue;
      }

      const cell = { row: targetRow, col: range.startCol + colOffset };
      applyCellMutationState(worksheet, cell, after, fallbackStyle);
      mutations.push({ after, before, cell });
    }
  }
  return mutations;
}

/**
 * Sorts any rectangular range while preserving whole rows.
 *
 * Returns the mutations to record, or `null` when the range is invalid or the
 * sort would not move anything. The caller decides how to record history and
 * refresh the workbook view.
 */
export function sortWorksheetRange(
  worksheet: ReturnType<Workbook["getSheet"]>,
  range: SortableRange,
  direction: XlsxTableSortDirection,
  fallbackStyle: ReturnType<typeof resolveCellFallbackStyle>
): RangeCellMutation[] | null {
  if (range.startRow > range.endRow || range.startCol > range.endCol || range.sortCol > range.endCol) {
    return null;
  }
  const rows = captureSortableRows(worksheet, range);
  const sortedRows = sortSortableRows(rows, direction);
  if (!sortableRowsMoved(rows, sortedRows)) {
    return null;
  }
  return writeSortedRows(worksheet, rows, sortedRows, range, range.startRow, fallbackStyle);
}

