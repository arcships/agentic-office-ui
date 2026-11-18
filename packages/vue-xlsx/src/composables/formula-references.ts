import type { XlsxCellAddress } from "@arcships/xlsx-core";
import { isCellReferenceBoundary } from "./cell-input";
import { columnLabel } from "./selection";

/**
 * Excel's hard worksheet bounds: 16,384 columns (XFD) by 1,048,576 rows.
 *
 * `XlsxCellAddress` is zero-based, so valid rows are `0..SHEET_MAX_ROWS - 1`
 * and valid columns are `0..SHEET_MAX_COLUMNS - 1`.
 */
export const SHEET_MAX_COLUMNS = 16_384;
export const SHEET_MAX_ROWS = 1_048_576;

export function isWithinSheetBounds(cell: XlsxCellAddress) {
  return (
    cell.row >= 0 &&
    cell.row < SHEET_MAX_ROWS &&
    cell.col >= 0 &&
    cell.col < SHEET_MAX_COLUMNS
  );
}

function columnIndexFromLabel(label: string) {
  let index = 0;
  for (const char of label.toUpperCase()) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index;
}

const CELL_REFERENCE = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/y;

function matchCellReference(formula: string, index: number) {
  CELL_REFERENCE.lastIndex = index;
  const match = CELL_REFERENCE.exec(formula);
  if (!match || !isCellReferenceBoundary(formula, index, index + match[0].length)) {
    return null;
  }

  return {
    columnAnchor: match[1],
    columnPart: match[2],
    rowAnchor: match[3],
    rowPart: match[4],
    token: match[0],
  };
}

/**
 * Rewrites the relative cell references in a formula the way Excel does when a
 * cell is filled or copied/pasted: every reference moves by the row/column
 * delta, absolute (`$`) parts stay put, and references pushed off the sheet
 * become `#REF!`.
 *
 * String literals and quoted sheet names are left untouched, and function
 * names that look like cell addresses (`LOG10`) are not rewritten.
 */
export function shiftFormulaReferences(
  formula: string,
  rowDelta: number,
  colDelta: number,
) {
  if ((rowDelta === 0 && colDelta === 0) || !formula.startsWith("=")) {
    return formula;
  }

  let result = "";
  let index = 0;
  while (index < formula.length) {
    const char = formula[index] ?? "";

    // String literals ("say ""hi""") and quoted sheet names ('My Sheet')
    // keep their contents verbatim.
    if (char === '"' || char === "'") {
      let end = index + 1;
      while (end < formula.length) {
        if (formula[end] === char) {
          if (formula[end + 1] === char) {
            end += 2;
            continue;
          }
          end += 1;
          break;
        }
        end += 1;
      }
      result += formula.slice(index, end);
      index = end;
      continue;
    }

    const reference = matchCellReference(formula, index);
    if (reference) {
      const { columnAnchor, columnPart, rowAnchor, rowPart, token } = reference;
      const column = columnIndexFromLabel(columnPart);
      const row = Number(rowPart);
      if (
        column >= 1 &&
        column <= SHEET_MAX_COLUMNS &&
        row >= 1 &&
        row <= SHEET_MAX_ROWS
      ) {
        const nextColumn = columnAnchor ? column : column + colDelta;
        const nextRow = rowAnchor ? row : row + rowDelta;
        if (
          nextColumn < 1 ||
          nextColumn > SHEET_MAX_COLUMNS ||
          nextRow < 1 ||
          nextRow > SHEET_MAX_ROWS
        ) {
          result += "#REF!";
        } else {
          result += `${columnAnchor}${columnLabel(nextColumn - 1)}${rowAnchor}${nextRow}`;
        }
        index += token.length;
        continue;
      }
    }

    result += char;
    index += 1;
  }

  return result;
}
