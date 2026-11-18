// replace.ts — find & replace inside one worksheet.
//
// Extracted from the editing domain: search reports one match per cell plus the
// offsets of the matched text, so the replacement has to resolve values and
// formulas itself, keep the cell's number format and type coercion rules, and
// record the whole sweep as one undo entry.

import type {
  XlsxCellAddress,
  XlsxCellRange,
  XlsxReplaceInSheetOptions,
  XlsxReplaceResult,
  XlsxSheetData,
} from "@arcships/xlsx-core";
import { createSurfaceSearchExpression } from "@arcships/office-runtime";
import { cellAddressToA1, normalizeRange } from "./selection";
import { canonicalizeFormula, resolveCellInput } from "./cell-input";
import { resolveEffectiveCellNumberFormat } from "./formatting";
import type { RangeCellMutation, XlsxControllerContext } from "./internal";

interface TextReplacement {
  text: string;
  occurrences: number;
}

/**
 * Literal replace inside `text`.
 *
 * Matching uses the shared surface-search expression, so case-insensitive
 * matching stays Unicode-aware (the same matcher the search session uses) and
 * `wholeCell` only accepts a match that spans the entire cell text.
 */
function replaceText(
  text: string,
  query: string,
  replacement: string,
  options: { matchCase: boolean; wholeCell: boolean },
): TextReplacement | undefined {
  if (!text || !query) return undefined;
  const expression = createSurfaceSearchExpression(query, {
    caseSensitive: options.matchCase,
  });
  let occurrences = 0;
  const next = text.replace(expression, (matched, offset: number) => {
    if (options.wholeCell && (offset !== 0 || matched.length !== text.length)) {
      return matched;
    }
    occurrences += 1;
    return replacement;
  });
  if (occurrences === 0) return undefined;
  return { text: next, occurrences };
}

/** Replaces exactly one occurrence at a known offset, as reported by search. */
function replaceOccurrence(
  text: string,
  query: string,
  replacement: string,
  occurrence: { start: number; end: number },
  options: { matchCase: boolean },
): TextReplacement | undefined {
  const { start, end } = occurrence;
  if (start < 0 || end > text.length || end <= start) return undefined;
  const matched = text.slice(start, end);
  const expected = options.matchCase ? query : query.toLowerCase();
  const actual = options.matchCase ? matched : matched.toLowerCase();
  if (expected !== actual) return undefined;
  return {
    text: `${text.slice(0, start)}${replacement}${text.slice(end)}`,
    occurrences: 1,
  };
}

function usedRange(sheet: XlsxSheetData): XlsxCellRange {
  return {
    start: { row: Math.max(0, sheet.minUsedRow), col: Math.max(0, sheet.minUsedCol) },
    end: { row: Math.max(0, sheet.maxUsedRow), col: Math.max(0, sheet.maxUsedCol) },
  };
}

export function createReplaceDomain(ctx: XlsxControllerContext) {
  function emptyResult(sheetIndex: number, workbookSheetIndex: number): XlsxReplaceResult {
    return {
      sheetIndex,
      workbookSheetIndex,
      replacedCells: 0,
      replacedOccurrences: 0,
      changed: false,
    };
  }

  function replaceInSheet(
    query: string,
    replacement: string,
    options: XlsxReplaceInSheetOptions = {},
  ): XlsxReplaceResult {
    const targetIndex = options.sheetIndex ?? ctx.activeSheetIndex.value;
    const targetSheet = ctx.sheets.value[targetIndex];
    if (!query || !targetSheet) {
      return emptyResult(targetIndex, targetSheet?.workbookSheetIndex ?? -1);
    }
    if (ctx.isEditingLocked.value || !ctx.workbook.value) {
      return emptyResult(targetIndex, targetSheet.workbookSheetIndex);
    }
    // The history entry names the active sheet, so a command aimed at another
    // sheet activates it first: undo then restores the sheet the user sees.
    if (targetIndex !== ctx.activeSheetIndex.value) {
      ctx.switchToSheet(targetIndex);
    }
    const worksheet = ctx.getActiveWorksheet();
    if (!worksheet) return emptyResult(targetIndex, targetSheet.workbookSheetIndex);

    const matchCase = options.matchCase === true;
    const wholeCell = options.wholeCell === true;
    const includeFormulas = options.formula !== false;
    const requestedRange = normalizeRange(options.range ?? usedRange(targetSheet));
    const occurrence = options.occurrence;
    const firstCell: XlsxCellAddress | undefined = occurrence?.cell;
    const mutations: RangeCellMutation[] = [];
    let replacedOccurrences = 0;

    // Empty cells hold nothing to replace, so intersecting the request with the
    // sheet's used range keeps a whole-column range from walking a million rows.
    const used = usedRange(targetSheet);
    const range = {
      start: {
        row: Math.max(requestedRange.start.row, used.start.row),
        col: Math.max(requestedRange.start.col, used.start.col),
      },
      end: {
        row: Math.min(requestedRange.end.row, used.end.row),
        col: Math.min(requestedRange.end.col, used.end.col),
      },
    };
    if (!occurrence && (range.start.row > range.end.row || range.start.col > range.end.col)) {
      return emptyResult(targetIndex, targetSheet.workbookSheetIndex);
    }
    const startRow = Math.max(0, Math.min(range.start.row, range.end.row));
    const endRow = Math.max(0, Math.max(range.start.row, range.end.row));
    const startCol = Math.max(0, Math.min(range.start.col, range.end.col));
    const endCol = Math.max(0, Math.max(range.start.col, range.end.col));

    const rows = occurrence
      ? [occurrence.cell.row]
      : Array.from({ length: endRow - startRow + 1 }, (_unused, index) => startRow + index);
    const cols = occurrence
      ? [occurrence.cell.col]
      : Array.from({ length: endCol - startCol + 1 }, (_unused, index) => startCol + index);

    for (const row of rows) {
      for (const col of cols) {
        const cell = { row, col };
        if (worksheet.isMergedSecondary(row, col)) continue;

        const address = cellAddressToA1(cell);
        const formula = worksheet.getFormulaAt(row, col);
        let before = null as ReturnType<XlsxControllerContext["captureCellMutationState"]>;
        let after = null as ReturnType<XlsxControllerContext["captureCellMutationState"]>;

        if (formula && includeFormulas) {
          const edited = occurrence
            ? replaceOccurrence(formula, query, replacement, occurrence, { matchCase })
            : replaceText(formula, query, replacement, { matchCase, wholeCell });
          if (!edited) continue;
          before = ctx.captureCellMutationState(cell);
          if (!before) continue;
          worksheet.setFormula(address, canonicalizeFormula(edited.text));
          after = ctx.captureCellMutationState(cell);
          if (!after) continue;
          replacedOccurrences += edited.occurrences;
        } else if (!formula) {
          const cellValue = worksheet.getCellAt(row, col);
          let rawValue: unknown;
          try {
            rawValue = cellValue.toJs();
          } finally {
            cellValue.free();
          }
          if (rawValue === null || rawValue === undefined) continue;
          const text = typeof rawValue === "string" ? rawValue : String(rawValue);
          const edited = occurrence
            ? replaceOccurrence(text, query, replacement, occurrence, { matchCase })
            : replaceText(text, query, replacement, { matchCase, wholeCell });
          if (!edited) continue;

          before = ctx.captureCellMutationState(cell);
          if (!before) continue;
          // Same coercion as typed input, so replacing "42" with "43" keeps a
          // number instead of silently turning the cell into text.
          const input = resolveCellInput(edited.text, {
            date1904: worksheet.date1904 === true,
            locale: ctx.locale.value,
            numberFormat: resolveEffectiveCellNumberFormat(
              worksheet.getCellStyleAt(row, col),
              ctx.activeSheet.value,
              row,
              col,
            ),
          });
          if (input.kind === "formula") {
            worksheet.setFormula(address, input.formula);
          } else {
            worksheet.setCell(address, input.value);
          }
          if (input.numberFormat) {
            worksheet.setCellStyleAt(row, col, { numberFormat: input.numberFormat });
          }
          after = ctx.captureCellMutationState(cell);
          if (!after) continue;
          replacedOccurrences += edited.occurrences;
        } else {
          // A formula cell whose value is not rewritten must not fall through to
          // the value path: its value is the formula's cached result.
          continue;
        }

        mutations.push({ before, after, cell });
      }
    }

    if (mutations.length === 0) {
      return emptyResult(targetIndex, targetSheet.workbookSheetIndex);
    }

    if (ctx.maybeRecalculateWorkbook(ctx.workbook.value)) {
      ctx.refreshWorkbookState(ctx.workbook.value);
    }

    const activeCellAfter = firstCell ?? mutations[0].cell;
    const selectionAfter: XlsxCellRange = { start: activeCellAfter, end: activeCellAfter };
    ctx.recordRangeEditHistory(mutations, selectionAfter, activeCellAfter);
    ctx.activeCell.value = activeCellAfter;
    ctx.selectionAnchorRef.value = activeCellAfter;
    ctx.selection.value = selectionAfter;

    return {
      sheetIndex: targetIndex,
      workbookSheetIndex: targetSheet.workbookSheetIndex,
      replacedCells: mutations.length,
      replacedOccurrences,
      changed: true,
    };
  }

  return { replaceInSheet };
}
