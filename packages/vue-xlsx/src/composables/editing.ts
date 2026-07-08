import type { Workbook } from "@dukelib/sheets-wasm";
import type {
  XlsxCellAddress,
  XlsxCellRange,
  XlsxCellStyleInput
} from "@extend-ai/xlsx-core";
import { loadWorkbookChartAssets } from "@extend-ai/xlsx-core";
import { cellAddressToA1, normalizeRange, rangeContainsCell, rangeToA1 } from "./selection";
import {
  applyCellMutationState,
  cloneCellStyle,
  coerceUserEnteredValue,
  normalizeCellValue,
  pushHistoryEntry,
  type CellMutationState,
  type HistoryEntry
} from "./internal";
import { buildSheetList, buildVisibleSheetIndexMap } from "./workbook-state";
import type { XlsxControllerContext } from "./internal";

export { applyCellMutationState, cloneCellStyle, coerceUserEnteredValue, normalizeCellValue, pushHistoryEntry };
export type { CellMutationState, HistoryEntry };

export function createEditingDomain(ctx: XlsxControllerContext) {
  function selectCell(cell: XlsxCellAddress, options?: { extend?: boolean }) {
    ctx.selectedChartId.value = null;
    ctx.selectedChartElement.value = null;
    ctx.selectedImageId.value = null;
    ctx.activeCell.value = cell;
    if (options?.extend && ctx.selectionAnchorRef.value) {
      ctx.selection.value = normalizeRange({ start: ctx.selectionAnchorRef.value, end: cell });
      return;
    }

    ctx.selectionAnchorRef.value = cell;
    ctx.selection.value = { start: cell, end: cell };
  }

  function selectRange(range: XlsxCellRange) {
    const normalized = normalizeRange(range);
    ctx.selectedChartId.value = null;
    ctx.selectedChartElement.value = null;
    ctx.selectedImageId.value = null;
    ctx.selectionAnchorRef.value = normalized.start;
    ctx.activeCell.value = normalized.end;
    ctx.selection.value = normalized;
  }

  function clearSelection() {
    ctx.selectionAnchorRef.value = null;
    ctx.activeCell.value = null;
    ctx.selection.value = null;
    ctx.selectedChartId.value = null;
    ctx.selectedChartElement.value = null;
    ctx.selectedImageId.value = null;
  }

  function clearSelectedCells() {
    const worksheet = ctx.getActiveWorksheet();
    const targetRange = ctx.selection.value ?? (ctx.activeCell.value ? { start: ctx.activeCell.value, end: ctx.activeCell.value } : null);
    if (ctx.readOnly.value || !worksheet || !ctx.workbook.value || !targetRange) {
      return;
    }

    const normalized = normalizeRange(targetRange);
    const mutations: import("./internal").RangeCellMutation[] = [];
    for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
      for (let col = normalized.start.col; col <= normalized.end.col; col += 1) {
        if (worksheet.isMergedSecondary(row, col)) {
          continue;
        }

        const cell = { row, col };
        const before = ctx.captureCellMutationState(cell);
        if (!before) {
          continue;
        }

        worksheet.setCell(cellAddressToA1({ row, col }), "");
        const after = ctx.captureCellMutationState(cell);
        if (!after) {
          continue;
        }
        mutations.push({
          after,
          before,
          cell
        });
      }
    }

    ctx.maybeRecalculateWorkbook(ctx.workbook.value);
    ctx.refreshWorkbookState(ctx.workbook.value);
    ctx.recordRangeEditHistory(mutations, normalized, ctx.activeCell.value ?? normalized.start);
  }

  function setCellValue(cell: XlsxCellAddress, value: string) {
    const worksheet = ctx.getActiveWorksheet();
    if (ctx.readOnly.value || !worksheet || !ctx.workbook.value) {
      return;
    }

    const before = ctx.captureCellMutationState(cell);
    if (!before) {
      return;
    }

    const nextValue = coerceUserEnteredValue(value);
    worksheet.setCell(cellAddressToA1(cell), nextValue);
    const after = ctx.captureCellMutationState(cell);
    if (!after) {
      return;
    }
    ctx.maybeRecalculateWorkbook(ctx.workbook.value);
    ctx.refreshWorkbookState(ctx.workbook.value);
    ctx.recordCellEditHistory(cell, before, after);
  }

  function setCellFormula(cell: XlsxCellAddress, formula: string) {
    const worksheet = ctx.getActiveWorksheet();
    if (ctx.readOnly.value || !worksheet || !ctx.workbook.value) {
      return;
    }

    const before = ctx.captureCellMutationState(cell);
    if (!before) {
      return;
    }

    const trimmedFormula = formula.trim();
    if (!trimmedFormula) {
      worksheet.setCell(cellAddressToA1(cell), "");
    } else {
      worksheet.setFormula(cellAddressToA1(cell), trimmedFormula);
    }
    const after = ctx.captureCellMutationState(cell);
    if (!after) {
      return;
    }
    ctx.maybeRecalculateWorkbook(ctx.workbook.value);
    ctx.refreshWorkbookState(ctx.workbook.value);
    ctx.recordCellEditHistory(cell, before, after);
  }

  function setCellStyle(cell: XlsxCellAddress, style: XlsxCellStyleInput) {
    const worksheet = ctx.getActiveWorksheet();
    if (ctx.readOnly.value || !worksheet || !ctx.workbook.value) {
      return;
    }

    const before = ctx.captureCellMutationState(cell);
    if (!before) {
      return;
    }

    worksheet.setCellStyleAt(cell.row, cell.col, style);
    const after = ctx.captureCellMutationState(cell);
    if (!after) {
      return;
    }

    ctx.refreshWorkbookState(ctx.workbook.value);
    ctx.recordCellEditHistory(cell, before, after);
  }

  function setSelectedCellValue(value: string) {
    if (!ctx.activeCell.value) {
      return;
    }

    setCellValue(ctx.activeCell.value, value);
  }

  function setSelectedCellFormula(formula: string) {
    if (!ctx.activeCell.value) {
      return;
    }

    setCellFormula(ctx.activeCell.value, formula);
  }

  function setSelectedFormula(formula: string) {
    const selectedFormulaTarget = ctx.selectedFormulaTarget.value;
    if (selectedFormulaTarget?.kind === "chartSeries") {
      return ctx.setChartSeriesFormula(selectedFormulaTarget.chartId, selectedFormulaTarget.seriesIndex, formula);
    }

    if (!ctx.activeCell.value) {
      return false;
    }

    setCellFormula(ctx.activeCell.value, formula);
    return true;
  }

  function setSelectedCellStyle(style: XlsxCellStyleInput) {
    if (!ctx.activeCell.value) {
      return;
    }

    setCellStyle(ctx.activeCell.value, style);
  }

  function setRangeStyle(range: XlsxCellRange, style: XlsxCellStyleInput) {
    const worksheet = ctx.getActiveWorksheet();
    if (ctx.readOnly.value || !worksheet || !ctx.workbook.value) {
      return;
    }

    const normalized = normalizeRange(range);
    const beforeStates: Array<{ before: CellMutationState; cell: XlsxCellAddress }> = [];
    for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
      for (let col = normalized.start.col; col <= normalized.end.col; col += 1) {
        const cell = { row, col };
        const before = ctx.captureCellMutationState(cell);
        if (!before) {
          continue;
        }
        beforeStates.push({
          before,
          cell
        });
      }
    }

    if (beforeStates.length === 0) {
      return;
    }

    worksheet.setRangeStyle(rangeToA1(normalized), style);
    const mutations: import("./internal").RangeCellMutation[] = [];
    for (const mutation of beforeStates) {
      const after = ctx.captureCellMutationState(mutation.cell);
      if (!after) {
        continue;
      }
      mutations.push({
        after,
        before: mutation.before,
        cell: mutation.cell
      });
    }

    ctx.refreshWorkbookState(ctx.workbook.value);
    ctx.recordRangeEditHistory(mutations, ctx.selection.value, ctx.activeCell.value);
  }

  function fillSelection(targetRange: XlsxCellRange) {
    const worksheet = ctx.getActiveWorksheet();
    if (ctx.readOnly.value || !worksheet || !ctx.workbook.value || !ctx.selection.value) {
      return;
    }

    const sourceRange = normalizeRange(ctx.selection.value);
    const nextRange = normalizeRange(targetRange);
    const sourceHeight = sourceRange.end.row - sourceRange.start.row + 1;
    const sourceWidth = sourceRange.end.col - sourceRange.start.col + 1;

    if (sourceHeight <= 0 || sourceWidth <= 0) {
      return;
    }

    const mutations: import("./internal").RangeCellMutation[] = [];
    for (let row = nextRange.start.row; row <= nextRange.end.row; row += 1) {
      for (let col = nextRange.start.col; col <= nextRange.end.col; col += 1) {
        if (rangeContainsCell(sourceRange, { row, col })) {
          continue;
        }

        const targetCell = { row, col };
        const before = ctx.captureCellMutationState(targetCell);
        if (!before) {
          continue;
        }

        const sourceRow = sourceRange.start.row + ((row - nextRange.start.row) % sourceHeight);
        const sourceCol = sourceRange.start.col + ((col - nextRange.start.col) % sourceWidth);
        const sourceFormula = worksheet.getFormulaAt(sourceRow, sourceCol);
        const sourceStyle = cloneCellStyle(worksheet.getCellStyleAt(sourceRow, sourceCol));

        if (sourceFormula) {
          worksheet.setFormula(cellAddressToA1(targetCell), sourceFormula);
        } else {
          const sourceValue = normalizeCellValue(worksheet.getCellAt(sourceRow, sourceCol).toJs());
          worksheet.setCell(cellAddressToA1(targetCell), sourceValue);
        }

        if (sourceStyle && typeof sourceStyle === "object") {
          worksheet.setCellStyleAt(targetCell.row, targetCell.col, sourceStyle);
        }

        const after = ctx.captureCellMutationState(targetCell);
        if (!after) {
          continue;
        }
        mutations.push({
          after,
          before,
          cell: targetCell
        });
      }
    }

    ctx.maybeRecalculateWorkbook(ctx.workbook.value);
    ctx.refreshWorkbookState(ctx.workbook.value);
    ctx.selection.value = nextRange;
    ctx.activeCell.value = nextRange.end;
    ctx.selectionAnchorRef.value = nextRange.start;
    ctx.recordRangeEditHistory(mutations, nextRange, nextRange.end);
  }

  function mergeSelection() {
    const worksheet = ctx.getActiveWorksheet();
    if (ctx.readOnly.value || !worksheet || !ctx.selection.value || !ctx.workbook.value) {
      return;
    }

    ctx.recordHistoryBeforeMutation();
    worksheet.mergeCells(rangeToA1(ctx.selection.value));
    ctx.refreshWorkbookState(ctx.workbook.value);
  }

  function unmergeSelection() {
    const worksheet = ctx.getActiveWorksheet();
    if (ctx.readOnly.value || !worksheet || !ctx.selection.value || !ctx.workbook.value) {
      return;
    }

    ctx.recordHistoryBeforeMutation();
    worksheet.unmergeCells(rangeToA1(ctx.selection.value));
    ctx.refreshWorkbookState(ctx.workbook.value);
  }

  function addSheet(name?: string) {
    if (ctx.readOnly.value || !ctx.workbook.value) {
      return;
    }

    ctx.recordHistoryBeforeMutation();
    const baseName = name?.trim() || "Sheet";
    let candidate = baseName;
    let counter = 2;
    while (ctx.workbook.value.sheetIndex(candidate) !== undefined) {
      candidate = `${baseName} ${counter}`;
      counter += 1;
    }

    ctx.workbook.value.addSheet(candidate);
    ctx.sheetOriginsRef.value = [...ctx.sheetOriginsRef.value, null];
    ctx.imagesByWorkbookSheetIndex.value = [...ctx.imagesByWorkbookSheetIndex.value, []];
    ctx.shapesByWorkbookSheetIndex.value = [...ctx.shapesByWorkbookSheetIndex.value, []];
    const nextSheets = buildSheetList(
      ctx.workbook.value,
      ctx.imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex,
      ctx.imageAssetsRef.value?.themePalette,
      ctx.imageAssetsRef.value?.styleById,
      ctx.imageAssetsRef.value?.namedCellStyleByName,
      ctx.imageAssetsRef.value?.tableStyleByName,
      ctx.showHiddenSheets
    );
    ctx.sheets.value = nextSheets;
    const nextChartAssets = ctx.imageAssetsRef.value
      ? loadWorkbookChartAssets(ctx.workbook.value, ctx.imageAssetsRef.value, buildVisibleSheetIndexMap(nextSheets), ctx.showHiddenSheets)
      : null;
    if (ctx.imageAssetsRef.value) {
      ctx.setChartAssets(nextChartAssets);
    }
    const nextIndex = nextSheets.findIndex((sheet) => sheet.name === candidate);
    ctx.activeSheetIndex.value = nextIndex >= 0 ? nextIndex : 0;
    const nextTabIndex = nextChartAssets?.tabs.findIndex((tab) => tab.kind === "sheet" && tab.name === candidate) ?? -1;
    if (nextTabIndex >= 0) {
      ctx.activeTabIndex.value = nextTabIndex;
    }
    ctx.revision.value = ctx.revision.value + 1;
  }

  function removeActiveSheet() {
    const activeSheetData = ctx.activeSheet.value;
    if (ctx.readOnly.value || !ctx.workbook.value || !activeSheetData) {
      return;
    }

    ctx.recordHistoryBeforeMutation();
    ctx.workbook.value.removeSheet(activeSheetData.workbookSheetIndex);
    ctx.sheetOriginsRef.value = ctx.sheetOriginsRef.value.filter((_, index) => index !== activeSheetData.workbookSheetIndex);
    ctx.imagesByWorkbookSheetIndex.value = ctx.imagesByWorkbookSheetIndex.value.filter((_, index) => index !== activeSheetData.workbookSheetIndex);
    ctx.shapesByWorkbookSheetIndex.value = ctx.shapesByWorkbookSheetIndex.value.filter((_, index) => index !== activeSheetData.workbookSheetIndex);
    if (ctx.imageAssetsRef.value) {
      ctx.imageAssetsRef.value.sheetStatesByWorkbookSheetIndex = ctx.imageAssetsRef.value.sheetStatesByWorkbookSheetIndex.filter(
        (_, index) => index !== activeSheetData.workbookSheetIndex
      );
    }
    const nextSheets = buildSheetList(
      ctx.workbook.value,
      ctx.imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex,
      ctx.imageAssetsRef.value?.themePalette,
      ctx.imageAssetsRef.value?.styleById,
      ctx.imageAssetsRef.value?.namedCellStyleByName,
      ctx.imageAssetsRef.value?.tableStyleByName,
      ctx.showHiddenSheets
    );
    ctx.sheets.value = nextSheets;
    if (ctx.imageAssetsRef.value) {
      ctx.setChartAssets(loadWorkbookChartAssets(ctx.workbook.value, ctx.imageAssetsRef.value, buildVisibleSheetIndexMap(nextSheets), ctx.showHiddenSheets));
    }
    ctx.activeSheetIndex.value = Math.max(0, Math.min(ctx.activeSheetIndex.value, nextSheets.length - 1));
    ctx.revision.value = ctx.revision.value + 1;
  }

  function defineNamedRange(name: string, range?: XlsxCellRange | null) {
    if (ctx.readOnly.value || !ctx.workbook.value) {
      return;
    }

    const trimmed = name.trim();
    if (!trimmed) {
      return;
    }

    const targetRange = range ?? ctx.selection.value;
    if (!targetRange) {
      return;
    }

    ctx.recordHistoryBeforeMutation();
    ctx.workbook.value.defineName(trimmed, rangeToA1(targetRange));
    ctx.revision.value = ctx.revision.value + 1;
  }

  return {
    selectCell,
    selectRange,
    clearSelection,
    clearSelectedCells,
    setCellValue,
    setCellFormula,
    setCellStyle,
    setSelectedCellValue,
    setSelectedCellFormula,
    setSelectedFormula,
    setSelectedCellStyle,
    setRangeStyle,
    fillSelection,
    mergeSelection,
    unmergeSelection,
    addSheet,
    removeActiveSheet,
    defineNamedRange
  };
}
