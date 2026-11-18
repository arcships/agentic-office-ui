import type { Workbook, Worksheet } from "@dukelib/sheets-wasm";
import {
  XLSX_MAX_COLUMNS,
  XLSX_MAX_ROWS,
  normalizeWorkbookDefinedNames,
  parseA1CellReference,
  parseSheetState,
  remapA1References,
  remapAxisIndexForStructure,
  remapCellForStructure,
  remapFreezePanesForStructure,
  remapImageAnchorForStructure,
  remapRangeForStructure,
  remapSparklineForStructure,
  resolveWorksheetMergedRegions,
  shiftWorkbookSheetStructure,
  updateWorkbookChartAnchor,
  updateWorkbookFormControl,
  updateWorkbookImageAnchor,
  updateWorkbookShapeAnchor,
  type WorkbookSheetState,
  type XlsxCellAddress,
  type XlsxCellRange,
  type XlsxStructureChange,
  type XlsxStructureAxis
} from "@arcships/xlsx-core";
import type { XlsxControllerContext } from "./internal";
import { resolveCellFallbackStyle } from "./internal";
import { cellAddressToA1, normalizeRange, rangeToA1 } from "./selection";

type StructureKind = XlsxStructureChange["kind"];

type CellState = {
  formula: string | null;
  style: unknown;
  value: unknown;
};

function axisLimit(axis: XlsxStructureAxis) {
  return axis === "row" ? XLSX_MAX_ROWS : XLSX_MAX_COLUMNS;
}

function createCell(cell: XlsxCellAddress, axis: XlsxStructureAxis, value: number): XlsxCellAddress {
  return axis === "row" ? { ...cell, row: value } : { ...cell, col: value };
}

/**
 * Row/column structure commands, paste-compatible selection helpers and
 * workbook-scoped features that depend on the sheet grid (names, number
 * formats).
 *
 * The WASM engine has no row/column insert/delete calls, so a structure edit
 * moves every cell, size override, merge and drawing itself, then rewrites the
 * authored worksheet parts (hyperlinks, comments, conditional formats, data
 * validations, tables, panes) through the xlsx-core helpers.
 */
export function createStructureDomain(ctx: XlsxControllerContext) {
  function activeWorkbookSheetIndex() {
    return ctx.activeSheet.value?.workbookSheetIndex ?? -1;
  }

  function canMutateStructure() {
    return Boolean(
      ctx.workbook.value
      && ctx.activeSheet.value
      && ctx.getActiveWorksheet()
      && !ctx.isEditingLocked.value
      && !ctx.isWorkerBacked.value
    );
  }

  function normalizeRequest(axis: XlsxStructureAxis, kind: StructureKind, index: number, count: number) {
    const limit = axisLimit(axis);
    const start = Math.max(0, Math.min(limit - 1, Math.floor(Number(index) || 0)));
    const requested = Math.max(1, Math.min(limit - start, Math.floor(Number(count) || 1)));
    return { axis, count: requested, index: start, kind } as XlsxStructureChange;
  }

  function readCellState(worksheet: Worksheet, cell: XlsxCellAddress): CellState {
    const handle = worksheet.getCellAt(cell.row, cell.col);
    let value: unknown = "";
    try {
      value = handle.toJs();
    } finally {
      handle.free();
    }
    return {
      formula: worksheet.getFormulaAt(cell.row, cell.col) ?? null,
      style: worksheet.getCellStyleAt(cell.row, cell.col) ?? null,
      value
    };
  }

  function isEmptyCellState(state: CellState) {
    const value = state.value === null || state.value === undefined ? "" : state.value;
    return !state.formula && !state.style && value === "";
  }

  function writeCellState(
    worksheet: Worksheet,
    cell: XlsxCellAddress,
    state: CellState,
    fallbackStyle: unknown
  ) {
    const address = cellAddressToA1(cell);
    if (state.formula) {
      worksheet.setFormula(address, state.formula);
    } else {
      worksheet.setCell(address, state.value ?? "");
    }
    if (state.style && typeof state.style === "object") {
      worksheet.setCellStyleAt(cell.row, cell.col, state.style);
    } else if (worksheet.getCellStyleAt(cell.row, cell.col)) {
      worksheet.setCellStyleAt(cell.row, cell.col, fallbackStyle ?? {});
    }
  }

  function clearCell(worksheet: Worksheet, cell: XlsxCellAddress, fallbackStyle: unknown) {
    worksheet.setCell(cellAddressToA1(cell), "");
    if (worksheet.getCellStyleAt(cell.row, cell.col)) {
      worksheet.setCellStyleAt(cell.row, cell.col, fallbackStyle ?? {});
    }
  }

  /**
   * Moves every cell of the used rectangle by `count` positions, then clears
   * the vacated strip. Iterating from the far edge inwards prevents an overwrite
   * from destroying a cell that has not been read yet.
   */
  function shiftCells(
    worksheet: Worksheet,
    change: XlsxStructureChange,
    bounds: { maxRow: number; maxCol: number },
    fallbackStyle: unknown
  ) {
    const rowAxis = change.axis === "row";
    const primaryEnd = rowAxis ? bounds.maxRow : bounds.maxCol;
    const secondaryEnd = rowAxis ? bounds.maxCol : bounds.maxRow;
    const moveState = (sourceIndex: number, targetIndex: number) => {
      for (let secondary = 0; secondary <= secondaryEnd; secondary += 1) {
        const source = rowAxis
          ? { col: secondary, row: sourceIndex }
          : { col: sourceIndex, row: secondary };
        const target = rowAxis
          ? { col: secondary, row: targetIndex }
          : { col: targetIndex, row: secondary };
        const state = readCellState(worksheet, source);
        // A truly empty source must erase its target: the target row/column has
        // already been moved out of the way, so a stale value cannot survive.
        if (change.kind === "delete" || !isEmptyCellState(state)) {
          writeCellState(worksheet, target, state, fallbackStyle);
        }
      }
    };
    const clearStrip = (from: number, to: number) => {
      for (let primary = from; primary <= to; primary += 1) {
        for (let secondary = 0; secondary <= secondaryEnd; secondary += 1) {
          const cell = rowAxis
            ? { col: secondary, row: primary }
            : { col: primary, row: secondary };
          clearCell(worksheet, cell, fallbackStyle);
        }
      }
    };

    if (change.kind === "insert") {
      for (let primary = primaryEnd; primary >= change.index; primary -= 1) {
        moveState(primary, primary + change.count);
      }
      clearStrip(change.index, change.index + change.count - 1);
      return;
    }

    const lastSource = Math.min(primaryEnd, change.index + change.count - 1);
    for (let primary = lastSource + 1; primary <= primaryEnd; primary += 1) {
      moveState(primary, primary - change.count);
    }
    clearStrip(Math.max(change.index, primaryEnd - change.count + 1), primaryEnd);
  }

  function shiftAxisSizes(
    worksheet: Worksheet,
    change: XlsxStructureChange,
    bounds: { maxRow: number; maxCol: number },
  ) {
    const rowAxis = change.axis === "row";
    const end = rowAxis ? bounds.maxRow : bounds.maxCol;
    const getSize = (index: number) => rowAxis
      ? worksheet.getRowHeight(index)
      : worksheet.getColumnWidth(index);
    const setSize = (index: number, size: number) => {
      if (rowAxis) {
        worksheet.setRowHeight(index, size);
      } else {
        worksheet.setColumnWidth(index, size);
      }
    };
    const fallback = rowAxis
      ? worksheet.defaultRowHeight
      : worksheet.defaultColumnWidth;
    const sizes = new Map<number, number>();
    for (let index = 0; index <= end; index += 1) {
      const size = getSize(index);
      if (typeof size === "number" && Number.isFinite(size)) {
        sizes.set(index, size);
      }
    }
    if (sizes.size === 0) {
      return;
    }
    const next = new Map<number, number>();
    for (const [index, size] of sizes) {
      const mapped = remapAxisIndexForStructure(index, change);
      if (mapped !== null && mapped >= 0) {
        next.set(mapped, size);
      }
    }
    for (const index of sizes.keys()) {
      if (!next.has(index)) {
        setSize(index, fallback);
      }
    }
    for (const [index, size] of next) {
      if (sizes.get(index) !== size) {
        setSize(index, size);
      }
    }
  }

  function rewriteFormulas(worksheet: Worksheet, change: XlsxStructureChange, sheetName: string) {
    const formulas = Array.isArray(worksheet.formulaCells) ? worksheet.formulaCells : [];
    for (const entry of formulas) {
      if (!entry || typeof entry.row !== "number" || typeof entry.col !== "number") {
        continue;
      }
      const formula = typeof entry.formula === "string"
        ? entry.formula
        : worksheet.getFormulaAt(entry.row, entry.col) ?? "";
      if (!formula.startsWith("=")) {
        continue;
      }
      const next = remapA1References(formula, change, { sheetName });
      if (next !== formula) {
        worksheet.setFormula(cellAddressToA1({ col: entry.col, row: entry.row }), next);
      }
    }
  }

  function remapNumberRecord(
    record: Record<number, number> | undefined,
    change: XlsxStructureChange,
  ): Record<number, number> {
    const next: Record<number, number> = {};
    for (const [rawIndex, value] of Object.entries(record ?? {})) {
      const index = Number(rawIndex);
      const mapped = remapAxisIndexForStructure(index, change);
      if (mapped !== null) {
        next[mapped] = value;
      }
    }
    return next;
  }

  function remapIndexList(values: number[] | undefined, change: XlsxStructureChange) {
    return (values ?? []).flatMap((index) => {
      const mapped = remapAxisIndexForStructure(index, change);
      return mapped === null ? [] : [mapped];
    });
  }

  function remapCachedFormulaValues(
    values: Record<string, string> | undefined,
    change: XlsxStructureChange,
  ) {
    const next: Record<string, string> = {};
    for (const [address, value] of Object.entries(values ?? {})) {
      const parsed = parseA1CellReference(address);
      if (!parsed) {
        continue;
      }
      const mapped = remapCellForStructure(parsed, change);
      if (mapped) {
        next[cellAddressToA1(mapped)] = value;
      }
    }
    return next;
  }

  function remapContentBound(value: number | undefined, change: XlsxStructureChange) {
    if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
      return value;
    }
    return remapAxisIndexForStructure(value, change) ?? change.index;
  }

  /**
   * Re-parses the authored worksheet part after `shiftWorkbookSheetStructure`
   * and layers the grid bookkeeping the XML does not carry (merges, sizes,
   * hidden lines, sparklines) on top.
   */
  function patchSheetState(
    change: XlsxStructureChange,
    merges: XlsxCellRange[],
    nextBounds: { maxCol: number; maxRow: number },
  ) {
    const assets = ctx.imageAssetsRef.value;
    const index = activeWorkbookSheetIndex();
    if (!assets || index < 0) {
      return;
    }
    const previous = assets.sheetStatesByWorkbookSheetIndex[index] ?? null;
    const sheetPath = assets.sheetPathsByWorkbookSheetIndex[index];
    const parsed = sheetPath
      ? parseSheetState(assets.archive, sheetPath, {
          includeCachedFormulaValues: false,
          themePalette: assets.themePalette
        })
      : null;
    const base: WorkbookSheetState | null = parsed
      ? {
          ...parsed,
          cachedFormulaValues: remapCachedFormulaValues(previous?.cachedFormulaValues, change),
          colStyleIds: remapNumberRecord(previous?.colStyleIds, change),
          colWidthOverridesPx: remapNumberRecord(previous?.colWidthOverridesPx, change),
          hiddenCols: remapIndexList(previous?.hiddenCols, change),
          hiddenRows: remapIndexList(previous?.hiddenRows, change),
          maxContentCol: Math.max(remapContentBound(previous?.maxContentCol, change) ?? 0, nextBounds.maxCol),
          maxContentRow: Math.max(remapContentBound(previous?.maxContentRow, change) ?? 0, nextBounds.maxRow),
          minContentCol: remapContentBound(previous?.minContentCol, change) ?? 0,
          minContentRow: remapContentBound(previous?.minContentRow, change) ?? 0,
          mergedRegions: merges,
          rowHeightOverridesPx: remapNumberRecord(previous?.rowHeightOverridesPx, change),
          rowStyleIds: remapNumberRecord(previous?.rowStyleIds, change),
          sparklines: (previous?.sparklines ?? []).flatMap((sparkline) => {
            const next = remapSparklineForStructure(sparkline, change);
            return next ? [next] : [];
          })
        }
      : null;
    assets.sheetStatesByWorkbookSheetIndex[index] = base;
  }

  function shiftDrawingAnchors(change: XlsxStructureChange) {
    const assets = ctx.imageAssetsRef.value;
    const index = activeWorkbookSheetIndex();
    if (!assets || index < 0) {
      return;
    }
    for (const image of assets.imagesByWorkbookSheetIndex[index] ?? []) {
      updateWorkbookImageAnchor(assets, image.id, remapImageAnchorForStructure(image.anchor, change));
    }
    for (const shape of assets.shapesByWorkbookSheetIndex[index] ?? []) {
      updateWorkbookShapeAnchor(assets, shape.id, remapImageAnchorForStructure(shape.anchor, change));
    }
    for (const formControl of assets.formControlsByWorkbookSheetIndex[index] ?? []) {
      updateWorkbookFormControl(assets, formControl.id, {
        anchor: remapImageAnchorForStructure(formControl.anchor, change)
      });
    }
    const chartAssets = ctx.chartAssetsRef.value;
    if (chartAssets) {
      for (const chart of chartAssets.chartsByWorkbookSheetIndex[index] ?? []) {
        const anchor = remapImageAnchorForStructure(chart.anchor, change);
        if (updateWorkbookChartAnchor(assets, chartAssets, chart.id, anchor)) {
          chart.anchor = anchor;
        }
      }
    }
  }

  function shiftDefinedNames(workbook: Workbook, change: XlsxStructureChange, sheetName: string) {
    const overrides = ctx.definedNameOverridesRef.value;
    const engineNames = normalizeWorkbookDefinedNames(workbook.namedRanges);
    for (const entry of engineNames) {
      const nextRefersTo = remapA1References(entry.refersTo, change, { sheetName });
      if (nextRefersTo !== entry.refersTo) {
        try {
          workbook.defineName(entry.name, nextRefersTo);
        } catch {
          // A rename that Excel rejects must not abort the structure edit.
        }
      }
    }
    const nextOverrides: typeof overrides = {};
    for (const [key, entry] of Object.entries(overrides)) {
      nextOverrides[key] = {
        ...entry,
        refersTo: remapA1References(entry.refersTo, change, { sheetName })
      };
    }
    ctx.definedNameOverridesRef.value = nextOverrides;
  }

  function remapSelection(change: XlsxStructureChange) {
    const selection = ctx.selection.value;
    if (selection) {
      const next = remapRangeForStructure(selection, change);
      if (next) {
        ctx.selection.value = next;
        ctx.selectionAnchorRef.value = next.start;
      } else {
        const origin = createCell(selection.start, change.axis, change.index);
        ctx.selection.value = { end: origin, start: origin };
        ctx.selectionAnchorRef.value = origin;
      }
    }
    const activeCell = ctx.activeCell.value;
    if (activeCell) {
      const next = remapCellForStructure(activeCell, change);
      ctx.activeCell.value = next ?? createCell(activeCell, change.axis, change.index);
    }
    const request = ctx.editRequestRef.value;
    if (request) {
      const next = remapCellForStructure(request.cell, change);
      ctx.editRequestRef.value = next ? { ...request, cell: next } : null;
    }
  }

  function applyStructure(
    axis: XlsxStructureAxis,
    kind: StructureKind,
    index: number,
    count: number,
  ) {
    const workbook = ctx.workbook.value;
    const sheet = ctx.activeSheet.value;
    const worksheet = ctx.getActiveWorksheet();
    if (!workbook || !sheet || !worksheet || !canMutateStructure()) {
      return false;
    }
    const change = normalizeRequest(axis, kind, index, count);
    const usedRange = worksheet.usedRange() as [number, number, number, number] | null;
    const bounds = {
      maxCol: Math.max(0, usedRange ? usedRange[3] : 0),
      maxRow: Math.max(0, usedRange ? usedRange[2] : 0)
    };
    const limit = axisLimit(axis);
    if (kind === "delete") {
      const usedEnd = axis === "row" ? bounds.maxRow : bounds.maxCol;
      if (change.index > usedEnd) {
        return false;
      }
      change.count = Math.min(change.count, usedEnd - change.index + 1, limit - change.index);
    } else if (change.index > (axis === "row" ? bounds.maxRow : bounds.maxCol) + 1) {
      return false;
    }

    ctx.recordHistoryBeforeMutation();
    const fallbackStyle = resolveCellFallbackStyle(sheet);

    const mergeRanges = resolveWorksheetMergedRegions(worksheet).map((range) => normalizeRange(range));
    for (const range of mergeRanges) {
      worksheet.unmergeCells(rangeToA1(range));
    }

    shiftCells(worksheet, change, bounds, fallbackStyle);
    shiftAxisSizes(worksheet, change, bounds);
    const nextMerges = mergeRanges.flatMap((range) => {
      const next = remapRangeForStructure(range, change);
      return next && (next.start.row !== next.end.row || next.start.col !== next.end.col)
        ? [next]
        : [];
    });
    for (const range of nextMerges) {
      worksheet.mergeCells(rangeToA1(range));
    }
    rewriteFormulas(worksheet, change, sheet.name);

    const assets = ctx.imageAssetsRef.value;
    const workbookSheetIndex = activeWorkbookSheetIndex();
    if (assets && workbookSheetIndex >= 0) {
      shiftWorkbookSheetStructure(assets, workbookSheetIndex, change, { sheetName: sheet.name });
    }
    shiftDrawingAnchors(change);
    const usedAfter = worksheet.usedRange() as [number, number, number, number] | null;
    patchSheetState(change, nextMerges, {
      maxCol: Math.max(0, usedAfter ? usedAfter[3] : 0),
      maxRow: Math.max(0, usedAfter ? usedAfter[2] : 0)
    });
    shiftDefinedNames(workbook, change, sheet.name);

    remapSelection(change);
    if (ctx.maybeRecalculateWorkbook(workbook)) {
      ctx.refreshWorkbookState(workbook);
    }
    return true;
  }

  function insertRows(startRow: number, count = 1) {
    return applyStructure("row", "insert", startRow, count);
  }

  function deleteRows(startRow: number, count = 1) {
    return applyStructure("row", "delete", startRow, count);
  }

  function insertColumns(startColumn: number, count = 1) {
    return applyStructure("column", "insert", startColumn, count);
  }

  function deleteColumns(startColumn: number, count = 1) {
    return applyStructure("column", "delete", startColumn, count);
  }

  function selectRangeOnSheet(range: XlsxCellRange) {
    const normalized = normalizeRange(range);
    ctx.selection.value = normalized;
    ctx.activeCell.value = normalized.end;
    ctx.selectionAnchorRef.value = normalized.start;
  }

  function selectRow(row: number) {
    const sheet = ctx.activeSheet.value;
    if (!sheet) {
      return;
    }
    const normalized = Math.max(0, Math.floor(Number(row) || 0));
    const lastCol = Math.max(sheet.maxUsedCol, sheet.visibleCols.at(-1) ?? 0);
    selectRangeOnSheet({ end: { col: lastCol, row: normalized }, start: { col: 0, row: normalized } });
  }

  function selectColumn(column: number) {
    const sheet = ctx.activeSheet.value;
    if (!sheet) {
      return;
    }
    const normalized = Math.max(0, Math.floor(Number(column) || 0));
    const lastRow = Math.max(sheet.maxUsedRow, sheet.visibleRows.at(-1) ?? 0);
    selectRangeOnSheet({ end: { col: normalized, row: lastRow }, start: { col: normalized, row: 0 } });
  }

  function selectAll() {
    const sheet = ctx.activeSheet.value;
    if (!sheet) {
      return;
    }
    const lastRow = Math.max(sheet.maxUsedRow, sheet.visibleRows.at(-1) ?? 0);
    const lastCol = Math.max(sheet.maxUsedCol, sheet.visibleCols.at(-1) ?? 0);
    selectRangeOnSheet({ end: { col: lastCol, row: lastRow }, start: { col: 0, row: 0 } });
  }

  function mergeAcrossSelection() {
    const worksheet = ctx.getActiveWorksheet();
    if (!worksheet || !ctx.selection.value || !ctx.workbook.value || ctx.isEditingLocked.value) {
      return;
    }
    const normalized = normalizeRange(ctx.selection.value);
    if (normalized.start.col === normalized.end.col) {
      return;
    }
    ctx.recordHistoryBeforeMutation();
    for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
      worksheet.mergeCells(rangeToA1({
        end: { col: normalized.end.col, row },
        start: { col: normalized.start.col, row }
      }));
    }
    const workbookSheetIndex = activeWorkbookSheetIndex();
    const assets = ctx.imageAssetsRef.value;
    const state = assets?.sheetStatesByWorkbookSheetIndex[workbookSheetIndex];
    if (assets && state) {
      assets.sheetStatesByWorkbookSheetIndex[workbookSheetIndex] = {
        ...state,
        mergedRegions: resolveWorksheetMergedRegions(worksheet).map((range) => normalizeRange(range))
      };
    }
    ctx.refreshWorkbookState(ctx.workbook.value);
  }

  function getNamedRanges() {
    const workbook = ctx.workbook.value;
    if (!workbook) {
      return [];
    }
    const engineNames = normalizeWorkbookDefinedNames(workbook.namedRanges);
    const removed = new Set(ctx.removedDefinedNamesRef.value.map((name) => name.toLowerCase()));
    const overrides = ctx.definedNameOverridesRef.value;
    const byName = new Map<string, import("@arcships/xlsx-core").XlsxNamedRange>();
    for (const entry of engineNames) {
      const key = entry.name.toLowerCase();
      if (removed.has(key)) {
        continue;
      }
      byName.set(key, overrides[key] ?? entry);
    }
    for (const [key, entry] of Object.entries(overrides)) {
      if (!removed.has(key)) {
        byName.set(key, entry);
      }
    }
    return [...byName.values()].sort((left, right) => left.name.localeCompare(right.name));
  }

  function setNamedRange(name: string, refersTo: string) {
    const workbook = ctx.workbook.value;
    const trimmedName = name.trim();
    const trimmedRefersTo = refersTo.trim();
    if (!workbook || !trimmedName || !trimmedRefersTo || ctx.isEditingLocked.value) {
      return false;
    }
    const previous = getNamedRanges().find(
      (entry) => entry.name.toLowerCase() === trimmedName.toLowerCase()
    );
    ctx.recordHistoryBeforeMutation();
    try {
      workbook.defineName(trimmedName, trimmedRefersTo);
    } catch {
      // The engine refuses to redefine an existing name; the Name Manager
      // overlay still carries the new reference and the export writes it.
      if (!previous) {
        return false;
      }
    }
    const key = trimmedName.toLowerCase();
    ctx.definedNameOverridesRef.value = {
      ...ctx.definedNameOverridesRef.value,
      [key]: {
        name: trimmedName,
        refersTo: trimmedRefersTo,
        sheetIndex: previous?.sheetIndex
      }
    };
    ctx.removedDefinedNamesRef.value = ctx.removedDefinedNamesRef.value.filter(
      (entry) => entry.toLowerCase() !== key
    );
    ctx.revision.value = ctx.revision.value + 1;
    return true;
  }

  function deleteNamedRange(name: string) {
    const trimmed = name.trim();
    if (!trimmed || ctx.isEditingLocked.value) {
      return false;
    }
    const key = trimmed.toLowerCase();
    if (!getNamedRanges().some((entry) => entry.name.toLowerCase() === key)) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    const nextOverrides = { ...ctx.definedNameOverridesRef.value };
    delete nextOverrides[key];
    ctx.definedNameOverridesRef.value = nextOverrides;
    ctx.removedDefinedNamesRef.value = [...ctx.removedDefinedNamesRef.value, trimmed];
    ctx.revision.value = ctx.revision.value + 1;
    return true;
  }

  return {
    deleteColumns,
    deleteNamedRange,
    deleteRows,
    getNamedRanges,
    insertColumns,
    insertRows,
    mergeAcrossSelection,
    selectAll,
    selectColumn,
    selectRow,
    setNamedRange
  };
}
