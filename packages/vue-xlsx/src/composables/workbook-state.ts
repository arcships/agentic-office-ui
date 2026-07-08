import { computed } from "vue";
import type { Workbook } from "@dukelib/sheets-wasm";
import {
  getSheetsWasmModule,
  loadWorkbookChartAssets,
  mergeWorkbookImageAssets,
  pxToSheetColumnWidth,
  resolveContentSheetAxisPixels,
  resolveWorksheetDefaultColumnWidthPixels,
  resolveWorksheetDefaultRowHeightPixels,
  resolveWorksheetHiddenCols,
  resolveWorksheetHiddenRows,
  resolveWorksheetMergeMetadata,
  resolveSheetColumnWidthPixels,
  resolveSheetRowHeightPixels,
  revokeWorkbookImageAssets,
  safeCalculate,
  tryRecalculate,
  type XlsxConditionalFormatRule,
  type XlsxResolvedCellStyle,
  type XlsxSheetData,
  type XlsxSheetVisibility,
  type XlsxSparkline,
  type XlsxTable,
  type XlsxTableSortDirection,
  type XlsxTableStyleDefinition,
  type XlsxThemePalette,
  type XlsxWorkbookTab
} from "@extend-ai/xlsx-core";
import type { UseXlsxViewerControllerOptions } from "@extend-ai/xlsx-core";
import {
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  DEFAULT_ZOOM_SCALE,
  DEFAULT_ZOOM_TAB_KEY,
  FORMULA_COUNT_THRESHOLD,
  MAX_ZOOM_SCALE,
  MIN_ROW_HEIGHT_PX,
  MIN_ZOOM_SCALE,
  XLSX_MIME_TYPE,
  CSV_MIME_TYPE,
  ZOOM_STEP,
  normalizeWorksheetVisibility,
  createAbortError,
  isAbortError,
  pushHistoryEntry,
  type CellMutationState,
  type HistoryEntry,
  type RangeCellMutation,
  type SnapshotHistoryEntry,
  type CellEditHistoryEntry,
  type RangeEditHistoryEntry
} from "./internal";
import { parseWorksheetDataValidations, parseWorksheetFreezePanes, parseA1RangeReference, rangeToA1, cellAddressToA1, normalizeRange } from "./selection";
import {
  XlsxFileSizeLimitExceededError,
  cloneBytes,
  createWorkbookTooLargeError,
  decodeHtmlEntities,
  fileStem,
  pxToSheetRowHeight,
  preflightWorkbookBuffer,
  resolveDisplayFileName,
  sanitizeSavedWorkbookBytes
} from "./formatting";
import { applyCellMutationState } from "./internal";
import { loadWorkbookImageAssets } from "./image-assets";
import { downloadArrayBuffer, downloadBytes, downloadText, downloadUrl } from "./clipboard";
import type { XlsxControllerContext } from "./navigation";

export {
  resolveDisplayFileName,
  XlsxFileSizeLimitExceededError,
  createWorkbookTooLargeError,
  preflightWorkbookBuffer,
  normalizeWorksheetVisibility
};

export function clampZoomScale(zoomScale: number) {
  if (!Number.isFinite(zoomScale)) {
    return DEFAULT_ZOOM_SCALE;
  }

  return Math.min(MAX_ZOOM_SCALE, Math.max(MIN_ZOOM_SCALE, Math.round(zoomScale)));
}

export function resolveDefaultZoomScale(activeTab: XlsxWorkbookTab | null, activeSheet: XlsxSheetData | null) {
  if (activeTab?.kind !== "sheet") {
    return DEFAULT_ZOOM_SCALE;
  }

  return clampZoomScale(activeSheet?.zoomScale ?? DEFAULT_ZOOM_SCALE);
}

export function resolveWorksheetZoomScale(
  worksheet: ReturnType<Workbook["getSheet"]>,
  sheetState?: Record<string, unknown> | null
) {
  const candidates = [
    typeof sheetState?.zoomScale === "number" ? sheetState.zoomScale : undefined,
    typeof worksheet.zoomScale === "number" ? worksheet.zoomScale : undefined
  ];
  const value = candidates.find((entry): entry is number => typeof entry === "number" && Number.isFinite(entry) && entry > 0);
  return clampZoomScale(value ?? DEFAULT_ZOOM_SCALE);
}

export function resolveNextZoomScale(currentZoomScale: number, direction: 1 | -1) {
  if (direction > 0) {
    return Math.min(
      MAX_ZOOM_SCALE,
      currentZoomScale % ZOOM_STEP === 0
        ? currentZoomScale + ZOOM_STEP
        : Math.ceil(currentZoomScale / ZOOM_STEP) * ZOOM_STEP
    );
  }

  return Math.max(
    MIN_ZOOM_SCALE,
    currentZoomScale % ZOOM_STEP === 0
      ? currentZoomScale - ZOOM_STEP
      : Math.floor(currentZoomScale / ZOOM_STEP) * ZOOM_STEP
  );
}

export { DEFAULT_ZOOM_TAB_KEY, MAX_ZOOM_SCALE, MIN_ZOOM_SCALE };

export function resolveSheetDisplayUsedRange(
  usedRange: [number, number, number, number],
  sheetState?: {
    maxContentCol?: number;
    maxContentRow?: number;
    maxHorizontalMergeEndCol?: number;
    maxVerticalMergeEndRow?: number;
    minContentCol?: number;
    minContentRow?: number;
  } | null
): [number, number, number, number] {
  const [minRow, minCol, maxRow, maxCol] = usedRange;
  const maxContentRow = sheetState?.maxContentRow ?? -1;
  const maxContentCol = sheetState?.maxContentCol ?? -1;
  const maxVerticalMergeEndRow = sheetState?.maxVerticalMergeEndRow ?? -1;
  const maxHorizontalMergeEndCol = sheetState?.maxHorizontalMergeEndCol ?? -1;
  const maxMeaningfulRow = Math.max(maxContentRow, maxVerticalMergeEndRow);
  const maxMeaningfulCol = Math.max(maxContentCol, maxHorizontalMergeEndCol);

  if (maxMeaningfulRow < 0 && maxMeaningfulCol < 0) {
    return usedRange;
  }

  return [
    sheetState?.minContentRow !== undefined && sheetState.minContentRow >= 0 ? Math.min(minRow, sheetState.minContentRow) : minRow,
    sheetState?.minContentCol !== undefined && sheetState.minContentCol >= 0 ? Math.min(minCol, sheetState.minContentCol) : minCol,
    maxMeaningfulRow >= 0
      ? (maxContentRow >= 0 ? Math.min(maxRow, maxMeaningfulRow) : Math.max(maxRow, maxMeaningfulRow))
      : maxRow,
    maxMeaningfulCol >= 0
      ? (maxContentCol >= 0 ? Math.min(maxCol, maxMeaningfulCol) : Math.max(maxCol, maxMeaningfulCol))
      : maxCol
  ];
}

export function buildSheetList(
  workbook: Workbook,
  sheetStatesByWorkbookSheetIndex?: Array<{
    cachedFormulaValues?: Record<string, string>;
    columnWidthCharacterWidthPx?: number;
    colWidthOverridesPx?: Record<number, number>;
    colStyleIds?: Record<number, number>;
    conditionalFormatRules?: XlsxConditionalFormatRule[];
    defaultColWidthPx?: number;
    defaultRowHeightPx?: number;
    hasHorizontalMerges?: boolean;
    hasVerticalMerges?: boolean;
    maxHorizontalMergeEndCol?: number;
    maxVerticalMergeEndRow?: number;
    maxContentCol?: number;
    maxContentRow?: number;
    minContentCol?: number;
    minContentRow?: number;
    hiddenCols?: number[];
    hiddenRows?: number[];
    rowHeightOverridesPx?: Record<number, number>;
    rowStyleIds?: Record<number, number>;
    showGridLines: boolean;
    sparklines?: XlsxSparkline[];
  } | null>,
  themePalette?: XlsxThemePalette | null,
  styleById?: Record<number, XlsxResolvedCellStyle> | null,
  namedCellStyleByName?: Record<string, XlsxResolvedCellStyle> | null,
  tableStyleByName?: Record<string, XlsxTableStyleDefinition> | null,
  showHiddenSheets = false
): XlsxSheetData[] {
  const sheets: XlsxSheetData[] = [];

  for (let index = 0; index < workbook.sheetCount; index += 1) {
    const worksheet = workbook.getSheet(index);
    const sheetState = sheetStatesByWorkbookSheetIndex?.[index] ?? null;
    const mergeMetadata = resolveWorksheetMergeMetadata(worksheet);
    const effectiveSheetState = {
      ...sheetState,
      ...mergeMetadata
    };
    const defaultColWidthPx = resolveWorksheetDefaultColumnWidthPixels(
      worksheet,
      sheetState?.columnWidthCharacterWidthPx,
      sheetState?.defaultColWidthPx ?? DEFAULT_COL_WIDTH
    );
    const defaultRowHeightPx = resolveWorksheetDefaultRowHeightPixels(
      worksheet,
      sheetState?.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT
    );
    const visibility = normalizeWorksheetVisibility(worksheet.visibility);
    if (!showHiddenSheets && visibility !== "visible") {
      continue;
    }

    const resolveColumnWidthPx = (col: number) => {
      const width = worksheet.getColumnWidth(col);
      if (width !== undefined && width !== null) {
        return resolveSheetColumnWidthPixels(width, sheetState?.columnWidthCharacterWidthPx);
      }

      return sheetState?.colWidthOverridesPx?.[col] ?? defaultColWidthPx;
    };

    const resolveRowHeightPx = (row: number) => {
      const height = worksheet.getRowHeight(row);
      if (height !== undefined && height !== null) {
        return Math.max(Math.round(height * 1.33), MIN_ROW_HEIGHT_PX);
      }

      return sheetState?.rowHeightOverridesPx?.[row] ?? defaultRowHeightPx;
    };

    const usedRange = worksheet.usedRange() as [number, number, number, number] | null;
    if (!usedRange) {
      sheets.push({
        cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
        columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
        colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
        colStyleIds: sheetState?.colStyleIds ?? {},
        conditionalFormatRules: sheetState?.conditionalFormatRules ?? [],
        dataValidations: parseWorksheetDataValidations(worksheet),
        defaultColWidthPx,
        defaultRowHeightPx,
        freezePanes: parseWorksheetFreezePanes(worksheet),
        hasHorizontalMerges: mergeMetadata.hasHorizontalMerges,
        hasVerticalMerges: mergeMetadata.hasVerticalMerges,
        maxHorizontalMergeEndCol: mergeMetadata.maxHorizontalMergeEndCol,
        maxVerticalMergeEndRow: mergeMetadata.maxVerticalMergeEndRow,
        hiddenCols: [],
        hiddenRows: [],
        minUsedCol: -1,
        minUsedRow: -1,
        maxUsedCol: -1,
        maxUsedRow: -1,
        name: worksheet.name,
        visibility,
        namedCellStyleByName: namedCellStyleByName ?? {},
        rowCount: 0,
        colCount: 0,
        rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
        rowStyleIds: sheetState?.rowStyleIds ?? {},
        styleById: styleById ?? {},
        sparklines: sheetState?.sparklines ?? [],
        tableStyleByName: tableStyleByName ?? {},
        visibleRows: [],
        visibleCols: [],
        colWidths: [],
        rowHeights: [],
        showGridLines: sheetState?.showGridLines ?? true,
        themePalette: themePalette ?? { colorsByIndex: {} },
        workbookSheetIndex: index,
        zoomScale: resolveWorksheetZoomScale(worksheet, sheetState)
      });
      continue;
    }

    const [minRow, minCol, maxRow, maxCol] = resolveSheetDisplayUsedRange(usedRange, effectiveSheetState);
    let visibleRowsCache: number[] | null = null;
    let visibleColsCache: number[] | null = null;
    let rowHeightsCache: number[] | null = null;
    let colWidthsCache: number[] | null = null;

    const getVisibleRows = () => {
      if (visibleRowsCache) {
        return visibleRowsCache;
      }

      const nextVisibleRows: number[] = [];
      for (let row = 0; row <= maxRow; row += 1) {
        if (!worksheet.isRowHidden(row)) {
          nextVisibleRows.push(row);
        }
      }

      visibleRowsCache = nextVisibleRows;
      return nextVisibleRows;
    };

    const getVisibleCols = () => {
      if (visibleColsCache) {
        return visibleColsCache;
      }

      const nextVisibleCols: number[] = [];
      for (let col = 0; col <= maxCol; col += 1) {
        if (!worksheet.isColumnHidden(col)) {
          nextVisibleCols.push(col);
        }
      }

      visibleColsCache = nextVisibleCols;
      return nextVisibleCols;
    };

    const getRowHeights = () => {
      if (rowHeightsCache) {
        return rowHeightsCache;
      }

      rowHeightsCache = getVisibleRows().map(resolveRowHeightPx);
      return rowHeightsCache;
    };

    const getColWidths = () => {
      if (colWidthsCache) {
        return colWidthsCache;
      }

      colWidthsCache = getVisibleCols().map(resolveColumnWidthPx);
      return colWidthsCache;
    };

    const sheet: XlsxSheetData = {
      cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
      columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
      colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
      colStyleIds: sheetState?.colStyleIds ?? {},
      conditionalFormatRules: sheetState?.conditionalFormatRules ?? [],
      dataValidations: parseWorksheetDataValidations(worksheet),
      defaultColWidthPx,
      defaultRowHeightPx,
      freezePanes: parseWorksheetFreezePanes(worksheet),
      hasHorizontalMerges: mergeMetadata.hasHorizontalMerges,
      hasVerticalMerges: mergeMetadata.hasVerticalMerges,
      maxHorizontalMergeEndCol: mergeMetadata.maxHorizontalMergeEndCol,
      maxVerticalMergeEndRow: mergeMetadata.maxVerticalMergeEndRow,
      hiddenCols: resolveWorksheetHiddenCols(worksheet, maxCol),
      hiddenRows: resolveWorksheetHiddenRows(worksheet, maxRow),
      minUsedCol: minCol,
      minUsedRow: minRow,
      maxUsedCol: maxCol,
      maxUsedRow: maxRow,
      name: worksheet.name,
      visibility,
      namedCellStyleByName: namedCellStyleByName ?? {},
      rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
      rowStyleIds: sheetState?.rowStyleIds ?? {},
      showGridLines: sheetState?.showGridLines ?? true,
      styleById: styleById ?? {},
      sparklines: sheetState?.sparklines ?? [],
      tableStyleByName: tableStyleByName ?? {},
      themePalette: themePalette ?? { colorsByIndex: {} },
      workbookSheetIndex: index,
      zoomScale: resolveWorksheetZoomScale(worksheet, sheetState),
      get rowCount() {
        return getVisibleRows().length;
      },
      get colCount() {
        return getVisibleCols().length;
      },
      get visibleRows() {
        return getVisibleRows();
      },
      get visibleCols() {
        return getVisibleCols();
      },
      get colWidths() {
        return getColWidths();
      },
      get rowHeights() {
        return getRowHeights();
      }
    };

    sheets.push(sheet);
  }

  return sheets;
}

export function buildVisibleSheetIndexMap(sheets: XlsxSheetData[]) {
  return new Map(sheets.map((sheet, index) => [sheet.workbookSheetIndex, index]));
}

export { resolveInheritedCellStyle } from "./formatting";

export function mapWorksheetTables(worksheet: ReturnType<Workbook["getSheet"]> | null): XlsxTable[] {
  const rawTables = (worksheet?.tables ?? []) as Array<Record<string, unknown>>;
  return rawTables.flatMap((table, index) => {
    const rawColumns = Array.isArray(table.columns) ? table.columns : [];
    const rawName = typeof table.name === "string" ? table.name : `Table${index + 1}`;
    const rawDisplayName =
      typeof table.displayName === "string"
        ? table.displayName
        : typeof table.name === "string"
          ? table.name
          : `Table ${index + 1}`;
    const rawReference = typeof table.reference === "string" ? table.reference : "";
    const reference = rawReference;
    const parsedRange = parseA1RangeReference(reference);
    if (!parsedRange) {
      return [];
    }

    return [{
      columns: rawColumns.map((column, columnIndex) => ({
        id: typeof (column as { id?: unknown }).id === "number" ? ((column as { id?: number }).id ?? columnIndex + 1) : columnIndex + 1,
        index: columnIndex,
        name: typeof (column as { name?: unknown }).name === "string" ? ((column as { name?: string }).name ?? `Column ${columnIndex + 1}`) : `Column ${columnIndex + 1}`
      })),
      displayName: rawDisplayName,
      end: parsedRange.end,
      headerRowCount: resolveWorkbookTableCount(table.headerRowCount, 1),
      headerRowCellStyle: typeof table.headerRowCellStyle === "string" ? table.headerRowCellStyle : undefined,
      name: rawName,
      reference,
      start: parsedRange.start,
      styleInfo: table.styleInfo as XlsxTable["styleInfo"] | undefined,
      totalsRowCount: resolveWorkbookTableCount(table.totalsRowCount, 0),
      totalsRowShown: resolveWorkbookTableBoolean(table.totalsRowShown)
    }];
  });
}

export function resolveWorkbookTableCount(value: unknown, fallback: number) {
  if (typeof value === "number" && Number.isFinite(value) && value >= 0) {
    return value;
  }

  if (typeof value === "string") {
    const parsed = Number.parseInt(value, 10);
    if (Number.isFinite(parsed) && parsed >= 0) {
      return parsed;
    }
  }

  return fallback;
}

export function resolveWorkbookTableBoolean(value: unknown) {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "number") {
    return value !== 0;
  }

  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (normalized === "0" || normalized === "false" || normalized === "") {
      return false;
    }
    if (normalized === "1" || normalized === "true") {
      return true;
    }
  }

  return false;
}

export async function resolveWorkbookBuffer(
  { file, src }: UseXlsxViewerControllerOptions,
  signal?: AbortSignal
): Promise<ArrayBuffer> {
  let buffer: ArrayBuffer;

  if (signal?.aborted) {
    throw createAbortError();
  }

  if (file) {
    buffer = file;
  } else if (src) {
    let response: Response;
    try {
      response = await fetch(src, { signal });
    } catch (error) {
      if (isAbortError(error)) {
        throw error;
      }

      throw new Error(
        "Failed to fetch workbook. The remote URL may be blocked by CORS, unavailable, or not directly downloadable from the browser."
      );
    }
    if (!response.ok) {
      throw new Error(`Failed to fetch workbook (status ${response.status})`);
    }
    buffer = await response.arrayBuffer();
  } else {
    throw new Error("Either `file` or `src` must be provided.");
  }

  return buffer;
}

export async function parseWorkbookBuffer(buffer: ArrayBuffer): Promise<{
  shouldAutoCalculate: boolean;
  workbook: Workbook;
}> {
  const wasmModule = await getSheetsWasmModule();
  const initialWorkbook = wasmModule.Workbook.fromBytes(new Uint8Array(buffer));
  let totalFormulas = 0;

  for (let index = 0; index < initialWorkbook.sheetCount; index += 1) {
    totalFormulas += initialWorkbook.getSheet(index).formulaCount;
  }

  const shouldAutoCalculate = totalFormulas <= FORMULA_COUNT_THRESHOLD;
  if (!shouldAutoCalculate) {
    return { shouldAutoCalculate, workbook: initialWorkbook };
  }

  const result = safeCalculate(initialWorkbook, {
    reparse: () => wasmModule.Workbook.fromBytes(new Uint8Array(buffer))
  });

  return {
    shouldAutoCalculate: result.calculated,
    workbook: result.workbook
  };
}

export { tryRecalculate };

export type { XlsxSheetVisibility };

export function createHistoryDomain(ctx: XlsxControllerContext) {
  const tables = computed(() => { return (
      ctx.isWorkerBacked.value
        ? ctx.workerTablesByWorkbookSheetIndex.value[ctx.activeSheet.value?.workbookSheetIndex ?? -1] ?? []
        : mapWorksheetTables(ctx.getActiveWorksheet())
    ); })

  function getCellSnapshotAsync(workbookSheetIndex: number, row: number, col: number) {
    if (!ctx.isWorkerBacked.value) {
      return Promise.resolve({
        displayValue: "",
        formula: ""
      });
    }

    return ctx.getWorkerClient().getCellSnapshot(workbookSheetIndex, row, col);
  }

  function getRowsBatchAsync(workbookSheetIndex: number, startRow: number, rowCount: number) {
    if (!ctx.isWorkerBacked.value) {
      return Promise.resolve(null);
    }

    return ctx.getWorkerClient().getRowsBatch(workbookSheetIndex, startRow, rowCount);
  }

  function createHistoryEntry(): SnapshotHistoryEntry | null {
    if (!ctx.workbook.value) {
      return null;
    }

    return {
      kind: "snapshot",
      activeCell: ctx.activeCell.value,
      activeSheetIndex: ctx.activeSheetIndex.value,
      bytes: ctx.createSavedWorkbookBytes(ctx.workbook.value),
      selection: ctx.selection.value
    };
  }

  async function restoreHistoryEntry(entry: SnapshotHistoryEntry) {
    const wasmModule = await getSheetsWasmModule();
    const nextWorkbook = wasmModule.Workbook.fromBytes(cloneBytes(entry.bytes));
    const nextImageAssets = loadWorkbookImageAssets(entry.bytes, nextWorkbook);
    const nextSheets = buildSheetList(
      nextWorkbook,
      nextImageAssets.sheetStatesByWorkbookSheetIndex,
      nextImageAssets.themePalette,
      nextImageAssets.styleById,
      nextImageAssets.namedCellStyleByName,
      nextImageAssets.tableStyleByName,
      ctx.showHiddenSheets
    );
    const nextSheetIndex = Math.max(0, Math.min(entry.activeSheetIndex, Math.max(0, nextSheets.length - 1)));

    ctx.error.value = null;
    ctx.isLoading.value = false;
    ctx.setImageAssets(nextImageAssets);
    ctx.workbook.value = nextWorkbook;
    ctx.sheets.value = nextSheets;
    const nextChartAssets = loadWorkbookChartAssets(nextWorkbook, nextImageAssets, buildVisibleSheetIndexMap(nextSheets), ctx.showHiddenSheets);
    ctx.setChartAssets(nextChartAssets);
    ctx.activeSheetIndex.value = nextSheetIndex;
    const nextTabIndex = nextChartAssets.tabs.findIndex((tab) => tab.kind === "sheet" && tab.sheetIndex === nextSheetIndex);
    if (nextTabIndex >= 0) {
      ctx.activeTabIndex.value = nextTabIndex;
    }
    ctx.activeCell.value = entry.activeCell;
    ctx.selection.value = entry.selection;
    ctx.selectionAnchorRef.value = entry.selection ? normalizeRange(entry.selection).start : entry.activeCell;
    ctx.revision.value = ctx.revision.value + 1;
  }

  function applyCellEditHistoryEntry(entry: CellEditHistoryEntry, direction: "undo" | "redo") {
    if (!ctx.workbook.value) {
      return;
    }

    const worksheet = ctx.workbook.value.getSheet(entry.sheetIndex);
    const visibleSheetIndex = ctx.sheets.value.findIndex((sheet) => sheet.workbookSheetIndex === entry.sheetIndex);
    const targetState = direction === "undo" ? entry.before : entry.after;

    ctx.isApplyingHistoryRef.value = true;
    applyCellMutationState(worksheet, entry.cell, targetState);
    ctx.maybeRecalculateWorkbook(ctx.workbook.value);
    ctx.refreshWorkbookState(ctx.workbook.value);

    const nextActiveCell = direction === "undo" ? entry.activeCellBefore : entry.activeCellAfter;
    const nextSelection = direction === "undo" ? entry.selectionBefore : entry.selectionAfter;
    if (visibleSheetIndex >= 0) {
      ctx.activeSheetIndex.value = visibleSheetIndex;
    }
    ctx.activeCell.value = nextActiveCell;
    ctx.selection.value = nextSelection;
    ctx.selectionAnchorRef.value = nextSelection ? normalizeRange(nextSelection).start : nextActiveCell;
    ctx.isApplyingHistoryRef.value = false;
  }

  function applyRangeEditHistoryEntry(entry: RangeEditHistoryEntry, direction: "undo" | "redo") {
    if (!ctx.workbook.value) {
      return;
    }

    const worksheet = ctx.workbook.value.getSheet(entry.sheetIndex);
    const visibleSheetIndex = ctx.sheets.value.findIndex((sheet) => sheet.workbookSheetIndex === entry.sheetIndex);

    ctx.isApplyingHistoryRef.value = true;
    for (const mutation of entry.mutations) {
      applyCellMutationState(worksheet, mutation.cell, direction === "undo" ? mutation.before : mutation.after);
    }
    ctx.maybeRecalculateWorkbook(ctx.workbook.value);
    ctx.refreshWorkbookState(ctx.workbook.value);

    const nextActiveCell = direction === "undo" ? entry.activeCellBefore : entry.activeCellAfter;
    const nextSelection = direction === "undo" ? entry.selectionBefore : entry.selectionAfter;
    if (visibleSheetIndex >= 0) {
      ctx.activeSheetIndex.value = visibleSheetIndex;
    }
    ctx.activeCell.value = nextActiveCell;
    ctx.selection.value = nextSelection;
    ctx.selectionAnchorRef.value = nextSelection ? normalizeRange(nextSelection).start : nextActiveCell;
    ctx.isApplyingHistoryRef.value = false;
  }

  function sortTable(tableName: string, columnIndex: number, direction: XlsxTableSortDirection) {
    const worksheet = ctx.getActiveWorksheet();
    const targetTable = tables.value.find((table) => table.name === tableName || table.displayName === tableName);
    if (!worksheet || !ctx.workbook.value || !ctx.activeSheet.value || !targetTable) {
      return;
    }

    const dataStartRow = targetTable.start.row + Math.max(targetTable.headerRowCount, 1);
    const totalsRowOffset = targetTable.totalsRowShown ? Math.max(targetTable.totalsRowCount, 1) : 0;
    const dataEndRow = targetTable.end.row - totalsRowOffset;
    const startCol = targetTable.start.col;
    const endCol = targetTable.end.col;
    const sortCol = startCol + columnIndex;

    if (columnIndex < 0 || sortCol > endCol || dataStartRow > dataEndRow) {
      return;
    }

    const rows: Array<{
      cells: CellMutationState[];
      index: number;
      sortBoolean: boolean | undefined;
      sortEmpty: boolean;
      sortNumber: number | undefined;
      sortText: string;
    }> = [];

    for (let row = dataStartRow; row <= dataEndRow; row += 1) {
      const cells: CellMutationState[] = [];
      for (let col = startCol; col <= endCol; col += 1) {
        cells.push({
          formula: worksheet.getFormulaAt(row, col) ?? null,
          style: worksheet.getCellStyleAt(row, col),
          value: worksheet.getCellAt(row, col).toJs()
        });
      }

      const calculated = worksheet.getCalculatedValueAt(row, sortCol);
      const formatted = decodeHtmlEntities(worksheet.getFormattedValueAt(row, sortCol) ?? "");
      rows.push({
        cells,
        index: row,
        sortBoolean: calculated.is_boolean ? calculated.asBoolean() : undefined,
        sortEmpty: calculated.is_empty || formatted.length === 0,
        sortNumber: calculated.is_number ? calculated.asNumber() : undefined,
        sortText: calculated.is_text ? (calculated.asText() ?? formatted) : formatted
      });
    }

    const sortedRows = [...rows].sort((left, right) => {
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

    if (sortedRows.every((row, index) => row.index === rows[index]?.index)) {
      ctx.sortState.value = { columnIndex, direction, tableName: targetTable.name };
      return;
    }

    const mutations: RangeCellMutation[] = [];
    for (let rowOffset = 0; rowOffset < rows.length; rowOffset += 1) {
      const targetRow = dataStartRow + rowOffset;
      const sourceRow = sortedRows[rowOffset];
      const beforeRow = rows[rowOffset];
      if (!sourceRow || !beforeRow) {
        continue;
      }

      for (let colOffset = 0; colOffset <= endCol - startCol; colOffset += 1) {
        const before = beforeRow.cells[colOffset];
        const after = sourceRow.cells[colOffset];
        if (!before || !after) {
          continue;
        }

        const cell = { row: targetRow, col: startCol + colOffset };
        applyCellMutationState(worksheet, cell, after);
        mutations.push({
          after,
          before,
          cell
        });
      }
    }

    ctx.maybeRecalculateWorkbook(ctx.workbook.value);
    ctx.refreshWorkbookState(ctx.workbook.value);
    ctx.sortState.value = { columnIndex, direction, tableName: targetTable.name };
    ctx.recordRangeEditHistory(mutations, ctx.selection.value, ctx.activeCell.value);
  }

  function download() {
    if (ctx.file) {
      downloadArrayBuffer(ctx.file, ctx.displayFileName.value);
      return;
    }

    if (ctx.src) {
      downloadUrl(ctx.src, ctx.displayFileName.value);
    }
  }

  function exportXlsx() {
    if (!ctx.workbook.value) {
      return;
    }

    downloadBytes(ctx.createSavedWorkbookBytes(ctx.workbook.value), `${fileStem(ctx.displayFileName.value)}.xlsx`, XLSX_MIME_TYPE);
  }

  function exportCsv() {
    if (!ctx.workbook.value) {
      return;
    }

    const activeSheetName = ctx.activeSheet.value?.name ?? "sheet";
    downloadText(ctx.workbook.value.saveCsvString(), `${fileStem(ctx.displayFileName.value)}-${activeSheetName}.csv`, CSV_MIME_TYPE);
  }

  function recalculate() {
    if (!ctx.workbook.value) {
      return;
    }

    const result = tryRecalculate(ctx.workbook.value);
    if (result.calculated) {
      ctx.refreshWorkbookState(ctx.workbook.value);
      return;
    }

    // Trap poisons the Workbook pointer; skip refreshWorkbookState so we
    // don't crash reading cells from it.
    ctx.shouldAutoCalculate.value = false;
  }

  function applyReadOnlyResizeOverride(axis: "column" | "row", actualIndex: number, sizePx: number) {
    if (!ctx.activeSheet.value) {
      return;
    }

    const contentSizePx = resolveContentSheetAxisPixels(sizePx, ctx.activeSheet.value.showGridLines);
    const targetWorkbookSheetIndex = ctx.activeSheet.value.workbookSheetIndex;
    ctx.sheets.value = ctx.sheets.value.map((sheet) => {
      if (sheet.workbookSheetIndex !== targetWorkbookSheetIndex) {
        return sheet;
      }

      if (axis === "column") {
        const nextColWidthOverridesPx = {
          ...sheet.colWidthOverridesPx,
          [actualIndex]: contentSizePx
        };
        const nextColWidths = [...sheet.colWidths];
        const visibleColIndex = sheet.visibleCols.indexOf(actualIndex);
        if (visibleColIndex >= 0) {
          nextColWidths[visibleColIndex] = contentSizePx;
        }

        return {
          ...sheet,
          colWidthOverridesPx: nextColWidthOverridesPx,
          colWidths: nextColWidths
        };
      }

      const nextRowHeightOverridesPx = {
        ...sheet.rowHeightOverridesPx,
        [actualIndex]: contentSizePx
      };
      const nextRowHeights = [...sheet.rowHeights];
      const visibleRowIndex = sheet.visibleRows.indexOf(actualIndex);
      if (visibleRowIndex >= 0) {
        nextRowHeights[visibleRowIndex] = contentSizePx;
      }

      return {
        ...sheet,
        rowHeightOverridesPx: nextRowHeightOverridesPx,
        rowHeights: nextRowHeights
      };
    });
    ctx.revision.value = ctx.revision.value + 1;
  }

  function resizeColumn(col: number, widthPx: number) {
    if ((ctx.readOnly.value && !ctx.canResizeReadOnly.value) || !ctx.activeSheet.value) {
      return;
    }

    if (ctx.isWorkerBacked.value) {
      applyReadOnlyResizeOverride("column", col, widthPx);
      return;
    }

    if (!ctx.workbook.value) {
      return;
    }

    ctx.recordHistoryBeforeMutation();
    const worksheet = ctx.workbook.value.getSheet(ctx.activeSheet.value.workbookSheetIndex);
    worksheet.setColumnWidth(
      col,
      pxToSheetColumnWidth(resolveContentSheetAxisPixels(widthPx, ctx.activeSheet.value.showGridLines))
    );
    ctx.refreshWorkbookState(ctx.workbook.value);
  }

  function resizeRow(row: number, heightPx: number) {
    if ((ctx.readOnly.value && !ctx.canResizeReadOnly.value) || !ctx.activeSheet.value) {
      return;
    }

    if (ctx.isWorkerBacked.value) {
      applyReadOnlyResizeOverride("row", row, heightPx);
      return;
    }

    if (!ctx.workbook.value) {
      return;
    }

    ctx.recordHistoryBeforeMutation();
    const worksheet = ctx.workbook.value.getSheet(ctx.activeSheet.value.workbookSheetIndex);
    worksheet.setRowHeight(
      row,
      pxToSheetRowHeight(resolveContentSheetAxisPixels(heightPx, ctx.activeSheet.value.showGridLines))
    );
    ctx.refreshWorkbookState(ctx.workbook.value);
  }

  function undo() {
    if (ctx.readOnly.value || !ctx.workbook.value || ctx.undoStackRef.value.length === 0) {
      return;
    }

    const entry = ctx.undoStackRef.value.pop();
    if (!entry) {
      return;
    }

    if (entry.kind === "cell-edit") {
      pushHistoryEntry(ctx.redoStackRef.value, entry);
      ctx.historyRevision.value = ctx.historyRevision.value + 1;
      applyCellEditHistoryEntry(entry, "undo");
      return;
    }

    if (entry.kind === "range-edit") {
      pushHistoryEntry(ctx.redoStackRef.value, entry);
      ctx.historyRevision.value = ctx.historyRevision.value + 1;
      applyRangeEditHistoryEntry(entry, "undo");
      return;
    }

    const currentSnapshot = createHistoryEntry();
    if (currentSnapshot) {
      pushHistoryEntry(ctx.redoStackRef.value, currentSnapshot);
    }
    ctx.historyRevision.value = ctx.historyRevision.value + 1;
    void restoreHistoryEntry(entry);
  }

  function redo() {
    if (ctx.readOnly.value || !ctx.workbook.value || ctx.redoStackRef.value.length === 0) {
      return;
    }

    const entry = ctx.redoStackRef.value.pop();
    if (!entry) {
      return;
    }

    if (entry.kind === "cell-edit") {
      pushHistoryEntry(ctx.undoStackRef.value, entry);
      ctx.historyRevision.value = ctx.historyRevision.value + 1;
      applyCellEditHistoryEntry(entry, "redo");
      return;
    }

    if (entry.kind === "range-edit") {
      pushHistoryEntry(ctx.undoStackRef.value, entry);
      ctx.historyRevision.value = ctx.historyRevision.value + 1;
      applyRangeEditHistoryEntry(entry, "redo");
      return;
    }

    const currentSnapshot = createHistoryEntry();
    if (currentSnapshot) {
      pushHistoryEntry(ctx.undoStackRef.value, currentSnapshot);
    }
    ctx.historyRevision.value = ctx.historyRevision.value + 1;
    void restoreHistoryEntry(entry);
  }

  return {
    tables,
    getCellSnapshotAsync,
    getRowsBatchAsync,
    sortTable,
    download,
    exportXlsx,
    exportCsv,
    recalculate,
    resizeColumn,
    resizeRow,
    undo,
    redo
  };
}
