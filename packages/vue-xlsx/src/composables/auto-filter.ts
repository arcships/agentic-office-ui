import { computed, ref, watch, type ComputedRef } from "vue";
import type {
  XlsxAutoFilterColumnState,
  XlsxAutoFilterRange,
  XlsxAutoFilterSnapshot,
  XlsxAutoFilterState,
  XlsxAutoFilterValueOption,
  XlsxCellRange,
  XlsxTableSortDirection
} from "@arcships/xlsx-core";
import { applyWorkbookAutoFilterOverrides } from "@arcships/xlsx-core";
import type { XlsxControllerContext } from "./internal";
import { resolveCellFallbackStyle } from "./internal";
import { sortWorksheetRange } from "./sorting";

/**
 * Upper bounds for a header menu, not for the filter itself.
 *
 * A value list with hundreds of thousands of entries is unusable, and a sheet
 * that large has bigger rendering costs than this scan. Filtering itself reads
 * at most this many data rows; the remainder stays visible.
 */
const MAX_AUTO_FILTER_VALUE_OPTIONS = 1000;
const MAX_AUTO_FILTER_SCAN_ROWS = 200_000;

/** Raw display value of one AutoFilter cell. */
type CellDisplayValueReader = (cell: { row: number; col: number }) => string;

export interface XlsxAutoFilterDomain {
  autoFilterRanges: ComputedRef<XlsxAutoFilterRange[]>;
  autoFilterStates: ComputedRef<XlsxAutoFilterState[]>;
  hiddenRows: ComputedRef<number[]>;
  applyAutoFilterValues: (
    reference: string,
    columnIndex: number,
    values: string[] | null,
    options?: { includeBlank?: boolean }
  ) => boolean;
  clearAutoFilter: (reference?: string) => boolean;
  getAutoFilterColumnValues: (range: XlsxAutoFilterRange, columnIndex: number) => XlsxAutoFilterValueOption[];
  sortAutoFilter: (reference: string, columnIndex: number, direction: XlsxTableSortDirection) => boolean;
  /** Writes applied/cleared criteria into saved workbook bytes. */
  applyToSavedBytes: (bytes: Uint8Array) => Uint8Array;
}

function resolveDataStartRow(range: XlsxAutoFilterRange) {
  return range.start.row + Math.max(range.headerRowCount, 1);
}

function resolveDataRowCount(range: XlsxAutoFilterRange) {
  return Math.max(0, range.dataEndRow - resolveDataStartRow(range) + 1);
}

function seededColumns(range: XlsxAutoFilterRange): XlsxAutoFilterColumnState[] {
  return range.columns
    .filter((column) => column.values.length > 0 || column.includeBlank)
    .map((column) => ({
      columnIndex: column.columnIndex,
      includeBlank: column.includeBlank,
      values: [...column.values]
    }));
}

function rangesIntersect(left: XlsxCellRange, right: XlsxCellRange) {
  return left.start.row <= right.end.row
    && left.end.row >= right.start.row
    && left.start.col <= right.end.col
    && left.end.col >= right.start.col;
}

/**
 * View-level AutoFilter state.
 *
 * Parsed definitions live on `XlsxSheetData.autoFilterRanges`; this domain
 * keeps only the user's current value filters. Hiding rows is a projection:
 * cells, merges, frozen panes and the exported package are untouched.
 */
export function createAutoFilterDomain(
  ctx: XlsxControllerContext,
  readCellDisplayValue: CellDisplayValueReader
): XlsxAutoFilterDomain {
  // Workbook sheet index → range reference → column criteria. A present key is
  // an explicit user override (an empty list clears the file's own criteria);
  // an absent key falls back to the criteria parsed from the sheet.
  const overridesByWorkbookSheetIndex = ref<Record<number, Record<string, XlsxAutoFilterColumnState[]>>>({});

  // A different source is a different workbook: criteria from the previous
  // document must not filter the new one. Undo/redo keeps `sourceBufferRef`.
  watch(ctx.sourceBufferRef, () => {
    overridesByWorkbookSheetIndex.value = {};
    workerSnapshots.value = {};
    requestedWorkerSnapshots.clear();
  });

  // Worker-backed (read-only) viewers have no main-thread workbook. The Worker
  // returns the formatted values of a range once; both the menu and the hidden
  // row projection read that snapshot.
  const workerSnapshots = ref<Record<string, XlsxAutoFilterSnapshot>>({});
  const workerSnapshotRevision = ref(0);
  const requestedWorkerSnapshots = new Set<string>();


  watch(ctx.sourceBufferRef, () => {
    overridesByWorkbookSheetIndex.value = {};
    workerSnapshots.value = {};
    requestedWorkerSnapshots.clear();
  });

  function workerSnapshotKey(range: XlsxAutoFilterRange) {
    const sheetIndex = ctx.activeSheet.value?.workbookSheetIndex ?? -1;
    return `${sheetIndex}:${range.reference}`;
  }

  function ensureWorkerSnapshot(range: XlsxAutoFilterRange) {
    if (!ctx.isWorkerBacked.value) return;
    const sheetIndex = ctx.activeSheet.value?.workbookSheetIndex;
    if (sheetIndex === undefined) return;
    const key = workerSnapshotKey(range);
    if (key in workerSnapshots.value || requestedWorkerSnapshots.has(key)) return;
    requestedWorkerSnapshots.add(key);
    const dataStartRow = resolveDataStartRow(range);
    void ctx.getWorkerClient()
      .getAutoFilterRows(
        sheetIndex,
        dataStartRow,
        range.dataEndRow,
        range.start.col,
        range.end.col,
        MAX_AUTO_FILTER_SCAN_ROWS,
      )
      .then((snapshot) => {
        if (!snapshot) return;
        workerSnapshots.value = { ...workerSnapshots.value, [key]: snapshot };
        workerSnapshotRevision.value += 1;
      })
      .catch(() => {
        // A transient Worker failure leaves the range unfiltered; the next
        // sheet switch or reload retries it.
        requestedWorkerSnapshots.delete(key);
      });
  }

  function readRangeCell(range: XlsxAutoFilterRange, row: number, col: number) {
    if (!ctx.isWorkerBacked.value) {
      return readCellDisplayValue({ row, col });
    }
    const snapshot = workerSnapshots.value[workerSnapshotKey(range)];
    if (!snapshot) return "";
    const rowData = snapshot.rows[row - resolveDataStartRow(range)];
    return rowData?.values[col - range.start.col] ?? "";
  }

  const autoFilterRanges = computed(() => ctx.activeSheet.value?.autoFilterRanges ?? []);

  // Worker-backed ranges need their display values before the seeded criteria
  // can hide rows; request them as soon as a sheet with a filter is active.
  watch(autoFilterRanges, (ranges) => {
    for (const range of ranges) {
      ensureWorkerSnapshot(range);
    }
  }, { immediate: true });

  function resolveColumns(range: XlsxAutoFilterRange): XlsxAutoFilterColumnState[] {
    const sheetIndex = ctx.activeSheet.value?.workbookSheetIndex;
    if (sheetIndex === undefined) {
      return seededColumns(range);
    }
    const sheetOverrides = overridesByWorkbookSheetIndex.value[sheetIndex];
    return Object.prototype.hasOwnProperty.call(sheetOverrides ?? {}, range.reference)
      ? sheetOverrides?.[range.reference] ?? []
      : seededColumns(range);
  }

  function activeColumns(range: XlsxAutoFilterRange) {
    return resolveColumns(range).filter(
      (column) => column.values.length > 0 || column.includeBlank
    );
  }

  function rowMatches(row: number, range: XlsxAutoFilterRange, columns: XlsxAutoFilterColumnState[]) {
    for (const column of columns) {
      const value = readRangeCell(range, row, range.start.col + column.columnIndex);
      const isBlank = value.length === 0;
      const matches = isBlank
        ? column.includeBlank
        : column.values.includes(value);
      if (!matches) {
        return false;
      }
    }
    return true;
  }

  function computeHiddenRows(range: XlsxAutoFilterRange, columns: XlsxAutoFilterColumnState[]) {
    if (columns.length === 0) {
      return [];
    }
    if (ctx.isWorkerBacked.value && !workerSnapshots.value[workerSnapshotKey(range)]) {
      // No values yet: keep every row until the Worker answers instead of
      // treating each cell as blank.
      return [];
    }
    const dataStartRow = resolveDataStartRow(range);
    const rowCount = Math.min(resolveDataRowCount(range), MAX_AUTO_FILTER_SCAN_ROWS);
    const hidden: number[] = [];
    for (let offset = 0; offset < rowCount; offset += 1) {
      const row = dataStartRow + offset;
      if (!rowMatches(row, range, columns)) {
        hidden.push(row);
      }
    }
    return hidden;
  }

  const autoFilterStates = computed<XlsxAutoFilterState[]>(() => {
    // Cell values are read through the worksheet, which is not reactive. The
    // revision makes the computed refresh after an edit, a recalc or a reload;
    // the Worker snapshot revision does the same once its values arrive.
    void ctx.revision.value;
    void ctx.historyRevision.value;
    void workerSnapshotRevision.value;
    return autoFilterRanges.value.flatMap((range) => {
      const columns = activeColumns(range);
      if (columns.length === 0) {
        return [];
      }
      return [{
        columns,
        hiddenRows: computeHiddenRows(range, columns),
        reference: range.reference
      }];
    });
  });

  const hiddenRows = computed(() => {
    const hidden = new Set<number>();
    for (const state of autoFilterStates.value) {
      for (const row of state.hiddenRows) {
        hidden.add(row);
      }
    }
    return [...hidden].sort((left, right) => left - right);
  });

  function findRange(reference: string) {
    return autoFilterRanges.value.find((range) => range.reference === reference) ?? null;
  }

  function writeColumns(
    range: XlsxAutoFilterRange,
    columns: XlsxAutoFilterColumnState[]
  ) {
    const sheetIndex = ctx.activeSheet.value?.workbookSheetIndex;
    if (sheetIndex === undefined) {
      return false;
    }
    overridesByWorkbookSheetIndex.value = {
      ...overridesByWorkbookSheetIndex.value,
      [sheetIndex]: {
        ...(overridesByWorkbookSheetIndex.value[sheetIndex] ?? {}),
        [range.reference]: columns
      }
    };
    return true;
  }

  function getAutoFilterColumnValues(
    range: XlsxAutoFilterRange,
    columnIndex: number
  ): XlsxAutoFilterValueOption[] {
    const columnCount = range.end.col - range.start.col + 1;
    if (!Number.isInteger(columnIndex) || columnIndex < 0 || columnIndex >= columnCount) {
      return [];
    }
    if (ctx.isWorkerBacked.value && !workerSnapshots.value[workerSnapshotKey(range)]) {
      return [];
    }
    const dataStartRow = resolveDataStartRow(range);
    const rowCount = Math.min(resolveDataRowCount(range), MAX_AUTO_FILTER_SCAN_ROWS);
    const col = range.start.col + columnIndex;
    const values = new Map<string, XlsxAutoFilterValueOption>();
    for (let offset = 0; offset < rowCount; offset += 1) {
      const value = readRangeCell(range, dataStartRow + offset, col);
      const key = value.length === 0 ? "\u0000blank" : value;
      if (!values.has(key)) {
        if (values.size >= MAX_AUTO_FILTER_VALUE_OPTIONS) {
          break;
        }
        values.set(key, {
          blank: value.length === 0,
          label: value,
          value
        });
      }
    }
    return [...values.values()].sort((left, right) => {
      if (left.blank !== right.blank) {
        return left.blank ? 1 : -1;
      }
      return left.value.localeCompare(right.value, undefined, {
        numeric: true,
        sensitivity: "base"
      });
    });
  }

  function applyAutoFilterValues(
    reference: string,
    columnIndex: number,
    values: string[] | null,
    options?: { includeBlank?: boolean }
  ) {
    const range = findRange(reference);
    if (!range) {
      return false;
    }
    const columnCount = range.end.col - range.start.col + 1;
    if (!Number.isInteger(columnIndex) || columnIndex < 0 || columnIndex >= columnCount) {
      return false;
    }

    const includeBlank = options?.includeBlank === true;
    const nextValues = values
      ? [...new Set(values.filter((value) => typeof value === "string" && value.length > 0))]
      : [];
    const nextColumns = resolveColumns(range).filter(
      (column) => column.columnIndex !== columnIndex
    );
    if (nextValues.length > 0 || includeBlank) {
      nextColumns.push({ columnIndex, includeBlank, values: nextValues });
    }
    nextColumns.sort((left, right) => left.columnIndex - right.columnIndex);
    return writeColumns(range, nextColumns);
  }

  function clearAutoFilter(reference?: string) {
    const ranges = reference
      ? autoFilterRanges.value.filter((range) => range.reference === reference)
      : autoFilterRanges.value;
    if (ranges.length === 0) {
      return false;
    }
    const sheetIndex = ctx.activeSheet.value?.workbookSheetIndex;
    if (sheetIndex === undefined) {
      return false;
    }
    const sheetOverrides = { ...(overridesByWorkbookSheetIndex.value[sheetIndex] ?? {}) };
    for (const range of ranges) {
      // An explicit empty list, not a deleted key: the file's own criteria must
      // not come back on the next repaint.
      sheetOverrides[range.reference] = [];
    }
    overridesByWorkbookSheetIndex.value = {
      ...overridesByWorkbookSheetIndex.value,
      [sheetIndex]: sheetOverrides
    };
    return true;
  }

  function sortAutoFilter(
    reference: string,
    columnIndex: number,
    direction: XlsxTableSortDirection
  ) {
    if (ctx.readOnly.value || ctx.isEditingLocked.value) {
      return false;
    }
    const range = findRange(reference);
    const worksheet = ctx.getActiveWorksheet();
    const activeSheet = ctx.activeSheet.value;
    const workbook = ctx.workbook.value;
    if (!range || !worksheet || !activeSheet || !workbook) {
      return false;
    }
    const columnCount = range.end.col - range.start.col + 1;
    if (!Number.isInteger(columnIndex) || columnIndex < 0 || columnIndex >= columnCount) {
      return false;
    }
    if (direction !== "ascending" && direction !== "descending") {
      return false;
    }
    const dataStartRow = resolveDataStartRow(range);
    if (dataStartRow > range.dataEndRow) {
      return false;
    }

    // Sorting moves whole rows; a merged region would stay behind and end up
    // covering unrelated values. Excel refuses this sort too.
    const sortBody: XlsxCellRange = {
      start: { row: dataStartRow, col: range.start.col },
      end: { row: range.dataEndRow, col: range.end.col }
    };
    if ((activeSheet.mergedRegions ?? []).some((region) => rangesIntersect(region, sortBody))) {
      return false;
    }

    const mutations = sortWorksheetRange(
      worksheet,
      {
        startRow: dataStartRow,
        endRow: range.dataEndRow,
        startCol: range.start.col,
        endCol: range.end.col,
        sortCol: range.start.col + columnIndex
      },
      direction,
      resolveCellFallbackStyle(activeSheet)
    );
    if (mutations === null) {
      // Already in that order: keep the menu's direction indicator in sync
      // without writing cells or a history entry.
      ctx.sortState.value = {
        columnIndex,
        direction,
        rangeReference: range.reference,
        tableName: range.tableName ?? range.reference
      };
      return true;
    }

    if (ctx.maybeRecalculateWorkbook(workbook)) {
      ctx.refreshWorkbookState(workbook);
    }
    ctx.sortState.value = {
      columnIndex,
      direction,
      rangeReference: range.reference,
      tableName: range.tableName ?? range.reference
    };
    ctx.recordRangeEditHistory(mutations, ctx.selection.value, ctx.activeCell.value);
    return true;
  }

  /**
   * Writes the current criteria into the saved package.
   *
   * The filter itself never mutates the workbook; this export-only patch is
   * what makes an applied or cleared filter survive a reload (Excel stores the
   * criteria in `<autoFilter>`). A malformed package falls back to the original
   * bytes instead of failing the export.
   */
  function applyToSavedBytes(bytes: Uint8Array) {
    const overrides = overridesByWorkbookSheetIndex.value;
    if (Object.keys(overrides).length === 0) {
      return bytes;
    }
    try {
      return applyWorkbookAutoFilterOverrides(bytes, overrides);
    } catch {
      return bytes;
    }
  }

  return {
    applyAutoFilterValues,
    applyToSavedBytes,
    autoFilterRanges,
    autoFilterStates,
    clearAutoFilter,
    getAutoFilterColumnValues,
    hiddenRows,
    sortAutoFilter
  };
}
