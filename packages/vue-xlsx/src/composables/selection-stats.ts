import { computed, ref, shallowRef, watch, type ComputedRef, type Ref } from "vue";
import type { Worksheet } from "@dukelib/sheets-wasm";
import {
  XLSX_SELECTION_STATS_CELL_LIMIT,
  summarizeSelectionValues,
  type XlsxCellNumberFormatInput,
  type XlsxSelectionStatsQuery,
  type XlsxSelectionSummary,
  type XlsxViewerController,
} from "@arcships/xlsx-core";
import { resolveEffectiveCellNumberFormat } from "./formatting";
import { normalizeRange } from "./selection";

/**
 * One selection read, resolved from the controller's reactive state.
 *
 * `hiddenRows` only carries the AutoFilter-hidden rows the bounded scan can
 * reach; `hiddenRowCount` counts every filtered-out row in the selection so
 * `cellCount` still describes the whole range.
 */
export interface XlsxSelectionStatsRequest extends XlsxSelectionStatsQuery {
  sheetIndex: number;
  /** Controller revision that invalidates a previously read summary. */
  revision: number;
  hiddenRowCount: number;
}

export interface XlsxSelectionStatsState {
  /** Aggregate of the current selection, or `null` when nothing is selected. */
  summary: ComputedRef<XlsxSelectionSummary | null>;
  /** True while the Worker is summarizing a Worker-backed selection. */
  pending: Ref<boolean>;
  /** True when the selection contains rows hidden by the active AutoFilter. */
  filtered: ComputedRef<boolean>;
  /** The resolved read; exposed for hosts that render their own status bar. */
  request: ComputedRef<XlsxSelectionStatsRequest | null>;
}

/** Index of the first hidden row at or after `row` in an ascending list. */
function lowerBound(rows: readonly number[], row: number): number {
  let low = 0;
  let high = rows.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if ((rows[middle] ?? 0) < row) low = middle + 1;
    else high = middle;
  }
  return low;
}

/** Index of the first hidden row after `row` in an ascending list. */
function upperBound(rows: readonly number[], row: number): number {
  let low = 0;
  let high = rows.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if ((rows[middle] ?? 0) <= row) low = middle + 1;
    else high = middle;
  }
  return low;
}

function countHiddenRowsInRange(rows: readonly number[], startRow: number, endRow: number): number {
  if (rows.length === 0 || endRow < startRow) return 0;
  return upperBound(rows, endRow) - lowerBound(rows, startRow);
}

/**
 * Hidden rows the row-major scan can reach.
 *
 * The scan budget is expressed in cells, so at most
 * `ceil(maxCells / columnCount)` rows are read. Rows after that cannot affect
 * the totals and stay out of the Worker request.
 */
function collectScanHiddenRows(
  rows: readonly number[],
  startRow: number,
  endRow: number,
  rowBudget: number,
): number[] {
  if (rows.length === 0) return [];
  const collected: number[] = [];
  const lastRow = Math.min(endRow, startRow + rowBudget - 1);
  let index = lowerBound(rows, startRow);
  for (let row = startRow; row <= lastRow; row += 1) {
    // Both lists are ascending; skipping ahead keeps the walk proportional to
    // the scanned rows rather than to the whole hidden-row list.
    while (index < rows.length && (rows[index] ?? 0) < row) index += 1;
    if ((rows[index] ?? 0) !== row) continue;
    collected.push(row);
    index += 1;
  }
  return collected;
}

function readWorksheetSummary(
  controller: XlsxViewerController,
  request: XlsxSelectionStatsRequest,
): XlsxSelectionSummary | null {
  const worksheet: Worksheet | null = controller.getActiveWorksheet();
  if (!worksheet) return null;

  try {
    const hidden = new Set(request.hiddenRows ?? []);
    const values: Array<number | null> = [];
    let scannedCells = 0;
    let numberFormat: XlsxCellNumberFormatInput | null = null;

    scan: for (let row = request.startRow; row <= request.endRow; row += 1) {
      for (let col = request.startCol; col <= request.endCol; col += 1) {
        if (scannedCells >= request.maxCells) break scan;
        scannedCells += 1;
        // Filtered rows still consume the budget: the cap describes how much
        // of the selection the status bar looked at, and it stays bounded even
        // when a filter hides most of a huge range.
        if (hidden.has(row)) continue;
        const value = worksheet.getCalculatedValueAt(row, col);
        try {
          if (value.is_empty) continue;
          let number: number | null = null;
          if (value.is_number) {
            const parsed = value.asNumber();
            if (parsed !== undefined && Number.isFinite(parsed)) number = parsed;
          }
          if (number !== null && numberFormat === null) {
            numberFormat = resolveEffectiveCellNumberFormat(
              worksheet.getCellStyleAt(row, col),
              controller.activeSheet,
              row,
              col,
            );
          }
          values.push(number);
        } finally {
          value.free();
        }
      }
    }

    return summarizeSelectionValues(values, {
      numberFormat,
      scannedCells,
      totalCells: request.cellCount,
    });
  } finally {
    worksheet.free();
  }
}

/**
 * Selection summary for the status bar.
 *
 * Main-thread viewers read the active worksheet directly, bounded by
 * {@link XLSX_SELECTION_STATS_CELL_LIMIT}. Worker-backed viewers ask the
 * Worker for the aggregate and keep the previous result only until the new
 * one arrives; the read is aborted when the selection moves again.
 */
export function useXlsxSelectionStats(
  controller: () => XlsxViewerController | null | undefined,
): XlsxSelectionStatsState {
  const workerSummary = shallowRef<XlsxSelectionSummary | null>(null);
  const pending = ref(false);

  const request = computed<XlsxSelectionStatsRequest | null>(() => {
    const current = controller();
    const sheet = current?.activeSheet;
    const selection = current?.selection;
    if (!current || !sheet || !selection) return null;

    const range = normalizeRange(selection);
    const columnCount = range.end.col - range.start.col + 1;
    const rowCount = range.end.row - range.start.row + 1;
    if (columnCount <= 0 || rowCount <= 0) return null;

    const hidden = current.hiddenAutoFilterRows ?? [];
    const hiddenRowCount = countHiddenRowsInRange(hidden, range.start.row, range.end.row);
    const visibleRows = Math.max(0, rowCount - hiddenRowCount);
    const rowBudget = Math.max(1, Math.ceil(XLSX_SELECTION_STATS_CELL_LIMIT / columnCount));

    return {
      cellCount: visibleRows * columnCount,
      endCol: range.end.col,
      endRow: range.end.row,
      hiddenRowCount,
      hiddenRows: collectScanHiddenRows(hidden, range.start.row, range.end.row, rowBudget),
      maxCells: XLSX_SELECTION_STATS_CELL_LIMIT,
      revision: current.revision,
      sheetIndex: sheet.workbookSheetIndex,
      startCol: range.start.col,
      startRow: range.start.row,
    };
  });

  watch(
    request,
    (next, _previous, onCleanup) => {
      const current = controller();
      const readFromWorker = current?.getSelectionStatsAsync;
      workerSummary.value = null;
      if (!next || !current?.isWorkerBacked || typeof readFromWorker !== "function") {
        pending.value = false;
        return;
      }

      const abort = new AbortController();
      let isCurrent = true;
      pending.value = true;
      readFromWorker.call(current, next.sheetIndex, next, abort.signal)
        .then((summary) => {
          if (!isCurrent) return;
          workerSummary.value = summary;
          pending.value = false;
        })
        .catch(() => {
          // A failed read leaves the totals empty; the next selection or sheet
          // switch retries with a fresh Worker client.
          if (isCurrent) pending.value = false;
        });

      onCleanup(() => {
        isCurrent = false;
        abort.abort();
      });
    },
    { immediate: true },
  );

  const summary = computed<XlsxSelectionSummary | null>(() => {
    const next = request.value;
    const current = controller();
    if (!next || !current) return null;
    if (current.isWorkerBacked) return workerSummary.value;
    return readWorksheetSummary(current, next);
  });

  const filtered = computed(() => (request.value?.hiddenRowCount ?? 0) > 0);

  return { filtered, pending, request, summary };
}
