import type { Workbook } from "@dukelib/sheets-wasm";
import type {
  XlsxCellAddress,
  XlsxCellRange,
  XlsxSheetVisibility
} from "@extend-ai/xlsx-core";
import { cellAddressToA1 } from "./selection";

export const FORMULA_COUNT_THRESHOLD = 1000;
export const DEFAULT_ROW_HEIGHT = 24;
export const DEFAULT_COL_WIDTH = 80;
export const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
export const CSV_MIME_TYPE = "text/csv;charset=utf-8";
export const MIN_COL_WIDTH_PX = 30;
export const MIN_ROW_HEIGHT_PX = 16;
export const GRID_HEADER_HEIGHT = 24;
export const GRID_ROW_HEADER_WIDTH = 40;
export const HISTORY_LIMIT = 100;
export const INTERNAL_CLIPBOARD_MIME = "application/x-react-xlsx-range+json";
export const DEFAULT_DEFER_LOADING_ABOVE_BYTES = 0;
export const DEFAULT_MAX_FILE_SIZE_BYTES = 25 * 1024 * 1024;
export const MAX_INTERACTIVE_WORKSHEET_XML_BYTES = 200 * 1024 * 1024;
export const MAX_INTERACTIVE_SHARED_STRINGS_BYTES = 50 * 1024 * 1024;
export const MAX_INTERACTIVE_TOTAL_XML_BYTES = 256 * 1024 * 1024;
export const EMU_PER_PIXEL = 9525;
export const IMAGE_BATCH_ROW_COUNT = 256;
export const DEFAULT_ZOOM_SCALE = 100;
export const MIN_ZOOM_SCALE = 10;
export const MAX_ZOOM_SCALE = 400;
export const ZOOM_STEP = 10;
export const DEFAULT_ZOOM_TAB_KEY = "__default__";

export function normalizeWorksheetVisibility(value: unknown): XlsxSheetVisibility {
  return value === "hidden" || value === "veryHidden" ? value : "visible";
}

export type IdleRequestHandle = number;

export type IdleDeadline = {
  didTimeout: boolean;
  timeRemaining: () => number;
};

export type IdleWindow = Window & {
  cancelIdleCallback?: (handle: IdleRequestHandle) => void;
  requestIdleCallback?: (
    callback: (deadline: IdleDeadline) => void,
    options?: {
      timeout: number;
    }
  ) => IdleRequestHandle;
};

export type SnapshotHistoryEntry = {
  kind: "snapshot";
  activeCell: XlsxCellAddress | null;
  activeSheetIndex: number;
  bytes: Uint8Array;
  selection: XlsxCellRange | null;
};

export type CellMutationState = {
  formula: string | null;
  style: unknown;
  value: unknown;
};

export type CellEditHistoryEntry = {
  kind: "cell-edit";
  activeCellAfter: XlsxCellAddress | null;
  activeCellBefore: XlsxCellAddress | null;
  after: CellMutationState;
  before: CellMutationState;
  cell: XlsxCellAddress;
  selectionAfter: XlsxCellRange | null;
  selectionBefore: XlsxCellRange | null;
  sheetIndex: number;
};

export type RangeCellMutation = {
  after: CellMutationState;
  before: CellMutationState;
  cell: XlsxCellAddress;
};

export type RangeEditHistoryEntry = {
  kind: "range-edit";
  activeCellAfter: XlsxCellAddress | null;
  activeCellBefore: XlsxCellAddress | null;
  mutations: RangeCellMutation[];
  selectionAfter: XlsxCellRange | null;
  selectionBefore: XlsxCellRange | null;
  sheetIndex: number;
};

export type HistoryEntry = SnapshotHistoryEntry | CellEditHistoryEntry | RangeEditHistoryEntry;

export type ClipboardMatrixCell = {
  colOffset: number;
  formula: string | null;
  rowOffset: number;
  value: string;
};

export type ClipboardMerge = {
  colSpan: number;
  colOffset: number;
  rowOffset: number;
  rowSpan: number;
};

export type ClipboardPayload = {
  cells: ClipboardMatrixCell[];
  cols: number;
  merges: ClipboardMerge[];
  rows: number;
};

export function createAbortError() {
  if (typeof DOMException !== "undefined") {
    return new DOMException("Aborted", "AbortError");
  }

  const error = new Error("Aborted");
  error.name = "AbortError";
  return error;
}

export function isAbortError(error: unknown) {
  return error instanceof Error && error.name === "AbortError";
}

export function pushHistoryEntry(stack: HistoryEntry[], entry: HistoryEntry) {
  stack.push(entry);
  if (stack.length > HISTORY_LIMIT) {
    stack.shift();
  }
}

export function normalizeCellValue(value: unknown) {
  return value ?? "";
}

export function cloneCellStyle(style: unknown): unknown {
  if (!style || typeof style !== "object") {
    return style;
  }

  if (typeof structuredClone === "function") {
    try {
      return structuredClone(style);
    } catch {
      // Fall through to the JSON clone below.
    }
  }

  try {
    return JSON.parse(JSON.stringify(style));
  } catch {
    return style;
  }
}

export function coerceUserEnteredValue(value: string): unknown {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }

  if (trimmed.startsWith("'")) {
    return trimmed.slice(1);
  }

  if (/^(true|false)$/i.test(trimmed)) {
    return trimmed.toLowerCase() === "true";
  }

  if (/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(trimmed)) {
    const parsed = Number(trimmed);
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }

  return value;
}

export function scheduleLowPriorityTask(task: () => void) {
  if (typeof window === "undefined") {
    const timeoutHandle = setTimeout(task, 0);
    return () => clearTimeout(timeoutHandle);
  }

  const idleWindow = window as IdleWindow;
  if (typeof idleWindow.requestIdleCallback === "function") {
    const idleHandle = idleWindow.requestIdleCallback(() => {
      task();
    }, { timeout: 120 });
    return () => {
      idleWindow.cancelIdleCallback?.(idleHandle);
    };
  }

  const timeoutHandle = window.setTimeout(task, 0);
  return () => window.clearTimeout(timeoutHandle);
}

export function asFiniteNumber(value: unknown) {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

export function applyCellMutationState(
  worksheet: ReturnType<Workbook["getSheet"]>,
  cell: XlsxCellAddress,
  state: CellMutationState
) {
  if (state.formula) {
    worksheet.setFormula(cellAddressToA1(cell), state.formula);
  } else {
    worksheet.setCell(cellAddressToA1(cell), normalizeCellValue(state.value));
  }

  if (state.style && typeof state.style === "object") {
    worksheet.setCellStyleAt(cell.row, cell.col, state.style);
  }
}
