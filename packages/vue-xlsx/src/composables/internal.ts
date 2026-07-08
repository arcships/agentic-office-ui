import type { ComputedRef, Ref, ShallowRef } from "vue";
import type { Workbook } from "@dukelib/sheets-wasm";
import type {
  XlsxCellAddress,
  XlsxCellRange,
  XlsxChart,
  XlsxChartElementSelection,
  XlsxChartsheet,
  XlsxFormControl,
  XlsxImage,
  XlsxShape,
  XlsxSheetData,
  XlsxTable,
  XlsxTableSortState,
  XlsxWorkbookTab,
  XlsxSheetVisibility,
  UseXlsxViewerControllerOptions
} from "@extend-ai/xlsx-core";
import type {
  XlsxWorkerClient,
  WorkbookChartAssets,
  WorkbookImageAssets,
  WorkbookImageSheetOrigin
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

/**
 * Shared mutable state and helpers across all controller domain modules.
 * Defined once inside `useXlsxViewerController` and passed to each domain
 * factory so the closure body can be split into ≤1000-line files.
 */
export interface XlsxControllerContext {
  // Options (destructured from UseXlsxViewerControllerOptions)
  file: ArrayBuffer | undefined;
  src: string | undefined;
  fileName: string | undefined;
  maxFileSizeBytes: number;
  deferLoadingAboveBytes: number;
  skipXmlParsing: boolean;
  showHiddenSheets: boolean;
  workerSupported: boolean;
  shouldDeferLoading: boolean;
  canUseWorkerForRequestedReadOnly: boolean;
  requestedReadOnly: boolean;

  // Core reactive state
  isLoading: Ref<boolean>;
  error: Ref<Error | null>;
  workbook: Ref<Workbook | null>;
  sheets: Ref<XlsxSheetData[]>;
  chartsByWorkbookSheetIndex: Ref<XlsxChart[][]>;
  chartsheets: Ref<XlsxChartsheet[]>;
  tabs: Ref<XlsxWorkbookTab[]>;
  isChartsLoading: Ref<boolean>;
  workerTablesByWorkbookSheetIndex: Ref<XlsxTable[][]>;
  formControlsByWorkbookSheetIndex: Ref<XlsxFormControl[][]>;
  imagesByWorkbookSheetIndex: Ref<XlsxImage[][]>;
  shapesByWorkbookSheetIndex: Ref<XlsxShape[][]>;
  activeSheetIndex: Ref<number>;
  activeTabIndex: Ref<number>;
  zoomScaleOverridesByTabId: Ref<Record<string, number>>;
  activeCell: Ref<XlsxCellAddress | null>;
  selection: Ref<XlsxCellRange | null>;
  selectedChartId: Ref<string | null>;
  selectedChartElement: Ref<XlsxChartElementSelection | null>;
  selectedImageId: Ref<string | null>;
  revision: Ref<number>;
  selectionAnchorRef: ShallowRef<XlsxCellAddress | null>;
  undoStackRef: ShallowRef<HistoryEntry[]>;
  redoStackRef: ShallowRef<HistoryEntry[]>;
  isApplyingHistoryRef: ShallowRef<boolean>;
  historyRevision: Ref<number>;
  shouldAutoCalculate: Ref<boolean>;
  workerCellSnapshotRevision: Ref<number>;
  isWorkerBacked: Ref<boolean>;
  sortState: Ref<XlsxTableSortState | null>;
  forcedReadOnly: Ref<boolean>;
  deferredBufferRef: ShallowRef<ArrayBuffer | null>;
  deferredLoadFileSize: Ref<number | null>;
  imageAssetsRef: ShallowRef<WorkbookImageAssets | null>;
  chartAssetsRef: ShallowRef<WorkbookChartAssets | null>;
  chartLoadRequestTokenRef: ShallowRef<number>;
  chartDisplayFallbackCleanupRef: ShallowRef<(() => void) | null>;
  sheetOriginsRef: ShallowRef<Array<WorkbookImageSheetOrigin | null>>;
  workerClientRef: ShallowRef<XlsxWorkerClient | null>;
  workerCellSnapshotCacheRef: ShallowRef<Map<string, { displayValue: string; formula: string }>>;

  // Derived computeds defined in the orchestrator
  activeTab: ComputedRef<XlsxWorkbookTab | null>;
  activeSheet: ComputedRef<XlsxSheetData | null>;
  deferredMetadataCell: ComputedRef<XlsxCellAddress | null>;
  deferredMetadataSheet: ComputedRef<XlsxSheetData | null>;
  activeZoomTabKey: ComputedRef<string>;
  defaultZoomScale: ComputedRef<number>;
  zoomScale: ComputedRef<number>;
  canZoomIn: ComputedRef<boolean>;
  canZoomOut: ComputedRef<boolean>;
  readOnly: ComputedRef<boolean>;
  canResizeReadOnly: ComputedRef<boolean>;
  displayFileName: ComputedRef<string>;
  visibleSheetIndexByWorkbookSheetIndex: ComputedRef<Map<number, number>>;

  // Shared mutation helpers (defined in editing/history domains, referenced widely)
  refreshWorkbookState: (targetWorkbook: Workbook) => void;
  maybeRecalculateWorkbook: (targetWorkbook: Workbook) => void;
  getActiveWorksheet: () => ReturnType<Workbook["getSheet"]> | null;
  setChartAssets: (assets: WorkbookChartAssets | null) => void;
  setImageAssets: (assets: WorkbookImageAssets | null) => void;
  clearImageAssets: () => void;
  clearChartAssets: () => void;
  getWorkerClient: () => XlsxWorkerClient;
  disposeWorkerClient: () => void;
  startChartDisplayHydration: (buffer: ArrayBuffer, targetWorkbook: Workbook, targetSheets: XlsxSheetData[]) => void;
  loadWorkbookOnMainThread: (buffer: ArrayBuffer) => Promise<{ imageAssets: WorkbookImageAssets; parsedWorkbook: { shouldAutoCalculate: boolean; workbook: Workbook } }>;
  hasIncompleteWorkerChartSnapshot: (snapshot: { chartsByWorkbookSheetIndex: XlsxChart[][] }) => boolean;
  shouldFallbackFromWorkerError: (error: unknown) => boolean;
  ensureChartAssetsHydrated: (targetWorkbook: Workbook | null, targetSheets: XlsxSheetData[]) => WorkbookChartAssets | null;
  shouldForceReadOnlyForBuffer: (bufferByteLength: number) => boolean;
  shouldUseWorkerForReadOnlyLoad: (willForceReadOnly: boolean) => boolean;
  createSavedWorkbookBytes: (targetWorkbook: Workbook) => Uint8Array;
  captureCellMutationState: (cell: XlsxCellAddress) => CellMutationState | null;
  recordHistoryBeforeMutation: () => void;
  recordCellEditHistory: (cell: XlsxCellAddress, before: CellMutationState, after: CellMutationState) => void;
  recordRangeEditHistory: (mutations: RangeCellMutation[], selectionAfter: XlsxCellRange | null, activeCellAfter: XlsxCellAddress | null) => void;
  setChartSeriesFormula: (chartId: string, seriesIndex: number, formula: string) => boolean;
  selectedFormulaTarget: ComputedRef<
    | { kind: "chartSeries"; chartId: string; seriesId: string; seriesIndex: number }
    | { kind: "cell"; cell: XlsxCellAddress | null }
  >;
}

export type { UseXlsxViewerControllerOptions };
