import type { ComputedRef, Ref, ShallowRef } from "vue";
import type { Workbook } from "@dukelib/sheets-wasm";
import type {
  XlsxCellAddress,
  XlsxCellEditRequest,
  XlsxCellNumberFormatInput,
  XlsxCellRange,
  XlsxChart,
  XlsxChartElementSelection,
  XlsxChartsheet,
  XlsxDataValidation,
  XlsxDataValidationFeedback,
  XlsxFormControl,
  XlsxImage,
  XlsxNamedRange,
  XlsxShape,
  XlsxSheetData,
  XlsxTable,
  XlsxTableSortState,
  XlsxWorkbookTab,
  XlsxSheetVisibility,
  XlsxLoadError,
  XlsxSourceState,
  UseXlsxViewerControllerOptions,
  SafeCalculateSkipReason
} from "@arcships/xlsx-core";
import type {
  XlsxWorkerClient,
  WorkbookChartAssets,
  WorkbookImageAssets,
  WorkbookImageSheetOrigin
} from "@arcships/xlsx-core";
import { cellAddressToA1 } from "./selection";
import { DEFAULT_CELL_STYLE_PATCH, coerceEnteredValue } from "./cell-input";
import type { CellInputOptions } from "./cell-input";
import { toEngineCellStyle } from "./engine-style";
import type { XlsxDataValidationCheck } from "./validation";

export const FORMULA_COUNT_THRESHOLD = 1000;

export type XlsxCalculationSkipReason =
  | "unresolved-sheet-refs"
  | "calculate-trapped"
  | "recovery-snapshot-unavailable";

/** User-visible reason why the viewer is showing cached formula values. */
export type XlsxCalculationStatus = {
  reason: XlsxCalculationSkipReason;
  message: string;
};

export type XlsxCalculationOutcome = {
  calculated: boolean;
  /** False when the workbook instance can no longer be read safely. */
  safeToRefresh: boolean;
};
export const DEFAULT_ROW_HEIGHT = 24;
export const DEFAULT_COL_WIDTH = 80;
export const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
export const CSV_MIME_TYPE = "text/csv;charset=utf-8";
export const MIN_COL_WIDTH_PX = 30;
export const MIN_ROW_HEIGHT_PX = 16;
// Must match the header strips painted by XlsxGrid; components/grid-metrics.ts
// re-exports these for the view layer, so this is the single source of truth.
export const GRID_HEADER_HEIGHT = 24;
export const GRID_ROW_HEADER_WIDTH = 48;
export const HISTORY_LIMIT = 100;
export const DEFAULT_HISTORY_MAX_BYTES = 32 * 1024 * 1024;
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
  /** Active tab, so undoing a chartsheet move restores that chartsheet. */
  activeTabIndex?: number;
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

export type SelectionSnapshot = {
  activeCell: XlsxCellAddress | null;
  selection: XlsxCellRange | null;
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

export type HistoryBudget = Readonly<{
  maxEntries: number;
  maxBytes: number;
}>;

export type ClipboardMatrixCell = {
  colOffset: number;
  formula: string | null;
  rowOffset: number;
  /** Raw typed value (number/boolean/text); keeps the value type on internal paste. */
  rawValue?: unknown;
  /**
   * Effective cell style captured with the cell. `null` means "no explicit
   * style", so an internal paste clears the target formatting like Excel does.
   */
  style?: unknown;
  /** Display string used for the plain-text and HTML flavours. */
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
  /**
   * Top-left cell the payload was copied from. Paste uses it to shift relative
   * formula references by the copy→paste distance. Optional so payloads from
   * older builds and other hosts still paste (verbatim).
   */
  origin?: XlsxCellAddress;
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

export function createHistoryBudget(
  requestedMaxEntries = HISTORY_LIMIT,
  requestedMaxBytes = DEFAULT_HISTORY_MAX_BYTES
): HistoryBudget {
  return Object.freeze({
    maxEntries:
      Number.isFinite(requestedMaxEntries) && requestedMaxEntries > 0
        ? Math.max(1, Math.floor(requestedMaxEntries))
        : HISTORY_LIMIT,
    maxBytes:
      Number.isFinite(requestedMaxBytes) && requestedMaxBytes > 0
        ? Math.max(1, Math.floor(requestedMaxBytes))
        : DEFAULT_HISTORY_MAX_BYTES
  });
}

export function estimateHistoryEntryBytes(entry: HistoryEntry): number {
  let binaryBytes = 0;
  const seen = new WeakSet<object>();
  try {
    const json = JSON.stringify(entry, (_key, candidate: unknown) => {
      if (!candidate || typeof candidate !== "object") return candidate;
      if (candidate instanceof ArrayBuffer) {
        binaryBytes += candidate.byteLength;
        return null;
      }
      if (ArrayBuffer.isView(candidate)) {
        binaryBytes += candidate.byteLength;
        return null;
      }
      if (seen.has(candidate)) return null;
      seen.add(candidate);
      return candidate;
    }) ?? "";
    return Math.min(
      Number.MAX_SAFE_INTEGER,
      new TextEncoder().encode(json).byteLength + binaryBytes
    );
  } catch {
    return Number.MAX_SAFE_INTEGER;
  }
}

export function pushHistoryEntry(
  stack: HistoryEntry[],
  entry: HistoryEntry,
  budget: HistoryBudget = createHistoryBudget()
) {
  stack.push(entry);

  let retainedEntries = 0;
  let retainedBytes = 0;
  let firstRetainedIndex = stack.length - 1;
  for (let index = stack.length - 1; index >= 0; index -= 1) {
    const entryBytes = estimateHistoryEntryBytes(stack[index]);
    if (
      retainedEntries > 0 &&
      (retainedEntries >= budget.maxEntries || retainedBytes + entryBytes > budget.maxBytes)
    ) {
      break;
    }
    retainedEntries += 1;
    retainedBytes = Math.min(Number.MAX_SAFE_INTEGER, retainedBytes + entryBytes);
    firstRetainedIndex = index;
  }

  if (firstRetainedIndex > 0) {
    stack.splice(0, firstRetainedIndex);
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

export function coerceUserEnteredValue(value: string, options: CellInputOptions = {}): unknown {
  return coerceEnteredValue(value, options);
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
  state: CellMutationState,
  fallbackStyle?: unknown
) {
  if (state.formula) {
    worksheet.setFormula(cellAddressToA1(cell), state.formula);
  } else {
    worksheet.setCell(cellAddressToA1(cell), normalizeCellValue(state.value));
  }

  if (state.style && typeof state.style === "object") {
    worksheet.setCellStyleAt(cell.row, cell.col, state.style);
    return;
  }

  if (state.style !== null) {
    return;
  }

  // `null` means the cell had no explicit style before the mutation. The wasm
  // workbook has no "unset style" entry point, so undo has to restore the
  // effective formatting explicitly: the inherited row/column style (or the
  // workbook default) when the caller knows it, Excel's defaults otherwise.
  // A cell that matches the workbook default normalizes back to "no style".
  // Only touch cells that have a style now, so plain value edits never add a
  // style to a cell that never had one.
  if (!worksheet.getCellStyleAt(cell.row, cell.col)) {
    return;
  }

  const restoreStyle = fallbackStyle && typeof fallbackStyle === "object"
    ? fallbackStyle
    : DEFAULT_CELL_STYLE_PATCH;
  worksheet.setCellStyleAt(cell.row, cell.col, restoreStyle);
}

/**
 * Style patch that restores the workbook's default formatting.
 *
 * Undo clears a cell that had no explicit style back to the workbook default
 * rather than to a hard-coded Calibri style, so a workbook with its own default
 * font (or a default number format) keeps it. The number format is always
 * explicit, because that is the group undo is usually restoring.
 *
 * The parsed model mirrors the source OOXML, so its colours are translated to
 * the engine's colour input before the patch reaches `setCellStyleAt`. The
 * result always names every style group and every property inside a group: the
 * engine merges groups and their properties, so a patch that only mentioned the
 * default font would leave the bold or the border undo is removing in place.
 */
export function resolveCellFallbackStyle(
  sheet: { styleById?: Record<number, unknown> } | null | undefined
): unknown {
  const defaultStyle = sheet?.styleById?.[0];
  if (!defaultStyle || typeof defaultStyle !== "object") {
    return null;
  }
  const converted = toEngineCellStyle(defaultStyle);
  const style: Record<string, unknown> = { ...DEFAULT_CELL_STYLE_PATCH };
  if (converted && typeof converted === "object") {
    for (const [group, value] of Object.entries(converted as Record<string, unknown>)) {
      const base = style[group];
      style[group] = base && typeof base === "object" && value && typeof value === "object"
        ? { ...base, ...value }
        : value;
    }
  }
  return style;
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
  sourceState: Ref<XlsxSourceState>;
  sourceError: Ref<XlsxLoadError | null>;
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
  selectedShapeId: Ref<string | null>;
  revision: Ref<number>;
  selectionAnchorRef: ShallowRef<XlsxCellAddress | null>;
  undoStackRef: ShallowRef<HistoryEntry[]>;
  redoStackRef: ShallowRef<HistoryEntry[]>;
  historyBudget: HistoryBudget;
  isApplyingHistoryRef: ShallowRef<boolean>;
  historyRevision: Ref<number>;
  /** Locale used by number/date formats that name none of their own. */
  locale: Ref<string>;
  shouldAutoCalculate: Ref<boolean>;
  workerCellSnapshotRevision: Ref<number>;
  isWorkerBacked: Ref<boolean>;
  sortState: Ref<XlsxTableSortState | null>;
  forcedReadOnly: Ref<boolean>;
  /** Inline-edit request watched by the grid; `token` retriggers the same cell. */
  editRequestRef: Ref<XlsxCellEditRequest | null>;
  /** Name Manager edits layered over the engine's defined names. */
  definedNameOverridesRef: Ref<Record<string, XlsxNamedRange>>;
  /** Names the Name Manager deleted; the engine cannot remove a defined name. */
  removedDefinedNamesRef: Ref<string[]>;
  dataValidationFeedback: Ref<XlsxDataValidationFeedback | null>;
  lastExportWarning: Ref<string | null>;
  deferredBufferRef: ShallowRef<ArrayBuffer | null>;
  /** Load generation that owns the currently deferred buffer. */
  deferredLoadRequestIdRef: ShallowRef<number | null>;
  /** Exact bytes accepted for the current source; downloads must use these bytes. */
  sourceBufferRef: ShallowRef<ArrayBuffer | null>;
  deferredLoadFileSize: Ref<number | null>;
  imageAssetsRef: ShallowRef<WorkbookImageAssets | null>;
  chartAssetsRef: ShallowRef<WorkbookChartAssets | null>;
  chartLoadRequestTokenRef: ShallowRef<number>;
  chartDisplayFallbackCleanupRef: ShallowRef<(() => void) | null>;
  sheetOriginsRef: ShallowRef<Array<WorkbookImageSheetOrigin | null>>;
  workerClientRef: ShallowRef<XlsxWorkerClient | null>;
  workerCellSnapshotCacheRef: ShallowRef<Map<string, { displayValue: string; formula: string; numberFormat?: XlsxCellNumberFormatInput | null }>>;

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
  isActiveSheetProtected: ComputedRef<boolean>;
  isEditingLocked: ComputedRef<boolean>;
  canResizeReadOnly: ComputedRef<boolean>;
  displayFileName: ComputedRef<string>;
  visibleSheetIndexByWorkbookSheetIndex: ComputedRef<Map<number, number>>;

  // Shared mutation helpers (defined in editing/history domains, referenced widely)
  /**
   * Activates a visible sheet together with the tab that renders it, so the grid,
   * the tab strip and the sheet-scoped overlays cannot point at different
   * sheets. Chartsheet tabs are never selected. `preserveSelection` keeps the
   * current selection across the switch instead of letting the tab watcher reset
   * it; only history restores use it, because they restore selection themselves.
   */
  switchToSheet: (visibleSheetIndex: number, options?: { preserveSelection?: boolean }) => void;
  /**
   * Activates one tab, worksheet or chartsheet, keeping the grid and the tab
   * strip in agreement. `preserveSelection` keeps the current selection across
   * a switch that does not change the sheet the user is working on.
   */
  activateTab: (tabIndex: number, options?: { preserveSelection?: boolean }) => void;
  /**
   * Restores a history snapshot's sheet and tab together without letting the
   * tab watcher reset the selection the snapshot restored.
   */
  restoreActiveSheetAndTab: (visibleSheetIndex: number, tabIndex?: number) => void;
  /**
   * Consumes the marker set by a `preserveSelection` switch. True means the
   * pending tab-index change must not reset sheet-scoped state.
   */
  consumePreservedTabSwitch: (nextTabIndex: number) => boolean;
  refreshWorkbookState: (targetWorkbook: Workbook) => void;
  /** Recalculates when enabled; true when workbook state is safe to refresh. */
  maybeRecalculateWorkbook: (targetWorkbook: Workbook) => boolean;
  /** Recalculates through safeCalculate, snapshotting before a possible trap. */
  recalculateWorkbookSafely: (targetWorkbook: Workbook) => XlsxCalculationOutcome;
  /** Restores automatic recalculation after a successful manual retry. */
  resumeAutoCalculate: () => void;
  calculationStatus: Ref<XlsxCalculationStatus | null>;
  noteCalculationResult: (result: {
    calculated: boolean;
    skipReason: XlsxCalculationSkipReason | null;
  }) => void;
  getActiveWorksheet: () => ReturnType<Workbook["getSheet"]> | null;
  setChartAssets: (assets: WorkbookChartAssets | null) => void;
  setImageAssets: (assets: WorkbookImageAssets | null) => void;
  clearImageAssets: () => void;
  clearChartAssets: () => void;
  getWorkerClient: () => XlsxWorkerClient;
  disposeWorkerClient: () => void;
  replaceWorkbook: (workbook: Workbook | null) => void;
  disposeWorkbook: (workbook: Workbook | null | undefined) => void;
  discardWorkbookLoadResult: (assets: WorkbookImageAssets | null, workbook: Workbook | null) => void;
  isCurrentLoadRequest: (requestId: number) => boolean;
  getCurrentLoadRequestId: () => number;
  /** Resumes the one instance-owned deferred load; repeated calls are ignored. */
  continueDeferredWorkbookLoad: () => void;
  /** Abandons the one instance-owned deferred load; repeated calls are ignored. */
  cancelDeferredWorkbookLoad: () => void;
  startChartDisplayHydration: (buffer: ArrayBuffer, targetWorkbook: Workbook, targetSheets: XlsxSheetData[]) => void;
  loadWorkbookOnMainThread: (buffer: ArrayBuffer) => Promise<{ imageAssets: WorkbookImageAssets; parsedWorkbook: { shouldAutoCalculate: boolean; calculationSkipReason: SafeCalculateSkipReason | null; workbook: Workbook } }>;
  hasIncompleteWorkerChartSnapshot: (snapshot: { chartsByWorkbookSheetIndex: XlsxChart[][] }) => boolean;
  ensureChartAssetsHydrated: (targetWorkbook: Workbook | null, targetSheets: XlsxSheetData[]) => WorkbookChartAssets | null;
  shouldForceReadOnlyForBuffer: (bufferByteLength: number) => boolean;
  shouldUseWorkerForReadOnlyLoad: (willForceReadOnly: boolean) => boolean;
  createSavedWorkbookBytes: (targetWorkbook: Workbook, reportSanitizationWarning?: boolean) => Uint8Array;
  captureCellMutationState: (cell: XlsxCellAddress) => CellMutationState | null;
  checkCellValue: (cell: XlsxCellAddress, value: string) => XlsxDataValidationCheck;
  getCellDataValidation: (cell?: XlsxCellAddress | null) => XlsxDataValidation | null;
  getDataValidationListOptions: (validation: XlsxDataValidation) => string[];
  setValidationFeedback: (feedback: XlsxDataValidationFeedback | null) => void;
  dismissValidationFeedback: () => void;
  setExportWarning: (message: string | null) => void;
  recordHistoryBeforeMutation: (precomputedBytes?: Uint8Array) => void;
  recordCellEditHistory: (cell: XlsxCellAddress, before: CellMutationState, after: CellMutationState) => void;
  recordRangeEditHistory: (mutations: RangeCellMutation[], selectionAfter: XlsxCellRange | null, activeCellAfter: XlsxCellAddress | null, before?: SelectionSnapshot) => void;
  setChartSeriesFormula: (chartId: string, seriesIndex: number, formula: string) => boolean;
  selectedFormulaTarget: ComputedRef<
    | { kind: "chartSeries"; chartId: string; seriesId: string; seriesIndex: number }
    | { kind: "cell"; cell: XlsxCellAddress | null }
  >;
}

export type { UseXlsxViewerControllerOptions };
