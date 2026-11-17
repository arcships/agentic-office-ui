import { ref, computed, watch, onUnmounted, shallowRef, nextTick } from "vue";
import {
  createLatestTaskCoordinator,
  createOfficeTaskSequence,
  sanitizeOfficeUrl,
  type OfficeLoadTask,
  type OfficeSource,
  type OfficeUrlPolicy,
} from "@arcships/office-runtime";
import type { Workbook } from "@dukelib/sheets-wasm";
import type {
  UseXlsxViewerControllerOptions,
  XlsxCellAddress,
  XlsxCellRange,
  XlsxChart,
  XlsxChartElementSelection,
  XlsxChartsheet,
  XlsxDataValidation,
  XlsxDataValidationFeedback,
  XlsxFormControl,
  XlsxImage,
  XlsxShape,
  XlsxSheetData,
  XlsxTable,
  XlsxTableSortState,
  XlsxLoadError,
  XlsxSourceState,
  XlsxViewerController,
  XlsxWorkbookTab,
  XlsxDiagnostic,
  XlsxRuntime
} from "@arcships/xlsx-core";
import {
  XlsxWorkerClient,
  XlsxWorkerError,
  createXlsxRuntime,
  validateXlsxArchive,
  loadWorkbookChartAssets,
  mergeWorkbookImageAssets,
  revokeWorkbookImageAssets,
  safeCalculate,
  trackXlsxWorkbookLifetime,
  type SafeCalculateSkipReason,
  type WorkbookChartAssets,
  type WorkbookImageAssets,
  type WorkbookImageSheetOrigin,
  type XlsxCellNumberFormatInput
} from "@arcships/xlsx-core";
import {
  DEFAULT_DEFER_LOADING_ABOVE_BYTES,
  DEFAULT_HISTORY_MAX_BYTES,
  DEFAULT_MAX_FILE_SIZE_BYTES,
  DEFAULT_ZOOM_TAB_KEY,
  MAX_ZOOM_SCALE,
  MIN_ZOOM_SCALE,
  isAbortError,
  createHistoryBudget,
  HISTORY_LIMIT,
  pushHistoryEntry,
  scheduleLowPriorityTask,
  type CellMutationState,
  type HistoryEntry,
  type RangeCellMutation,
  type SelectionSnapshot,
  type SnapshotHistoryEntry,
  type XlsxCalculationOutcome,
  type XlsxCalculationStatus
} from "./internal";
import {
  clampZoomScale,
  resolveDefaultZoomScale,
  resolveNextZoomScale,
  buildSheetList,
  buildVisibleSheetIndexMap,
  convertCsvBufferToXlsx,
  convertXlsbBufferToXlsx,
  resolveWorkbookSource,
  parseWorkbookBuffer,
  createHistoryDomain
} from "./workbook-state";
import { createEditingDomain } from "./editing";
import { createClipboardDomain } from "./clipboard";
import {
  createChartImageDomain,
  loadWorkbookImageAssets,
  shouldSkipXmlParsingForWorkbook
} from "./chart-controller";
import { createNavigationDomain } from "./navigation";
import {
  XlsxFileSizeLimitExceededError,
  createWorkbookTooLargeError,
  decodeHtmlEntities,
  fileStem,
  preflightWorkbookBuffer,
  resolveDisplayFileName,
  sanitizeSavedWorkbookBytes
} from "./formatting";
import { cellAddressToA1, parseA1RangeReference, rangeToA1 } from "./selection";
import { normalizeDocumentLocale } from "./number-locale";
import {
  checkCellValueAgainstDataValidation,
  findDataValidationForCell,
  isSheetEditingLocked,
  resolveDataValidationListOptions
} from "./validation";
import { toXlsxLoadError } from "@arcships/xlsx-core";

type DeferredWorkbookLoad = {
  buffer: ArrayBuffer;
  inputFormat: "csv" | "xlsb" | "workbook";
  csvConverted?: boolean;
  xlsbConverted?: boolean;
  /** Original source bytes retained for download and input-size diagnostics. */
  sourceBuffer: ArrayBuffer;
  requestId: number;
  resolvedUrl?: string;
  sourceKind: "file" | "url";
  state: "deferred" | "running" | "settled";
  task: OfficeLoadTask;
  archiveValidated?: boolean;
};

const WORKER_CELL_SNAPSHOT_CACHE_MAX_ENTRIES = 512;

type WorkerCellSnapshot = {
  displayValue: string;
  formula: string;
  numberFormat?: XlsxCellNumberFormatInput | null;
};

/** Bounded LRU for Worker cell snapshots; browsing a sheet must not grow forever. */
class WorkerCellSnapshotCache extends Map<string, WorkerCellSnapshot> {
  private readonly maxEntries: number;

  constructor(maxEntries = WORKER_CELL_SNAPSHOT_CACHE_MAX_ENTRIES) {
    super();
    this.maxEntries = maxEntries;
  }

  override get(key: string) {
    const value = super.get(key);
    if (value !== undefined) {
      super.delete(key);
      super.set(key, value);
    }
    return value;
  }

  override set(key: string, value: WorkerCellSnapshot) {
    super.delete(key);
    super.set(key, value);
    while (this.size > this.maxEntries) {
      const oldest = this.keys().next().value;
      if (oldest === undefined) break;
      super.delete(oldest);
    }
    return this;
  }
}

export function useXlsxViewerController(options: UseXlsxViewerControllerOptions): XlsxViewerController {
  const {
    allowProtectedSheetEditing = false,
    allowResizeInReadOnly = true,
    createWorker,
    deferLoadingAboveBytes = DEFAULT_DEFER_LOADING_ABOVE_BYTES,
    file,
    fileName,
    historyMaxBytes = DEFAULT_HISTORY_MAX_BYTES,
    historyMaxEntries = HISTORY_LIMIT,
    locale: configuredLocale,
    maxFileSizeBytes: configuredMaxFileSizeBytes,
    onDiagnostic,
    readOnly: initialRequestedReadOnly = false,
    readOnlyAboveBytes = 0,
    runtime: configuredRuntime,
    showHiddenSheets: configuredShowHiddenSheets,
    skipXmlParsing: configuredSkipXmlParsing,
    src,
    urlPolicy,
    useWorker = true,
    workerUrl,
  } = options;
  const ownsXlsxRuntime = configuredRuntime === undefined;
  const historyBudget = createHistoryBudget(historyMaxEntries, historyMaxBytes);
  const maxFileSizeBytes = configuredMaxFileSizeBytes
    ?? configuredRuntime?.limits?.maxInputBytes
    ?? DEFAULT_MAX_FILE_SIZE_BYTES;
  const xlsxRuntime = (configuredRuntime as XlsxRuntime | undefined) ?? createXlsxRuntime({
    createWorker,
    workerUrl,
    limits: { maxInputBytes: maxFileSizeBytes },
  });
  const sourceInputLimits = [maxFileSizeBytes, xlsxRuntime.limits?.maxInputBytes]
    .filter((value): value is number => typeof value === "number" && Number.isFinite(value) && value > 0);
  const sourceMaxInputBytes = sourceInputLimits.length > 0
    ? Math.min(...sourceInputLimits)
    : undefined;
  const showHiddenSheets = configuredShowHiddenSheets ?? xlsxRuntime.parseOptions.showHiddenSheets ?? false;
  const skipXmlParsing = configuredSkipXmlParsing ?? xlsxRuntime.parseOptions.skipXmlParsing ?? false;

  // ═══════════════════════════════════════════════════════════════════
  // State
  // ═══════════════════════════════════════════════════════════════════
  const isLoading = ref(Boolean(file ?? src));
  const error = ref<Error | null>(null);
  const sourceState = ref<XlsxSourceState>(file || src ? "loading" : "idle");
  const sourceError = ref<XlsxLoadError | null>(null);
  const workbook = ref<Workbook | null>(null);
  const sheets = ref<XlsxSheetData[]>([]);
  const chartsByWorkbookSheetIndex = ref<XlsxChart[][]>([]);
  const chartsheets = ref<XlsxChartsheet[]>([]);
  const tabs = ref<XlsxWorkbookTab[]>([]);
  const isChartsLoading = ref(false);
  const workerTablesByWorkbookSheetIndex = ref<XlsxTable[][]>([]);
  const formControlsByWorkbookSheetIndex = ref<XlsxFormControl[][]>([]);
  const imagesByWorkbookSheetIndex = ref<XlsxImage[][]>([]);
  const shapesByWorkbookSheetIndex = ref<XlsxShape[][]>([]);
  const activeSheetIndex = ref(0);
  const activeTabIndex = ref(0);
  const zoomScaleOverridesByTabId = ref<Record<string, number>>({});
  const activeCell = ref<XlsxCellAddress | null>(null);
  const selection = ref<XlsxCellRange | null>(null);
  const selectedChartId = ref<string | null>(null);
  const selectedChartElement = ref<XlsxChartElementSelection | null>(null);
  const selectedImageId = ref<string | null>(null);
  const selectedShapeId = ref<string | null>(null);
  const revision = ref(0);
  const selectionAnchorRef = shallowRef<XlsxCellAddress | null>(null);
  const undoStackRef = shallowRef<HistoryEntry[]>([]);
  const redoStackRef = shallowRef<HistoryEntry[]>([]);
  const isApplyingHistoryRef = shallowRef(false);
  const historyRevision = ref(0);
  const shouldAutoCalculate = ref(false);
  const calculationStatus = ref<XlsxCalculationStatus | null>(null);
  let autoCalculateDisabledByFailure = false;
  const workerCellSnapshotRevision = ref(0);
  const isWorkerBacked = ref(false);
  const sortState = ref<XlsxTableSortState | null>(null);
  const forcedReadOnly = ref(false);
  const dataValidationFeedback = ref<XlsxDataValidationFeedback | null>(null);
  const lastExportWarning = ref<string | null>(null);
  const deferredBufferRef = shallowRef<ArrayBuffer | null>(null);
  const deferredLoadRequestIdRef = shallowRef<number | null>(null);
  const deferredLoadRef = shallowRef<DeferredWorkbookLoad | null>(null);
  const sourceBufferRef = shallowRef<ArrayBuffer | null>(null);
  const deferredLoadFileSize = ref<number | null>(null);
  const imageAssetsRef = shallowRef<WorkbookImageAssets | null>(null);
  const chartAssetsRef = shallowRef<WorkbookChartAssets | null>(null);
  const chartLoadRequestTokenRef = shallowRef(0);
  const chartDisplayFallbackCleanupRef = shallowRef<(() => void) | null>(null);
  const sheetOriginsRef = shallowRef<Array<WorkbookImageSheetOrigin | null>>([]);
  const workerClientRef = shallowRef<XlsxWorkerClient | null>(null);
  const workerCellSnapshotCacheRef = shallowRef<Map<string, WorkerCellSnapshot>>(new WorkerCellSnapshotCache());
  const documentLocale = ref(normalizeDocumentLocale(configuredLocale));
  let latestLoadRequestId = 0;
  const xlsxRuntimeId = xlsxRuntime.id;
  const xlsxTaskSequence = createOfficeTaskSequence(xlsxRuntimeId);
  const xlsxLoadCoordinator = createLatestTaskCoordinator(xlsxTaskSequence);
  const disposedWorkbooks = new WeakSet<object>();
  let activeWorksheetHandle: {
    workbook: Workbook;
    workbookSheetIndex: number;
    worksheet: ReturnType<Workbook["getSheet"]>;
  } | null = null;

  function isCurrentLoadRequest(requestId: number): boolean {
    return requestId === latestLoadRequestId;
  }
  function getCurrentLoadRequestId(): number {
    return latestLoadRequestId;
  }

  function reportDiagnostic(diagnostic: XlsxDiagnostic) {
    try {
      onDiagnostic?.({
        ...diagnostic,
        ...(diagnostic.url ? { url: sanitizeOfficeUrl(diagnostic.url) } : {}),
      });
    } catch {
      // Host diagnostics must not alter workbook loading.
    }
  }
  function setExportWarning(message: string | null) {
    lastExportWarning.value = message;
    if (!message) return;
    reportDiagnostic({
      type: "export-warning",
      requestId: latestLoadRequestId,
      sourceKind: src ? "url" : "file",
      message,
    });
  }

  function reportImageDecodeError(imageId: string, message = "XLSX image could not be decoded.") {
    reportDiagnostic({
      type: "image-decode-error",
      requestId: latestLoadRequestId,
      sourceKind: src ? "url" : "file",
      imageId,
      error: {
        code: "IMAGE_DECODE_FAILED",
        message,
        sourceKind: src ? "url" : "file",
        phase: "image-decode",
      },
    });
  }

  /**
   * Select the locale used by number and date formats that name none of their
   * own (`[$-…]` wins over this value) and repaint the workbook.
   */
  function setDocumentLocale(nextLocale?: string) {
    const normalized = normalizeDocumentLocale(nextLocale);
    if (normalized === documentLocale.value) return;
    documentLocale.value = normalized;
    revision.value = revision.value + 1;
  }

  // ═══════════════════════════════════════════════════════════════════
  // Computed
  // ═══════════════════════════════════════════════════════════════════
  const displayFileName = computed(() => resolveDisplayFileName(src, fileName));
  const shouldDeferLoading = deferLoadingAboveBytes > 0;
  // Read through the options object so hosts may provide a reactive getter.
  // Plain boolean options retain the existing one-time configuration behavior.
  const requestedReadOnly = computed(() => Boolean(options.readOnly ?? initialRequestedReadOnly));
  const readOnly = computed(() => requestedReadOnly.value || forcedReadOnly.value);
  const canResizeReadOnly = computed(() => requestedReadOnly.value && allowResizeInReadOnly && !forcedReadOnly.value);
  const workerSupported = useWorker && typeof Worker !== "undefined";
  const canUseWorkerForRequestedReadOnly = initialRequestedReadOnly;

  const activeTab = computed(() => tabs.value[activeTabIndex.value] ?? null);
  const activeSheet = computed(() => {
    const tab = activeTab.value;
    return tab?.kind === "sheet" ? sheets.value[tab.sheetIndex ?? -1] ?? null : null;
  });
  const isActiveSheetProtected = computed(() => activeSheet.value?.protection?.enabled === true);
  const isEditingLocked = computed(() => isSheetEditingLocked(
    readOnly.value,
    isActiveSheetProtected.value,
    allowProtectedSheetEditing
  ));
  watch(() => activeSheet.value?.workbookSheetIndex, () => { dataValidationFeedback.value = null; });
  const activeZoomTabKey = computed(() => activeTab.value?.id ?? DEFAULT_ZOOM_TAB_KEY);
  const defaultZoomScale = computed(() => resolveDefaultZoomScale(activeTab.value, activeSheet.value));
  const zoomScale = computed(() => clampZoomScale(zoomScaleOverridesByTabId.value[activeZoomTabKey.value] ?? defaultZoomScale.value));
  const canZoomIn = computed(() => zoomScale.value < MAX_ZOOM_SCALE);
  const canZoomOut = computed(() => zoomScale.value > MIN_ZOOM_SCALE);
  const isLoadDeferred = computed(() => deferredLoadFileSize.value !== null);
  const canLoadDeferred = computed(() => !isLoading.value && isLoadDeferred.value);
  const canUndo = computed(() => {
    historyRevision.value;
    return !isEditingLocked.value && !isApplyingHistoryRef.value && undoStackRef.value.length > 0;
  });
  const canRedo = computed(() => {
    historyRevision.value;
    return !isEditingLocked.value && !isApplyingHistoryRef.value && redoStackRef.value.length > 0;
  });
  const visibleSheetIndexByWorkbookSheetIndex = computed(() => new Map(sheets.value.map((sheet, index) => [sheet.workbookSheetIndex, index])));
  const deferredMetadataCell = computed(() => activeCell.value);
  const deferredMetadataSheet = computed(() => activeSheet.value);

  // ═══════════════════════════════════════════════════════════════════
  // Helpers (need defined before they can go on ctx)
  // ═══════════════════════════════════════════════════════════════════
  function shouldForceReadOnlyForBuffer(bufferByteLength: number) {
    return !requestedReadOnly.value && readOnlyAboveBytes > 0 && bufferByteLength > readOnlyAboveBytes;
  }
  function shouldUseWorkerForReadOnlyLoad(willForceReadOnly: boolean) {
    return workerSupported && (willForceReadOnly || canUseWorkerForRequestedReadOnly);
  }
  function disposeWorkerClient() { workerClientRef.value?.dispose(); workerClientRef.value = null; }
  function getWorkerClient() {
    if (!workerClientRef.value) {
      workerClientRef.value = xlsxRuntime.createWorkerClient();
    }
    return workerClientRef.value;
  }
  function releaseActiveWorksheetHandle(targetWorkbook?: Workbook | null) {
    const handle = activeWorksheetHandle;
    if (!handle || (targetWorkbook && handle.workbook !== targetWorkbook)) return;
    activeWorksheetHandle = null;
    try {
      handle.worksheet.free();
    } catch {
      // The tracked workbook still performs ordered best-effort cleanup.
    }
  }
  function disposeWorkbook(target: Workbook | null | undefined) {
    if (!target || disposedWorkbooks.has(target)) return;
    releaseActiveWorksheetHandle(target);
    disposedWorkbooks.add(target);
    try {
      target.free();
    } catch {
      // WASM resources are best-effort on teardown; the instance is never reused.
    }
  }
  function replaceWorkbook(nextWorkbook: Workbook | null) {
    const previousWorkbook = workbook.value;
    if (previousWorkbook === nextWorkbook) return;
    workbook.value = nextWorkbook;
    disposeWorkbook(previousWorkbook);
  }
  function discardWorkbookLoadResult(
    nextImageAssets: WorkbookImageAssets | null,
    nextWorkbook: Workbook | null,
  ) {
    revokeWorkbookImageAssets(nextImageAssets);
    disposeWorkbook(nextWorkbook);
  }
  function clearDeferredLoad(expected?: DeferredWorkbookLoad) {
    const current = deferredLoadRef.value;
    if (!current || (expected && current !== expected)) return;
    deferredLoadRef.value = null;
    deferredBufferRef.value = null;
    deferredLoadRequestIdRef.value = null;
    deferredLoadFileSize.value = null;
  }
  function clearImageAssets() {
    revokeWorkbookImageAssets(imageAssetsRef.value);
    imageAssetsRef.value = null;
    sheetOriginsRef.value = [];
    formControlsByWorkbookSheetIndex.value = [];
    imagesByWorkbookSheetIndex.value = [];
    shapesByWorkbookSheetIndex.value = [];
  }
  function clearChartAssets() {
    chartLoadRequestTokenRef.value += 1;
    chartDisplayFallbackCleanupRef.value?.();
    chartDisplayFallbackCleanupRef.value = null;
    chartAssetsRef.value = null;
    chartsByWorkbookSheetIndex.value = [];
    chartsheets.value = [];
    tabs.value = [];
    isChartsLoading.value = false;
  }
  function setImageAssets(assets: WorkbookImageAssets | null) {
    revokeWorkbookImageAssets(imageAssetsRef.value);
    imageAssetsRef.value = assets;
    sheetOriginsRef.value = assets?.sheetOrigins.slice() ?? [];
    formControlsByWorkbookSheetIndex.value = assets?.formControlsByWorkbookSheetIndex ?? [];
    imagesByWorkbookSheetIndex.value = assets?.imagesByWorkbookSheetIndex ?? [];
    shapesByWorkbookSheetIndex.value = assets?.shapesByWorkbookSheetIndex ?? [];
  }
  function setWorkerImageAssets(snapshot: {
    formControlsByWorkbookSheetIndex?: XlsxFormControl[][];
    imagesByWorkbookSheetIndex?: XlsxImage[][];
    objectUrls?: string[];
    shapesByWorkbookSheetIndex?: XlsxShape[][];
    sheets: XlsxSheetData[];
  }) {
    const formControls = snapshot.formControlsByWorkbookSheetIndex ?? [];
    const images = snapshot.imagesByWorkbookSheetIndex ?? [];
    const shapes = snapshot.shapesByWorkbookSheetIndex ?? [];
    const workbookSheetCount = Math.max(
      formControls.length,
      images.length,
      shapes.length,
      ...snapshot.sheets.map((sheet) => sheet.workbookSheetIndex + 1),
      0,
    );
    setImageAssets({
      archive: {},
      formControlOriginsById: new Map(),
      formControlsByWorkbookSheetIndex: formControls,
      imageOriginsById: new Map(),
      imagesByWorkbookSheetIndex: images,
      namedCellStyleByName: {},
      objectUrls: snapshot.objectUrls ?? [],
      shapeOriginsById: new Map(),
      shapesByWorkbookSheetIndex: shapes,
      sheetPathsByWorkbookSheetIndex: [],
      sheetOrigins: Array.from({ length: workbookSheetCount }, () => null),
      sheetStatesByWorkbookSheetIndex: Array.from({ length: workbookSheetCount }, () => null),
      styleById: {},
      tableMetadataByWorkbookSheetIndex: Array.from({ length: workbookSheetCount }, () => []),
      tableStyleByName: {},
      themePalette: { colorsByIndex: {} },
    });
  }
  function revokeWorkerSnapshotObjectUrls(snapshot: { objectUrls?: string[] }) {
    for (const objectUrl of snapshot.objectUrls ?? []) {
      if (objectUrl.startsWith("blob:")) URL.revokeObjectURL(objectUrl);
    }
    snapshot.objectUrls?.splice(0, snapshot.objectUrls.length);
  }
  function setChartAssets(assets: WorkbookChartAssets | null) {
    // Any authoritative chart reload invalidates an in-flight Worker hydration.
    // Otherwise a Worker result parsed from the original bytes can overwrite
    // charts that were already refreshed after an edit.
    chartLoadRequestTokenRef.value += 1;
    chartAssetsRef.value = assets;
    chartsByWorkbookSheetIndex.value = assets?.chartsByWorkbookSheetIndex ?? [];
    chartsheets.value = assets?.chartsheets ?? [];
    tabs.value = assets?.tabs ?? [];
    isChartsLoading.value = false;
  }
  function hasIncompleteWorkerChartSnapshot(snapshot: { chartsByWorkbookSheetIndex: XlsxChart[][] }) {
    for (const sheetCharts of snapshot.chartsByWorkbookSheetIndex) {
      for (const chart of sheetCharts) {
        if (chart.chartType !== "Bubble") continue;
        for (const series of chart.series) {
          const pointCount = Math.max(series.values.length, series.categories.length);
          if (pointCount <= 1) continue;
          const numericBubbleSizes = (series.bubbleSizes ?? []).filter((v): v is number => typeof v === "number" && Number.isFinite(v));
          if (numericBubbleSizes.length < pointCount) return true;
        }
      }
    }
    return false;
  }
  function ensureChartAssetsHydrated(targetWorkbook: Workbook | null, targetSheets: XlsxSheetData[]) {
    const currentAssets = chartAssetsRef.value;
    if (currentAssets && (currentAssets.chartOriginsById.size > 0 || !targetWorkbook || !imageAssetsRef.value)) return currentAssets;
    if (!targetWorkbook || !imageAssetsRef.value) return chartAssetsRef.value;
    const assets = loadWorkbookChartAssets(targetWorkbook, imageAssetsRef.value, buildVisibleSheetIndexMap(targetSheets), showHiddenSheets);
    chartAssetsRef.value = assets;
    return assets;
  }
  function startChartDisplayHydration(buffer: ArrayBuffer, targetWorkbook: Workbook, targetSheets: XlsxSheetData[]) {
    const effectiveSkipXmlParsing = shouldSkipXmlParsingForWorkbook(new Uint8Array(buffer), skipXmlParsing);
    const visibleMap = buildVisibleSheetIndexMap(targetSheets);
    const quickAssets = loadWorkbookChartAssets(targetWorkbook, imageAssetsRef.value, visibleMap, showHiddenSheets);
    setChartAssets(quickAssets);
    if (effectiveSkipXmlParsing) return;
    isChartsLoading.value = true;
    const requestToken = chartLoadRequestTokenRef.value + 1;
    chartLoadRequestTokenRef.value = requestToken;
    chartDisplayFallbackCleanupRef.value?.();
    chartDisplayFallbackCleanupRef.value = null;
    let workerSettled = false;
    const applyWorkerResult = (result: { chartsByWorkbookSheetIndex: XlsxChart[][]; chartsheets: XlsxChartsheet[]; tabs: XlsxWorkbookTab[] }) => {
      if (workerSettled || requestToken !== chartLoadRequestTokenRef.value) return;
      workerSettled = true;
      chartsByWorkbookSheetIndex.value = result.chartsByWorkbookSheetIndex;
      chartsheets.value = result.chartsheets;
      tabs.value = result.tabs;
      isChartsLoading.value = false;
    };
    const finishWorkerFailure = (error: unknown) => {
      if (workerSettled || requestToken !== chartLoadRequestTokenRef.value) return;
      workerSettled = true;
      isChartsLoading.value = false;
      disposeWorkerClient();
      reportDiagnostic({
        type: "worker-error",
        requestId: latestLoadRequestId,
        sourceKind: src ? "url" : "file",
        url: src,
        error: toXlsxLoadError(error, src ? "url" : "file", src),
      });
    };
    const runMainThreadFallback = () => {
      chartDisplayFallbackCleanupRef.value = scheduleLowPriorityTask(() => {
        if (requestToken !== chartLoadRequestTokenRef.value) return;
        try {
          const hydratedAssets = loadWorkbookChartAssets(targetWorkbook, imageAssetsRef.value, visibleMap, showHiddenSheets);
          if (requestToken !== chartLoadRequestTokenRef.value) return;
          setChartAssets(hydratedAssets);
        } catch { if (requestToken !== chartLoadRequestTokenRef.value) return; setChartAssets(quickAssets); }
        finally { if (requestToken === chartLoadRequestTokenRef.value) isChartsLoading.value = false; }
      });
    };
    if (!workerSupported) { runMainThreadFallback(); return; }
    void getWorkerClient().parseCharts(buffer, effectiveSkipXmlParsing, showHiddenSheets)
      .then((result) => {
        const quickChartCount = quickAssets.chartsByWorkbookSheetIndex.reduce((total, charts) => total + charts.length, 0);
        const workerChartCount = result.chartsByWorkbookSheetIndex.reduce((total, charts) => total + charts.length, 0);
        if (workerChartCount < quickChartCount || hasIncompleteWorkerChartSnapshot(result)) {
          finishWorkerFailure(new XlsxWorkerError(
            "WORKER_FAILED",
            "XLSX Worker returned incomplete chart data.",
            { phase: "chart" },
          ));
          return;
        }
        applyWorkerResult(result);
      })
      .catch((error: unknown) => {
        if (isAbortError(error)) return;
        finishWorkerFailure(error);
      });
  }
  function createParseTimeoutError(maxParseMs: number, actual: number) {
    return new XlsxWorkerError(
      "TIMEOUT",
      `XLSX parsing exceeded ${maxParseMs} ms.`,
      {
        phase: "parse",
        limit: "maxParseMs",
        actual,
        allowed: maxParseMs,
      },
    );
  }
  async function loadWorkbookOnMainThread(buffer: ArrayBuffer) {
    const parsePromise = (async () => {
      const nextParsedWorkbook = await parseWorkbookBuffer(buffer);
      try {
        const bytes = new Uint8Array(buffer);
        const nextImageAssets = loadWorkbookImageAssets(bytes, nextParsedWorkbook.workbook, shouldSkipXmlParsingForWorkbook(bytes, skipXmlParsing));
        return { imageAssets: nextImageAssets, parsedWorkbook: nextParsedWorkbook };
      } catch (error) {
        disposeWorkbook(nextParsedWorkbook.workbook);
        throw error;
      }
    })();

    const maxParseMs = xlsxRuntime.limits?.maxParseMs;
    if (maxParseMs === undefined || maxParseMs <= 0) return parsePromise;

    const startedAt = Date.now();
    let timedOut = false;
    let timeout: ReturnType<typeof setTimeout> | undefined;
    const timeoutPromise = new Promise<never>((_resolve, reject) => {
      timeout = setTimeout(() => {
        timedOut = true;
        reject(createParseTimeoutError(maxParseMs, Math.max(maxParseMs + 1, Date.now() - startedAt)));
      }, maxParseMs);
    });
    try {
      const result = await Promise.race([parsePromise, timeoutPromise]);
      const elapsed = Date.now() - startedAt;
      // The synchronous WASM parse cannot be interrupted, so also enforce the
      // budget once it returns instead of only racing module initialization.
      if (elapsed > maxParseMs) {
        discardWorkbookLoadResult(result.imageAssets, result.parsedWorkbook.workbook);
        throw createParseTimeoutError(maxParseMs, elapsed);
      }
      return result;
    } finally {
      if (timeout !== undefined) clearTimeout(timeout);
      if (timedOut) {
        void parsePromise.then((result) => {
          discardWorkbookLoadResult(result.imageAssets, result.parsedWorkbook.workbook);
        }).catch(() => {
          // Late failures after the timeout need no further cleanup.
        });
      }
    }
  }
  function refreshWorkbookState(targetWorkbook: Workbook) {
    const nextSheets = buildSheetList(
      targetWorkbook, imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex,
      imageAssetsRef.value?.themePalette, imageAssetsRef.value?.styleById,
      imageAssetsRef.value?.namedCellStyleByName, imageAssetsRef.value?.tableStyleByName, showHiddenSheets
    );
    sheets.value = nextSheets;
    setChartAssets(loadWorkbookChartAssets(targetWorkbook, imageAssetsRef.value, buildVisibleSheetIndexMap(nextSheets), showHiddenSheets));
    revision.value = revision.value + 1;
  }
  async function refreshChartData() {
    if (!isWorkerBacked.value) {
      if (!workbook.value) return false;
      refreshWorkbookState(workbook.value);
      return true;
    }
    try {
      const result = await getWorkerClient().refreshCharts(showHiddenSheets);
      chartsByWorkbookSheetIndex.value = result.chartsByWorkbookSheetIndex;
      chartsheets.value = result.chartsheets;
      tabs.value = result.tabs;
      revision.value = revision.value + 1;
      return true;
    } catch (error) {
      reportDiagnostic({
        type: "worker-error",
        requestId: latestLoadRequestId,
        sourceKind: src ? "url" : "file",
        url: src,
        error: toXlsxLoadError(error, src ? "url" : "file", src),
      });
      return false;
    }
  }
  // Set by switchToSheet(..., { preserveSelection: true }) and consumed by the
  // tab watcher, which otherwise resets sheet-scoped state on every tab change.
  // The marker expires with the flush that runs the watcher, so a switch that
  // never changes the tab cannot leak into a later user tab click.
  const preservedTabSwitchIndexRef = shallowRef(-1);
  function consumePreservedTabSwitch(nextTabIndex: number) {
    const preservedTabIndex = preservedTabSwitchIndexRef.value;
    if (preservedTabIndex < 0) return false;
    preservedTabSwitchIndexRef.value = -1;
    return preservedTabIndex === nextTabIndex;
  }
  function resolveSheetTabIndex(visibleSheetIndex: number) {
    const targetSheet = sheets.value[visibleSheetIndex];
    if (!targetSheet) return -1;
    const tabIndexBySheetIndex = tabs.value.findIndex(
      (tab) => tab.kind === "sheet" && tab.sheetIndex === visibleSheetIndex,
    );
    if (tabIndexBySheetIndex >= 0) return tabIndexBySheetIndex;
    return tabs.value.findIndex(
      (tab) => tab.kind === "sheet" && tab.workbookSheetIndex === targetSheet.workbookSheetIndex,
    );
  }
  function switchToSheet(visibleSheetIndex: number, options?: { preserveSelection?: boolean }) {
    if (visibleSheetIndex < 0 || visibleSheetIndex >= sheets.value.length) return;
    activeSheetIndex.value = visibleSheetIndex;
    const nextTabIndex = resolveSheetTabIndex(visibleSheetIndex);
    if (nextTabIndex < 0 || nextTabIndex === activeTabIndex.value) return;
    if (options?.preserveSelection) {
      preservedTabSwitchIndexRef.value = nextTabIndex;
      void nextTick(() => {
        if (preservedTabSwitchIndexRef.value === nextTabIndex) preservedTabSwitchIndexRef.value = -1;
      });
    }
    activeTabIndex.value = nextTabIndex;
  }
  function getInternalActiveWorksheet() {
    const activeWorkbook = workbook.value;
    const workbookSheetIndex = activeSheet.value?.workbookSheetIndex;
    if (!activeWorkbook || workbookSheetIndex === undefined) {
      releaseActiveWorksheetHandle();
      return null;
    }
    if (
      activeWorksheetHandle?.workbook === activeWorkbook &&
      activeWorksheetHandle.workbookSheetIndex === workbookSheetIndex
    ) {
      return activeWorksheetHandle.worksheet;
    }
    releaseActiveWorksheetHandle();
    const worksheet = activeWorkbook.getSheet(workbookSheetIndex);
    activeWorksheetHandle = { workbook: activeWorkbook, workbookSheetIndex, worksheet };
    return worksheet;
  }
  function readValidationRangeValues(reference: string): string[] | null {
    const activeWorkbook = workbook.value;
    if (!activeWorkbook) return null;

    const bangIndex = reference.lastIndexOf("!");
    const rawSheetName = bangIndex >= 0 ? reference.slice(0, bangIndex) : null;
    const rangeReference = (bangIndex >= 0 ? reference.slice(bangIndex + 1) : reference).replace(/\$/g, "");
    const range = parseA1RangeReference(rangeReference);
    if (!range) return null;

    let worksheet: ReturnType<Workbook["getSheet"]> | null = null;
    let ownsWorksheet = false;
    if (rawSheetName) {
      const sheetName = rawSheetName.replace(/^'|'$/g, "").replace(/''/g, "'");
      try {
        worksheet = activeWorkbook.getSheetByName(sheetName) ?? null;
      } catch {
        return null;
      }
      ownsWorksheet = true;
    } else {
      worksheet = getInternalActiveWorksheet();
    }
    if (!worksheet) return null;

    try {
      const values: string[] = [];
      for (let row = range.start.row; row <= range.end.row; row += 1) {
        for (let col = range.start.col; col <= range.end.col; col += 1) {
          const value = worksheet.getFormattedValueAt(row, col);
          if (value) values.push(value);
        }
      }
      return values;
    } catch {
      return null;
    } finally {
      if (ownsWorksheet) {
        try { worksheet.free(); } catch { /* best-effort cleanup */ }
      }
    }
  }

  function getCellDataValidation(cell?: XlsxCellAddress | null): XlsxDataValidation | null {
    return findDataValidationForCell(activeSheet.value?.dataValidations, cell ?? activeCell.value);
  }

  function getDataValidationListOptions(validation: XlsxDataValidation): string[] {
    return resolveDataValidationListOptions(validation, {
      getNamedRange: (name) => workbook.value?.getNamedRange(name) ?? null,
      readRangeValues: readValidationRangeValues
    });
  }

  function checkCellValue(cell: XlsxCellAddress, value: string) {
    return checkCellValueAgainstDataValidation(activeSheet.value, cell, value, {
      getNamedRange: (name) => workbook.value?.getNamedRange(name) ?? null,
      readRangeValues: readValidationRangeValues
    });
  }

  function setValidationFeedback(feedback: XlsxDataValidationFeedback | null) {
    dataValidationFeedback.value = feedback;
  }

  function dismissValidationFeedback() {
    dataValidationFeedback.value = null;
  }

  function hasFormulaCells(targetWorkbook: Workbook): boolean {
    for (let index = 0; index < targetWorkbook.sheetCount; index += 1) {
      let worksheet;
      try {
        worksheet = targetWorkbook.getSheet(index);
      } catch {
        continue;
      }
      try {
        if (worksheet.formulaCount > 0) return true;
      } finally {
        worksheet.free();
      }
    }
    return false;
  }
  function calculationMessage(reason: SafeCalculateSkipReason | "recovery-snapshot-unavailable"): string {
    if (reason === "unresolved-sheet-refs") {
      return "A formula references a sheet that no longer exists; cached values are shown.";
    }
    if (reason === "recovery-snapshot-unavailable") {
      return "The workbook snapshot could not be saved, so recalculation was skipped to protect the current data; cached values are shown.";
    }
    return "Formulas could not be recalculated; cached values are shown. Recalculate manually to retry.";
  }
  function noteCalculationResult(result: {
    calculated: boolean;
    skipReason: SafeCalculateSkipReason | "recovery-snapshot-unavailable" | null;
  }) {
    if (result.calculated) {
      calculationStatus.value = null;
      return;
    }
    if (!result.skipReason) return;
    calculationStatus.value = {
      reason: result.skipReason,
      message: calculationMessage(result.skipReason)
    };
    shouldAutoCalculate.value = false;
    autoCalculateDisabledByFailure = true;
  }
  function resumeAutoCalculate() {
    if (!autoCalculateDisabledByFailure) return;
    autoCalculateDisabledByFailure = false;
    shouldAutoCalculate.value = true;
  }
  function recalculateWorkbookSafely(targetWorkbook: Workbook): XlsxCalculationOutcome {
    const WorkbookConstructor = targetWorkbook.constructor as typeof Workbook;
    let snapshot: Uint8Array | null = null;
    if (hasFormulaCells(targetWorkbook)) {
      try {
        snapshot = createSavedWorkbookBytes(targetWorkbook);
      } catch (error) {
        console.warn("[xlsx] failed to snapshot workbook before recalculation", error);
        noteCalculationResult({ calculated: false, skipReason: "recovery-snapshot-unavailable" });
        return { calculated: false, safeToRefresh: true };
      }
    }

    const result = safeCalculate(targetWorkbook, snapshot ? {
      // The snapshot contains the mutation that triggered this recalculation,
      // so rebuilding from it keeps the edit while replacing the trapped WASM
      // instance with a readable one.
      reparse: () => trackXlsxWorkbookLifetime(WorkbookConstructor.fromBytes(snapshot as Uint8Array))
    } : {});
    if (result.workbook !== targetWorkbook) {
      replaceWorkbook(result.workbook);
    }
    noteCalculationResult({ calculated: result.calculated, skipReason: result.skipReason });
    if (!result.calculated && result.skipReason === "calculate-trapped" && result.workbook === targetWorkbook) {
      // The reparse itself failed; the trapped instance must never be read.
      replaceWorkbook(null);
      return { calculated: false, safeToRefresh: false };
    }
    return { calculated: result.calculated, safeToRefresh: true };
  }
  function maybeRecalculateWorkbook(targetWorkbook: Workbook) {
    if (!shouldAutoCalculate.value) return true;
    return recalculateWorkbookSafely(targetWorkbook).safeToRefresh;
  }
  function createSavedWorkbookBytes(targetWorkbook: Workbook, reportSanitizationWarning = false) {
    const sanitized = sanitizeSavedWorkbookBytes(targetWorkbook.saveXlsxBytes());
    if (sanitized.error) {
      console.warn("[xlsx] failed to sanitize saved workbook XML", sanitized.error);
      if (reportSanitizationWarning) {
        setExportWarning("The XML escape repair before export failed; the exported workbook may contain double escaping.");
      }
    }
    return mergeWorkbookImageAssets(sanitized.bytes, imageAssetsRef.value, sheetOriginsRef.value);
  }
  function captureCellMutationState(cell: XlsxCellAddress) {
    const worksheet = getInternalActiveWorksheet();
    if (!worksheet) return null;
    const cellValue = worksheet.getCellAt(cell.row, cell.col);
    try {
      return {
        formula: worksheet.getFormulaAt(cell.row, cell.col) ?? null,
        style: worksheet.getCellStyleAt(cell.row, cell.col),
        value: cellValue.toJs()
      };
    } finally {
      cellValue.free();
    }
  }
  function recordHistoryBeforeMutation() {
    if (isApplyingHistoryRef.value || !workbook.value) return;
    pushHistoryEntry(undoStackRef.value, {
      kind: "snapshot", activeCell: activeCell.value, activeSheetIndex: activeSheetIndex.value,
      bytes: createSavedWorkbookBytes(workbook.value), selection: selection.value
    }, historyBudget);
    redoStackRef.value = [];
    historyRevision.value = historyRevision.value + 1;
  }
  function recordCellEditHistory(cell: XlsxCellAddress, before: CellMutationState, after: CellMutationState) {
    if (!activeSheet.value || isApplyingHistoryRef.value) return;
    pushHistoryEntry(undoStackRef.value, {
      kind: "cell-edit", activeCellAfter: cell, activeCellBefore: activeCell.value,
      after, before, cell,
      selectionAfter: { start: cell, end: cell }, selectionBefore: selection.value,
      sheetIndex: activeSheet.value.workbookSheetIndex
    }, historyBudget);
    redoStackRef.value = [];
    historyRevision.value = historyRevision.value + 1;
  }
  function recordRangeEditHistory(
    mutations: RangeCellMutation[],
    selectionAfter: XlsxCellRange | null,
    activeCellAfter: XlsxCellAddress | null,
    before?: SelectionSnapshot
  ) {
    if (!activeSheet.value || isApplyingHistoryRef.value || mutations.length === 0) return;
    pushHistoryEntry(undoStackRef.value, {
      kind: "range-edit", activeCellAfter, activeCellBefore: before ? before.activeCell : activeCell.value,
      mutations, selectionAfter, selectionBefore: before ? before.selection : selection.value,
      sheetIndex: activeSheet.value.workbookSheetIndex
    }, historyBudget);
    redoStackRef.value = [];
    historyRevision.value = historyRevision.value + 1;
  }

  function isCurrentWorkbookLoad(load: DeferredWorkbookLoad) {
    return load.task.isCurrent() && isCurrentLoadRequest(load.requestId);
  }

  async function executeWorkbookLoad(load: DeferredWorkbookLoad) {
    const { requestId, resolvedUrl, sourceBuffer, sourceKind, task } = load;
    if (!isCurrentWorkbookLoad(load)) return;
    if (maxFileSizeBytes > 0 && sourceBuffer.byteLength > maxFileSizeBytes) {
      throw new XlsxFileSizeLimitExceededError(sourceBuffer.byteLength, maxFileSizeBytes);
    }
    if (load.inputFormat === "csv" && !load.csvConverted) {
      load.buffer = await convertCsvBufferToXlsx(sourceBuffer);
      load.csvConverted = true;
      if (!isCurrentWorkbookLoad(load)) return;
    }
    if (load.inputFormat === "xlsb" && !load.xlsbConverted) {
      validateXlsxArchive(sourceBuffer, xlsxRuntime.limits, task.signal);
      load.buffer = await convertXlsbBufferToXlsx(sourceBuffer);
      load.xlsbConverted = true;
      load.archiveValidated = false;
      if (!isCurrentWorkbookLoad(load)) return;
    }
    const buffer = load.buffer;
    if (!load.archiveValidated) {
      validateXlsxArchive(buffer, xlsxRuntime.limits, task.signal);
      load.archiveValidated = true;
    }
    const preflight = preflightWorkbookBuffer(buffer);
    if (preflight?.tooLarge) throw createWorkbookTooLargeError(preflight);

    const willForceReadOnly = shouldForceReadOnlyForBuffer(sourceBuffer.byteLength);
    forcedReadOnly.value = willForceReadOnly;
    const shouldUseWorkerForLoad = shouldUseWorkerForReadOnlyLoad(willForceReadOnly);
    const effectiveSkipXmlParsing = shouldSkipXmlParsingForWorkbook(new Uint8Array(buffer), skipXmlParsing);

    if (shouldUseWorkerForLoad) {
      const client = getWorkerClient();
      const terminateOwnedWorker = () => {
        if (workerClientRef.value === client) disposeWorkerClient();
      };
      task.signal.addEventListener("abort", terminateOwnedWorker, { once: true });
      try {
        const snapshot = await client.loadWorkbook(
          buffer,
          effectiveSkipXmlParsing,
          showHiddenSheets,
          task.signal,
        );
        if (!isCurrentWorkbookLoad(load)) {
          revokeWorkerSnapshotObjectUrls(snapshot);
          return;
        }
        if (!effectiveSkipXmlParsing && hasIncompleteWorkerChartSnapshot(snapshot)) {
          // Treat this as a Worker defect, not a workbook failure: the catch
          // below reports a diagnostic and retries on the main thread.
          throw new XlsxWorkerError(
            "WORKER_FAILED",
            "XLSX Worker returned incomplete chart data.",
            { phase: "chart" },
          );
        }
        replaceWorkbook(null);
        setWorkerImageAssets(snapshot);
        sheets.value = snapshot.sheets;
        chartsByWorkbookSheetIndex.value = snapshot.chartsByWorkbookSheetIndex;
        chartsheets.value = snapshot.chartsheets;
        tabs.value = snapshot.tabs;
        chartLoadRequestTokenRef.value += 1;
        chartAssetsRef.value = null;
        workerTablesByWorkbookSheetIndex.value = snapshot.tablesByWorkbookSheetIndex;
        shouldAutoCalculate.value = false;
        calculationStatus.value = null;
        autoCalculateDisabledByFailure = false;
        isWorkerBacked.value = true;
        sourceBufferRef.value = sourceBuffer;
        sortState.value = null;
        isChartsLoading.value = false;
        isLoading.value = false;
        sourceState.value = "ready";
        sourceError.value = null;
        load.state = "settled";
        reportDiagnostic({
          type: "load-success",
          requestId,
          taskId: task.context.taskId,
          sourceKind,
          url: resolvedUrl,
          bytes: sourceBuffer.byteLength,
        });
        return;
      } catch (workerError) {
        if (!isCurrentWorkbookLoad(load) || isAbortError(workerError)) throw workerError;
        disposeWorkerClient();
        reportDiagnostic({
          type: "worker-error",
          requestId,
          taskId: task.context.taskId,
          sourceKind,
          url: resolvedUrl,
          error: toXlsxLoadError(workerError, sourceKind, resolvedUrl),
        });
        if (
          workerError instanceof XlsxWorkerError
          && (workerError.code === "LIMIT_EXCEEDED" || workerError.code === "INVALID_SOURCE")
        ) {
          throw workerError;
        }
        // A Worker-side failure (protocol error, timeout, incomplete chart
        // payload) must not fail a legal workbook: retry once on the main
        // thread, which is also where non-Worker loads normally run.
      } finally {
        task.signal.removeEventListener("abort", terminateOwnedWorker);
      }
    }

    const { imageAssets: nextImageAssets, parsedWorkbook: nextParsedWorkbook } = await loadWorkbookOnMainThread(buffer);
    let committed = false;
    try {
      if (!isCurrentWorkbookLoad(load)) return;
      const nextSheets = buildSheetList(
        nextParsedWorkbook.workbook,
        nextImageAssets.sheetStatesByWorkbookSheetIndex,
        nextImageAssets.themePalette,
        nextImageAssets.styleById,
        nextImageAssets.namedCellStyleByName,
        nextImageAssets.tableStyleByName,
        showHiddenSheets,
      );
      if (!isCurrentWorkbookLoad(load)) return;
      setImageAssets(nextImageAssets);
      replaceWorkbook(nextParsedWorkbook.workbook);
      committed = true;
      sheets.value = nextSheets;
      startChartDisplayHydration(buffer, nextParsedWorkbook.workbook, nextSheets);
      shouldAutoCalculate.value = nextParsedWorkbook.shouldAutoCalculate;
      autoCalculateDisabledByFailure = false;
      noteCalculationResult({
        calculated: nextParsedWorkbook.shouldAutoCalculate,
        skipReason: nextParsedWorkbook.calculationSkipReason
      });
      workerTablesByWorkbookSheetIndex.value = [];
      isWorkerBacked.value = false;
      sourceBufferRef.value = sourceBuffer;
      sortState.value = null;
      isLoading.value = false;
      sourceState.value = "ready";
      sourceError.value = null;
      load.state = "settled";
      reportDiagnostic({
        type: "load-success",
        requestId,
        taskId: task.context.taskId,
        sourceKind,
        url: resolvedUrl,
        bytes: sourceBuffer.byteLength,
      });
    } finally {
      if (!committed) {
        discardWorkbookLoadResult(nextImageAssets, nextParsedWorkbook.workbook);
      }
    }
  }

  function handleWorkbookLoadFailure(load: DeferredWorkbookLoad, nextError: unknown) {
    if (!isCurrentWorkbookLoad(load)) return;
    const { requestId, resolvedUrl, sourceKind, task } = load;
    load.state = "settled";
    if (isAbortError(nextError)) {
      const cancelledError = toXlsxLoadError(nextError, sourceKind, resolvedUrl);
      isLoading.value = false;
      sourceState.value = "idle";
      sourceError.value = cancelledError;
      reportDiagnostic({
        type: "load-cancelled",
        requestId,
        taskId: task.context.taskId,
        sourceKind,
        url: resolvedUrl,
        error: cancelledError,
      });
      return;
    }
    replaceWorkbook(null);
    sheets.value = [];
    clearChartAssets();
    workerTablesByWorkbookSheetIndex.value = [];
    clearImageAssets();
    shouldAutoCalculate.value = false;
    calculationStatus.value = null;
    autoCalculateDisabledByFailure = false;
    isWorkerBacked.value = false;
    sourceBufferRef.value = null;
    sortState.value = null;
    const loadError = toXlsxLoadError(nextError, sourceKind, resolvedUrl);
    error.value = nextError instanceof Error ? nextError : new Error(loadError.message);
    isLoading.value = false;
    sourceState.value = "error";
    sourceError.value = loadError;
    reportDiagnostic({
      type: "load-error",
      requestId,
      taskId: task.context.taskId,
      sourceKind,
      url: resolvedUrl,
      error: loadError,
    });
  }

  function continueDeferredWorkbookLoad() {
    const load = deferredLoadRef.value;
    if (!load || load.state !== "deferred" || !isCurrentWorkbookLoad(load)) return;
    load.state = "running";
    isLoading.value = true;
    error.value = null;
    sourceState.value = "loading";
    sourceError.value = null;
    reportDiagnostic({
      type: "load-resumed",
      requestId: load.requestId,
      taskId: load.task.context.taskId,
      sourceKind: load.sourceKind,
      url: load.resolvedUrl,
      bytes: load.sourceBuffer.byteLength,
    });
    void executeWorkbookLoad(load)
      .catch((nextError: unknown) => handleWorkbookLoadFailure(load, nextError))
      .finally(() => {
        clearDeferredLoad(load);
        load.task.finish();
      });
  }

  function cancelDeferredWorkbookLoad() {
    const load = deferredLoadRef.value;
    if (!load || load.state !== "deferred" || !isCurrentWorkbookLoad(load)) return;
    load.state = "settled";
    clearDeferredLoad(load);
    replaceWorkbook(null);
    sheets.value = [];
    clearChartAssets();
    workerTablesByWorkbookSheetIndex.value = [];
    clearImageAssets();
    error.value = null;
    isLoading.value = false;
    isWorkerBacked.value = false;
    sourceState.value = "idle";
    sourceError.value = null;
    sourceBufferRef.value = null;
    activeSheetIndex.value = 0;
    activeTabIndex.value = 0;
    activeCell.value = null;
    selection.value = null;
    selectedChartId.value = null;
    selectedChartElement.value = null;
    selectedImageId.value = null;
    selectedShapeId.value = null;
    selectionAnchorRef.value = null;
    sortState.value = null;
    reportDiagnostic({
      type: "load-cancelled",
      requestId: load.requestId,
      taskId: load.task.context.taskId,
      sourceKind: load.sourceKind,
      url: load.resolvedUrl,
      bytes: load.sourceBuffer.byteLength,
    });
    load.task.finish();
  }

  // ═══════════════════════════════════════════════════════════════════
  // Lifecycle
  // ═══════════════════════════════════════════════════════════════════
  onUnmounted(() => {
    latestLoadRequestId += 1;
    xlsxLoadCoordinator.dispose();
    xlsxTaskSequence.dispose();
    if (ownsXlsxRuntime) xlsxRuntime.dispose();
    const deferredLoad = deferredLoadRef.value;
    clearDeferredLoad(deferredLoad ?? undefined);
    deferredLoad?.task.finish();
    sourceBufferRef.value = null;
    sourceState.value = "idle";
    sourceError.value = null;
    isApplyingHistoryRef.value = false;
    chartDisplayFallbackCleanupRef.value?.();
    chartDisplayFallbackCleanupRef.value = null;
    revokeWorkbookImageAssets(imageAssetsRef.value);
    imageAssetsRef.value = null;
    replaceWorkbook(null);
    disposeWorkerClient();
  });

  // ═══════════════════════════════════════════════════════════════════
  // Loading watcher
  // ═══════════════════════════════════════════════════════════════════
  watch(() => [file, src], (_value, _oldValue, onCleanup) => {
    if (!file && !src) {
      latestLoadRequestId += 1;
      xlsxLoadCoordinator.cancel();
      const deferredLoad = deferredLoadRef.value;
      clearDeferredLoad(deferredLoad ?? undefined);
      deferredLoad?.task.finish();
      disposeWorkerClient();
      forcedReadOnly.value = false;
      replaceWorkbook(null); sheets.value = []; clearChartAssets();
      workerTablesByWorkbookSheetIndex.value = []; clearImageAssets();
      error.value = null; sourceState.value = "idle"; sourceError.value = null; isLoading.value = false; isWorkerBacked.value = false;
      lastExportWarning.value = null;
      sourceBufferRef.value = null;
      activeSheetIndex.value = 0; activeTabIndex.value = 0;
      activeCell.value = null; selection.value = null;
      selectedChartId.value = null; selectedChartElement.value = null; selectedImageId.value = null; selectedShapeId.value = null;
      selectionAnchorRef.value = null;
      undoStackRef.value = []; redoStackRef.value = []; isApplyingHistoryRef.value = false; historyRevision.value = 0;
      shouldAutoCalculate.value = false;
      calculationStatus.value = null;
      autoCalculateDisabledByFailure = false;
      workerCellSnapshotCacheRef.value.clear(); workerCellSnapshotRevision.value = 0;
      sortState.value = null; zoomScaleOverridesByTabId.value = {}; revision.value = 0;
      return;
    }
    const requestId = ++latestLoadRequestId;
    const sourceKind = src ? "url" as const : "file" as const;
    const officeSource: OfficeSource = src
      ? { kind: "url", url: src, name: fileName }
      : { kind: "bytes", bytes: file as ArrayBuffer, name: fileName };
    const loadTask = xlsxLoadCoordinator.start(officeSource, {
      limits: { maxInputBytes: sourceMaxInputBytes },
      resources: {
        workerUrl: xlsxRuntime.workerUrl,
        wasmUrl: typeof xlsxRuntime.wasmSource === "string" ? xlsxRuntime.wasmSource : undefined,
      },
      urlPolicy: urlPolicy as OfficeUrlPolicy | undefined,
    });
    isLoading.value = true; error.value = null; sourceState.value = "loading"; sourceError.value = null; clearImageAssets(); clearChartAssets();
    lastExportWarning.value = null;
    replaceWorkbook(null);
    sheets.value = [];
    workerTablesByWorkbookSheetIndex.value = []; isWorkerBacked.value = false;
    deferredLoadRef.value = null; deferredBufferRef.value = null; deferredLoadRequestIdRef.value = null; sourceBufferRef.value = null; deferredLoadFileSize.value = null;
    activeSheetIndex.value = 0; activeTabIndex.value = 0;
    activeCell.value = null; selection.value = null;
    selectedChartId.value = null; selectedChartElement.value = null; selectedImageId.value = null; selectedShapeId.value = null;
    selectionAnchorRef.value = null;
    undoStackRef.value = []; redoStackRef.value = []; isApplyingHistoryRef.value = false; historyRevision.value = 0;
    shouldAutoCalculate.value = false;
    calculationStatus.value = null;
    autoCalculateDisabledByFailure = false;
    workerCellSnapshotCacheRef.value.clear(); workerCellSnapshotRevision.value = 0;
    sortState.value = null; zoomScaleOverridesByTabId.value = {}; revision.value = 0;
    disposeWorkerClient();
    reportDiagnostic({ type: "load-start", requestId, taskId: loadTask.context.taskId, sourceKind, url: src });

    const activeLoad: DeferredWorkbookLoad = {
      buffer: new ArrayBuffer(0),
      inputFormat: "workbook",
      sourceBuffer: new ArrayBuffer(0),
      requestId,
      resolvedUrl: src,
      sourceKind,
      state: "running",
      task: loadTask,
    };
    let handedToDeferredLoad = false;

    void resolveWorkbookSource(
      { file, fileName, src, urlPolicy },
      loadTask.signal,
      { maxInputBytes: sourceMaxInputBytes },
    )
      .then(async (resolvedSource) => {
        const sourceBuffer = resolvedSource.buffer;
        activeLoad.buffer = sourceBuffer;
        activeLoad.inputFormat = resolvedSource.inputFormat;
        activeLoad.sourceBuffer = sourceBuffer;
        activeLoad.resolvedUrl = resolvedSource.resolvedUrl;
        activeLoad.sourceKind = resolvedSource.sourceKind;
        if (!isCurrentWorkbookLoad(activeLoad)) return;
        if (maxFileSizeBytes > 0 && sourceBuffer.byteLength > maxFileSizeBytes) {
          throw new XlsxFileSizeLimitExceededError(sourceBuffer.byteLength, maxFileSizeBytes);
        }
        if (resolvedSource.inputFormat !== "csv") {
          validateXlsxArchive(sourceBuffer, xlsxRuntime.limits, loadTask.signal);
          activeLoad.archiveValidated = resolvedSource.inputFormat === "workbook";
          if (resolvedSource.inputFormat === "workbook") {
            const preflight = preflightWorkbookBuffer(sourceBuffer);
            if (preflight?.tooLarge) throw createWorkbookTooLargeError(preflight);
          }
        }
        if (shouldDeferLoading && sourceBuffer.byteLength > deferLoadingAboveBytes) {
          activeLoad.state = "deferred";
          handedToDeferredLoad = true;
          deferredLoadRef.value = activeLoad;
          deferredBufferRef.value = sourceBuffer;
          deferredLoadRequestIdRef.value = requestId;
          deferredLoadFileSize.value = sourceBuffer.byteLength;
          replaceWorkbook(null);
          sheets.value = [];
          clearChartAssets();
          workerTablesByWorkbookSheetIndex.value = [];
          isLoading.value = false;
          reportDiagnostic({
            type: "load-deferred",
            requestId,
            taskId: loadTask.context.taskId,
            sourceKind: activeLoad.sourceKind,
            url: activeLoad.resolvedUrl,
            bytes: sourceBuffer.byteLength,
          });
          return;
        }
        await executeWorkbookLoad(activeLoad);
      })
      .catch((nextError: unknown) => handleWorkbookLoadFailure(activeLoad, nextError))
      .finally(() => {
        if (!handedToDeferredLoad) loadTask.finish();
      });
    onCleanup(() => {
      const shouldReportCancellation = loadTask.isCurrent() && activeLoad.state !== "settled";
      xlsxLoadCoordinator.cancel();
      disposeWorkerClient();
      if (deferredLoadRef.value === activeLoad) {
        clearDeferredLoad(activeLoad);
        loadTask.finish();
      }
      if (shouldReportCancellation) {
        reportDiagnostic({
          type: "load-cancelled",
          requestId,
          taskId: loadTask.context.taskId,
          sourceKind: activeLoad.sourceKind,
          url: activeLoad.resolvedUrl,
        });
      }
    });
  }, { immediate: true });

  // ═══════════════════════════════════════════════════════════════════
  // Domain ctx and factories
  // ═══════════════════════════════════════════════════════════════════
  // Placeholder for setChartSeriesFormula — will be replaced by chartDomain
  let _setChartSeriesFormula: (chartId: string, seriesIndex: number, formula: string) => boolean = () => false;
  const selectedFormulaTarget = computed<{ kind: "chartSeries"; chartId: string; seriesId: string; seriesIndex: number } | { kind: "cell"; cell: XlsxCellAddress | null }>(() => {
    const sf = selectedFormula.value;
    const el = selectedChartElement.value;
    if (sf && el && el.kind !== "chart") {
      return { kind: "chartSeries", chartId: el.chartId, seriesId: el.seriesId, seriesIndex: el.seriesIndex };
    }
    return { kind: "cell", cell: deferredMetadataCell.value };
  });

  const ctx = {
    file, src, fileName, maxFileSizeBytes, deferLoadingAboveBytes, skipXmlParsing, showHiddenSheets,
    workerSupported, shouldDeferLoading, canUseWorkerForRequestedReadOnly, requestedReadOnly: initialRequestedReadOnly,
    isLoading, error, sourceState, sourceError, workbook, sheets, chartsByWorkbookSheetIndex, chartsheets, tabs,
    isChartsLoading, workerTablesByWorkbookSheetIndex, formControlsByWorkbookSheetIndex,
    imagesByWorkbookSheetIndex, shapesByWorkbookSheetIndex, activeSheetIndex, activeTabIndex,
    zoomScaleOverridesByTabId, activeCell, selection, selectedChartId, selectedChartElement,
    selectedImageId, selectedShapeId, revision, selectionAnchorRef, undoStackRef, redoStackRef, historyBudget,
    isApplyingHistoryRef, historyRevision, shouldAutoCalculate, workerCellSnapshotRevision,
    isWorkerBacked, sortState, forcedReadOnly, dataValidationFeedback, deferredBufferRef, deferredLoadRequestIdRef, sourceBufferRef, deferredLoadFileSize,
    imageAssetsRef, chartAssetsRef, chartLoadRequestTokenRef, chartDisplayFallbackCleanupRef,
    sheetOriginsRef, workerClientRef, workerCellSnapshotCacheRef, locale: documentLocale,
    activeTab, activeSheet, deferredMetadataCell, deferredMetadataSheet,
    activeZoomTabKey, defaultZoomScale, zoomScale, canZoomIn, canZoomOut,
    readOnly, isActiveSheetProtected, isEditingLocked, canResizeReadOnly, displayFileName, visibleSheetIndexByWorkbookSheetIndex,
    refreshWorkbookState, maybeRecalculateWorkbook,
    switchToSheet, consumePreservedTabSwitch,
    recalculateWorkbookSafely, resumeAutoCalculate, calculationStatus, noteCalculationResult,
    getActiveWorksheet: getInternalActiveWorksheet,
    setChartAssets, setImageAssets, clearImageAssets, clearChartAssets,
    getWorkerClient, disposeWorkerClient, replaceWorkbook, disposeWorkbook, discardWorkbookLoadResult,
    isCurrentLoadRequest, getCurrentLoadRequestId, continueDeferredWorkbookLoad, cancelDeferredWorkbookLoad,
    startChartDisplayHydration, loadWorkbookOnMainThread,
    hasIncompleteWorkerChartSnapshot,
    ensureChartAssetsHydrated,
    shouldForceReadOnlyForBuffer, shouldUseWorkerForReadOnlyLoad,
    createSavedWorkbookBytes, captureCellMutationState,
    checkCellValue, getCellDataValidation, getDataValidationListOptions,
    setValidationFeedback, dismissValidationFeedback,
    lastExportWarning, setExportWarning,
    recordHistoryBeforeMutation, recordCellEditHistory, recordRangeEditHistory,
    selectedFormulaTarget,
    get setChartSeriesFormula() { return _setChartSeriesFormula; },
    set setChartSeriesFormula(v: typeof _setChartSeriesFormula) { _setChartSeriesFormula = v; },
  };

  // Initialize domains (order matters — navigation registers watchers, editing depends on ctx)
  const navigationDomain = createNavigationDomain(ctx);
  const clipboardDomain = createClipboardDomain(ctx);
  const chartImageDomain = createChartImageDomain(ctx);
  const historyDomain = createHistoryDomain(ctx);
  const editingDomain = createEditingDomain(ctx);

  // Replace placeholder with actual implementation from chart domain
  _setChartSeriesFormula = chartImageDomain.setChartSeriesFormula;

  function downloadSource() {
    const sourceBuffer = sourceBufferRef.value;
    if (!sourceBuffer) return;
    historyDomain.download();
    reportDiagnostic({
      type: "download",
      requestId: latestLoadRequestId,
      sourceKind: src ? "url" : "file",
      url: src,
      bytes: sourceBuffer.byteLength,
    });
  }

  // ═══════════════════════════════════════════════════════════════════
  // Additional watchers (worker cell snapshot)
  // ═══════════════════════════════════════════════════════════════════
  watch(
    () => [deferredMetadataCell.value, deferredMetadataSheet.value, isWorkerBacked.value],
    (_value, _oldValue, onCleanup) => {
      if (!isWorkerBacked.value || !deferredMetadataSheet.value || !deferredMetadataCell.value) return;
      const cacheKey = `${deferredMetadataSheet.value.workbookSheetIndex}:${deferredMetadataCell.value.row}:${deferredMetadataCell.value.col}`;
      if (workerCellSnapshotCacheRef.value.has(cacheKey)) return;
      let isCurrent = true;
      const controller = new AbortController();
      void historyDomain.getCellSnapshotAsync(
        deferredMetadataSheet.value.workbookSheetIndex,
        deferredMetadataCell.value.row,
        deferredMetadataCell.value.col,
        controller.signal,
      )
        .then((snapshot) => {
          if (!isCurrent) return;
          workerCellSnapshotCacheRef.value.set(cacheKey, snapshot);
          workerCellSnapshotRevision.value = workerCellSnapshotRevision.value + 1;
        })
        .catch((error: unknown) => {
          if (!isCurrent || isAbortError(error)) return;
          // Transient Worker failures must not be cached as a permanent empty
          // cell. Recreate the client so the next selection can retry.
          disposeWorkerClient();
        });
      onCleanup(() => {
        isCurrent = false;
        controller.abort();
      });
    }
  );

  // ═══════════════════════════════════════════════════════════════════
  // Computed (post-domains — depend on domain results)
  // ═══════════════════════════════════════════════════════════════════
  const activeCellAddress = computed(() => activeCell.value ? cellAddressToA1(activeCell.value) : null);
  const selectedRangeAddress = computed(() => selection.value ? rangeToA1(selection.value) : null);
  const selectedValue = computed(() => {
    workerCellSnapshotRevision.value;
    return clipboardDomain.getCellDisplayValue(deferredMetadataCell.value);
  });
  const selectedCellFormula = computed(() => {
    workerCellSnapshotRevision.value;
    return clipboardDomain.getCellFormula(deferredMetadataCell.value);
  });
  const selectedChartFormula = computed(() => {
    if (!selectedChartElement.value || selectedChartElement.value.kind === "chart" || selectedChartElement.value.seriesIndex < 0) return null;
    return chartImageDomain.getChartSeriesFormula(selectedChartElement.value.chartId, selectedChartElement.value.seriesIndex);
  });
  const selectedFormula = computed(() => selectedChartFormula.value ?? selectedCellFormula.value);

  // ═══════════════════════════════════════════════════════════════════
  // Return composed controller
  // ═══════════════════════════════════════════════════════════════════
  return {
    get activeCell() { return activeCell.value; },
    get activeCellAddress() { return activeCellAddress.value; },
    get activeSheet() { return activeSheet.value; },
    get activeSheetIndex() { return activeSheetIndex.value; },
    get activeTab() { return activeTab.value; },
    get activeTabIndex() { return activeTabIndex.value; },
    addSheet: editingDomain.addSheet,
    get canRedo() { return canRedo.value; },
    get canDownload() { return sourceState.value === "ready" && sourceBufferRef.value !== null; },
    get canExport() {
      return sourceState.value === "ready"
        && !isLoading.value
        && sourceBufferRef.value !== null
        && workbook.value !== null;
    },
    get canLoadDeferred() { return canLoadDeferred.value; },
    get calculationStatus() { return calculationStatus.value; },
    get canUndo() { return canUndo.value; },
    get canZoomIn() { return canZoomIn.value; },
    get canZoomOut() { return canZoomOut.value; },
    get charts() { return chartImageDomain.charts.value; },
    get chartsheets() { return chartsheets.value; },
    clearSelectedChart: chartImageDomain.clearSelectedChart,
    clearSelectedChartElement: chartImageDomain.clearSelectedChartElement,
    clearSelectedCells: editingDomain.clearSelectedCells,
    clearSelectedImage: chartImageDomain.clearSelectedImage,
    clearSelection: editingDomain.clearSelection,
    continueDeferredLoad: navigationDomain.continueDeferredLoad,
    cancelDeferredLoad: navigationDomain.cancelDeferredLoad,
    copySelectionToClipboard: clipboardDomain.copySelectionToClipboard,
    get defaultZoomScale() { return defaultZoomScale.value; },
    get deferredLoadFileSize() { return deferredLoadFileSize.value; },
    defineNamedRange: editingDomain.defineNamedRange,
    get displayFileName() { return displayFileName.value; },
    download: downloadSource,
    exportCsv: historyDomain.exportCsv,
    exportXlsx: historyDomain.exportXlsx,
    get error() { return error.value; },
    fillSelection: editingDomain.fillSelection,
    get formControls() { return chartImageDomain.formControls.value; },
    getFormControlById: chartImageDomain.getFormControlById,
    getFormControlItems: chartImageDomain.getFormControlItems,
    addFormControl: chartImageDomain.addFormControl,
    updateFormControl: chartImageDomain.updateFormControl,
    removeFormControl: chartImageDomain.removeFormControl,
    getChartById: chartImageDomain.getChartById,
    getChartSeriesFormula: chartImageDomain.getChartSeriesFormula,
    getChartsheetById: chartImageDomain.getChartsheetById,
    getImageById: chartImageDomain.getImageById,
    getShapeById: chartImageDomain.getShapeById,
    getSheetCharts: chartImageDomain.getSheetCharts,
    getSheetFormControls: chartImageDomain.getSheetFormControls,
    getSheetImages: chartImageDomain.getSheetImages,
    getSheetShapes: chartImageDomain.getSheetShapes,
    file,
    getClipboardData: clipboardDomain.getClipboardData,
    findCells: (query: string) => {
      const sheet = activeSheet.value
      if (!sheet || !query.trim()) return []
      const q = query.trim().toLowerCase()
      const results: XlsxCellAddress[] = []
      const minRow = sheet.minUsedRow ?? 0
      const maxRow = sheet.maxUsedRow ?? -1
      const minCol = sheet.minUsedCol ?? 0
      const maxCol = sheet.maxUsedCol ?? -1
      for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
          const cell: XlsxCellAddress = { row: r, col: c }
          const value = clipboardDomain.getCellDisplayValue(cell)
          if (typeof value === "string" && value.toLowerCase().includes(q)) {
            results.push(cell)
          }
        }
      }
      return results
    },
    getCellDisplayValue: clipboardDomain.getCellDisplayValue,
    getCellFormula: clipboardDomain.getCellFormula,
    getCellDataValidation,
    getDataValidationListOptions,
    get getCellSnapshotAsync() { return isWorkerBacked.value ? historyDomain.getCellSnapshotAsync : undefined; },
    getActiveWorksheet() {
      if (!workbook.value || !activeSheet.value) return null;
      return workbook.value.getSheet(activeSheet.value.workbookSheetIndex);
    },
    get getRowsBatchAsync() { return isWorkerBacked.value ? historyDomain.getRowsBatchAsync : undefined; },
    get images() { return chartImageDomain.images.value; },
    get isLoadDeferred() { return isLoadDeferred.value; },
    get isLoading() { return isLoading.value; },
    get isActiveSheetProtected() { return isActiveSheetProtected.value; },
    get isEditingLocked() { return isEditingLocked.value; },
    get isChartsLoading() { return isChartsLoading.value; },
    get isWorkerBacked() { return isWorkerBacked.value; },
    get lastExportWarning() { return lastExportWarning.value; },
    get maxConcurrentImageDecodes() { return xlsxRuntime.limits?.maxConcurrentImageDecodes ?? 4; },
    get sourceState() { return sourceState.value; },
    get sourceError() { return sourceError.value; },
    reportImageDecodeError,
    mergeSelection: editingDomain.mergeSelection,
    maxZoomScale: MAX_ZOOM_SCALE,
    minZoomScale: MIN_ZOOM_SCALE,
    moveChartBy: chartImageDomain.moveChartBy,
    moveImageBy: chartImageDomain.moveImageBy,
    moveShapeBy: chartImageDomain.moveShapeBy,
    deleteChart: chartImageDomain.deleteChart,
    deleteImage: chartImageDomain.deleteImage,
    deleteShape: chartImageDomain.deleteShape,
    pasteFromClipboard: clipboardDomain.pasteFromClipboard,
    pasteStructuredClipboardData: clipboardDomain.pasteStructuredClipboardData,
    pasteText: clipboardDomain.pasteText,
    removeActiveSheet: editingDomain.removeActiveSheet,
    get readOnly() { return readOnly.value; },
    get locale() { return documentLocale.value; },
    setDocumentLocale,
    get dataValidationFeedback() { return dataValidationFeedback.value; },
    dismissDataValidationFeedback: dismissValidationFeedback,
    recalculate: historyDomain.recalculate,
    refreshChartData,
    redo: historyDomain.redo,
    resetZoom: navigationDomain.resetZoom,
    get revision() { return revision.value; },
    resizeChartBy: chartImageDomain.resizeChartBy,
    resizeImageBy: chartImageDomain.resizeImageBy,
    resizeShapeBy: chartImageDomain.resizeShapeBy,
    resizeColumn: historyDomain.resizeColumn,
    resizeRow: historyDomain.resizeRow,
    setCellFormula: editingDomain.setCellFormula,
    setCellStyle: editingDomain.setCellStyle,
    setCellValue: editingDomain.setCellValue,
    commitCellInput: editingDomain.commitCellInput,
    setRangeStyle: editingDomain.setRangeStyle,
    setSelectedFormula: editingDomain.setSelectedFormula,
    setZoomScale: navigationDomain.setZoomScale,
    setChartRect: chartImageDomain.setChartRect,
    setChartSeriesFormula: chartImageDomain.setChartSeriesFormula,
    setImageRect: chartImageDomain.setImageRect,
    setShapeRect: chartImageDomain.setShapeRect,
    get selectedChart() { return chartImageDomain.selectedChart.value; },
    get selectedChartElement() { return selectedChartElement.value; },
    get selectedChartFormula() { return selectedChartFormula.value; },
    get selectedChartId() { return selectedChartId.value; },
    get selectedCellFormula() { return selectedCellFormula.value; },
    get selectedFormula() { return selectedFormula.value; },
    get selectedFormulaTarget() { return selectedFormulaTarget.value; },
    get selectedImage() { return chartImageDomain.selectedImage.value; },
    get selectedImageId() { return selectedImageId.value; },
    get selectedShape() { return chartImageDomain.selectedShape.value; },
    get selectedShapeId() { return selectedShapeId.value; },
    get selectedRangeAddress() { return selectedRangeAddress.value; },
    get selectedValue() { return selectedValue.value; },
    selectCell: editingDomain.selectCell,
    selectChart: chartImageDomain.selectChart,
    selectChartElement: chartImageDomain.selectChartElement,
    selectImage: chartImageDomain.selectImage,
    selectShape: chartImageDomain.selectShape,
    clearSelectedShape: chartImageDomain.clearSelectedShape,
    selectRange: editingDomain.selectRange,
    get selection() { return selection.value; },
    setActiveSheetIndex: navigationDomain.setActiveSheetIndex,
    setActiveTabIndex: navigationDomain.setActiveTabIndex,
    setSelectedCellFormula: editingDomain.setSelectedCellFormula,
    setSelectedCellStyle: editingDomain.setSelectedCellStyle,
    setSelectedCellValue: editingDomain.setSelectedCellValue,
    get sheets() { return sheets.value; },
    get shapes() { return chartImageDomain.shapes.value; },
    src,
    get sortState() { return sortState.value; },
    sortTable: historyDomain.sortTable,
    get tabs() { return tabs.value; },
    get tables() { return historyDomain.tables.value; },
    undo: historyDomain.undo,
    unmergeSelection: editingDomain.unmergeSelection,
    updateChart: chartImageDomain.updateChart,
    get workbook() { return workbook.value; },
    zoomIn: navigationDomain.zoomIn,
    zoomOut: navigationDomain.zoomOut,
    get zoomScale() { return zoomScale.value; }
  };
}
