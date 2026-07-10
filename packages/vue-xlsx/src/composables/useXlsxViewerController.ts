import { ref, computed, watch, onUnmounted, shallowRef } from "vue";
import {
  createLatestTaskCoordinator,
  createOfficeTaskSequence,
  sanitizeOfficeUrl,
  type OfficeLoadTask,
  type OfficeSource,
  type OfficeUrlPolicy,
} from "@extend-ai/office-runtime";
import type { Workbook } from "@dukelib/sheets-wasm";
import type {
  UseXlsxViewerControllerOptions,
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
  XlsxLoadError,
  XlsxSourceState,
  XlsxViewerController,
  XlsxWorkbookTab,
  XlsxDiagnostic,
  XlsxRuntime
} from "@extend-ai/xlsx-core";
import {
  XlsxWorkerClient,
  createXlsxRuntime,
  loadWorkbookChartAssets,
  mergeWorkbookImageAssets,
  revokeWorkbookImageAssets,
  type WorkbookChartAssets,
  type WorkbookImageAssets,
  type WorkbookImageSheetOrigin
} from "@extend-ai/xlsx-core";
import {
  DEFAULT_DEFER_LOADING_ABOVE_BYTES,
  DEFAULT_MAX_FILE_SIZE_BYTES,
  DEFAULT_ZOOM_TAB_KEY,
  MAX_ZOOM_SCALE,
  MIN_ZOOM_SCALE,
  isAbortError,
  pushHistoryEntry,
  scheduleLowPriorityTask,
  type CellMutationState,
  type HistoryEntry,
  type RangeCellMutation,
  type SnapshotHistoryEntry
} from "./internal";
import {
  clampZoomScale,
  resolveDefaultZoomScale,
  resolveNextZoomScale,
  buildSheetList,
  buildVisibleSheetIndexMap,
  resolveWorkbookSource,
  parseWorkbookBuffer,
  createHistoryDomain,
  tryRecalculate
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
import { cellAddressToA1, rangeToA1 } from "./selection";
import { toXlsxLoadError } from "@extend-ai/xlsx-core";

type DeferredWorkbookLoad = {
  buffer: ArrayBuffer;
  requestId: number;
  resolvedUrl?: string;
  sourceKind: "file" | "url";
  state: "deferred" | "running" | "settled";
  task: OfficeLoadTask;
};

export function useXlsxViewerController(options: UseXlsxViewerControllerOptions): XlsxViewerController {
  const {
    allowResizeInReadOnly = false,
    createWorker,
    deferLoadingAboveBytes = DEFAULT_DEFER_LOADING_ABOVE_BYTES,
    file,
    fileName,
    maxFileSizeBytes = DEFAULT_MAX_FILE_SIZE_BYTES,
    onDiagnostic,
    readOnly: requestedReadOnly = false,
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
  const xlsxRuntime = (configuredRuntime as XlsxRuntime | undefined) ?? createXlsxRuntime({ createWorker, workerUrl });
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
  const revision = ref(0);
  const selectionAnchorRef = shallowRef<XlsxCellAddress | null>(null);
  const undoStackRef = shallowRef<HistoryEntry[]>([]);
  const redoStackRef = shallowRef<HistoryEntry[]>([]);
  const isApplyingHistoryRef = shallowRef(false);
  const historyRevision = ref(0);
  const shouldAutoCalculate = ref(false);
  const workerCellSnapshotRevision = ref(0);
  const isWorkerBacked = ref(false);
  const sortState = ref<XlsxTableSortState | null>(null);
  const forcedReadOnly = ref(false);
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
  const workerCellSnapshotCacheRef = shallowRef(new Map<string, { displayValue: string; formula: string }>());
  let latestLoadRequestId = 0;
  const xlsxRuntimeId = xlsxRuntime.id;
  const xlsxTaskSequence = createOfficeTaskSequence(xlsxRuntimeId);
  const xlsxLoadCoordinator = createLatestTaskCoordinator(xlsxTaskSequence);
  const disposedWorkbooks = new WeakSet<object>();

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

  // ═══════════════════════════════════════════════════════════════════
  // Computed
  // ═══════════════════════════════════════════════════════════════════
  const displayFileName = computed(() => resolveDisplayFileName(src, fileName));
  const shouldDeferLoading = deferLoadingAboveBytes > 0;
  const readOnly = computed(() => requestedReadOnly || forcedReadOnly.value);
  const canResizeReadOnly = computed(() => requestedReadOnly && allowResizeInReadOnly && !forcedReadOnly.value);
  const workerSupported = useWorker && typeof Worker !== "undefined";
  const canUseWorkerForRequestedReadOnly = requestedReadOnly;

  const activeTab = computed(() => tabs.value[activeTabIndex.value] ?? null);
  const activeSheet = computed(() => {
    const tab = activeTab.value;
    return tab?.kind === "sheet" ? sheets.value[tab.sheetIndex ?? -1] ?? null : null;
  });
  const activeZoomTabKey = computed(() => activeTab.value?.id ?? DEFAULT_ZOOM_TAB_KEY);
  const defaultZoomScale = computed(() => resolveDefaultZoomScale(activeTab.value, activeSheet.value));
  const zoomScale = computed(() => clampZoomScale(zoomScaleOverridesByTabId.value[activeZoomTabKey.value] ?? defaultZoomScale.value));
  const canZoomIn = computed(() => zoomScale.value < MAX_ZOOM_SCALE);
  const canZoomOut = computed(() => zoomScale.value > MIN_ZOOM_SCALE);
  const isLoadDeferred = computed(() => deferredLoadFileSize.value !== null);
  const canLoadDeferred = computed(() => !isLoading.value && isLoadDeferred.value);
  const canUndo = computed(() => {
    historyRevision.value;
    return !readOnly.value && !isApplyingHistoryRef.value && undoStackRef.value.length > 0;
  });
  const canRedo = computed(() => {
    historyRevision.value;
    return !readOnly.value && !isApplyingHistoryRef.value && redoStackRef.value.length > 0;
  });
  const visibleSheetIndexByWorkbookSheetIndex = computed(() => new Map(sheets.value.map((sheet, index) => [sheet.workbookSheetIndex, index])));
  const deferredMetadataCell = computed(() => activeCell.value);
  const deferredMetadataSheet = computed(() => activeSheet.value);

  // ═══════════════════════════════════════════════════════════════════
  // Helpers (need defined before they can go on ctx)
  // ═══════════════════════════════════════════════════════════════════
  function shouldForceReadOnlyForBuffer(bufferByteLength: number) {
    return !requestedReadOnly && readOnlyAboveBytes > 0 && bufferByteLength > readOnlyAboveBytes;
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
  function disposeWorkbook(target: Workbook | null | undefined) {
    if (!target || disposedWorkbooks.has(target)) return;
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
  function setChartAssets(assets: WorkbookChartAssets | null) {
    chartAssetsRef.value = assets;
    chartsByWorkbookSheetIndex.value = assets?.chartsByWorkbookSheetIndex ?? [];
    chartsheets.value = assets?.chartsheets ?? [];
    tabs.value = assets?.tabs ?? [];
    isChartsLoading.value = false;
  }
  function shouldFallbackFromWorkerError(error: unknown) {
    const message = error instanceof Error ? error.message : String(error ?? "");
    return message.includes("DOMParser is not defined") || message.includes("XMLSerializer is not defined") || message.includes("Worker chart payload incomplete");
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
    const quickAssets = loadWorkbookChartAssets(targetWorkbook, null, visibleMap, showHiddenSheets);
    setChartAssets(quickAssets);
    if (effectiveSkipXmlParsing) return;
    const hasCharts = quickAssets.chartsByWorkbookSheetIndex.some((sc) => sc.length > 0);
    if (!hasCharts) { isChartsLoading.value = false; return; }
    isChartsLoading.value = true;
    const requestToken = chartLoadRequestTokenRef.value + 1;
    chartLoadRequestTokenRef.value = requestToken;
    chartDisplayFallbackCleanupRef.value?.();
    chartDisplayFallbackCleanupRef.value = null;
    let fallbackTriggered = false;
    const triggerFallback = () => {
      if (fallbackTriggered || requestToken !== chartLoadRequestTokenRef.value) return;
      fallbackTriggered = true;
      runMainThreadFallback();
    };
    const workerTimeoutHandle = typeof window !== "undefined" ? window.setTimeout(() => triggerFallback(), 1500) : null;
    const applyWorkerResult = (result: { chartsByWorkbookSheetIndex: XlsxChart[][]; chartsheets: XlsxChartsheet[]; tabs: XlsxWorkbookTab[] }) => {
      if (requestToken !== chartLoadRequestTokenRef.value) return;
      chartsByWorkbookSheetIndex.value = result.chartsByWorkbookSheetIndex;
      chartsheets.value = result.chartsheets;
      tabs.value = result.tabs;
      isChartsLoading.value = false;
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
        if (workerTimeoutHandle !== null) window.clearTimeout(workerTimeoutHandle);
        if (fallbackTriggered) return;
        try { if (hasIncompleteWorkerChartSnapshot(result)) { triggerFallback(); return; } applyWorkerResult(result); }
        catch { triggerFallback(); }
      })
      .catch((error: unknown) => {
        if (workerTimeoutHandle !== null) window.clearTimeout(workerTimeoutHandle);
        if (isAbortError(error)) return;
        triggerFallback();
      });
  }
  async function loadWorkbookOnMainThread(buffer: ArrayBuffer) {
    const nextParsedWorkbook = await parseWorkbookBuffer(buffer);
    try {
      const bytes = new Uint8Array(buffer);
      const nextImageAssets = loadWorkbookImageAssets(bytes, nextParsedWorkbook.workbook, shouldSkipXmlParsingForWorkbook(bytes, skipXmlParsing));
      return { imageAssets: nextImageAssets, parsedWorkbook: nextParsedWorkbook };
    } catch (error) {
      disposeWorkbook(nextParsedWorkbook.workbook);
      throw error;
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
  function getActiveWorksheet() {
    if (!workbook.value || !activeSheet.value) return null;
    return workbook.value.getSheet(activeSheet.value.workbookSheetIndex);
  }
  function maybeRecalculateWorkbook(targetWorkbook: Workbook) {
    if (!shouldAutoCalculate.value) return;
    const result = tryRecalculate(targetWorkbook);
    if (!result.calculated) shouldAutoCalculate.value = false;
  }
  function createSavedWorkbookBytes(targetWorkbook: Workbook) {
    return mergeWorkbookImageAssets(sanitizeSavedWorkbookBytes(targetWorkbook.saveXlsxBytes()), imageAssetsRef.value, sheetOriginsRef.value);
  }
  function captureCellMutationState(cell: XlsxCellAddress) {
    const worksheet = getActiveWorksheet();
    if (!worksheet) return null;
    return {
      formula: worksheet.getFormulaAt(cell.row, cell.col) ?? null,
      style: worksheet.getCellStyleAt(cell.row, cell.col),
      value: worksheet.getCellAt(cell.row, cell.col).toJs()
    };
  }
  function recordHistoryBeforeMutation() {
    if (isApplyingHistoryRef.value || !workbook.value) return;
    pushHistoryEntry(undoStackRef.value, {
      kind: "snapshot", activeCell: activeCell.value, activeSheetIndex: activeSheetIndex.value,
      bytes: createSavedWorkbookBytes(workbook.value), selection: selection.value
    });
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
    });
    redoStackRef.value = [];
    historyRevision.value = historyRevision.value + 1;
  }
  function recordRangeEditHistory(mutations: RangeCellMutation[], selectionAfter: XlsxCellRange | null, activeCellAfter: XlsxCellAddress | null) {
    if (!activeSheet.value || isApplyingHistoryRef.value || mutations.length === 0) return;
    pushHistoryEntry(undoStackRef.value, {
      kind: "range-edit", activeCellAfter, activeCellBefore: activeCell.value,
      mutations, selectionAfter, selectionBefore: selection.value,
      sheetIndex: activeSheet.value.workbookSheetIndex
    });
    redoStackRef.value = [];
    historyRevision.value = historyRevision.value + 1;
  }

  function isCurrentWorkbookLoad(load: DeferredWorkbookLoad) {
    return load.task.isCurrent() && isCurrentLoadRequest(load.requestId);
  }

  async function executeWorkbookLoad(load: DeferredWorkbookLoad) {
    const { buffer, requestId, resolvedUrl, sourceKind, task } = load;
    if (!isCurrentWorkbookLoad(load)) return;
    if (maxFileSizeBytes > 0 && buffer.byteLength > maxFileSizeBytes) {
      throw new XlsxFileSizeLimitExceededError(buffer.byteLength, maxFileSizeBytes);
    }
    const preflight = preflightWorkbookBuffer(buffer);
    if (preflight?.tooLarge) throw createWorkbookTooLargeError(preflight);

    const willForceReadOnly = shouldForceReadOnlyForBuffer(buffer.byteLength);
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
        if (!isCurrentWorkbookLoad(load)) return;
        if (!effectiveSkipXmlParsing && hasIncompleteWorkerChartSnapshot(snapshot)) {
          throw new Error("Worker chart payload incomplete");
        }
        replaceWorkbook(null);
        sheets.value = snapshot.sheets;
        chartsByWorkbookSheetIndex.value = snapshot.chartsByWorkbookSheetIndex;
        chartsheets.value = snapshot.chartsheets;
        tabs.value = snapshot.tabs;
        chartAssetsRef.value = null;
        workerTablesByWorkbookSheetIndex.value = snapshot.tablesByWorkbookSheetIndex;
        shouldAutoCalculate.value = false;
        isWorkerBacked.value = true;
        sourceBufferRef.value = buffer.slice(0);
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
          bytes: buffer.byteLength,
        });
        return;
      } catch (workerError) {
        if (!isCurrentWorkbookLoad(load) || isAbortError(workerError)) throw workerError;
        if (!shouldFallbackFromWorkerError(workerError)) throw workerError;
        disposeWorkerClient();
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
      workerTablesByWorkbookSheetIndex.value = [];
      isWorkerBacked.value = false;
      sourceBufferRef.value = buffer.slice(0);
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
        bytes: buffer.byteLength,
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
      bytes: load.buffer.byteLength,
    });
    void executeWorkbookLoad(load)
      .catch((nextError: unknown) => handleWorkbookLoadFailure(load, nextError))
      .finally(() => {
        clearDeferredLoad(load);
        load.task.finish();
      });
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
      sourceBufferRef.value = null;
      activeSheetIndex.value = 0; activeTabIndex.value = 0;
      activeCell.value = null; selection.value = null;
      selectedChartId.value = null; selectedChartElement.value = null; selectedImageId.value = null;
      selectionAnchorRef.value = null;
      undoStackRef.value = []; redoStackRef.value = []; isApplyingHistoryRef.value = false; historyRevision.value = 0;
      shouldAutoCalculate.value = false;
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
      limits: { maxInputBytes: maxFileSizeBytes },
      resources: {
        workerUrl: xlsxRuntime.workerUrl,
        wasmUrl: typeof xlsxRuntime.wasmSource === "string" ? xlsxRuntime.wasmSource : undefined,
      },
      urlPolicy: urlPolicy as OfficeUrlPolicy | undefined,
    });
    isLoading.value = true; error.value = null; sourceState.value = "loading"; sourceError.value = null; clearImageAssets(); clearChartAssets();
    workerTablesByWorkbookSheetIndex.value = []; isWorkerBacked.value = false;
    deferredLoadRef.value = null; deferredBufferRef.value = null; deferredLoadRequestIdRef.value = null; sourceBufferRef.value = null; deferredLoadFileSize.value = null;
    activeSheetIndex.value = 0; activeTabIndex.value = 0;
    activeCell.value = null; selection.value = null;
    selectedChartId.value = null; selectedChartElement.value = null; selectedImageId.value = null;
    selectionAnchorRef.value = null;
    undoStackRef.value = []; redoStackRef.value = []; isApplyingHistoryRef.value = false; historyRevision.value = 0;
    shouldAutoCalculate.value = false;
    workerCellSnapshotCacheRef.value.clear(); workerCellSnapshotRevision.value = 0;
    sortState.value = null; zoomScaleOverridesByTabId.value = {}; revision.value = 0;
    disposeWorkerClient();
    reportDiagnostic({ type: "load-start", requestId, taskId: loadTask.context.taskId, sourceKind, url: src });

    const activeLoad: DeferredWorkbookLoad = {
      buffer: new ArrayBuffer(0),
      requestId,
      resolvedUrl: src,
      sourceKind,
      state: "running",
      task: loadTask,
    };
    let handedToDeferredLoad = false;

    void resolveWorkbookSource({ file, src, urlPolicy }, loadTask.signal)
      .then(async (resolvedSource) => {
        const buffer = resolvedSource.buffer;
        activeLoad.buffer = buffer;
        activeLoad.resolvedUrl = resolvedSource.resolvedUrl;
        activeLoad.sourceKind = resolvedSource.sourceKind;
        if (!isCurrentWorkbookLoad(activeLoad)) return;
        if (maxFileSizeBytes > 0 && buffer.byteLength > maxFileSizeBytes) {
          throw new XlsxFileSizeLimitExceededError(buffer.byteLength, maxFileSizeBytes);
        }
        const preflight = preflightWorkbookBuffer(buffer);
        if (preflight?.tooLarge) throw createWorkbookTooLargeError(preflight);
        if (shouldDeferLoading && buffer.byteLength > deferLoadingAboveBytes) {
          activeLoad.state = "deferred";
          handedToDeferredLoad = true;
          deferredLoadRef.value = activeLoad;
          deferredBufferRef.value = buffer;
          deferredLoadRequestIdRef.value = requestId;
          deferredLoadFileSize.value = buffer.byteLength;
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
            bytes: buffer.byteLength,
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
    workerSupported, shouldDeferLoading, canUseWorkerForRequestedReadOnly, requestedReadOnly,
    isLoading, error, sourceState, sourceError, workbook, sheets, chartsByWorkbookSheetIndex, chartsheets, tabs,
    isChartsLoading, workerTablesByWorkbookSheetIndex, formControlsByWorkbookSheetIndex,
    imagesByWorkbookSheetIndex, shapesByWorkbookSheetIndex, activeSheetIndex, activeTabIndex,
    zoomScaleOverridesByTabId, activeCell, selection, selectedChartId, selectedChartElement,
    selectedImageId, revision, selectionAnchorRef, undoStackRef, redoStackRef,
    isApplyingHistoryRef, historyRevision, shouldAutoCalculate, workerCellSnapshotRevision,
    isWorkerBacked, sortState, forcedReadOnly, deferredBufferRef, deferredLoadRequestIdRef, sourceBufferRef, deferredLoadFileSize,
    imageAssetsRef, chartAssetsRef, chartLoadRequestTokenRef, chartDisplayFallbackCleanupRef,
    sheetOriginsRef, workerClientRef, workerCellSnapshotCacheRef,
    activeTab, activeSheet, deferredMetadataCell, deferredMetadataSheet,
    activeZoomTabKey, defaultZoomScale, zoomScale, canZoomIn, canZoomOut,
    readOnly, canResizeReadOnly, displayFileName, visibleSheetIndexByWorkbookSheetIndex,
    refreshWorkbookState, maybeRecalculateWorkbook, getActiveWorksheet,
    setChartAssets, setImageAssets, clearImageAssets, clearChartAssets,
    getWorkerClient, disposeWorkerClient, replaceWorkbook, disposeWorkbook, discardWorkbookLoadResult,
    isCurrentLoadRequest, getCurrentLoadRequestId, continueDeferredWorkbookLoad,
    startChartDisplayHydration, loadWorkbookOnMainThread,
    hasIncompleteWorkerChartSnapshot, shouldFallbackFromWorkerError,
    ensureChartAssetsHydrated,
    shouldForceReadOnlyForBuffer, shouldUseWorkerForReadOnlyLoad,
    createSavedWorkbookBytes, captureCellMutationState,
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
        .catch(() => {
          if (!isCurrent) return;
          workerCellSnapshotCacheRef.value.set(cacheKey, { displayValue: "", formula: "" });
          workerCellSnapshotRevision.value = workerCellSnapshotRevision.value + 1;
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
  const selectedValue = computed(() => clipboardDomain.getCellDisplayValue(deferredMetadataCell.value));
  const selectedCellFormula = computed(() => clipboardDomain.getCellFormula(deferredMetadataCell.value));
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
    get canExport() { return Boolean(workbook.value); },
    get canLoadDeferred() { return canLoadDeferred.value; },
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
    getChartById: chartImageDomain.getChartById,
    getChartSeriesFormula: chartImageDomain.getChartSeriesFormula,
    getChartsheetById: chartImageDomain.getChartsheetById,
    getImageById: chartImageDomain.getImageById,
    getSheetCharts: chartImageDomain.getSheetCharts,
    getSheetFormControls: chartImageDomain.getSheetFormControls,
    getSheetImages: chartImageDomain.getSheetImages,
    getSheetShapes: chartImageDomain.getSheetShapes,
    file,
    getClipboardData: clipboardDomain.getClipboardData,
    getCellDisplayValue: clipboardDomain.getCellDisplayValue,
    getCellFormula: clipboardDomain.getCellFormula,
    get getCellSnapshotAsync() { return isWorkerBacked.value ? historyDomain.getCellSnapshotAsync : undefined; },
    getActiveWorksheet,
    get getRowsBatchAsync() { return isWorkerBacked.value ? historyDomain.getRowsBatchAsync : undefined; },
    get images() { return chartImageDomain.images.value; },
    get isLoadDeferred() { return isLoadDeferred.value; },
    get isLoading() { return isLoading.value; },
    get isChartsLoading() { return isChartsLoading.value; },
    get isWorkerBacked() { return isWorkerBacked.value; },
    get sourceState() { return sourceState.value; },
    get sourceError() { return sourceError.value; },
    mergeSelection: editingDomain.mergeSelection,
    maxZoomScale: MAX_ZOOM_SCALE,
    minZoomScale: MIN_ZOOM_SCALE,
    moveChartBy: chartImageDomain.moveChartBy,
    moveImageBy: chartImageDomain.moveImageBy,
    pasteFromClipboard: clipboardDomain.pasteFromClipboard,
    pasteStructuredClipboardData: clipboardDomain.pasteStructuredClipboardData,
    pasteText: clipboardDomain.pasteText,
    removeActiveSheet: editingDomain.removeActiveSheet,
    get readOnly() { return readOnly.value; },
    recalculate: historyDomain.recalculate,
    redo: historyDomain.redo,
    resetZoom: navigationDomain.resetZoom,
    get revision() { return revision.value; },
    resizeChartBy: chartImageDomain.resizeChartBy,
    resizeImageBy: chartImageDomain.resizeImageBy,
    resizeColumn: historyDomain.resizeColumn,
    resizeRow: historyDomain.resizeRow,
    setCellFormula: editingDomain.setCellFormula,
    setCellStyle: editingDomain.setCellStyle,
    setCellValue: editingDomain.setCellValue,
    setRangeStyle: editingDomain.setRangeStyle,
    setSelectedFormula: editingDomain.setSelectedFormula,
    setZoomScale: navigationDomain.setZoomScale,
    setChartRect: chartImageDomain.setChartRect,
    setChartSeriesFormula: chartImageDomain.setChartSeriesFormula,
    setImageRect: chartImageDomain.setImageRect,
    get selectedChart() { return chartImageDomain.selectedChart.value; },
    get selectedChartElement() { return selectedChartElement.value; },
    get selectedChartFormula() { return selectedChartFormula.value; },
    get selectedChartId() { return selectedChartId.value; },
    get selectedCellFormula() { return selectedCellFormula.value; },
    get selectedFormula() { return selectedFormula.value; },
    get selectedFormulaTarget() { return selectedFormulaTarget.value; },
    get selectedImage() { return chartImageDomain.selectedImage.value; },
    get selectedImageId() { return selectedImageId.value; },
    get selectedRangeAddress() { return selectedRangeAddress.value; },
    get selectedValue() { return selectedValue.value; },
    selectCell: editingDomain.selectCell,
    selectChart: chartImageDomain.selectChart,
    selectChartElement: chartImageDomain.selectChartElement,
    selectImage: chartImageDomain.selectImage,
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
