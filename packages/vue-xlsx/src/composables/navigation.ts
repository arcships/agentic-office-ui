import { computed, watch, type ComputedRef, type Ref, type ShallowRef } from "vue";
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
  UseXlsxViewerControllerOptions
} from "@extend-ai/xlsx-core";
import {
  loadWorkbookChartAssets,
  revokeWorkbookImageAssets,
  XlsxWorkerClient,
  type WorkbookChartAssets,
  type WorkbookImageAssets,
  type WorkbookImageSheetOrigin
} from "@extend-ai/xlsx-core";
import {
  buildSheetList,
  buildVisibleSheetIndexMap,
  clampZoomScale,
  resolveDefaultZoomScale,
  resolveNextZoomScale,
  resolveWorkbookBuffer,
  parseWorkbookBuffer,
  DEFAULT_ZOOM_TAB_KEY,
  MAX_ZOOM_SCALE,
  MIN_ZOOM_SCALE,
  tryRecalculate
} from "./workbook-state";
import {
  DEFAULT_DEFER_LOADING_ABOVE_BYTES,
  DEFAULT_MAX_FILE_SIZE_BYTES,
  isAbortError,
  scheduleLowPriorityTask,
  type CellMutationState,
  type HistoryEntry,
  type RangeCellMutation
} from "./internal";
import {
  XlsxFileSizeLimitExceededError,
  createWorkbookTooLargeError,
  preflightWorkbookBuffer
} from "./formatting";
import {
  loadWorkbookImageAssets,
  shouldSkipXmlParsingForWorkbook
} from "./chart-controller";

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

export function createNavigationDomain(ctx: XlsxControllerContext) {
  watch(() => ctx.activeTabIndex.value, () => {
    ctx.activeCell.value = null;
    ctx.selection.value = null;
    ctx.selectedChartId.value = null;
    ctx.selectedChartElement.value = null;
    ctx.selectedImageId.value = null;
    ctx.selectionAnchorRef.value = null;
    ctx.sortState.value = null;
  });

  function setActiveSheetIndex(index: number) {
    ctx.activeSheetIndex.value = (() => {
      if (index < 0 || index >= ctx.sheets.value.length) {
        return ctx.activeSheetIndex.value;
      }
      const targetSheet = ctx.sheets.value[index];
      const tabIndex = ctx.tabs.value.findIndex((tab) => tab.kind === "sheet" && tab.workbookSheetIndex === targetSheet?.workbookSheetIndex);
      if (tabIndex >= 0) {
        ctx.activeTabIndex.value = tabIndex;
      }
      return index;
    })();
  }

  function setActiveTabIndex(index: number) {
    ctx.activeTabIndex.value = (() => {
      if (index < 0 || index >= ctx.tabs.value.length) {
        return ctx.activeTabIndex.value;
      }

      const targetTab = ctx.tabs.value[index];
      if (targetTab?.kind === "sheet" && typeof targetTab.sheetIndex === "number") {
        ctx.activeSheetIndex.value = targetTab.sheetIndex;
      }
      return index;
    })();
  }

  function setZoomScale(nextZoomScale: number) {
    const normalizedZoomScale = clampZoomScale(nextZoomScale);
    ctx.zoomScaleOverridesByTabId.value = (() => {
      if (ctx.zoomScaleOverridesByTabId.value[ctx.activeZoomTabKey.value] === normalizedZoomScale) {
        return ctx.zoomScaleOverridesByTabId.value;
      }

      return {
        ...ctx.zoomScaleOverridesByTabId.value,
        [ctx.activeZoomTabKey.value]: normalizedZoomScale
      };
    })();
  }

  function resetZoom() {
    ctx.zoomScaleOverridesByTabId.value = (() => {
      if (ctx.zoomScaleOverridesByTabId.value[ctx.activeZoomTabKey.value] === undefined) {
        return ctx.zoomScaleOverridesByTabId.value;
      }

      const next = { ...ctx.zoomScaleOverridesByTabId.value };
      delete next[ctx.activeZoomTabKey.value];
      return next;
    })();
  }

  function zoomIn() {
    setZoomScale(resolveNextZoomScale(ctx.zoomScale.value, 1));
  }

  function zoomOut() {
    setZoomScale(resolveNextZoomScale(ctx.zoomScale.value, -1));
  }

  watch(() => ctx.tabs.value.length, () => {
    ctx.activeTabIndex.value = (() => {
      if (ctx.tabs.value.length === 0) {
        return 0;
      }
      return Math.min(ctx.activeTabIndex.value, ctx.tabs.value.length - 1);
    })();
  });

  function continueDeferredLoad() {
    const deferredBuffer = ctx.deferredBufferRef.value;
    if (!deferredBuffer) {
      return;
    }

    ctx.isLoading.value = true;
    ctx.error.value = null;

    if (ctx.maxFileSizeBytes > 0 && deferredBuffer.byteLength > ctx.maxFileSizeBytes) {
      ctx.deferredBufferRef.value = null;
      ctx.deferredLoadFileSize.value = null;
      ctx.workbook.value = null;
      ctx.sheets.value = [];
      ctx.clearChartAssets();
      ctx.workerTablesByWorkbookSheetIndex.value = [];
      ctx.clearImageAssets();
      ctx.shouldAutoCalculate.value = false;
      ctx.isWorkerBacked.value = false;
      ctx.sortState.value = null;
      ctx.error.value = new XlsxFileSizeLimitExceededError(deferredBuffer.byteLength, ctx.maxFileSizeBytes);
      ctx.isLoading.value = false;
      return;
    }

    const preflight = preflightWorkbookBuffer(deferredBuffer);
    if (preflight?.tooLarge) {
      ctx.deferredBufferRef.value = null;
      ctx.deferredLoadFileSize.value = null;
      ctx.workbook.value = null;
      ctx.sheets.value = [];
      ctx.clearChartAssets();
      ctx.workerTablesByWorkbookSheetIndex.value = [];
      ctx.shouldAutoCalculate.value = false;
      ctx.isWorkerBacked.value = false;
      ctx.sortState.value = null;
      ctx.error.value = createWorkbookTooLargeError(preflight);
      ctx.isLoading.value = false;
      return;
    }

    const shouldForceReadOnly = ctx.shouldForceReadOnlyForBuffer(deferredBuffer.byteLength);
    ctx.forcedReadOnly.value = shouldForceReadOnly;
    const shouldUseWorkerForLoad = ctx.shouldUseWorkerForReadOnlyLoad(shouldForceReadOnly);
    const effectiveSkipXmlParsing = shouldSkipXmlParsingForWorkbook(new Uint8Array(deferredBuffer), ctx.skipXmlParsing);

    if (shouldUseWorkerForLoad) {
      void ctx.getWorkerClient().loadWorkbook(deferredBuffer, effectiveSkipXmlParsing, ctx.showHiddenSheets)
        .then((snapshot) => {
          if (!effectiveSkipXmlParsing && ctx.hasIncompleteWorkerChartSnapshot(snapshot)) {
            throw new Error("Worker chart payload incomplete");
          }
          ctx.deferredBufferRef.value = null;
          ctx.deferredLoadFileSize.value = null;
          ctx.workbook.value = null;
          ctx.sheets.value = snapshot.sheets;
          ctx.chartsByWorkbookSheetIndex.value = snapshot.chartsByWorkbookSheetIndex;
          ctx.chartsheets.value = snapshot.chartsheets;
          ctx.tabs.value = snapshot.tabs;
          ctx.chartAssetsRef.value = null;
          ctx.workerTablesByWorkbookSheetIndex.value = snapshot.tablesByWorkbookSheetIndex;
          ctx.shouldAutoCalculate.value = false;
          ctx.isWorkerBacked.value = true;
          ctx.sortState.value = null;
          ctx.isChartsLoading.value = false;
          ctx.isLoading.value = false;
        })
        .catch(async (workerError: unknown) => {
          if (isAbortError(workerError)) {
            return;
          }
          if (!ctx.shouldFallbackFromWorkerError(workerError)) {
            throw workerError;
          }

          ctx.disposeWorkerClient();
          const { imageAssets: nextImageAssets, parsedWorkbook: nextParsedWorkbook } = await ctx.loadWorkbookOnMainThread(deferredBuffer);
          ctx.deferredBufferRef.value = null;
          ctx.deferredLoadFileSize.value = null;
          ctx.setImageAssets(nextImageAssets);
          ctx.workbook.value = nextParsedWorkbook.workbook;
          const nextSheets = buildSheetList(
            nextParsedWorkbook.workbook,
            nextImageAssets.sheetStatesByWorkbookSheetIndex,
            nextImageAssets.themePalette,
            nextImageAssets.styleById,
            nextImageAssets.namedCellStyleByName,
            nextImageAssets.tableStyleByName,
            ctx.showHiddenSheets
          );
          ctx.sheets.value = nextSheets;
          ctx.startChartDisplayHydration(deferredBuffer, nextParsedWorkbook.workbook, nextSheets);
          ctx.shouldAutoCalculate.value = nextParsedWorkbook.shouldAutoCalculate;
          ctx.workerTablesByWorkbookSheetIndex.value = [];
          ctx.isWorkerBacked.value = false;
          ctx.sortState.value = null;
          ctx.isLoading.value = false;
        })
        .catch((nextError: unknown) => {
          ctx.deferredBufferRef.value = null;
          ctx.deferredLoadFileSize.value = null;
          ctx.workbook.value = null;
          ctx.sheets.value = [];
          ctx.clearChartAssets();
          ctx.workerTablesByWorkbookSheetIndex.value = [];
          ctx.clearImageAssets();
          ctx.shouldAutoCalculate.value = false;
          ctx.isWorkerBacked.value = false;
          ctx.sortState.value = null;
          ctx.error.value = nextError instanceof Error ? nextError : new Error("Could not load workbook.");
          ctx.isLoading.value = false;
        });
      return;
    }

    void parseWorkbookBuffer(deferredBuffer)
      .then((nextParsedWorkbook) => {
        const bytes = new Uint8Array(deferredBuffer);
        const nextImageAssets = loadWorkbookImageAssets(
          bytes,
          nextParsedWorkbook.workbook,
          shouldSkipXmlParsingForWorkbook(bytes, ctx.skipXmlParsing)
        );
        ctx.deferredBufferRef.value = null;
        ctx.deferredLoadFileSize.value = null;
        ctx.setImageAssets(nextImageAssets);
        ctx.workbook.value = nextParsedWorkbook.workbook;
        const nextSheets = buildSheetList(
          nextParsedWorkbook.workbook,
          nextImageAssets.sheetStatesByWorkbookSheetIndex,
          nextImageAssets.themePalette,
          nextImageAssets.styleById,
          nextImageAssets.namedCellStyleByName,
          nextImageAssets.tableStyleByName,
          ctx.showHiddenSheets
        );
        ctx.sheets.value = nextSheets;
        ctx.startChartDisplayHydration(deferredBuffer, nextParsedWorkbook.workbook, nextSheets);
        ctx.shouldAutoCalculate.value = nextParsedWorkbook.shouldAutoCalculate;
        ctx.workerTablesByWorkbookSheetIndex.value = [];
        ctx.isWorkerBacked.value = false;
        ctx.sortState.value = null;
        ctx.isLoading.value = false;
      })
      .catch((nextError: unknown) => {
        ctx.deferredBufferRef.value = null;
        ctx.deferredLoadFileSize.value = null;
        ctx.workbook.value = null;
        ctx.sheets.value = [];
        ctx.clearChartAssets();
        ctx.workerTablesByWorkbookSheetIndex.value = [];
        ctx.clearImageAssets();
        ctx.shouldAutoCalculate.value = false;
        ctx.isWorkerBacked.value = false;
        ctx.sortState.value = null;
        ctx.error.value = nextError instanceof Error ? nextError : new Error("Could not load workbook.");
        ctx.isLoading.value = false;
      });
  }

  return {
    setActiveSheetIndex,
    setActiveTabIndex,
    setZoomScale,
    resetZoom,
    zoomIn,
    zoomOut,
    continueDeferredLoad
  };
}

export {
  buildSheetList,
  buildVisibleSheetIndexMap,
  clampZoomScale,
  resolveDefaultZoomScale,
  resolveNextZoomScale,
  resolveWorkbookBuffer,
  parseWorkbookBuffer,
  tryRecalculate,
  scheduleLowPriorityTask,
  DEFAULT_DEFER_LOADING_ABOVE_BYTES,
  DEFAULT_MAX_FILE_SIZE_BYTES,
  DEFAULT_ZOOM_TAB_KEY,
  MAX_ZOOM_SCALE,
  MIN_ZOOM_SCALE
};

export type { UseXlsxViewerControllerOptions };
