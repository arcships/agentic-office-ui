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
  type RangeCellMutation,
  type XlsxControllerContext
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
