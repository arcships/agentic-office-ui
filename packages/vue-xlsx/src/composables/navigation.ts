import { watch } from "vue";
import {
  clampZoomScale,
  resolveNextZoomScale,
} from "./workbook-state";
import type { XlsxControllerContext } from "./internal";

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
    ctx.continueDeferredWorkbookLoad();
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
