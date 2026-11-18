import { computed, type ComputedRef } from "vue";
import {
  applyWorkbookSheetVisibility,
  loadWorkbookChartAssets,
  type XlsxHiddenSheet,
  type XlsxSheetVisibility,
} from "@arcships/xlsx-core";
import type { Workbook } from "@dukelib/sheets-wasm";
import { normalizeWorksheetVisibility, type XlsxControllerContext } from "./internal";
import { buildSheetList, buildVisibleSheetIndexMap } from "./workbook-state";
import { commitWorkbookRebuild, rebuildWorkbookFromBytes } from "./workbook-rebuild";

/**
 * Worksheet visibility.
 *
 * The WASM workbook exposes `Worksheet.visibility` as read-only, so every
 * visibility change rewrites the `state` attribute in the saved package and
 * reparses it — the same save/parse round trip the reorder and rename domains
 * use. That keeps `sheets`, `tabs`, images, charts, shapes, form controls,
 * sheet state, undo snapshots and exported bytes aligned on one source of
 * truth.
 */
export function createSheetVisibilityDomain(ctx: XlsxControllerContext) {
  /** Clears state that belongs to the sheet we are leaving. */
  function clearSheetScopedState() {
    ctx.activeCell.value = null;
    ctx.selection.value = null;
    ctx.selectedChartId.value = null;
    ctx.selectedChartElement.value = null;
    ctx.selectedImageId.value = null;
    ctx.selectedShapeId.value = null;
    ctx.selectionAnchorRef.value = null;
    ctx.sortState.value = null;
  }

  function countVisibleSheets(workbook: Workbook) {
    let count = 0;
    for (let index = 0; index < workbook.sheetCount; index += 1) {
      const worksheet = workbook.getSheet(index);
      try {
        if (normalizeWorksheetVisibility(worksheet.visibility) === "visible") {
          count += 1;
        }
      } finally {
        worksheet.free();
      }
    }
    return count;
  }

  function isVisibleSheetTab(
    tab: { kind: string; workbookSheetIndex?: number }
  ) {
    if (tab.kind !== "sheet" || tab.workbookSheetIndex === undefined) {
      return false;
    }
    const sheet = ctx.sheets.value.find(
      (entry) => entry.workbookSheetIndex === tab.workbookSheetIndex
    );
    return sheet?.visibility === "visible";
  }

  /**
   * Picks the visible worksheet a hiding active tab should fall back to.
   *
   * The next visible tab wins, then the nearest previous one, then the first
   * visible worksheet — the same order Excel uses when the active sheet is
   * hidden. Chartsheet tabs never become the fallback because they do not own
   * a `sheets` entry.
   */
  function resolveNearestVisibleWorkbookIndex(
    excludedWorkbookSheetIndex: number,
    fromTabIndex: number
  ): number | null {
    const candidates = ctx.tabs.value
      .map((tab, tabIndex) => ({ tab, tabIndex }))
      .filter(({ tab }) => (
        isVisibleSheetTab(tab) && tab.workbookSheetIndex !== excludedWorkbookSheetIndex
      ));
    if (candidates.length === 0) {
      return null;
    }
    const after = candidates.find(({ tabIndex }) => tabIndex > fromTabIndex);
    if (after) {
      return after.tab.workbookSheetIndex ?? null;
    }
    const before = [...candidates].reverse().find(({ tabIndex }) => tabIndex < fromTabIndex);
    return before?.tab.workbookSheetIndex ?? candidates[0].tab.workbookSheetIndex ?? null;
  }

  function activateSheetByWorkbookIndex(
    workbookSheetIndex: number | null,
    options?: { preserveSelection?: boolean }
  ) {
    if (workbookSheetIndex === null) {
      return false;
    }
    const nextSheetIndex = ctx.sheets.value.findIndex(
      (sheet) => sheet.workbookSheetIndex === workbookSheetIndex
    );
    if (nextSheetIndex < 0) {
      return false;
    }
    ctx.switchToSheet(nextSheetIndex, options);
    return true;
  }

  function activateChartsheetByName(
    name: string | null,
    options?: { preserveSelection?: boolean }
  ) {
    if (!name) {
      return false;
    }
    const nextTabIndex = ctx.tabs.value.findIndex(
      (tab) => tab.kind === "chartsheet" && tab.name === name
    );
    if (nextTabIndex < 0) {
      return false;
    }
    ctx.activateTab(nextTabIndex, options);
    return true;
  }

  function rebuildSheetModel(includeHidden: boolean) {
    const workbook = ctx.workbook.value;
    if (!workbook) {
      return false;
    }
    const nextSheets = buildSheetList(
      workbook,
      ctx.imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex,
      ctx.imageAssetsRef.value?.themePalette,
      ctx.imageAssetsRef.value?.styleById,
      ctx.imageAssetsRef.value?.namedCellStyleByName,
      ctx.imageAssetsRef.value?.tableStyleByName,
      includeHidden
    );
    const nextChartAssets = loadWorkbookChartAssets(
      workbook,
      ctx.imageAssetsRef.value,
      buildVisibleSheetIndexMap(nextSheets),
      includeHidden
    );
    ctx.sheets.value = nextSheets;
    ctx.setChartAssets(nextChartAssets);
    ctx.revision.value = ctx.revision.value + 1;
    return true;
  }

  /**
   * Hidden worksheets the Unhide menu can offer, in workbook order.
   *
   * Reads the live workbook so the list exists even when hidden tabs are not
   * rendered. A Worker-backed viewer has no workbook; its snapshot carries
   * hidden entries only when `showHiddenSheets` was enabled, which is still
   * enough to list them.
   */
  const hiddenSheets: ComputedRef<XlsxHiddenSheet[]> = computed(() => {
    const workbook = ctx.workbook.value;
    const sheets: XlsxHiddenSheet[] = [];
    if (!workbook) {
      for (const sheet of ctx.sheets.value) {
        if (sheet.visibility === "hidden") {
          sheets.push({
            name: sheet.name,
            visibility: "hidden",
            workbookSheetIndex: sheet.workbookSheetIndex
          });
        }
      }
      return sheets;
    }
    for (let index = 0; index < workbook.sheetCount; index += 1) {
      const worksheet = workbook.getSheet(index);
      try {
        if (normalizeWorksheetVisibility(worksheet.visibility) === "hidden") {
          sheets.push({
            name: worksheet.name,
            visibility: "hidden",
            workbookSheetIndex: index
          });
        }
      } finally {
        worksheet.free();
      }
    }
    return sheets;
  });

  function getSheetVisibility(workbookSheetIndex: number): XlsxSheetVisibility | null {
    if (!Number.isInteger(workbookSheetIndex) || workbookSheetIndex < 0) {
      return null;
    }
    const workbook = ctx.workbook.value;
    if (workbook) {
      if (workbookSheetIndex >= workbook.sheetCount) {
        return null;
      }
      const worksheet = workbook.getSheet(workbookSheetIndex);
      try {
        return normalizeWorksheetVisibility(worksheet.visibility);
      } finally {
        worksheet.free();
      }
    }
    return ctx.sheets.value.find(
      (sheet) => sheet.workbookSheetIndex === workbookSheetIndex
    )?.visibility ?? null;
  }

  /**
   * Hides, unhides or very-hides one worksheet as a single undo entry.
   *
   * Refuses to remove the last visible worksheet, matching `removeActiveSheet`.
   * Hiding the active worksheet falls back to its nearest visible neighbour;
   * unhiding a worksheet activates it. The package is rewritten and reparsed
   * so every `workbookSheetIndex`-keyed asset and the exported file follow.
   */
  function setSheetVisibility(
    workbookSheetIndex: number,
    visibility: XlsxSheetVisibility
  ): boolean {
    const workbook = ctx.workbook.value;
    if (
      !workbook
      || ctx.readOnly.value
      || ctx.isApplyingHistoryRef.value
      || (visibility !== "visible" && visibility !== "hidden" && visibility !== "veryHidden")
      || !Number.isInteger(workbookSheetIndex)
      || workbookSheetIndex < 0
      || workbookSheetIndex >= workbook.sheetCount
    ) {
      return false;
    }

    const worksheet = workbook.getSheet(workbookSheetIndex);
    let sheetName: string;
    let currentVisibility: XlsxSheetVisibility;
    try {
      sheetName = worksheet.name;
      currentVisibility = normalizeWorksheetVisibility(worksheet.visibility);
    } finally {
      worksheet.free();
    }
    if (currentVisibility === visibility) {
      return false;
    }
    if (
      visibility !== "visible"
      && currentVisibility === "visible"
      && countVisibleSheets(workbook) <= 1
    ) {
      return false;
    }

    const wasActiveSheet = ctx.activeSheet.value?.workbookSheetIndex === workbookSheetIndex;
    const activeChartsheetName = ctx.activeTab.value?.kind === "chartsheet"
      ? ctx.activeTab.value.name
      : null;
    let nextActiveWorkbookIndex = ctx.activeSheet.value?.workbookSheetIndex ?? null;
    if (visibility === "visible") {
      // Unhiding makes the restored worksheet the one the user wants to see.
      nextActiveWorkbookIndex = workbookSheetIndex;
    } else if (wasActiveSheet) {
      nextActiveWorkbookIndex = resolveNearestVisibleWorkbookIndex(
        workbookSheetIndex,
        ctx.activeTabIndex.value
      );
    }

    const originalBytes = ctx.createSavedWorkbookBytes(workbook);
    const rewrittenBytes = applyWorkbookSheetVisibility(
      originalBytes,
      sheetName,
      visibility
    );
    if (!rewrittenBytes || rewrittenBytes === originalBytes) {
      return false;
    }
    const rebuild = rebuildWorkbookFromBytes(ctx, rewrittenBytes);
    if (!rebuild) {
      return false;
    }

    commitWorkbookRebuild(ctx, rebuild, originalBytes);

    const activateTarget = visibility === "visible";
    const preserveSelection = !activateTarget && !wasActiveSheet;
    if (wasActiveSheet && !activateTarget) {
      clearSheetScopedState();
    }
    if (!activateTarget && activeChartsheetName) {
      activateChartsheetByName(activeChartsheetName, { preserveSelection });
    } else if (
      !activateSheetByWorkbookIndex(nextActiveWorkbookIndex, { preserveSelection })
      && ctx.sheets.value.length > 0
    ) {
      activateSheetByWorkbookIndex(ctx.sheets.value[0].workbookSheetIndex);
    }
    return true;
  }

  /**
   * Includes or removes hidden tabs without touching the workbook bytes.
   *
   * No history entry is recorded: this is a view mode. The tab that was active
   * stays active when it is still listed; if it was a hidden tab being removed,
   * the nearest visible worksheet takes over.
   */
  function setShowHiddenSheets(includeHidden: boolean): boolean {
    const nextMode = Boolean(includeHidden);
    if (ctx.showHiddenSheets === nextMode) {
      return true;
    }
    const workbook = ctx.workbook.value;
    if (!workbook) {
      return false;
    }

    const activeSheetData = ctx.activeSheet.value;
    const activeWorkbookIndex = activeSheetData?.workbookSheetIndex ?? null;
    const activeChartsheetName = ctx.activeTab.value?.kind === "chartsheet"
      ? ctx.activeTab.value.name
      : null;
    const activeBecomesUnlisted = !nextMode
      && activeSheetData !== null
      && activeSheetData.visibility !== "visible";
    const nextActiveWorkbookIndex = activeBecomesUnlisted && activeWorkbookIndex !== null
      ? resolveNearestVisibleWorkbookIndex(activeWorkbookIndex, ctx.activeTabIndex.value)
      : activeWorkbookIndex;

    ctx.setShowHiddenSheetsMode(nextMode);
    if (!rebuildSheetModel(nextMode)) {
      ctx.setShowHiddenSheetsMode(!nextMode);
      return false;
    }

    if (activeBecomesUnlisted) {
      clearSheetScopedState();
    }
    const preserveSelection = !activeBecomesUnlisted;
    if (activeChartsheetName) {
      activateChartsheetByName(activeChartsheetName, { preserveSelection });
    } else if (
      !activateSheetByWorkbookIndex(nextActiveWorkbookIndex, { preserveSelection })
      && ctx.sheets.value.length > 0
    ) {
      activateSheetByWorkbookIndex(ctx.sheets.value[0].workbookSheetIndex);
    }
    return true;
  }

  return {
    hiddenSheets,
    getSheetVisibility,
    setSheetVisibility,
    setShowHiddenSheets
  };
}
