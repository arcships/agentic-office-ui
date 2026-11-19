import type { Workbook } from "@dukelib/sheets-wasm";
import { applyWorkbookSheetOrder, type XlsxWorkbookTab } from "@arcships/xlsx-core";
import type { XlsxControllerContext } from "./internal";
import { commitWorkbookRebuild, rebuildWorkbookFromBytes } from "./workbook-rebuild";

/** Number of slots in `Workbook.sheetOrder`, including hidden sheets. */
function resolveWorkbookSlotCount(workbook: Workbook, tabs: XlsxWorkbookTab[]) {
  if (Array.isArray(workbook.sheetOrder) && workbook.sheetOrder.length > 0) {
    return workbook.sheetOrder.length;
  }
  // `buildTabs` falls back to `sheetNames` when the engine reports no order;
  // in that case every tab index is a worksheet index.
  return Math.max(workbook.sheetCount, tabs.length);
}

/**
 * Rebuilds the full workbook slot order after moving one visible tab.
 *
 * Hidden sheets keep their slots (only visible slots are refilled), so a tab
 * move never drags an invisible sheet along with it. `visibleSlots` must be
 * the raw slot of each visible tab in ascending order, which is what
 * `buildTabs` produces.
 */
export function moveVisibleSheetSlot(
  slotCount: number,
  visibleSlots: readonly number[],
  fromIndex: number,
  toIndex: number
): number[] | null {
  const visible = new Set(visibleSlots);
  if (visible.size !== visibleSlots.length) {
    return null;
  }
  for (let index = 0; index < visibleSlots.length; index += 1) {
    const slot = visibleSlots[index];
    if (
      !Number.isInteger(slot)
      || slot < 0
      || slot >= slotCount
      || (index > 0 && slot <= visibleSlots[index - 1])
    ) {
      return null;
    }
  }
  if (fromIndex < 0 || fromIndex >= visibleSlots.length || toIndex < 0 || toIndex >= visibleSlots.length) {
    return null;
  }

  const desiredVisible = visibleSlots.slice();
  const [movedSlot] = desiredVisible.splice(fromIndex, 1);
  desiredVisible.splice(toIndex, 0, movedSlot);
  let cursor = 0;
  return Array.from({ length: slotCount }, (_, slot) => {
    if (!visible.has(slot)) {
      return slot;
    }
    const visibleSlot = desiredVisible[cursor];
    cursor += 1;
    return visibleSlot;
  });
}

/**
 * Worksheet tab reorder.
 *
 * The tab strip interleaves worksheets and chartsheets, so the controller API
 * takes tab indices while `Workbook` uses separate worksheet/chartsheet slots.
 * `XlsxWorkbookTab.index` is the workbook slot order position and is what
 * `applyWorkbookSheetOrder` rewrites.
 */
export function createSheetOrderDomain(ctx: XlsxControllerContext) {
  /**
   * Moves a tab from one slot to another and leaves it active.
   *
   * Same-slot drops, out-of-range indices, read-only viewers and a package
   * that cannot be rewritten are rejected before the workbook or the undo
   * stack changes.
   */
  function moveSheet(fromIndex: number, toIndex: number): boolean {
    const workbook = ctx.workbook.value;
    if (!workbook || ctx.readOnly.value || ctx.isApplyingHistoryRef.value) {
      return false;
    }
    if (!Number.isInteger(fromIndex) || !Number.isInteger(toIndex) || fromIndex === toIndex) {
      return false;
    }
    const tabs = ctx.tabs.value;
    if (fromIndex < 0 || fromIndex >= tabs.length || toIndex < 0 || toIndex >= tabs.length) {
      return false;
    }

    const slotCount = resolveWorkbookSlotCount(workbook, tabs);
    const visibleSlots = tabs.map((tab) => tab.index);
    const nextOrder = moveVisibleSheetSlot(slotCount, visibleSlots, fromIndex, toIndex);
    if (!nextOrder || nextOrder.every((slot, index) => slot === index)) {
      return false;
    }

    const movedTab = tabs[fromIndex];
    const movedOldSlot = visibleSlots[fromIndex];
    const movedNewSlot = nextOrder.indexOf(movedOldSlot);
    const wasActive = ctx.activeTabIndex.value === fromIndex;

    const originalBytes = ctx.createSavedWorkbookBytes(workbook);
    const reorderedBytes = applyWorkbookSheetOrder(originalBytes, nextOrder);
    if (!reorderedBytes) {
      return false;
    }
    const rebuild = rebuildWorkbookFromBytes(ctx, reorderedBytes);
    if (!rebuild) {
      return false;
    }

    // The tab id is not stable across the reparse, but the slot is: a worksheet
    // keeps its name while its `workbookSheetIndex` changes. Prefer the slot and
    // fall back to kind + name for an engine that reports no `sheetOrder`.
    const nextTabs = rebuild.chartAssets.tabs;
    let nextTabIndex = nextTabs.findIndex(
      (tab) => tab.index === movedNewSlot && tab.kind === movedTab.kind
    );
    if (nextTabIndex < 0) {
      nextTabIndex = nextTabs.findIndex(
        (tab) => tab.kind === movedTab.kind && tab.name === movedTab.name
      );
    }

    commitWorkbookRebuild(ctx, rebuild, originalBytes);
    if (nextTabIndex >= 0) {
      ctx.activateTab(nextTabIndex, { preserveSelection: wasActive });
    }
    return true;
  }

  return {
    moveSheet
  };
}
