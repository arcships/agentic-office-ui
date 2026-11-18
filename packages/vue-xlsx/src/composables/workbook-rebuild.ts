import type { Workbook } from "@dukelib/sheets-wasm";
import {
  loadWorkbookChartAssets,
  revokeWorkbookImageAssets,
  safeCalculate,
  trackXlsxWorkbookLifetime,
  type WorkbookChartAssets,
  type WorkbookImageAssets
} from "@arcships/xlsx-core";
import type { XlsxSheetData } from "@arcships/xlsx-core";
import { loadWorkbookImageAssets } from "./image-assets";
import { buildSheetList, buildVisibleSheetIndexMap } from "./workbook-state";
import type { XlsxControllerContext } from "./internal";

export type WorkbookRebuildResult = {
  calculation: ReturnType<typeof safeCalculate>;
  chartAssets: WorkbookChartAssets;
  imageAssets: ReturnType<typeof loadWorkbookImageAssets>;
  sheets: XlsxSheetData[];
  workbook: Workbook;
};

/**
 * Parses rewritten workbook bytes back into a live model.
 *
 * Sheet rename and sheet reorder both rewrite the saved package and need the
 * same follow-up: a fresh `Workbook`, recalculated formulas, and the worksheet
 * state, images, shapes, form controls and charts re-read from the rewritten
 * package. Rebuilding the assets (rather than splicing the old per-sheet
 * arrays) is what keeps every `workbookSheetIndex`-keyed array aligned with
 * the new sheet order.
 *
 * @returns `null` when the bytes cannot be parsed; the caller leaves the
 *   current workbook untouched.
 */
export function rebuildWorkbookFromBytes(
  ctx: XlsxControllerContext,
  bytes: Uint8Array
): WorkbookRebuildResult | null {
  const currentWorkbook = ctx.workbook.value;
  if (!currentWorkbook) {
    return null;
  }

  const WorkbookConstructor = currentWorkbook.constructor as typeof Workbook;
  let nextWorkbook: Workbook | null = null;
  let nextImageAssets: WorkbookImageAssets | null = null;
  try {
    nextWorkbook = trackXlsxWorkbookLifetime(WorkbookConstructor.fromBytes(bytes));
    const calculation = safeCalculate(nextWorkbook, {
      reparse: () => trackXlsxWorkbookLifetime(WorkbookConstructor.fromBytes(bytes))
    });
    if (calculation.workbook !== nextWorkbook) {
      try {
        nextWorkbook.free();
      } catch {
        // A trapped WASM workbook may reject cleanup; it is never reused.
      }
      nextWorkbook = calculation.workbook;
    }
    nextImageAssets = loadWorkbookImageAssets(bytes, nextWorkbook);
    const sheets = buildSheetList(
      nextWorkbook,
      nextImageAssets.sheetStatesByWorkbookSheetIndex,
      nextImageAssets.themePalette,
      nextImageAssets.styleById,
      nextImageAssets.namedCellStyleByName,
      nextImageAssets.tableStyleByName,
      ctx.showHiddenSheets
    );
    const chartAssets = loadWorkbookChartAssets(
      nextWorkbook,
      nextImageAssets,
      buildVisibleSheetIndexMap(sheets),
      ctx.showHiddenSheets
    );
    return {
      calculation,
      chartAssets,
      imageAssets: nextImageAssets,
      sheets,
      workbook: nextWorkbook
    };
  } catch {
    revokeWorkbookImageAssets(nextImageAssets);
    if (nextWorkbook) {
      try {
        nextWorkbook.free();
      } catch {
        // Best-effort cleanup for an instance that was never committed.
      }
    }
    return null;
  }
}

/**
 * Commits a rebuilt workbook as one undo step.
 *
 * `originalBytes` is the pre-mutation workbook and doubles as the history
 * snapshot, so the undo entry and the rewrite share a single serialization.
 * The calculation status is published before the workbook swap so a host
 * observing the revision sees why formula values may be cached.
 */
export function commitWorkbookRebuild(
  ctx: XlsxControllerContext,
  rebuild: WorkbookRebuildResult,
  originalBytes: Uint8Array
) {
  ctx.noteCalculationResult({
    calculated: rebuild.calculation.calculated,
    skipReason: rebuild.calculation.skipReason
  });
  if (rebuild.calculation.calculated) {
    ctx.resumeAutoCalculate();
  }

  ctx.recordHistoryBeforeMutation(originalBytes);
  ctx.setImageAssets(rebuild.imageAssets);
  ctx.replaceWorkbook(rebuild.workbook);
  ctx.sheets.value = rebuild.sheets;
  ctx.setChartAssets(rebuild.chartAssets);
  ctx.revision.value = ctx.revision.value + 1;
}
