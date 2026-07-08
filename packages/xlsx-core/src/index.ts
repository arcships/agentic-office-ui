export { initWasm, setWasmSource, canUseConfiguredWasmSourceInWorker, getConfiguredWorkerWasmSource, getSheetsWasmModule } from "./wasm";
export type { XlsxWasmSource, WorkerWasmSource } from "./wasm";

export { safeCalculate, tryRecalculate } from "./safe-calculate";
export type { SafeCalculateSkipReason, SafeCalculateResult, SafeCalculateOptions } from "./safe-calculate";

export { XlsxWorkerClient } from "./worker-client";

export {
  resolveWorkbookColor,
  resolveWorkbookFillColor,
  resolveWorkbookFillStyle,
} from "./colors";

export {
  resolveWorksheetDefaultColumnWidthPixels,
  resolveWorksheetDefaultRowHeightPixels,
  resolveWorksheetHiddenRows,
  resolveWorksheetHiddenCols,
  resolveWorksheetMergeMetadata,
  revokeWorkbookImageAssets,
  parseWorkbookStructureAssets,
  parseWorkbookChartStyleAssets,
  parseWorkbookImageAssets,
  updateWorkbookImageAnchor,
  mergeWorkbookImageAssets,
  emuToPixels,
  pixelsToEmu,
  pxToSheetColumnWidth,
  resolveSheetColumnWidthPixels,
  resolveSheetRowHeightPixels,
  resolveRenderedSheetAxisPixels,
  resolveContentSheetAxisPixels,
  rectToImageAnchor,
  resizeImageRect,
} from "./images";
export type {
  WorkbookImageSheetOrigin,
  WorkbookTableMetadata,
  WorkbookImageAssets,
  WorkbookStructureAssets,
  WorkbookChartStyleAssets,
} from "./images";

export {
  buildChartSeriesFormula,
  parseChartSeriesFormula,
  applyChartSeriesFormula,
  loadWorkbookChartAssets,
  updateWorkbookChartAnchor,
  updateWorkbookChartDefinition,
} from "./charts";
export type {
  ParsedChartSeriesFormula,
  WorkbookChartOrigin,
  WorkbookChartAssets,
} from "./charts";

// Types — all data model types
export type {
  UseXlsxViewerControllerOptions,
  XlsxCellAddress,
  XlsxCellAlignmentInput,
  XlsxCellBorderEdgeInput,
  XlsxCellBorderStyleInput,
  XlsxCellFillStyleInput,
  XlsxCellFontStyleInput,
  XlsxCellGradientStopInput,
  XlsxCellNumberFormatInput,
  XlsxCellProtectionInput,
  XlsxCellRange,
  XlsxCellStyleColorInput,
  XlsxCellStyleContext,
  XlsxCellStyleInput,
  XlsxChart,
  XlsxChartAxis,
  XlsxChartDataLabels,
  XlsxChartElementSelection,
  XlsxChartLegend,
  XlsxChartLoadingRenderProps,
  XlsxChartPointDataLabel,
  XlsxChartPointStyle,
  XlsxChartReference,
  XlsxChartSeries,
  XlsxChartTypeGroup,
  XlsxChartWall,
  XlsxChartsheet,
  XlsxClipboardData,
  XlsxConditionalColorScaleRule,
  XlsxConditionalDataBarRule,
  XlsxConditionalFormatIcon,
  XlsxConditionalFormatRule,
  XlsxConditionalFormatValueObject,
  XlsxConditionalIconSetRule,
  XlsxDataValidation,
  XlsxFileTooLargeRenderProps,
  XlsxFormControl,
  XlsxFormControlKind,
  XlsxFreezePanes,
  XlsxImage,
  XlsxImageAnchor,
  XlsxImageMarker,
  XlsxImageRect,
  XlsxImageRenderProps,
  XlsxImageResizeHandlePosition,
  XlsxImageSelectionRenderProps,
  XlsxResolvedCellStyle,
  XlsxScrollerRenderProps,
  XlsxShape,
  XlsxShapeFill,
  XlsxShapeParagraph,
  XlsxShapeStroke,
  XlsxShapeTextBox,
  XlsxShapeTextRun,
  XlsxSheetData,
  XlsxSheetThumbnail,
  XlsxSheetThumbnailResolution,
  XlsxSheetVisibility,
  XlsxSparkline,
  XlsxTable,
  XlsxTableColumn,
  XlsxTableHeaderMenuRenderProps,
  XlsxTableSortDirection,
  XlsxTableSortState,
  XlsxTableStyleDefinition,
  XlsxTableStyleInfo,
  XlsxThemePalette,
  XlsxWorkbookTab,
  XlsxViewerController,
  XlsxViewerEditing,
  XlsxViewerImages,
  XlsxViewerCharts,
  XlsxViewerSelection,
  XlsxViewerThumbnails,
  XlsxViewerZoom,
  XlsxViewerTables,
} from "./types";

import type { XlsxCellAddress, XlsxCellRange } from "./types";

export function columnLabel(col: number): string {
  let label = "";
  let nextValue = col;

  while (nextValue >= 0) {
    label = String.fromCharCode(65 + (nextValue % 26)) + label;
    nextValue = Math.floor(nextValue / 26) - 1;
  }

  return label;
}

export function rangeToA1(range: XlsxCellRange): string {
  const normalized = {
    start: {
      row: Math.min(range.start.row, range.end.row),
      col: Math.min(range.start.col, range.end.col),
    },
    end: {
      row: Math.max(range.start.row, range.end.row),
      col: Math.max(range.start.col, range.end.col),
    },
  };
  const start = `${columnLabel(normalized.start.col)}${normalized.start.row + 1}`;
  const end = `${columnLabel(normalized.end.col)}${normalized.end.row + 1}`;
  return start === end ? start : `${start}:${end}`;
}
