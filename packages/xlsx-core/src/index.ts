export { initWasm, setWasmSource, canUseConfiguredWasmSourceInWorker, getConfiguredWorkerWasmSource, getSheetsWasmModule } from "./wasm";
export type { XlsxWasmSource, WorkerWasmSource } from "./wasm";
export { bundledXlsxWasmUrl } from "./wasm-asset";

export { safeCalculate, tryRecalculate } from "./safe-calculate";
export type { SafeCalculateSkipReason, SafeCalculateResult, SafeCalculateOptions } from "./safe-calculate";

export { XlsxWorkerClient, createBundledXlsxWorker } from "./worker-client";
export type { XlsxWorkerClientOptions } from "./worker-client";
export { createXlsxRuntime } from "./runtime/xlsx-runtime";
export type {
  XlsxRuntime,
  XlsxRuntimeConfig,
  XlsxRuntimeDiagnostic,
  XlsxRuntimeParseOptions,
} from "./runtime/xlsx-runtime";

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
  XlsxRuntimeLike,
  XlsxUrlPolicy,
  XlsxSourceKind,
  XlsxSourceState,
  XlsxLoadErrorCode,
  XlsxLoadError,
  XlsxDiagnostic,
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
  UseXlsxViewerThumbnailsOptions,
  XlsxViewerZoom,
  XlsxViewerTables,
} from "./types";

export {
  XlsxSourceError,
  loadVerifiedXlsxSource,
  resolveAllowedXlsxUrl,
  toXlsxLoadError,
} from "./runtime/xlsx-url-policy";
export type { ResolvedXlsxSource } from "./runtime/xlsx-url-policy";

import {
  columnLabel as coreColumnLabel,
  rangeToA1 as coreRangeToA1,
} from "./core";
import type { XlsxCellAddress, XlsxCellRange } from "./types";

export function columnLabel(col: number): string {
  return coreColumnLabel(col);
}

export function rangeToA1(range: XlsxCellRange): string {
  return coreRangeToA1(range);
}
