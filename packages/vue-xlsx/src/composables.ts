// Barrel — all public composable surface is in composables/.
// Kept as a re-export to preserve the public import path for consumers.
export { useXlsxViewerController } from "./composables/useXlsxViewerController";
export { XlsxFileSizeLimitExceededError } from "./composables/formatting";

// Re-export domain helpers for direct consumers
export { createEditingDomain } from "./composables/editing";
export { createClipboardDomain } from "./composables/clipboard";
export { createChartImageDomain } from "./composables/chart-controller";
export { createNavigationDomain } from "./composables/navigation";
export { createHistoryDomain } from "./composables/workbook-state";

// Re-export utilities
export {
  clampZoomScale,
  resolveDefaultZoomScale,
  resolveNextZoomScale,
  buildSheetList,
  buildVisibleSheetIndexMap,
  resolveWorkbookBuffer,
  resolveWorkbookSource,
  parseWorkbookBuffer,
  tryRecalculate
} from "./composables/workbook-state";

export {
  cellAddressToA1,
  columnLabel,
  normalizeRange,
  parseA1CellReference,
  parseA1RangeReference,
  parseWorksheetDataValidations,
  parseWorksheetFreezePanes,
  rangeContainsCell,
  rangeToA1,
  resolveSheetDisplayUsedRange
} from "./composables/selection";

export {
  createWorkbookTooLargeError,
  decodeHtmlEntities,
  escapeHtml,
  fileStem,
  formatBinaryBytes,
  mapBorder,
  preflightWorkbookBuffer,
  resolveDisplayFileName,
  sanitizeSavedWorkbookBytes,
  type WorkbookPreflightResult,
  type ZipEntryMetadata
} from "./composables/formatting";

export {
  createAbortError,
  isAbortError,
  normalizeCellValue,
  normalizeWorksheetVisibility,
  pushHistoryEntry,
  scheduleLowPriorityTask,
  type CellEditHistoryEntry,
  type CellMutationState,
  type ClipboardPayload,
  type HistoryEntry,
  type RangeCellMutation,
  type RangeEditHistoryEntry,
  type SnapshotHistoryEntry,
  type XlsxControllerContext
} from "./composables/internal";

export {
  downloadArrayBuffer,
  downloadBytes,
  downloadText,
  downloadUrl,
  parseClipboardText,
  INTERNAL_CLIPBOARD_MIME
} from "./composables/clipboard";

export {
  loadWorkbookImageAssets,
  shouldSkipXmlParsingForWorkbook,
  createBasicWorkbookAssets,
  collectWorksheetApiImages,
  collectWorksheetApiShapes,
  collectWorksheetBatchImages,
  mergeParsedAndApiImages
} from "./composables/chart-controller";

export {
  DEFAULT_COL_WIDTH,
  DEFAULT_DEFER_LOADING_ABOVE_BYTES,
  DEFAULT_MAX_FILE_SIZE_BYTES,
  DEFAULT_ROW_HEIGHT,
  DEFAULT_ZOOM_SCALE,
  DEFAULT_ZOOM_TAB_KEY,
  EMU_PER_PIXEL,
  GRID_HEADER_HEIGHT,
  GRID_ROW_HEADER_WIDTH,
  MAX_ZOOM_SCALE,
  MIN_COL_WIDTH_PX,
  MIN_ROW_HEIGHT_PX,
  MIN_ZOOM_SCALE,
  XLSX_MIME_TYPE,
  CSV_MIME_TYPE
} from "./composables/internal";
