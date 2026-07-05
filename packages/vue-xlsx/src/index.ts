/**
 * @extend-ai/vue-xlsx
 *
 * Vue 3 components and composables for XLSX viewing and editing.
 * Vue API aligned with the public @extend-ai/react-xlsx package.
 */

// Components
export { default as XlsxViewer } from "./XlsxViewer.vue"
export { default as XlsxViewerProvider } from "./XlsxViewerProvider.vue"
export type { XlsxViewerProps, XlsxViewerProviderProps } from "./types"

// Composables
export {
  useXlsxViewerController,
  useXlsxViewer,
  useXlsxViewerSelection,
  useXlsxViewerZoom,
  useXlsxViewerEditing,
  useXlsxViewerTables,
  useXlsxViewerImages,
  useXlsxViewerCharts,
  useXlsxViewerThumbnails,
  XLSX_VIEWER_KEY,
  XLSX_VIEWER_DARK_KEY,
} from "./composables"

export type {
  XlsxViewerController,
  UseXlsxViewerControllerOptions,
  XlsxViewerSelection,
  XlsxViewerZoom,
  XlsxViewerEditing,
  XlsxViewerTables,
  XlsxViewerImages,
  XlsxViewerCharts,
  XlsxViewerThumbnails,
  XlsxSheetThumbnail,
  UseXlsxViewerThumbnailsOptions,
  XlsxCellAddress,
  XlsxCellRange,
  XlsxSheetData,
  XlsxSheetVisibility,
  XlsxWorkbookTab,
  XlsxTable,
  XlsxTableColumn,
  XlsxTableSortDirection,
  XlsxTableSortState,
  XlsxImage,
  XlsxImageAnchor,
  XlsxImageRect,
  XlsxImageResizeHandlePosition,
  XlsxChart,
  XlsxChartAxis,
  XlsxChartSeries,
  XlsxChartElementSelection,
  XlsxChartsheet,
  XlsxCellStyleInput,
  XlsxCellStyleColorInput,
  XlsxCellFontStyleInput,
  XlsxCellAlignmentInput,
  XlsxCellBorderStyleInput,
  XlsxCellBorderEdgeInput,
  XlsxCellFillStyleInput,
  XlsxCellNumberFormatInput,
  XlsxCellProtectionInput,
  XlsxResolvedCellStyle,
  XlsxClipboardData,
  XlsxFormulaTarget,
  XlsxShape,
} from "./composables"

