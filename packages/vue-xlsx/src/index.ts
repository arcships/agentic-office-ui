import XlsxViewerComponent from "./components/XlsxViewer.vue"

export type {
  OfficeDocumentRevision,
  OfficeObjectReference,
  OfficeReferenceCandidateChange,
  OfficeReferenceConfirmEvent,
  OfficeReferenceError,
  OfficeReferenceResolveEvent,
  OfficeReferenceSurfaceExposed,
  OfficeReferenceSurfaceProps,
  OfficeRegionDraftChange,
  OfficeSelectionCancelEvent,
  OfficeSelectionMode,
} from "@arcships/office-interaction"

export { useXlsxViewerController } from "./composables"
export type { XlsxGridRenderMode } from "./components/grid-dom"
export { resolveChartAriaLabel } from "./render/chart-shared"
export { localizeXlsxLoadError } from "./errors"
export type { XlsxLoadErrorMessages } from "./errors"
export { useClipboardCopyNotice } from "./composables/useClipboardCopyNotice"
export type {
  ClipboardCopyNotice,
  UseClipboardCopyNoticeOptions,
} from "./composables/useClipboardCopyNotice"
export {
  checkCellValueAgainstDataValidation,
  localizeDataValidationFeedback,
} from "./composables/validation"
export type { XlsxValidationMessages } from "./composables/validation"
export {
  CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE,
  XLSX_RANGE_HIGHLIGHT_DEFAULT_COLOR,
  chartRangeCellAt,
  collectChartRangeHighlights,
  normalizeHighlightRange,
  parseChartSourceA1Range,
  parseChartSourceRange,
} from "./composables/chart-range-highlight"
export type {
  ChartSourceHighlightRole,
  ChartSourceRange,
  ChartSourceSheet,
  CollectChartRangeHighlightsOptions,
  ParseChartSourceRangeOptions,
} from "./composables/chart-range-highlight"
export {
  createRangeHighlightLayout,
  rangeHighlightStyle,
} from "./components/range-highlight-geometry"
export type {
  RangeHighlightLayout,
  RangeHighlightSegment,
  RangeHighlightViewport,
} from "./components/range-highlight-geometry"
export {
  FORMULA_REFERENCE_COLORS,
  collectFormulaReferences,
  formulaReferenceColor,
  formulaReferenceSegments,
  parseFormulaReferences,
  resolveNamedRangeTarget,
} from "./composables/formula-reference-highlight"
export type {
  FormulaReference,
  FormulaReferenceCollection,
  FormulaReferenceKind,
  FormulaReferenceNameTarget,
  FormulaReferenceOptions,
  FormulaReferenceSegment,
} from "./composables/formula-reference-highlight"
export {
  FORMULA_REFERENCE_OWNER_FORMULA_BAR,
  FORMULA_REFERENCE_OWNER_GRID,
  createNamedRangeResolver,
  isReferenceOnActiveSheet,
  useFormulaReferenceEditing,
} from "./composables/useFormulaReferenceEditing"
export type {
  FormulaReferenceEditing,
  FormulaReferenceEditingSession,
} from "./composables/useFormulaReferenceEditing"
export { useXlsxViewerThumbnails } from "./composables/useXlsxViewerThumbnails"
export { useXlsxSelectionStats } from "./composables/selection-stats"
export type {
  XlsxSelectionStatsRequest,
  XlsxSelectionStatsState,
} from "./composables/selection-stats"
export { formatSelectionStatValue } from "./components/selection-status-format"
export { useXlsxSearch } from "./composables/useXlsxSearch"
export type {
  XlsxSearchMatch,
  XlsxSearchOptions,
  XlsxSearchState,
  XlsxSearchStatus,
  XlsxSurfaceSearch,
} from "./composables/useXlsxSearch"
export { XlsxFileSizeLimitExceededError } from "./composables"
export type {
  XlsxDiagnostic,
  XlsxLoadError,
  XlsxLoadErrorCode,
  XlsxSourceKind,
  XlsxSourceState,
  XlsxUrlPolicy,
} from "@arcships/xlsx-core"

// Optional renderers keep their historical root names, but load the heavy
// implementation only when a chart actually mounts.
/**
 * @deprecated Since 0.2.0. Rendering modules are owned by `XlsxViewer`; use
 * the high-level component and controller. Earliest removal: 1.0.0.
 */
export { MemoChartSvg, MemoSurfaceChartComposite } from "./optional/lazy-renderers"
/**
 * @deprecated Since 0.2.0. Rendering types are implementation details of
 * `XlsxViewer`. Kept throughout 0.x; earliest removal is 1.0.0.
 */
export type {
  ChartRendererPalette,
  ChartSvgProps,
  ChartLayout,
  LegendItem,
} from "./render/chart-types"

// Component exports
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxGrid } from "./components/XlsxGrid.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxToolbar } from "./components/XlsxToolbar.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxRibbon } from "./components/XlsxRibbon.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxFormulaBar } from "./components/XlsxFormulaBar.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxSheetTabs } from "./components/XlsxSheetTabs.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxChartOverlay } from "./components/XlsxChartOverlay.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxImageLayer } from "./components/XlsxImageLayer.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
/** @deprecated Internal overlay layer; compose XlsxSheetSurface instead. */
export { default as XlsxDrawingLayer } from "./components/XlsxDrawingLayer.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxSelectionOverlay } from "./components/XlsxSelectionOverlay.vue"
/** Range-highlight overlay used internally by `XlsxViewer` / `XlsxSheetSurface`. */
export { default as XlsxRangeHighlightLayer } from "./components/XlsxRangeHighlightLayer.vue"
/** @deprecated Since 0.2.0. Use `XlsxViewer`. Earliest removal: 1.0.0. */
export { default as XlsxContextMenu } from "./components/XlsxContextMenu.vue"
/**
 * Row/column header menu rendered by `XlsxGrid` on the shared
 * `OfficeDropdownMenu` primitive. Hosts that embed the grid alone can mount it.
 */
export { default as XlsxHeaderMenu } from "./components/XlsxHeaderMenu.vue"
/**
 * AutoFilter value panel rendered by `XlsxGrid` on the shared
 * `OfficeDropdownMenu` primitive. Hosts that embed the grid alone can mount it.
 */
export { default as XlsxAutoFilterMenu } from "./components/XlsxAutoFilterMenu.vue"
/**
 * Authoring dialog for charts, comments, hyperlinks, tables, conditional
 * formats and data validation. `XlsxViewer` renders it from the ribbon and
 * context menu; hosts that embed `XlsxSheetSurface` can mount it directly.
 */
export { default as XlsxAuthoringDialog } from "./components/XlsxAuthoringDialog.vue"
/**
 * Paginated print preview. `XlsxViewer` mounts it from the controller's print
 * state; hosts that embed `XlsxSheetSurface` can mount it directly.
 */
export { default as XlsxPrintPreview } from "./components/XlsxPrintPreview.vue"
/**
 * Selection status bar. `XlsxViewer` renders it under the sheet tabs; hosts
 * that embed `XlsxSheetSurface` can mount it directly.
 */
export { default as XlsxStatusBar } from "./components/XlsxStatusBar.vue"

export {
  XLSX_FILE_ACCEPT,
  XLSX_FILE_EXTENSIONS,
  XLSX_FILE_MIME_TYPES,
  isSupportedXlsxFileName,
  xlsxFileExtension,
} from "./composables/xlsx-file-types"

export const XlsxViewer = XlsxViewerComponent

/**
 * Minimal embeddable XLSX renderer for third-party consumers.
 *
 * Renders the sheet grid + chart/image/drawing/selection overlays + context
 * menu + chart sheet surface. No toolbar, ribbon, formula bar, or sheet tabs
 * — host owns all controls.
 *
 * @example
 * ```vue
 * <XlsxSheetSurface
 *   :controller="ctrl"
 *   :read-only="false"
 *   @cell-double-click="onDoubleClick"
 * />
 * ```
 */
export { default as XlsxSheetSurface } from "./components/XlsxSheetSurface.vue"
