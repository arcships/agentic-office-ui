/**
 * @extend-ai/vue-xlsx — Vue 3 composables & components for XLSX
 *
 * Vue API aligned with the public @extend-ai/react-xlsx hooks.
 * Built on @extend-ai/xlsx-core (framework-agnostic engine).
 */

import {
  ref, computed, inject,
} from "vue"

// ============================================================
// Symbol-based provide / inject key
// ============================================================

export const XLSX_VIEWER_KEY: unique symbol = Symbol("xlsx-viewer")
export const XLSX_VIEWER_DARK_KEY: unique symbol = Symbol("xlsx-viewer-dark")

// ============================================================
// Core types aligned with the public @extend-ai/react-xlsx API surface
// ============================================================

export interface XlsxCellAddress {
  col: number
  row: number
}

export interface XlsxCellRange {
  end: XlsxCellAddress
  start: XlsxCellAddress
}

export type XlsxSheetVisibility = "hidden" | "veryHidden" | "visible"

export interface XlsxFreezePanes {
  col: number
  row: number
}

export interface XlsxThemePalette {
  colorsByIndex: Record<number, string>
  majorLatinFont?: string
  minorLatinFont?: string
}

export interface XlsxResolvedCellStyle {
  [key: string]: unknown
  alignment?: Record<string, unknown>
  border?: Record<string, Record<string, unknown>>
  cellControl?: { kind: "checkbox" }
  fill?: Record<string, unknown>
  font?: Record<string, unknown>
}

export interface XlsxTableStyleDefinition {
  [elementType: string]: XlsxResolvedCellStyle
}

export interface XlsxConditionalFormatRule {
  kind: string
  priority: number
  ranges: XlsxCellRange[]
  [key: string]: unknown
}

export interface XlsxDataValidation {
  allowBlank?: boolean
  errorMessage?: string
  errorStyle?: string
  inputMessage?: string
  ranges: XlsxCellRange[]
  showDropdown?: boolean
  validationType: string
  [key: string]: unknown
}

export interface XlsxSparkline {
  [key: string]: unknown
}

export interface XlsxSheetData {
  cachedFormulaValues: Record<string, string>
  colWidthOverridesPx: Record<number, number>
  colStyleIds: Record<number, number>
  hiddenCols?: number[]
  hiddenRows?: number[]
  conditionalFormatRules: XlsxConditionalFormatRule[]
  dataValidations: XlsxDataValidation[]
  name: string
  visibility: XlsxSheetVisibility
  columnWidthCharacterWidthPx?: number
  defaultColWidthPx: number
  defaultRowHeightPx: number
  freezePanes: XlsxFreezePanes | null
  hasHorizontalMerges: boolean
  hasVerticalMerges: boolean
  maxHorizontalMergeEndCol: number
  maxVerticalMergeEndRow: number
  minUsedCol: number
  minUsedRow: number
  maxUsedCol: number
  maxUsedRow: number
  rowCount: number
  colCount: number
  rowHeightOverridesPx: Record<number, number>
  rowStyleIds: Record<number, number>
  namedCellStyleByName: Record<string, XlsxResolvedCellStyle>
  styleById: Record<number, XlsxResolvedCellStyle>
  tableStyleByName: Record<string, XlsxTableStyleDefinition>
  visibleRows: number[]
  visibleCols: number[]
  colWidths: number[]
  rowHeights: number[]
  showGridLines: boolean
  sparklines: XlsxSparkline[]
  themePalette: XlsxThemePalette
  workbookSheetIndex: number
  zoomScale?: number
}

export interface XlsxWorkbookTab {
  chartsheetIndex?: number
  id: string
  index: number
  kind: "chartsheet" | "sheet"
  name: string
  sheetIndex?: number
  visibility?: XlsxSheetVisibility
  workbookSheetIndex?: number
}

export interface XlsxTableColumn {
  [key: string]: unknown
  name?: string
  index?: number
}

export type XlsxTableSortDirection = "ascending" | "descending"

export interface XlsxTableSortState {
  columnIndex: number
  direction: XlsxTableSortDirection
  tableName: string
}

export interface XlsxTable {
  columns: XlsxTableColumn[]
  displayName: string
  end: XlsxCellAddress
  headerRowCount: number
  headerRowCellStyle?: string
  name: string
  reference: string
  start: XlsxCellAddress
  styleInfo?: Record<string, unknown>
  totalsRowCount: number
  totalsRowShown: boolean
}

export interface XlsxImageMarker {
  col: number
  colOffsetEmu: number
  row: number
  rowOffsetEmu: number
}

export type XlsxImageAnchor =
  | { from: XlsxImageMarker; kind: "one-cell"; sizeEmu: { cx: number; cy: number } }
  | { kind: "absolute"; positionEmu: { x: number; y: number }; sizeEmu: { cx: number; cy: number } }
  | { from: XlsxImageMarker; kind: "two-cell"; to: XlsxImageMarker }

export interface XlsxImage {
  anchor: XlsxImageAnchor
  description?: string
  editable?: boolean
  hyperlink?: string
  id: string
  mediaPath?: string
  mimeType: string
  name?: string
  sheetIndex: number
  src: string
  workbookSheetIndex: number
  zIndex: number
}

export interface XlsxImageRect {
  height: number
  left: number
  top: number
  width: number
}

export type XlsxImageResizeHandlePosition = "n" | "s" | "e" | "w" | "ne" | "nw" | "se" | "sw"

export interface XlsxChartAxis {
  [key: string]: unknown
}

export interface XlsxChartSeries {
  [key: string]: unknown
}

export interface XlsxChartLegend {
  [key: string]: unknown
}

export interface XlsxChart {
  anchor: XlsxImageAnchor
  autoTitleDeleted?: boolean
  axes: XlsxChartAxis[]
  axisLabelColor?: string
  axisLineColor?: string
  categoryAxis?: XlsxChartAxis | null
  chartAreaBorderColor?: string
  chartAreaFillColor?: string
  chartColorPalette?: string[]
  chartColorPaletteOffset?: number
  chartPath?: string
  chartStyleId?: number
  chartType: string
  dataLabels?: Record<string, unknown> | null
  editable?: boolean
  fontFamily?: string
  gapWidth?: number
  id: string
  is3d?: boolean
  legend?: XlsxChartLegend | null
  name?: string
  overlap?: number
  roundedCorners?: boolean
  series: XlsxChartSeries[]
  sheetIndex: number
  textColor?: string
  title?: string
  titleColor?: string
  titleFontFamily?: string
  valueAxis?: XlsxChartAxis | null
  varyColors?: boolean
  wireframe?: boolean
  workbookSheetIndex: number
  zIndex: number
}

export interface XlsxChartsheet {
  chartIds: string[]
  chartPath?: string
  id: string
  index: number
  name: string
  raw?: Record<string, unknown>
  workbookSheetIndex?: number
}

export interface XlsxChartElementSelection {
  [key: string]: unknown
}

export interface XlsxClipboardData {
  [key: string]: unknown
}

export interface XlsxFormulaTarget {
  [key: string]: unknown
}

export interface XlsxCellStyleColorInput {
  colorType?: "auto" | "rgb" | "argb" | "theme" | "indexed"
  hex?: string
  r?: number
  g?: number
  b?: number
  a?: number
  themeIndex?: number
  tint?: number
  paletteIndex?: number
}

export interface XlsxCellFontStyleInput {
  name?: string
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: "none" | "single" | "double" | "singleAccounting" | "doubleAccounting"
  strikethrough?: boolean
  color?: XlsxCellStyleColorInput
  verticalAlign?: "baseline" | "superscript" | "subscript"
  family?: number
  charset?: number
  scheme?: string
}

export interface XlsxCellAlignmentInput {
  horizontal?: "center" | "fill" | "general" | "justify" | "left" | "right"
  vertical?: "bottom" | "center" | "top"
  wrapText?: boolean
  textRotation?: number
  indent?: number
}

export interface XlsxCellBorderEdgeInput {
  style?: "none" | "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | "hair" | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot"
  color?: XlsxCellStyleColorInput
}

export interface XlsxCellBorderStyleInput {
  top?: XlsxCellBorderEdgeInput
  bottom?: XlsxCellBorderEdgeInput
  left?: XlsxCellBorderEdgeInput
  right?: XlsxCellBorderEdgeInput
  diagonal?: XlsxCellBorderEdgeInput
  diagonalDown?: boolean
  diagonalUp?: boolean
}

export interface XlsxCellGradientStopInput {
  position: number
  color: XlsxCellStyleColorInput
}

export interface XlsxCellFillStyleInput {
  fillType?: "none" | "solid" | "gradient" | "pattern"
  color?: XlsxCellStyleColorInput
  backgroundColor?: XlsxCellStyleColorInput
  patternType?: string
  gradientType?: "linear" | "path"
  gradientStops?: XlsxCellGradientStopInput[]
  gradientDegree?: number
  gradientLeft?: number
  gradientRight?: number
  gradientTop?: number
  gradientBottom?: number
}

export interface XlsxCellNumberFormatInput {
  formatCode?: string
  numFmtId?: number
}

export interface XlsxCellProtectionInput {
  locked?: boolean
  hidden?: boolean
}

export interface XlsxCellStyleInput {
  font?: XlsxCellFontStyleInput
  alignment?: XlsxCellAlignmentInput
  border?: XlsxCellBorderStyleInput
  fill?: XlsxCellFillStyleInput
  numberFormat?: XlsxCellNumberFormatInput
  protection?: XlsxCellProtectionInput
}

export interface XlsxShape {
  id: string
  anchor: XlsxImageAnchor
  sheetIndex: number
  workbookSheetIndex: number
  zIndex: number
  [key: string]: unknown
}

export interface UseXlsxViewerControllerOptions {
  allowResizeInReadOnly?: boolean
  deferLoadingAboveBytes?: number
  file?: ArrayBuffer
  fileName?: string
  maxFileSizeBytes?: number
  readOnly?: boolean
  readOnlyAboveBytes?: number
  showHiddenSheets?: boolean
  skipXmlParsing?: boolean
  src?: string
  useWorker?: boolean
}

// ============================================================
// XlsxViewerController — reactive workbook controller
// ============================================================

export interface XlsxViewerController {
  activeCell: XlsxCellAddress | null
  activeCellAddress: string | null
  activeSheet: XlsxSheetData | null
  activeSheetIndex: number
  activeTab: XlsxWorkbookTab | null
  activeTabIndex: number
  canDownload: boolean
  canExport: boolean
  canLoadDeferred: boolean
  canRedo: boolean
  canUndo: boolean
  canZoomIn: boolean
  canZoomOut: boolean
  clearSelectedCells: () => void
  clearSelection: () => void
  continueDeferredLoad: () => void
  copySelectionToClipboard: () => Promise<boolean>
  defaultZoomScale: number
  deferredLoadFileSize: number | null
  defineNamedRange: (name: string, range?: XlsxCellRange | null) => void
  displayFileName: string
  download: () => void
  charts: XlsxChart[]
  chartsheets: XlsxChartsheet[]
  exportCsv: () => void
  exportXlsx: () => void
  error: Error | null
  file?: ArrayBuffer
  fillSelection: (targetRange: XlsxCellRange) => void
  clearSelectedChart: () => void
  clearSelectedChartElement: () => void
  clearSelectedImage: () => void
  getChartById: (id: string) => XlsxChart | null
  getChartSeriesFormula: (chartId: string, seriesIndex: number) => string
  getChartsheetById: (id: string) => XlsxChartsheet | null
  getImageById: (id: string) => XlsxImage | null
  getSheetCharts: (sheetIndex?: number) => XlsxChart[]
  getSheetImages: (sheetIndex?: number) => XlsxImage[]
  getSheetShapes: (sheetIndex?: number) => XlsxShape[]
  getSheetFormControls: (sheetIndex?: number) => unknown[]
  getClipboardData: () => XlsxClipboardData | null
  getCellDisplayValue: (cell?: XlsxCellAddress | null) => string
  getCellFormula: (cell?: XlsxCellAddress | null) => string
  isLoadDeferred: boolean
  isLoading: boolean
  isChartsLoading: boolean
  images: XlsxImage[]
  shapes: XlsxShape[]
  formControls: unknown[]
  mergeSelection: () => void
  maxZoomScale: number
  minZoomScale: number
  moveChartBy: (id: string, deltaX: number, deltaY: number) => void
  moveImageBy: (id: string, deltaX: number, deltaY: number) => void
  removeActiveSheet: () => void
  readOnly: boolean
  recalculate: () => void
  revision: number
  resetZoom: () => void
  resizeChartBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void
  resizeImageBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void
  resizeColumn: (col: number, widthPx: number) => void
  resizeRow: (row: number, heightPx: number) => void
  redo: () => void
  pasteFromClipboard: () => Promise<boolean>
  pasteStructuredClipboardData: (payload: string) => boolean
  pasteText: (text: string) => boolean
  selectedRangeAddress: string | null
  selectedValue: string
  selectedFormula: string
  setCellFormula: (cell: XlsxCellAddress, formula: string) => void
  setCellStyle: (cell: XlsxCellAddress, style: XlsxCellStyleInput) => void
  setCellValue: (cell: XlsxCellAddress, value: string) => void
  setRangeStyle: (range: XlsxCellRange, style: XlsxCellStyleInput) => void
  setZoomScale: (zoomScale: number) => void
  selectCell: (cell: XlsxCellAddress, options?: { extend?: boolean }) => void
  selectChart: (id: string | null) => void
  selectChartElement: (selection: XlsxChartElementSelection | null) => void
  selectRange: (range: XlsxCellRange) => void
  selection: XlsxCellRange | null
  setActiveSheetIndex: (index: number) => void
  setActiveTabIndex: (index: number) => void
  selectedChart: XlsxChart | null
  selectedChartElement: XlsxChartElementSelection | null
  selectedChartFormula: string | null
  selectedChartId: string | null
  selectedCellFormula: string
  selectedFormulaTarget: XlsxFormulaTarget
  selectedImage: XlsxImage | null
  selectedImageId: string | null
  setSelectedFormula: (formula: string) => boolean
  setSelectedCellFormula: (formula: string) => void
  setSelectedCellStyle: (style: XlsxCellStyleInput) => void
  setSelectedCellValue: (value: string) => void
  sheets: XlsxSheetData[]
  src?: string
  sortState: XlsxTableSortState | null
  sortTable: (tableName: string, columnIndex: number, direction: XlsxTableSortDirection) => void
  setChartSeriesFormula: (chartId: string, seriesIndex: number, formula: string) => boolean
  selectImage: (id: string | null) => void
  setChartRect: (id: string, rect: XlsxImageRect) => void
  setImageRect: (id: string, rect: XlsxImageRect) => void
  tables: XlsxTable[]
  tabs: XlsxWorkbookTab[]
  undo: () => void
  unmergeSelection: () => void
  updateChart: (id: string, patch: Partial<XlsxChart>) => void
  workbook: unknown | null
  zoomIn: () => void
  zoomOut: () => void
  zoomScale: number
  getActiveWorksheet: () => unknown | null
  addSheet: (name?: string) => void
}

// ============================================================
// Sub-interfaces for domain hooks
// ============================================================

export interface XlsxViewerSelection {
  activeCell: XlsxCellAddress | null
  activeCellAddress: string | null
  clearSelection: () => void
  selectedRangeAddress: string | null
  selectCell: (cell: XlsxCellAddress, options?: { extend?: boolean }) => void
  selectRange: (range: XlsxCellRange) => void
  selection: XlsxCellRange | null
}

export interface XlsxViewerZoom {
  canZoomIn: boolean
  canZoomOut: boolean
  defaultZoomScale: number
  maxZoomScale: number
  minZoomScale: number
  resetZoom: () => void
  setZoomScale: (zoomScale: number) => void
  zoomIn: () => void
  zoomOut: () => void
  zoomScale: number
}

export interface XlsxViewerEditing {
  addSheet: (name?: string) => void
  canRedo: boolean
  canUndo: boolean
  clearSelectedCells: () => void
  copySelectionToClipboard: () => Promise<boolean>
  defineNamedRange: (name: string, range?: XlsxCellRange | null) => void
  fillSelection: (targetRange: XlsxCellRange) => void
  getClipboardData: () => XlsxClipboardData | null
  getCellDisplayValue: (cell?: XlsxCellAddress | null) => string
  getCellFormula: (cell?: XlsxCellAddress | null) => string
  mergeSelection: () => void
  pasteFromClipboard: () => Promise<boolean>
  pasteStructuredClipboardData: (payload: string) => boolean
  pasteText: (text: string) => boolean
  removeActiveSheet: () => void
  readOnly: boolean
  redo: () => void
  selectedCellFormula: string
  selectedChartFormula: string | null
  selectedFormula: string
  selectedFormulaTarget: XlsxFormulaTarget
  selectedValue: string
  setCellFormula: (cell: XlsxCellAddress, formula: string) => void
  setCellStyle: (cell: XlsxCellAddress, style: XlsxCellStyleInput) => void
  setCellValue: (cell: XlsxCellAddress, value: string) => void
  setRangeStyle: (range: XlsxCellRange, style: XlsxCellStyleInput) => void
  setSelectedFormula: (formula: string) => boolean
  setSelectedCellFormula: (formula: string) => void
  setSelectedCellStyle: (style: XlsxCellStyleInput) => void
  setSelectedCellValue: (value: string) => void
  undo: () => void
  unmergeSelection: () => void
}

export interface XlsxViewerTables {
  sortState: XlsxTableSortState | null
  sortTable: (tableName: string, columnIndex: number, direction: XlsxTableSortDirection) => void
  tables: XlsxTable[]
}

export interface XlsxViewerImages {
  charts: XlsxChart[]
  clearSelectedChart: () => void
  clearSelectedChartElement: () => void
  clearSelectedImage: () => void
  getChartById: (id: string) => XlsxChart | null
  getChartSeriesFormula: (chartId: string, seriesIndex: number) => string
  getSheetCharts: (sheetIndex?: number) => XlsxChart[]
  getImageById: (id: string) => XlsxImage | null
  getSheetImages: (sheetIndex?: number) => XlsxImage[]
  images: XlsxImage[]
  moveChartBy: (id: string, deltaX: number, deltaY: number) => void
  moveImageBy: (id: string, deltaX: number, deltaY: number) => void
  isChartsLoading: boolean
  readOnly: boolean
  resizeChartBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void
  resizeImageBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void
  selectedChart: XlsxChart | null
  selectedChartElement: XlsxChartElementSelection | null
  selectedChartFormula: string | null
  selectedChartId: string | null
  selectedImage: XlsxImage | null
  selectedImageId: string | null
  selectChart: (id: string | null) => void
  selectChartElement: (selection: XlsxChartElementSelection | null) => void
  selectImage: (id: string | null) => void
  setChartSeriesFormula: (chartId: string, seriesIndex: number, formula: string) => boolean
  setChartRect: (id: string, rect: XlsxImageRect) => void
  setImageRect: (id: string, rect: XlsxImageRect) => void
  updateChart: (id: string, patch: Partial<XlsxChart>) => void
}

export interface XlsxViewerCharts {
  activeTab: XlsxWorkbookTab | null
  activeTabIndex: number
  charts: XlsxChart[]
  chartsheets: XlsxChartsheet[]
  clearSelectedChart: () => void
  clearSelectedChartElement: () => void
  getChartById: (id: string) => XlsxChart | null
  getChartSeriesFormula: (chartId: string, seriesIndex: number) => string
  getChartsheetById: (id: string) => XlsxChartsheet | null
  getSheetCharts: (sheetIndex?: number) => XlsxChart[]
  isChartsLoading: boolean
  moveChartBy: (id: string, deltaX: number, deltaY: number) => void
  readOnly: boolean
  resizeChartBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void
  selectChart: (id: string | null) => void
  selectedChart: XlsxChart | null
  selectedChartElement: XlsxChartElementSelection | null
  selectedChartFormula: string | null
  selectedChartId: string | null
  selectChartElement: (selection: XlsxChartElementSelection | null) => void
  setChartSeriesFormula: (chartId: string, seriesIndex: number, formula: string) => boolean
  setActiveTabIndex: (index: number) => void
  setChartRect: (id: string, rect: XlsxImageRect) => void
  tabs: XlsxWorkbookTab[]
  updateChart: (id: string, patch: Partial<XlsxChart>) => void
}

export interface XlsxSheetThumbnail {
  aspectRatio: number
  contentHeight: number
  contentWidth: number
  height: number
  paint: (canvas: HTMLCanvasElement | null) => boolean
  sheet: XlsxSheetData
  sheetIndex: number
  sourceRange: XlsxCellRange
  width: number
  workbookSheetIndex: number
}

export interface XlsxViewerThumbnails {
  paintThumbnail: (sheetIndex: number, canvas: HTMLCanvasElement | null) => boolean
  thumbnails: XlsxSheetThumbnail[]
}

export interface UseXlsxViewerThumbnailsOptions {
  includeHeaders?: boolean
  resolution?: number | { maxHeight?: number; maxWidth?: number }
}

// ============================================================
// useXlsxViewerController — creates a reactive viewer controller
// ============================================================

function cellAddressToString(addr: XlsxCellAddress): string {
  const colLetter = String.fromCharCode(65 + addr.col)
  return `${colLetter}${addr.row + 1}`
}

function cellRangeToString(range: XlsxCellRange): string {
  return `${cellAddressToString(range.start)}:${cellAddressToString(range.end)}`
}

export function useXlsxViewerController(options: UseXlsxViewerControllerOptions = {}): XlsxViewerController {
  // Core state
  const activeCell = ref<XlsxCellAddress | null>(null)
  const selection = ref<XlsxCellRange | null>(null)
  const activeSheetIndex = ref(0)
  const activeTabIndex = ref(0)
  const zoomScale = ref(1)
  const isDark = ref(false)
  const isLoading = ref(false)
  const isChartsLoading = ref(false)
  const isLoadDeferred = ref(false)
  const deferredLoadFileSize = ref<number | null>(null)
  const error = ref<Error | null>(null)
  const revision = ref(0)
  const readOnly = ref(options.readOnly ?? false)

  // History stacks
  const history: unknown[] = []
  const future: unknown[] = []

  // Workbook data
  const sheets = ref<XlsxSheetData[]>([])
  const tabs = ref<XlsxWorkbookTab[]>([])
  const images = ref<XlsxImage[]>([])
  const charts = ref<XlsxChart[]>([])
  const chartsheets = ref<XlsxChartsheet[]>([])
  const shapes = ref<XlsxShape[]>([])
  const formControls = ref<unknown[]>([])
  const tables = ref<XlsxTable[]>([])
  const sortState = ref<XlsxTableSortState | null>(null)

  // Image/chart selection
  const selectedImageId = ref<string | null>(null)
  const selectedChartId = ref<string | null>(null)
  const selectedChartElement = ref<XlsxChartElementSelection | null>(null)
  const selectedFormula = ref("")
  const selectedFormulaTarget = ref<XlsxFormulaTarget>({})
  const selectedValue = ref("")

  // Derived
  const activeSheet = computed(() => sheets.value[activeSheetIndex.value] ?? null)
  const activeTab = computed(() => tabs.value[activeTabIndex.value] ?? null)
  const activeCellAddress = computed(() => activeCell.value ? cellAddressToString(activeCell.value) : null)
  const selectedRangeAddress = computed(() => selection.value ? cellRangeToString(selection.value) : null)
  const selectedCellFormula = computed(() => "")
  const selectedChartFormula = computed(() => null)
  const selectedImage = computed(() => selectedImageId.value ? images.value.find(i => i.id === selectedImageId.value) ?? null : null)
  const selectedChart = computed(() => selectedChartId.value ? charts.value.find(c => c.id === selectedChartId.value) ?? null : null)

  const canUndo = computed(() => history.length > 0)
  const canRedo = computed(() => future.length > 0)
  const canZoomIn = computed(() => zoomScale.value < 4)
  const canZoomOut = computed(() => zoomScale.value > 0.25)
  const canDownload = computed(() => sheets.value.length > 0)
  const canExport = computed(() => sheets.value.length > 0)
  const canLoadDeferred = computed(() => isLoadDeferred.value)

  const displayFileName = options.fileName ?? options.src?.split(/[?#]/)[0]?.split("/").pop() ?? "Untitled.xlsx"
  const defaultZoomScale = 1
  const maxZoomScale = 4
  const minZoomScale = 0.25
  const workbook = null

  // === Sheet parsing ===
  async function loadFile(buffer: ArrayBuffer) {
    const size = buffer.byteLength
    const maxSize = options.maxFileSizeBytes ?? 25 * 1024 * 1024
    const deferThreshold = options.deferLoadingAboveBytes ?? 0
    const readOnlyThreshold = options.readOnlyAboveBytes ?? 0

    if (size > maxSize) {
      error.value = new Error(`File size (${size} bytes) exceeds maximum (${maxSize} bytes)`)
      return
    }

    if (deferThreshold > 0 && size > deferThreshold) {
      isLoadDeferred.value = true
      deferredLoadFileSize.value = size
      return
    }

    if (readOnlyThreshold > 0 && size > readOnlyThreshold) {
      readOnly.value = true
    }

    isLoading.value = true
    error.value = null
    try {
      // Try to use @extend-ai/xlsx-core for parsing
      const { parseXlsxBuffer } = await import("@extend-ai/xlsx-core")
      const result = parseXlsxBuffer(buffer) as unknown as {
        sheets: XlsxSheetData[]
        tabs: XlsxWorkbookTab[]
        images?: XlsxImage[]
        charts?: XlsxChart[]
        tables?: XlsxTable[]
      }
      sheets.value = result.sheets
      tabs.value = result.tabs
      images.value = result.images ?? []
      charts.value = result.charts ?? []
      tables.value = result.tables ?? []
      activeSheetIndex.value = 0
      activeTabIndex.value = 0
      revision.value++
    } catch (e) {
      const err = e instanceof Error ? e : new Error(String(e))
      error.value = new Error(`Failed to load workbook: ${err.message}`)
      sheets.value = []
      tabs.value = []
      images.value = []
      charts.value = []
      tables.value = []
      activeSheetIndex.value = 0
      activeTabIndex.value = 0
      revision.value++
    } finally {
      isLoading.value = false
    }
  }

  function createEmptySheet(name: string, workbookSheetIndex: number): XlsxSheetData {
    return {
      cachedFormulaValues: {},
      colWidthOverridesPx: {},
      colStyleIds: {},
      conditionalFormatRules: [],
      dataValidations: [],
      name,
      visibility: "visible",
      defaultColWidthPx: 80,
      defaultRowHeightPx: 24,
      freezePanes: null,
      hasHorizontalMerges: false,
      hasVerticalMerges: false,
      maxHorizontalMergeEndCol: 0,
      maxVerticalMergeEndRow: 0,
      minUsedCol: 0,
      minUsedRow: 0,
      maxUsedCol: 0,
      maxUsedRow: 0,
      rowCount: 200,
      colCount: 50,
      rowHeightOverridesPx: {},
      rowStyleIds: {},
      namedCellStyleByName: {},
      styleById: {},
      tableStyleByName: {},
      visibleRows: [],
      visibleCols: [],
      colWidths: new Array(50).fill(80),
      rowHeights: new Array(200).fill(24),
      showGridLines: true,
      sparklines: [],
      themePalette: { colorsByIndex: {} },
      workbookSheetIndex,
      zoomScale: 1,
    }
  }

  function continueDeferredLoad() {
    isLoadDeferred.value = false
    deferredLoadFileSize.value = null
    if (options.file) loadFile(options.file)
  }

  // Watch file option
  if (options.file) {
    loadFile(options.file)
  }

  // Watch src option
  if (options.src) {
    isLoading.value = true
    fetch(options.src)
      .then(r => {
        if (!r.ok) throw new Error(`Failed to fetch workbook (${r.status})`)
        return r.arrayBuffer()
      })
      .then(buf => loadFile(buf))
      .catch(e => { error.value = e instanceof Error ? e : new Error(String(e)) })
      .finally(() => { isLoading.value = false })
  }

  // === Controller object ===
  const controller: XlsxViewerController = {
    // Getters
    get activeCell() { return activeCell.value },
    get activeCellAddress() { return activeCellAddress.value },
    get activeSheet() { return activeSheet.value },
    get activeSheetIndex() { return activeSheetIndex.value },
    get activeTab() { return activeTab.value },
    get activeTabIndex() { return activeTabIndex.value },
    get canDownload() { return canDownload.value },
    get canExport() { return canExport.value },
    get canLoadDeferred() { return canLoadDeferred.value },
    get canRedo() { return canRedo.value },
    get canUndo() { return canUndo.value },
    get canZoomIn() { return canZoomIn.value },
    get canZoomOut() { return canZoomOut.value },
    defaultZoomScale,
    get deferredLoadFileSize() { return deferredLoadFileSize.value },
    displayFileName,
    get charts() { return charts.value },
    get chartsheets() { return chartsheets.value },
    get error() { return error.value },
    get file() { return options.file },
    get images() { return images.value },
    get shapes() { return shapes.value },
    get formControls() { return formControls.value },
    get isLoadDeferred() { return isLoadDeferred.value },
    get isLoading() { return isLoading.value },
    get isChartsLoading() { return isChartsLoading.value },
    maxZoomScale,
    minZoomScale,
    get readOnly() { return readOnly.value },
    get revision() { return revision.value },
    get selectedRangeAddress() { return selectedRangeAddress.value },
    get selectedValue() { return selectedValue.value },
    get selectedFormula() { return selectedFormula.value },
    get selection() { return selection.value },
    get selectedChart() { return selectedChart.value },
    get selectedChartElement() { return selectedChartElement.value },
    get selectedChartFormula() { return selectedChartFormula.value },
    get selectedChartId() { return selectedChartId.value },
    get selectedCellFormula() { return selectedCellFormula.value },
    get selectedFormulaTarget() { return selectedFormulaTarget.value },
    get selectedImage() { return selectedImage.value },
    get selectedImageId() { return selectedImageId.value },
    get sheets() { return sheets.value },
    get sortState() { return sortState.value },
    get src() { return options.src },
    get tables() { return tables.value },
    get tabs() { return tabs.value },
    workbook,
    get zoomScale() { return zoomScale.value },

    // Actions
    clearSelectedCells() {
      selectedValue.value = ""
      selectedFormula.value = ""
    },
    clearSelection() {
      activeCell.value = null
      selection.value = null
    },
    continueDeferredLoad,
    async copySelectionToClipboard(): Promise<boolean> {
      if (!selection.value) return false
      try {
        const addr = selectedRangeAddress.value ?? ""
        await navigator.clipboard.writeText(addr)
        return true
      } catch { return false }
    },
    defineNamedRange(_name: string, _range?: XlsxCellRange | null) { /* noop */ },
    download() {
      const payload = JSON.stringify({ sheets: sheets.value.map((sheet) => ({ name: sheet.name, rowCount: sheet.rowCount, colCount: sheet.colCount })) }, null, 2)
      const blob = new Blob([payload], { type: "application/json" })
      const url = URL.createObjectURL(blob)
      const a = document.createElement("a")
      a.href = url
      a.download = displayFileName.toLowerCase().endsWith(".xlsx") ? displayFileName : `${displayFileName}.xlsx`
      a.rel = "noopener"
      document.body.append(a)
      a.click()
      a.remove()
      setTimeout(() => URL.revokeObjectURL(url), 0)
    },
    exportCsv() { /* noop */ },
    exportXlsx() { this.download() },
    fillSelection(_targetRange: XlsxCellRange) { /* noop */ },
    clearSelectedChart() { selectedChartId.value = null },
    clearSelectedChartElement() { selectedChartElement.value = null },
    clearSelectedImage() { selectedImageId.value = null },
    getChartById(id: string) { return charts.value.find(c => c.id === id) ?? null },
    getChartSeriesFormula(_chartId: string, _seriesIndex: number): string { return "" },
    getChartsheetById(id: string) { return chartsheets.value.find(c => c.id === id) ?? null },
    getImageById(id: string) { return images.value.find(i => i.id === id) ?? null },
    getSheetCharts(sheetIndex?: number) {
      const idx = sheetIndex ?? activeSheetIndex.value
      return charts.value.filter(c => c.workbookSheetIndex === idx)
    },
    getSheetImages(sheetIndex?: number) {
      const idx = sheetIndex ?? activeSheetIndex.value
      return images.value.filter(i => i.workbookSheetIndex === idx)
    },
    getSheetShapes(sheetIndex?: number) {
      const idx = sheetIndex ?? activeSheetIndex.value
      return shapes.value.filter(s => s.workbookSheetIndex === idx)
    },
    getSheetFormControls(sheetIndex?: number) {
      const idx = sheetIndex ?? activeSheetIndex.value
      return formControls.value.filter(
        (f: unknown) => (f as Record<string, unknown>).workbookSheetIndex === idx
      )
    },
    getClipboardData() { return null },
    getCellDisplayValue(_cell?: XlsxCellAddress | null): string { return "" },
    getCellFormula(_cell?: XlsxCellAddress | null): string { return "" },
    mergeSelection() { /* noop */ },
    moveChartBy(_id: string, _deltaX: number, _deltaY: number) { /* noop */ },
    moveImageBy(_id: string, _deltaX: number, _deltaY: number) { /* noop */ },
    removeActiveSheet() { /* noop */ },
    recalculate() { /* noop */ },
    resetZoom() { zoomScale.value = defaultZoomScale },
    resizeChartBy(_id: string, _handle: XlsxImageResizeHandlePosition, _deltaX: number, _deltaY: number) { /* noop */ },
    resizeImageBy(_id: string, _handle: XlsxImageResizeHandlePosition, _deltaX: number, _deltaY: number) { /* noop */ },
    resizeColumn(_col: number, _widthPx: number) { /* noop */ },
    resizeRow(_row: number, _heightPx: number) { /* noop */ },
    redo() { /* noop */ },
    async pasteFromClipboard(): Promise<boolean> {
      try {
        const text = await navigator.clipboard.readText()
        return this.pasteText(text)
      } catch { return false }
    },
    pasteStructuredClipboardData(_payload: string): boolean { return false },
    pasteText(_text: string): boolean { return false },
    setCellFormula(_cell: XlsxCellAddress, _formula: string) { /* noop */ },
    setCellStyle(_cell: XlsxCellAddress, _style: XlsxCellStyleInput) { /* noop */ },
    setCellValue(_cell: XlsxCellAddress, value: string) {
      selectedValue.value = value
      revision.value++
    },
    setRangeStyle(_range: XlsxCellRange, _style: XlsxCellStyleInput) { /* noop */ },
    setZoomScale(s: number) {
      zoomScale.value = Math.max(minZoomScale, Math.min(maxZoomScale, s))
    },
    selectCell(cell: XlsxCellAddress, opts?: { extend?: boolean }) {
      activeCell.value = cell
      if (opts?.extend && selection.value) {
        selection.value = { start: selection.value.start, end: cell }
      } else {
        selection.value = { start: cell, end: cell }
      }
      revision.value++
    },
    selectChart(id: string | null) { selectedChartId.value = id },
    selectChartElement(sel: XlsxChartElementSelection | null) { selectedChartElement.value = sel },
    selectRange(range: XlsxCellRange) {
      selection.value = range
      activeCell.value = range.start
      revision.value++
    },
    setActiveSheetIndex(index: number) {
      if (index >= 0 && index < sheets.value.length) {
        activeSheetIndex.value = index
        activeTabIndex.value = index
      }
    },
    setActiveTabIndex(index: number) {
      if (index >= 0 && index < tabs.value.length) {
        activeTabIndex.value = index
        const tab = tabs.value[index]
        if (tab.kind === "sheet" && tab.sheetIndex != null) {
          activeSheetIndex.value = tab.sheetIndex
        }
      }
    },
    setSelectedFormula(_formula: string): boolean { return false },
    setSelectedCellFormula(_formula: string) { /* noop */ },
    setSelectedCellStyle(_style: XlsxCellStyleInput) { /* noop */ },
    setSelectedCellValue(value: string) {
      selectedValue.value = value
      revision.value++
    },
    sortTable(tableName: string, columnIndex: number, direction: XlsxTableSortDirection) {
      sortState.value = { tableName, columnIndex, direction }
    },
    setChartSeriesFormula(_chartId: string, _seriesIndex: number, _formula: string): boolean { return false },
    selectImage(id: string | null) { selectedImageId.value = id },
    setChartRect(_id: string, _rect: XlsxImageRect) { /* noop */ },
    setImageRect(_id: string, _rect: XlsxImageRect) { /* noop */ },
    undo() { /* noop */ },
    unmergeSelection() { /* noop */ },
    updateChart(_id: string, _patch: Partial<XlsxChart>) { /* noop */ },
    zoomIn() { zoomScale.value = Math.min(maxZoomScale, zoomScale.value + 0.1) },
    zoomOut() { zoomScale.value = Math.max(minZoomScale, zoomScale.value - 0.1) },
    getActiveWorksheet() { return null },
    addSheet(_name?: string) {
      const sheetName = _name ?? `Sheet${sheets.value.length + 1}`
      const idx = sheets.value.length
      sheets.value.push(createEmptySheet(sheetName, idx))
      tabs.value.push({ id: `tab-${idx}`, index: idx, kind: "sheet", name: sheetName, sheetIndex: idx, workbookSheetIndex: idx })
    },
  }

  return controller
}

// ============================================================
// useXlsxViewer — obtain controller from provide/inject
// ============================================================

export function useXlsxViewer(): XlsxViewerController {
  const controller = inject<XlsxViewerController>(XLSX_VIEWER_KEY)
  if (!controller) {
    throw new Error("useXlsxViewer must be used inside XlsxViewer or XlsxViewerProvider.")
  }
  return controller
}

// ============================================================
// useXlsxViewerSelection
// ============================================================

export function useXlsxViewerSelection(): XlsxViewerSelection {
  const ctrl = useXlsxViewer()
  return {
    get activeCell() { return ctrl.activeCell },
    get activeCellAddress() { return ctrl.activeCellAddress },
    clearSelection: ctrl.clearSelection,
    get selectedRangeAddress() { return ctrl.selectedRangeAddress },
    selectCell: ctrl.selectCell,
    selectRange: ctrl.selectRange,
    get selection() { return ctrl.selection },
  }
}

// ============================================================
// useXlsxViewerZoom
// ============================================================

export function useXlsxViewerZoom(): XlsxViewerZoom {
  const ctrl = useXlsxViewer()
  return {
    get canZoomIn() { return ctrl.canZoomIn },
    get canZoomOut() { return ctrl.canZoomOut },
    defaultZoomScale: ctrl.defaultZoomScale,
    maxZoomScale: ctrl.maxZoomScale,
    minZoomScale: ctrl.minZoomScale,
    resetZoom: ctrl.resetZoom,
    setZoomScale: ctrl.setZoomScale,
    zoomIn: ctrl.zoomIn,
    zoomOut: ctrl.zoomOut,
    get zoomScale() { return ctrl.zoomScale },
  }
}

// ============================================================
// useXlsxViewerEditing
// ============================================================

export function useXlsxViewerEditing(): XlsxViewerEditing {
  const ctrl = useXlsxViewer()
  return {
    addSheet: ctrl.addSheet,
    get canRedo() { return ctrl.canRedo },
    get canUndo() { return ctrl.canUndo },
    clearSelectedCells: ctrl.clearSelectedCells,
    copySelectionToClipboard: ctrl.copySelectionToClipboard,
    defineNamedRange: ctrl.defineNamedRange,
    fillSelection: ctrl.fillSelection,
    getClipboardData: ctrl.getClipboardData,
    getCellDisplayValue: ctrl.getCellDisplayValue,
    getCellFormula: ctrl.getCellFormula,
    mergeSelection: ctrl.mergeSelection,
    pasteFromClipboard: ctrl.pasteFromClipboard,
    pasteStructuredClipboardData: ctrl.pasteStructuredClipboardData,
    pasteText: ctrl.pasteText,
    removeActiveSheet: ctrl.removeActiveSheet,
    get readOnly() { return ctrl.readOnly },
    redo: ctrl.redo,
    get selectedCellFormula() { return ctrl.selectedCellFormula },
    get selectedChartFormula() { return ctrl.selectedChartFormula },
    get selectedFormula() { return ctrl.selectedFormula },
    get selectedFormulaTarget() { return ctrl.selectedFormulaTarget },
    get selectedValue() { return ctrl.selectedValue },
    setCellFormula: ctrl.setCellFormula,
    setCellStyle: ctrl.setCellStyle,
    setCellValue: ctrl.setCellValue,
    setRangeStyle: ctrl.setRangeStyle,
    setSelectedFormula: ctrl.setSelectedFormula,
    setSelectedCellFormula: ctrl.setSelectedCellFormula,
    setSelectedCellStyle: ctrl.setSelectedCellStyle,
    setSelectedCellValue: ctrl.setSelectedCellValue,
    undo: ctrl.undo,
    unmergeSelection: ctrl.unmergeSelection,
  }
}

// ============================================================
// useXlsxViewerTables
// ============================================================

export function useXlsxViewerTables(): XlsxViewerTables {
  const ctrl = useXlsxViewer()
  return {
    get sortState() { return ctrl.sortState },
    sortTable: ctrl.sortTable,
    get tables() { return ctrl.tables },
  }
}

// ============================================================
// useXlsxViewerImages
// ============================================================

export function useXlsxViewerImages(): XlsxViewerImages {
  const ctrl = useXlsxViewer()
  return {
    get charts() { return ctrl.charts },
    clearSelectedChart: ctrl.clearSelectedChart,
    clearSelectedChartElement: ctrl.clearSelectedChartElement,
    clearSelectedImage: ctrl.clearSelectedImage,
    getChartById: ctrl.getChartById,
    getChartSeriesFormula: ctrl.getChartSeriesFormula,
    getSheetCharts: ctrl.getSheetCharts,
    getImageById: ctrl.getImageById,
    getSheetImages: ctrl.getSheetImages,
    get images() { return ctrl.images },
    moveChartBy: ctrl.moveChartBy,
    moveImageBy: ctrl.moveImageBy,
    get isChartsLoading() { return ctrl.isChartsLoading },
    get readOnly() { return ctrl.readOnly },
    resizeChartBy: ctrl.resizeChartBy,
    resizeImageBy: ctrl.resizeImageBy,
    get selectedChart() { return ctrl.selectedChart },
    get selectedChartElement() { return ctrl.selectedChartElement },
    get selectedChartFormula() { return ctrl.selectedChartFormula },
    get selectedChartId() { return ctrl.selectedChartId },
    get selectedImage() { return ctrl.selectedImage },
    get selectedImageId() { return ctrl.selectedImageId },
    selectChart: ctrl.selectChart,
    selectChartElement: ctrl.selectChartElement,
    selectImage: ctrl.selectImage,
    setChartSeriesFormula: ctrl.setChartSeriesFormula,
    setChartRect: ctrl.setChartRect,
    setImageRect: ctrl.setImageRect,
    updateChart: ctrl.updateChart,
  }
}

// ============================================================
// useXlsxViewerCharts
// ============================================================

export function useXlsxViewerCharts(): XlsxViewerCharts {
  const ctrl = useXlsxViewer()
  return {
    get activeTab() { return ctrl.activeTab },
    get activeTabIndex() { return ctrl.activeTabIndex },
    get charts() { return ctrl.charts },
    get chartsheets() { return ctrl.chartsheets },
    clearSelectedChart: ctrl.clearSelectedChart,
    clearSelectedChartElement: ctrl.clearSelectedChartElement,
    getChartById: ctrl.getChartById,
    getChartSeriesFormula: ctrl.getChartSeriesFormula,
    getChartsheetById: ctrl.getChartsheetById,
    getSheetCharts: ctrl.getSheetCharts,
    get isChartsLoading() { return ctrl.isChartsLoading },
    moveChartBy: ctrl.moveChartBy,
    get readOnly() { return ctrl.readOnly },
    resizeChartBy: ctrl.resizeChartBy,
    selectChart: ctrl.selectChart,
    get selectedChart() { return ctrl.selectedChart },
    get selectedChartElement() { return ctrl.selectedChartElement },
    get selectedChartFormula() { return ctrl.selectedChartFormula },
    get selectedChartId() { return ctrl.selectedChartId },
    selectChartElement: ctrl.selectChartElement,
    setChartSeriesFormula: ctrl.setChartSeriesFormula,
    setActiveTabIndex: ctrl.setActiveTabIndex,
    setChartRect: ctrl.setChartRect,
    get tabs() { return ctrl.tabs },
    updateChart: ctrl.updateChart,
  }
}

// ============================================================
// useXlsxViewerThumbnails
// ============================================================

export function useXlsxViewerThumbnails(
  _options: UseXlsxViewerThumbnailsOptions = {}
): XlsxViewerThumbnails {
  const ctrl = useXlsxViewer()

  const thumbnails = computed<XlsxSheetThumbnail[]>(() => {
    return ctrl.sheets.map((sheet, i) => {
      const defaultColW = sheet.defaultColWidthPx
      const defaultRowH = sheet.defaultRowHeightPx
      const contentW = sheet.colCount * defaultColW
      const contentH = sheet.rowCount * defaultRowH
      const aspect = contentH > 0 ? contentW / contentH : 1

      return {
        aspectRatio: aspect,
        contentHeight: contentH,
        contentWidth: contentW,
        height: 150,
        width: Math.round(150 * aspect),
        paint(_canvas: HTMLCanvasElement | null): boolean { return false },
        sheet,
        sheetIndex: i,
        sourceRange: {
          start: { col: 0, row: 0 },
          end: { col: sheet.colCount - 1, row: sheet.rowCount - 1 },
        },
        workbookSheetIndex: sheet.workbookSheetIndex,
      }
    })
  })

  function paintThumbnail(sheetIndex: number, canvas: HTMLCanvasElement | null): boolean {
    const t = thumbnails.value[sheetIndex]
    return t ? t.paint(canvas) : false
  }

  return {
    paintThumbnail,
    get thumbnails() { return thumbnails.value },
  }
}

