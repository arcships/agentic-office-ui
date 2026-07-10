import type { Workbook, Worksheet } from "@dukelib/sheets-wasm";
import type { XlsxChart, XlsxChartElementSelection, XlsxChartsheet } from "./chart-types";
import type { XlsxThemePalette } from "./theme-types";
import type {
  XlsxFormControl,
  XlsxImage,
  XlsxImageRect,
  XlsxImageResizeHandlePosition,
  XlsxShape,
} from "./image-types";

// ── Cell addressing ──────────────────────────────────────────────────────

export interface XlsxCellAddress {
  col: number;
  row: number;
}

export interface XlsxCellRange {
  end: XlsxCellAddress;
  start: XlsxCellAddress;
}

// ── Theme / styles ───────────────────────────────────────────────────────

export interface XlsxResolvedCellStyle {
  [key: string]: unknown;
  alignment?: Record<string, unknown>;
  border?: Record<string, Record<string, unknown>>;
  cellControl?: {
    kind: "checkbox";
  };
  fill?: Record<string, unknown>;
  font?: Record<string, unknown>;
}

export interface XlsxTableStyleDefinition {
  [elementType: string]: XlsxResolvedCellStyle;
}

// ── Conditional formatting ───────────────────────────────────────────────

export interface XlsxConditionalFormatValueObject {
  type: string;
  value?: number;
}

export interface XlsxConditionalDataBarRule {
  axisColor?: Record<string, unknown>;
  border?: boolean;
  color?: Record<string, unknown>;
  borderColor?: Record<string, unknown>;
  cfvos: XlsxConditionalFormatValueObject[];
  gradient?: boolean;
  kind: "dataBar";
  maxLength?: number;
  minLength?: number;
  negativeBarBorderColorSameAsPositive?: boolean;
  negativeBorderColor?: Record<string, unknown>;
  negativeFillColor?: Record<string, unknown>;
  priority: number;
  ranges: XlsxCellRange[];
  showValue?: boolean;
}

export interface XlsxConditionalColorScaleRule {
  cfvos: XlsxConditionalFormatValueObject[];
  colors: Record<string, unknown>[];
  kind: "colorScale";
  priority: number;
  ranges: XlsxCellRange[];
}

export interface XlsxConditionalFormatIcon {
  iconId: number;
  iconSet: string;
}

export interface XlsxConditionalIconSetRule {
  cfvos: XlsxConditionalFormatValueObject[];
  icons: XlsxConditionalFormatIcon[];
  kind: "iconSet";
  priority: number;
  ranges: XlsxCellRange[];
  reverse?: boolean;
  showValue?: boolean;
}

export type XlsxConditionalFormatRule =
  | XlsxConditionalColorScaleRule
  | XlsxConditionalDataBarRule
  | XlsxConditionalIconSetRule;

// ── Data validation / freeze panes ───────────────────────────────────────

export interface XlsxDataValidation {
  allowBlank?: boolean;
  errorMessage?: string;
  errorStyle?: string;
  inputMessage?: string;
  listSource?: string;
  ranges: XlsxCellRange[];
  showDropdown?: boolean;
  showErrorAlert?: boolean;
  showInputMessage?: boolean;
  validationType: string;
}

export interface XlsxFreezePanes {
  col: number;
  row: number;
}

export type XlsxSheetVisibility = "hidden" | "veryHidden" | "visible";

// ── Sheet data ───────────────────────────────────────────────────────────

export interface XlsxSheetData {
  cachedFormulaValues: Record<string, string>;
  colWidthOverridesPx: Record<number, number>;
  colStyleIds: Record<number, number>;
  hiddenCols?: number[];
  hiddenRows?: number[];
  conditionalFormatRules: XlsxConditionalFormatRule[];
  dataValidations: XlsxDataValidation[];
  name: string;
  visibility: XlsxSheetVisibility;
  columnWidthCharacterWidthPx?: number;
  defaultColWidthPx: number;
  defaultRowHeightPx: number;
  freezePanes: XlsxFreezePanes | null;
  hasHorizontalMerges: boolean;
  hasVerticalMerges: boolean;
  maxHorizontalMergeEndCol: number;
  maxVerticalMergeEndRow: number;
  minUsedCol: number;
  minUsedRow: number;
  maxUsedCol: number;
  maxUsedRow: number;
  rowCount: number;
  colCount: number;
  rowHeightOverridesPx: Record<number, number>;
  rowStyleIds: Record<number, number>;
  namedCellStyleByName: Record<string, XlsxResolvedCellStyle>;
  styleById: Record<number, XlsxResolvedCellStyle>;
  tableStyleByName: Record<string, XlsxTableStyleDefinition>;
  visibleRows: number[];
  visibleCols: number[];
  colWidths: number[];
  rowHeights: number[];
  showGridLines: boolean;
  sparklines: XlsxSparkline[];
  themePalette: XlsxThemePalette;
  workbookSheetIndex: number;
  zoomScale?: number;
}

// ── Cell style inputs ────────────────────────────────────────────────────

/**
 * Color value accepted by persisted workbook style mutation APIs.
 *
 * Use `rgb` with a six-character `hex` value such as `"2563EB"` for the
 * simplest browser-to-Excel mapping. `argb`, `theme`, and `indexed` are
 * available for callers that need to preserve Excel-native color references.
 */
export interface XlsxCellStyleColorInput {
  /** Excel color representation to write. Defaults are handled by the workbook engine. */
  colorType?: "auto" | "rgb" | "argb" | "theme" | "indexed";
  /** Hex color string, for example `"2563EB"` for RGB or `"FF2563EB"` for ARGB. */
  hex?: string;
  /** Red channel, 0-255. */
  r?: number;
  /** Green channel, 0-255. */
  g?: number;
  /** Blue channel, 0-255. */
  b?: number;
  /** Alpha channel, 0-255, used by ARGB colors. */
  a?: number;
  /** Theme color index for Excel theme colors. */
  themeIndex?: number;
  /** Theme tint/shade adjustment. */
  tint?: number;
  /** Indexed color palette entry. */
  paletteIndex?: number;
}

/** Font styling to persist on a cell or range. */
export interface XlsxCellFontStyleInput {
  /** Font face name, such as `"Aptos"` or `"Calibri"`. */
  name?: string;
  /** Font size in points. */
  size?: number;
  /** Whether the font is bold. */
  bold?: boolean;
  /** Whether the font is italic. */
  italic?: boolean;
  /** Underline style. */
  underline?: "none" | "single" | "double" | "singleAccounting" | "doubleAccounting";
  /** Whether the font is struck through. */
  strikethrough?: boolean;
  /** Font color. */
  color?: XlsxCellStyleColorInput;
  /** Baseline, superscript, or subscript positioning. */
  verticalAlign?: "baseline" | "superscript" | "subscript";
  /** Excel font family classification. */
  family?: number;
  /** Font charset identifier. */
  charset?: number;
  /** Excel font scheme value. */
  scheme?: string;
}

/** A color stop used by gradient fills. */
export interface XlsxCellGradientStopInput {
  /** Stop position from 0 to 1. */
  position: number;
  /** Stop color. */
  color: XlsxCellStyleColorInput;
}

/** Fill styling to persist on a cell or range. */
export interface XlsxCellFillStyleInput {
  /** Fill type. Use `"solid"` with `color` for typical Excel fill color. */
  fillType?: "none" | "solid" | "pattern" | "gradient";
  /** Primary fill color. */
  color?: XlsxCellStyleColorInput;
  /** Excel pattern name for pattern fills. */
  pattern?: string;
  /** Foreground color for pattern fills. */
  foreground?: XlsxCellStyleColorInput;
  /** Background color for pattern fills. */
  background?: XlsxCellStyleColorInput;
  /** Gradient type for gradient fills. */
  gradientType?: "linear" | "path";
  /** Gradient angle in degrees. */
  angle?: number;
  /** Gradient color stops. */
  stops?: XlsxCellGradientStopInput[];
}

/** A single border edge style. */
export interface XlsxCellBorderEdgeInput {
  /** Excel border line style. */
  style?: "none" | "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | "hair" | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot";
  /** Border line color. */
  color?: XlsxCellStyleColorInput;
}

/** Border styling to persist on a cell or range. */
export interface XlsxCellBorderStyleInput {
  /** Left border edge. */
  left?: XlsxCellBorderEdgeInput;
  /** Right border edge. */
  right?: XlsxCellBorderEdgeInput;
  /** Top border edge. */
  top?: XlsxCellBorderEdgeInput;
  /** Bottom border edge. */
  bottom?: XlsxCellBorderEdgeInput;
  /** Diagonal border edge. */
  diagonal?: XlsxCellBorderEdgeInput;
  /** Diagonal border direction. */
  diagonalDirection?: "none" | "down" | "up" | "both";
}

/** Alignment styling to persist on a cell or range. */
export interface XlsxCellAlignmentInput {
  /** Horizontal alignment. */
  horizontal?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
  /** Vertical alignment. */
  vertical?: "top" | "center" | "bottom" | "justify" | "distributed";
  /** Whether text wraps inside the cell. */
  wrapText?: boolean;
  /** Whether text shrinks to fit the cell. */
  shrinkToFit?: boolean;
  /** Indentation level. */
  indent?: number;
  /** Text rotation in degrees. */
  rotation?: number;
  /** Text reading order. */
  readingOrder?: "contextDependent" | "leftToRight" | "rightToLeft";
}

/** Number format styling to persist on a cell or range. */
export interface XlsxCellNumberFormatInput {
  /** Number format source. Use `"custom"` with `formatString` for custom Excel formats. */
  formatType?: "general" | "builtin" | "custom";
  /** Built-in Excel number format id. */
  id?: number;
  /** Excel number format string, for example `"$#,##0.00"` or `"m/d/yyyy"`. */
  formatString?: string;
}

/** Protection flags to persist on a cell or range. */
export interface XlsxCellProtectionInput {
  /** Whether the cell is locked when sheet protection is enabled. */
  locked?: boolean;
  /** Whether the cell formula is hidden when sheet protection is enabled. */
  hidden?: boolean;
}

/**
 * Persisted Excel cell style patch.
 *
 * Passing a partial style updates the supplied style groups on the target cell
 * or range. Unlike the `getCellStyle` render override prop, this mutates the
 * workbook and is included in exported XLSX bytes.
 */
export interface XlsxCellStyleInput {
  /** Font formatting such as bold, italic, size, family, and font color. */
  font?: XlsxCellFontStyleInput;
  /** Cell fill formatting such as background color or gradient fill. */
  fill?: XlsxCellFillStyleInput;
  /** Border formatting for cell edges. */
  border?: XlsxCellBorderStyleInput;
  /** Alignment, wrapping, indentation, rotation, and reading order. */
  alignment?: XlsxCellAlignmentInput;
  /** Excel number/date/currency/percent format metadata. */
  numberFormat?: XlsxCellNumberFormatInput;
  /** Cell protection flags. */
  protection?: XlsxCellProtectionInput;
}

// ── Sparklines ───────────────────────────────────────────────────────────

export interface XlsxSparkline {
  color?: string;
  firstColor?: string;
  highColor?: string;
  lastColor?: string;
  lowColor?: string;
  markerColor?: string;
  markers?: boolean;
  negative?: boolean;
  negativeColor?: string;
  range: XlsxCellRange;
  sheetName?: string;
  target: XlsxCellAddress;
  type: "column" | "line" | "winLoss";
}

// ── Clipboard ────────────────────────────────────────────────────────────

export interface XlsxClipboardData {
  html: string;
  structured: string;
  text: string;
}

// ── Tables ───────────────────────────────────────────────────────────────

export interface XlsxTableColumn {
  id: number;
  index: number;
  name: string;
}

export interface XlsxTableStyleInfo {
  name?: string;
  showColumnStripes?: boolean;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
}

export interface XlsxTable {
  columns: XlsxTableColumn[];
  displayName: string;
  end: XlsxCellAddress;
  headerRowCount: number;
  headerRowCellStyle?: string;
  name: string;
  reference: string;
  start: XlsxCellAddress;
  styleInfo?: XlsxTableStyleInfo;
  totalsRowCount: number;
  totalsRowShown: boolean;
}

export type XlsxTableSortDirection = "ascending" | "descending";

export interface XlsxTableSortState {
  columnIndex: number;
  direction: XlsxTableSortDirection;
  tableName: string;
}

// ── Workbook tab ─────────────────────────────────────────────────────────

export interface XlsxWorkbookTab {
  chartsheetIndex?: number;
  id: string;
  index: number;
  kind: "chartsheet" | "sheet";
  name: string;
  sheetIndex?: number;
  visibility?: XlsxSheetVisibility;
  workbookSheetIndex?: number;
}

// ── Formula target ───────────────────────────────────────────────────────

export type XlsxFormulaTarget =
  | { kind: "cell"; cell: XlsxCellAddress | null }
  | { kind: "chartSeries"; chartId: string; seriesId: string; seriesIndex: number }
  | null;

// ── Cell style context ───────────────────────────────────────────────────

export interface XlsxCellStyleContext {
  /** Address of the cell being styled. */
  cell: XlsxCellAddress;
  /** True when the cell is part of a selected chart's highlighted source range. */
  hasChartHighlight: boolean;
  /** True when a conditional format (color scale, data bar, or icon set) applies to the cell. */
  hasConditionalFormat: boolean;
  /** True when the cell has a hyperlink. */
  hasHyperlink: boolean;
  /** True when the cell has a data validation rule. */
  hasValidation: boolean;
  /** True when the cell is the anchor of a merged range. */
  isMerged: boolean;
  /** True when the cell is a table header cell. */
  isTableHeader: boolean;
  /**
   * The fully resolved style the viewer computed for this cell, combining the workbook's
   * own formatting with the viewer's built-in styling. Treat this as read-only; returning
   * a partial style from `getCellStyle` merges on top of it.
   */
  resolvedStyle: Record<string, string | number | undefined>;
  /** Display name of the sheet the cell belongs to. */
  sheetName: string;
  /** The cell's resolved display value. */
  value: string;
  /** Workbook sheet index of the cell's sheet. */
  workbookSheetIndex: number;
}

// ── Table header menu ────────────────────────────────────────────────────

export interface XlsxTableHeaderMenuRenderProps {
  /** Address of the table header cell that owns this menu. */
  cell: XlsxCellAddress;
  /** Table column metadata for the active header. */
  column: XlsxTableColumn;
  /** Current sort direction for this column, or `null` when unsorted. */
  direction: XlsxTableSortDirection | null;
  /** Sorts the table by this column in ascending order. */
  sortAscending: () => void;
  /** Sorts the table by this column in descending order. */
  sortDescending: () => void;
  /** Table metadata for the active header. */
  table: XlsxTable;
  /** Built-in trigger text/icon. Useful when composing your own trigger button. */
  triggerIcon: string;
  /** Props that must be applied to your menu trigger button so grid selection does not receive the click. */
  triggerProps: Record<string, unknown>;
}

// ── Viewer controller ────────────────────────────────────────────────────

/** Minimal structural contract accepted from an explicit XLSX Runtime. */
export interface XlsxRuntimeLike {
  readonly id: string;
  readonly wasmSource: string | ArrayBuffer | WebAssembly.Module;
  readonly workerUrl?: string;
  readonly parseOptions: Readonly<{
    showHiddenSheets?: boolean;
    skipXmlParsing?: boolean;
  }>;
  createWorkerClient(): unknown;
  dispose(): void;
}

export type XlsxSourceKind = "file" | "url";
export type XlsxSourceState = "idle" | "loading" | "ready" | "error";

/**
 * Per-controller URL rules for workbook sources. The host supplies the base
 * URL and allowed origins; core and Vue packages do not read browser globals
 * to widen their own permissions.
 */
export interface XlsxUrlPolicy {
  enabled?: boolean;
  baseUrl?: string;
  allowRelativeUrl?: boolean;
  allowedProtocols?: readonly string[];
  allowedOrigins?: readonly string[];
  allowHttpOnLocalhost?: boolean;
  fetch?: typeof fetch;
}

export type XlsxLoadErrorCode = "SOURCE_NOT_ALLOWED" | "FETCH_FAILED" | "INVALID_WORKBOOK" | "ABORTED";

export interface XlsxLoadError {
  code: XlsxLoadErrorCode;
  message: string;
  sourceKind: XlsxSourceKind;
  url?: string;
}

export interface XlsxDiagnostic {
  type: "load-start" | "load-deferred" | "load-resumed" | "load-success" | "load-error" | "load-cancelled" | "download";
  requestId: number;
  /** Globally unique task identity supplied by the instance loading boundary. */
  taskId?: string;
  sourceKind?: XlsxSourceKind;
  url?: string;
  bytes?: number;
  error?: XlsxLoadError;
}

export interface UseXlsxViewerControllerOptions {
  /** Explicit immutable runtime configuration shared with this controller. */
  runtime?: XlsxRuntimeLike;
  /**
   * Allows row and column resizing even while editing is disabled by `readOnly`.
   *
   * @default false
   */
  allowResizeInReadOnly?: boolean;
  /**
   * Defers loading until `continueDeferredLoad()` is called when the file is larger than this byte threshold.
   * Set to `0` to parse immediately.
   *
   * @default 0
   * @example
   * ```tsx
   * useXlsxViewerController({ file, deferLoadingAboveBytes: 10 * 1024 * 1024 })
   * ```
   */
  deferLoadingAboveBytes?: number;
  /**
   * Local workbook bytes to load. Use either `file` or `src`.
   *
   * @example
   * ```tsx
   * <XlsxViewer file={await uploadedFile.arrayBuffer()} fileName={uploadedFile.name} />
   * ```
   */
  file?: ArrayBuffer;
  /**
   * Optional display and download name for the workbook.
   *
   * @example
   * ```tsx
   * <XlsxViewer file={buffer} fileName="forecast.xlsx" />
   * ```
   */
  fileName?: string;
  /**
   * Rejects files larger than this byte limit and renders `fileTooLargeState`.
   *
   * @default 25 * 1024 * 1024
   */
  maxFileSizeBytes?: number;
  /**
   * Disables workbook edits, paste, fill, undo/redo, and other mutation actions.
   *
   * @default false
   */
  readOnly?: boolean;
  /**
   * Automatically enables read-only mode for files larger than this byte threshold.
   * Set to `0` to disable the automatic switch.
   *
   * @default 0
   */
  readOnlyAboveBytes?: number;
  /**
   * Includes hidden and very hidden workbook sheets in the sheet tabs and controller state.
   *
   * @default false
   */
  showHiddenSheets?: boolean;
  /**
   * Skips OOXML ZIP/XML parsing and relies on `@dukelib/sheets-wasm` metadata only.
   * This is automatically used for legacy `.xls` files.
   *
   * @default false
   */
  skipXmlParsing?: boolean;
  /**
   * Remote workbook URL to fetch and load. Use either `src` or `file`.
   *
   * @example
   * ```tsx
   * <XlsxViewer src="/workbooks/report.xlsx" />
   * ```
   */
  src?: string;
  /**
   * Explicit URL rules required for URL sources. Local ArrayBuffer sources do
   * not perform a network request.
   */
  urlPolicy?: XlsxUrlPolicy;
  /** Receives public loading and download diagnostics for this controller instance. */
  onDiagnostic?: (diagnostic: XlsxDiagnostic) => void;
  /**
   * Creates the Worker owned by this controller. Use this when a host bundler
   * exposes the package Worker through its public URL import.
   */
  createWorker?: () => Worker;
  /**
   * Parses supported workbook data in a Web Worker so large files do not block React rendering.
   *
   * @default true
   */
  useWorker?: boolean;
  /**
   * Public module Worker URL for this controller. It is ignored when a Worker
   * factory is supplied.
   */
  workerUrl?: string;
}

export interface XlsxViewerController {
  activeCell: XlsxCellAddress | null;
  activeCellAddress: string | null;
  activeSheet: XlsxSheetData | null;
  activeSheetIndex: number;
  activeTab: XlsxWorkbookTab | null;
  activeTabIndex: number;
  canDownload: boolean;
  canExport: boolean;
  canLoadDeferred: boolean;
  canRedo: boolean;
  canUndo: boolean;
  canZoomIn: boolean;
  canZoomOut: boolean;
  clearSelectedCells: () => void;
  clearSelection: () => void;
  continueDeferredLoad: () => void;
  copySelectionToClipboard: () => Promise<boolean>;
  defaultZoomScale: number;
  deferredLoadFileSize: number | null;
  defineNamedRange: (name: string, range?: XlsxCellRange | null) => void;
  displayFileName: string;
  download: () => void;
  charts: XlsxChart[];
  chartsheets: XlsxChartsheet[];
  exportCsv: () => void;
  exportXlsx: () => void;
  error: Error | null;
  file?: ArrayBuffer;
  fillSelection: (targetRange: XlsxCellRange) => void;
  clearSelectedChart: () => void;
  clearSelectedChartElement: () => void;
  clearSelectedImage: () => void;
  getChartById: (id: string) => XlsxChart | null;
  getChartSeriesFormula: (chartId: string, seriesIndex: number) => string;
  getChartsheetById: (id: string) => XlsxChartsheet | null;
  formControls: XlsxFormControl[];
  getSheetCharts: (sheetIndex?: number) => XlsxChart[];
  getSheetFormControls: (sheetIndex?: number) => XlsxFormControl[];
  getImageById: (id: string) => XlsxImage | null;
  getSheetImages: (sheetIndex?: number) => XlsxImage[];
  getSheetShapes: (sheetIndex?: number) => XlsxShape[];
  getClipboardData: () => XlsxClipboardData | null;
  getCellDisplayValue: (cell?: XlsxCellAddress | null) => string;
  getCellFormula: (cell?: XlsxCellAddress | null) => string;
  getCellSnapshotAsync?: (workbookSheetIndex: number, row: number, col: number, signal?: AbortSignal) => Promise<{
    displayValue: string;
    formula: string;
  }>;
  isLoadDeferred: boolean;
  isLoading: boolean;
  isChartsLoading: boolean;
  isWorkerBacked?: boolean;
  sourceState: XlsxSourceState;
  sourceError: XlsxLoadError | null;
  images: XlsxImage[];
  moveChartBy: (id: string, deltaX: number, deltaY: number) => void;
  shapes: XlsxShape[];
  mergeSelection: () => void;
  maxZoomScale: number;
  minZoomScale: number;
  moveImageBy: (id: string, deltaX: number, deltaY: number) => void;
  removeActiveSheet: () => void;
  readOnly: boolean;
  recalculate: () => void;
  revision: number;
  resetZoom: () => void;
  resizeChartBy: (
    id: string,
    handle: XlsxImageResizeHandlePosition,
    deltaX: number,
    deltaY: number
  ) => void;
  resizeImageBy: (
    id: string,
    handle: XlsxImageResizeHandlePosition,
    deltaX: number,
    deltaY: number
  ) => void;
  resizeColumn: (col: number, widthPx: number) => void;
  resizeRow: (row: number, heightPx: number) => void;
  redo: () => void;
  pasteFromClipboard: () => Promise<boolean>;
  pasteStructuredClipboardData: (payload: string) => boolean;
  pasteText: (text: string) => boolean;
  selectedRangeAddress: string | null;
  selectedValue: string;
  selectedFormula: string;
  /**
   * Sets the formula for a cell in the active worksheet.
   *
   * Pass an A1-style formula body or formula string, for example `"SUM(A1:A3)"`
   * or `"=SUM(A1:A3)"`. The mutation is recorded in undo history and is included
   * in exported workbook bytes.
   *
   * @param cell Zero-based row and column address in the active worksheet.
   * @param formula Formula text to write. Pass an empty string to clear the formula.
   */
  setCellFormula: (cell: XlsxCellAddress, formula: string) => void;
  /**
   * Applies persisted Excel style properties to a cell in the active worksheet.
   *
   * This mutates workbook data, records undo history, refreshes the viewer, and
   * is included in exported XLSX bytes. Use the `getCellStyle` viewer prop only
   * for render-time CSS overrides that should not modify the workbook.
   *
   * @param cell Zero-based row and column address in the active worksheet.
   * @param style Partial Excel style patch to apply to the target cell.
   */
  setCellStyle: (cell: XlsxCellAddress, style: XlsxCellStyleInput) => void;
  /**
   * Sets the value for a cell in the active worksheet.
   *
   * User-entered strings are coerced to booleans or numbers when they match
   * Excel-like input. Prefix text with an apostrophe to force literal text.
   *
   * @param cell Zero-based row and column address in the active worksheet.
   * @param value Text entered by the user.
   */
  setCellValue: (cell: XlsxCellAddress, value: string) => void;
  /**
   * Applies persisted Excel style properties to every cell in a range.
   *
   * This mutates workbook data, records undo history, refreshes the viewer, and
   * is included in exported XLSX bytes.
   *
   * @param range Zero-based cell range in the active worksheet.
   * @param style Partial Excel style patch to apply to the target range.
   */
  setRangeStyle: (range: XlsxCellRange, style: XlsxCellStyleInput) => void;
  setZoomScale: (zoomScale: number) => void;
  selectCell: (cell: XlsxCellAddress, options?: { extend?: boolean }) => void;
  selectChart: (id: string | null) => void;
  selectChartElement: (selection: XlsxChartElementSelection | null) => void;
  selectRange: (range: XlsxCellRange) => void;
  selection: XlsxCellRange | null;
  setActiveSheetIndex: (index: number) => void;
  setActiveTabIndex: (index: number) => void;
  selectedChart: XlsxChart | null;
  selectedChartElement: XlsxChartElementSelection | null;
  selectedChartFormula: string | null;
  selectedChartId: string | null;
  selectedCellFormula: string;
  selectedFormulaTarget: XlsxFormulaTarget;
  selectedImage: XlsxImage | null;
  selectedImageId: string | null;
  setSelectedFormula: (formula: string) => boolean;
  setSelectedCellFormula: (formula: string) => void;
  /**
   * Applies persisted Excel style properties to the active cell.
   *
   * No-ops when there is no active cell, when editing is disabled, or when no
   * workbook is loaded. The mutation is recorded in undo history and is included
   * in exported XLSX bytes.
   *
   * @param style Partial Excel style patch to apply to the active cell.
   */
  setSelectedCellStyle: (style: XlsxCellStyleInput) => void;
  setSelectedCellValue: (value: string) => void;
  sheets: XlsxSheetData[];
  src?: string;
  sortState: XlsxTableSortState | null;
  sortTable: (tableName: string, columnIndex: number, direction: XlsxTableSortDirection) => void;
  setChartSeriesFormula: (chartId: string, seriesIndex: number, formula: string) => boolean;
  selectImage: (id: string | null) => void;
  setChartRect: (id: string, rect: XlsxImageRect) => void;
  setImageRect: (id: string, rect: XlsxImageRect) => void;
  getRowsBatchAsync?: (workbookSheetIndex: number, startRow: number, rowCount: number, signal?: AbortSignal) => Promise<unknown[] | null>;
  tables: XlsxTable[];
  tabs: XlsxWorkbookTab[];
  undo: () => void;
  unmergeSelection: () => void;
  updateChart: (id: string, patch: Partial<XlsxChart>) => void;
  workbook: Workbook | null;
  zoomIn: () => void;
  zoomOut: () => void;
  zoomScale: number;
  getActiveWorksheet: () => Worksheet | null;
  addSheet: (name?: string) => void;
}

// ── Sub-controller interfaces ────────────────────────────────────────────

export interface XlsxViewerSelection {
  activeCell: XlsxCellAddress | null;
  activeCellAddress: string | null;
  clearSelection: () => void;
  selectedRangeAddress: string | null;
  selectCell: (cell: XlsxCellAddress, options?: { extend?: boolean }) => void;
  selectRange: (range: XlsxCellRange) => void;
  selection: XlsxCellRange | null;
}

export interface XlsxViewerZoom {
  canZoomIn: boolean;
  canZoomOut: boolean;
  defaultZoomScale: number;
  maxZoomScale: number;
  minZoomScale: number;
  resetZoom: () => void;
  setZoomScale: (zoomScale: number) => void;
  zoomIn: () => void;
  zoomOut: () => void;
  zoomScale: number;
}

export interface XlsxViewerEditing {
  addSheet: (name?: string) => void;
  canRedo: boolean;
  canUndo: boolean;
  clearSelectedCells: () => void;
  copySelectionToClipboard: () => Promise<boolean>;
  defineNamedRange: (name: string, range?: XlsxCellRange | null) => void;
  fillSelection: (targetRange: XlsxCellRange) => void;
  getClipboardData: () => XlsxClipboardData | null;
  getCellDisplayValue: (cell?: XlsxCellAddress | null) => string;
  getCellFormula: (cell?: XlsxCellAddress | null) => string;
  mergeSelection: () => void;
  pasteFromClipboard: () => Promise<boolean>;
  pasteStructuredClipboardData: (payload: string) => boolean;
  pasteText: (text: string) => boolean;
  removeActiveSheet: () => void;
  readOnly: boolean;
  redo: () => void;
  selectedCellFormula: string;
  selectedChartFormula: string | null;
  selectedFormula: string;
  selectedFormulaTarget: XlsxFormulaTarget;
  selectedValue: string;
  /**
   * Sets the formula for a cell in the active worksheet.
   *
   * @param cell Zero-based row and column address in the active worksheet.
   * @param formula Formula text to write. Pass an empty string to clear the formula.
   */
  setCellFormula: (cell: XlsxCellAddress, formula: string) => void;
  /**
   * Applies persisted Excel style properties to a cell in the active worksheet.
   *
   * The mutation is recorded in undo history and is included in exported XLSX
   * bytes. This is different from the `getCellStyle` render override prop.
   *
   * @param cell Zero-based row and column address in the active worksheet.
   * @param style Partial Excel style patch to apply to the target cell.
   */
  setCellStyle: (cell: XlsxCellAddress, style: XlsxCellStyleInput) => void;
  /**
   * Sets the value for a cell in the active worksheet.
   *
   * @param cell Zero-based row and column address in the active worksheet.
   * @param value Text entered by the user.
   */
  setCellValue: (cell: XlsxCellAddress, value: string) => void;
  /**
   * Applies persisted Excel style properties to every cell in a range.
   *
   * @param range Zero-based cell range in the active worksheet.
   * @param style Partial Excel style patch to apply to the target range.
   */
  setRangeStyle: (range: XlsxCellRange, style: XlsxCellStyleInput) => void;
  setSelectedFormula: (formula: string) => boolean;
  setSelectedCellFormula: (formula: string) => void;
  /**
   * Applies persisted Excel style properties to the active cell.
   *
   * @param style Partial Excel style patch to apply to the active cell.
   */
  setSelectedCellStyle: (style: XlsxCellStyleInput) => void;
  setSelectedCellValue: (value: string) => void;
  undo: () => void;
  unmergeSelection: () => void;
}

export interface XlsxViewerTables {
  sortState: XlsxTableSortState | null;
  sortTable: (tableName: string, columnIndex: number, direction: XlsxTableSortDirection) => void;
  tables: XlsxTable[];
}

export interface XlsxViewerImages {
  charts: XlsxChart[];
  clearSelectedChart: () => void;
  clearSelectedChartElement: () => void;
  clearSelectedImage: () => void;
  getChartById: (id: string) => XlsxChart | null;
  getChartSeriesFormula: (chartId: string, seriesIndex: number) => string;
  getSheetCharts: (sheetIndex?: number) => XlsxChart[];
  getImageById: (id: string) => XlsxImage | null;
  getSheetImages: (sheetIndex?: number) => XlsxImage[];
  images: XlsxImage[];
  moveChartBy: (id: string, deltaX: number, deltaY: number) => void;
  moveImageBy: (id: string, deltaX: number, deltaY: number) => void;
  isChartsLoading: boolean;
  readOnly: boolean;
  resizeChartBy: (
    id: string,
    handle: XlsxImageResizeHandlePosition,
    deltaX: number,
    deltaY: number
  ) => void;
  resizeImageBy: (
    id: string,
    handle: XlsxImageResizeHandlePosition,
    deltaX: number,
    deltaY: number
  ) => void;
  selectedChart: XlsxChart | null;
  selectedChartElement: XlsxChartElementSelection | null;
  selectedChartFormula: string | null;
  selectedChartId: string | null;
  selectedImage: XlsxImage | null;
  selectedImageId: string | null;
  selectChart: (id: string | null) => void;
  selectChartElement: (selection: XlsxChartElementSelection | null) => void;
  selectImage: (id: string | null) => void;
  setChartSeriesFormula: (chartId: string, seriesIndex: number, formula: string) => boolean;
  setChartRect: (id: string, rect: XlsxImageRect) => void;
  setImageRect: (id: string, rect: XlsxImageRect) => void;
  updateChart: (id: string, patch: Partial<XlsxChart>) => void;
}

export interface XlsxViewerCharts {
  activeTab: XlsxWorkbookTab | null;
  activeTabIndex: number;
  charts: XlsxChart[];
  chartsheets: XlsxChartsheet[];
  clearSelectedChart: () => void;
  clearSelectedChartElement: () => void;
  getChartById: (id: string) => XlsxChart | null;
  getChartSeriesFormula: (chartId: string, seriesIndex: number) => string;
  getChartsheetById: (id: string) => XlsxChartsheet | null;
  getSheetCharts: (sheetIndex?: number) => XlsxChart[];
  isChartsLoading: boolean;
  moveChartBy: (id: string, deltaX: number, deltaY: number) => void;
  readOnly: boolean;
  resizeChartBy: (
    id: string,
    handle: XlsxImageResizeHandlePosition,
    deltaX: number,
    deltaY: number
  ) => void;
  selectChart: (id: string | null) => void;
  selectedChart: XlsxChart | null;
  selectedChartElement: XlsxChartElementSelection | null;
  selectedChartFormula: string | null;
  selectedChartId: string | null;
  selectChartElement: (selection: XlsxChartElementSelection | null) => void;
  setChartSeriesFormula: (chartId: string, seriesIndex: number, formula: string) => boolean;
  setActiveTabIndex: (index: number) => void;
  setChartRect: (id: string, rect: XlsxImageRect) => void;
  tabs: XlsxWorkbookTab[];
  updateChart: (id: string, patch: Partial<XlsxChart>) => void;
}

// ── Thumbnails ───────────────────────────────────────────────────────────

export type XlsxSheetThumbnailResolution =
  | number
  | {
      maxHeight?: number;
      maxWidth?: number;
    };

export interface UseXlsxViewerThumbnailsOptions {
  includeHeaders?: boolean;
  resolution?: XlsxSheetThumbnailResolution;
}

export interface XlsxSheetThumbnail {
  aspectRatio: number;
  contentHeight: number;
  contentWidth: number;
  height: number;
  paint: (canvas: HTMLCanvasElement | null) => boolean;
  sheet: XlsxSheetData;
  sheetIndex: number;
  sourceRange: XlsxCellRange;
  width: number;
  workbookSheetIndex: number;
}

export interface XlsxViewerThumbnails {
  paintThumbnail: (sheetIndex: number, canvas: HTMLCanvasElement | null) => boolean;
  thumbnails: XlsxSheetThumbnail[];
}

// ── Render props ─────────────────────────────────────────────────────────

export interface XlsxFileTooLargeRenderProps {
  /** The built-in file-too-large message. */
  defaultNode: unknown;
  /** File name displayed in the viewer UI. */
  displayFileName: string;
  /** Actual file size in bytes. */
  fileSizeBytes: number;
  /** Configured maximum file size in bytes. */
  maxFileSizeBytes: number;
}

export interface XlsxScrollerRenderProps {
  /** Workbook grid content that should be rendered inside the scrollable viewport. */
  children: unknown;
  /** Props that must be applied to the actual scrollable viewport element. */
  viewportProps: Record<string, unknown> & {
    ref: unknown;
    style: Record<string, string | number | undefined>;
    tabIndex: number;
  };
}
