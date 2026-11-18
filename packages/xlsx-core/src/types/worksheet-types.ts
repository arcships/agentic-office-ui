import type { Workbook, Worksheet } from "@dukelib/sheets-wasm";
import type { XlsxChart, XlsxChartElementSelection, XlsxChartFamily, XlsxChartsheet } from "./chart-types";
import type { XlsxThemePalette } from "./theme-types";
import type {
  XlsxFormControl,
  XlsxFormControlInput,
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

/**
 * Reason a range is highlighted.
 *
 * `category`, `value` and `bubbleSize` are derived from a selected chart
 * element, so the viewer can paint the chart's data sources with a colour per
 * role. `seriesName` is reserved for chart name references and `custom` marks
 * ranges a host highlighted through `XlsxViewerController.highlightRanges`.
 */
export type XlsxRangeHighlightRole = "category" | "value" | "bubbleSize" | "seriesName" | "custom";

/**
 * One range painted by the viewer's range-highlight layer.
 *
 * Ranges are always normalized (`start` is the top-left cell) and refer to a
 * visible worksheet index, matching the `sheetIndex` used by charts and
 * images. The controller returns a fresh, read-only snapshot from
 * `highlightedRanges`; calling `highlightRanges` again replaces the ranges a
 * host supplied, while chart source ranges follow the chart selection.
 */
export interface XlsxRangeHighlight {
  /** CSS colour used for the translucent fill and its outline. */
  color: string;
  /** Optional stable identity for host diffing. */
  id?: string;
  /** Optional description a host can show next to the highlight. */
  label?: string;
  /** Normalized zero-based range inside `sheetIndex`. */
  range: XlsxCellRange;
  /** Why the range is highlighted. */
  role?: XlsxRangeHighlightRole;
  /** Visible worksheet index the range belongs to. */
  sheetIndex: number;
  /** Zero-based chart series index when the range feeds a chart series. */
  seriesIndex?: number;
  /** Chart series name, so a highlighted range can be tied to its legend entry. */
  seriesName?: string;
}

/** Options for {@link XlsxViewerController.highlightRanges}. */
export interface XlsxHighlightRangesOptions {
  /** CSS colour shared by every range in the call; defaults to amber. */
  color?: string;
  /** Optional stable id shared by the highlighted ranges. */
  id?: string;
  /** Optional description copied to every highlighted range. */
  label?: string;
  /** Defaults to `"custom"`. */
  role?: XlsxRangeHighlightRole;
  /** Visible worksheet index; defaults to the active sheet. */
  sheetIndex?: number;
}

export interface XlsxCellHyperlink {
  cell: XlsxCellAddress;
  target: string;
  tooltip?: string;
}

export interface XlsxCellComment {
  author?: string;
  cell: XlsxCellAddress;
  text: string;
}

/**
 * Formatting of one inline run inside a rich-text cell.
 *
 * The fields mirror the OOXML `rPr` element: they are per-run overrides, so a
 * field that is absent inherits from the cell style.
 */
export interface XlsxCellTextRunFont {
  bold?: boolean;
  /** Raw spreadsheet color record; resolve with `resolveWorkbookColor`. */
  color?: Record<string, unknown>;
  italic?: boolean;
  /** Font family name from `rFont`/`name`, when the run overrides it. */
  name?: string;
  /** Font size in points, when the run overrides the cell size. */
  sizePt?: number;
  strikethrough?: boolean;
  underline?: string;
  /** OOXML `vertAlign`: `superscript`, `subscript` or `baseline`. */
  verticalAlign?: string;
}

/**
 * One inline run of a rich-text cell.
 *
 * A cell that stores several runs with their own fonts (for example a bold
 * title with a red word and a superscript footnote marker) keeps them here so
 * renderers can draw each run with its own formatting. The cell's plain
 * `value` stays the concatenation of every run and remains the text used for
 * formulas, search and clipboard.
 */
export interface XlsxCellTextRun {
  text: string;
  /** Per-run formatting; absent means the run uses the cell style as-is. */
  font?: XlsxCellTextRunFont;
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
  /** Raw cfvo type: min, max, num, percent, percentile, formula, autoMin/autoMax. */
  type: string;
  /** Numeric bound, or the raw formula text for `formula` cfvos. */
  value?: number | string;
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
  errorTitle?: string;
  inputMessage?: string;
  inputTitle?: string;
  listSource?: string;
  operator?: string;
  ranges: XlsxCellRange[];
  showDropdown?: boolean;
  showErrorAlert?: boolean;
  showInputMessage?: boolean;
  validationType: string;
  value1?: string;
  value2?: string;
}

/** Built-in validation message keys that hosts can localize. */
export type XlsxDataValidationMessageKey =
  | "list"
  | "whole"
  | "decimal"
  | "textLength"
  | "date"
  | "time"
  | "generic";

/** Feedback produced when user input violates a data validation rule. */
export interface XlsxDataValidationFeedback {
  cell: XlsxCellAddress;
  kind: "error" | "information" | "warning";
  message: string;
  /** Set for built-in messages; hosts may localize on this key. */
  defaultMessageKey?: XlsxDataValidationMessageKey;
  title?: string;
}

export type XlsxSheetProtectionAction =
  | "autoFilter"
  | "deleteColumns"
  | "deleteRows"
  | "formatCells"
  | "formatColumns"
  | "formatRows"
  | "insertColumns"
  | "insertHyperlinks"
  | "insertRows"
  | "objects"
  | "pivotTables"
  | "scenarios"
  | "selectLockedCells"
  | "selectUnlockedCells"
  | "sort";

/**
 * Sheet protection parsed from the worksheet `sheetProtection` record.
 *
 * `flags` mirrors the protection state reported by the workbook engine for
 * each OOXML attribute. Treat `enabled` as the authoritative signal for
 * read-only behaviour and read individual flags only when a host needs the
 * finer-grained Excel permissions.
 */
export interface XlsxSheetProtection {
  enabled: boolean;
  flags: Partial<Record<XlsxSheetProtectionAction, boolean>>;
}

export interface XlsxFreezePanes {
  col: number;
  row: number;
}

export type XlsxSheetVisibility = "hidden" | "veryHidden" | "visible";

// ── Print / page layout ──────────────────────────────────────────────────

export type XlsxPageOrientation = "portrait" | "landscape";

export type XlsxPageOrder = "downThenOver" | "overThenDown";

/**
 * Worksheet print set-up parsed from `pageSetup`, `sheetPr/pageSetUpPr` and
 * `pageMargins`.
 *
 * Lengths are in inches, matching OOXML. `scale` is a percentage in the
 * 10–400 range Excel accepts. `fitToPage` mirrors
 * `sheetPr/pageSetUpPr/@fitToPage`; when it is `true`, `fitToWidth` and
 * `fitToHeight` drive the effective scale instead of `scale`.
 */
export interface XlsxPageSetup {
  paperSize: number;
  orientation: XlsxPageOrientation;
  scale: number;
  fitToWidth: number;
  fitToHeight: number;
  fitToPage: boolean;
  firstPageNumber: number;
  useFirstPageNumber: boolean;
  blackAndWhite: boolean;
  draft: boolean;
  cellComments: "none" | "asDisplayed" | "atEnd";
  errors: "displayed" | "blank" | "dash" | "NA";
  pageOrder: XlsxPageOrder;
  horizontalDpi?: number;
  verticalDpi?: number;
  copies?: number;
}

/** Worksheet `printOptions` record. */
export interface XlsxPrintOptions {
  horizontalCentered: boolean;
  verticalCentered: boolean;
  headings: boolean;
  gridLines: boolean;
}

/** Worksheet `pageMargins` record; every length is in inches. */
export interface XlsxPageMargins {
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

/**
 * A manual page break from `rowBreaks` / `colBreaks`.
 *
 * `id` is the zero-based row or column index before which the next page
 * starts, matching the OOXML `brk/@id` attribute (a break before Excel row 2
 * is written as `id="1"`).
 */
export interface XlsxPageBreak {
  id: number;
  min: number;
  max: number;
  manual: boolean;
}

/** Inclusive zero-based range on one worksheet axis. */
export interface XlsxAxisRange {
  start: number;
  end: number;
}

/**
 * Repeated rows and columns from `_xlnm.Print_Titles`, applied to every
 * printed page.
 */
export interface XlsxPrintTitles {
  rows: XlsxAxisRange[];
  columns: XlsxAxisRange[];
}

/** A paper size resolved from the OOXML `paperSize` identifier. */
export interface XlsxPaperSize {
  id: number;
  name: string;
  widthIn: number;
  heightIn: number;
}

/** One printed page: a cell range on a sheet plus its paper geometry. */
export interface XlsxPrintPage {
  /** Zero-based index in the sheet's flat page list. */
  index: number;
  /** One-based page number shown to the user; honours `firstPageNumber`. */
  pageNumber: number;
  /** Cell range printed as this page's body, excluding repeated titles. */
  range: XlsxCellRange;
  /** Print area this page belongs to. */
  printArea: XlsxCellRange;
  paperSize: number;
  paperSizeName: string;
  orientation: XlsxPageOrientation;
  /** Paper width and height in CSS pixels at 96 dpi. */
  widthPx: number;
  heightPx: number;
  /** Printable box after page margins, in CSS pixels. */
  contentWidthPx: number;
  contentHeightPx: number;
  /** Effective print scale as a percentage. */
  scale: number;
  /** Zero-based column page index inside the print area. */
  colPage: number;
  /** Zero-based row page index inside the print area. */
  rowPage: number;
  /** Number of spreadsheet columns/rows in `range`. */
  colCount: number;
  rowCount: number;
}

/** Flat pagination result for one worksheet. */
export interface XlsxPrintLayout {
  workbookSheetIndex: number;
  sheetName: string;
  pages: XlsxPrintPage[];
  pageSetup: XlsxPageSetup;
  printOptions: XlsxPrintOptions;
  pageMargins: XlsxPageMargins;
  printAreas: XlsxCellRange[];
  printTitles: XlsxPrintTitles;
  paperSize: XlsxPaperSize;
  /** Effective print scale as a percentage. */
  scale: number;
  fitToPage: boolean;
  /** `true` when a host-provided page cap stopped pagination early. */
  truncated: boolean;
}

/** Options accepted by `buildWorksheetPrintLayout`. */
export interface XlsxPrintLayoutOptions {
  /** Maximum number of pages to generate. Defaults to 5000. */
  maxPages?: number;
}

/** One cell slot rendered by `buildPrintPageGrid`. */
export interface XlsxPrintPageGridCell {
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
  /** `true` when a merge anchored earlier on this page covers the slot. */
  covered: boolean;
}

/** Row or column slot on a printed page. */
export interface XlsxPrintPageGridAxis {
  index: number;
  /** Unscaled size in CSS pixels at 100%. */
  sizePx: number;
}

/** Cell layout for one printed page, including repeated titles and merges. */
export interface XlsxPrintPageGrid {
  rows: XlsxPrintPageGridAxis[];
  columns: XlsxPrintPageGridAxis[];
  cells: XlsxPrintPageGridCell[];
}

/**
 * Why Excel would reject a worksheet name.
 *
 * `empty` means the trimmed name has no characters, `too-long` means it
 * exceeds 31 code points, `invalid-characters` covers `\ / : * ? [ ]`,
 * `apostrophe` covers a leading or trailing apostrophe, and `duplicate` means
 * another worksheet or chartsheet already uses the name (case-insensitively).
 */
export type XlsxSheetNameIssue =
  | "apostrophe"
  | "duplicate"
  | "empty"
  | "invalid-characters"
  | "too-long";

// ── Sheet data ───────────────────────────────────────────────────────────

export interface XlsxSheetData {
  cachedFormulaValues: Record<string, string>;
  /**
<<<<<<< HEAD
   * Rotation and shrink-to-fit alignment of cells whose style overrides the
   * row/column alignment, keyed by A1 address.
   *
   * The values are the raw OOXML alignment flags (`textRotation` 0–180 and
   * `shrinkToFit`), so both main-thread and Worker-backed renderers can lay
   * out rotated headers and long values without reading the workbook again.
   */
  cellTextLayouts?: Record<string, { shrinkToFit?: boolean; textRotation?: number }>;
=======
   * AutoFilter ranges of the sheet, composed from the worksheet's own
   * `<autoFilter>` and every table part's filter. Empty when the sheet has no
   * filter or when the viewer skipped XML parsing.
   */
  autoFilterRanges?: XlsxAutoFilterRange[];
>>>>>>> fix/133
  /**
   * Rich-text runs of every cell that stores more than one formatted run,
   * keyed by A1 address. Cells without an entry render their plain `value`.
   */
  richTextCells?: Record<string, XlsxCellTextRun[]>;
  colWidthOverridesPx: Record<number, number>;
  colStyleIds: Record<number, number>;
  hiddenCols?: number[];
  hiddenRows?: number[];
  conditionalFormatRules: XlsxConditionalFormatRule[];
  hyperlinks: XlsxCellHyperlink[];
  comments: XlsxCellComment[];
  dataValidations: XlsxDataValidation[];
  name: string;
  visibility: XlsxSheetVisibility;
  /** Parsed sheet protection; `null` or omitted for unprotected sheets. */
  protection?: XlsxSheetProtection | null;
  /** Page size, orientation and scaling from `pageSetup`. */
  pageSetup: XlsxPageSetup;
  /** Print centering, headings and gridline flags from `printOptions`. */
  printOptions: XlsxPrintOptions;
  /** Page margins in inches from `pageMargins`. */
  pageMargins: XlsxPageMargins;
  /** Manual row breaks from `rowBreaks`. */
  rowBreaks: XlsxPageBreak[];
  /** Manual column breaks from `colBreaks`. */
  colBreaks: XlsxPageBreak[];
  /** Areas from `_xlnm.Print_Area`; empty means the whole used range prints. */
  printAreas: XlsxCellRange[];
  /** Repeated rows/columns from `_xlnm.Print_Titles`. */
  printTitles: XlsxPrintTitles;
  columnWidthCharacterWidthPx?: number;
  defaultColWidthPx: number;
  defaultRowHeightPx: number;
  freezePanes: XlsxFreezePanes | null;
  hasHorizontalMerges: boolean;
  hasVerticalMerges: boolean;
  maxHorizontalMergeEndCol: number;
  maxVerticalMergeEndRow: number;
  /** Exact merged ranges used by Worker-backed viewers. */
  mergedRegions?: XlsxCellRange[];
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

/**
 * Border presets the ribbon offers for the current selection.
 *
 * `"outside"` draws the perimeter of a multi-cell selection only, so the
 * inside of the range keeps whatever borders it already had.
 */
export type XlsxBorderPreset =
  | "all"
  | "outside"
  | "left"
  | "right"
  | "top"
  | "bottom"
  | "none";

/** Selection border command: which edges to draw, with what line and colour. */
export interface XlsxBorderPresetInput {
  /** Edges the command draws; defaults to `"all"`. */
  preset?: XlsxBorderPreset;
  /** Excel border line style; defaults to `"thin"`. */
  style?: XlsxCellBorderEdgeInput["style"];
  /** Border line colour; defaults to the ribbon's grey (`A1A1AA`). */
  color?: XlsxCellStyleColorInput;
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

/**
 * A cut waiting for its paste: the source range that still has to be cleared
 * once the payload lands somewhere else.
 *
 * The sheet index is stored alongside the range because a host may switch
 * sheets between the cut and the paste; the marquee belongs to the source
 * sheet, and the move clears it there.
 */
export interface XlsxClipboardCutState {
  /** Absolute (zero-based) source range of the pending move. */
  range: XlsxCellRange;
  /** Zero-based workbook sheet index of the cut source. */
  sheetIndex: number;
}

// ── Tables ───────────────────────────────────────────────────────────────

export interface XlsxTableColumn {
  id: number;
  index: number;
  name: string;
}

/**
 * One `<filterColumn>` of an AutoFilter definition.
 *
 * `columnIndex` is relative to the filter range's first column, matching the
 * `colId` attribute in the sheet XML. `values` holds the allowed display
 * values of a value filter; `includeBlank` keeps empty cells visible.
 */
export interface XlsxAutoFilterColumn {
  columnIndex: number;
  includeBlank: boolean;
  showButton: boolean;
  values: string[];
}

/**
 * A parsed `<autoFilter>` definition from a worksheet or a table part.
 *
 * Excel stores the filter range next to the criteria that created it; the
 * viewer turns each definition into a header menu and a view-level filter.
 */
export interface XlsxAutoFilterRange {
  /** Filter columns declared by the definition, keyed by relative column. */
  columns: XlsxAutoFilterColumn[];
  /** Last data row, already excluding a table's totals row. */
  dataEndRow: number;
  end: XlsxCellAddress;
  /** Header rows inside the range; worksheet filters always have one. */
  headerRowCount: number;
  reference: string;
  source: "table" | "worksheet";
  start: XlsxCellAddress;
  /** Owning table name when `source` is `table`. */
  tableName?: string;
}

/** A filter value offered by an AutoFilter header menu. */
export interface XlsxAutoFilterValueOption {
  blank: boolean;
  label: string;
  value: string;
}

/** One active column filter of an AutoFilter range. */
export interface XlsxAutoFilterColumnState {
  columnIndex: number;
  includeBlank: boolean;
  values: string[];
}

/** Active view-level filter state of one AutoFilter range. */
export interface XlsxAutoFilterState {
  columns: XlsxAutoFilterColumnState[];
  /** Absolute worksheet rows hidden by this range's filters. */
  hiddenRows: number[];
  reference: string;
}

/** One data row of an AutoFilter range, as read by the Worker viewer. */
export interface XlsxAutoFilterRowData {
  row: number;
  values: string[];
}

/**
 * Display values of an AutoFilter range.
 *
 * Worker-backed (read-only) viewers have no main-thread workbook, so the Worker
 * returns the formatted values the header menu needs to list and filter.
 */
export interface XlsxAutoFilterSnapshot {
  rows: XlsxAutoFilterRowData[];
  truncated: boolean;
}

export interface XlsxTableStyleInfo {
  name?: string;
  showColumnStripes?: boolean;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
}

export interface XlsxTable {
  /** AutoFilter declared by the table part, when it has one. */
  autoFilter?: XlsxAutoFilterRange;
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
  /** Table name, or the AutoFilter range reference for a worksheet filter. */
  tableName: string;
  /** Set when the sort came from an AutoFilter header menu. */
  rangeReference?: string;
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

/**
 * A worksheet the user can bring back with "Unhide".
 *
 * `veryHidden` worksheets are deliberately absent from this list: Excel does
 * not show them in its Unhide dialog either, so they can only be restored
 * through `setSheetVisibility`. `getSheetVisibility` still reports them.
 */
export interface XlsxHiddenSheet {
  name: string;
  visibility: "hidden";
  workbookSheetIndex: number;
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
  /** The AutoFilter range backing this header, when the sheet has one. */
  autoFilter?: XlsxAutoFilterRange;
  /** Applies a value filter to this AutoFilter column. */
  applyFilter?: (values: string[], options?: { includeBlank?: boolean }) => void;
  /** Removes this range's filter and restores its hidden rows. */
  clearFilter?: () => void;
  /** Current sort direction for this column, or `null` when unsorted. */
  direction: XlsxTableSortDirection | null;
  /** Filter values currently selected for this column, if any. */
  filterValues?: string[];
  /** Resolves the distinct display values of the column and the blanks flag. */
  getValueOptions?: () => XlsxAutoFilterValueOption[];
  /** Whether this range currently hides rows. */
  hasFilter?: boolean;
  /** Whether this column has an active value filter. */
  hasValueFilter?: boolean;
  /** Sorts the table by this column in ascending order. */
  sortAscending: () => void;
  /** Sorts the table by this column in descending order. */
  sortDescending: () => void;
  /** Table metadata for the active header; absent for a worksheet filter. */
  table: XlsxTable | null;
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
  readonly limits?: Readonly<{
    maxInputBytes?: number;
    maxConcurrentImageDecodes?: number;
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

export type XlsxLoadErrorCode =
  | "SOURCE_NOT_ALLOWED"
  | "FETCH_FAILED"
  | "INVALID_WORKBOOK"
  | "LIMIT_EXCEEDED"
  | "IMAGE_LIMIT_EXCEEDED"
  | "INVALID_IMAGE"
  | "IMAGE_DECODE_FAILED"
  | "TIMEOUT"
  | "WORKER_UNAVAILABLE"
  | "ABORTED";

export interface XlsxLoadError {
  code: XlsxLoadErrorCode;
  message: string;
  sourceKind: XlsxSourceKind;
  url?: string;
  phase?: string;
  limit?: string;
  actual?: number;
  allowed?: number;
}

export interface XlsxDiagnostic {
  type: "load-start" | "load-deferred" | "load-resumed" | "load-success" | "load-error" | "load-cancelled" | "download" | "export-warning" | "worker-error" | "image-decode-error";
  requestId: number;
  /** Globally unique task identity supplied by the instance loading boundary. */
  taskId?: string;
  sourceKind?: XlsxSourceKind;
  url?: string;
  bytes?: number;
  imageId?: string;
  /** Human-readable export warning, such as macro loss on .xlsm/.xltm sources. */
  message?: string;
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
   * Allows mutations on worksheets that declare OOXML sheet protection.
   *
   * By default a protected worksheet behaves like a read-only sheet while it
   * is active. Set this to `true` when the host enforces protection itself or
   * intentionally offers an unlock path.
   *
   * @default false
   */
  allowProtectedSheetEditing?: boolean;
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
   * Local XLSX, XLS, XLSB, XLSM, XLTX, XLTM, or CSV bytes to load. Use either `file` or `src`.
   * CSV byte sources must provide a `.csv` file name so the input format can
   * be identified; UTF-8 and BOM-marked UTF-16 are supported.
   *
   * @example
   * ```tsx
   * <XlsxViewer file={await uploadedFile.arrayBuffer()} fileName={uploadedFile.name} />
   * ```
   */
  file?: ArrayBuffer;
  /**
   * Optional display and download name for the workbook. Required for local
   * CSV byte sources and must end in `.csv`.
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
   * Locale used for number and date formats that do not name one of their own
   * (`[$-407]`-style sections always win), for example `"de-DE"`. Typed values
   * and displayed values both follow it.
   *
   * @default "en-US"
   */
  locale?: string;
  /**
   * Maximum number of undo and redo entries retained by this controller
   * instance.
   *
   * @default 100
   */
  historyMaxEntries?: number;
  /**
   * Approximate byte budget for undo and redo entries retained by this
   * controller instance. The newest oversized entry remains available by
   * itself so the latest edit can still be undone.
   *
   * @default 32 * 1024 * 1024
   */
  historyMaxBytes?: number;
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
   * Includes hidden workbook sheets in the sheet tabs and controller state.
   *
   * The mode can be changed after load with
   * `XlsxViewerController.setShowHiddenSheets`. `veryHidden` sheets are never
   * listed: Excel keeps them out of the tab strip and Unhide dialog, so they
   * can only be restored through `setSheetVisibility`.
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
   * Remote XLSX, XLS, XLSB, XLSM, XLTX, XLTM, or CSV URL to fetch and load. Use either `src` or `file`.
   * CSV is detected from the final URL extension or response Content-Type.
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


// ── Authoring inputs ─────────────────────────────────────────────────────

/** Comment text a host wants to store on a cell. */
export interface XlsxCellCommentInput {
  /** Optional author name shown in the note. */
  author?: string;
  /** Comment body. Leading/trailing whitespace is trimmed. */
  text: string;
}

/** Hyperlink a host wants to store on a cell. */
export interface XlsxCellHyperlinkInput {
  /** Absolute URL, `mailto:` address or `#Sheet!A1` internal reference. */
  target: string;
  /** Optional tooltip shown on hover. */
  tooltip?: string;
}

/**
 * Host-facing description of a conditional-format rule.
 *
 * The controller fills in thresholds, colours and priorities, so a host only
 * describes what the user picked.
 */
export interface XlsxConditionalFormatDraft {
  /** Rule kind to author. */
  kind: "colorScale" | "dataBar" | "iconSet";
  /** Colour stops as `RRGGBB`/`AARRGGBB`/`#RRGGBB`; 2–3 for a scale, one for a bar. */
  colors?: string[];
  /** Built-in icon set name such as `3TrafficLights1`. */
  iconSet?: string;
  /** Target ranges; the current selection is used when omitted. */
  ranges?: XlsxCellRange[];
  reverse?: boolean;
  showValue?: boolean;
  /** Thresholds for each stop, in order. Defaults to min/percentile/max. */
  thresholds?: Array<{ type: string; value?: number | string }>;
}

/** Chart a host wants to insert from a cell range. */
export interface XlsxChartInsertInput {
  /** Chart family to generate. Defaults to `"bar"` (clustered column). */
  family?: XlsxChartFamily;
  /** First column of the range holds category labels. Defaults to `true`. */
  hasHeaderColumn?: boolean;
  /** First row of the range holds series names. Defaults to `true`. */
  hasHeaderRow?: boolean;
  /** Source range. Defaults to the current selection, then the used range. */
  range?: XlsxCellRange;
  /** Optional chart title. */
  title?: string;
}

/** ListObject table a host wants to create from a cell range. */
export interface XlsxTableCreateInput {
  /** First row of the range holds the column headers. Defaults to `true`. */
  hasHeaders?: boolean;
  /** Table name; Excel requires a unique, identifier-like name. */
  name?: string;
  /** Source range. Defaults to the current selection. */
  range?: XlsxCellRange;
  /** Built-in table style such as `TableStyleMedium2`. */
  styleName?: string;
}

/** Data-validation rule a host wants to apply to the selection. */
export interface XlsxDataValidationInput {
  allowBlank?: boolean;
  errorMessage?: string;
  inputMessage?: string;
  /** Inline list (`"a,b,c"`), range reference (`A1:A5`) or formula for list rules. */
  listSource?: string;
  operator?: string;
  /** Target ranges; the current selection is used when omitted. */
  ranges?: XlsxCellRange[];
  showDropdown?: boolean;
  validationType: "list" | "whole" | "decimal" | "date" | "textLength";
  value1?: string;
  value2?: string;
}

/**
 * A workbook or sheet-scoped defined name (Excel's Name Manager row).
 *
 * `refersTo` keeps Excel's A1 syntax, for example `Sheet1!$A$1:$C$3` or
 * `SUM(Sheet1!$A$1:$A$3)`. Sheet-scoped names set `sheetIndex` to the
 * zero-based workbook sheet index; workbook-scoped names omit it.
 */
export interface XlsxNamedRange {
  /** Optional comment shown in Excel's Name Manager. */
  comment?: string;
  hidden?: boolean;
  name: string;
  refersTo: string;
  /** Zero-based workbook sheet index for a sheet-scoped name. */
  sheetIndex?: number;
}

/**
 * Excel-style paste options for `pasteSpecialFromClipboard` /
 * `pasteSpecialStructuredData`.
 *
 * `all` keeps formulas, values and styles; `values` writes the displayed value
 * only; `formats` writes only cell styles. `transpose` swaps rows and columns.
 */
export interface XlsxPasteSpecialOptions {
  mode?: "all" | "values" | "formats";
  transpose?: boolean;
}

/** Inline-edit request the grid watches; `token` retriggers the same cell. */
export interface XlsxCellEditRequest {
  cell: XlsxCellAddress;
  token: number;
}

/**
 * Authoring actions for non-cell workbook objects.
 *
 * Implemented by `XlsxViewerController`. Every action mutates the workbook
 * model, re-renders the viewer, records an undo entry and is written into
 * exported XLSX bytes.
 */
export interface XlsxViewerAuthoring {
  // ── Structure ──────────────────────────────────────────────────────
  /**
   * Inserts blank rows at `startRow`, pushing existing rows down.
   *
   * Cell values, formulas, styles, merges, row sizes, hyperlinks, comments,
   * conditional formats, data validations, tables, defined names, frozen panes
   * and drawing anchors are shifted together so the workbook stays consistent.
   *
   * @returns `true` when the workbook changed.
   */
  insertRows: (startRow: number, count?: number) => boolean;
  /** Deletes rows starting at `startRow`, pulling the rows below up. */
  deleteRows: (startRow: number, count?: number) => boolean;
  /** Inserts blank columns at `startColumn`, pushing existing columns right. */
  insertColumns: (startColumn: number, count?: number) => boolean;
  /** Deletes columns starting at `startColumn`, pulling the columns right left. */
  deleteColumns: (startColumn: number, count?: number) => boolean;
  // ── Selection ──────────────────────────────────────────────────────
  /** Selects a whole row (every displayed column). */
  selectRow: (row: number) => void;
  /** Selects a whole column (every displayed row). */
  selectColumn: (column: number) => void;
  /** Selects the whole used range. */
  selectAll: () => void;
  // ── Clipboard ──────────────────────────────────────────────────────
  /**
   * Cuts the selection: writes the payload and marks the source as a pending
   * move (`cutState`); the source cells are cleared by the next paste.
   *
   * The move is one undo entry: the paste snapshots the workbook before it
   * writes, then clears the source, so a single undo restores both ranges.
   *
   * @returns `true` when a cut is pending.
   */
  cutSelection: () => boolean;
  /**
   * Cancels a pending cut, dropping the marquee and the internal payload
   * without touching the workbook (Excel's Escape).
   *
   * @returns `true` when there was a pending cut to cancel.
   */
  cancelCut: () => boolean;
  /**
   * Pastes the last copied or cut payload with Excel-style paste options.
   *
   * `values` writes values only, `formats` writes styles only and `transpose`
   * swaps rows and columns. Falls back to the system clipboard when nothing was
   * copied inside the viewer.
   *
   * @returns `true` when the workbook changed.
   */
  pasteSpecialFromClipboard: (options?: XlsxPasteSpecialOptions) => Promise<boolean>;
  /** Synchronous paste-special entry point for hosts and tests. */
  pasteSpecialStructuredData: (payload: string, options?: XlsxPasteSpecialOptions) => boolean;
  // ── Number formats ─────────────────────────────────────────────────
  /** Applies one number format to the selection as a single undo step. */
  setSelectionNumberFormat: (format: XlsxCellNumberFormatInput) => boolean;
  // ── Defined names ──────────────────────────────────────────────────
  /** Defined names of the workbook, including sheet-scoped names. */
  getNamedRanges: () => XlsxNamedRange[];
  /** Adds or replaces a defined name pointing at an A1 reference. */
  setNamedRange: (name: string, refersTo: string) => boolean;
  /** Removes a defined name. */
  deleteNamedRange: (name: string) => boolean;
  // ── Charts ─────────────────────────────────────────────────────────
  /**
   * Inserts a bar, line or pie chart built from a cell range.
   *
   * @returns The inserted chart, or `null` when the range is empty, the sheet
   * part is unavailable, or editing is locked.
   */
  insertChart: (input?: XlsxChartInsertInput | null) => XlsxChart | null;
  /**
   * Switches an existing chart to another chart type.
   *
   * The chart part is rewritten in place: classic families rebuild the plot
   * area while keeping series names, colours, data labels and axis formatting,
   * modern families rebuild the part as a chartEx chart (with the style parts
   * it needs), and migrations between the two also update the drawing
   * relationship, the `mc:AlternateContent` wrapper and the content type.
   * The mutation is recorded in undo history and is written into exported XLSX
   * bytes.
   *
   * @param id Chart id, as returned by `getSheetCharts()` or `insertChart()`.
   * @param type Chart type option id, for example `"Pie3D"` or `"Treemap"`.
   * @returns `true` when the chart part was rewritten.
   */
  changeChartType: (id: string, type: string) => boolean;
  // ── Freeze panes ───────────────────────────────────────────────────
  /** Frozen rows/columns of the active sheet. */
  readonly freezePanes: XlsxFreezePanes | null;
  /**
   * Freezes rows/columns above and left of `freeze`, or clears the frozen pane
   * when both counts are zero or `null` is passed.
   *
   * @returns `true` when the workbook changed.
   */
  setFreezePanes: (freeze: XlsxFreezePanes | null) => boolean;
  /**
   * Freezes panes at the active cell, or unfreezes when that cell is already
   * the freeze origin.
   *
   * @returns `true` when the workbook changed.
   */
  toggleFreezePanesAtActiveCell: () => boolean;
  // ── Comments ───────────────────────────────────────────────────────
  /** Comment on a cell, or `null`. Defaults to the active cell. */
  getCellComment: (cell?: XlsxCellAddress | null) => XlsxCellComment | null;
  /**
   * Creates, edits or deletes the note on a cell. Pass `null` to delete.
   *
   * @returns `true` when the workbook changed.
   */
  setCellComment: (cell: XlsxCellAddress, input: XlsxCellCommentInput | null) => boolean;
  // ── Hyperlinks ─────────────────────────────────────────────────────
  /** Hyperlink on a cell, or `null`. Defaults to the active cell. */
  getCellHyperlink: (cell?: XlsxCellAddress | null) => XlsxCellHyperlink | null;
  /**
   * Creates, edits or removes the hyperlink on a cell. Pass `null` to remove.
   *
   * @returns `true` when the workbook changed.
   */
  setCellHyperlink: (cell: XlsxCellAddress, input: XlsxCellHyperlinkInput | null) => boolean;
  // ── Conditional formatting ─────────────────────────────────────────
  /** Conditional-format rules of the active sheet, in priority order. */
  getConditionalFormatRules: () => XlsxConditionalFormatRule[];
  /**
   * Adds a conditional-format rule to the active sheet.
   *
   * @returns `true` when the rule was written.
   */
  addConditionalFormatRule: (draft: XlsxConditionalFormatDraft) => boolean;
  /** Replaces the rule at `index`; `draft.ranges` default to the rule's ranges. */
  updateConditionalFormatRule: (index: number, draft: XlsxConditionalFormatDraft) => boolean;
  /** Deletes the rule at `index`. */
  deleteConditionalFormatRule: (index: number) => boolean;
  // ── Tables ─────────────────────────────────────────────────────────
  /**
   * Creates a ListObject table from a range (defaults to the selection).
   *
   * @returns The created table, or `null` when the range or name is invalid.
   */
  createTable: (input?: XlsxTableCreateInput | null) => XlsxTable | null;
  /** Removes a table definition from the active sheet. */
  removeTable: (name: string) => boolean;
  /** Table whose range contains `cell`. Defaults to the active cell. */
  getTableAtCell: (cell?: XlsxCellAddress | null) => XlsxTable | null;
  // ── Data validation ────────────────────────────────────────────────
  /**
   * Applies a data-validation rule to the selection, replacing rules on the
   * same cells. Pass `null` to remove validations from the selection.
   */
  setSelectionDataValidation: (input: XlsxDataValidationInput | null) => boolean;
}

/**
 * Options for {@link XlsxViewerController.replaceInSheet}.
 *
 * The defaults mirror Excel's Find and Replace: case-insensitive literal
 * matching over the whole used range of one sheet, values and formulas.
 */
export interface XlsxReplaceInSheetOptions {
  /** Match case exactly. Defaults to `false`. */
  matchCase?: boolean;
  /** Only replace a cell whose entire text equals the query. Defaults to `false`. */
  wholeCell?: boolean;
  /**
   * Replace inside formulas as well. Defaults to `true` so replace covers the
   * same cells the search session reports. A formula cell is edited through its
   * formula text, never through the cached value it calculates.
   */
  formula?: boolean;
  /** Visible sheet index (a `sheets` index). Defaults to the active sheet. */
  sheetIndex?: number;
  /** Limit the scan to this range. Defaults to the sheet's used range. */
  range?: XlsxCellRange;
  /**
   * Replace only this occurrence instead of every match in each cell.
   *
   * Spreadsheet search reports one match per cell plus the offsets of the
   * matched text, so the Replace button passes the active match here while
   * Replace All omits it.
   */
  occurrence?: {
    cell: XlsxCellAddress;
    start: number;
    end: number;
  };
}

/** What {@link XlsxViewerController.replaceInSheet} changed. */
export interface XlsxReplaceResult {
  /** Visible sheet index the command ran against. */
  sheetIndex: number;
  /** Workbook sheet index of the same sheet, as used by the engine. */
  workbookSheetIndex: number;
  /** Number of cells written. */
  replacedCells: number;
  /** Number of matched occurrences replaced (a cell can hold several). */
  replacedOccurrences: number;
  /** True when the workbook changed and one undo entry was recorded. */
  changed: boolean;
}

export interface XlsxViewerController extends XlsxViewerAuthoring {
  activeCell: XlsxCellAddress | null;
  activeCellAddress: string | null;
  activeSheet: XlsxSheetData | null;
  activeSheetIndex: number;
  activeTab: XlsxWorkbookTab | null;
  activeTabIndex: number;
  canDownload: boolean;
  canExport: boolean;
  canLoadDeferred: boolean;
  /**
   * Why the most recent recalculation fell back to cached formula values, or
   * `null` when formulas are current. Recalculation never traps the workbook
   * silently: hosts can render `message` and call `recalculate()` to retry.
   */
  calculationStatus: {
    reason: "unresolved-sheet-refs" | "calculate-trapped" | "recovery-snapshot-unavailable";
    message: string;
  } | null;
  canRedo: boolean;
  canUndo: boolean;
  canZoomIn: boolean;
  canZoomOut: boolean;
  /**
   * Abandons a deferred load and returns the viewer to its empty state.
   *
   * Ignored when no load is deferred; the pending source bytes are released
   * without parsing.
   */
  cancelDeferredLoad: () => void;
  clearSelectedCells: () => void;
  clearSelection: () => void;
  /**
   * Commits user-entered text to a cell in the active worksheet.
   *
   * Text that starts with `=` is written as a formula through
   * `setCellFormula`; any other text is written as a value through
   * `setCellValue`, so editing a formula cell never replaces the formula with
   * the value it displays. Text that already matches the cell's formula or
   * coerced value is dropped without touching the workbook or undo history.
   *
   * Inline cell editing and the formula bar both commit through this method.
   *
   * @param cell Zero-based row and column address in the active worksheet.
   * @param text Text entered by the user, for example `"=SUM(A1:A3)"` or `"42"`.
   * @returns `true` when the cell was written.
   */
  commitCellInput: (cell: XlsxCellAddress, text: string) => boolean;
  continueDeferredLoad: () => void;
  copySelectionToClipboard: () => Promise<boolean>;
  /**
   * The pending cut (source range and sheet), or `null` when nothing is cut.
   *
   * Hosts draw the moving marquee from it and hide it once a paste lands or
   * `cancelCut()` is called.
   */
  cutState: XlsxClipboardCutState | null;
  /** Pending inline-edit request; the grid starts editing when it changes. */
  editRequest: XlsxCellEditRequest | null;
  /**
   * Requests inline editing of a cell (defaults to the active cell).
   *
   * @returns `true` when a request was issued.
   */
  editCell: (cell?: XlsxCellAddress | null) => boolean;
  dataValidationFeedback: XlsxDataValidationFeedback | null;
  dismissDataValidationFeedback: () => void;
  defaultZoomScale: number;
  deferredLoadFileSize: number | null;
  defineNamedRange: (name: string, range?: XlsxCellRange | null) => void;
  displayFileName: string;
  download: () => void;
  charts: XlsxChart[];
  chartsheets: XlsxChartsheet[];
  /**
   * Ranges currently painted by the viewer's highlight layer.
   *
   * The list merges ranges sourced from the selected chart element with the
   * ranges supplied through {@link highlightRanges}. It is reactive and
   * read-only; mutate it through those methods.
   */
  highlightedRanges: XlsxRangeHighlight[];
  /**
   * Highlights one or more cell ranges in the built-in overlay layer.
   *
   * Ranges are normalized (start = top-left). The call replaces the previous
   * host-supplied ranges, defaults to the active worksheet and to the amber
   * custom colour, and is read-only: no workbook mutation, undo entry or
   * exported byte changes. Chart source highlights are derived from the chart
   * selection and are not replaced. Highlights follow scrolling, zoom and
   * freeze panes.
   *
   * @param ranges Zero-based A1-less ranges on one worksheet.
   * @param options Colour, worksheet index, role and optional identity.
   */
  highlightRanges: (ranges: readonly XlsxCellRange[], options?: XlsxHighlightRangesOptions) => void;
  /** Removes the ranges supplied through {@link highlightRanges}. */
  clearHighlightedRanges: () => void;
  exportCsv: () => void;
  exportXlsx: () => void;
  findCells: (query: string) => XlsxCellAddress[];
  error: Error | null;
  file?: ArrayBuffer;
  fillSelection: (targetRange: XlsxCellRange) => void;
  clearSelectedChart: () => void;
  clearSelectedChartElement: () => void;
  clearSelectedImage: () => void;
  getChartById: (id: string) => XlsxChart | null;
  getChartSeriesFormula: (chartId: string, seriesIndex: number) => string;
  getChartsheetById: (id: string) => XlsxChartsheet | null;
  getCellDataValidation: (cell?: XlsxCellAddress | null) => XlsxDataValidation | null;
  getDataValidationListOptions: (validation: XlsxDataValidation) => string[];
  formControls: XlsxFormControl[];
  getSheetCharts: (sheetIndex?: number) => XlsxChart[];
  getSheetFormControls: (sheetIndex?: number) => XlsxFormControl[];
  getImageById: (id: string) => XlsxImage | null;
  getSheetImages: (sheetIndex?: number) => XlsxImage[];
  getSheetShapes: (sheetIndex?: number) => XlsxShape[];
  getClipboardData: () => XlsxClipboardData | null;
  getCellDisplayValue: (cell?: XlsxCellAddress | null) => string;
  /**
   * Rich-text runs of a cell, or `undefined` when the cell has none.
   *
   * The runs concatenate to `getCellDisplayValue(cell)`. They are dropped once
   * the cell value no longer matches the parsed runs (an edit or recalculation
   * replaced the rich string), so a host can render them without checking.
   */
  getCellRuns: (cell?: XlsxCellAddress | null) => XlsxCellTextRun[] | undefined;
  getCellFormula: (cell?: XlsxCellAddress | null) => string;
  getCellSnapshotAsync?: (workbookSheetIndex: number, row: number, col: number, signal?: AbortSignal) => Promise<{
    displayValue: string;
    formula: string;
    /** Resolved number format of the cell, so the viewer can localize it. */
    numberFormat?: XlsxCellNumberFormatInput | null;
    /** Rich-text runs, so Worker-backed viewers keep per-run formatting. */
    runs?: XlsxCellTextRun[];
  }>;
  /** Flat print layout of the active sheet, or `null` when none is loaded. */
  getPrintLayout: () => XlsxPrintLayout | null;
  /** Pages of the active sheet's current print layout. */
  printPages: XlsxPrintPage[];
  /** Current print layout of the active sheet. */
  printLayout: XlsxPrintLayout | null;
  /** Whether the grid shows Excel-style dashed page-break preview lines. */
  isPageBreakPreview: boolean;
  /** Enables or disables page-break preview on the active sheet. */
  setPageBreakPreview: (enabled: boolean) => void;
  /** Whether the print-preview surface is open. */
  isPrintPreviewOpen: boolean;
  /**
   * Opens the print preview for the active sheet.
   *
   * @returns `false` when no worksheet is loaded.
   */
  openPrintPreview: () => boolean;
  /** Closes the print preview without printing. */
  closePrintPreview: () => void;
  /**
   * Opens the print preview and asks the host window to print it.
   *
   * @returns `false` when no worksheet is loaded or the environment has no
   *   `window.print`.
   */
  print: () => boolean;
  isLoadDeferred: boolean;
  isLoading: boolean;
  isActiveSheetProtected: boolean;
  isEditingLocked: boolean;
  isChartsLoading: boolean;
  isWorkerBacked?: boolean;
  /**
   * Warning raised by the most recent export, for example when an `.xlsm` or
   * `.xltm` source is written as a macro-free `.xlsx`. Cleared when a new
   * source starts loading.
   */
  lastExportWarning: string | null;
  /** Locale used for number and date formats that do not name one. */
  locale: string;
  maxConcurrentImageDecodes?: number;
  sourceState: XlsxSourceState;
  sourceError: XlsxLoadError | null;
  reportImageDecodeError?: (imageId: string, message?: string) => void;
  images: XlsxImage[];
  moveChartBy: (id: string, deltaX: number, deltaY: number) => void;
  shapes: XlsxShape[];
  /**
   * Merges the current selection without touching its alignment.
   *
   * Hosts that want Excel's default action use `mergeAndCenterSelection`;
   * this entry stays available for hosts that merge without centring.
   */
  mergeSelection: () => void;
  /**
   * Merges every row of the selection across its columns (Excel's Merge
   * across), leaving each row's alignment untouched.
   */
  mergeAcrossSelection: () => void;
  /**
   * Merges the current selection and centres its text both ways.
   *
   * A single-cell selection is only centred, because Excel never writes a
   * one-cell merge region. The merge and the alignment land in one undo step.
   */
  mergeAndCenterSelection: () => void;
  /**
   * Draws a border preset on the current selection.
   *
   * `all` borders every cell of the selection, `outside` only its perimeter and
   * `left`/`right`/`top`/`bottom` that edge of every cell. The line style and
   * colour come from the command, so per-edge colour is available without
   * dropping to `setRangeStyle`. The mutation is recorded in undo history and is
   * included in exported XLSX bytes.
   *
   * @returns `false` when there is no selection, the sheet is locked, or the
   *   command changed nothing.
   */
  setSelectionBorders: (input?: XlsxBorderPresetInput) => boolean;
  maxZoomScale: number;
  minZoomScale: number;
  moveImageBy: (id: string, deltaX: number, deltaY: number) => void;
  removeActiveSheet: () => void;
  readOnly: boolean;
  recalculate: () => void;
  /**
   * Re-resolves chart categories and values from the current workbook model.
   *
   * Main-thread mode refreshes from the live workbook. Worker-backed mode asks
   * the Worker to refresh from its own edited workbook snapshot, so charts do
   * not stay on the source file's cached chart XML values after an in-worker
   * edit.
   *
   * @returns `false` when no workbook or Worker is available.
   */
  refreshChartData: () => Promise<boolean>;
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
  /**
   * Replaces `query` with `replacement` in one worksheet.
   *
   * Covers both cell values and formulas (see
   * {@link XlsxReplaceInSheetOptions}), optionally limited to a range or to one
   * previously reported search occurrence. Everything the command writes is
   * recorded as a *single* undo entry, and a sheet other than the active one is
   * activated first so the history entry and the visible grid agree.
   *
   * Read-only viewers, protected sheets and worker-backed (read-only) loads
   * return an empty result without touching the workbook.
   *
   * @param query Literal text to find.
   * @param replacement Text written in place of every match.
   * @param options Case/whole-cell/formula switches, sheet and range scoping.
   */
  replaceInSheet: (
    query: string,
    replacement: string,
    options?: XlsxReplaceInSheetOptions
  ) => XlsxReplaceResult;
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
  selectedShape: XlsxShape | null;
  selectedShapeId: string | null;
  getShapeById: (id: string) => XlsxShape | null;
  selectShape: (id: string | null) => void;
  clearSelectedShape: () => void;
  setShapeRect: (id: string, rect: XlsxImageRect) => void;
  moveShapeBy: (id: string, deltaX: number, deltaY: number) => void;
  resizeShapeBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void;
  deleteShape: (id: string) => boolean;
  deleteSelectedDrawing: () => boolean;
  deleteImage: (id: string) => boolean;
  deleteChart: (id: string) => boolean;
  copyDrawing: () => XlsxClipboardData | null;
  cutDrawing: () => XlsxClipboardData | null;
  pasteDrawing: () => boolean;
  duplicateDrawing: () => boolean;
  addFormControl: (input: XlsxFormControlInput) => XlsxFormControl | null;
  updateFormControl: (id: string, patch: Partial<XlsxFormControlInput>) => boolean;
  removeFormControl: (id: string) => boolean;
  getFormControlById: (id: string) => XlsxFormControl | null;
  getFormControlItems: (sheetIndex?: number) => XlsxFormControl[];
  setSelectedFormula: (formula: string) => boolean;
  setSelectedCellFormula: (formula: string) => void;
  /**
   * Switches the locale used by number and date formats that name none of
   * their own, and repaints the workbook.
   *
   * @param locale BCP-47 tag such as `"de-DE"`; omitted or invalid values fall
   * back to `"en-US"`.
   */
  setDocumentLocale: (locale?: string) => void;
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
  /** AutoFilter ranges of the active sheet, in sheet order. */
  autoFilterRanges: XlsxAutoFilterRange[];
  /** Active view-level filters of the active sheet. */
  autoFilterStates: XlsxAutoFilterState[];
  /** Absolute rows hidden by the active sheet's AutoFilter menus. */
  hiddenAutoFilterRows: number[];
  /**
   * Distinct display values of one AutoFilter column, blanks included.
   *
   * The list is capped (the menu reports the remainder) and never mutates the
   * workbook.
   */
  getAutoFilterColumnValues: (range: XlsxAutoFilterRange, columnIndex: number) => XlsxAutoFilterValueOption[];
  /**
   * Applies a value filter to one AutoFilter column.
   *
   * Passing `null` or an empty list removes that column's filter. Rows that do
   * not match every filtered column are hidden; the criterion is view state, so
   * no cell, merge or frozen pane is modified. Returns `false` when the range
   * or column does not exist.
   */
  applyAutoFilterValues: (
    reference: string,
    columnIndex: number,
    values: string[] | null,
    options?: { includeBlank?: boolean }
  ) => boolean;
  /**
   * Removes the filters of one range, or every AutoFilter filter of the sheet
   * when `reference` is omitted, restoring the hidden rows.
   */
  clearAutoFilter: (reference?: string) => boolean;
  /**
   * Sorts the rows of one AutoFilter range by a column.
   *
   * The sort is a normal undoable range edit: one entry restores the previous
   * row order. It is refused in read-only/protected sheets and when a merged
   * region intersects the data rows, because moving those rows would corrupt
   * the merge.
   */
  sortAutoFilter: (reference: string, columnIndex: number, direction: XlsxTableSortDirection) => boolean;
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
  /**
   * Moves a sheet tab to another slot in the tab strip.
   *
   * `fromIndex` and `toIndex` are tab indices — positions in `tabs`, which
   * interleave worksheets and chartsheets — not `sheets` or `chartsheets`
   * indices. The order after the move is `tabs.splice(toIndex, 0,
   * ...tabs.splice(fromIndex, 1))`: the moved tab becomes active and stays
   * selected, and dropping a tab on its own slot is a no-op.
   *
   * `@dukelib/sheets-wasm` has no reorder setter, so the viewer rewrites the
   * `<sheets>` order in its saved workbook and parses the result back into the
   * workbook model. Sheet-scoped names and workbook-view tabs keep pointing at
   * the same sheet, every per-sheet asset is re-read from the reordered
   * package (images, charts, shapes, form controls and sheet state follow their
   * sheet), the move is one undo entry, and exported XLSX bytes keep the new
   * order.
   *
   * @param fromIndex Zero-based tab index being moved.
   * @param toIndex Zero-based destination tab index in the reordered strip.
   * @returns `true` when the order changed; `false` for an out-of-range index,
   *   an unchanged order, a read-only viewer, or a workbook that cannot be
   *   rewritten. Sheet protection does not block a move, matching Excel.
   */
  moveSheet: (fromIndex: number, toIndex: number) => boolean;
  /**
   * Renames a worksheet, the way Excel's tab rename does.
   *
   * The name is validated first (trimmed, non-empty, at most 31 characters, no
   * `[ ] : * ? / \\`, no leading or trailing apostrophe, unique across
   * worksheets and chartsheets ignoring case) and the workbook is left
   * untouched when a rule fails. A successful rename keeps worksheet order, so
   * `activeSheetIndex` and `activeTabIndex` do not move; it is one undo entry
   * and is written into exported XLSX bytes.
   *
   * `@dukelib/sheets-wasm` has no rename setter, so the viewer rewrites the
   * worksheet name and every reference to it in its saved workbook and parses
   * the result back into the workbook model.
   *
   * @param index Zero-based visible sheet index, as used by `sheets`,
   *   `activeSheetIndex` and `setActiveSheetIndex`.
   * @param name New worksheet name. Surrounding whitespace is trimmed.
   * @returns `true` when the sheet was renamed; `false` when the name is
   *   rejected, the index does not name a worksheet, or the viewer is
   *   read-only. Sheet protection does not block a rename, matching Excel.
   */
  renameSheet: (index: number, name: string) => boolean;
  /** Renames the active worksheet; see `renameSheet`. */
  renameActiveSheet: (name: string) => boolean;
  /**
   * Reports why Excel would reject a worksheet name before it is committed.
   *
   * Inline rename inputs use this for live validation. The sheet at `index` is
   * left out of the uniqueness check so a case-only rename stays possible.
   *
   * @param name Candidate name, as typed.
   * @param index Visible sheet index being renamed, when known.
   * @returns The first broken rule, or `null` when the name is acceptable.
   */
  getSheetNameError: (name: string, index?: number) => XlsxSheetNameIssue | null;
  /**
   * Hides, unhides or very-hides a worksheet at runtime.
   *
   * The sheet is addressed by workbook sheet index — the value every
   * `XlsxSheetData` and worksheet tab exposes as `workbookSheetIndex` — not by
   * tab or visible-sheet index. A hidden sheet is absent from `tabs`, so it
   * has no tab index, while `workbookSheetIndex` identifies both listed and
   * hidden worksheets.
   *
   * Hiding the active worksheet selects the nearest remaining visible tab and
   * clears the sheet-scoped selection and overlay state; hiding the last
   * visible worksheet is refused, mirroring `removeActiveSheet`. Unhiding a
   * worksheet makes it active. `veryHidden` worksheets leave `tabs` and the
   * `hiddenSheets` unhide list even when hidden tabs are shown, like Excel.
   *
   * `@dukelib/sheets-wasm` exposes visibility as read-only, so the viewer
   * rewrites the `state` attribute of the worksheet in its saved workbook,
   * reparses the package and commits the result as one undo entry. Visibility
   * is therefore preserved by `exportXlsx` and by a later reload, and images,
   * charts, shapes, form controls and sheet state are re-read from the rebuilt
   * package so their `workbookSheetIndex` keys stay aligned.
   *
   * @param workbookSheetIndex Zero-based worksheet index in the workbook.
   * @param visibility Excel visibility to write.
   * @returns `true` when visibility changed; `false` for a no-op, an invalid
   *   index, a read-only viewer, an attempt to hide the last visible sheet, or
   *   a package that cannot be rewritten.
   */
  setSheetVisibility: (
    workbookSheetIndex: number,
    visibility: XlsxSheetVisibility
  ) => boolean;
  /**
   * Current visibility of a worksheet, including `veryHidden`.
   *
   * Worker-backed read-only viewers only know visibility for sheets present in
   * their snapshot; `null` means the index is not a worksheet the controller
   * can see.
   */
  getSheetVisibility: (workbookSheetIndex: number) => XlsxSheetVisibility | null;
  /**
   * Worksheets the tab context menu offers to unhide, in workbook order.
   *
   * `veryHidden` worksheets are omitted, matching Excel's Unhide dialog. The
   * list comes from the live workbook, so it is populated even when
   * `showHiddenSheets` is off and hidden tabs are not rendered.
   */
  hiddenSheets: XlsxHiddenSheet[];
  /** Whether hidden worksheet tabs are included in `tabs` and the tab strip. */
  showHiddenSheets: boolean;
  /**
   * Includes or removes hidden worksheet tabs at runtime.
   *
   * This is a view option, not a workbook edit: it rebuilds `sheets`, `tabs`
   * and the visible-index map from the current workbook and records no undo
   * entry. `veryHidden` worksheets are never listed. Turning it off while a
   * hidden sheet is active switches to the nearest visible worksheet.
   *
   * @returns `true` when the mode changed (or already matched); `false` when
   *   the viewer has no main-thread workbook to rebuild from, such as a
   *   Worker-backed read-only viewer.
   */
  setShowHiddenSheets: (includeHidden: boolean) => boolean;
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
  /** Moves a sheet tab; see `XlsxViewerController.moveSheet`. */
  moveSheet: (fromIndex: number, toIndex: number) => boolean;
  /** Renames a worksheet; see `XlsxViewerController.renameSheet`. */
  renameSheet: (index: number, name: string) => boolean;
  /** Renames the active worksheet; see `XlsxViewerController.renameSheet`. */
  renameActiveSheet: (name: string) => boolean;
  /** Explains why a worksheet name is rejected; `null` when it is acceptable. */
  getSheetNameError: (name: string, index?: number) => XlsxSheetNameIssue | null;
  /** Hides, unhides or very-hides a worksheet; see `XlsxViewerController.setSheetVisibility`. */
  setSheetVisibility: (
    workbookSheetIndex: number,
    visibility: XlsxSheetVisibility
  ) => boolean;
  /** Current worksheet visibility; see `XlsxViewerController.getSheetVisibility`. */
  getSheetVisibility: (workbookSheetIndex: number) => XlsxSheetVisibility | null;
  /** Worksheets offered by the Unhide menu; see `XlsxViewerController.hiddenSheets`. */
  hiddenSheets: XlsxHiddenSheet[];
  /** Includes hidden worksheet tabs; see `XlsxViewerController.showHiddenSheets`. */
  showHiddenSheets: boolean;
  /** Includes or removes hidden tabs at runtime; see `XlsxViewerController.setShowHiddenSheets`. */
  setShowHiddenSheets: (includeHidden: boolean) => boolean;
  canRedo: boolean;
  canUndo: boolean;
  clearSelectedCells: () => void;
  /**
   * Commits user-entered text to a cell in the active worksheet.
   *
   * Text that starts with `=` is written as a formula, anything else as a
   * value, so an edited formula cell keeps its formula. Text that already
   * matches the cell's formula or coerced value is dropped without touching
   * the workbook or undo history.
   *
   * @param cell Zero-based row and column address in the active worksheet.
   * @param text Text entered by the user, for example `"=SUM(A1:A3)"` or `"42"`.
   * @returns `true` when the cell was written.
   */
  commitCellInput: (cell: XlsxCellAddress, text: string) => boolean;
  copySelectionToClipboard: () => Promise<boolean>;
  defineNamedRange: (name: string, range?: XlsxCellRange | null) => void;
  fillSelection: (targetRange: XlsxCellRange) => void;
  findCells: (query: string) => XlsxCellAddress[];
  insertRows: (startRow: number, count?: number) => boolean;
  deleteRows: (startRow: number, count?: number) => boolean;
  insertColumns: (startColumn: number, count?: number) => boolean;
  deleteColumns: (startColumn: number, count?: number) => boolean;
  selectRow: (row: number) => void;
  selectColumn: (column: number) => void;
  selectAll: () => void;
  cutSelection: () => boolean;
  cancelCut: () => boolean;
  /**
   * Captures the selected chart, image or shape as clipboard data.
   *
   * Returns `null` when no drawing is selected. The caller can write the
   * returned flavours to the system clipboard; the viewer also keeps the
   * payload so `pasteDrawing()` works where the clipboard is unreadable.
   */
  copyDrawing: () => XlsxClipboardData | null;
  /** Captures the selected drawing for a cut-paste move. */
  cutDrawing: () => XlsxClipboardData | null;
  /** Pastes the drawing clipboard onto the active sheet at the active cell. */
  pasteDrawing: () => boolean;
  /** Copies the selected drawing in place with a small offset. */
  duplicateDrawing: () => boolean;
  /**
   * The pending cut (source range and sheet), or `null` when nothing is cut.
   *
   * Hosts draw the moving marquee from it and hide it once a paste lands or
   * `cancelCut()` is called.
   */
  cutState: XlsxClipboardCutState | null;
  pasteSpecialFromClipboard: (options?: XlsxPasteSpecialOptions) => Promise<boolean>;
  pasteSpecialStructuredData: (payload: string, options?: XlsxPasteSpecialOptions) => boolean;
  setSelectionNumberFormat: (format: XlsxCellNumberFormatInput) => boolean;
  getNamedRanges: () => XlsxNamedRange[];
  setNamedRange: (name: string, refersTo: string) => boolean;
  deleteNamedRange: (name: string) => boolean;
  getClipboardData: () => XlsxClipboardData | null;
  getCellDisplayValue: (cell?: XlsxCellAddress | null) => string;
  getCellFormula: (cell?: XlsxCellAddress | null) => string;
  /** Merges the current selection without touching its alignment. */
  mergeSelection: () => void;
  /** Merges every row of the selection across its columns. */
  mergeAcrossSelection: () => void;
  /** Merges the current selection and centres its text both ways. */
  mergeAndCenterSelection: () => void;
  /**
   * Draws a border preset on the current selection.
   *
   * @returns `false` when there is no selection, the sheet is locked, or the
   *   command changed nothing.
   */
  setSelectionBorders: (input?: XlsxBorderPresetInput) => boolean;
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
  autoFilterRanges: XlsxAutoFilterRange[];
  autoFilterStates: XlsxAutoFilterState[];
  applyAutoFilterValues: (
    reference: string,
    columnIndex: number,
    values: string[] | null,
    options?: { includeBlank?: boolean }
  ) => boolean;
  clearAutoFilter: (reference?: string) => boolean;
  getAutoFilterColumnValues: (range: XlsxAutoFilterRange, columnIndex: number) => XlsxAutoFilterValueOption[];
  hiddenAutoFilterRows: number[];
  sortAutoFilter: (reference: string, columnIndex: number, direction: XlsxTableSortDirection) => boolean;
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
  selectedShape: XlsxShape | null;
  selectedShapeId: string | null;
  getShapeById: (id: string) => XlsxShape | null;
  selectShape: (id: string | null) => void;
  clearSelectedShape: () => void;
  setShapeRect: (id: string, rect: XlsxImageRect) => void;
  moveShapeBy: (id: string, deltaX: number, deltaY: number) => void;
  resizeShapeBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void;
  deleteShape: (id: string) => boolean;
  deleteImage: (id: string) => boolean;
  deleteChart: (id: string) => boolean;
  deleteSelectedDrawing: () => boolean;
  copyDrawing: () => XlsxClipboardData | null;
  cutDrawing: () => XlsxClipboardData | null;
  pasteDrawing: () => boolean;
  duplicateDrawing: () => boolean;
  addFormControl: (input: XlsxFormControlInput) => XlsxFormControl | null;
  updateFormControl: (id: string, patch: Partial<XlsxFormControlInput>) => boolean;
  removeFormControl: (id: string) => boolean;
  getFormControlById: (id: string) => XlsxFormControl | null;
  getFormControlItems: (sheetIndex?: number) => XlsxFormControl[];
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
  changeChartType: (id: string, type: string) => boolean;
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
  /**
   * Warms the preview data (geometry and cell values) for the given sheets so a
   * later `paintThumbnail` only draws. Callers schedule this on idle, hovered
   * tabs first; the composable dedupes per sheet revision, caps how many sheets
   * one call warms and ignores unknown indices.
   */
  prefetchThumbnails: (sheetIndices: readonly number[]) => void;
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
