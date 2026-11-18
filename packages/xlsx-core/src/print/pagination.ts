import type {
  XlsxAxisRange,
  XlsxCellRange,
  XlsxPageBreak,
  XlsxPageMargins,
  XlsxPageSetup,
  XlsxPaperSize,
  XlsxPrintLayout,
  XlsxPrintLayoutOptions,
  XlsxPrintPage,
  XlsxPrintPageGrid,
  XlsxPrintPageGridAxis,
  XlsxPrintPageGridCell,
  XlsxPrintOptions,
  XlsxPrintTitles,
} from "../types";

/**
 * Structural sheet subset used by pagination. Keeping it local avoids pulling
 * the WASM Workbook type into the pure `@arcships/xlsx-core/core` entry.
 */
export interface XlsxPrintSheetData {
  name: string;
  workbookSheetIndex: number;
  colBreaks: XlsxPageBreak[];
  colWidthOverridesPx: Record<number, number>;
  colWidths?: number[];
  visibleCols?: number[];
  defaultColWidthPx: number;
  defaultRowHeightPx: number;
  hiddenCols?: number[];
  hiddenRows?: number[];
  maxUsedCol: number;
  maxUsedRow: number;
  mergedRegions?: XlsxCellRange[];
  minUsedCol: number;
  minUsedRow: number;
  pageMargins: XlsxPageMargins;
  pageSetup: XlsxPageSetup;
  printAreas: XlsxCellRange[];
  printOptions: XlsxPrintOptions;
  printTitles: XlsxPrintTitles;
  rowBreaks: XlsxPageBreak[];
  rowHeightOverridesPx: Record<number, number>;
  rowHeights?: number[];
  visibleRows?: number[];
}

// ── Constants ────────────────────────────────────────────────────────────

const CSS_PX_PER_INCH = 96;
const MAX_SHEET_ROWS = 1_048_576;
const MAX_SHEET_COLUMNS = 16_384;
const DEFAULT_MAX_PAGES = 5_000;
const MIN_SCALE_PERCENT = 10;
const MAX_SCALE_PERCENT = 400;
const MAX_PRINT_TITLE_AXIS_COUNT = 100;

/**
 * OOXML paper-size identifiers (ECMA-376 `ST_PaperSize`) in inches.
 *
 * Unknown identifiers fall back to Letter so a malformed workbook still gets a
 * printable page instead of an empty preview.
 */
export const PAPER_SIZES: Readonly<Record<number, Omit<XlsxPaperSize, "id">>> = {
  1: { heightIn: 11, name: "Letter", widthIn: 8.5 },
  2: { heightIn: 11, name: "Letter Small", widthIn: 8.5 },
  3: { heightIn: 17, name: "Tabloid", widthIn: 11 },
  4: { heightIn: 11, name: "Ledger", widthIn: 17 },
  5: { heightIn: 14, name: "Legal", widthIn: 8.5 },
  6: { heightIn: 8.5, name: "Statement", widthIn: 5.5 },
  7: { heightIn: 10.5, name: "Executive", widthIn: 7.25 },
  8: { heightIn: 16.5354, name: "A3", widthIn: 11.6929 },
  9: { heightIn: 11.6929, name: "A4", widthIn: 8.2677 },
  10: { heightIn: 11.6929, name: "A4 Small", widthIn: 8.2677 },
  11: { heightIn: 8.2677, name: "A5", widthIn: 5.8268 },
  12: { heightIn: 13.8976, name: "B4", widthIn: 9.8425 },
  13: { heightIn: 9.8425, name: "B5", widthIn: 6.9291 },
  14: { heightIn: 13, name: "Folio", widthIn: 8.5 },
  15: { heightIn: 10.8268, name: "Quarto", widthIn: 8.5 },
  16: { heightIn: 14, name: "10 × 14", widthIn: 10 },
  17: { heightIn: 17, name: "11 × 17", widthIn: 11 },
  18: { heightIn: 11, name: "Note", widthIn: 8.5 },
  19: { heightIn: 8.875, name: "Envelope #9", widthIn: 3.875 },
  20: { heightIn: 9.5, name: "Envelope #10", widthIn: 4.125 },
  21: { heightIn: 10.375, name: "Envelope #11", widthIn: 4.5 },
  22: { heightIn: 11, name: "Envelope #12", widthIn: 4.75 },
  23: { heightIn: 11.5, name: "Envelope #14", widthIn: 5 },
  24: { heightIn: 22, name: "C", widthIn: 17 },
  25: { heightIn: 34, name: "D", widthIn: 22 },
  26: { heightIn: 44, name: "E", widthIn: 34 },
  27: { heightIn: 8.6614, name: "Envelope DL", widthIn: 4.3307 },
  28: { heightIn: 9.0157, name: "Envelope C5", widthIn: 6.378 },
  29: { heightIn: 18.0315, name: "Envelope C3", widthIn: 12.7559 },
  30: { heightIn: 12.7559, name: "Envelope C4", widthIn: 9.0157 },
  31: { heightIn: 6.378, name: "Envelope C6", widthIn: 4.4882 },
  32: { heightIn: 9.0157, name: "Envelope C65", widthIn: 4.4882 },
  33: { heightIn: 13.8976, name: "Envelope B4", widthIn: 9.8425 },
  34: { heightIn: 9.8425, name: "Envelope B5", widthIn: 6.9291 },
  35: { heightIn: 6.9291, name: "Envelope B6", widthIn: 4.9213 },
  36: { heightIn: 9.0551, name: "Envelope Italy", widthIn: 4.3307 },
  37: { heightIn: 7.5, name: "Envelope Monarch", widthIn: 3.875 },
  38: { heightIn: 6.5, name: "Envelope 6¾", widthIn: 3.625 },
  39: { heightIn: 11, name: "US Standard Fanfold", widthIn: 14.875 },
  40: { heightIn: 12, name: "German Standard Fanfold", widthIn: 8.5 },
  41: { heightIn: 13, name: "German Legal Fanfold", widthIn: 8.5 },
  42: { heightIn: 13.8976, name: "ISO B4", widthIn: 9.8425 },
  43: { heightIn: 5.8268, name: "Japanese Postcard", widthIn: 3.937 },
  44: { heightIn: 11, name: "9 × 11", widthIn: 9 },
  45: { heightIn: 11, name: "10 × 11", widthIn: 10 },
  46: { heightIn: 11, name: "15 × 11", widthIn: 15 },
  47: { heightIn: 8.6614, name: "Invite Envelope", widthIn: 8.6614 },
  48: { heightIn: 12, name: "Letter Extra", widthIn: 9.275 },
  49: { heightIn: 15, name: "Legal Extra", widthIn: 9.275 },
  50: { heightIn: 18, name: "Tabloid Extra", widthIn: 11.6929 },
  51: { heightIn: 12.6929, name: "A4 Extra", widthIn: 9.275 },
  52: { heightIn: 8.5, name: "Letter Transverse", widthIn: 11 },
  53: { heightIn: 8.2677, name: "A4 Transverse", widthIn: 11.6929 },
  54: { heightIn: 9.275, name: "Letter Extra Transverse", widthIn: 12 },
  55: { heightIn: 14, name: "Super A", widthIn: 8.937 },
  56: { heightIn: 17, name: "Super B", widthIn: 12 },
  57: { heightIn: 12.6929, name: "Letter Plus", widthIn: 8.5 },
  58: { heightIn: 12.9921, name: "A4 Plus", widthIn: 8.2677 },
  59: { heightIn: 5.8268, name: "A5 Transverse", widthIn: 8.2677 },
  60: { heightIn: 6.9291, name: "B5 JIS Transverse", widthIn: 9.8425 },
  61: { heightIn: 17.5197, name: "A3 Extra", widthIn: 12.6929 },
  62: { heightIn: 8.6614, name: "A5 Extra", widthIn: 6.1417 },
  63: { heightIn: 10.1181, name: "B5 ISO Extra", widthIn: 7.1654 },
  64: { heightIn: 23.3858, name: "A2", widthIn: 16.5354 },
  65: { heightIn: 11.6929, name: "A3 Transverse", widthIn: 16.5354 },
  66: { heightIn: 12.6929, name: "A3 Extra Transverse", widthIn: 17.5197 },
  67: { heightIn: 5.8268, name: "Japanese Double Postcard", widthIn: 7.874 },
  68: { heightIn: 5.8268, name: "A6", widthIn: 4.1339 },
  69: { heightIn: 13.0709, name: "Japanese Envelope Kaku #2", widthIn: 9.4488 },
  70: { heightIn: 10.9055, name: "Japanese Envelope Kaku #3", widthIn: 8.5039 },
  71: { heightIn: 9.252, name: "Japanese Envelope Chou #3", widthIn: 4.7244 },
  72: { heightIn: 8.0709, name: "Japanese Envelope Chou #4", widthIn: 3.5433 },
  73: { heightIn: 8.5, name: "Letter Rotated", widthIn: 11 },
  74: { heightIn: 11.6929, name: "A3 Rotated", widthIn: 16.5354 },
  75: { heightIn: 8.2677, name: "A4 Rotated", widthIn: 11.6929 },
  76: { heightIn: 5.8268, name: "A5 Rotated", widthIn: 8.2677 },
  77: { heightIn: 9.8425, name: "B4 JIS Rotated", widthIn: 13.8976 },
  78: { heightIn: 6.9291, name: "B5 JIS Rotated", widthIn: 9.8425 },
  79: { heightIn: 3.937, name: "Japanese Postcard Rotated", widthIn: 5.8268 },
  80: { heightIn: 5.8268, name: "Japanese Double Postcard Rotated", widthIn: 7.874 },
  81: { heightIn: 4.1339, name: "A6 Rotated", widthIn: 5.8268 },
  82: { heightIn: 9.4488, name: "Japanese Envelope Kaku #2 Rotated", widthIn: 13.0709 },
  83: { heightIn: 8.5039, name: "Japanese Envelope Kaku #3 Rotated", widthIn: 10.9055 },
  84: { heightIn: 4.7244, name: "Japanese Envelope Chou #3 Rotated", widthIn: 9.252 },
  85: { heightIn: 3.5433, name: "Japanese Envelope Chou #4 Rotated", widthIn: 8.0709 },
  86: { heightIn: 7.1654, name: "JIS B6", widthIn: 5.0394 },
  87: { heightIn: 5.0394, name: "JIS B6 Rotated", widthIn: 7.1654 },
  88: { heightIn: 11, name: "12 × 11", widthIn: 12 },
  89: { heightIn: 9.252, name: "Japanese Envelope You #4", widthIn: 4.1339 },
  90: { heightIn: 4.1339, name: "Japanese Envelope You #4 Rotated", widthIn: 9.252 },
  91: { heightIn: 10.2362, name: "PRC 16K", widthIn: 7.2441 },
  92: { heightIn: 7.2441, name: "PRC 32K", widthIn: 5.1181 },
  93: { heightIn: 7.9921, name: "PRC 32K Big", widthIn: 5.5118 },
  94: { heightIn: 6.4961, name: "PRC Envelope #1", widthIn: 4.0157 },
  95: { heightIn: 6.9291, name: "PRC Envelope #2", widthIn: 4.0157 },
  96: { heightIn: 6.9291, name: "PRC Envelope #3", widthIn: 4.9213 },
  97: { heightIn: 8.189, name: "PRC Envelope #4", widthIn: 4.3307 },
  98: { heightIn: 8.6614, name: "PRC Envelope #5", widthIn: 4.3307 },
  99: { heightIn: 9.0551, name: "PRC Envelope #6", widthIn: 4.7244 },
  100: { heightIn: 9.0551, name: "PRC Envelope #7", widthIn: 6.2992 },
  101: { heightIn: 12.2047, name: "PRC Envelope #8", widthIn: 4.7244 },
  102: { heightIn: 12.2047, name: "PRC Envelope #9", widthIn: 6.2992 },
  103: { heightIn: 12.7559, name: "PRC Envelope #10", widthIn: 4.9213 },
  104: { heightIn: 12.4409, name: "PRC Envelope 16K", widthIn: 8.7402 },
  105: { heightIn: 8.7402, name: "PRC Envelope 32K", widthIn: 5.1181 },
  106: { heightIn: 9.8425, name: "PRC Envelope 32K Big", widthIn: 5.1181 },
};

/** Resolves an OOXML paper-size id, falling back to Letter. */
export function resolvePaperSize(paperSize: number): XlsxPaperSize {
  const id = Number.isInteger(paperSize) && paperSize > 0 ? paperSize : 1;
  const entry = PAPER_SIZES[id] ?? PAPER_SIZES[1]!;
  return { id: PAPER_SIZES[id] ? id : 1, ...entry };
}

/** Converts a paper size to CSS pixels, swapping axes for landscape. */
export function resolvePaperPixelSize(
  paperSize: XlsxPaperSize,
  orientation: XlsxPageSetup["orientation"],
): { heightPx: number; widthPx: number } {
  const shortEdge = Math.min(paperSize.widthIn, paperSize.heightIn);
  const longEdge = Math.max(paperSize.widthIn, paperSize.heightIn);
  const widthIn = orientation === "landscape" ? longEdge : shortEdge;
  const heightIn = orientation === "landscape" ? shortEdge : longEdge;
  return {
    heightPx: Math.round(heightIn * CSS_PX_PER_INCH),
    widthPx: Math.round(widthIn * CSS_PX_PER_INCH),
  };
}

/** Margins in CSS pixels for the printable box. */
export function resolvePageMarginsPx(pageMargins: XlsxPageMargins) {
  return {
    bottom: Math.round(Math.max(0, pageMargins.bottom) * CSS_PX_PER_INCH),
    footer: Math.round(Math.max(0, pageMargins.footer) * CSS_PX_PER_INCH),
    header: Math.round(Math.max(0, pageMargins.header) * CSS_PX_PER_INCH),
    left: Math.round(Math.max(0, pageMargins.left) * CSS_PX_PER_INCH),
    right: Math.round(Math.max(0, pageMargins.right) * CSS_PX_PER_INCH),
    top: Math.round(Math.max(0, pageMargins.top) * CSS_PX_PER_INCH),
  };
}

// ── Range helpers ────────────────────────────────────────────────────────

function resolveColumnWidthPx(sheet: XlsxPrintSheetData, col: number) {
  if (col < 0) return 0;
  const fallback = sheet.defaultColWidthPx;
  const { colWidths, visibleCols } = sheet;
  if (visibleCols?.length && colWidths?.length) {
    const displayIndex = visibleCols.indexOf(col);
    if (displayIndex >= 0) return colWidths[displayIndex] ?? fallback;
    return col > visibleCols[visibleCols.length - 1]! && !sheet.hiddenCols?.includes(col) ? fallback : 0;
  }
  return sheet.colWidthOverridesPx?.[col] ?? fallback;
}

function resolveRowHeightPx(sheet: XlsxPrintSheetData, row: number) {
  if (row < 0) return 0;
  const fallback = sheet.defaultRowHeightPx;
  const { rowHeights, visibleRows } = sheet;
  if (visibleRows?.length && rowHeights?.length) {
    const displayIndex = visibleRows.indexOf(row);
    if (displayIndex >= 0) return rowHeights[displayIndex] ?? fallback;
    return row > visibleRows[visibleRows.length - 1]! && !sheet.hiddenRows?.includes(row) ? fallback : 0;
  }
  return sheet.rowHeightOverridesPx?.[row] ?? fallback;
}

function normalizeRange(range: XlsxCellRange): XlsxCellRange {
  return {
    end: {
      col: Math.max(range.start.col, range.end.col),
      row: Math.max(range.start.row, range.end.row),
    },
    start: {
      col: Math.min(range.start.col, range.end.col),
      row: Math.min(range.start.row, range.end.row),
    },
  };
}

function clampRange(range: XlsxCellRange, maxRow: number, maxCol: number): XlsxCellRange {
  const normalized = normalizeRange(range);
  return {
    end: {
      col: Math.min(maxCol, Math.max(0, normalized.end.col)),
      row: Math.min(maxRow, Math.max(0, normalized.end.row)),
    },
    start: {
      col: Math.min(maxCol, Math.max(0, normalized.start.col)),
      row: Math.min(maxRow, Math.max(0, normalized.start.row)),
    },
  };
}

function axisRangeIndices(
  ranges: XlsxAxisRange[] | undefined,
  maximum: number,
  maxCount = Number.POSITIVE_INFINITY,
): number[] {
  const indices: number[] = [];
  const seen = new Set<number>();
  for (const range of ranges ?? []) {
    const start = Math.max(0, Math.min(maximum, Math.floor(Math.min(range.start, range.end))));
    const end = Math.max(0, Math.min(maximum, Math.floor(Math.max(range.start, range.end))));
    for (let index = start; index <= end && indices.length < maxCount; index += 1) {
      if (seen.has(index)) continue;
      seen.add(index);
      indices.push(index);
    }
    if (indices.length >= maxCount) break;
  }
  return indices;
}

interface AxisSlice {
  index: number;
  sizePx: number;
}

function buildAxisSlices(
  start: number,
  end: number,
  hidden: ReadonlySet<number>,
  sizeOf: (index: number) => number,
): AxisSlice[] {
  const slices: AxisSlice[] = [];
  for (let index = start; index <= end; index += 1) {
    if (hidden.has(index)) continue;
    const sizePx = Math.max(0, sizeOf(index));
    if (sizePx <= 0) continue;
    slices.push({ index, sizePx });
  }
  return slices;
}

function sumSlices(slices: AxisSlice[], indices?: ReadonlySet<number>) {
  let total = 0;
  for (const slice of slices) {
    if (indices && !indices.has(slice.index)) continue;
    total += slice.sizePx;
  }
  return total;
}

// ── Pagination ───────────────────────────────────────────────────────────

function printAreasOf(sheet: XlsxPrintSheetData): XlsxCellRange[] {
  const areas = (sheet.printAreas ?? [])
    .filter((range) => range && Number.isFinite(range.start?.row) && Number.isFinite(range.start?.col))
    .map((range) => clampRange(range, MAX_SHEET_ROWS - 1, MAX_SHEET_COLUMNS - 1));
  if (areas.length > 0) return areas;
  if (sheet.minUsedRow >= 0 && sheet.minUsedCol >= 0 && sheet.maxUsedRow >= sheet.minUsedRow && sheet.maxUsedCol >= sheet.minUsedCol) {
    return [clampRange({
      end: { col: sheet.maxUsedCol, row: sheet.maxUsedRow },
      start: { col: sheet.minUsedCol, row: sheet.minUsedRow },
    }, Math.max(0, Math.min(MAX_SHEET_ROWS - 1, sheet.maxUsedRow)), Math.max(0, Math.min(MAX_SHEET_COLUMNS - 1, sheet.maxUsedCol)))];
  }
  return [{ end: { col: 0, row: 0 }, start: { col: 0, row: 0 } }];
}

function resolveEffectiveScale(
  pageSetup: XlsxPageSetup,
  bodyWidth: number,
  bodyHeight: number,
  titleWidth: number,
  titleHeight: number,
  contentWidth: number,
  contentHeight: number,
): number {
  if (!pageSetup.fitToPage) {
    return Math.max(MIN_SCALE_PERCENT, Math.min(MAX_SCALE_PERCENT, Math.floor(pageSetup.scale || 100)));
  }
  const fitToWidth = Math.max(0, Math.floor(pageSetup.fitToWidth || 0));
  const fitToHeight = Math.max(0, Math.floor(pageSetup.fitToHeight || 0));
  const availableWidth = fitToWidth > 0
    ? Math.max(1, contentWidth * fitToWidth - titleWidth * fitToWidth)
    : Number.POSITIVE_INFINITY;
  const availableHeight = fitToHeight > 0
    ? Math.max(1, (contentHeight - titleHeight) * fitToHeight)
    : Number.POSITIVE_INFINITY;
  const widthScale = bodyWidth > 0 ? availableWidth / bodyWidth : 1;
  const heightScale = bodyHeight > 0 ? availableHeight / bodyHeight : 1;
  const scale = Math.min(1, widthScale, heightScale);
  if (!Number.isFinite(scale) || scale <= 0) return 100;
  return Math.max(MIN_SCALE_PERCENT, Math.min(100, Math.floor(scale * 100)));
}

function splitAxis(
  slices: AxisSlice[],
  availablePx: number,
  scale: number,
  manualBreaks: readonly XlsxPageBreak[],
): number[] {
  // Boundaries are the first slice index of every page after the first.
  const boundaries = new Set<number>();
  const breakIds = new Set<number>();
  for (const pageBreak of manualBreaks) {
    const target = slices.find((slice) => slice.index >= pageBreak.id);
    if (target) breakIds.add(target.index);
  }
  let pageStart = 0;
  let used = 0;
  for (let index = 0; index < slices.length; index += 1) {
    const slice = slices[index]!;
    const atManualBreak = breakIds.has(slice.index);
    const scaledSize = slice.sizePx * scale;
    if ((atManualBreak || (used > 0 && used + scaledSize > availablePx + 0.5)) && index > pageStart) {
      boundaries.add(slice.index);
      pageStart = index;
      used = 0;
    }
    used += scaledSize;
  }
  return [...boundaries].sort((left, right) => left - right);
}

function axisPages(
  slices: AxisSlice[],
  boundaries: number[],
  fallbackIndex: number,
): Array<{ end: number; start: number }> {
  if (slices.length === 0) return [{ end: fallbackIndex, start: fallbackIndex }];
  const pages: Array<{ end: number; start: number }> = [];
  let currentStart = slices[0]!.index;
  for (const boundary of boundaries) {
    const current = slices.filter((slice) => slice.index < boundary && slice.index >= currentStart);
    if (current.length === 0) continue;
    pages.push({ end: current[current.length - 1]!.index, start: currentStart });
    currentStart = boundary;
  }
  const tail = slices.filter((slice) => slice.index >= currentStart);
  if (tail.length > 0) {
    pages.push({ end: tail[tail.length - 1]!.index, start: currentStart });
  }
  return pages.length > 0 ? pages : [{ end: slices[slices.length - 1]!.index, start: slices[0]!.index }];
}

function expandBoundariesAroundMerges(
  boundaries: number[],
  area: XlsxCellRange,
  merges: readonly XlsxCellRange[],
  axis: "row" | "col",
): number[] {
  if (boundaries.length === 0 || merges.length === 0) return boundaries;
  let next = [...boundaries];
  for (let pass = 0; pass < 16; pass += 1) {
    let changed = false;
    const expanded = next.map((boundary) => {
      const crossing = merges.find((merge) => {
        const start = merge.start[axis === "row" ? "row" : "col"];
        const end = merge.end[axis === "row" ? "row" : "col"];
        return start < boundary && end >= boundary;
      });
      if (!crossing) return boundary;
      const end = crossing.end[axis === "row" ? "row" : "col"];
      const areaEnd = area.end[axis === "row" ? "row" : "col"];
      if (end >= areaEnd) return boundary;
      changed = true;
      return end + 1;
    });
    next = [...new Set(expanded)].sort((left, right) => left - right);
    if (!changed) break;
  }
  return next;
}

function pageBreakIds(breaks: readonly XlsxPageBreak[] | undefined) {
  return (breaks ?? []).filter((pageBreak) => pageBreak.manual !== false);
}

/**
 * Computes Excel-style pagination for one worksheet.
 *
 * The function is platform-neutral: it uses only the sheet's row heights,
 * column widths, print set-up and merge metadata, so hosts, tests and the
 * viewer all derive identical page boundaries. Pages are returned as a flat
 * list of cell ranges plus the paper size used for each page.
 */
export function buildWorksheetPrintLayout(
  sheet: XlsxPrintSheetData,
  options: XlsxPrintLayoutOptions = {},
): XlsxPrintLayout {
  const pageSetup = sheet.pageSetup;
  const printOptions = sheet.printOptions;
  const pageMargins = sheet.pageMargins;
  const printTitles = sheet.printTitles ?? { columns: [], rows: [] };
  const paperSize = resolvePaperSize(pageSetup.paperSize);
  const paperPixels = resolvePaperPixelSize(paperSize, pageSetup.orientation);
  const marginPixels = resolvePageMarginsPx(pageMargins);
  const contentWidth = Math.max(1, paperPixels.widthPx - marginPixels.left - marginPixels.right);
  const contentHeight = Math.max(1, paperPixels.heightPx - marginPixels.top - marginPixels.bottom);
  const hiddenRows = new Set(sheet.hiddenRows ?? []);
  const hiddenCols = new Set(sheet.hiddenCols ?? []);
  const rowSize = (row: number) => resolveRowHeightPx(sheet, row);
  const colSize = (col: number) => resolveColumnWidthPx(sheet, col);
  const merges = sheet.mergedRegions ?? [];
  const maxPages = Math.max(1, Math.floor(options.maxPages ?? DEFAULT_MAX_PAGES));
  const areas = printAreasOf(sheet);

  const titleRows = axisRangeIndices(printTitles.rows, MAX_SHEET_ROWS - 1, MAX_PRINT_TITLE_AXIS_COUNT);
  const titleCols = axisRangeIndices(printTitles.columns, MAX_SHEET_COLUMNS - 1, MAX_PRINT_TITLE_AXIS_COUNT);

  // Titles are printed on every page. They reduce the body area of each page,
  // so the fit-to-page scale and the automatic split both account for them.
  let repeatedRowHeight = 0;
  for (const row of titleRows) {
    if (!hiddenRows.has(row)) repeatedRowHeight += rowSize(row);
  }
  let repeatedColWidth = 0;
  for (const col of titleCols) {
    if (!hiddenCols.has(col)) repeatedColWidth += colSize(col);
  }

  const pages: XlsxPrintPage[] = [];
  let truncated = false;

  for (const area of areas) {
    const rowSlices = buildAxisSlices(area.start.row, area.end.row, hiddenRows, rowSize);
    const colSlices = buildAxisSlices(area.start.col, area.end.col, hiddenCols, colSize);
    const bodyRowIndices = new Set(rowSlices.map((slice) => slice.index));
    const bodyColIndices = new Set(colSlices.map((slice) => slice.index));
    const titleRowsOutsideBody = titleRows.filter((row) => !bodyRowIndices.has(row) && !hiddenRows.has(row));
    const titleColsOutsideBody = titleCols.filter((col) => !bodyColIndices.has(col) && !hiddenCols.has(col));
    const titleHeightInPage = titleRowsOutsideBody.reduce((total, row) => total + rowSize(row), 0);
    const titleWidthInPage = titleColsOutsideBody.reduce((total, col) => total + colSize(col), 0);
    const bodyWidth = sumSlices(colSlices);
    const bodyHeight = sumSlices(rowSlices);
    const scale = resolveEffectiveScale(
      pageSetup,
      bodyWidth,
      bodyHeight,
      titleWidthInPage,
      titleHeightInPage,
      contentWidth,
      contentHeight,
    );
    const scaleFactor = scale / 100;
    const availableBodyWidth = Math.max(1, contentWidth - titleWidthInPage * scaleFactor);
    const availableBodyHeight = Math.max(1, contentHeight - titleHeightInPage * scaleFactor);
    const rowBoundaries = expandBoundariesAroundMerges(
      splitAxis(rowSlices, availableBodyHeight, scaleFactor, pageBreakIds(sheet.rowBreaks)),
      area,
      merges,
      "row",
    );
    const colBoundaries = expandBoundariesAroundMerges(
      splitAxis(colSlices, availableBodyWidth, scaleFactor, pageBreakIds(sheet.colBreaks)),
      area,
      merges,
      "col",
    );
    const rowPages = axisPages(rowSlices, rowBoundaries, area.start.row);
    const colPages = axisPages(colSlices, colBoundaries, area.start.col);
    const pageCombinations: Array<{ colPage: number; rowPage: number; range: XlsxCellRange }> = [];
    const maximumCombinations = Math.max(0, maxPages - pages.length);
    const order = pageSetup.pageOrder;
    if (order === "overThenDown") {
      outer: for (let rowPage = 0; rowPage < rowPages.length; rowPage += 1) {
        for (let colPage = 0; colPage < colPages.length; colPage += 1) {
          if (pageCombinations.length >= maximumCombinations) {
            truncated = true;
            break outer;
          }
          pageCombinations.push(buildPageCombination(rowPages[rowPage]!, colPages[colPage]!, rowPage, colPage));
        }
      }
    } else {
      outer: for (let colPage = 0; colPage < colPages.length; colPage += 1) {
        for (let rowPage = 0; rowPage < rowPages.length; rowPage += 1) {
          if (pageCombinations.length >= maximumCombinations) {
            truncated = true;
            break outer;
          }
          pageCombinations.push(buildPageCombination(rowPages[rowPage]!, colPages[colPage]!, rowPage, colPage));
        }
      }
    }

    for (const combination of pageCombinations) {
      if (pages.length >= maxPages) {
        truncated = true;
        break;
      }
      const index = pages.length;
      pages.push({
        colCount: combination.range.end.col - combination.range.start.col + 1,
        colPage: combination.colPage,
        contentHeightPx: contentHeight,
        contentWidthPx: contentWidth,
        heightPx: paperPixels.heightPx,
        index,
        orientation: pageSetup.orientation,
        pageNumber: (pageSetup.useFirstPageNumber ? pageSetup.firstPageNumber : 1) + index,
        paperSize: paperSize.id,
        paperSizeName: paperSize.name,
        printArea: area,
        range: combination.range,
        rowCount: combination.range.end.row - combination.range.start.row + 1,
        rowPage: combination.rowPage,
        scale,
        widthPx: paperPixels.widthPx,
      });
    }
    if (truncated) break;
  }

  return {
    fitToPage: pageSetup.fitToPage,
    pageMargins,
    pageSetup,
    pages,
    paperSize,
    printAreas: areas,
    printOptions,
    printTitles,
    scale: pages[0]?.scale ?? pageSetup.scale,
    sheetName: sheet.name,
    truncated,
    workbookSheetIndex: sheet.workbookSheetIndex,
  };
}

function buildPageCombination(
  rowPage: { end: number; start: number },
  colPage: { end: number; start: number },
  rowIndex: number,
  colIndex: number,
): { colPage: number; range: XlsxCellRange; rowPage: number } {
  return {
    colPage: colIndex,
    range: {
      end: { col: colPage.end, row: rowPage.end },
      start: { col: colPage.start, row: rowPage.start },
    },
    rowPage: rowIndex,
  };
}

/** Paginates every worksheet that has print data. */
export function buildWorkbookPrintLayouts(
  sheets: readonly XlsxPrintSheetData[],
  options: XlsxPrintLayoutOptions = {},
): XlsxPrintLayout[] {
  return sheets.map((sheet) => buildWorksheetPrintLayout(sheet, options));
}

// ── Page cell layout ─────────────────────────────────────────────────────

/**
 * Builds the rows, columns and merge spans of one printed page.
 *
 * Repeated print titles are included at the top/left of every page. A merge
 * that starts inside the page spans as many visible slots as it covers on that
 * page; slots covered by the merge are marked `covered` so a renderer can skip
 * them. A merge whose anchor is on another page is left unspanned because its
 * text belongs to that other page.
 */
export function buildPrintPageGrid(
  sheet: XlsxPrintSheetData,
  page: XlsxPrintPage,
  printTitles: XlsxPrintTitles = sheet.printTitles,
): XlsxPrintPageGrid {
  const hiddenRows = new Set(sheet.hiddenRows ?? []);
  const hiddenCols = new Set(sheet.hiddenCols ?? []);
  const rowSize = (row: number) => (hiddenRows.has(row) ? 0 : resolveRowHeightPx(sheet, row));
  const colSize = (col: number) => (hiddenCols.has(col) ? 0 : resolveColumnWidthPx(sheet, col));
  const bodyRows: number[] = [];
  for (let row = page.range.start.row; row <= page.range.end.row; row += 1) {
    if (rowSize(row) > 0) bodyRows.push(row);
  }
  const bodyCols: number[] = [];
  for (let col = page.range.start.col; col <= page.range.end.col; col += 1) {
    if (colSize(col) > 0) bodyCols.push(col);
  }

  const titleRows = axisRangeIndices(
    printTitles.rows,
    MAX_SHEET_ROWS - 1,
    MAX_PRINT_TITLE_AXIS_COUNT,
  );
  const titleCols = axisRangeIndices(
    printTitles.columns,
    MAX_SHEET_COLUMNS - 1,
    MAX_PRINT_TITLE_AXIS_COUNT,
  );
  const titleRowSet = new Set(titleRows);
  const titleColSet = new Set(titleCols);
  // Excel paints repeated titles first, in their own workbook order, then the
  // page body. A title row/column that also belongs to the body is not painted
  // twice.
  const rows = [...titleRows, ...bodyRows.filter((row) => !titleRowSet.has(row))];
  const columns = [...titleCols, ...bodyCols.filter((col) => !titleColSet.has(col))];
  const rowIndexInPage = new Map(rows.map((row, index) => [row, index]));
  const colIndexInPage = new Map(columns.map((col, index) => [col, index]));
  const merges = sheet.mergedRegions ?? [];
  const cells: XlsxPrintPageGridCell[] = [];

  for (const row of rows) {
    for (const col of columns) {
      const merge = merges.find((range) =>
        range.start.row <= row && row <= range.end.row
        && range.start.col <= col && col <= range.end.col
      );
      if (!merge) {
        cells.push({ col, colSpan: 1, covered: false, row, rowSpan: 1 });
        continue;
      }
      const anchorRow = merge.start.row;
      const anchorCol = merge.start.col;
      const anchorVisible = rowIndexInPage.has(anchorRow) && colIndexInPage.has(anchorCol);
      if (!anchorVisible) {
        cells.push({ col, colSpan: 1, covered: true, row, rowSpan: 1 });
        continue;
      }
      const endRow = Math.min(merge.end.row, rows[rows.length - 1] ?? merge.end.row);
      const endCol = Math.min(merge.end.col, columns[columns.length - 1] ?? merge.end.col);
      const rowSpan = rows.filter((candidate) => candidate >= anchorRow && candidate <= endRow).length;
      const colSpan = columns.filter((candidate) => candidate >= anchorCol && candidate <= endCol).length;
      cells.push({
        col,
        colSpan: Math.max(1, colSpan),
        covered: row !== anchorRow || col !== anchorCol,
        row,
        rowSpan: Math.max(1, rowSpan),
      });
    }
  }

  const gridRows: XlsxPrintPageGridAxis[] = rows.map((row) => ({ index: row, sizePx: rowSize(row) }));
  const gridColumns: XlsxPrintPageGridAxis[] = columns.map((col) => ({ index: col, sizePx: colSize(col) }));
  return { cells, columns: gridColumns, rows: gridRows };
}
