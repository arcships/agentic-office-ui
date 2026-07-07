import { ref, computed, watch, onUnmounted, shallowRef } from "vue";
import type { Workbook } from "@dukelib/sheets-wasm";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  applyChartSeriesFormula,
  buildChartSeriesFormula,
  loadWorkbookChartAssets,
  updateWorkbookChartAnchor,
  updateWorkbookChartDefinition,
  type WorkbookChartAssets
} from "@extend-ai/xlsx-core";
import { resolveWorkbookColor, resolveWorkbookFillStyle } from "@extend-ai/xlsx-core";
import {
  mergeWorkbookImageAssets,
  parseWorkbookImageAssets,
  pxToSheetColumnWidth,
  rectToImageAnchor,
  resolveContentSheetAxisPixels,
  resolveWorksheetDefaultColumnWidthPixels,
  resolveWorksheetDefaultRowHeightPixels,
  resolveWorksheetHiddenCols,
  resolveWorksheetHiddenRows,
  resolveWorksheetMergeMetadata,
  resolveSheetColumnWidthPixels,
  resolveRenderedSheetAxisPixels,
  resolveSheetRowHeightPixels,
  resizeImageRect,
  revokeWorkbookImageAssets,
  updateWorkbookImageAnchor,
  type WorkbookImageAssets,
  type WorkbookImageSheetOrigin
} from "@extend-ai/xlsx-core";
import { safeCalculate, tryRecalculate } from "@extend-ai/xlsx-core";
import { canUseConfiguredWasmSourceInWorker, getSheetsWasmModule } from "@extend-ai/xlsx-core";
import { XlsxWorkerClient } from "@extend-ai/xlsx-core";
import type {
  UseXlsxViewerControllerOptions,
  XlsxChart,
  XlsxChartElementSelection,
  XlsxChartsheet,
  XlsxCellAddress,
  XlsxCellRange,
  XlsxCellStyleInput,
  XlsxClipboardData,
  XlsxConditionalFormatRule,
  XlsxDataValidation,
  XlsxFormControl,
  XlsxFreezePanes,
  XlsxImage,
  XlsxImageRect,
  XlsxImageResizeHandlePosition,
  XlsxResolvedCellStyle,
  XlsxShape,
  XlsxSheetData,
  XlsxSheetVisibility,
  XlsxSparkline,
  XlsxThemePalette,
  XlsxTable,
  XlsxTableStyleDefinition,
  XlsxTableSortDirection,
  XlsxTableSortState,
  XlsxViewerController,
  XlsxWorkbookTab
} from "@extend-ai/xlsx-core";

const FORMULA_COUNT_THRESHOLD = 1000;
const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const CSV_MIME_TYPE = "text/csv;charset=utf-8";
const MIN_COL_WIDTH_PX = 30;
const MIN_ROW_HEIGHT_PX = 16;
const GRID_HEADER_HEIGHT = 24;
const GRID_ROW_HEADER_WIDTH = 40;
const HISTORY_LIMIT = 100;
const INTERNAL_CLIPBOARD_MIME = "application/x-react-xlsx-range+json";
const DEFAULT_DEFER_LOADING_ABOVE_BYTES = 0;
const DEFAULT_MAX_FILE_SIZE_BYTES = 25 * 1024 * 1024;
const MAX_INTERACTIVE_WORKSHEET_XML_BYTES = 200 * 1024 * 1024;
const MAX_INTERACTIVE_SHARED_STRINGS_BYTES = 50 * 1024 * 1024;
const MAX_INTERACTIVE_TOTAL_XML_BYTES = 256 * 1024 * 1024;
const EMU_PER_PIXEL = 9525;
const IMAGE_BATCH_ROW_COUNT = 256;
const DEFAULT_ZOOM_SCALE = 100;
const MIN_ZOOM_SCALE = 10;
const MAX_ZOOM_SCALE = 400;
const ZOOM_STEP = 10;
const DEFAULT_ZOOM_TAB_KEY = "__default__";

function normalizeWorksheetVisibility(value: unknown): XlsxSheetVisibility {
  return value === "hidden" || value === "veryHidden" ? value : "visible";
}

type IdleRequestHandle = number;

type IdleDeadline = {
  didTimeout: boolean;
  timeRemaining: () => number;
};

type IdleWindow = Window & {
  cancelIdleCallback?: (handle: IdleRequestHandle) => void;
  requestIdleCallback?: (
    callback: (deadline: IdleDeadline) => void,
    options?: {
      timeout: number;
    }
  ) => IdleRequestHandle;
};

type SnapshotHistoryEntry = {
  kind: "snapshot";
  activeCell: XlsxCellAddress | null;
  activeSheetIndex: number;
  bytes: Uint8Array;
  selection: XlsxCellRange | null;
};

type CellMutationState = {
  formula: string | null;
  style: unknown;
  value: unknown;
};

type CellEditHistoryEntry = {
  kind: "cell-edit";
  activeCellAfter: XlsxCellAddress | null;
  activeCellBefore: XlsxCellAddress | null;
  after: CellMutationState;
  before: CellMutationState;
  cell: XlsxCellAddress;
  selectionAfter: XlsxCellRange | null;
  selectionBefore: XlsxCellRange | null;
  sheetIndex: number;
};

type RangeCellMutation = {
  after: CellMutationState;
  before: CellMutationState;
  cell: XlsxCellAddress;
};

type RangeEditHistoryEntry = {
  kind: "range-edit";
  activeCellAfter: XlsxCellAddress | null;
  activeCellBefore: XlsxCellAddress | null;
  mutations: RangeCellMutation[];
  selectionAfter: XlsxCellRange | null;
  selectionBefore: XlsxCellRange | null;
  sheetIndex: number;
};

function clampZoomScale(zoomScale: number) {
  if (!Number.isFinite(zoomScale)) {
    return DEFAULT_ZOOM_SCALE;
  }

  return Math.min(MAX_ZOOM_SCALE, Math.max(MIN_ZOOM_SCALE, Math.round(zoomScale)));
}

function resolveDefaultZoomScale(activeTab: XlsxWorkbookTab | null, activeSheet: XlsxSheetData | null) {
  if (activeTab?.kind !== "sheet") {
    return DEFAULT_ZOOM_SCALE;
  }

  return clampZoomScale(activeSheet?.zoomScale ?? DEFAULT_ZOOM_SCALE);
}

function resolveWorksheetZoomScale(
  worksheet: ReturnType<Workbook["getSheet"]>,
  sheetState?: Record<string, unknown> | null
) {
  const candidates = [
    typeof sheetState?.zoomScale === "number" ? sheetState.zoomScale : undefined,
    typeof worksheet.zoomScale === "number" ? worksheet.zoomScale : undefined
  ];
  const value = candidates.find((entry): entry is number => typeof entry === "number" && Number.isFinite(entry) && entry > 0);
  return clampZoomScale(value ?? DEFAULT_ZOOM_SCALE);
}

function resolveNextZoomScale(currentZoomScale: number, direction: 1 | -1) {
  if (direction > 0) {
    return Math.min(
      MAX_ZOOM_SCALE,
      currentZoomScale % ZOOM_STEP === 0
        ? currentZoomScale + ZOOM_STEP
        : Math.ceil(currentZoomScale / ZOOM_STEP) * ZOOM_STEP
    );
  }

  return Math.max(
    MIN_ZOOM_SCALE,
    currentZoomScale % ZOOM_STEP === 0
      ? currentZoomScale - ZOOM_STEP
      : Math.floor(currentZoomScale / ZOOM_STEP) * ZOOM_STEP
  );
}

type WorksheetApiImageInfo = {
  altText?: unknown;
  height?: unknown;
  source?: unknown;
  width?: unknown;
};

type WorksheetDirectImageAnchorInfo = {
  fromCol?: unknown;
  fromColOffset?: unknown;
  fromRow?: unknown;
  fromRowOffset?: unknown;
  toCol?: unknown;
  toColOffset?: unknown;
  toRow?: unknown;
  toRowOffset?: unknown;
};

type WorksheetDirectImageInfo = {
  anchor?: unknown;
  data?: unknown;
  format?: unknown;
  id?: unknown;
  mediaPath?: unknown;
  name?: unknown;
  widthEmu?: unknown;
  heightEmu?: unknown;
};

type WorksheetDirectShapeParagraphRunInfo = {
  bold?: unknown;
  color?: unknown;
  fontFamily?: unknown;
  fontSizePt?: unknown;
  italic?: unknown;
  text?: unknown;
  underline?: unknown;
};

type WorksheetDirectShapeParagraphInfo = {
  align?: unknown;
  runs?: unknown;
};

type WorksheetDirectShapeTextBoxInfo = {
  horizontalAlign?: unknown;
  insetPx?: {
    bottom?: unknown;
    left?: unknown;
    right?: unknown;
    top?: unknown;
  } | null;
  verticalAlign?: unknown;
};

type WorksheetDirectShapeInfo = {
  anchor?: unknown;
  description?: unknown;
  fill?: {
    color?: unknown;
    none?: unknown;
    opacity?: unknown;
  } | null;
  flipH?: unknown;
  flipV?: unknown;
  geometry?: unknown;
  geometryAdjustments?: unknown;
  hyperlink?: unknown;
  id?: unknown;
  name?: unknown;
  paragraphs?: unknown;
  rotationDeg?: unknown;
  scaleX?: unknown;
  scaleY?: unknown;
  stroke?: {
    color?: unknown;
    dash?: unknown;
    headEndType?: unknown;
    none?: unknown;
    opacity?: unknown;
    tailEndType?: unknown;
    widthPx?: unknown;
  } | null;
  svgPath?: unknown;
  svgViewBox?: {
    height?: unknown;
    width?: unknown;
  } | null;
  text?: unknown;
  textBox?: WorksheetDirectShapeTextBoxInfo | null;
};

type WorksheetApiRowCell = {
  col?: unknown;
  image?: WorksheetApiImageInfo | null;
};

type WorksheetApiRow = {
  cells?: unknown;
  index?: unknown;
};

type WorksheetWithRowsBatch = ReturnType<Workbook["getSheet"]> & {
  getRowsBatch?: (startRow: number, maxRows: number, options?: unknown) => unknown;
};

type ZipEntryMetadata = {
  compressedSize: number;
  name: string;
  uncompressedSize: number;
};

type WorkbookPreflightResult = {
  largestWorksheetXmlBytes: number;
  sharedStringsBytes: number;
  totalWorksheetXmlBytes: number;
  tooLarge: boolean;
};

function formatBinaryBytes(bytes: number) {
  if (!Number.isFinite(bytes) || bytes <= 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }

  return `${value >= 10 || unitIndex === 0 ? value.toFixed(0) : value.toFixed(1)} ${units[unitIndex]}`;
}

function findZipEndOfCentralDirectoryOffset(bytes: Uint8Array) {
  const minLength = 22;
  if (bytes.byteLength < minLength) {
    return -1;
  }

  const searchStart = Math.max(0, bytes.byteLength - (0xffff + minLength));
  for (let offset = bytes.byteLength - minLength; offset >= searchStart; offset -= 1) {
    if (
      bytes[offset] === 0x50 &&
      bytes[offset + 1] === 0x4b &&
      bytes[offset + 2] === 0x05 &&
      bytes[offset + 3] === 0x06
    ) {
      return offset;
    }
  }

  return -1;
}

function readZipCentralDirectoryEntries(buffer: ArrayBuffer): ZipEntryMetadata[] | null {
  const bytes = new Uint8Array(buffer);
  const eocdOffset = findZipEndOfCentralDirectoryOffset(bytes);
  if (eocdOffset < 0) {
    return null;
  }

  const view = new DataView(buffer, bytes.byteOffset, bytes.byteLength);
  const centralDirectorySize = view.getUint32(eocdOffset + 12, true);
  const centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);
  const decoder = new TextDecoder();
  const entries: ZipEntryMetadata[] = [];

  let offset = centralDirectoryOffset;
  const endOffset = centralDirectoryOffset + centralDirectorySize;
  while (offset + 46 <= endOffset && offset + 46 <= bytes.byteLength) {
    if (view.getUint32(offset, true) !== 0x02014b50) {
      return null;
    }

    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const fileNameStart = offset + 46;
    const fileNameEnd = fileNameStart + fileNameLength;
    if (fileNameEnd > bytes.byteLength) {
      return null;
    }

    entries.push({
      compressedSize,
      name: decoder.decode(bytes.subarray(fileNameStart, fileNameEnd)),
      uncompressedSize
    });

    offset = fileNameEnd + extraLength + commentLength;
  }

  return entries;
}

function preflightWorkbookBuffer(buffer: ArrayBuffer): WorkbookPreflightResult | null {
  const entries = readZipCentralDirectoryEntries(buffer);
  if (!entries) {
    return null;
  }

  let largestWorksheetXmlBytes = 0;
  let totalWorksheetXmlBytes = 0;
  let sharedStringsBytes = 0;

  for (const entry of entries) {
    if (/^xl\/worksheets\/[^/]+\.xml$/i.test(entry.name)) {
      largestWorksheetXmlBytes = Math.max(largestWorksheetXmlBytes, entry.uncompressedSize);
      totalWorksheetXmlBytes += entry.uncompressedSize;
      continue;
    }

    if (entry.name === "xl/sharedStrings.xml") {
      sharedStringsBytes = entry.uncompressedSize;
    }
  }

  const tooLarge =
    largestWorksheetXmlBytes > MAX_INTERACTIVE_WORKSHEET_XML_BYTES ||
    sharedStringsBytes > MAX_INTERACTIVE_SHARED_STRINGS_BYTES ||
    totalWorksheetXmlBytes + sharedStringsBytes > MAX_INTERACTIVE_TOTAL_XML_BYTES;

  return {
    largestWorksheetXmlBytes,
    sharedStringsBytes,
    tooLarge,
    totalWorksheetXmlBytes
  };
}

function createWorkbookTooLargeError(preflight: WorkbookPreflightResult) {
  return new Error(
    `XLSX is too large to preview interactively. `
    + `Largest worksheet XML: ${formatBinaryBytes(preflight.largestWorksheetXmlBytes)}; `
    + `shared strings: ${formatBinaryBytes(preflight.sharedStringsBytes)}.`
  );
}

export class XlsxFileSizeLimitExceededError extends Error {
  fileSizeBytes: number;
  maxFileSizeBytes: number;

  constructor(fileSizeBytes: number, maxFileSizeBytes: number) {
    super(
      `XLSX file size ${formatBinaryBytes(fileSizeBytes)} exceeds the configured limit of ${formatBinaryBytes(maxFileSizeBytes)}.`
    );
    this.name = "XlsxFileSizeLimitExceededError";
    this.fileSizeBytes = fileSizeBytes;
    this.maxFileSizeBytes = maxFileSizeBytes;
  }
}

type HistoryEntry = SnapshotHistoryEntry | CellEditHistoryEntry | RangeEditHistoryEntry;

type ClipboardMatrixCell = {
  colOffset: number;
  formula: string | null;
  rowOffset: number;
  value: string;
};

type ClipboardMerge = {
  colSpan: number;
  colOffset: number;
  rowOffset: number;
  rowSpan: number;
};

type ClipboardPayload = {
  cells: ClipboardMatrixCell[];
  cols: number;
  merges: ClipboardMerge[];
  rows: number;
};

function resolveDisplayFileName(src?: string, fileName?: string): string {
  if (typeof fileName === "string" && fileName.trim().length > 0) {
    return fileName.trim();
  }

  if (!src) {
    return "Workbook.xlsx";
  }

  const pathWithoutQuery = src.split("?")[0] ?? "";
  const pathSegments = pathWithoutQuery.split("/");
  const lastSegment = pathSegments[pathSegments.length - 1] ?? "";

  if (!lastSegment) {
    return "Workbook.xlsx";
  }

  try {
    return decodeURIComponent(lastSegment);
  } catch {
    return lastSegment;
  }
}

function resolveSheetDisplayUsedRange(
  usedRange: [number, number, number, number],
  sheetState?: {
    maxContentCol?: number;
    maxContentRow?: number;
    maxHorizontalMergeEndCol?: number;
    maxVerticalMergeEndRow?: number;
    minContentCol?: number;
    minContentRow?: number;
  } | null
): [number, number, number, number] {
  const [minRow, minCol, maxRow, maxCol] = usedRange;
  const maxContentRow = sheetState?.maxContentRow ?? -1;
  const maxContentCol = sheetState?.maxContentCol ?? -1;
  const maxVerticalMergeEndRow = sheetState?.maxVerticalMergeEndRow ?? -1;
  const maxHorizontalMergeEndCol = sheetState?.maxHorizontalMergeEndCol ?? -1;
  const maxMeaningfulRow = Math.max(maxContentRow, maxVerticalMergeEndRow);
  const maxMeaningfulCol = Math.max(maxContentCol, maxHorizontalMergeEndCol);

  if (maxMeaningfulRow < 0 && maxMeaningfulCol < 0) {
    return usedRange;
  }

  return [
    sheetState?.minContentRow !== undefined && sheetState.minContentRow >= 0 ? Math.min(minRow, sheetState.minContentRow) : minRow,
    sheetState?.minContentCol !== undefined && sheetState.minContentCol >= 0 ? Math.min(minCol, sheetState.minContentCol) : minCol,
    maxMeaningfulRow >= 0
      ? (maxContentRow >= 0 ? Math.min(maxRow, maxMeaningfulRow) : Math.max(maxRow, maxMeaningfulRow))
      : maxRow,
    maxMeaningfulCol >= 0
      ? (maxContentCol >= 0 ? Math.min(maxCol, maxMeaningfulCol) : Math.max(maxCol, maxMeaningfulCol))
      : maxCol
  ];
}

function buildSheetList(
  workbook: Workbook,
  sheetStatesByWorkbookSheetIndex?: Array<{
    cachedFormulaValues?: Record<string, string>;
    columnWidthCharacterWidthPx?: number;
    colWidthOverridesPx?: Record<number, number>;
    colStyleIds?: Record<number, number>;
    conditionalFormatRules?: XlsxConditionalFormatRule[];
    defaultColWidthPx?: number;
    defaultRowHeightPx?: number;
    hasHorizontalMerges?: boolean;
    hasVerticalMerges?: boolean;
    maxHorizontalMergeEndCol?: number;
    maxVerticalMergeEndRow?: number;
    maxContentCol?: number;
    maxContentRow?: number;
    minContentCol?: number;
    minContentRow?: number;
    hiddenCols?: number[];
    hiddenRows?: number[];
    rowHeightOverridesPx?: Record<number, number>;
    rowStyleIds?: Record<number, number>;
    showGridLines: boolean;
    sparklines?: XlsxSparkline[];
  } | null>,
  themePalette?: XlsxThemePalette | null,
  styleById?: Record<number, XlsxResolvedCellStyle> | null,
  namedCellStyleByName?: Record<string, XlsxResolvedCellStyle> | null,
  tableStyleByName?: Record<string, XlsxTableStyleDefinition> | null,
  showHiddenSheets = false
): XlsxSheetData[] {
  const sheets: XlsxSheetData[] = [];

  for (let index = 0; index < workbook.sheetCount; index += 1) {
    const worksheet = workbook.getSheet(index);
    const sheetState = sheetStatesByWorkbookSheetIndex?.[index] ?? null;
    const mergeMetadata = resolveWorksheetMergeMetadata(worksheet);
    const effectiveSheetState = {
      ...sheetState,
      ...mergeMetadata
    };
    const defaultColWidthPx = resolveWorksheetDefaultColumnWidthPixels(
      worksheet,
      sheetState?.columnWidthCharacterWidthPx,
      sheetState?.defaultColWidthPx ?? DEFAULT_COL_WIDTH
    );
    const defaultRowHeightPx = resolveWorksheetDefaultRowHeightPixels(
      worksheet,
      sheetState?.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT
    );
    const visibility = normalizeWorksheetVisibility(worksheet.visibility);
    if (!showHiddenSheets && visibility !== "visible") {
      continue;
    }

    const resolveColumnWidthPx = (col: number) => {
      const width = worksheet.getColumnWidth(col);
      if (width !== undefined && width !== null) {
        return resolveSheetColumnWidthPixels(width, sheetState?.columnWidthCharacterWidthPx);
      }

      return sheetState?.colWidthOverridesPx?.[col] ?? defaultColWidthPx;
    };

    const resolveRowHeightPx = (row: number) => {
      const height = worksheet.getRowHeight(row);
      if (height !== undefined && height !== null) {
        return Math.max(Math.round(height * 1.33), MIN_ROW_HEIGHT_PX);
      }

      return sheetState?.rowHeightOverridesPx?.[row] ?? defaultRowHeightPx;
    };

    const usedRange = worksheet.usedRange() as [number, number, number, number] | null;
    if (!usedRange) {
      sheets.push({
        cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
        columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
        colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
        colStyleIds: sheetState?.colStyleIds ?? {},
        conditionalFormatRules: sheetState?.conditionalFormatRules ?? [],
        dataValidations: parseWorksheetDataValidations(worksheet),
        defaultColWidthPx,
        defaultRowHeightPx,
        freezePanes: parseWorksheetFreezePanes(worksheet),
        hasHorizontalMerges: mergeMetadata.hasHorizontalMerges,
        hasVerticalMerges: mergeMetadata.hasVerticalMerges,
        maxHorizontalMergeEndCol: mergeMetadata.maxHorizontalMergeEndCol,
        maxVerticalMergeEndRow: mergeMetadata.maxVerticalMergeEndRow,
        hiddenCols: [],
        hiddenRows: [],
        minUsedCol: -1,
        minUsedRow: -1,
        maxUsedCol: -1,
        maxUsedRow: -1,
        name: worksheet.name,
        visibility,
        namedCellStyleByName: namedCellStyleByName ?? {},
        rowCount: 0,
        colCount: 0,
        rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
        rowStyleIds: sheetState?.rowStyleIds ?? {},
        styleById: styleById ?? {},
        sparklines: sheetState?.sparklines ?? [],
        tableStyleByName: tableStyleByName ?? {},
        visibleRows: [],
        visibleCols: [],
        colWidths: [],
        rowHeights: [],
        showGridLines: sheetState?.showGridLines ?? true,
        themePalette: themePalette ?? { colorsByIndex: {} },
        workbookSheetIndex: index,
        zoomScale: resolveWorksheetZoomScale(worksheet, sheetState)
      });
      continue;
    }

    const [minRow, minCol, maxRow, maxCol] = resolveSheetDisplayUsedRange(usedRange, effectiveSheetState);
    let visibleRowsCache: number[] | null = null;
    let visibleColsCache: number[] | null = null;
    let rowHeightsCache: number[] | null = null;
    let colWidthsCache: number[] | null = null;

    const getVisibleRows = () => {
      if (visibleRowsCache) {
        return visibleRowsCache;
      }

      const nextVisibleRows: number[] = [];
      for (let row = 0; row <= maxRow; row += 1) {
        if (!worksheet.isRowHidden(row)) {
          nextVisibleRows.push(row);
        }
      }

      visibleRowsCache = nextVisibleRows;
      return nextVisibleRows;
    };

    const getVisibleCols = () => {
      if (visibleColsCache) {
        return visibleColsCache;
      }

      const nextVisibleCols: number[] = [];
      for (let col = 0; col <= maxCol; col += 1) {
        if (!worksheet.isColumnHidden(col)) {
          nextVisibleCols.push(col);
        }
      }

      visibleColsCache = nextVisibleCols;
      return nextVisibleCols;
    };

    const getRowHeights = () => {
      if (rowHeightsCache) {
        return rowHeightsCache;
      }

      rowHeightsCache = getVisibleRows().map(resolveRowHeightPx);
      return rowHeightsCache;
    };

    const getColWidths = () => {
      if (colWidthsCache) {
        return colWidthsCache;
      }

      colWidthsCache = getVisibleCols().map(resolveColumnWidthPx);
      return colWidthsCache;
    };

    const sheet: XlsxSheetData = {
      cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
      columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
      colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
      colStyleIds: sheetState?.colStyleIds ?? {},
      conditionalFormatRules: sheetState?.conditionalFormatRules ?? [],
      dataValidations: parseWorksheetDataValidations(worksheet),
      defaultColWidthPx,
      defaultRowHeightPx,
      freezePanes: parseWorksheetFreezePanes(worksheet),
      hasHorizontalMerges: mergeMetadata.hasHorizontalMerges,
      hasVerticalMerges: mergeMetadata.hasVerticalMerges,
      maxHorizontalMergeEndCol: mergeMetadata.maxHorizontalMergeEndCol,
      maxVerticalMergeEndRow: mergeMetadata.maxVerticalMergeEndRow,
      hiddenCols: resolveWorksheetHiddenCols(worksheet, maxCol),
      hiddenRows: resolveWorksheetHiddenRows(worksheet, maxRow),
      minUsedCol: minCol,
      minUsedRow: minRow,
      maxUsedCol: maxCol,
      maxUsedRow: maxRow,
      name: worksheet.name,
      visibility,
      namedCellStyleByName: namedCellStyleByName ?? {},
      rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
      rowStyleIds: sheetState?.rowStyleIds ?? {},
      showGridLines: sheetState?.showGridLines ?? true,
      styleById: styleById ?? {},
      sparklines: sheetState?.sparklines ?? [],
      tableStyleByName: tableStyleByName ?? {},
      themePalette: themePalette ?? { colorsByIndex: {} },
      workbookSheetIndex: index,
      zoomScale: resolveWorksheetZoomScale(worksheet, sheetState),
      get rowCount() {
        return getVisibleRows().length;
      },
      get colCount() {
        return getVisibleCols().length;
      },
      get visibleRows() {
        return getVisibleRows();
      },
      get visibleCols() {
        return getVisibleCols();
      },
      get colWidths() {
        return getColWidths();
      },
      get rowHeights() {
        return getRowHeights();
      }
    };

    sheets.push(sheet);
  }

  return sheets;
}

function buildVisibleSheetIndexMap(sheets: XlsxSheetData[]) {
  return new Map(sheets.map((sheet, index) => [sheet.workbookSheetIndex, index]));
}

function resolveInheritedCellStyle(sheet: XlsxSheetData | null | undefined, row: number, col: number) {
  if (!sheet) {
    return null;
  }

  const rowStyleId = sheet.rowStyleIds[row];
  if (rowStyleId !== undefined) {
    return sheet.styleById[rowStyleId] ?? null;
  }

  const colStyleId = sheet.colStyleIds[col];
  if (colStyleId !== undefined) {
    return sheet.styleById[colStyleId] ?? null;
  }

  return null;
}

function columnLabel(col: number): string {
  let label = "";
  let nextValue = col;

  while (nextValue >= 0) {
    label = String.fromCharCode(65 + (nextValue % 26)) + label;
    nextValue = Math.floor(nextValue / 26) - 1;
  }

  return label;
}

function cellAddressToA1(cell: XlsxCellAddress): string {
  return `${columnLabel(cell.col)}${cell.row + 1}`;
}

function parseA1CellReference(reference: string): XlsxCellAddress | null {
  const match = /^([A-Z]+)(\d+)$/i.exec(reference.trim());
  if (!match) {
    return null;
  }

  const [, columnPart, rowPart] = match;
  let col = 0;
  for (const char of columnPart.toUpperCase()) {
    col = col * 26 + (char.charCodeAt(0) - 64);
  }

  return {
    col: col - 1,
    row: Number(rowPart) - 1
  };
}

function parseA1RangeReference(reference: string): XlsxCellRange | null {
  const [startRef, endRef = startRef] = reference.split(":");
  const start = parseA1CellReference(startRef ?? "");
  const end = parseA1CellReference(endRef ?? "");
  if (!start || !end) {
    return null;
  }

  return normalizeRange({ start, end });
}

function parseWorksheetFreezePanes(worksheet: ReturnType<Workbook["getSheet"]>): XlsxFreezePanes | null {
  const rawFreezePanes = worksheet.freezePanes as Record<string, unknown> | null | undefined;
  const row = typeof rawFreezePanes?.row === "number" && rawFreezePanes.row >= 0 ? rawFreezePanes.row : null;
  const col = typeof rawFreezePanes?.col === "number" && rawFreezePanes.col >= 0 ? rawFreezePanes.col : null;
  if (row === null && col === null) {
    return null;
  }

  return {
    col: col ?? 0,
    row: row ?? 0
  };
}

function parseWorksheetDataValidations(worksheet: ReturnType<Workbook["getSheet"]>): XlsxDataValidation[] {
  const rawDataValidations = Array.isArray(worksheet.dataValidations) ? worksheet.dataValidations : [];

  return rawDataValidations.flatMap((entry) => {
    if (!entry || typeof entry !== "object") {
      return [];
    }

    const validation = entry as Record<string, unknown>;
    const ranges = Array.isArray(validation.ranges)
      ? validation.ranges.flatMap((range) => {
          if (typeof range !== "string") {
            return [];
          }

          const parsedRange = parseA1RangeReference(range);
          return parsedRange ? [parsedRange] : [];
        })
      : [];
    const validationType = typeof validation.validationType === "string" ? validation.validationType : null;
    if (!validationType || ranges.length === 0) {
      return [];
    }

    return [{
      allowBlank: typeof validation.allowBlank === "boolean" ? validation.allowBlank : undefined,
      errorMessage: typeof validation.errorMessage === "string" ? validation.errorMessage : undefined,
      errorStyle: typeof validation.errorStyle === "string" ? validation.errorStyle : undefined,
      inputMessage: typeof validation.inputMessage === "string" ? validation.inputMessage : undefined,
      listSource: typeof validation.listSource === "string" ? validation.listSource : undefined,
      ranges,
      showDropdown: typeof validation.showDropdown === "boolean" ? validation.showDropdown : undefined,
      showErrorAlert: typeof validation.showErrorAlert === "boolean" ? validation.showErrorAlert : undefined,
      showInputMessage: typeof validation.showInputMessage === "boolean" ? validation.showInputMessage : undefined,
      validationType
    } satisfies XlsxDataValidation];
  });
}

function normalizeRange(range: XlsxCellRange): XlsxCellRange {
  return {
    start: {
      row: Math.min(range.start.row, range.end.row),
      col: Math.min(range.start.col, range.end.col)
    },
    end: {
      row: Math.max(range.start.row, range.end.row),
      col: Math.max(range.start.col, range.end.col)
    }
  };
}

function rangeToA1(range: XlsxCellRange): string {
  const normalized = normalizeRange(range);
  const start = cellAddressToA1(normalized.start);
  const end = cellAddressToA1(normalized.end);
  return start === end ? start : `${start}:${end}`;
}

function rangeContainsCell(range: XlsxCellRange, cell: XlsxCellAddress): boolean {
  const normalized = normalizeRange(range);
  return (
    cell.row >= normalized.start.row &&
    cell.row <= normalized.end.row &&
    cell.col >= normalized.start.col &&
    cell.col <= normalized.end.col
  );
}

function mapWorksheetTables(worksheet: ReturnType<Workbook["getSheet"]> | null): XlsxTable[] {
  const rawTables = (worksheet?.tables ?? []) as Array<Record<string, unknown>>;
  return rawTables.flatMap((table, index) => {
    const rawColumns = Array.isArray(table.columns) ? table.columns : [];
    const rawName = typeof table.name === "string" ? table.name : `Table${index + 1}`;
    const rawDisplayName =
      typeof table.displayName === "string"
        ? table.displayName
        : typeof table.name === "string"
          ? table.name
          : `Table ${index + 1}`;
    const rawReference = typeof table.reference === "string" ? table.reference : "";
    const reference = rawReference;
    const parsedRange = parseA1RangeReference(reference);
    if (!parsedRange) {
      return [];
    }

    return [{
      columns: rawColumns.map((column, columnIndex) => ({
        id: typeof (column as { id?: unknown }).id === "number" ? ((column as { id?: number }).id ?? columnIndex + 1) : columnIndex + 1,
        index: columnIndex,
        name: typeof (column as { name?: unknown }).name === "string" ? ((column as { name?: string }).name ?? `Column ${columnIndex + 1}`) : `Column ${columnIndex + 1}`
      })),
      displayName: rawDisplayName,
      end: parsedRange.end,
      headerRowCount: resolveWorkbookTableCount(table.headerRowCount, 1),
      headerRowCellStyle: typeof table.headerRowCellStyle === "string" ? table.headerRowCellStyle : undefined,
      name: rawName,
      reference,
      start: parsedRange.start,
      styleInfo: table.styleInfo as XlsxTable["styleInfo"] | undefined,
      totalsRowCount: resolveWorkbookTableCount(table.totalsRowCount, 0),
      totalsRowShown: resolveWorkbookTableBoolean(table.totalsRowShown)
    }];
  });
}

function resolveWorkbookTableCount(value: unknown, fallback: number) {
  if (typeof value === "number" && Number.isFinite(value) && value >= 0) {
    return value;
  }

  if (typeof value === "string") {
    const parsed = Number.parseInt(value, 10);
    if (Number.isFinite(parsed) && parsed >= 0) {
      return parsed;
    }
  }

  return fallback;
}

function resolveWorkbookTableBoolean(value: unknown) {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "number") {
    return value !== 0;
  }

  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (normalized === "0" || normalized === "false" || normalized === "") {
      return false;
    }
    if (normalized === "1" || normalized === "true") {
      return true;
    }
  }

  return false;
}

function fileStem(fileName: string): string {
  const normalized = fileName.trim();
  const lastDot = normalized.lastIndexOf(".");
  return lastDot > 0 ? normalized.slice(0, lastDot) : normalized;
}

function pxToSheetRowHeight(heightPx: number): number {
  return Math.max(heightPx, MIN_ROW_HEIGHT_PX) / 1.33;
}

function cssColor(color: Record<string, unknown> | undefined): string | null {
  if (!color?.hex) {
    return null;
  }

  const hex = String(color.hex);
  const rgb = hex.length === 8 ? hex.slice(2) : hex;
  return `#${rgb}`;
}

function mapBorder(edge: { style: string; color?: { hex?: string } }): string {
  const color = cssColor(edge.color as Record<string, unknown> | undefined) ?? "#000000";
  const widthMap: Record<string, string> = {
    dashed: "1px",
    dotted: "1px",
    double: "3px",
    hair: "1px",
    medium: "2px",
    thick: "3px",
    thin: "1px"
  };
  const styleMap: Record<string, string> = {
    dashDot: "dashed",
    dashDotDot: "dotted",
    dashed: "dashed",
    dotted: "dotted",
    double: "double",
    hair: "solid",
    medium: "solid",
    mediumDashDot: "dashed",
    mediumDashDotDot: "dotted",
    mediumDashed: "dashed",
    slantDashDot: "dashed",
    thick: "solid",
    thin: "solid"
  };

  return `${widthMap[edge.style] ?? "1px"} ${styleMap[edge.style] ?? "solid"} ${color}`;
}

function decodeHtmlEntities(value: string): string {
  return value
    .replace(/&quot;/g, "\"")
    .replace(/&#34;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function cloneBytes(bytes: Uint8Array): Uint8Array {
  const nextBytes = new Uint8Array(bytes.byteLength);
  nextBytes.set(bytes);
  return nextBytes;
}

function sanitizeSavedWorkbookBytes(bytes: Uint8Array): Uint8Array {
  try {
    const archive = unzipSync(bytes);
    const stylesEntry = archive["xl/styles.xml"];
    if (stylesEntry) {
      const stylesXml = strFromU8(stylesEntry)
        .replace(/&amp;quot;/g, "&quot;")
        .replace(/&amp;apos;/g, "&apos;");
      archive["xl/styles.xml"] = strToU8(stylesXml);
    }

    return zipSync(archive, { level: 6 });
  } catch {
    return cloneBytes(bytes);
  }
}

function pushHistoryEntry(stack: HistoryEntry[], entry: HistoryEntry) {
  stack.push(entry);
  if (stack.length > HISTORY_LIMIT) {
    stack.shift();
  }
}

function normalizeCellValue(value: unknown) {
  return value ?? "";
}

function cloneCellStyle(style: unknown): unknown {
  if (!style || typeof style !== "object") {
    return style;
  }

  if (typeof structuredClone === "function") {
    try {
      return structuredClone(style);
    } catch {
      // Fall through to the JSON clone below.
    }
  }

  try {
    return JSON.parse(JSON.stringify(style));
  } catch {
    return style;
  }
}

function coerceUserEnteredValue(value: string): unknown {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }

  if (trimmed.startsWith("'")) {
    return trimmed.slice(1);
  }

  if (/^(true|false)$/i.test(trimmed)) {
    return trimmed.toLowerCase() === "true";
  }

  if (/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(trimmed)) {
    const parsed = Number(trimmed);
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }

  return value;
}

function applyCellMutationState(
  worksheet: ReturnType<Workbook["getSheet"]>,
  cell: XlsxCellAddress,
  state: CellMutationState
) {
  if (state.formula) {
    worksheet.setFormula(cellAddressToA1(cell), state.formula);
  } else {
    worksheet.setCell(cellAddressToA1(cell), normalizeCellValue(state.value));
  }

  if (state.style && typeof state.style === "object") {
    worksheet.setCellStyleAt(cell.row, cell.col, state.style);
  }
}

function escapeHtml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function parseClipboardText(text: string): string[][] {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const rows = normalized.split("\n");

  if (rows.length > 1 && rows[rows.length - 1] === "") {
    rows.pop();
  }

  return rows.map((row) => row.split("\t"));
}

function createAbortError() {
  if (typeof DOMException !== "undefined") {
    return new DOMException("Aborted", "AbortError");
  }

  const error = new Error("Aborted");
  error.name = "AbortError";
  return error;
}

function isAbortError(error: unknown) {
  return error instanceof Error && error.name === "AbortError";
}

async function resolveWorkbookBuffer(
  { file, src }: UseXlsxViewerControllerOptions,
  signal?: AbortSignal
): Promise<ArrayBuffer> {
  let buffer: ArrayBuffer;

  if (signal?.aborted) {
    throw createAbortError();
  }

  if (file) {
    buffer = file;
  } else if (src) {
    let response: Response;
    try {
      response = await fetch(src, { signal });
    } catch (error) {
      if (isAbortError(error)) {
        throw error;
      }

      throw new Error(
        "Failed to fetch workbook. The remote URL may be blocked by CORS, unavailable, or not directly downloadable from the browser."
      );
    }
    if (!response.ok) {
      throw new Error(`Failed to fetch workbook (status ${response.status})`);
    }
    buffer = await response.arrayBuffer();
  } else {
    throw new Error("Either `file` or `src` must be provided.");
  }

  return buffer;
}

async function parseWorkbookBuffer(buffer: ArrayBuffer): Promise<{
  shouldAutoCalculate: boolean;
  workbook: Workbook;
}> {
  const wasmModule = await getSheetsWasmModule();
  const initialWorkbook = wasmModule.Workbook.fromBytes(new Uint8Array(buffer));
  let totalFormulas = 0;

  for (let index = 0; index < initialWorkbook.sheetCount; index += 1) {
    totalFormulas += initialWorkbook.getSheet(index).formulaCount;
  }

  const shouldAutoCalculate = totalFormulas <= FORMULA_COUNT_THRESHOLD;
  if (!shouldAutoCalculate) {
    return { shouldAutoCalculate, workbook: initialWorkbook };
  }

  const result = safeCalculate(initialWorkbook, {
    reparse: () => wasmModule.Workbook.fromBytes(new Uint8Array(buffer))
  });

  return {
    shouldAutoCalculate: result.calculated,
    workbook: result.workbook
  };
}

function scheduleLowPriorityTask(task: () => void) {
  if (typeof window === "undefined") {
    const timeoutHandle = setTimeout(task, 0);
    return () => clearTimeout(timeoutHandle);
  }

  const idleWindow = window as IdleWindow;
  if (typeof idleWindow.requestIdleCallback === "function") {
    const idleHandle = idleWindow.requestIdleCallback(() => {
      task();
    }, { timeout: 120 });
    return () => {
      idleWindow.cancelIdleCallback?.(idleHandle);
    };
  }

  const timeoutHandle = window.setTimeout(task, 0);
  return () => window.clearTimeout(timeoutHandle);
}

function asFiniteNumber(value: unknown) {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function inferImageMimeType(source: string) {
  if (source.startsWith("data:")) {
    const separatorIndex = source.indexOf(";");
    if (separatorIndex > 5) {
      return source.slice(5, separatorIndex);
    }
  }

  const normalized = source.split("?")[0]?.toLowerCase() ?? "";
  if (normalized.endsWith(".gif")) {
    return "image/gif";
  }
  if (normalized.endsWith(".jpg") || normalized.endsWith(".jpeg")) {
    return "image/jpeg";
  }
  if (normalized.endsWith(".svg")) {
    return "image/svg+xml";
  }
  if (normalized.endsWith(".webp")) {
    return "image/webp";
  }

  return "image/png";
}

function inferWorksheetDirectImageMimeType(info: WorksheetDirectImageInfo) {
  const format = typeof info.format === "string" ? info.format.trim().toLowerCase() : "";
  if (format === "gif") {
    return "image/gif";
  }
  if (format === "jpg" || format === "jpeg") {
    return "image/jpeg";
  }
  if (format === "svg") {
    return "image/svg+xml";
  }
  if (format === "webp") {
    return "image/webp";
  }
  if (format === "png") {
    return "image/png";
  }

  const mediaPath = typeof info.mediaPath === "string" ? info.mediaPath : "";
  if (mediaPath) {
    return inferImageMimeType(mediaPath);
  }

  return "image/png";
}

function createWorksheetDirectImageSource(
  data: unknown,
  mimeType: string,
  objectUrls: string[]
) {
  const bytes = data instanceof Uint8Array
    ? data
    : Array.isArray(data)
      ? Uint8Array.from(data.filter((value): value is number => typeof value === "number"))
      : null;
  if (!bytes || bytes.byteLength === 0) {
    return null;
  }

  const blobBuffer = new Uint8Array(bytes.byteLength);
  blobBuffer.set(bytes);
  const objectUrl = URL.createObjectURL(new Blob([blobBuffer.buffer], { type: mimeType }));
  objectUrls.push(objectUrl);
  return objectUrl;
}

function buildWorksheetDirectImageAnchor(
  rawAnchor: unknown,
  widthEmu: number,
  heightEmu: number
): XlsxImage["anchor"] {
  const anchor = rawAnchor && typeof rawAnchor === "object" ? rawAnchor as WorksheetDirectImageAnchorInfo : {};
  const fromCol = asFiniteNumber(anchor.fromCol) ?? 0;
  const fromRow = asFiniteNumber(anchor.fromRow) ?? 0;
  const fromColOffset = asFiniteNumber(anchor.fromColOffset) ?? 0;
  const fromRowOffset = asFiniteNumber(anchor.fromRowOffset) ?? 0;
  const toCol = asFiniteNumber(anchor.toCol);
  const toRow = asFiniteNumber(anchor.toRow);
  const toColOffset = asFiniteNumber(anchor.toColOffset) ?? 0;
  const toRowOffset = asFiniteNumber(anchor.toRowOffset) ?? 0;

  if (toCol !== null && toRow !== null) {
    return {
      from: {
        col: Math.max(0, Math.round(fromCol)),
        colOffsetEmu: Math.max(0, Math.round(fromColOffset)),
        row: Math.max(0, Math.round(fromRow)),
        rowOffsetEmu: Math.max(0, Math.round(fromRowOffset))
      },
      kind: "two-cell",
      to: {
        col: Math.max(0, Math.round(toCol)),
        colOffsetEmu: Math.max(0, Math.round(toColOffset)),
        row: Math.max(0, Math.round(toRow)),
        rowOffsetEmu: Math.max(0, Math.round(toRowOffset))
      }
    };
  }

  return {
    from: {
      col: Math.max(0, Math.round(fromCol)),
      colOffsetEmu: Math.max(0, Math.round(fromColOffset)),
      row: Math.max(0, Math.round(fromRow)),
      rowOffsetEmu: Math.max(0, Math.round(fromRowOffset))
    },
    kind: "one-cell",
    sizeEmu: {
      cx: Math.max(EMU_PER_PIXEL, Math.round(widthEmu)),
      cy: Math.max(EMU_PER_PIXEL, Math.round(heightEmu))
    }
  };
}

function normalizeWorksheetDirectShapeParagraphs(rawParagraphs: unknown, fallbackText: unknown): XlsxShape["paragraphs"] {
  const normalizedParagraphs: XlsxShape["paragraphs"] = [];

  if (Array.isArray(rawParagraphs)) {
    for (const entry of rawParagraphs) {
        const paragraph = entry && typeof entry === "object" ? entry as WorksheetDirectShapeParagraphInfo : {};
        const runs: XlsxShape["paragraphs"][number]["runs"] = [];
        if (Array.isArray(paragraph.runs)) {
          for (const runEntry of paragraph.runs) {
            const run = runEntry && typeof runEntry === "object" ? runEntry as WorksheetDirectShapeParagraphRunInfo : {};
            const text = typeof run.text === "string" ? run.text : "";
            if (!text) {
              continue;
            }
            runs.push({
              bold: typeof run.bold === "boolean" ? run.bold : undefined,
              color: typeof run.color === "string" && run.color.trim() ? run.color : undefined,
              fontFamily: typeof run.fontFamily === "string" && run.fontFamily.trim() ? run.fontFamily : undefined,
              fontSizePt: asFiniteNumber(run.fontSizePt) ?? undefined,
              italic: typeof run.italic === "boolean" ? run.italic : undefined,
              text,
              underline: typeof run.underline === "boolean" ? run.underline : undefined
            });
          }
        }
        if (runs.length === 0) {
          continue;
        }
        const align = paragraph.align;
        normalizedParagraphs.push({
          align: align === "center" || align === "justify" || align === "left" || align === "right" ? align : undefined,
          runs
        });
    }
  }

  if (normalizedParagraphs.length > 0) {
    return normalizedParagraphs;
  }

  const text = typeof fallbackText === "string" ? fallbackText : "";
  return text
    ? [{ runs: [{ text }] }]
    : [];
}

function buildWorksheetDirectApiShape(
  workbookSheetIndex: number,
  info: WorksheetDirectShapeInfo,
  zIndex: number
): XlsxShape {
  const fill = info.fill && typeof info.fill === "object"
    ? {
        color: typeof info.fill.color === "string" && info.fill.color.trim() ? info.fill.color : undefined,
        none: typeof info.fill.none === "boolean" ? info.fill.none : undefined,
        opacity: asFiniteNumber(info.fill.opacity) ?? undefined
      }
    : undefined;
  const stroke = info.stroke && typeof info.stroke === "object"
    ? {
        color: typeof info.stroke.color === "string" && info.stroke.color.trim() ? info.stroke.color : undefined,
        dash: typeof info.stroke.dash === "string" && info.stroke.dash.trim() ? info.stroke.dash : undefined,
        headEndType: typeof info.stroke.headEndType === "string" && info.stroke.headEndType.trim() ? info.stroke.headEndType : undefined,
        none: typeof info.stroke.none === "boolean" ? info.stroke.none : undefined,
        opacity: asFiniteNumber(info.stroke.opacity) ?? undefined,
        tailEndType: typeof info.stroke.tailEndType === "string" && info.stroke.tailEndType.trim() ? info.stroke.tailEndType : undefined,
        widthPx: asFiniteNumber(info.stroke.widthPx) ?? undefined
      }
    : undefined;
  const rawSvgViewBox = info.svgViewBox && typeof info.svgViewBox === "object" ? info.svgViewBox : null;
  const rawTextBox = info.textBox && typeof info.textBox === "object" ? info.textBox : null;
  const rawInset = rawTextBox?.insetPx && typeof rawTextBox.insetPx === "object" ? rawTextBox.insetPx : null;

  return {
    anchor: buildWorksheetDirectImageAnchor(
      info.anchor,
      DEFAULT_COL_WIDTH * EMU_PER_PIXEL,
      DEFAULT_ROW_HEIGHT * EMU_PER_PIXEL
    ),
    description: typeof info.description === "string" && info.description.trim() ? info.description : undefined,
    fill,
    flipH: typeof info.flipH === "boolean" ? info.flipH : undefined,
    flipV: typeof info.flipV === "boolean" ? info.flipV : undefined,
    geometry: typeof info.geometry === "string" && info.geometry.trim() ? info.geometry : "rect",
    geometryAdjustments: info.geometryAdjustments && typeof info.geometryAdjustments === "object"
      ? Object.fromEntries(
          Object.entries(info.geometryAdjustments as Record<string, unknown>)
            .map(([key, value]) => [key, asFiniteNumber(value)])
            .filter((entry): entry is [string, number] => typeof entry[1] === "number")
        )
      : undefined,
    hyperlink: typeof info.hyperlink === "string" && info.hyperlink.trim() ? info.hyperlink : undefined,
    id: `shape-${workbookSheetIndex}-${String(info.id ?? zIndex)}`,
    name: typeof info.name === "string" && info.name.trim() ? info.name : undefined,
    paragraphs: normalizeWorksheetDirectShapeParagraphs(info.paragraphs, info.text),
    rotationDeg: asFiniteNumber(info.rotationDeg) ?? undefined,
    scaleX: asFiniteNumber(info.scaleX) ?? undefined,
    scaleY: asFiniteNumber(info.scaleY) ?? undefined,
    sheetIndex: workbookSheetIndex,
    svgPath: typeof info.svgPath === "string" && info.svgPath.trim() ? info.svgPath : undefined,
    svgViewBox: rawSvgViewBox
      && asFiniteNumber(rawSvgViewBox.width) !== null
      && asFiniteNumber(rawSvgViewBox.height) !== null
      ? {
          height: asFiniteNumber(rawSvgViewBox.height) ?? 0,
          width: asFiniteNumber(rawSvgViewBox.width) ?? 0
        }
      : undefined,
    stroke,
    textBox: rawTextBox
      ? {
          horizontalAlign: rawTextBox.horizontalAlign === "center" || rawTextBox.horizontalAlign === "left"
            ? rawTextBox.horizontalAlign
            : undefined,
          insetPx: rawInset
            ? {
                bottom: asFiniteNumber(rawInset.bottom) ?? 0,
                left: asFiniteNumber(rawInset.left) ?? 0,
                right: asFiniteNumber(rawInset.right) ?? 0,
                top: asFiniteNumber(rawInset.top) ?? 0
              }
            : undefined,
          verticalAlign: rawTextBox.verticalAlign === "bottom" || rawTextBox.verticalAlign === "middle" || rawTextBox.verticalAlign === "top"
            ? rawTextBox.verticalAlign
            : undefined
        }
      : undefined,
    workbookSheetIndex,
    zIndex
  };
}

function buildWorksheetApiImage(
  workbookSheetIndex: number,
  row: number,
  col: number,
  info: WorksheetApiImageInfo,
  zIndex: number
): XlsxImage | null {
  if (typeof info.source !== "string" || !info.source) {
    return null;
  }

  const width = Math.max(1, Math.round(asFiniteNumber(info.width) ?? DEFAULT_COL_WIDTH));
  const height = Math.max(1, Math.round(asFiniteNumber(info.height) ?? DEFAULT_ROW_HEIGHT));
  const description = typeof info.altText === "string" && info.altText.trim() ? info.altText : undefined;

  return {
    anchor: {
      from: {
        col,
        colOffsetEmu: 0,
        row,
        rowOffsetEmu: 0
      },
      kind: "one-cell",
      sizeEmu: {
        cx: width * EMU_PER_PIXEL,
        cy: height * EMU_PER_PIXEL
      }
    },
    description,
    editable: false,
    id: `worksheet-image-${workbookSheetIndex}-${row}-${col}-${zIndex}`,
    mimeType: inferImageMimeType(info.source),
    sheetIndex: workbookSheetIndex,
    src: info.source,
    workbookSheetIndex,
    zIndex
  };
}

function buildWorksheetDirectApiImage(
  workbookSheetIndex: number,
  info: WorksheetDirectImageInfo,
  zIndex: number,
  objectUrls: string[]
): XlsxImage | null {
  const mimeType = inferWorksheetDirectImageMimeType(info);
  const src = createWorksheetDirectImageSource(info.data, mimeType, objectUrls);
  if (!src) {
    return null;
  }

  const widthEmu = Math.max(EMU_PER_PIXEL, Math.round(asFiniteNumber(info.widthEmu) ?? DEFAULT_COL_WIDTH * EMU_PER_PIXEL));
  const heightEmu = Math.max(EMU_PER_PIXEL, Math.round(asFiniteNumber(info.heightEmu) ?? DEFAULT_ROW_HEIGHT * EMU_PER_PIXEL));
  return {
    anchor: buildWorksheetDirectImageAnchor(info.anchor, widthEmu, heightEmu),
    editable: false,
    id: `worksheet-image-${workbookSheetIndex}-${String(info.id ?? zIndex)}`,
    mediaPath: typeof info.mediaPath === "string" && info.mediaPath.trim() ? info.mediaPath : undefined,
    mimeType,
    name: typeof info.name === "string" && info.name.trim() ? info.name : undefined,
    sheetIndex: workbookSheetIndex,
    src,
    workbookSheetIndex,
    zIndex
  };
}

function collectWorksheetBatchImages(workbook: Workbook) {
  const imagesByWorkbookSheetIndex = Array.from({ length: workbook.sheetCount }, () => [] as XlsxImage[]);

  for (let workbookSheetIndex = 0; workbookSheetIndex < workbook.sheetCount; workbookSheetIndex += 1) {
    const worksheet = workbook.getSheet(workbookSheetIndex) as WorksheetWithRowsBatch;
    if (typeof worksheet.getRowsBatch !== "function") {
      continue;
    }

    const usedRange = worksheet.usedRange() as [number, number, number, number] | null;
    const maxRow = usedRange?.[2] ?? -1;
    if (maxRow < 0) {
      continue;
    }

    let zIndex = 1;
    let sheetFailed = false;
    for (let startRow = 0; startRow <= maxRow; startRow += IMAGE_BATCH_ROW_COUNT) {
      let rows: unknown;
      try {
        rows = worksheet.getRowsBatch(startRow, IMAGE_BATCH_ROW_COUNT, { includeImages: true });
      } catch {
        sheetFailed = true;
        break;
      }

      if (!Array.isArray(rows)) {
        continue;
      }

      for (const rowEntry of rows as WorksheetApiRow[]) {
        const row = typeof rowEntry.index === "number" ? rowEntry.index : null;
        if (row === null || !Array.isArray(rowEntry.cells)) {
          continue;
        }

        for (const cellEntry of rowEntry.cells as WorksheetApiRowCell[]) {
          const col = typeof cellEntry.col === "number" ? cellEntry.col : null;
          if (col === null || !cellEntry.image || typeof cellEntry.image !== "object") {
            continue;
          }

          const image = buildWorksheetApiImage(workbookSheetIndex, row, col, cellEntry.image, zIndex);
          if (!image) {
            continue;
          }

          imagesByWorkbookSheetIndex[workbookSheetIndex].push(image);
          zIndex += 1;
        }
      }
    }

    if (sheetFailed) {
      imagesByWorkbookSheetIndex[workbookSheetIndex] = [];
    }
  }

  return imagesByWorkbookSheetIndex;
}

function collectWorksheetApiImages(workbook: Workbook, objectUrls: string[]) {
  const directImagesByWorkbookSheetIndex = Array.from({ length: workbook.sheetCount }, () => [] as XlsxImage[]);
  let didUseDirectImages = false;

  for (let workbookSheetIndex = 0; workbookSheetIndex < workbook.sheetCount; workbookSheetIndex += 1) {
    const worksheet = workbook.getSheet(workbookSheetIndex) as ReturnType<Workbook["getSheet"]> & {
      images?: unknown;
    };
    const rawImages = Array.isArray(worksheet.images) ? worksheet.images as WorksheetDirectImageInfo[] : [];
    if (rawImages.length === 0) {
      continue;
    }

    const nextImages = rawImages
      .map((info, index) => buildWorksheetDirectApiImage(workbookSheetIndex, info, index + 1, objectUrls))
      .filter((image): image is XlsxImage => Boolean(image));
    if (nextImages.length > 0) {
      directImagesByWorkbookSheetIndex[workbookSheetIndex] = nextImages;
      didUseDirectImages = true;
    }
  }

  if (didUseDirectImages) {
    return directImagesByWorkbookSheetIndex;
  }

  return collectWorksheetBatchImages(workbook);
}

function collectWorksheetApiShapes(workbook: Workbook) {
  return Array.from({ length: workbook.sheetCount }, (_, workbookSheetIndex) => {
    const worksheet = workbook.getSheet(workbookSheetIndex) as ReturnType<Workbook["getSheet"]> & {
      shapes?: unknown;
    };
    const rawShapes = Array.isArray(worksheet.shapes) ? worksheet.shapes as WorksheetDirectShapeInfo[] : [];
    return rawShapes
      .map((shape, index) => buildWorksheetDirectApiShape(workbookSheetIndex, shape, index + 1));
  });
}

function mergeParsedAndApiImages(parsedImages: XlsxImage[], apiImages: XlsxImage[]) {
  if (parsedImages.length === 0) {
    return apiImages;
  }
  if (apiImages.length === 0) {
    return parsedImages;
  }

  const normalizeTextKey = (value: string | undefined) => value?.trim().toLowerCase() ?? "";
  const anchorKey = (anchor: XlsxImage["anchor"]) => {
    if (anchor.kind === "absolute") {
      return [
        "absolute",
        Math.round(anchor.positionEmu.x),
        Math.round(anchor.positionEmu.y),
        Math.round(anchor.sizeEmu.cx),
        Math.round(anchor.sizeEmu.cy)
      ].join(":");
    }
    if (anchor.kind === "one-cell") {
      return [
        "one",
        anchor.from.col,
        anchor.from.row,
        Math.round(anchor.from.colOffsetEmu),
        Math.round(anchor.from.rowOffsetEmu),
        Math.round(anchor.sizeEmu.cx),
        Math.round(anchor.sizeEmu.cy)
      ].join(":");
    }
    return [
      "two",
      anchor.from.col,
      anchor.from.row,
      Math.round(anchor.from.colOffsetEmu),
      Math.round(anchor.from.rowOffsetEmu),
      anchor.to.col,
      anchor.to.row,
      Math.round(anchor.to.colOffsetEmu),
      Math.round(anchor.to.rowOffsetEmu)
    ].join(":");
  };
  const imageKeys = (image: XlsxImage) => {
    const keys = [
      `${normalizeTextKey(image.mediaPath)}|${normalizeTextKey(image.name)}|${anchorKey(image.anchor)}`,
      `${normalizeTextKey(image.mediaPath)}|${anchorKey(image.anchor)}`,
      `${normalizeTextKey(image.name)}|${anchorKey(image.anchor)}`,
      `${anchorKey(image.anchor)}`
    ];
    return keys.filter((key, index) => key && keys.indexOf(key) === index);
  };

  const apiBuckets = new Map<string, XlsxImage[]>();
  for (const apiImage of apiImages) {
    for (const key of imageKeys(apiImage)) {
      const bucket = apiBuckets.get(key);
      if (bucket) {
        bucket.push(apiImage);
      } else {
        apiBuckets.set(key, [apiImage]);
      }
    }
  }

  const usedApiImages = new Set<XlsxImage>();
  const takeApiMatch = (image: XlsxImage) => {
    for (const key of imageKeys(image)) {
      const bucket = apiBuckets.get(key);
      if (!bucket) {
        continue;
      }
      const match = bucket.find((candidate) => !usedApiImages.has(candidate));
      if (match) {
        usedApiImages.add(match);
        return match;
      }
    }
    return null;
  };

  const merged = parsedImages.map((image) => {
    const apiImage = takeApiMatch(image);
    if (!apiImage) {
      return image;
    }

    return {
      ...image,
      anchor: apiImage.anchor,
      mediaPath: apiImage.mediaPath ?? image.mediaPath,
      mimeType: apiImage.mimeType,
      name: apiImage.name ?? image.name,
      src: apiImage.src
    };
  });

  return merged;
}

function isZipWorkbook(bytes: Uint8Array) {
  return bytes.byteLength >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b;
}

function isLegacyXlsWorkbook(bytes: Uint8Array) {
  return bytes.byteLength >= 8
    && bytes[0] === 0xd0
    && bytes[1] === 0xcf
    && bytes[2] === 0x11
    && bytes[3] === 0xe0
    && bytes[4] === 0xa1
    && bytes[5] === 0xb1
    && bytes[6] === 0x1a
    && bytes[7] === 0xe1;
}

function shouldSkipXmlParsingForWorkbook(bytes: Uint8Array, skipXmlParsing = false) {
  return skipXmlParsing || isLegacyXlsWorkbook(bytes);
}

function createBasicWorkbookAssets(workbook: Workbook): WorkbookImageAssets {
  const objectUrls: string[] = [];
  return {
    archive: {},
    formControlsByWorkbookSheetIndex: Array.from({ length: workbook.sheetCount }, () => [] as XlsxFormControl[]),
    imageOriginsById: new Map(),
    imagesByWorkbookSheetIndex: collectWorksheetApiImages(workbook, objectUrls),
    namedCellStyleByName: {},
    objectUrls,
    shapesByWorkbookSheetIndex: collectWorksheetApiShapes(workbook),
    sheetOrigins: Array.from({ length: workbook.sheetCount }, () => null as WorkbookImageSheetOrigin | null),
    sheetStatesByWorkbookSheetIndex: Array.from({ length: workbook.sheetCount }, () => null),
    styleById: {},
    tableMetadataByWorkbookSheetIndex: Array.from({ length: workbook.sheetCount }, () => []),
    tableStyleByName: {},
    themePalette: { colorsByIndex: {} }
  };
}

function loadWorkbookImageAssets(bytes: Uint8Array, workbook: Workbook, skipXmlParsing = false) {
  if (shouldSkipXmlParsingForWorkbook(bytes, skipXmlParsing) || !isZipWorkbook(bytes)) {
    return createBasicWorkbookAssets(workbook);
  }

  const parsedAssets = parseWorkbookImageAssets(bytes);
  const apiImagesByWorkbookSheetIndex = collectWorksheetApiImages(workbook, parsedAssets.objectUrls);

  const imagesByWorkbookSheetIndex = Array.from(
    { length: Math.max(workbook.sheetCount, parsedAssets.imagesByWorkbookSheetIndex.length, apiImagesByWorkbookSheetIndex.length) },
    (_, index) => {
      const parsedImages = parsedAssets.imagesByWorkbookSheetIndex[index] ?? [];
      const apiImages = apiImagesByWorkbookSheetIndex[index] ?? [];
      return mergeParsedAndApiImages(parsedImages, apiImages);
    }
  );

  return {
    ...parsedAssets,
    imagesByWorkbookSheetIndex
  };
}

function downloadArrayBuffer(file: ArrayBuffer, fileName: string) {
  const blob = new Blob([file], { type: XLSX_MIME_TYPE });
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = objectUrl;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
}

function downloadBytes(bytes: Uint8Array, fileName: string, mimeType: string) {
  const normalizedBytes = new Uint8Array(bytes.byteLength);
  normalizedBytes.set(bytes);
  const blob = new Blob([normalizedBytes], { type: mimeType });
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = objectUrl;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
}

function downloadText(text: string, fileName: string, mimeType: string) {
  const blob = new Blob([text], { type: mimeType });
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = objectUrl;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
}

function downloadUrl(src: string, fileName: string) {
  const anchor = document.createElement("a");
  anchor.href = src;
  anchor.download = fileName;
  anchor.rel = "noreferrer";
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
}

export function useXlsxViewerController(options: UseXlsxViewerControllerOptions): XlsxViewerController {
  const {
    allowResizeInReadOnly = false,
    deferLoadingAboveBytes = DEFAULT_DEFER_LOADING_ABOVE_BYTES,
    file,
    fileName,
    maxFileSizeBytes = DEFAULT_MAX_FILE_SIZE_BYTES,
    readOnly: requestedReadOnly = false,
    readOnlyAboveBytes = 0,
    showHiddenSheets = false,
    skipXmlParsing = false,
    src,
    useWorker = true
  } = options;
  const isLoading = ref(Boolean(file ?? src));
  const error = ref<Error | null>(null);
  const workbook = ref<Workbook | null>(null);
  const sheets = ref<XlsxSheetData[]>([]);
  const chartsByWorkbookSheetIndex = ref<XlsxChart[][]>([]);
  const chartsheets = ref<XlsxChartsheet[]>([]);
  const tabs = ref<XlsxWorkbookTab[]>([]);
  const isChartsLoading = ref(false);
  const workerTablesByWorkbookSheetIndex = ref<XlsxTable[][]>([]);
  const formControlsByWorkbookSheetIndex = ref<XlsxFormControl[][]>([]);
  const imagesByWorkbookSheetIndex = ref<XlsxImage[][]>([]);
  const shapesByWorkbookSheetIndex = ref<XlsxShape[][]>([]);
  const activeSheetIndex = ref(0);
  const activeTabIndex = ref(0);
  const zoomScaleOverridesByTabId = ref<Record<string, number>>({});
  const activeCell = ref<XlsxCellAddress | null>(null);
  const selection = ref<XlsxCellRange | null>(null);
  const selectedChartId = ref<string | null>(null);
  const selectedChartElement = ref<XlsxChartElementSelection | null>(null);
  const selectedImageId = ref<string | null>(null);
  const revision = ref(0);
  const selectionAnchorRef = shallowRef<XlsxCellAddress | null>(null)
  const undoStackRef = shallowRef<HistoryEntry[]>([])
  const redoStackRef = shallowRef<HistoryEntry[]>([])
  const isApplyingHistoryRef = shallowRef(false)
  const historyRevision = ref(0)
  const shouldAutoCalculate = ref(false)
  const workerCellSnapshotRevision = ref(0)
  const isWorkerBacked = ref(false)
  const sortState = ref<XlsxTableSortState | null>(null)
  const forcedReadOnly = ref(false)
  const deferredBufferRef = shallowRef<ArrayBuffer | null>(null)
  const deferredLoadFileSize = ref<number | null>(null)
  const imageAssetsRef = shallowRef<WorkbookImageAssets | null>(null)
  const chartAssetsRef = shallowRef<WorkbookChartAssets | null>(null)
  const chartLoadRequestTokenRef = shallowRef(0)
  const chartDisplayFallbackCleanupRef = shallowRef<(() => void) | null>(null)
  const sheetOriginsRef = shallowRef<Array<WorkbookImageSheetOrigin | null>>([])
  const workerClientRef = shallowRef<XlsxWorkerClient | null>(null)
  const workerCellSnapshotCacheRef = shallowRef(new Map<string, { displayValue: string; formula: string }>())
  const displayFileName = computed(() => { return resolveDisplayFileName(src, fileName); })
  const shouldDeferLoading = deferLoadingAboveBytes > 0;
  const readOnly = computed(() => requestedReadOnly || forcedReadOnly.value);
  const canResizeReadOnly = computed(() => requestedReadOnly && allowResizeInReadOnly && !forcedReadOnly.value);
  const workerSupported = useWorker && typeof Worker !== "undefined" && canUseConfiguredWasmSourceInWorker();
  const canUseWorkerForRequestedReadOnly = requestedReadOnly;
  function shouldForceReadOnlyForBuffer(bufferByteLength: number) {
  return (
    !requestedReadOnly && readOnlyAboveBytes > 0 && bufferByteLength > readOnlyAboveBytes
  );
}
  function shouldUseWorkerForReadOnlyLoad(willForceReadOnly: boolean) {
  return (
    workerSupported && (willForceReadOnly || canUseWorkerForRequestedReadOnly)
  );
}

  function disposeWorkerClient() {
    workerClientRef.value?.dispose();
    workerClientRef.value = null;
  }

  function getWorkerClient() {
    if (!workerClientRef.value) {
      workerClientRef.value = new XlsxWorkerClient();
    }

    return workerClientRef.value;
  }

  function clearImageAssets() {
    revokeWorkbookImageAssets(imageAssetsRef.value);
    imageAssetsRef.value = null;
    sheetOriginsRef.value = [];
    formControlsByWorkbookSheetIndex.value = [];
    imagesByWorkbookSheetIndex.value = [];
    shapesByWorkbookSheetIndex.value = [];
  }

  function clearChartAssets() {
    chartLoadRequestTokenRef.value += 1;
    chartDisplayFallbackCleanupRef.value?.();
    chartDisplayFallbackCleanupRef.value = null;
    chartAssetsRef.value = null;
    chartsByWorkbookSheetIndex.value = [];
    chartsheets.value = [];
    tabs.value = [];
    isChartsLoading.value = false;
  }

  function setImageAssets(assets: WorkbookImageAssets | null) {
    revokeWorkbookImageAssets(imageAssetsRef.value);
    imageAssetsRef.value = assets;
    sheetOriginsRef.value = assets?.sheetOrigins.slice() ?? [];
    formControlsByWorkbookSheetIndex.value = assets?.formControlsByWorkbookSheetIndex ?? [];
    imagesByWorkbookSheetIndex.value = assets?.imagesByWorkbookSheetIndex ?? [];
    shapesByWorkbookSheetIndex.value = assets?.shapesByWorkbookSheetIndex ?? [];
  }

  function setChartAssets(assets: WorkbookChartAssets | null) {
    chartAssetsRef.value = assets;
    chartsByWorkbookSheetIndex.value = assets?.chartsByWorkbookSheetIndex ?? [];
    chartsheets.value = assets?.chartsheets ?? [];
    tabs.value = assets?.tabs ?? [];
    isChartsLoading.value = false;
  }

  function shouldFallbackFromWorkerError(error: unknown) {
    const message = error instanceof Error ? error.message : String(error ?? "");
    return (
      message.includes("DOMParser is not defined")
      || message.includes("XMLSerializer is not defined")
      || message.includes("Worker chart payload incomplete")
    );
  }

  function hasIncompleteWorkerChartSnapshot(snapshot: {
    chartsByWorkbookSheetIndex: XlsxChart[][];
  }) {
    for (const sheetCharts of snapshot.chartsByWorkbookSheetIndex) {
      for (const chart of sheetCharts) {
        if (chart.chartType !== "Bubble") {
          continue;
        }
        for (const series of chart.series) {
          const pointCount = Math.max(series.values.length, series.categories.length);
          if (pointCount <= 1) {
            continue;
          }
          const numericBubbleSizes = (series.bubbleSizes ?? []).filter(
            (value): value is number => typeof value === "number" && Number.isFinite(value)
          );
          if (numericBubbleSizes.length < pointCount) {
            return true;
          }
        }
      }
    }

    return false;
  }

  function ensureChartAssetsHydrated(targetWorkbook: Workbook | null, targetSheets: XlsxSheetData[]) {
    const currentAssets = chartAssetsRef.value;
    if (
      currentAssets
      && (
        currentAssets.chartOriginsById.size > 0
        || !targetWorkbook
        || !imageAssetsRef.value
      )
    ) {
      return currentAssets;
    }

    if (!targetWorkbook || !imageAssetsRef.value) {
      return chartAssetsRef.value;
    }

    const assets = loadWorkbookChartAssets(
      targetWorkbook,
      imageAssetsRef.value,
      buildVisibleSheetIndexMap(targetSheets),
      showHiddenSheets
    );
    chartAssetsRef.value = assets;
    return assets;
  }

  function startChartDisplayHydration(buffer: ArrayBuffer, targetWorkbook: Workbook, targetSheets: XlsxSheetData[]) {
    const effectiveSkipXmlParsing = shouldSkipXmlParsingForWorkbook(new Uint8Array(buffer), skipXmlParsing);
    const visibleSheetIndexByWorkbookSheetIndex = buildVisibleSheetIndexMap(targetSheets);
    const quickAssets = loadWorkbookChartAssets(targetWorkbook, null, visibleSheetIndexByWorkbookSheetIndex, showHiddenSheets);
    setChartAssets(quickAssets);

    if (effectiveSkipXmlParsing) {
      return;
    }

    const hasCharts = quickAssets.chartsByWorkbookSheetIndex.some((sheetCharts) => sheetCharts.length > 0);
    if (!hasCharts) {
      isChartsLoading.value = false;
      return;
    }

    isChartsLoading.value = true;
    const requestToken = chartLoadRequestTokenRef.value + 1;
    chartLoadRequestTokenRef.value = requestToken;
    chartDisplayFallbackCleanupRef.value?.();
    chartDisplayFallbackCleanupRef.value = null;
    let fallbackTriggered = false;
    const triggerFallback = () => {
      if (fallbackTriggered || requestToken !== chartLoadRequestTokenRef.value) {
        return;
      }
      fallbackTriggered = true;
      runMainThreadFallback();
    };
    const workerTimeoutHandle = typeof window !== "undefined"
      ? window.setTimeout(() => {
          triggerFallback();
        }, 1500)
      : null;

    const applyWorkerResult = (result: {
      chartsByWorkbookSheetIndex: XlsxChart[][];
      chartsheets: XlsxChartsheet[];
      tabs: XlsxWorkbookTab[];
    }) => {
      if (requestToken !== chartLoadRequestTokenRef.value) {
        return;
      }
      chartsByWorkbookSheetIndex.value = result.chartsByWorkbookSheetIndex;
      chartsheets.value = result.chartsheets;
      tabs.value = result.tabs;
      isChartsLoading.value = false;
    };

    const runMainThreadFallback = () => {
      chartDisplayFallbackCleanupRef.value = scheduleLowPriorityTask(() => {
        if (requestToken !== chartLoadRequestTokenRef.value) {
          return;
        }
        try {
          const hydratedAssets = loadWorkbookChartAssets(
            targetWorkbook,
            imageAssetsRef.value,
            visibleSheetIndexByWorkbookSheetIndex,
            showHiddenSheets
          );
          if (requestToken !== chartLoadRequestTokenRef.value) {
            return;
          }
          setChartAssets(hydratedAssets);
        } catch {
          if (requestToken !== chartLoadRequestTokenRef.value) {
            return;
          }
          setChartAssets(quickAssets);
        } finally {
          if (requestToken === chartLoadRequestTokenRef.value) {
            isChartsLoading.value = false;
          }
        }
      });
    };

    if (!workerSupported) {
      runMainThreadFallback();
      return;
    }

    void getWorkerClient().parseCharts(buffer, effectiveSkipXmlParsing, showHiddenSheets)
      .then((result) => {
        if (workerTimeoutHandle !== null) {
          window.clearTimeout(workerTimeoutHandle);
        }
        if (fallbackTriggered) {
          return;
        }
        try {
          if (hasIncompleteWorkerChartSnapshot(result)) {
            triggerFallback();
            return;
          }
          applyWorkerResult(result);
        } catch {
          triggerFallback();
        }
      })
      .catch((error: unknown) => {
        if (workerTimeoutHandle !== null) {
          window.clearTimeout(workerTimeoutHandle);
        }
        if (isAbortError(error)) {
          return;
        }
        triggerFallback();
      });
  }

  async function loadWorkbookOnMainThread(buffer: ArrayBuffer) {
    const nextParsedWorkbook = await parseWorkbookBuffer(buffer);
    const bytes = new Uint8Array(buffer);
    const nextImageAssets = loadWorkbookImageAssets(
      bytes,
      nextParsedWorkbook.workbook,
      shouldSkipXmlParsingForWorkbook(bytes, skipXmlParsing)
    );
    return {
      imageAssets: nextImageAssets,
      parsedWorkbook: nextParsedWorkbook
    };
  }

  function refreshWorkbookState(targetWorkbook: Workbook) {
    const nextSheets = buildSheetList(
      targetWorkbook,
      imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex,
      imageAssetsRef.value?.themePalette,
      imageAssetsRef.value?.styleById,
      imageAssetsRef.value?.namedCellStyleByName,
      imageAssetsRef.value?.tableStyleByName,
      showHiddenSheets
    );
    sheets.value = nextSheets;
    setChartAssets(
      loadWorkbookChartAssets(
        targetWorkbook,
        imageAssetsRef.value,
        buildVisibleSheetIndexMap(nextSheets),
        showHiddenSheets
      )
    );
    revision.value = revision.value + 1;
  }

  onUnmounted(() => {
    chartDisplayFallbackCleanupRef.value?.();
    chartDisplayFallbackCleanupRef.value = null;
    revokeWorkbookImageAssets(imageAssetsRef.value);
    disposeWorkerClient();
  });

  watch(() => [file, src], (_value, _oldValue, onCleanup) => {
    if (!file && !src) {
      disposeWorkerClient();
      forcedReadOnly.value = false;
      workbook.value = null;
      sheets.value = [];
      clearChartAssets();
      workerTablesByWorkbookSheetIndex.value = [];
      clearImageAssets();
      error.value = null;
      isLoading.value = false;
      isWorkerBacked.value = false;
      deferredBufferRef.value = null;
      deferredLoadFileSize.value = null;
      activeSheetIndex.value = 0;
      activeTabIndex.value = 0;
      activeCell.value = null;
      selection.value = null;
      selectedChartId.value = null;
      selectedChartElement.value = null;
      selectedImageId.value = null;
      selectionAnchorRef.value = null;
      undoStackRef.value = [];
      redoStackRef.value = [];
      historyRevision.value = 0;
      shouldAutoCalculate.value = false;
      workerCellSnapshotCacheRef.value.clear();
      workerCellSnapshotRevision.value = 0;
      sortState.value = null;
      zoomScaleOverridesByTabId.value = {};
      revision.value = 0;
      return;
    }

    let isCurrent = true;
    const abortController = new AbortController();
    isLoading.value = true;
    error.value = null;
    clearImageAssets();
    clearChartAssets();
    workerTablesByWorkbookSheetIndex.value = [];
    isWorkerBacked.value = false;
    deferredBufferRef.value = null;
    deferredLoadFileSize.value = null;
    activeSheetIndex.value = 0;
    activeTabIndex.value = 0;
    activeCell.value = null;
    selection.value = null;
    selectedChartId.value = null;
    selectedChartElement.value = null;
    selectedImageId.value = null;
    selectionAnchorRef.value = null;
    undoStackRef.value = [];
    redoStackRef.value = [];
    historyRevision.value = 0;
    shouldAutoCalculate.value = false;
    workerCellSnapshotCacheRef.value.clear();
    workerCellSnapshotRevision.value = 0;
    sortState.value = null;
    zoomScaleOverridesByTabId.value = {};
    revision.value = 0;
    disposeWorkerClient();

    void resolveWorkbookBuffer({ file, src }, abortController.signal)
      .then(async (buffer) => {
        if (!isCurrent || abortController.signal.aborted) {
          return;
        }

        if (maxFileSizeBytes > 0 && buffer.byteLength > maxFileSizeBytes) {
          throw new XlsxFileSizeLimitExceededError(buffer.byteLength, maxFileSizeBytes);
        }

        const preflight = preflightWorkbookBuffer(buffer);
        if (preflight?.tooLarge) {
          throw createWorkbookTooLargeError(preflight);
        }

        const shouldForceReadOnly = shouldForceReadOnlyForBuffer(buffer.byteLength);
        forcedReadOnly.value = shouldForceReadOnly;
        const shouldUseWorkerForLoad = shouldUseWorkerForReadOnlyLoad(shouldForceReadOnly);
        const effectiveSkipXmlParsing = shouldSkipXmlParsingForWorkbook(new Uint8Array(buffer), skipXmlParsing);

        if (shouldDeferLoading && buffer.byteLength > deferLoadingAboveBytes) {
          deferredBufferRef.value = buffer;
          deferredLoadFileSize.value = buffer.byteLength;
          workbook.value = null;
          sheets.value = [];
          clearChartAssets();
          workerTablesByWorkbookSheetIndex.value = [];
          isLoading.value = false;
          return;
        }

        if (shouldUseWorkerForLoad) {
          try {
            const snapshot = await getWorkerClient().loadWorkbook(buffer, effectiveSkipXmlParsing, showHiddenSheets);
            if (!isCurrent || abortController.signal.aborted) {
              return;
            }
            if (!effectiveSkipXmlParsing && hasIncompleteWorkerChartSnapshot(snapshot)) {
              throw new Error("Worker chart payload incomplete");
            }

            workbook.value = null;
            sheets.value = snapshot.sheets;
            chartsByWorkbookSheetIndex.value = snapshot.chartsByWorkbookSheetIndex;
            chartsheets.value = snapshot.chartsheets;
            tabs.value = snapshot.tabs;
            chartAssetsRef.value = null;
            workerTablesByWorkbookSheetIndex.value = snapshot.tablesByWorkbookSheetIndex;
            shouldAutoCalculate.value = false;
            isWorkerBacked.value = true;
            sortState.value = null;
            isChartsLoading.value = false;
            isLoading.value = false;
            return;
          } catch (workerError) {
            if (!isCurrent || isAbortError(workerError)) {
              return;
            }
            if (!shouldFallbackFromWorkerError(workerError)) {
              throw workerError;
            }

            disposeWorkerClient();
          }
        }

        const { imageAssets: nextImageAssets, parsedWorkbook: nextParsedWorkbook } = await loadWorkbookOnMainThread(buffer);
        if (!isCurrent || abortController.signal.aborted) {
          revokeWorkbookImageAssets(nextImageAssets);
          return;
        }

        setImageAssets(nextImageAssets);
        workbook.value = nextParsedWorkbook.workbook;
        const nextSheets = buildSheetList(
          nextParsedWorkbook.workbook,
          nextImageAssets.sheetStatesByWorkbookSheetIndex,
          nextImageAssets.themePalette,
          nextImageAssets.styleById,
          nextImageAssets.namedCellStyleByName,
          nextImageAssets.tableStyleByName,
          showHiddenSheets
        );
        sheets.value = nextSheets;
        startChartDisplayHydration(buffer, nextParsedWorkbook.workbook, nextSheets);
        shouldAutoCalculate.value = nextParsedWorkbook.shouldAutoCalculate;
        workerTablesByWorkbookSheetIndex.value = [];
        isWorkerBacked.value = false;
        sortState.value = null;
        isLoading.value = false;
      })
      .catch((nextError: unknown) => {
        if (!isCurrent || isAbortError(nextError)) {
          return;
        }

        workbook.value = null;
        sheets.value = [];
        clearChartAssets();
        workerTablesByWorkbookSheetIndex.value = [];
        clearImageAssets();
        shouldAutoCalculate.value = false;
        isWorkerBacked.value = false;
        sortState.value = null;
        error.value = nextError instanceof Error ? nextError : new Error("Could not load workbook.");
        isLoading.value = false;
      });

    onCleanup(() => {
      isCurrent = false;
      abortController.abort();
      disposeWorkerClient();
    });
  }, { immediate: true });

  const activeTab = computed(() => tabs.value[activeTabIndex.value] ?? null);
  const activeSheet = computed(() => {
    const tab = activeTab.value;
    return tab?.kind === "sheet"
      ? sheets.value[tab.sheetIndex ?? -1] ?? null
      : null;
  });
  const deferredMetadataCell = computed(() => activeCell.value);
  const deferredMetadataSheet = computed(() => activeSheet.value);
  const activeZoomTabKey = computed(() => activeTab.value?.id ?? DEFAULT_ZOOM_TAB_KEY);
  const defaultZoomScale = computed(() => resolveDefaultZoomScale(activeTab.value, activeSheet.value));
  const zoomScale = computed(() => clampZoomScale(zoomScaleOverridesByTabId.value[activeZoomTabKey.value] ?? defaultZoomScale.value));
  const canZoomIn = computed(() => zoomScale.value < MAX_ZOOM_SCALE);
  const canZoomOut = computed(() => zoomScale.value > MIN_ZOOM_SCALE);

  watch(() => activeTabIndex.value, () => {
    activeCell.value = null;
    selection.value = null;
    selectedChartId.value = null;
    selectedChartElement.value = null;
    selectedImageId.value = null;
    selectionAnchorRef.value = null;
    sortState.value = null;
  });

  function setActiveSheetIndex(index: number) {
    activeSheetIndex.value = (() => {
      if (index < 0 || index >= sheets.value.length) {
        return activeSheetIndex.value;
      }
      const targetSheet = sheets.value[index];
      const tabIndex = tabs.value.findIndex((tab) => tab.kind === "sheet" && tab.workbookSheetIndex === targetSheet?.workbookSheetIndex);
      if (tabIndex >= 0) {
        activeTabIndex.value = tabIndex;
      }
      return index;
    })();
  }

  function setActiveTabIndex(index: number) {
    activeTabIndex.value = (() => {
      if (index < 0 || index >= tabs.value.length) {
        return activeTabIndex.value;
      }

      const targetTab = tabs.value[index];
      if (targetTab?.kind === "sheet" && typeof targetTab.sheetIndex === "number") {
        activeSheetIndex.value = targetTab.sheetIndex;
      }
      return index;
    })();
  }

  function setZoomScale(nextZoomScale: number) {
    const normalizedZoomScale = clampZoomScale(nextZoomScale);
    zoomScaleOverridesByTabId.value = (() => {
      if (zoomScaleOverridesByTabId.value[activeZoomTabKey.value] === normalizedZoomScale) {
        return zoomScaleOverridesByTabId.value;
      }

      return {
        ...zoomScaleOverridesByTabId.value,
        [activeZoomTabKey.value]: normalizedZoomScale
      };
    })();
  }

  function resetZoom() {
    zoomScaleOverridesByTabId.value = (() => {
      if (zoomScaleOverridesByTabId.value[activeZoomTabKey.value] === undefined) {
        return zoomScaleOverridesByTabId.value;
      }

      const next = { ...zoomScaleOverridesByTabId.value };
      delete next[activeZoomTabKey.value];
      return next;
    })();
  }

  function zoomIn() {
    setZoomScale(resolveNextZoomScale(zoomScale.value, 1));
  }

  function zoomOut() {
    setZoomScale(resolveNextZoomScale(zoomScale.value, -1));
  }

  watch(() => tabs.value.length, () => {
    activeTabIndex.value = (() => {
      if (tabs.value.length === 0) {
        return 0;
      }
      return Math.min(activeTabIndex.value, tabs.value.length - 1);
    })();
  });

  function continueDeferredLoad() {
    const deferredBuffer = deferredBufferRef.value;
    if (!deferredBuffer) {
      return;
    }

    isLoading.value = true;
    error.value = null;

    if (maxFileSizeBytes > 0 && deferredBuffer.byteLength > maxFileSizeBytes) {
      deferredBufferRef.value = null;
      deferredLoadFileSize.value = null;
      workbook.value = null;
      sheets.value = [];
      clearChartAssets();
      workerTablesByWorkbookSheetIndex.value = [];
      clearImageAssets();
      shouldAutoCalculate.value = false;
      isWorkerBacked.value = false;
      sortState.value = null;
      error.value = new XlsxFileSizeLimitExceededError(deferredBuffer.byteLength, maxFileSizeBytes);
      isLoading.value = false;
      return;
    }

    const preflight = preflightWorkbookBuffer(deferredBuffer);
    if (preflight?.tooLarge) {
      deferredBufferRef.value = null;
      deferredLoadFileSize.value = null;
      workbook.value = null;
      sheets.value = [];
      clearChartAssets();
      workerTablesByWorkbookSheetIndex.value = [];
      clearImageAssets();
      shouldAutoCalculate.value = false;
      isWorkerBacked.value = false;
      sortState.value = null;
      error.value = createWorkbookTooLargeError(preflight);
      isLoading.value = false;
      return;
    }

    const shouldForceReadOnly = shouldForceReadOnlyForBuffer(deferredBuffer.byteLength);
    forcedReadOnly.value = shouldForceReadOnly;
    const shouldUseWorkerForLoad = shouldUseWorkerForReadOnlyLoad(shouldForceReadOnly);
    const effectiveSkipXmlParsing = shouldSkipXmlParsingForWorkbook(new Uint8Array(deferredBuffer), skipXmlParsing);

    if (shouldUseWorkerForLoad) {
      void getWorkerClient().loadWorkbook(deferredBuffer, effectiveSkipXmlParsing, showHiddenSheets)
        .then((snapshot) => {
          if (!effectiveSkipXmlParsing && hasIncompleteWorkerChartSnapshot(snapshot)) {
            throw new Error("Worker chart payload incomplete");
          }
          deferredBufferRef.value = null;
          deferredLoadFileSize.value = null;
          workbook.value = null;
          sheets.value = snapshot.sheets;
          chartsByWorkbookSheetIndex.value = snapshot.chartsByWorkbookSheetIndex;
          chartsheets.value = snapshot.chartsheets;
          tabs.value = snapshot.tabs;
          chartAssetsRef.value = null;
          workerTablesByWorkbookSheetIndex.value = snapshot.tablesByWorkbookSheetIndex;
          shouldAutoCalculate.value = false;
          isWorkerBacked.value = true;
          sortState.value = null;
          isChartsLoading.value = false;
          isLoading.value = false;
        })
        .catch(async (workerError: unknown) => {
          if (isAbortError(workerError)) {
            return;
          }
          if (!shouldFallbackFromWorkerError(workerError)) {
            throw workerError;
          }

          disposeWorkerClient();
          const { imageAssets: nextImageAssets, parsedWorkbook: nextParsedWorkbook } = await loadWorkbookOnMainThread(deferredBuffer);
          deferredBufferRef.value = null;
          deferredLoadFileSize.value = null;
          setImageAssets(nextImageAssets);
          workbook.value = nextParsedWorkbook.workbook;
          const nextSheets = buildSheetList(
            nextParsedWorkbook.workbook,
            nextImageAssets.sheetStatesByWorkbookSheetIndex,
            nextImageAssets.themePalette,
            nextImageAssets.styleById,
            nextImageAssets.namedCellStyleByName,
            nextImageAssets.tableStyleByName,
            showHiddenSheets
          );
          sheets.value = nextSheets;
          startChartDisplayHydration(deferredBuffer, nextParsedWorkbook.workbook, nextSheets);
          shouldAutoCalculate.value = nextParsedWorkbook.shouldAutoCalculate;
          workerTablesByWorkbookSheetIndex.value = [];
          isWorkerBacked.value = false;
          sortState.value = null;
          isLoading.value = false;
        })
        .catch((nextError: unknown) => {
          deferredBufferRef.value = null;
          deferredLoadFileSize.value = null;
          workbook.value = null;
          sheets.value = [];
          clearChartAssets();
          workerTablesByWorkbookSheetIndex.value = [];
          clearImageAssets();
          shouldAutoCalculate.value = false;
          isWorkerBacked.value = false;
          sortState.value = null;
          error.value = nextError instanceof Error ? nextError : new Error("Could not load workbook.");
          isLoading.value = false;
        });
      return;
    }

    void parseWorkbookBuffer(deferredBuffer)
      .then((nextParsedWorkbook) => {
        const bytes = new Uint8Array(deferredBuffer);
        const nextImageAssets = loadWorkbookImageAssets(
          bytes,
          nextParsedWorkbook.workbook,
          shouldSkipXmlParsingForWorkbook(bytes, skipXmlParsing)
        );
        deferredBufferRef.value = null;
        deferredLoadFileSize.value = null;
        setImageAssets(nextImageAssets);
        workbook.value = nextParsedWorkbook.workbook;
        const nextSheets = buildSheetList(
          nextParsedWorkbook.workbook,
          nextImageAssets.sheetStatesByWorkbookSheetIndex,
          nextImageAssets.themePalette,
          nextImageAssets.styleById,
          nextImageAssets.namedCellStyleByName,
          nextImageAssets.tableStyleByName,
          showHiddenSheets
        );
        sheets.value = nextSheets;
        startChartDisplayHydration(deferredBuffer, nextParsedWorkbook.workbook, nextSheets);
        shouldAutoCalculate.value = nextParsedWorkbook.shouldAutoCalculate;
        workerTablesByWorkbookSheetIndex.value = [];
        isWorkerBacked.value = false;
        sortState.value = null;
        isLoading.value = false;
      })
      .catch((nextError: unknown) => {
        deferredBufferRef.value = null;
        deferredLoadFileSize.value = null;
        workbook.value = null;
        sheets.value = [];
        clearChartAssets();
        workerTablesByWorkbookSheetIndex.value = [];
        clearImageAssets();
        shouldAutoCalculate.value = false;
        isWorkerBacked.value = false;
        sortState.value = null;
        error.value = nextError instanceof Error ? nextError : new Error("Could not load workbook.");
        isLoading.value = false;
      });
  }

  function maybeRecalculateWorkbook(targetWorkbook: Workbook) {
    if (!shouldAutoCalculate.value) {
      return;
    }

    const result = tryRecalculate(targetWorkbook);
    if (!result.calculated) {
      shouldAutoCalculate.value = false;
    }
  }

  function getActiveWorksheet() {
    if (!workbook.value || !activeSheet.value) {
      return null;
    }

    return workbook.value.getSheet(activeSheet.value.workbookSheetIndex);
  }

  const tables = computed(() => { return (
      isWorkerBacked.value
        ? workerTablesByWorkbookSheetIndex.value[activeSheet.value?.workbookSheetIndex ?? -1] ?? []
        : mapWorksheetTables(getActiveWorksheet())
    ); })

  function getCellSnapshotAsync(workbookSheetIndex: number, row: number, col: number) {
    if (!isWorkerBacked.value) {
      return Promise.resolve({
        displayValue: "",
        formula: ""
      });
    }

    return getWorkerClient().getCellSnapshot(workbookSheetIndex, row, col);
  }

  function getRowsBatchAsync(workbookSheetIndex: number, startRow: number, rowCount: number) {
    if (!isWorkerBacked.value) {
      return Promise.resolve(null);
    }

    return getWorkerClient().getRowsBatch(workbookSheetIndex, startRow, rowCount);
  }

  const visibleSheetIndexByWorkbookSheetIndex = computed(() => { return new Map(sheets.value.map((sheet, index) => [sheet.workbookSheetIndex, index])); })

  function mapPublicChart(chart: XlsxChart) {
    const visibleSheetIndex = visibleSheetIndexByWorkbookSheetIndex.value.get(chart.workbookSheetIndex);
    return {
      ...chart,
      sheetIndex: visibleSheetIndex ?? chart.workbookSheetIndex
    };
  }

  function mapPublicImage(image: XlsxImage) {
    const visibleSheetIndex = visibleSheetIndexByWorkbookSheetIndex.value.get(image.workbookSheetIndex);
    return {
      ...image,
      sheetIndex: visibleSheetIndex ?? image.workbookSheetIndex
    };
  }

  function mapPublicFormControl(control: XlsxFormControl) {
    const visibleSheetIndex = visibleSheetIndexByWorkbookSheetIndex.value.get(control.workbookSheetIndex);
    return {
      ...control,
      sheetIndex: visibleSheetIndex ?? control.workbookSheetIndex
    };
  }

  const publicChartsByWorkbookSheetIndex = computed(() => { return chartsByWorkbookSheetIndex.value.map((sheetCharts) => sheetCharts.map(mapPublicChart)); })
  const publicChartById = computed(() => {
    const lookup = new Map<string, XlsxChart>();
    for (const sheetCharts of publicChartsByWorkbookSheetIndex.value) {
      for (const chart of sheetCharts) {
        lookup.set(chart.id, chart);
      }
    }
    return lookup;
  })

  function getSheetCharts(sheetIndex = activeSheetIndex.value) {
    const targetSheet = sheets.value[sheetIndex];
    if (!targetSheet) {
      return [];
    }

    return publicChartsByWorkbookSheetIndex.value[targetSheet.workbookSheetIndex] ?? [];
  }

  function getChartById(id: string) {
    return publicChartById.value.get(id) ?? null;
  }

  function getChartsheetById(id: string) {
  return (
    chartsheets.value.find((chartsheet) => chartsheet.id === id) ?? null
  );
}

  const charts = computed(() => {
    if (activeTab.value?.kind === "chartsheet" && typeof activeTab.value.chartsheetIndex === "number") {
      const chartsheet = chartsheets.value[activeTab.value.chartsheetIndex];
      return (chartsheet?.chartIds ?? []).map((id) => getChartById(id)).filter((value): value is XlsxChart => Boolean(value));
    }

    return getSheetCharts(activeSheetIndex.value);
  })

  const selectedChart = computed(() => { return (selectedChartId.value ? getChartById(selectedChartId.value) : null); })

  watch(
    () => [selectedChart.value, selectedChartElement.value, selectedChartId.value],
    () => {
      if (!selectedChartId.value) {
        if (selectedChartElement.value) {
          selectedChartElement.value = null;
        }
        return;
      }

      if (!selectedChart.value) {
        selectedChartId.value = null;
        selectedChartElement.value = null;
        return;
      }

      if (!selectedChartElement.value) {
        selectedChartElement.value = { chartId: selectedChartId.value, kind: "chart" };
        return;
      }

      if (selectedChartElement.value.chartId !== selectedChartId.value) {
        selectedChartElement.value = { chartId: selectedChartId.value, kind: "chart" };
        return;
      }

      if (selectedChartElement.value.kind !== "chart") {
        const selectedSeries = selectedChart.value.series[selectedChartElement.value.seriesIndex];
        if (!selectedSeries || selectedSeries.id !== selectedChartElement.value.seriesId) {
          selectedChartElement.value = { chartId: selectedChartId.value, kind: "chart" };
        }
      }
    }
  );

  function selectChart(id: string | null) {
    selectedImageId.value = null;
    selectedChartId.value = id;
    selectedChartElement.value = id ? { chartId: id, kind: "chart" } : null;
  }

  function clearSelectedChart() {
    selectedChartId.value = null;
    selectedChartElement.value = null;
  }

  function clearSelectedChartElement() {
    selectedChartElement.value = selectedChartId.value ? { chartId: selectedChartId.value, kind: "chart" } : null;
  }

  function selectChartElement(selection: XlsxChartElementSelection | null) {
    selectedImageId.value = null;
    selectedChartId.value = selection?.chartId ?? null;
    selectedChartElement.value = selection;
  }

  function getSheetImages(sheetIndex = activeSheetIndex.value) {
    const targetSheet = sheets.value[sheetIndex];
    if (!targetSheet) {
      return [];
    }

    return (imagesByWorkbookSheetIndex.value[targetSheet.workbookSheetIndex] ?? []).map(mapPublicImage);
  }

  const images = computed(() => { return getSheetImages(activeSheetIndex.value); })

  function getSheetFormControls(sheetIndex = activeSheetIndex.value) {
    const targetSheet = sheets.value[sheetIndex];
    if (!targetSheet) {
      return [];
    }

    return (formControlsByWorkbookSheetIndex.value[targetSheet.workbookSheetIndex] ?? []).map(mapPublicFormControl);
  }

  const formControls = computed(() => { return getSheetFormControls(activeSheetIndex.value); })

  function getSheetShapes(sheetIndex = activeSheetIndex.value) {
    const targetSheet = sheets.value[sheetIndex];
    if (!targetSheet) {
      return [];
    }

    return (shapesByWorkbookSheetIndex.value[targetSheet.workbookSheetIndex] ?? []).map((shape) => {
      const visibleSheetIndex = visibleSheetIndexByWorkbookSheetIndex.value.get(shape.workbookSheetIndex);
      return {
        ...shape,
        sheetIndex: visibleSheetIndex ?? shape.workbookSheetIndex
      };
    });
  }

  const shapes = computed(() => { return getSheetShapes(activeSheetIndex.value); })

  function getImageById(id: string) {
    for (const sheetImages of imagesByWorkbookSheetIndex.value) {
      const match = sheetImages?.find((image) => image.id === id);
      if (match) {
        return mapPublicImage(match);
      }
    }

    return null;
  }

  const selectedImage = computed(() => { return (selectedImageId.value ? getImageById(selectedImageId.value) : null); })

  function selectImage(id: string | null) {
    selectedChartId.value = null;
    selectedChartElement.value = null;
    selectedImageId.value = id;
  }

  function clearSelectedImage() {
    selectedImageId.value = null;
  }

  function getColumnWidthPx(worksheet: ReturnType<Workbook["getSheet"]>, col: number) {
    const sheetState = imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex[activeSheet.value?.workbookSheetIndex ?? -1] ?? null;
    const width = worksheet.getColumnWidth(col);
    const showGridLines = activeSheet.value?.showGridLines ?? true;
    if (width !== undefined && width !== null) {
      return resolveRenderedSheetAxisPixels(
        resolveSheetColumnWidthPixels(width, sheetState?.columnWidthCharacterWidthPx),
        showGridLines
      );
    }

    return resolveRenderedSheetAxisPixels(
      sheetState?.colWidthOverridesPx?.[col] ?? sheetState?.defaultColWidthPx ?? DEFAULT_COL_WIDTH,
      showGridLines
    );
  }

  function getRowHeightPx(worksheet: ReturnType<Workbook["getSheet"]>, row: number) {
    const sheetState = imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex[activeSheet.value?.workbookSheetIndex ?? -1] ?? null;
    const height = worksheet.getRowHeight(row);
    const showGridLines = activeSheet.value?.showGridLines ?? true;
    if (height !== undefined && height !== null) {
      return resolveRenderedSheetAxisPixels(resolveSheetRowHeightPixels(height), showGridLines);
    }

    return resolveRenderedSheetAxisPixels(
      sheetState?.rowHeightOverridesPx?.[row] ?? sheetState?.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT,
      showGridLines
    );
  }

  function getCellDisplayValue(cell?: XlsxCellAddress | null) {
    if (cell && activeSheet.value) {
      const workerSnapshot = workerCellSnapshotCacheRef.value.get(`${activeSheet.value.workbookSheetIndex}:${cell.row}:${cell.col}`);
      if (workerSnapshot) {
        return workerSnapshot.displayValue;
      }
    }

    const worksheet = getActiveWorksheet();
    if (!worksheet || !cell) {
      return "";
    }

    const formula = worksheet.getFormulaAt(cell.row, cell.col);
    const cachedFormulaValue = formula ? activeSheet.value?.cachedFormulaValues?.[cellAddressToA1(cell)] : undefined;
    const formatted = worksheet.getFormattedValueAt(cell.row, cell.col);
    if (formatted && !(formula && cachedFormulaValue !== undefined && formatted.startsWith("#"))) {
      return decodeHtmlEntities(formatted);
    }

    const calculated = worksheet.getCalculatedValueAt(cell.row, cell.col);
    if (formula && cachedFormulaValue !== undefined && calculated.is_error) {
      return cachedFormulaValue;
    }
    if (calculated.is_error) {
      return calculated.asError() ?? "";
    }
    if (calculated.is_empty) {
      return "";
    }

    return calculated.toString();
  }

  function getCellFormula(cell?: XlsxCellAddress | null) {
    if (cell && activeSheet.value) {
      const workerSnapshot = workerCellSnapshotCacheRef.value.get(`${activeSheet.value.workbookSheetIndex}:${cell.row}:${cell.col}`);
      if (workerSnapshot) {
        return workerSnapshot.formula;
      }
    }

    const worksheet = getActiveWorksheet();
    if (!worksheet || !cell) {
      return "";
    }

    return worksheet.getFormulaAt(cell.row, cell.col) ?? "";
  }

  function getClipboardData() {
    const worksheet = getActiveWorksheet();
    const targetRange = selection.value ?? (activeCell.value ? { start: activeCell.value, end: activeCell.value } : null);
    if (!worksheet || !targetRange) {
      return null;
    }

    const normalized = normalizeRange(targetRange);
    const rows: string[] = [];
    const htmlRows: string[] = [];
    const payload: ClipboardPayload = {
      cells: [],
      cols: normalized.end.col - normalized.start.col + 1,
      merges: [],
      rows: normalized.end.row - normalized.start.row + 1
    };

    for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
      const textCells: string[] = [];
      const htmlCells: string[] = [];

      for (let col = normalized.start.col; col <= normalized.end.col; col += 1) {
        if (worksheet.isMergedSecondary(row, col)) {
          textCells.push("");
          continue;
        }

        const formula = worksheet.getFormulaAt(row, col) ?? null;
        const value = getCellDisplayValue({ row, col });
        const merge = worksheet.getMergeSpan(row, col) as
          | { colSpan?: number; rowSpan?: number }
          | null
          | undefined;
        const rawStyle = (
          worksheet.getCellStyleAt(row, col) as Record<string, unknown> | null | undefined
        ) ?? resolveInheritedCellStyle(activeSheet.value, row, col);
        const cellStyles: string[] = [
          "padding:2px 4px",
          "white-space:pre-wrap",
          "vertical-align:top"
        ];

        const fill = rawStyle?.fill as Record<string, unknown> | undefined;
        if (fill) {
          const fillStyle = resolveWorkbookFillStyle(fill, activeSheet.value?.themePalette);
          if (fillStyle.backgroundColor && fillStyle.backgroundColor.toLowerCase() !== "#ffffff") {
            cellStyles.push(`background-color:${fillStyle.backgroundColor}`);
          }
          if (fillStyle.backgroundImage) {
            cellStyles.push(`background-image:${fillStyle.backgroundImage}`);
          }
        }

        const font = rawStyle?.font as Record<string, unknown> | undefined;
        if (font) {
          if (font.bold) {
            cellStyles.push("font-weight:700");
          }
          if (font.italic) {
            cellStyles.push("font-style:italic");
          }
          if (font.underline && font.underline !== "none") {
            cellStyles.push("text-decoration:underline");
          }
          if (font.strikethrough) {
            cellStyles.push("text-decoration:line-through");
          }
          const fontColor = resolveWorkbookColor(font.color as Record<string, unknown> | undefined, activeSheet.value?.themePalette);
          if (fontColor) {
            cellStyles.push(`color:${fontColor}`);
          }
          if (typeof font.size === "number") {
            cellStyles.push(`font-size:${font.size}pt`);
          }
        }

        const alignment = rawStyle?.alignment as Record<string, unknown> | undefined;
        if (alignment?.horizontal && alignment.horizontal !== "general") {
          cellStyles.push(`text-align:${String(alignment.horizontal)}`);
        }
        if (alignment?.wrapText) {
          cellStyles.push("white-space:pre-wrap");
          cellStyles.push("word-break:break-word");
        }

        const border = rawStyle?.border as Record<string, Record<string, unknown>> | undefined;
        if (border?.top?.style && border.top.style !== "none") {
          cellStyles.push(`border-top:${mapBorder(border.top as { color?: { hex?: string }; style: string })}`);
        }
        if (border?.right?.style && border.right.style !== "none") {
          cellStyles.push(`border-right:${mapBorder(border.right as { color?: { hex?: string }; style: string })}`);
        }
        if (border?.bottom?.style && border.bottom.style !== "none") {
          cellStyles.push(`border-bottom:${mapBorder(border.bottom as { color?: { hex?: string }; style: string })}`);
        }
        if (border?.left?.style && border.left.style !== "none") {
          cellStyles.push(`border-left:${mapBorder(border.left as { color?: { hex?: string }; style: string })}`);
        }

        const rowSpan = Math.min(merge?.rowSpan ?? 1, normalized.end.row - row + 1);
        const colSpan = Math.min(merge?.colSpan ?? 1, normalized.end.col - col + 1);

        payload.cells.push({
          colOffset: col - normalized.start.col,
          formula,
          rowOffset: row - normalized.start.row,
          value
        });

        if (rowSpan > 1 || colSpan > 1) {
          payload.merges.push({
            colOffset: col - normalized.start.col,
            colSpan,
            rowOffset: row - normalized.start.row,
            rowSpan
          });
        }

        textCells.push(value);
        htmlCells.push(
          `<td${rowSpan > 1 ? ` rowspan="${rowSpan}"` : ""}${colSpan > 1 ? ` colspan="${colSpan}"` : ""} style="${escapeHtml(cellStyles.join(";"))}">${escapeHtml(value)}</td>`
        );
      }

      rows.push(textCells.join("\t"));
      htmlRows.push(`<tr>${htmlCells.join("")}</tr>`);
    }

    return {
      html: `<table style="border-collapse:collapse">${htmlRows.join("")}</table>`,
      structured: JSON.stringify(payload),
      text: rows.join("\n")
    };
  }

  watch(
    () => [deferredMetadataCell.value, deferredMetadataSheet.value, isWorkerBacked.value],
    (_value, _oldValue, onCleanup) => {
      if (!isWorkerBacked.value || !deferredMetadataSheet.value || !deferredMetadataCell.value) {
        return;
      }

      const cacheKey = `${deferredMetadataSheet.value.workbookSheetIndex}:${deferredMetadataCell.value.row}:${deferredMetadataCell.value.col}`;
      if (workerCellSnapshotCacheRef.value.has(cacheKey)) {
        return;
      }

      let isCurrent = true;
      void getCellSnapshotAsync(deferredMetadataSheet.value.workbookSheetIndex, deferredMetadataCell.value.row, deferredMetadataCell.value.col)
        .then((snapshot) => {
          if (!isCurrent) {
            return;
          }

          workerCellSnapshotCacheRef.value.set(cacheKey, snapshot);
          workerCellSnapshotRevision.value = workerCellSnapshotRevision.value + 1;
        })
        .catch(() => {
          if (!isCurrent) {
            return;
          }

          workerCellSnapshotCacheRef.value.set(cacheKey, {
            displayValue: "",
            formula: ""
          });
          workerCellSnapshotRevision.value = workerCellSnapshotRevision.value + 1;
        });

      onCleanup(() => {
        isCurrent = false;
      });
    }
  );

  const activeCellAddress = computed(() => { return (activeCell.value ? cellAddressToA1(activeCell.value) : null); })
  const selectedRangeAddress = computed(() => { return (selection.value ? rangeToA1(selection.value) : null); })
  const selectedValue = computed(() => { return getCellDisplayValue(deferredMetadataCell.value); })
  const selectedCellFormula = computed(() => { return getCellFormula(deferredMetadataCell.value); })
  function getChartSeriesFormula(chartId: string, seriesIndex: number) {
  return (
    buildChartSeriesFormula(getChartById(chartId), seriesIndex)
  );
}
  const selectedChartFormula = computed(() => {
    if (
      !selectedChartElement.value
      || selectedChartElement.value.kind === "chart"
      || selectedChartElement.value.seriesIndex < 0
    ) {
      return null;
    }

    return getChartSeriesFormula(selectedChartElement.value.chartId, selectedChartElement.value.seriesIndex);
  })
  const selectedFormulaTarget = computed(() => {
    if (selectedChartFormula.value && selectedChartElement.value && selectedChartElement.value.kind !== "chart") {
      return {
        chartId: selectedChartElement.value.chartId,
        kind: "chartSeries" as const,
        seriesId: selectedChartElement.value.seriesId,
        seriesIndex: selectedChartElement.value.seriesIndex
      };
    }

    return {
      cell: deferredMetadataCell.value,
      kind: "cell" as const
    };
  })
  const selectedFormula = computed(() => selectedChartFormula.value ?? selectedCellFormula.value);
  const isLoadDeferred = computed(() => deferredLoadFileSize.value !== null);
  const canLoadDeferred = computed(() => !isLoading.value && isLoadDeferred.value);
  const canUndo = computed(() => !readOnly.value && undoStackRef.value.length > 0);
  const canRedo = computed(() => !readOnly.value && redoStackRef.value.length > 0);

  function createSavedWorkbookBytes(targetWorkbook: Workbook) {
    const sanitizedBytes = sanitizeSavedWorkbookBytes(targetWorkbook.saveXlsxBytes());
    return mergeWorkbookImageAssets(sanitizedBytes, imageAssetsRef.value, sheetOriginsRef.value);
  }

  function createHistoryEntry(): SnapshotHistoryEntry | null {
    if (!workbook.value) {
      return null;
    }

    return {
      kind: "snapshot",
      activeCell: activeCell.value,
      activeSheetIndex: activeSheetIndex.value,
      bytes: createSavedWorkbookBytes(workbook.value),
      selection: selection.value
    };
  }

  function captureCellMutationState(cell: XlsxCellAddress) {
    const worksheet = getActiveWorksheet();
    if (!worksheet) {
      return null;
    }

    return {
      formula: worksheet.getFormulaAt(cell.row, cell.col) ?? null,
      style: worksheet.getCellStyleAt(cell.row, cell.col),
      value: worksheet.getCellAt(cell.row, cell.col).toJs()
    };
  }

  async function restoreHistoryEntry(entry: SnapshotHistoryEntry) {
    const wasmModule = await getSheetsWasmModule();
    const nextWorkbook = wasmModule.Workbook.fromBytes(cloneBytes(entry.bytes));
    const nextImageAssets = loadWorkbookImageAssets(entry.bytes, nextWorkbook);
    const nextSheets = buildSheetList(
      nextWorkbook,
      nextImageAssets.sheetStatesByWorkbookSheetIndex,
      nextImageAssets.themePalette,
      nextImageAssets.styleById,
      nextImageAssets.namedCellStyleByName,
      nextImageAssets.tableStyleByName,
      showHiddenSheets
    );
    const nextSheetIndex = Math.max(0, Math.min(entry.activeSheetIndex, Math.max(0, nextSheets.length - 1)));

    error.value = null;
    isLoading.value = false;
    setImageAssets(nextImageAssets);
    workbook.value = nextWorkbook;
    sheets.value = nextSheets;
    const nextChartAssets = loadWorkbookChartAssets(nextWorkbook, nextImageAssets, buildVisibleSheetIndexMap(nextSheets), showHiddenSheets);
    setChartAssets(nextChartAssets);
    activeSheetIndex.value = nextSheetIndex;
    const nextTabIndex = nextChartAssets.tabs.findIndex((tab) => tab.kind === "sheet" && tab.sheetIndex === nextSheetIndex);
    if (nextTabIndex >= 0) {
      activeTabIndex.value = nextTabIndex;
    }
    activeCell.value = entry.activeCell;
    selection.value = entry.selection;
    selectionAnchorRef.value = entry.selection ? normalizeRange(entry.selection).start : entry.activeCell;
    revision.value = revision.value + 1;
  }

  function applyCellEditHistoryEntry(entry: CellEditHistoryEntry, direction: "undo" | "redo") {
    if (!workbook.value) {
      return;
    }

    const worksheet = workbook.value.getSheet(entry.sheetIndex);
    const visibleSheetIndex = sheets.value.findIndex((sheet) => sheet.workbookSheetIndex === entry.sheetIndex);
    const targetState = direction === "undo" ? entry.before : entry.after;

    isApplyingHistoryRef.value = true;
    applyCellMutationState(worksheet, entry.cell, targetState);
    maybeRecalculateWorkbook(workbook.value);
    refreshWorkbookState(workbook.value);

    const nextActiveCell = direction === "undo" ? entry.activeCellBefore : entry.activeCellAfter;
    const nextSelection = direction === "undo" ? entry.selectionBefore : entry.selectionAfter;
    if (visibleSheetIndex >= 0) {
      activeSheetIndex.value = visibleSheetIndex;
    }
    activeCell.value = nextActiveCell;
    selection.value = nextSelection;
    selectionAnchorRef.value = nextSelection ? normalizeRange(nextSelection).start : nextActiveCell;
    isApplyingHistoryRef.value = false;
  }

  function applyRangeEditHistoryEntry(entry: RangeEditHistoryEntry, direction: "undo" | "redo") {
    if (!workbook.value) {
      return;
    }

    const worksheet = workbook.value.getSheet(entry.sheetIndex);
    const visibleSheetIndex = sheets.value.findIndex((sheet) => sheet.workbookSheetIndex === entry.sheetIndex);

    isApplyingHistoryRef.value = true;
    for (const mutation of entry.mutations) {
      applyCellMutationState(worksheet, mutation.cell, direction === "undo" ? mutation.before : mutation.after);
    }
    maybeRecalculateWorkbook(workbook.value);
    refreshWorkbookState(workbook.value);

    const nextActiveCell = direction === "undo" ? entry.activeCellBefore : entry.activeCellAfter;
    const nextSelection = direction === "undo" ? entry.selectionBefore : entry.selectionAfter;
    if (visibleSheetIndex >= 0) {
      activeSheetIndex.value = visibleSheetIndex;
    }
    activeCell.value = nextActiveCell;
    selection.value = nextSelection;
    selectionAnchorRef.value = nextSelection ? normalizeRange(nextSelection).start : nextActiveCell;
    isApplyingHistoryRef.value = false;
  }

  function recordHistoryBeforeMutation() {
    if (isApplyingHistoryRef.value) {
      return;
    }

    const snapshot = createHistoryEntry();
    if (!snapshot) {
      return;
    }

    pushHistoryEntry(undoStackRef.value, snapshot);
    redoStackRef.value = [];
    historyRevision.value = historyRevision.value + 1;
  }

  function recordCellEditHistory(cell: XlsxCellAddress, before: CellMutationState, after: CellMutationState) {
    if (!activeSheet.value || isApplyingHistoryRef.value) {
      return;
    }

    pushHistoryEntry(undoStackRef.value, {
      kind: "cell-edit",
      activeCellAfter: cell,
      activeCellBefore: activeCell.value,
      after,
      before,
      cell,
      selectionAfter: { start: cell, end: cell },
      selectionBefore: selection.value,
      sheetIndex: activeSheet.value.workbookSheetIndex
    });
    redoStackRef.value = [];
    historyRevision.value = historyRevision.value + 1;
  }

  function recordRangeEditHistory(mutations: RangeCellMutation[], selectionAfter: XlsxCellRange | null, activeCellAfter: XlsxCellAddress | null) {
    if (!activeSheet.value || isApplyingHistoryRef.value || mutations.length === 0) {
      return;
    }

    pushHistoryEntry(undoStackRef.value, {
      kind: "range-edit",
      activeCellAfter,
      activeCellBefore: activeCell.value,
      mutations,
      selectionAfter,
      selectionBefore: selection.value,
      sheetIndex: activeSheet.value.workbookSheetIndex
    });
    redoStackRef.value = [];
    historyRevision.value = historyRevision.value + 1;
  }

  function sortTable(tableName: string, columnIndex: number, direction: XlsxTableSortDirection) {
    const worksheet = getActiveWorksheet();
    const targetTable = tables.value.find((table) => table.name === tableName || table.displayName === tableName);
    if (!worksheet || !workbook.value || !activeSheet.value || !targetTable) {
      return;
    }

    const dataStartRow = targetTable.start.row + Math.max(targetTable.headerRowCount, 1);
    const totalsRowOffset = targetTable.totalsRowShown ? Math.max(targetTable.totalsRowCount, 1) : 0;
    const dataEndRow = targetTable.end.row - totalsRowOffset;
    const startCol = targetTable.start.col;
    const endCol = targetTable.end.col;
    const sortCol = startCol + columnIndex;

    if (columnIndex < 0 || sortCol > endCol || dataStartRow > dataEndRow) {
      return;
    }

    const rows: Array<{
      cells: CellMutationState[];
      index: number;
      sortBoolean: boolean | undefined;
      sortEmpty: boolean;
      sortNumber: number | undefined;
      sortText: string;
    }> = [];

    for (let row = dataStartRow; row <= dataEndRow; row += 1) {
      const cells: CellMutationState[] = [];
      for (let col = startCol; col <= endCol; col += 1) {
        cells.push({
          formula: worksheet.getFormulaAt(row, col) ?? null,
          style: worksheet.getCellStyleAt(row, col),
          value: worksheet.getCellAt(row, col).toJs()
        });
      }

      const calculated = worksheet.getCalculatedValueAt(row, sortCol);
      const formatted = decodeHtmlEntities(worksheet.getFormattedValueAt(row, sortCol) ?? "");
      rows.push({
        cells,
        index: row,
        sortBoolean: calculated.is_boolean ? calculated.asBoolean() : undefined,
        sortEmpty: calculated.is_empty || formatted.length === 0,
        sortNumber: calculated.is_number ? calculated.asNumber() : undefined,
        sortText: calculated.is_text ? (calculated.asText() ?? formatted) : formatted
      });
    }

    const sortedRows = [...rows].sort((left, right) => {
      if (left.sortEmpty !== right.sortEmpty) {
        return left.sortEmpty ? 1 : -1;
      }

      if (left.sortNumber !== undefined && right.sortNumber !== undefined) {
        return direction === "ascending" ? left.sortNumber - right.sortNumber : right.sortNumber - left.sortNumber;
      }

      if (left.sortBoolean !== undefined && right.sortBoolean !== undefined) {
        const leftValue = left.sortBoolean ? 1 : 0;
        const rightValue = right.sortBoolean ? 1 : 0;
        return direction === "ascending" ? leftValue - rightValue : rightValue - leftValue;
      }

      const comparedText = left.sortText.localeCompare(right.sortText, undefined, { numeric: true, sensitivity: "base" });
      return direction === "ascending" ? comparedText : -comparedText;
    });

    if (sortedRows.every((row, index) => row.index === rows[index]?.index)) {
      sortState.value = { columnIndex, direction, tableName: targetTable.name };
      return;
    }

    const mutations: RangeCellMutation[] = [];
    for (let rowOffset = 0; rowOffset < rows.length; rowOffset += 1) {
      const targetRow = dataStartRow + rowOffset;
      const sourceRow = sortedRows[rowOffset];
      const beforeRow = rows[rowOffset];
      if (!sourceRow || !beforeRow) {
        continue;
      }

      for (let colOffset = 0; colOffset <= endCol - startCol; colOffset += 1) {
        const before = beforeRow.cells[colOffset];
        const after = sourceRow.cells[colOffset];
        if (!before || !after) {
          continue;
        }

        const cell = { row: targetRow, col: startCol + colOffset };
        applyCellMutationState(worksheet, cell, after);
        mutations.push({
          after,
          before,
          cell
        });
      }
    }

    maybeRecalculateWorkbook(workbook.value);
    refreshWorkbookState(workbook.value);
    sortState.value = { columnIndex, direction, tableName: targetTable.name };
    recordRangeEditHistory(mutations, selection.value, activeCell.value);
  }

  function download() {
    if (file) {
      downloadArrayBuffer(file, displayFileName.value);
      return;
    }

    if (src) {
      downloadUrl(src, displayFileName.value);
    }
  }

  function exportXlsx() {
    if (!workbook.value) {
      return;
    }

    downloadBytes(createSavedWorkbookBytes(workbook.value), `${fileStem(displayFileName.value)}.xlsx`, XLSX_MIME_TYPE);
  }

  function exportCsv() {
    if (!workbook.value) {
      return;
    }

    const activeSheetName = activeSheet.value?.name ?? "sheet";
    downloadText(workbook.value.saveCsvString(), `${fileStem(displayFileName.value)}-${activeSheetName}.csv`, CSV_MIME_TYPE);
  }

  function recalculate() {
    if (!workbook.value) {
      return;
    }

    const result = tryRecalculate(workbook.value);
    if (result.calculated) {
      refreshWorkbookState(workbook.value);
      return;
    }

    // Trap poisons the Workbook pointer; skip refreshWorkbookState so we
    // don't crash reading cells from it.
    shouldAutoCalculate.value = false;
  }

  function applyReadOnlyResizeOverride(axis: "column" | "row", actualIndex: number, sizePx: number) {
    if (!activeSheet.value) {
      return;
    }

    const contentSizePx = resolveContentSheetAxisPixels(sizePx, activeSheet.value.showGridLines);
    const targetWorkbookSheetIndex = activeSheet.value.workbookSheetIndex;
    sheets.value = sheets.value.map((sheet) => {
      if (sheet.workbookSheetIndex !== targetWorkbookSheetIndex) {
        return sheet;
      }

      if (axis === "column") {
        const nextColWidthOverridesPx = {
          ...sheet.colWidthOverridesPx,
          [actualIndex]: contentSizePx
        };
        const nextColWidths = [...sheet.colWidths];
        const visibleColIndex = sheet.visibleCols.indexOf(actualIndex);
        if (visibleColIndex >= 0) {
          nextColWidths[visibleColIndex] = contentSizePx;
        }

        return {
          ...sheet,
          colWidthOverridesPx: nextColWidthOverridesPx,
          colWidths: nextColWidths
        };
      }

      const nextRowHeightOverridesPx = {
        ...sheet.rowHeightOverridesPx,
        [actualIndex]: contentSizePx
      };
      const nextRowHeights = [...sheet.rowHeights];
      const visibleRowIndex = sheet.visibleRows.indexOf(actualIndex);
      if (visibleRowIndex >= 0) {
        nextRowHeights[visibleRowIndex] = contentSizePx;
      }

      return {
        ...sheet,
        rowHeightOverridesPx: nextRowHeightOverridesPx,
        rowHeights: nextRowHeights
      };
    });
    revision.value = revision.value + 1;
  }

  function resizeColumn(col: number, widthPx: number) {
    if ((readOnly.value && !canResizeReadOnly.value) || !activeSheet.value) {
      return;
    }

    if (isWorkerBacked.value) {
      applyReadOnlyResizeOverride("column", col, widthPx);
      return;
    }

    if (!workbook.value) {
      return;
    }

    recordHistoryBeforeMutation();
    const worksheet = workbook.value.getSheet(activeSheet.value.workbookSheetIndex);
    worksheet.setColumnWidth(
      col,
      pxToSheetColumnWidth(resolveContentSheetAxisPixels(widthPx, activeSheet.value.showGridLines))
    );
    refreshWorkbookState(workbook.value);
  }

  function resizeRow(row: number, heightPx: number) {
    if ((readOnly.value && !canResizeReadOnly.value) || !activeSheet.value) {
      return;
    }

    if (isWorkerBacked.value) {
      applyReadOnlyResizeOverride("row", row, heightPx);
      return;
    }

    if (!workbook.value) {
      return;
    }

    recordHistoryBeforeMutation();
    const worksheet = workbook.value.getSheet(activeSheet.value.workbookSheetIndex);
    worksheet.setRowHeight(
      row,
      pxToSheetRowHeight(resolveContentSheetAxisPixels(heightPx, activeSheet.value.showGridLines))
    );
    refreshWorkbookState(workbook.value);
  }

  function resolveAnchoredObjectRect(anchor: XlsxImage["anchor"], worksheet: ReturnType<Workbook["getSheet"]>) {
    const resolveAxisSum = (
      index: number,
      getSize: (target: number) => number
    ) => {
      let total = 0;
      for (let cursor = 0; cursor < index; cursor += 1) {
        total += getSize(cursor);
      }
      return total;
    };

    if (anchor.kind === "absolute") {
      return {
        height: anchor.sizeEmu.cy / 9525,
        left: GRID_ROW_HEADER_WIDTH + anchor.positionEmu.x / 9525,
        top: GRID_HEADER_HEIGHT + anchor.positionEmu.y / 9525,
        width: anchor.sizeEmu.cx / 9525
      };
    }

    const left = GRID_ROW_HEADER_WIDTH + resolveAxisSum(anchor.from.col, (col) => getColumnWidthPx(worksheet, col)) + anchor.from.colOffsetEmu / 9525;
    const top = GRID_HEADER_HEIGHT + resolveAxisSum(anchor.from.row, (row) => getRowHeightPx(worksheet, row)) + anchor.from.rowOffsetEmu / 9525;

    if (anchor.kind === "one-cell") {
      return {
        height: anchor.sizeEmu.cy / 9525,
        left,
        top,
        width: anchor.sizeEmu.cx / 9525
      };
    }

    const right = GRID_ROW_HEADER_WIDTH + resolveAxisSum(anchor.to.col, (col) => getColumnWidthPx(worksheet, col)) + anchor.to.colOffsetEmu / 9525;
    const bottom = GRID_HEADER_HEIGHT + resolveAxisSum(anchor.to.row, (row) => getRowHeightPx(worksheet, row)) + anchor.to.rowOffsetEmu / 9525;

    return {
      height: Math.max(1, bottom - top),
      left,
      top,
      width: Math.max(1, right - left)
    };
  }

  function setChartRect(id: string, rect: XlsxImageRect) {
    const hydratedChartAssets = ensureChartAssetsHydrated(workbook.value, sheets.value);
    console.info("[react-xlsx debug] setChartRect", {
      hasActiveSheet: Boolean(activeSheet.value),
      hasHydratedChartAssets: Boolean(hydratedChartAssets),
      hasImageAssets: Boolean(imageAssetsRef.value),
      hasWorkbook: Boolean(workbook.value),
      id,
      readOnly: readOnly.value,
      rect
    });
    if (readOnly.value || !workbook.value || !activeSheet.value || !imageAssetsRef.value || !hydratedChartAssets) {
      return;
    }

    const worksheet = workbook.value.getSheet(activeSheet.value.workbookSheetIndex);
    const currentChart = getChartById(id);
    console.info("[react-xlsx debug] currentChart", {
      activeWorkbookSheetIndex: activeSheet.value.workbookSheetIndex,
      editable: currentChart?.editable,
      found: Boolean(currentChart),
      originCount: hydratedChartAssets.chartOriginsById.size,
      workbookSheetIndex: currentChart?.workbookSheetIndex
    });
    if (!currentChart || currentChart.editable === false || currentChart.workbookSheetIndex !== activeSheet.value.workbookSheetIndex) {
      return;
    }

    const nextAnchor = rectToImageAnchor(rect, currentChart.anchor, {
      contentOffsetLeft: GRID_ROW_HEADER_WIDTH,
      contentOffsetTop: GRID_HEADER_HEIGHT,
      getColumnWidthPx: (col) => getColumnWidthPx(worksheet, col),
      getRowHeightPx: (row) => getRowHeightPx(worksheet, row)
    });

    recordHistoryBeforeMutation();
    const didUpdateAnchor = updateWorkbookChartAnchor(imageAssetsRef.value, hydratedChartAssets, id, nextAnchor);
    console.info("[react-xlsx debug] updateWorkbookChartAnchor", { didUpdateAnchor, nextAnchor });

    hydratedChartAssets.chartsByWorkbookSheetIndex = hydratedChartAssets.chartsByWorkbookSheetIndex.map((sheetCharts) => (
      sheetCharts.map((chart) => chart.id === id ? { ...chart, anchor: nextAnchor } : chart)
    ));

    chartsByWorkbookSheetIndex.value = chartsByWorkbookSheetIndex.value.map((sheetCharts) => (
      sheetCharts.map((chart) => chart.id === id ? { ...chart, anchor: nextAnchor } : chart)
    ));
    revision.value = revision.value + 1;
  }

  function setImageRect(id: string, rect: XlsxImageRect) {
    if (readOnly.value || !workbook.value || !activeSheet.value || !imageAssetsRef.value) {
      return;
    }

    const worksheet = workbook.value.getSheet(activeSheet.value.workbookSheetIndex);
    const currentImage = getImageById(id);
    if (!currentImage || currentImage.editable === false || currentImage.workbookSheetIndex !== activeSheet.value.workbookSheetIndex) {
      return;
    }

    const nextAnchor = rectToImageAnchor(rect, currentImage.anchor, {
      contentOffsetLeft: GRID_ROW_HEADER_WIDTH,
      contentOffsetTop: GRID_HEADER_HEIGHT,
      getColumnWidthPx: (col) => getColumnWidthPx(worksheet, col),
      getRowHeightPx: (row) => getRowHeightPx(worksheet, row)
    });

    recordHistoryBeforeMutation();
    if (!updateWorkbookImageAnchor(imageAssetsRef.value, id, nextAnchor)) {
      return;
    }

    imagesByWorkbookSheetIndex.value = [...imageAssetsRef.value.imagesByWorkbookSheetIndex];
    revision.value = revision.value + 1;
  }

  function moveChartBy(id: string, deltaX: number, deltaY: number) {
    const currentChart = getChartById(id);
    if (!currentChart || currentChart.editable === false) {
      return;
    }

    const worksheet = getActiveWorksheet();
    if (!worksheet) {
      return;
    }

    const currentRect = resolveAnchoredObjectRect(currentChart.anchor, worksheet);
    setChartRect(id, {
      ...currentRect,
      left: currentRect.left + deltaX,
      top: currentRect.top + deltaY
    });
  }

  function moveImageBy(id: string, deltaX: number, deltaY: number) {
    const currentImage = getImageById(id);
    if (!currentImage || currentImage.editable === false) {
      return;
    }

    const currentRect = (() => {
      const worksheet = getActiveWorksheet();
      if (!worksheet) {
        return null;
      }

      const resolveAxisSum = (
        index: number,
        getSize: (target: number) => number
      ) => {
        let total = 0;
        for (let cursor = 0; cursor < index; cursor += 1) {
          total += getSize(cursor);
        }
        return total;
      };

      if (currentImage.anchor.kind === "absolute") {
        return {
          height: currentImage.anchor.sizeEmu.cy / 9525,
          left: GRID_ROW_HEADER_WIDTH + currentImage.anchor.positionEmu.x / 9525,
          top: GRID_HEADER_HEIGHT + currentImage.anchor.positionEmu.y / 9525,
          width: currentImage.anchor.sizeEmu.cx / 9525
        };
      }

      const left = GRID_ROW_HEADER_WIDTH + resolveAxisSum(currentImage.anchor.from.col, (col) => getColumnWidthPx(worksheet, col)) + currentImage.anchor.from.colOffsetEmu / 9525;
      const top = GRID_HEADER_HEIGHT + resolveAxisSum(currentImage.anchor.from.row, (row) => getRowHeightPx(worksheet, row)) + currentImage.anchor.from.rowOffsetEmu / 9525;

      if (currentImage.anchor.kind === "one-cell") {
        return {
          height: currentImage.anchor.sizeEmu.cy / 9525,
          left,
          top,
          width: currentImage.anchor.sizeEmu.cx / 9525
        };
      }

      const right = GRID_ROW_HEADER_WIDTH + resolveAxisSum(currentImage.anchor.to.col, (col) => getColumnWidthPx(worksheet, col)) + currentImage.anchor.to.colOffsetEmu / 9525;
      const bottom = GRID_HEADER_HEIGHT + resolveAxisSum(currentImage.anchor.to.row, (row) => getRowHeightPx(worksheet, row)) + currentImage.anchor.to.rowOffsetEmu / 9525;

      return {
        height: Math.max(1, bottom - top),
        left,
        top,
        width: Math.max(1, right - left)
      };
    })();

    if (!currentRect) {
      return;
    }

    setImageRect(id, {
      ...currentRect,
      left: currentRect.left + deltaX,
      top: currentRect.top + deltaY
    });
  }

  function resizeChartBy(id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) {
    const currentChart = getChartById(id);
    if (!currentChart || currentChart.editable === false) {
      return;
    }

    const worksheet = getActiveWorksheet();
    if (!worksheet) {
      return;
    }

    const currentRect = resolveAnchoredObjectRect(currentChart.anchor, worksheet);
    setChartRect(id, resizeImageRect(currentRect, handle, deltaX, deltaY, 48));
  }

  function resizeImageBy(id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) {
    const currentImage = getImageById(id);
    if (!currentImage || currentImage.editable === false) {
      return;
    }

    const worksheet = getActiveWorksheet();
    if (!worksheet) {
      return;
    }

    const resolveAxisSum = (
      index: number,
      getSize: (target: number) => number
    ) => {
      let total = 0;
      for (let cursor = 0; cursor < index; cursor += 1) {
        total += getSize(cursor);
      }
      return total;
    };

    const left = currentImage.anchor.kind === "absolute"
      ? GRID_ROW_HEADER_WIDTH + currentImage.anchor.positionEmu.x / 9525
      : GRID_ROW_HEADER_WIDTH + resolveAxisSum(currentImage.anchor.from.col, (col) => getColumnWidthPx(worksheet, col)) + currentImage.anchor.from.colOffsetEmu / 9525;
    const top = currentImage.anchor.kind === "absolute"
      ? GRID_HEADER_HEIGHT + currentImage.anchor.positionEmu.y / 9525
      : GRID_HEADER_HEIGHT + resolveAxisSum(currentImage.anchor.from.row, (row) => getRowHeightPx(worksheet, row)) + currentImage.anchor.from.rowOffsetEmu / 9525;
    const width = currentImage.anchor.kind === "two-cell"
      ? Math.max(
          1,
          GRID_ROW_HEADER_WIDTH + resolveAxisSum(currentImage.anchor.to.col, (col) => getColumnWidthPx(worksheet, col)) + currentImage.anchor.to.colOffsetEmu / 9525 - left
        )
      : currentImage.anchor.sizeEmu.cx / 9525;
    const height = currentImage.anchor.kind === "two-cell"
      ? Math.max(
          1,
          GRID_HEADER_HEIGHT + resolveAxisSum(currentImage.anchor.to.row, (row) => getRowHeightPx(worksheet, row)) + currentImage.anchor.to.rowOffsetEmu / 9525 - top
        )
      : currentImage.anchor.sizeEmu.cy / 9525;

    const nextRect = resizeImageRect({ height, left, top, width }, handle, deltaX, deltaY);
    setImageRect(id, nextRect);
  }

  function updateChart(id: string, patch: Partial<XlsxChart>) {
    const currentChart = getChartById(id);
    const hydratedChartAssets = ensureChartAssetsHydrated(workbook.value, sheets.value);
    if (readOnly.value || !currentChart) {
      return;
    }

    recordHistoryBeforeMutation();
    if (patch.anchor && imageAssetsRef.value && hydratedChartAssets) {
      updateWorkbookChartAnchor(imageAssetsRef.value, hydratedChartAssets, id, patch.anchor);
    }
    if (imageAssetsRef.value && hydratedChartAssets) {
      updateWorkbookChartDefinition(imageAssetsRef.value, hydratedChartAssets, id, patch);
    }

    chartsByWorkbookSheetIndex.value = chartsByWorkbookSheetIndex.value.map((sheetCharts) => (
      sheetCharts.map((chart) => chart.id === id ? { ...chart, ...patch } : chart)
    ));
    revision.value = revision.value + 1;
  }

  function setChartSeriesFormula(chartId: string, seriesIndex: number, formula: string) {
    if (readOnly.value) {
      return false;
    }

    const chart = getChartById(chartId);
    if (!chart || chart.editable === false) {
      return false;
    }

    const nextChart = applyChartSeriesFormula(chart, seriesIndex, formula, workbook.value);
    if (!nextChart) {
      return false;
    }

    updateChart(chartId, { series: nextChart.series });
    const selectedSeries = nextChart.series[seriesIndex];
    if (selectedSeries) {
      selectedChartElement.value = (
        selectedChartElement.value && selectedChartElement.value.chartId === chartId && selectedChartElement.value.kind !== "chart"
          ? {
              ...selectedChartElement.value,
              seriesId: selectedSeries.id,
              seriesIndex
            }
          : selectedChartElement.value
      );
    }
    return true;
  }

  function selectCell(cell: XlsxCellAddress, options?: { extend?: boolean }) {
    selectedChartId.value = null;
    selectedChartElement.value = null;
    selectedImageId.value = null;
    activeCell.value = cell;
    if (options?.extend && selectionAnchorRef.value) {
      selection.value = normalizeRange({ start: selectionAnchorRef.value, end: cell });
      return;
    }

    selectionAnchorRef.value = cell;
    selection.value = { start: cell, end: cell };
  }

  function selectRange(range: XlsxCellRange) {
    const normalized = normalizeRange(range);
    selectedChartId.value = null;
    selectedChartElement.value = null;
    selectedImageId.value = null;
    selectionAnchorRef.value = normalized.start;
    activeCell.value = normalized.end;
    selection.value = normalized;
  }

  function clearSelection() {
    selectionAnchorRef.value = null;
    activeCell.value = null;
    selection.value = null;
    selectedChartId.value = null;
    selectedChartElement.value = null;
    selectedImageId.value = null;
  }

  function clearSelectedCells() {
    const worksheet = getActiveWorksheet();
    const targetRange = selection.value ?? (activeCell.value ? { start: activeCell.value, end: activeCell.value } : null);
    if (readOnly.value || !worksheet || !workbook.value || !targetRange) {
      return;
    }

    const normalized = normalizeRange(targetRange);
    const mutations: RangeCellMutation[] = [];
    for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
      for (let col = normalized.start.col; col <= normalized.end.col; col += 1) {
        if (worksheet.isMergedSecondary(row, col)) {
          continue;
        }

        const cell = { row, col };
        const before = captureCellMutationState(cell);
        if (!before) {
          continue;
        }

        worksheet.setCell(cellAddressToA1({ row, col }), "");
        const after = captureCellMutationState(cell);
        if (!after) {
          continue;
        }
        mutations.push({
          after,
          before,
          cell
        });
      }
    }

    maybeRecalculateWorkbook(workbook.value);
    refreshWorkbookState(workbook.value);
    recordRangeEditHistory(mutations, normalized, activeCell.value ?? normalized.start);
  }

  function setCellValue(cell: XlsxCellAddress, value: string) {
    const worksheet = getActiveWorksheet();
    if (readOnly.value || !worksheet || !workbook.value) {
      return;
    }

    const before = captureCellMutationState(cell);
    if (!before) {
      return;
    }

    const nextValue = coerceUserEnteredValue(value);
    worksheet.setCell(cellAddressToA1(cell), nextValue);
    const after = captureCellMutationState(cell);
    if (!after) {
      return;
    }
    maybeRecalculateWorkbook(workbook.value);
    refreshWorkbookState(workbook.value);
    recordCellEditHistory(cell, before, after);
  }

  function setCellFormula(cell: XlsxCellAddress, formula: string) {
    const worksheet = getActiveWorksheet();
    if (readOnly.value || !worksheet || !workbook.value) {
      return;
    }

    const before = captureCellMutationState(cell);
    if (!before) {
      return;
    }

    const trimmedFormula = formula.trim();
    if (!formula.trim()) {
      worksheet.setCell(cellAddressToA1(cell), "");
    } else {
      worksheet.setFormula(cellAddressToA1(cell), formula);
    }
    const after = captureCellMutationState(cell);
    if (!after) {
      return;
    }
    maybeRecalculateWorkbook(workbook.value);
    refreshWorkbookState(workbook.value);
    recordCellEditHistory(cell, before, after);
  }

  function setCellStyle(cell: XlsxCellAddress, style: XlsxCellStyleInput) {
    const worksheet = getActiveWorksheet();
    if (readOnly.value || !worksheet || !workbook.value) {
      return;
    }

    const before = captureCellMutationState(cell);
    if (!before) {
      return;
    }

    worksheet.setCellStyleAt(cell.row, cell.col, style);
    const after = captureCellMutationState(cell);
    if (!after) {
      return;
    }

    refreshWorkbookState(workbook.value);
    recordCellEditHistory(cell, before, after);
  }

  function setSelectedCellValue(value: string) {
    if (!activeCell.value) {
      return;
    }

    setCellValue(activeCell.value, value);
  }

  function setSelectedCellFormula(formula: string) {
    if (!activeCell.value) {
      return;
    }

    setCellFormula(activeCell.value, formula);
  }

  function setSelectedFormula(formula: string) {
    if (selectedFormulaTarget.value?.kind === "chartSeries") {
      return setChartSeriesFormula(selectedFormulaTarget.value.chartId, selectedFormulaTarget.value.seriesIndex, formula);
    }

    if (!activeCell.value) {
      return false;
    }

    setCellFormula(activeCell.value, formula);
    return true;
  }

  function setSelectedCellStyle(style: XlsxCellStyleInput) {
    if (!activeCell.value) {
      return;
    }

    setCellStyle(activeCell.value, style);
  }

  function setRangeStyle(range: XlsxCellRange, style: XlsxCellStyleInput) {
    const worksheet = getActiveWorksheet();
    if (readOnly.value || !worksheet || !workbook.value) {
      return;
    }

    const normalized = normalizeRange(range);
    const beforeStates: Array<{ before: CellMutationState; cell: XlsxCellAddress }> = [];
    for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
      for (let col = normalized.start.col; col <= normalized.end.col; col += 1) {
        const cell = { row, col };
        const before = captureCellMutationState(cell);
        if (!before) {
          continue;
        }
        beforeStates.push({
          before,
          cell
        });
      }
    }

    if (beforeStates.length === 0) {
      return;
    }

    worksheet.setRangeStyle(rangeToA1(normalized), style);
    const mutations: RangeCellMutation[] = [];
    for (const mutation of beforeStates) {
      const after = captureCellMutationState(mutation.cell);
      if (!after) {
        continue;
      }
      mutations.push({
        after,
        before: mutation.before,
        cell: mutation.cell
      });
    }

    refreshWorkbookState(workbook.value);
    recordRangeEditHistory(mutations, selection.value, activeCell.value);
  }

  function fillSelection(targetRange: XlsxCellRange) {
    const worksheet = getActiveWorksheet();
    if (readOnly.value || !worksheet || !workbook.value || !selection.value) {
      return;
    }

    const sourceRange = normalizeRange(selection.value);
    const nextRange = normalizeRange(targetRange);
    const sourceHeight = sourceRange.end.row - sourceRange.start.row + 1;
    const sourceWidth = sourceRange.end.col - sourceRange.start.col + 1;

    if (sourceHeight <= 0 || sourceWidth <= 0) {
      return;
    }

    const mutations: RangeCellMutation[] = [];
    for (let row = nextRange.start.row; row <= nextRange.end.row; row += 1) {
      for (let col = nextRange.start.col; col <= nextRange.end.col; col += 1) {
        if (rangeContainsCell(sourceRange, { row, col })) {
          continue;
        }

        const targetCell = { row, col };
        const before = captureCellMutationState(targetCell);
        if (!before) {
          continue;
        }

        const sourceRow = sourceRange.start.row + ((row - nextRange.start.row) % sourceHeight);
        const sourceCol = sourceRange.start.col + ((col - nextRange.start.col) % sourceWidth);
        const sourceFormula = worksheet.getFormulaAt(sourceRow, sourceCol);
        const sourceStyle = cloneCellStyle(worksheet.getCellStyleAt(sourceRow, sourceCol));

        if (sourceFormula) {
          worksheet.setFormula(cellAddressToA1(targetCell), sourceFormula);
        } else {
          const sourceValue = normalizeCellValue(worksheet.getCellAt(sourceRow, sourceCol).toJs());
          worksheet.setCell(cellAddressToA1(targetCell), sourceValue);
        }

        if (sourceStyle && typeof sourceStyle === "object") {
          worksheet.setCellStyleAt(targetCell.row, targetCell.col, sourceStyle);
        }

        const after = captureCellMutationState(targetCell);
        if (!after) {
          continue;
        }
        mutations.push({
          after,
          before,
          cell: targetCell
        });
      }
    }

    maybeRecalculateWorkbook(workbook.value);
    refreshWorkbookState(workbook.value);
    selection.value = nextRange;
    activeCell.value = nextRange.end;
    selectionAnchorRef.value = nextRange.start;
    recordRangeEditHistory(mutations, nextRange, nextRange.end);
  }

  function mergeSelection() {
    const worksheet = getActiveWorksheet();
    if (readOnly.value || !worksheet || !selection.value || !workbook.value) {
      return;
    }

    recordHistoryBeforeMutation();
    worksheet.mergeCells(rangeToA1(selection.value));
    refreshWorkbookState(workbook.value);
  }

  function unmergeSelection() {
    const worksheet = getActiveWorksheet();
    if (readOnly.value || !worksheet || !selection.value || !workbook.value) {
      return;
    }

    recordHistoryBeforeMutation();
    worksheet.unmergeCells(rangeToA1(selection.value));
    refreshWorkbookState(workbook.value);
  }

  function addSheet(name?: string) {
    if (readOnly.value || !workbook.value) {
      return;
    }

    recordHistoryBeforeMutation();
    const baseName = name?.trim() || "Sheet";
    let candidate = baseName;
    let counter = 2;
    while (workbook.value.sheetIndex(candidate) !== undefined) {
      candidate = `${baseName} ${counter}`;
      counter += 1;
    }

    workbook.value.addSheet(candidate);
    sheetOriginsRef.value = [...sheetOriginsRef.value, null];
    imagesByWorkbookSheetIndex.value = [...imagesByWorkbookSheetIndex.value, []];
    shapesByWorkbookSheetIndex.value = [...shapesByWorkbookSheetIndex.value, []];
    const nextSheets = buildSheetList(
      workbook.value,
      imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex,
      imageAssetsRef.value?.themePalette,
      imageAssetsRef.value?.styleById,
      imageAssetsRef.value?.namedCellStyleByName,
      imageAssetsRef.value?.tableStyleByName,
      showHiddenSheets
    );
    sheets.value = nextSheets;
    const nextChartAssets = imageAssetsRef.value
      ? loadWorkbookChartAssets(workbook.value, imageAssetsRef.value, buildVisibleSheetIndexMap(nextSheets), showHiddenSheets)
      : null;
    if (imageAssetsRef.value) {
      setChartAssets(nextChartAssets);
    }
    const nextIndex = nextSheets.findIndex((sheet) => sheet.name === candidate);
    activeSheetIndex.value = nextIndex >= 0 ? nextIndex : 0;
    const nextTabIndex = nextChartAssets?.tabs.findIndex((tab) => tab.kind === "sheet" && tab.name === candidate) ?? -1;
    if (nextTabIndex >= 0) {
      activeTabIndex.value = nextTabIndex;
    }
    revision.value = revision.value + 1;
  }

  function removeActiveSheet() {
    const activeSheetData = activeSheet.value;
    if (readOnly.value || !workbook.value || !activeSheetData) {
      return;
    }

    recordHistoryBeforeMutation();
    workbook.value.removeSheet(activeSheetData.workbookSheetIndex);
    sheetOriginsRef.value = sheetOriginsRef.value.filter((_, index) => index !== activeSheetData.workbookSheetIndex);
    imagesByWorkbookSheetIndex.value = imagesByWorkbookSheetIndex.value.filter((_, index) => index !== activeSheetData.workbookSheetIndex);
    shapesByWorkbookSheetIndex.value = shapesByWorkbookSheetIndex.value.filter((_, index) => index !== activeSheetData.workbookSheetIndex);
    if (imageAssetsRef.value) {
      imageAssetsRef.value.sheetStatesByWorkbookSheetIndex = imageAssetsRef.value.sheetStatesByWorkbookSheetIndex.filter(
        (_, index) => index !== activeSheetData.workbookSheetIndex
      );
    }
    const nextSheets = buildSheetList(
      workbook.value,
      imageAssetsRef.value?.sheetStatesByWorkbookSheetIndex,
      imageAssetsRef.value?.themePalette,
      imageAssetsRef.value?.styleById,
      imageAssetsRef.value?.namedCellStyleByName,
      imageAssetsRef.value?.tableStyleByName,
      showHiddenSheets
    );
    sheets.value = nextSheets;
    if (imageAssetsRef.value) {
      setChartAssets(loadWorkbookChartAssets(workbook.value, imageAssetsRef.value, buildVisibleSheetIndexMap(nextSheets), showHiddenSheets));
    }
    activeSheetIndex.value = Math.max(0, Math.min(activeSheetIndex.value, nextSheets.length - 1));
    revision.value = revision.value + 1;
  }

  function defineNamedRange(name: string, range?: XlsxCellRange | null) {
    if (readOnly.value || !workbook.value) {
      return;
    }

    const trimmed = name.trim();
    if (!trimmed) {
      return;
    }

    const targetRange = range ?? selection.value;
    if (!targetRange) {
      return;
    }

    recordHistoryBeforeMutation();
    workbook.value.defineName(trimmed, rangeToA1(targetRange));
    revision.value = revision.value + 1;
  }

  function pasteText(text: string) {
    const worksheet = getActiveWorksheet();
    const targetCell = activeCell.value ?? selection.value?.start ?? null;
    if (readOnly.value || !worksheet || !workbook.value || !targetCell || !text) {
      return false;
    }

    const grid = parseClipboardText(text);
    if (grid.length === 0 || grid.every((row) => row.length === 0)) {
      return false;
    }

    const mutations: RangeCellMutation[] = [];
    for (let rowIndex = 0; rowIndex < grid.length; rowIndex += 1) {
      const row = grid[rowIndex] ?? [];
      for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
        const rawValue = row[colIndex] ?? "";
        const nextCell = {
          col: targetCell.col + colIndex,
          row: targetCell.row + rowIndex
        };
        const before = captureCellMutationState(nextCell);
        if (!before) {
          continue;
        }
        if (rawValue.startsWith("=") && rawValue.length > 1) {
          worksheet.setFormula(cellAddressToA1(nextCell), rawValue);
          const after = captureCellMutationState(nextCell);
          if (!after) {
            continue;
          }
          mutations.push({
            after,
            before,
            cell: nextCell
          });
        } else {
          const nextValue = coerceUserEnteredValue(rawValue);
          worksheet.setCell(cellAddressToA1(nextCell), nextValue);
          const after = captureCellMutationState(nextCell);
          if (!after) {
            continue;
          }
          mutations.push({
            after,
            before,
            cell: nextCell
          });
        }
      }
    }

    maybeRecalculateWorkbook(workbook.value);
    refreshWorkbookState(workbook.value);
    const nextRange = normalizeRange({
      start: targetCell,
      end: {
        col: targetCell.col + Math.max(0, Math.max(...grid.map((row) => row.length), 1) - 1),
        row: targetCell.row + grid.length - 1
      }
    });
    activeCell.value = targetCell;
    selection.value = nextRange;
    selectionAnchorRef.value = targetCell;
    recordRangeEditHistory(mutations, nextRange, targetCell);
    return true;
  }

  function pasteStructuredClipboardData(serializedPayload: string) {
    const worksheet = getActiveWorksheet();
    const targetCell = activeCell.value ?? selection.value?.start ?? null;
    if (readOnly.value || !worksheet || !workbook.value || !targetCell || !serializedPayload) {
      return false;
    }

    let payload: ClipboardPayload;
    try {
      payload = JSON.parse(serializedPayload) as ClipboardPayload;
    } catch {
      return false;
    }

    if (!Array.isArray(payload.cells) || payload.cells.length === 0) {
      return false;
    }

    const hasMergeOperations = Array.isArray(payload.merges) && payload.merges.some((merge) => (merge.rowSpan ?? 1) > 1 || (merge.colSpan ?? 1) > 1);
    const mutations: RangeCellMutation[] = [];
    if (hasMergeOperations) {
      recordHistoryBeforeMutation();
    }
    for (const cell of payload.cells) {
      const nextCell = {
        col: targetCell.col + cell.colOffset,
        row: targetCell.row + cell.rowOffset
      };
      const before = hasMergeOperations ? null : captureCellMutationState(nextCell);

      if (cell.formula) {
        worksheet.setFormula(cellAddressToA1(nextCell), cell.formula);
        if (before) {
          const after = captureCellMutationState(nextCell);
          if (!after) {
            continue;
          }
          mutations.push({
            after,
            before,
            cell: nextCell
          });
        }
      } else {
        worksheet.setCell(cellAddressToA1(nextCell), cell.value);
        if (before) {
          const after = captureCellMutationState(nextCell);
          if (!after) {
            continue;
          }
          mutations.push({
            after,
            before,
            cell: nextCell
          });
        }
      }
    }

    if (Array.isArray(payload.merges)) {
      for (const merge of payload.merges) {
        if ((merge.rowSpan ?? 1) <= 1 && (merge.colSpan ?? 1) <= 1) {
          continue;
        }

        const mergeRange = normalizeRange({
          start: {
            col: targetCell.col + merge.colOffset,
            row: targetCell.row + merge.rowOffset
          },
          end: {
            col: targetCell.col + merge.colOffset + merge.colSpan - 1,
            row: targetCell.row + merge.rowOffset + merge.rowSpan - 1
          }
        });
        worksheet.mergeCells(rangeToA1(mergeRange));
      }
    }

    maybeRecalculateWorkbook(workbook.value);
    refreshWorkbookState(workbook.value);
    const nextRange = normalizeRange({
      start: targetCell,
      end: {
        col: targetCell.col + Math.max((payload.cols ?? 1) - 1, 0),
        row: targetCell.row + Math.max((payload.rows ?? 1) - 1, 0)
      }
    });
    activeCell.value = targetCell;
    selection.value = nextRange;
    selectionAnchorRef.value = targetCell;
    if (!hasMergeOperations) {
      recordRangeEditHistory(mutations, nextRange, targetCell);
    }
    return true;
  }

  async function copySelectionToClipboard() {
    const clipboardData = getClipboardData();
    if (!clipboardData || typeof navigator === "undefined" || !navigator.clipboard) {
      return false;
    }

    if (typeof ClipboardItem === "function" && navigator.clipboard.write) {
      const item = new ClipboardItem({
        [INTERNAL_CLIPBOARD_MIME]: new Blob([clipboardData.structured], { type: INTERNAL_CLIPBOARD_MIME }),
        "text/html": new Blob([clipboardData.html], { type: "text/html" }),
        "text/plain": new Blob([clipboardData.text], { type: "text/plain" })
      });
      await navigator.clipboard.write([item]);
      return true;
    }

    await navigator.clipboard.writeText(clipboardData.text);
    return true;
  }

  async function pasteFromClipboard() {
    if (readOnly.value || typeof navigator === "undefined" || !navigator.clipboard) {
      return false;
    }

    if (navigator.clipboard.read) {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (item.types.includes(INTERNAL_CLIPBOARD_MIME)) {
          const blob = await item.getType(INTERNAL_CLIPBOARD_MIME);
          return pasteStructuredClipboardData(await blob.text());
        }
      }

      for (const item of items) {
        if (item.types.includes("text/plain")) {
          const blob = await item.getType("text/plain");
          return pasteText(await blob.text());
        }
      }
    }

    return pasteText(await navigator.clipboard.readText());
  }

  function undo() {
    if (readOnly.value || !workbook.value || undoStackRef.value.length === 0) {
      return;
    }

    const entry = undoStackRef.value.pop();
    if (!entry) {
      return;
    }

    if (entry.kind === "cell-edit") {
      pushHistoryEntry(redoStackRef.value, entry);
      historyRevision.value = historyRevision.value + 1;
      applyCellEditHistoryEntry(entry, "undo");
      return;
    }

    if (entry.kind === "range-edit") {
      pushHistoryEntry(redoStackRef.value, entry);
      historyRevision.value = historyRevision.value + 1;
      applyRangeEditHistoryEntry(entry, "undo");
      return;
    }

    const currentSnapshot = createHistoryEntry();
    if (currentSnapshot) {
      pushHistoryEntry(redoStackRef.value, currentSnapshot);
    }
    historyRevision.value = historyRevision.value + 1;
    void restoreHistoryEntry(entry);
  }

  function redo() {
    if (readOnly.value || !workbook.value || redoStackRef.value.length === 0) {
      return;
    }

    const entry = redoStackRef.value.pop();
    if (!entry) {
      return;
    }

    if (entry.kind === "cell-edit") {
      pushHistoryEntry(undoStackRef.value, entry);
      historyRevision.value = historyRevision.value + 1;
      applyCellEditHistoryEntry(entry, "redo");
      return;
    }

    if (entry.kind === "range-edit") {
      pushHistoryEntry(undoStackRef.value, entry);
      historyRevision.value = historyRevision.value + 1;
      applyRangeEditHistoryEntry(entry, "redo");
      return;
    }

    const currentSnapshot = createHistoryEntry();
    if (currentSnapshot) {
      pushHistoryEntry(undoStackRef.value, currentSnapshot);
    }
    historyRevision.value = historyRevision.value + 1;
    void restoreHistoryEntry(entry);
  }

  return {
      get activeCell() { return activeCell.value; },
      get activeCellAddress() { return activeCellAddress.value; },
      get activeSheet() { return activeSheet.value; },
      get activeSheetIndex() { return activeSheetIndex.value; },
      get activeTab() { return activeTab.value; },
      get activeTabIndex() { return activeTabIndex.value; },
      addSheet,
      get canRedo() { return canRedo.value; },
      get canDownload() { return Boolean(file ?? src); },
      get canExport() { return Boolean(workbook.value); },
      get canLoadDeferred() { return canLoadDeferred.value; },
      get canUndo() { return canUndo.value; },
      get canZoomIn() { return canZoomIn.value; },
      get canZoomOut() { return canZoomOut.value; },
      get charts() { return charts.value; },
      get chartsheets() { return chartsheets.value; },
      clearSelectedChart,
      clearSelectedChartElement,
      clearSelectedCells,
      clearSelectedImage,
      clearSelection,
      continueDeferredLoad,
      copySelectionToClipboard,
      get defaultZoomScale() { return defaultZoomScale.value; },
      get deferredLoadFileSize() { return deferredLoadFileSize.value; },
      defineNamedRange,
      get displayFileName() { return displayFileName.value; },
      download,
      exportCsv,
      exportXlsx,
      get error() { return error.value; },
      fillSelection,
      get formControls() { return formControls.value; },
      getChartById,
      getChartSeriesFormula,
      getChartsheetById,
      getImageById,
      getSheetCharts,
      getSheetFormControls,
      getSheetImages,
      getSheetShapes,
      file,
      getClipboardData,
      getCellDisplayValue,
      getCellFormula,
      get getCellSnapshotAsync() { return isWorkerBacked.value ? getCellSnapshotAsync : undefined; },
      getActiveWorksheet,
      get getRowsBatchAsync() { return isWorkerBacked.value ? getRowsBatchAsync : undefined; },
      get images() { return images.value; },
      get isLoadDeferred() { return isLoadDeferred.value; },
      get isLoading() { return isLoading.value; },
      get isChartsLoading() { return isChartsLoading.value; },
      get isWorkerBacked() { return isWorkerBacked.value; },
      mergeSelection,
      maxZoomScale: MAX_ZOOM_SCALE,
      minZoomScale: MIN_ZOOM_SCALE,
      moveChartBy,
      moveImageBy,
      pasteFromClipboard,
      pasteStructuredClipboardData,
      pasteText,
      removeActiveSheet,
      get readOnly() { return readOnly.value; },
      recalculate,
      redo,
      resetZoom,
      get revision() { return revision.value; },
      resizeChartBy,
      resizeImageBy,
      resizeColumn,
      resizeRow,
      setCellFormula,
      setCellStyle,
      setCellValue,
      setRangeStyle,
      setSelectedFormula,
      setZoomScale,
      setChartRect,
      setChartSeriesFormula,
      setImageRect,
      get selectedChart() { return selectedChart.value; },
      get selectedChartElement() { return selectedChartElement.value; },
      get selectedChartFormula() { return selectedChartFormula.value; },
      get selectedChartId() { return selectedChartId.value; },
      get selectedCellFormula() { return selectedCellFormula.value; },
      get selectedFormula() { return selectedFormula.value; },
      get selectedFormulaTarget() { return selectedFormulaTarget.value; },
      get selectedImage() { return selectedImage.value; },
      get selectedImageId() { return selectedImageId.value; },
      get selectedRangeAddress() { return selectedRangeAddress.value; },
      get selectedValue() { return selectedValue.value; },
      selectCell,
      selectChart,
      selectChartElement,
      selectImage,
      selectRange,
      get selection() { return selection.value; },
      setActiveSheetIndex,
      setActiveTabIndex,
      setSelectedCellFormula,
      setSelectedCellStyle,
      setSelectedCellValue,
      get sheets() { return sheets.value; },
      get shapes() { return shapes.value; },
      src,
      get sortState() { return sortState.value; },
      sortTable,
      get tabs() { return tabs.value; },
      get tables() { return tables.value; },
      undo,
      unmergeSelection,
      updateChart,
      get workbook() { return workbook.value; },
      zoomIn,
      zoomOut,
      get zoomScale() { return zoomScale.value; }
    };
}
