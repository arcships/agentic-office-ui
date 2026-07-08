import type {
  XlsxCellRange,
  XlsxImage,
  XlsxImageRect,
  XlsxImageResizeHandlePosition,
} from "../types";
import {
  EMU_PER_PIXEL,
  type DukeWorksheet,
  type DrawingRectEmu,
  emuToPixels,
  parseA1RangeReference,
  pixelsToEmu,
} from "./image-parser";
import type { WorkbookSheetState } from "./grid-render";

// ── Constants ───────────────────────────────────────────────────────────────

export const MIN_COL_WIDTH_PX = 30;
export const MIN_ROW_HEIGHT_PX = 16;
const DEFAULT_COL_WIDTH_EMU = 64 * EMU_PER_PIXEL;
const DEFAULT_ROW_HEIGHT_EMU = 20 * EMU_PER_PIXEL;
const DEFAULT_COLUMN_CHARACTER_WIDTH_PX = 7;
const columnCharacterWidthCache = new Map<string, number>();

// ── Gridline thickness ─────────────────────────────────────────────────────

export function resolveDeviceGridlineThicknessPx() {
  if (typeof window === "undefined") {
    return 1;
  }

  const devicePixelRatio = window.devicePixelRatio;
  if (!Number.isFinite(devicePixelRatio) || devicePixelRatio <= 0) {
    return 1;
  }

  return 1 / devicePixelRatio;
}

// ── Column character width measurement ─────────────────────────────────────

export function measureColumnCharacterWidthPx(fontFamily?: string | null, fontSizePt?: number | null) {
  const normalizedFamily = typeof fontFamily === "string" && fontFamily.trim().length > 0
    ? fontFamily.trim()
    : "Calibri";
  const normalizedSizePt = typeof fontSizePt === "number" && Number.isFinite(fontSizePt) && fontSizePt > 0
    ? fontSizePt
    : 11;
  const cacheKey = `${normalizedFamily}|${normalizedSizePt}`;
  const cached = columnCharacterWidthCache.get(cacheKey);
  if (cached !== undefined) {
    return cached;
  }

  const fontSizePx = normalizedSizePt * (96 / 72);
  const font = `${fontSizePx}px "${normalizedFamily}"`;
  let width = DEFAULT_COLUMN_CHARACTER_WIDTH_PX;

  try {
    const context = typeof document !== "undefined"
      ? document.createElement("canvas").getContext("2d")
      : typeof OffscreenCanvas !== "undefined"
        ? new OffscreenCanvas(32, 32).getContext("2d")
        : null;
    if (context) {
      context.font = font;
      width = Math.max(1, context.measureText("0").width);
    }
  } catch {
    width = DEFAULT_COLUMN_CHARACTER_WIDTH_PX;
  }

  columnCharacterWidthCache.set(cacheKey, width);
  return width;
}

// ── Column/row sizing ──────────────────────────────────────────────────────

export function sheetColumnWidthToPixels(width: number, columnCharacterWidthPx = DEFAULT_COLUMN_CHARACTER_WIDTH_PX) {
  if (!Number.isFinite(width) || width <= 0) {
    return MIN_COL_WIDTH_PX;
  }

  const digitWidth = Math.max(1, columnCharacterWidthPx);
  const pixels = width < 1
    ? Math.floor(width * (digitWidth + 5) + 0.5)
    : Math.floor(((256 * width + Math.floor(128 / digitWidth)) / 256) * digitWidth);
  return Math.max(MIN_COL_WIDTH_PX, pixels);
}

export function resolveWorksheetDefaultColumnWidthPixels(
  worksheet: DukeWorksheet,
  columnCharacterWidthPx = DEFAULT_COLUMN_CHARACTER_WIDTH_PX,
  fallbackPx = sheetColumnWidthToPixels(8.43, columnCharacterWidthPx)
) {
  const width = typeof worksheet.defaultColumnWidth === "number" ? worksheet.defaultColumnWidth : Number.NaN;
  return Number.isFinite(width) && width > 0
    ? sheetColumnWidthToPixels(width, columnCharacterWidthPx)
    : fallbackPx;
}

export function resolveWorksheetDefaultRowHeightPixels(
  worksheet: DukeWorksheet,
  fallbackPx = Math.max(MIN_ROW_HEIGHT_PX, Math.round(15 * 1.33))
) {
  const height = typeof worksheet.defaultRowHeight === "number" ? worksheet.defaultRowHeight : Number.NaN;
  return Number.isFinite(height) && height > 0
    ? Math.max(MIN_ROW_HEIGHT_PX, Math.round(height * 1.33))
    : fallbackPx;
}

// ── Hidden rows/cols ───────────────────────────────────────────────────────

export function resolveWorksheetHiddenRows(worksheet: DukeWorksheet, maxRow: number) {
  if (!Number.isFinite(maxRow) || maxRow < 0 || typeof worksheet.isRowHidden !== "function") {
    return [] as number[];
  }

  const hiddenRows: number[] = [];
  for (let row = 0; row <= maxRow; row += 1) {
    if (worksheet.isRowHidden(row)) {
      hiddenRows.push(row);
    }
  }
  return hiddenRows;
}

export function resolveWorksheetHiddenCols(worksheet: DukeWorksheet, maxCol: number) {
  if (!Number.isFinite(maxCol) || maxCol < 0 || typeof worksheet.isColumnHidden !== "function") {
    return [] as number[];
  }

  const hiddenCols: number[] = [];
  for (let col = 0; col <= maxCol; col += 1) {
    if (worksheet.isColumnHidden(col)) {
      hiddenCols.push(col);
    }
  }
  return hiddenCols;
}

// ── Merge metadata ─────────────────────────────────────────────────────────

export function resolveWorksheetMergeMetadata(worksheet: DukeWorksheet) {
  const mergeMetadata = {
    hasHorizontalMerges: false,
    hasVerticalMerges: false,
    maxHorizontalMergeEndCol: -1,
    maxVerticalMergeEndRow: -1
  };
  const mergedRegions = Array.isArray(worksheet.mergedRegions) ? worksheet.mergedRegions : [];

  for (const rawRegion of mergedRegions) {
    let range: XlsxCellRange | null = null;
    if (typeof rawRegion === "string") {
      range = parseA1RangeReference(rawRegion);
    } else if (rawRegion && typeof rawRegion === "object") {
      const region = rawRegion as Record<string, unknown>;
      const startRow = typeof region.startRow === "number" ? region.startRow : Number.NaN;
      const startCol = typeof region.startCol === "number" ? region.startCol : Number.NaN;
      const endRow = typeof region.endRow === "number" ? region.endRow : Number.NaN;
      const endCol = typeof region.endCol === "number" ? region.endCol : Number.NaN;
      if ([startRow, startCol, endRow, endCol].every((value) => Number.isFinite(value) && value >= 0)) {
        range = {
          end: {
            col: Math.max(startCol, endCol),
            row: Math.max(startRow, endRow)
          },
          start: {
            col: Math.min(startCol, endCol),
            row: Math.min(startRow, endRow)
          }
        };
      } else if (typeof region.range === "string") {
        range = parseA1RangeReference(region.range);
      }
    }

    if (!range) {
      continue;
    }

    if (range.end.col > range.start.col) {
      mergeMetadata.hasHorizontalMerges = true;
      mergeMetadata.maxHorizontalMergeEndCol = Math.max(mergeMetadata.maxHorizontalMergeEndCol, range.end.col);
    }
    if (range.end.row > range.start.row) {
      mergeMetadata.hasVerticalMerges = true;
      mergeMetadata.maxVerticalMergeEndRow = Math.max(mergeMetadata.maxVerticalMergeEndRow, range.end.row);
    }
  }

  return mergeMetadata;
}

// ── Sheet-level sizing ─────────────────────────────────────────────────────

export function resolveSheetColumnWidthPx(sheetState: WorkbookSheetState | null, col: number) {
  if (col < 0) {
    return 0;
  }
  return sheetState?.colWidthOverridesPx[col] ?? sheetState?.defaultColWidthPx ?? emuToPixels(DEFAULT_COL_WIDTH_EMU);
}

export function resolveSheetRowHeightPx(sheetState: WorkbookSheetState | null, row: number) {
  if (row < 0) {
    return 0;
  }
  return sheetState?.rowHeightOverridesPx[row] ?? sheetState?.defaultRowHeightPx ?? emuToPixels(DEFAULT_ROW_HEIGHT_EMU);
}

export function sumSheetColumnWidthsEmu(sheetState: WorkbookSheetState | null, beforeCol: number) {
  let total = 0;
  for (let col = 0; col < beforeCol; col += 1) {
    total += pixelsToEmu(resolveSheetColumnWidthPx(sheetState, col));
  }
  return total;
}

export function sumSheetRowHeightsEmu(sheetState: WorkbookSheetState | null, beforeRow: number) {
  let total = 0;
  for (let row = 0; row < beforeRow; row += 1) {
    total += pixelsToEmu(resolveSheetRowHeightPx(sheetState, row));
  }
  return total;
}

// ── Anchor → absolute rect ─────────────────────────────────────────────────

export function anchorToAbsoluteRect(anchor: XlsxImage["anchor"], sheetState: WorkbookSheetState | null): DrawingRectEmu {
  if (anchor.kind === "absolute") {
    return {
      cx: anchor.sizeEmu.cx,
      cy: anchor.sizeEmu.cy,
      x: anchor.positionEmu.x,
      y: anchor.positionEmu.y
    };
  }

  if (anchor.kind === "one-cell") {
    return {
      cx: anchor.sizeEmu.cx,
      cy: anchor.sizeEmu.cy,
      x: sumSheetColumnWidthsEmu(sheetState, anchor.from.col) + anchor.from.colOffsetEmu,
      y: sumSheetRowHeightsEmu(sheetState, anchor.from.row) + anchor.from.rowOffsetEmu
    };
  }

  const left = sumSheetColumnWidthsEmu(sheetState, anchor.from.col) + anchor.from.colOffsetEmu;
  const top = sumSheetRowHeightsEmu(sheetState, anchor.from.row) + anchor.from.rowOffsetEmu;
  const right = sumSheetColumnWidthsEmu(sheetState, anchor.to.col) + anchor.to.colOffsetEmu;
  const bottom = sumSheetRowHeightsEmu(sheetState, anchor.to.row) + anchor.to.rowOffsetEmu;

  return {
    cx: Math.max(0, right - left),
    cy: Math.max(0, bottom - top),
    x: left,
    y: top
  };
}

// ── Public column width / row height helpers ───────────────────────────────

export function pxToSheetColumnWidth(widthPx: number) {
  return (Math.max(widthPx, MIN_COL_WIDTH_PX) - 5) / 7;
}

export function resolveSheetColumnWidthPixels(width: number, columnWidthCharacterWidthPx?: number) {
  return sheetColumnWidthToPixels(width, columnWidthCharacterWidthPx);
}

export function resolveSheetRowHeightPixels(height: number) {
  return Math.max(Math.round(height * 1.33), MIN_ROW_HEIGHT_PX);
}

export function resolveRenderedSheetAxisPixels(sizePx: number, showGridLines = true) {
  return Math.max(0, sizePx) + (showGridLines ? resolveDeviceGridlineThicknessPx() : 0);
}

export function resolveContentSheetAxisPixels(sizePx: number, showGridLines = true) {
  return Math.max(0, sizePx - (showGridLines ? resolveDeviceGridlineThicknessPx() : 0));
}

// ── Marker from offset ─────────────────────────────────────────────────────

function markerFromOffset(offsetPx: number, getSizePx: (index: number) => number) {
  let remaining = Math.max(0, offsetPx);
  let index = 0;
  while (remaining > 0) {
    const size = Math.max(1, getSizePx(index));
    if (remaining < size) {
      break;
    }
    remaining -= size;
    index += 1;
  }

  return {
    index,
    offsetPx: remaining
  };
}

// ── Rect → image anchor ────────────────────────────────────────────────────

export function rectToImageAnchor(
  rect: XlsxImageRect,
  currentAnchor: XlsxImage["anchor"],
  options: {
    contentOffsetLeft: number;
    contentOffsetTop: number;
    getColumnWidthPx: (col: number) => number;
    getRowHeightPx: (row: number) => number;
  }
): XlsxImage["anchor"] {
  const contentLeft = Math.max(0, rect.left - options.contentOffsetLeft);
  const contentTop = Math.max(0, rect.top - options.contentOffsetTop);
  const contentRight = Math.max(contentLeft + 1, rect.left + rect.width - options.contentOffsetLeft);
  const contentBottom = Math.max(contentTop + 1, rect.top + rect.height - options.contentOffsetTop);

  if (currentAnchor.kind === "absolute") {
    return {
      kind: "absolute",
      positionEmu: {
        x: pixelsToEmu(contentLeft),
        y: pixelsToEmu(contentTop)
      },
      sizeEmu: {
        cx: pixelsToEmu(rect.width),
        cy: pixelsToEmu(rect.height)
      }
    };
  }

  const fromCol = markerFromOffset(contentLeft, options.getColumnWidthPx);
  const fromRow = markerFromOffset(contentTop, options.getRowHeightPx);
  if (currentAnchor.kind === "one-cell") {
    return {
      from: {
        col: fromCol.index,
        colOffsetEmu: pixelsToEmu(fromCol.offsetPx),
        row: fromRow.index,
        rowOffsetEmu: pixelsToEmu(fromRow.offsetPx)
      },
      kind: "one-cell",
      sizeEmu: {
        cx: pixelsToEmu(rect.width),
        cy: pixelsToEmu(rect.height)
      }
    };
  }

  const toCol = markerFromOffset(contentRight, options.getColumnWidthPx);
  const toRow = markerFromOffset(contentBottom, options.getRowHeightPx);
  return {
    from: {
      col: fromCol.index,
      colOffsetEmu: pixelsToEmu(fromCol.offsetPx),
      row: fromRow.index,
      rowOffsetEmu: pixelsToEmu(fromRow.offsetPx)
    },
    kind: "two-cell",
    to: {
      col: toCol.index,
      colOffsetEmu: pixelsToEmu(toCol.offsetPx),
      row: toRow.index,
      rowOffsetEmu: pixelsToEmu(toRow.offsetPx)
    }
  };
}

// ── Image resize ───────────────────────────────────────────────────────────

export function resizeImageRect(
  rect: XlsxImageRect,
  handle: XlsxImageResizeHandlePosition,
  deltaX: number,
  deltaY: number,
  minimumSize = 16
) {
  let left = rect.left;
  let top = rect.top;
  let width = rect.width;
  let height = rect.height;

  if (handle.includes("w")) {
    left += deltaX;
    width -= deltaX;
  }
  if (handle.includes("e")) {
    width += deltaX;
  }
  if (handle.includes("n")) {
    top += deltaY;
    height -= deltaY;
  }
  if (handle.includes("s")) {
    height += deltaY;
  }

  if (width < minimumSize) {
    if (handle.includes("w")) {
      left -= minimumSize - width;
    }
    width = minimumSize;
  }
  if (height < minimumSize) {
    if (handle.includes("n")) {
      top -= minimumSize - height;
    }
    height = minimumSize;
  }

  return { height, left, top, width };
}
