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

type SheetAxisSizeState = {
  colWidthOverridesPx?: Record<number, number>;
  defaultColWidthPx?: number;
  hiddenCols?: number[];
  hiddenRows?: number[];
  rowHeightOverridesPx?: Record<number, number>;
  defaultRowHeightPx?: number;
  /** Display-indexed widths after hidden columns are removed (grid paint order). */
  colWidths?: number[];
  /** Display-indexed heights after hidden rows are removed (grid paint order). */
  rowHeights?: number[];
  /** Workbook column indices that the grid actually paints, in paint order. */
  visibleCols?: number[];
  /** Workbook row indices that the grid actually paints, in paint order. */
  visibleRows?: number[];
} | null;

// ── Constants ───────────────────────────────────────────────────────────────

export const MIN_COL_WIDTH_PX = 30;
export const MIN_ROW_HEIGHT_PX = 16;
const DEFAULT_COL_WIDTH_EMU = 64 * EMU_PER_PIXEL;
const DEFAULT_ROW_HEIGHT_EMU = 20 * EMU_PER_PIXEL;
const DEFAULT_COLUMN_CHARACTER_WIDTH_PX = 7;

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

/** Exact merged regions reported by the engine, normalized to zero-based ranges. */
export function resolveWorksheetMergedRegions(worksheet: DukeWorksheet): XlsxCellRange[] {
  const mergedRegions = Array.isArray(worksheet.mergedRegions) ? worksheet.mergedRegions : [];
  return mergedRegions.flatMap((rawRegion) => {
    if (typeof rawRegion === "string") {
      const range = parseA1RangeReference(rawRegion);
      return range ? [range] : [];
    }
    if (!rawRegion || typeof rawRegion !== "object") {
      return [];
    }
    const region = rawRegion as Record<string, unknown>;
    const startRow = typeof region.startRow === "number" ? region.startRow : Number.NaN;
    const startCol = typeof region.startCol === "number" ? region.startCol : Number.NaN;
    const endRow = typeof region.endRow === "number" ? region.endRow : Number.NaN;
    const endCol = typeof region.endCol === "number" ? region.endCol : Number.NaN;
    if ([startRow, startCol, endRow, endCol].every((value) => Number.isFinite(value) && value >= 0)) {
      return [{
        end: {
          col: Math.max(startCol, endCol),
          row: Math.max(startRow, endRow)
        },
        start: {
          col: Math.min(startCol, endCol),
          row: Math.min(startRow, endRow)
        }
      }];
    }
    if (typeof region.range === "string") {
      const range = parseA1RangeReference(region.range);
      return range ? [range] : [];
    }
    return [];
  });
}

export function resolveWorksheetMergeMetadata(worksheet: DukeWorksheet) {
  const mergeMetadata = {
    hasHorizontalMerges: false,
    hasVerticalMerges: false,
    maxHorizontalMergeEndCol: -1,
    maxVerticalMergeEndRow: -1
  };

  for (const range of resolveWorksheetMergedRegions(worksheet)) {
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

export function resolveSheetColumnWidthPx(sheetState: SheetAxisSizeState | null, col: number) {
  if (col < 0) {
    return 0;
  }
  const fallback = sheetState?.defaultColWidthPx ?? emuToPixels(DEFAULT_COL_WIDTH_EMU);
  const { colWidths, visibleCols } = sheetState ?? {};
  if (visibleCols?.length && colWidths?.length) {
    const displayIndex = visibleCols.indexOf(col);
    if (displayIndex >= 0) {
      return colWidths[displayIndex] ?? fallback;
    }
    // Columns past the tracked visible range still use the default width;
    // indices inside the range that are absent are hidden and take no space.
    return col > visibleCols[visibleCols.length - 1]! && !sheetState?.hiddenCols?.includes(col)
      ? fallback
      : 0;
  }
  return sheetState?.colWidthOverridesPx?.[col] ?? fallback;
}

export function resolveSheetRowHeightPx(sheetState: SheetAxisSizeState | null, row: number) {
  if (row < 0) {
    return 0;
  }
  const fallback = sheetState?.defaultRowHeightPx ?? emuToPixels(DEFAULT_ROW_HEIGHT_EMU);
  const { rowHeights, visibleRows } = sheetState ?? {};
  if (visibleRows?.length && rowHeights?.length) {
    const displayIndex = visibleRows.indexOf(row);
    if (displayIndex >= 0) {
      return rowHeights[displayIndex] ?? fallback;
    }
    return row > visibleRows[visibleRows.length - 1]! && !sheetState?.hiddenRows?.includes(row)
      ? fallback
      : 0;
  }
  return sheetState?.rowHeightOverridesPx?.[row] ?? fallback;
}

export function sumSheetColumnWidthsEmu(sheetState: SheetAxisSizeState, beforeCol: number) {
  return sumSheetAxisSizesEmu(sheetState, beforeCol, "column");
}

export function sumSheetRowHeightsEmu(sheetState: SheetAxisSizeState, beforeRow: number) {
  return sumSheetAxisSizesEmu(sheetState, beforeRow, "row");
}

/**
 * Sum the rendered size of every axis slot before `beforeIndex`.
 *
 * Anchors are stored with workbook indices, but the grid removes hidden rows /
 * columns from layout and paints `colWidths` / `rowHeights` in display order.
 * Walking workbook indices and skipping hidden slots keeps anchored objects
 * aligned with the cells the canvas actually draws, including sheets whose
 * hidden columns have a stored width (Excel keeps `ht` / `width` on hidden
 * rows / columns).
 */
function sumSheetAxisSizesEmu(
  sheetState: SheetAxisSizeState,
  beforeIndex: number,
  axis: "column" | "row",
) {
  if (!Number.isFinite(beforeIndex) || beforeIndex <= 0) {
    return 0;
  }

  const hidden = new Set(axis === "column" ? sheetState?.hiddenCols : sheetState?.hiddenRows);
  const visible = axis === "column" ? sheetState?.visibleCols : sheetState?.visibleRows;
  const sizes = axis === "column" ? sheetState?.colWidths : sheetState?.rowHeights;
  const sortedVisible = visible && visible.length > 0 ? visible : null;
  const lastVisible = sortedVisible ? sortedVisible[sortedVisible.length - 1] ?? -1 : -1;
  let total = 0;

  for (let index = 0; index < beforeIndex; index += 1) {
    if (hidden.has(index)) {
      continue;
    }

    let sizePx: number | undefined;
    if (sortedVisible) {
      const displayIndex = findSortedIndex(sortedVisible, index);
      if (displayIndex >= 0) {
        const candidate = sizes?.[displayIndex];
        if (typeof candidate === "number" && candidate > 0) {
          sizePx = candidate;
        }
      } else if (index <= lastVisible) {
        // The grid's visible list covers every index up to the last displayed
        // one; a missing slot in that range is hidden even when the caller did
        // not populate `hiddenCols`.
        continue;
      }
    }

    if (sizePx === undefined) {
      sizePx = axis === "column"
        ? resolveSheetColumnWidthPx(sheetState, index)
        : resolveSheetRowHeightPx(sheetState, index);
    }
    total += pixelsToEmu(sizePx);
  }

  return total;
}

function findSortedIndex(indices: number[], target: number) {
  let low = 0;
  let high = indices.length - 1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    const value = indices[middle];
    if (value === target) {
      return middle;
    }
    if (value < target) {
      low = middle + 1;
    } else {
      high = middle - 1;
    }
  }
  return -1;
}

// ── Anchor → absolute rect ─────────────────────────────────────────────────

export function anchorToAbsoluteRect(
  anchor: XlsxImage["anchor"],
  sheetState: SheetAxisSizeState | null,
): DrawingRectEmu {
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

/**
 * Content-space rect cache for anchored objects. Anchor objects survive scroll
 * ticks, so geometry only needs recomputation when the sheet model or zoom
 * changes. Returns the same rect instance for cache hits.
 */
export function createAnchoredRectCache() {
  const cache = new WeakMap<object, { sheetState: SheetAxisSizeState | null; zoom: number; rect: XlsxImageRect }>();
  return {
    resolve(anchor: XlsxImage["anchor"], sheetState: SheetAxisSizeState | null, zoomScale: number): XlsxImageRect {
      const scaleValue = zoomScale / 100;
      const zoom = Number.isFinite(scaleValue) ? Math.max(0.1, scaleValue) : 1;
      const cached = cache.get(anchor);
      if (cached && cached.sheetState === sheetState && cached.zoom === zoom) {
        return cached.rect;
      }
      const rect = anchorToAbsoluteRect(anchor, sheetState);
      const resolved = {
        height: Math.max(1, emuToPixels(rect.cy) * zoom),
        left: emuToPixels(rect.x) * zoom,
        top: emuToPixels(rect.y) * zoom,
        width: Math.max(1, emuToPixels(rect.cx) * zoom),
      };
      cache.set(anchor, { rect: resolved, sheetState, zoom });
      return resolved;
    },
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
