import { computed, type ComputedRef } from "vue";
import type {
  XlsxSheetData,
  XlsxSheetThumbnail,
  XlsxSheetThumbnailResolution,
  XlsxViewerController,
  XlsxViewerThumbnails,
  XlsxCellRange,
} from "@arcships/xlsx-core";
import { cellAddressToA1 } from "./selection";

export interface UseXlsxViewerThumbnailsOptions {
  includeHeaders?: boolean;
  resolution?: XlsxSheetThumbnailResolution;
}

const THUMBNAIL_MAX_ROWS = 60;
const THUMBNAIL_MAX_COLS = 26;
const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const HEADER_HEIGHT = 24;
const ROW_HEADER_WIDTH = 40;

interface ThumbnailOutputSize {
  height: number;
  scale: number;
  width: number;
}

interface ThumbnailGeometry {
  colWidths: number[];
  headerHeight: number;
  outputSize: ThumbnailOutputSize;
  rowHeaderWidth: number;
  rowHeights: number[];
  sourceHeight: number;
  sourceWidth: number;
  totalColWidth: number;
  totalRowHeight: number;
}

interface ThumbnailCells {
  values: string[][];
  revision: number;
}

function resolveOutputSize(
  sourceWidth: number,
  sourceHeight: number,
  resolution?: UseXlsxViewerThumbnailsOptions["resolution"],
): ThumbnailOutputSize {
  let maxWidth = 200;
  let maxHeight = 132;

  if (typeof resolution === "number") {
    maxWidth = resolution;
    maxHeight = resolution;
  } else if (resolution && typeof resolution === "object") {
    if (resolution.maxWidth) maxWidth = resolution.maxWidth;
    if (resolution.maxHeight) maxHeight = resolution.maxHeight;
  }

  const scaleX = sourceWidth > maxWidth ? maxWidth / sourceWidth : 1;
  const scaleY = sourceHeight > maxHeight ? maxHeight / sourceHeight : 1;
  const scale = Math.min(scaleX, scaleY);

  return {
    height: Math.max(1, Math.round(sourceHeight * scale)),
    scale,
    width: Math.max(1, Math.round(sourceWidth * scale)),
  };
}

function columnLabel(index: number): string {
  let label = "";
  let n = index;
  while (n >= 0) {
    label = String.fromCharCode((n % 26) + 65) + label;
    n = Math.floor(n / 26) - 1;
  }
  return label;
}

/**
 * Builds a thumbnail descriptor whose geometry and cell values are read on
 * demand. A workbook with dozens of sheets used to pay `rows × cols` formatted
 * lookups per sheet before the first hover; now nothing is read until a preview
 * is actually shown, and the cell values are reused until the sheet revision
 * changes.
 */
function createThumbnail(
  readController: () => XlsxViewerController,
  sheet: XlsxSheetData,
  sheetIndex: number,
  options: UseXlsxViewerThumbnailsOptions,
): XlsxSheetThumbnail {
  const includeHeaders = options.includeHeaders ?? true;
  const minRow = sheet.minUsedRow ?? 0;
  const maxRow = Math.max(sheet.maxUsedRow ?? -1, 0);
  const minCol = sheet.minUsedCol ?? 0;
  const maxCol = Math.max(sheet.maxUsedCol ?? -1, 0);
  const previewRows = Math.min(maxRow - minRow + 3, THUMBNAIL_MAX_ROWS);
  const previewCols = Math.min(maxCol - minCol + 3, THUMBNAIL_MAX_COLS);
  const headerHeight = includeHeaders ? HEADER_HEIGHT : 0;
  const rowHeaderWidth = includeHeaders ? ROW_HEADER_WIDTH : 0;

  let geometry: ThumbnailGeometry | null = null;
  let cells: ThumbnailCells | null = null;

  function loadGeometry(): ThumbnailGeometry {
    const worksheet = readController().workbook?.getSheet(sheet.workbookSheetIndex) ?? null;
    try {
      // Worker-backed viewers have no main-thread Workbook; their snapshot
      // already carries the resolved column widths and row heights.
      const workerWidthByColumn = worksheet
        ? null
        : new Map<number, number>((sheet.visibleCols ?? []).flatMap((column, index) => {
          const width = sheet.colWidths?.[index];
          return typeof width === "number" && Number.isFinite(width) && width > 0
            ? [[column, width] as const]
            : [];
        }));
      const workerHeightByRow = worksheet
        ? null
        : new Map<number, number>((sheet.visibleRows ?? []).flatMap((row, index) => {
          const height = sheet.rowHeights?.[index];
          return typeof height === "number" && Number.isFinite(height) && height > 0
            ? [[row, height] as const]
            : [];
        }));

      const colWidths: number[] = [];
      let totalColWidth = 0;
      for (let c = 0; c < previewCols; c++) {
        const actualCol = minCol + c;
        const rawWidth = worksheet
          ? worksheet.getColumnWidth(actualCol)
          : workerWidthByColumn?.get(actualCol)
            ?? sheet.colWidthOverridesPx?.[actualCol]
            ?? sheet.defaultColWidthPx;
        const width = rawWidth !== undefined && rawWidth !== null
          ? Math.max(rawWidth, DEFAULT_COL_WIDTH / 2)
          : DEFAULT_COL_WIDTH;
        colWidths.push(width);
        totalColWidth += width;
      }

      const rowHeights: number[] = [];
      let totalRowHeight = 0;
      for (let r = 0; r < previewRows; r++) {
        const actualRow = minRow + r;
        const rawHeight = worksheet
          ? worksheet.getRowHeight(actualRow)
          : workerHeightByRow?.get(actualRow)
            ?? sheet.rowHeightOverridesPx?.[actualRow]
            ?? sheet.defaultRowHeightPx;
        const height = rawHeight !== undefined && rawHeight !== null
          ? Math.max(rawHeight, DEFAULT_ROW_HEIGHT / 1.5)
          : DEFAULT_ROW_HEIGHT;
        rowHeights.push(height);
        totalRowHeight += height;
      }

      const sourceWidth = rowHeaderWidth + totalColWidth;
      const sourceHeight = headerHeight + totalRowHeight;
      return {
        colWidths,
        headerHeight,
        outputSize: resolveOutputSize(sourceWidth, sourceHeight, options.resolution),
        rowHeaderWidth,
        rowHeights,
        sourceHeight,
        sourceWidth,
        totalColWidth,
        totalRowHeight,
      };
    } finally {
      worksheet?.free();
    }
  }

  function loadCells(): ThumbnailCells {
    const revision = readController().revision;
    const worksheet = readController().workbook?.getSheet(sheet.workbookSheetIndex) ?? null;
    try {
      const values = Array.from({ length: previewRows }, (_, rowOffset) =>
        Array.from({ length: previewCols }, (_, colOffset) => {
          if (!worksheet) {
            // Worker snapshots carry cached formula values; the grid hydrates
            // literal cell values separately.
            const cached = sheet.cachedFormulaValues?.[cellAddressToA1({
              col: minCol + colOffset,
              row: minRow + rowOffset,
            })];
            return cached != null ? String(cached) : "";
          }
          const formatted = worksheet.getFormattedValueAt(minRow + rowOffset, minCol + colOffset);
          return formatted != null ? String(formatted) : "";
        }),
      );
      return { revision, values };
    } finally {
      worksheet?.free();
    }
  }

  function ensureGeometry(): ThumbnailGeometry {
    geometry ??= loadGeometry();
    return geometry;
  }

  function ensureCells(): string[][] {
    const revision = readController().revision;
    if (!cells || cells.revision !== revision) cells = loadCells();
    return cells.values;
  }

  const paint = (canvas: HTMLCanvasElement | null): boolean => {
    if (!canvas) return false;

    const context = canvas.getContext("2d");
    if (!context) return false;

    const { colWidths, rowHeights, totalColWidth, totalRowHeight, sourceWidth, sourceHeight, outputSize } = ensureGeometry();
    const cellValues = ensureCells();
    const dpr = typeof window === "undefined" ? 1 : Math.max(1, window.devicePixelRatio || 1);
    const scale = outputSize.scale;
    const dw = Math.max(1, Math.round(outputSize.width * dpr));
    const dh = Math.max(1, Math.round(outputSize.height * dpr));

    canvas.width = dw;
    canvas.height = dh;
    canvas.style.width = `${outputSize.width}px`;
    canvas.style.height = `${outputSize.height}px`;

    context.setTransform(dpr * scale, 0, 0, dpr * scale, 0, 0);

    // Background
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, sourceWidth, sourceHeight);

    // Grid surface
    context.fillStyle = "#fafafa";
    context.fillRect(rowHeaderWidth, headerHeight, totalColWidth, totalRowHeight);

    if (includeHeaders) {
      // Column headers
      context.fillStyle = "#f0f0f0";
      context.fillRect(rowHeaderWidth, 0, totalColWidth, headerHeight);
      context.fillStyle = "#333";
      context.font = "10px sans-serif";
      context.textAlign = "center";
      context.textBaseline = "middle";
      let cx = rowHeaderWidth;
      for (let c = 0; c < previewCols; c++) {
        const label = columnLabel(minCol + c);
        const w = colWidths[c];
        context.fillText(label, cx + w / 2, headerHeight / 2);
        // Grid line
        context.strokeStyle = "#d4d4d8";
        context.lineWidth = 0.5;
        context.beginPath();
        context.moveTo(cx, 0);
        context.lineTo(cx, sourceHeight);
        context.stroke();
        cx += w;
      }

      // Row headers
      context.fillStyle = "#f0f0f0";
      context.fillRect(0, headerHeight, rowHeaderWidth, totalRowHeight);
      context.fillStyle = "#666";
      context.font = "9px sans-serif";
      context.textAlign = "center";
      context.textBaseline = "middle";
      let ry = headerHeight;
      for (let r = 0; r < previewRows; r++) {
        const h = rowHeights[r];
        context.fillText(`${minRow + r + 1}`, rowHeaderWidth / 2, ry + h / 2);
        // Grid line
        context.strokeStyle = "#d4d4d8";
        context.lineWidth = 0.5;
        context.beginPath();
        context.moveTo(0, ry);
        context.lineTo(sourceWidth, ry);
        context.stroke();
        ry += h;
      }
    }

    // Cell values
    context.fillStyle = "#333";
    context.font = "8px sans-serif";
    context.textAlign = "left";
    context.textBaseline = "top";
    let ry = headerHeight;
    for (let r = 0; r < previewRows; r++) {
      let cx = rowHeaderWidth;
      for (let c = 0; c < previewCols; c++) {
        const value = cellValues[r]?.[c] ?? "";
        if (value) {
          const maxW = colWidths[c] - 4;
          const display = value.length > 15 ? value.slice(0, 14) + "…" : value;
          context.fillText(display, cx + 2, ry + 2, maxW);
        }
        cx += colWidths[c];
      }
      ry += rowHeights[r];
    }

    return true;
  };

  const sourceRange: XlsxCellRange = {
    start: { col: minCol, row: minRow },
    end: { col: minCol + previewCols - 1, row: minRow + previewRows - 1 },
  };

  return {
    get aspectRatio() {
      const { sourceHeight, sourceWidth } = ensureGeometry();
      return sourceWidth / Math.max(1, sourceHeight);
    },
    get contentHeight() {
      return ensureGeometry().totalRowHeight;
    },
    get contentWidth() {
      return ensureGeometry().totalColWidth;
    },
    get height() {
      return ensureGeometry().outputSize.height;
    },
    paint,
    sheet,
    sheetIndex,
    sourceRange,
    get width() {
      return ensureGeometry().outputSize.width;
    },
    workbookSheetIndex: sheet.workbookSheetIndex,
  };
}

export function useXlsxViewerThumbnails(
  controller: ComputedRef<XlsxViewerController> | (() => XlsxViewerController),
  options: UseXlsxViewerThumbnailsOptions = {},
): XlsxViewerThumbnails {
  const resolvedController = typeof controller === "function" ? controller : () => controller.value;

  const thumbnails = computed<XlsxSheetThumbnail[]>(() =>
    resolvedController().sheets.map((sheet, sheetIndex) =>
      createThumbnail(() => resolvedController(), sheet, sheetIndex, options),
    ),
  );

  function paintThumbnail(sheetIndex: number, canvas: HTMLCanvasElement | null): boolean {
    const t = thumbnails.value[sheetIndex];
    if (!t) return false;
    return t.paint(canvas);
  }

  return {
    paintThumbnail,
    get thumbnails(): XlsxSheetThumbnail[] {
      return thumbnails.value;
    },
  };
}
