import { computed, type ComputedRef } from "vue";
import type {
  XlsxCellTextRun,
  XlsxSheetData,
  XlsxSheetThumbnail,
  XlsxSheetThumbnailResolution,
  XlsxViewerController,
  XlsxViewerThumbnails,
  XlsxCellRange,
} from "@arcships/xlsx-core";
import { resolveRichTextRuns, resolveWorkbookColor } from "@arcships/xlsx-core";
import { cellAddressToA1 } from "./selection";
import {
  RICH_TEXT_VERTICAL_ALIGN_SCALE,
  isRichTextVerticalAlign,
  measureRichTextSegments,
  richTextBaselineOffset,
  truncateRichTextRuns,
} from "../components/rich-text-layout";
import {
  resolveCellTextLayout,
  resolveRotatedTextFitWidth,
  resolveShrinkToFitFontSize,
  type CellTextLayout,
} from "../components/grid-paint";

export interface UseXlsxViewerThumbnailsOptions {
  includeHeaders?: boolean;
  resolution?: XlsxSheetThumbnailResolution;
}

const THUMBNAIL_MAX_ROWS = 60;
const THUMBNAIL_MAX_COLS = 26;
/** Sheets warmed by one prefetch call; a hover asks for the tab and its neighbours. */
const THUMBNAIL_PREFETCH_LIMIT = 3;
const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const HEADER_HEIGHT = 24;
const ROW_HEADER_WIDTH = 40;
/** Base font of thumbnail cell text, shared by plain and rich cells. */
const THUMBNAIL_FONT_SIZE_PX = 8;
const THUMBNAIL_FONT = `${THUMBNAIL_FONT_SIZE_PX}px sans-serif`;

/** Thumbnail font of one run: weight and superscript scale on the cell's size. */
function thumbnailRunFont(
  font: XlsxCellTextRun["font"],
  fontSizePx = THUMBNAIL_FONT_SIZE_PX,
): string {
  const sizePx = Math.max(1, fontSizePx * (isRichTextVerticalAlign(font, "superscript") || isRichTextVerticalAlign(font, "subscript") ? RICH_TEXT_VERTICAL_ALIGN_SCALE : 1));
  const weight = font?.bold ? "bold " : "";
  const style = font?.italic ? "italic " : "";
  return `${style}${weight}${sizePx}px sans-serif`;
}

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
  /** Rotation / shrink-to-fit flags per cell, `null` for plain layout. */
  layouts: Array<Array<CellTextLayout | null>>;
  /** Rich-text runs per cell, `undefined` when the cell paints plain text. */
  runs: Array<Array<XlsxCellTextRun[] | undefined>>;
  revision: number;
  values: string[][];
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
): ThumbnailEntry {
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
      const runs = values.map((row, rowOffset) =>
        row.map((value, colOffset) => resolveThumbnailRuns(value, minRow + rowOffset, minCol + colOffset)),
      );
      const layouts = values.map((row, rowOffset) =>
        row.map((_, colOffset) => resolveThumbnailLayout(worksheet, minRow + rowOffset, minCol + colOffset)),
      );
      return { layouts, revision, runs, values };
    } finally {
      worksheet?.free();
    }
  }

  /**
   * Runs of a thumbnail cell.
   *
   * The sheet state is the only source: it carries the parsed runs in both
   * main-thread and Worker-backed viewers, and `resolveRichTextRuns` drops them
   * once an edit made the cell's text differ from the parsed runs.
   */
  function resolveThumbnailRuns(value: string, row: number, col: number): XlsxCellTextRun[] | undefined {
    return resolveRichTextRuns(sheet.richTextCells?.[cellAddressToA1({ col, row })], value);
  }

  /**
   * Rotation / shrink-to-fit of a thumbnail cell.
   *
   * A main-thread worksheet resolves the exact cell style; a Worker snapshot
   * carries only the cells whose explicit style needed text layout, so the
   * row/column style is the last fallback for everything else.
   */
  function resolveThumbnailLayout(
    worksheet: ReturnType<XlsxViewerController["getActiveWorksheet"]>,
    row: number,
    col: number,
  ): CellTextLayout | null {
    let alignment: Record<string, unknown> | undefined;
    if (worksheet) {
      try {
        const style = worksheet.getCellStyleAt(row, col) as { alignment?: Record<string, unknown> } | null | undefined;
        alignment = style?.alignment;
      } catch {
        alignment = undefined;
      }
    }
    if (!alignment) {
      const address = cellAddressToA1({ col, row });
      const parsed = sheet.cellTextLayouts?.[address];
      if (parsed) {
        alignment = {
          ...(parsed.textRotation !== undefined ? { textRotation: parsed.textRotation } : {}),
          ...(parsed.shrinkToFit ? { shrinkToFit: true } : {}),
        };
      } else {
        const rowStyleId = sheet.rowStyleIds?.[row];
        const colStyleId = sheet.colStyleIds?.[col];
        const fallback = rowStyleId !== undefined
          ? sheet.styleById?.[rowStyleId]
          : colStyleId !== undefined ? sheet.styleById?.[colStyleId] : undefined;
        alignment = fallback?.alignment;
      }
    }
    const layout = alignment ? resolveCellTextLayout(alignment) : { rotation: null, shrinkToFit: false };
    return layout.rotation !== null || layout.shrinkToFit ? layout : null;
  }

  function ensureGeometry(): ThumbnailGeometry {
    geometry ??= loadGeometry();
    return geometry;
  }

  function ensureCells(): ThumbnailCells {
    const revision = readController().revision;
    if (!cells || cells.revision !== revision) cells = loadCells();
    return cells;
  }

  /** Loads whatever `paint` would load, without needing a canvas. */
  function warm(): void {
    ensureGeometry();
    ensureCells();
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
    context.font = THUMBNAIL_FONT;
    context.textAlign = "left";
    context.textBaseline = "top";
    let ry = headerHeight;
    for (let r = 0; r < previewRows; r++) {
      let cx = rowHeaderWidth;
      for (let c = 0; c < previewCols; c++) {
        const value = cellValues.values[r]?.[c] ?? "";
        if (value) {
          paintThumbnailText(
            context,
            value,
            cellValues.runs[r]?.[c],
            cellValues.layouts[r]?.[c] ?? null,
            cx,
            ry,
            colWidths[c] ?? DEFAULT_COL_WIDTH,
            rowHeights[r] ?? DEFAULT_ROW_HEIGHT,
          );
        }
        cx += colWidths[c];
      }
      ry += rowHeights[r];
    }

    return true;
  };

  /** Natural width of the thumbnail cell's runs at one font size. */
  function measureThumbnailRuns(
    context: CanvasRenderingContext2D,
    runs: XlsxCellTextRun[],
    fontSizePx: number,
  ): number {
    const measure = (text: string, font: XlsxCellTextRun["font"]) => {
      context.font = thumbnailRunFont(font, fontSizePx);
      return context.measureText(text).width;
    };
    return measureRichTextSegments(runs, measure);
  }

  /**
   * Paints one thumbnail cell, applying rotation and shrink-to-fit the same
   * way the grid does.
   *
   * Rotated text is clipped to its cell rectangle and anchored by its rotated
   * bounds, so a 90° header uses the row height as its line length instead of
   * painting over the next column. Shrink-to-fit binary-searches the 8px
   * thumbnail base down to 0.4× using the width the thumbnail actually paints.
   */
  function paintThumbnailText(
    context: CanvasRenderingContext2D,
    value: string,
    runs: XlsxCellTextRun[] | undefined,
    layout: CellTextLayout | null,
    x: number,
    y: number,
    width: number,
    height: number,
  ): void {
    const maxWidth = Math.max(0, width - 4);
    const rotation = layout?.rotation ?? null;
    let fontSizePx = THUMBNAIL_FONT_SIZE_PX;
    if (rotation !== null) {
      context.save();
      context.beginPath();
      context.rect(x, y, width, height);
      context.clip();
      if (layout?.shrinkToFit) {
        const fitWidth = Math.max(0, resolveRotatedTextFitWidth(
          maxWidth,
          Math.max(0, height - 2),
          rotation,
          Math.max(8, fontSizePx * 1.25),
        ));
        fontSizePx = resolveShrinkToFitFontSize(
          THUMBNAIL_FONT_SIZE_PX,
          fitWidth,
          (candidate) => runs
            ? measureThumbnailRuns(context, runs, candidate)
            : (context.font = `${candidate}px sans-serif`, context.measureText(value).width),
        );
      }
      context.translate(x + width / 2, y + height / 2);
      context.rotate(rotation * Math.PI / 180);
      context.textAlign = "left";
      context.textBaseline = "middle";
      context.fillStyle = "#333";
      if (runs) {
        const naturalWidth = measureThumbnailRuns(context, runs, fontSizePx);
        paintRichTextRuns(context, runs, -naturalWidth / 2, 0, Number.POSITIVE_INFINITY, fontSizePx);
      } else {
        context.font = `${fontSizePx}px sans-serif`;
        const textWidth = context.measureText(value).width;
        context.fillText(value, -textWidth / 2, 0);
      }
      context.restore();
      return;
    }

    if (layout?.shrinkToFit) {
      fontSizePx = resolveShrinkToFitFontSize(
        THUMBNAIL_FONT_SIZE_PX,
        maxWidth,
        (candidate) => runs
          ? measureThumbnailRuns(context, runs, candidate)
          : (context.font = `${candidate}px sans-serif`, context.measureText(value).width),
      );
    }
    if (runs) {
      paintRichTextRuns(context, runs, x + 2, y + 2, maxWidth, fontSizePx);
    } else {
      const display = layout?.shrinkToFit || value.length <= 15 ? value : value.slice(0, 14) + "…";
      context.font = `${fontSizePx}px sans-serif`;
      context.fillText(display, x + 2, y + 2, maxWidth);
    }
  }

  /**
   * Paints one thumbnail cell run by run.
   *
   * Runs keep their weight, colour and superscript/subscript shape at the
   * requested size and clip to the width measured with each run's own font.
   */
  function paintRichTextRuns(
    context: CanvasRenderingContext2D,
    runs: XlsxCellTextRun[],
    x: number,
    y: number,
    maxWidth: number,
    fontSizePx = THUMBNAIL_FONT_SIZE_PX,
  ): void {
    const measure = (text: string, font: XlsxCellTextRun["font"]) => {
      context.font = thumbnailRunFont(font, fontSizePx);
      return context.measureText(text).width;
    };
    const visible = truncateRichTextRuns(runs, measure, maxWidth);
    let cursor = x;
    for (const segment of visible) {
      if (!segment.text) continue;
      const segmentWidth = measure(segment.text, segment.font);
      const color = segment.font?.color
        ? resolveWorkbookColor(segment.font.color as Record<string, unknown>, sheet.themePalette)
        : null;
      context.font = thumbnailRunFont(segment.font, fontSizePx);
      context.fillStyle = color ?? "#333";
      context.fillText(
        segment.text,
        cursor,
        y + richTextBaselineOffset(segment.font, fontSizePx),
      );
      cursor += segmentWidth;
    }
    context.fillStyle = "#333";
    context.font = THUMBNAIL_FONT;
  }

  const sourceRange: XlsxCellRange = {
    start: { col: minCol, row: minRow },
    end: { col: minCol + previewCols - 1, row: minRow + previewRows - 1 },
  };

  const thumbnail: XlsxSheetThumbnail = {
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
  return { thumbnail, warm };
}

interface ThumbnailEntry {
  thumbnail: XlsxSheetThumbnail;
  warm: () => void;
}

export function useXlsxViewerThumbnails(
  controller: ComputedRef<XlsxViewerController> | (() => XlsxViewerController),
  options: UseXlsxViewerThumbnailsOptions = {},
): XlsxViewerThumbnails {
  const resolvedController = typeof controller === "function" ? controller : () => controller.value;

  const thumbnailEntries = computed<ThumbnailEntry[]>(() =>
    resolvedController().sheets.map((sheet, sheetIndex) =>
      createThumbnail(() => resolvedController(), sheet, sheetIndex, options),
    ),
  );

  const thumbnails = computed<XlsxSheetThumbnail[]>(() =>
    thumbnailEntries.value.map((entry) => entry.thumbnail),
  );

  // Prefetch bookkeeping: at most `THUMBNAIL_PREFETCH_LIMIT` sheets per call,
  // one warm-up per sheet per revision, unknown indices ignored.
  const prefetchedSheets = new Set<number>();
  let prefetchedRevision: number | null = null;

  function prefetchThumbnails(sheetIndices: readonly number[]): void {
    const revision = resolvedController().revision;
    if (prefetchedRevision !== revision) {
      prefetchedRevision = revision;
      prefetchedSheets.clear();
    }
    const entries = thumbnailEntries.value;
    let warmed = 0;
    for (const sheetIndex of sheetIndices) {
      if (warmed >= THUMBNAIL_PREFETCH_LIMIT) break;
      if (!Number.isInteger(sheetIndex) || sheetIndex < 0 || sheetIndex >= entries.length) continue;
      if (prefetchedSheets.has(sheetIndex)) continue;
      prefetchedSheets.add(sheetIndex);
      entries[sheetIndex].warm();
      warmed += 1;
    }
  }

  function paintThumbnail(sheetIndex: number, canvas: HTMLCanvasElement | null): boolean {
    const t = thumbnails.value[sheetIndex];
    if (!t) return false;
    return t.paint(canvas);
  }

  return {
    paintThumbnail,
    prefetchThumbnails,
    get thumbnails(): XlsxSheetThumbnail[] {
      return thumbnails.value;
    },
  };
}
