import type { XlsxCellRange } from "@arcships/xlsx-core";

/**
 * Pure geometry / paint helpers shared by `XlsxGrid.vue`.
 *
 * They live outside the component so the fidelity rules (device-pixel caps,
 * merged-region viewport hits, icon-set buckets, overflow width, reference
 * outline clamps and resize zoom basis) can be unit-tested without a Canvas
 * implementation.
 */

// ── Canvas backing store ───────────────────────────────────────────────────

/**
 * Browser canvas limits vary, but 16.7 M pixels (4096 × 4096) and a 32 767 px
 * edge are safe across Blink, Gecko and WebKit. Beyond them browsers either
 * refuse to allocate the backing store or paint nothing, which blanks the grid.
 */
export const MAX_CANVAS_AREA_PX = 16_777_216;
export const MAX_CANVAS_DIMENSION_PX = 32_767;
export const MIN_CANVAS_PIXEL_RATIO = 0.1;

/** Current device-pixel ratio, read lazily so display / browser zoom changes apply. */
export function readDevicePixelRatio(): number {
  if (typeof window === "undefined") return 1;
  const value = window.devicePixelRatio;
  return Number.isFinite(value) && value > 0 ? value : 1;
}

/**
 * Effective backing-store ratio for a CSS-sized canvas. Starts from the live
 * `devicePixelRatio` and lowers it when the requested pixel area / edge would
 * exceed the browser limit, so the grid degrades in sharpness instead of
 * disappearing.
 */
export function resolveCanvasPixelRatio(
  cssWidth: number,
  cssHeight: number,
  devicePixelRatio: number,
): number {
  const width = Number.isFinite(cssWidth) ? Math.max(1, cssWidth) : 1;
  const height = Number.isFinite(cssHeight) ? Math.max(1, cssHeight) : 1;
  const requested = Number.isFinite(devicePixelRatio) && devicePixelRatio > 0 ? devicePixelRatio : 1;
  const areaScale = Math.sqrt(MAX_CANVAS_AREA_PX / (width * height));
  const dimensionScale = Math.min(
    MAX_CANVAS_DIMENSION_PX / width,
    MAX_CANVAS_DIMENSION_PX / height,
  );
  return Math.max(MIN_CANVAS_PIXEL_RATIO, Math.min(requested, areaScale, dimensionScale));
}

// ── Zoom basis ─────────────────────────────────────────────────────────────

export function resolveZoomFactor(zoomScale: number): number {
  const scale = zoomScale / 100;
  return Number.isFinite(scale) ? Math.max(0.1, scale) : 1;
}

/**
 * Row / column resize uses the same zoom basis as painting. The controller
 * allows zooming down to 10 %, so clamping at 0.5 made the first drag jump by
 * the ratio between the two scales.
 */
export function resolveAxisResizeSize(
  startSize: number,
  deltaPx: number,
  zoomScale: number,
  minimum: number,
): number {
  return Math.max(minimum, startSize + deltaPx / resolveZoomFactor(zoomScale));
}

// ── Per-frame identity keys ────────────────────────────────────────────────

/** XLSX's maximum column count; every cell key packs (row, col) into one number. */
export const CELL_KEY_STRIDE = 16_384;
/** XLSX's maximum row count, so (sheet, row, col) also stays one number. */
export const ROW_KEY_STRIDE = 1_048_576;

/**
 * Numeric identity for a sheet cell. Repaints resolve merge anchors, comments,
 * hyperlinks, sparklines and worker snapshots per visible cell; string keys
 * (`` `${row}:${col}` ``) allocated several throwaway strings per cell per
 * frame, numbers do not.
 */
export function cellKey(row: number, col: number): number {
  return row * CELL_KEY_STRIDE + col;
}

/** Numeric identity for a cell that also carries its workbook sheet index. */
export function sheetCellKey(sheetIndex: number, row: number, col: number): number {
  return (sheetIndex * ROW_KEY_STRIDE + row) * CELL_KEY_STRIDE + col;
}

// ── Merged regions ─────────────────────────────────────────────────────────

export interface PixelRect {
  height: number;
  left: number;
  top: number;
  width: number;
}

export interface PixelViewport {
  bottom: number;
  left: number;
  right: number;
  top: number;
}

export function intersectsViewport(rect: PixelRect, viewport: PixelViewport): boolean {
  return rect.width > 0
    && rect.height > 0
    && rect.left < viewport.right
    && rect.left + rect.width > viewport.left
    && rect.top < viewport.bottom
    && rect.top + rect.height > viewport.top;
}

// ── Formula reference outlines ─────────────────────────────────────────────

/**
 * Maps a sheet-coordinate reference range onto display coordinates for the
 * grid's outline pass.
 *
 * Whole-column (`A:A`) and whole-row (`1:1`) references end at Excel's maximum
 * index, which is not in the display map. Clamping the missing end to the last
 * displayed row/column draws the visible portion of the reference instead of
 * dropping the outline. A single off-screen cell has both endpoints unmapped and
 * is skipped rather than stretched over the viewport.
 */
export function resolvePaintedReferenceRange(
  range: XlsxCellRange,
  rowToDisplayIndex: ReadonlyMap<number, number>,
  colToDisplayIndex: ReadonlyMap<number, number>,
  displayRowCount: number,
  displayColCount: number,
): XlsxCellRange | null {
  if (displayRowCount <= 0 || displayColCount <= 0) return null;
  const startRow = rowToDisplayIndex.get(range.start.row);
  const endRow = rowToDisplayIndex.get(range.end.row);
  const startCol = colToDisplayIndex.get(range.start.col);
  const endCol = colToDisplayIndex.get(range.end.col);
  const wholeColumn = range.start.row === 0 && range.end.row === ROW_KEY_STRIDE - 1;
  const wholeRow = range.start.col === 0 && range.end.col === CELL_KEY_STRIDE - 1;

  if (startRow === undefined && endRow === undefined) {
    if (!wholeColumn) return null;
  }
  if (startCol === undefined && endCol === undefined) {
    if (!wholeRow) return null;
  }

  return {
    start: {
      row: startRow ?? 0,
      col: startCol ?? 0,
    },
    end: {
      row: endRow ?? Math.max(0, displayRowCount - 1),
      col: endCol ?? Math.max(0, displayColCount - 1),
    },
  };
}

// ── Text overflow ──────────────────────────────────────────────────────────

/**
 * Excel lets left-aligned text spill into adjacent empty cells. `neighbourWidths`
 * must only contain the cells between this one and the first occupied cell.
 */
export function resolveOverflowTextWidth(
  cellWidth: number,
  neighbourWidths: readonly number[],
): number {
  let width = Number.isFinite(cellWidth) ? Math.max(0, cellWidth) : 0;
  for (const neighbour of neighbourWidths) {
    if (!Number.isFinite(neighbour) || neighbour <= 0) continue;
    width += neighbour;
  }
  return width;
}

// ── Text rotation and shrink-to-fit ────────────────────────────────────────

/**
 * Parsed text-layout flags of a cell's alignment.
 *
 * `rotation` is the canvas angle: Excel stores 0–90 for counter-clockwise
 * rotation and 91–180 for clockwise rotation, while canvas `rotate()` treats
 * positive angles as clockwise. `null` means horizontal text.
 */
export interface CellTextLayout {
  rotation: number | null;
  shrinkToFit: boolean;
}

/**
 * Converts OOXML `textRotation` to the angle passed to `CanvasRenderingContext2D.rotate`.
 *
 * Excel measures rotation counter-clockwise, so 45° becomes -45 canvas
 * degrees; values above 90° mean clockwise rotation, e.g. 180° is +90 canvas
 * degrees (top-to-bottom vertical text). Values outside 0–180 (including the
 * legacy 255 vertical-stack marker) are horizontal text.
 */
export function resolveSpreadsheetTextRotation(textRotation: unknown): number | null {
  if (typeof textRotation !== "number" || !Number.isFinite(textRotation)) return null;
  if (textRotation < 0 || textRotation > 180) return null;
  const excelDegrees = textRotation > 90 ? 90 - textRotation : textRotation;
  return excelDegrees === 0 ? null : -excelDegrees;
}

/**
 * Text-layout flags of a resolved cell style.
 *
 * The workbook engine reports `alignment.rotation` while the OOXML parser
 * keeps the raw `alignment.textRotation`; both spellings must render.
 */
export function resolveCellTextLayout(
  alignment: Record<string, unknown> | undefined,
): CellTextLayout {
  return {
    rotation: resolveSpreadsheetTextRotation(alignment?.rotation ?? alignment?.textRotation),
    shrinkToFit: alignment?.shrinkToFit === true,
  };
}

/** Axis-aligned bounds of a `width × height` box rotated by `degrees`. */
export function resolveRotatedTextBounds(width: number, height: number, degrees: number): {
  height: number;
  width: number;
} {
  const radians = degrees * Math.PI / 180;
  const cos = Math.abs(Math.cos(radians));
  const sin = Math.abs(Math.sin(radians));
  const safeWidth = Number.isFinite(width) ? Math.max(0, width) : 0;
  const safeHeight = Number.isFinite(height) ? Math.max(0, height) : 0;
  return {
    height: Math.abs(safeWidth * sin) + Math.abs(safeHeight * cos),
    width: Math.abs(safeWidth * cos) + Math.abs(safeHeight * sin),
  };
}

/**
 * Longest single line that keeps a rotated text box inside the cell.
 *
 * Rotation swaps the space text can occupy: at 90° the line may be as tall as
 * the cell, and at 45° both the cell's width and height constrain the rotated
 * bounds. Callers that shrink text use this as the target width; normal
 * rotated text is clipped to the same cell rectangle, so it never paints over
 * a neighbour.
 */
export function resolveRotatedTextFitWidth(
  cellWidth: number,
  cellHeight: number,
  degrees: number,
  lineHeight: number,
): number {
  const width = Number.isFinite(cellWidth) ? Math.max(0, cellWidth) : 0;
  const height = Number.isFinite(cellHeight) ? Math.max(0, cellHeight) : 0;
  const line = Number.isFinite(lineHeight) ? Math.max(0, lineHeight) : 0;
  const radians = degrees * Math.PI / 180;
  const cos = Math.abs(Math.cos(radians));
  const sin = Math.abs(Math.sin(radians));
  let maxWidth = Number.POSITIVE_INFINITY;
  if (cos > 1e-6) maxWidth = Math.min(maxWidth, (width - line * sin) / cos);
  if (sin > 1e-6) maxWidth = Math.min(maxWidth, (height - line * cos) / sin);
  return Number.isFinite(maxWidth) ? Math.max(0, maxWidth) : width;
}

/**
 * Centre of a rotated text block, positioned the way Excel anchors it.
 *
 * The rotated bounds — not the unrotated text size — decide the distance from
 * the left/right and top/bottom edges, which is what swaps the occupied
 * width/height at 90°.
 */
export function resolveRotatedTextCenter(
  rect: PixelRect,
  bounds: { height: number; width: number },
  alignment: "left" | "center" | "right",
  vertical: "top" | "middle" | "bottom",
): { x: number; y: number } {
  return {
    x: alignment === "right"
      ? rect.left + rect.width - bounds.width / 2
      : alignment === "center"
        ? rect.left + rect.width / 2
        : rect.left + bounds.width / 2,
    y: vertical === "top"
      ? rect.top + bounds.height / 2
      : vertical === "middle"
        ? rect.top + rect.height / 2
        : rect.top + rect.height - bounds.height / 2,
  };
}

/**
 * Largest font size at or below `baseFontSizePx` whose measured width still
 * fits `availableWidth`.
 *
 * Excel's shrink-to-fit never goes below 0.4× the base font; the binary search
 * stops within a quarter pixel, which is enough for the canvas to round the
 * same way the measure pass did. Text that is still too wide at the floor
 * returns the floor so the caller can clip or ellipsize it.
 */
export function resolveShrinkToFitFontSize(
  baseFontSizePx: number,
  availableWidth: number,
  measureText: (fontSizePx: number) => number,
  minimumScale = 0.4,
): number {
  const base = Number.isFinite(baseFontSizePx) ? Math.max(1, baseFontSizePx) : 1;
  if (!Number.isFinite(availableWidth) || availableWidth <= 0) return base;
  if (measureText(base) <= availableWidth) return base;

  const scale = Number.isFinite(minimumScale) ? Math.min(1, Math.max(0.01, minimumScale)) : 0.4;
  const minimum = Math.max(1, base * scale);
  if (measureText(minimum) > availableWidth) return minimum;

  let lower = minimum;
  let upper = base;
  while (upper - lower > 0.25) {
    const candidate = (lower + upper) / 2;
    if (measureText(candidate) <= availableWidth) {
      lower = candidate;
    } else {
      upper = candidate;
    }
  }
  return Math.min(base, lower);
}

// ── Axis auto-fit ──────────────────────────────────────────────────────────

/** Text inset the paint path keeps on each side of a cell, in CSS pixels. */
export const AUTOFIT_CELL_PADDING_PX = 5;
/** Rounding room for the last glyph's overhang and the cell's gridline. */
export const AUTOFIT_EXTRA_PX = 2;
/** Vertical chrome the paint path reserves above and below a cell's text. */
export const AUTOFIT_ROW_CHROME_PX = 4;
/** Excel will not auto-fit a column narrower than the drag-resize floor. */
export const AUTOFIT_MIN_COLUMN_WIDTH_PX = 32;
/**
 * Excel's widest column is 255 characters; 1 800 px keeps the fitted width
 * inside what the engine can round-trip back to a character count.
 */
export const AUTOFIT_MAX_COLUMN_WIDTH_PX = 1_800;
/** Excel's shortest row still shows one line of the default font. */
export const AUTOFIT_MIN_ROW_HEIGHT_PX = 18;
/** Excel's tallest row (409 points at 96 DPI). */
export const AUTOFIT_MAX_ROW_HEIGHT_PX = 546;
/**
 * Cells an axis auto-fit measures before it stops.
 *
 * The scan is bounded by the sheet's used range, which a formatted-but-empty
 * far-away cell can stretch to a million rows; measuring stops at this budget
 * so a double-click can never freeze the viewer.
 */
export const AUTOFIT_SCAN_LIMIT = 20_000;

/** Measures one piece of text in one canvas font, in CSS pixels. */
export type MeasureAutofitText = (text: string, font: string) => number;

/**
 * Everything the paint path knows about a cell's text, handed to the auto-fit
 * measurement so it sizes cells with the same font, line height, wrapping and
 * rotation that `drawCellText` uses.
 */
export interface AutofitCellText {
  /** Text width the cell currently has inside its column, padding excluded. */
  availableWidth: number;
  /** Canvas font string `drawCellText` sets for the base style. */
  font: string;
  /** Line height `drawCellText` steps by, in CSS pixels. */
  lineHeight: number;
  /** Canvas angle of rotated text; `null` for horizontal text. */
  rotation: number | null;
  /** Inline rich-text runs, each measured with its own font like the paint path. */
  runs?: readonly { font: string; text: string }[];
  /** Formatted display text. */
  value: string;
  /** Whether the cell wraps its text inside its column. */
  wrap: boolean;
}

/** Text box an axis has to fit, in the same (zoomed) pixels it was measured in. */
export interface AutofitContentBox {
  height: number;
  width: number;
}

export interface AutofitOptions {
  /**
   * Extra room added to the measured box; defaults to
   * {@link AUTOFIT_EXTRA_PX} for a column and {@link AUTOFIT_ROW_CHROME_PX}
   * for a row.
   */
  extraPx?: number;
  maximumPx?: number;
  minimumPx?: number;
  /** Per-side text inset; defaults to {@link AUTOFIT_CELL_PADDING_PX}. */
  paddingPx?: number;
  /** Grid zoom factor the cell fonts were resolved with (1 = 100 %). */
  zoom?: number;
}

/**
 * Advance width of `text` when no canvas is available (SSR, tests).
 *
 * CJK and full-width glyphs are square, Latin text averages a little over half
 * an em; the font string is the same one the canvas would have been given, so
 * the estimate still tracks the cell's real font size.
 */
export function estimateAutofitTextWidth(text: string, font: string): number {
  const sizeMatch = /([\d.]+)px/.exec(font);
  const size = sizeMatch ? Number(sizeMatch[1]) : 12;
  const fontSize = Number.isFinite(size) && size > 0 ? size : 12;
  let width = 0;
  for (const character of text) {
    width += /[\u1100-\u115f\u2e80-\ua4cf\uac00-\ud7a3\uf900-\ufaff\ufe30-\ufe4f\uff00-\uff60\uffe0-\uffe6]/u.test(character)
      ? fontSize
      : fontSize * 0.55;
  }
  return width;
}

/** Hard line breaks of a cell value, matching the paint path's `\r\n` handling. */
function autofitLines(value: string): string[] {
  return value.replace(/\r\n?/g, "\n").split("\n");
}

/**
 * Natural width of a cell's text: the widest hard line, or the sum of the rich
 * runs' own fonts, exactly as the paint path draws it without wrapping.
 */
export function measureAutofitTextWidth(
  cell: AutofitCellText,
  measure: MeasureAutofitText,
): number {
  if (cell.runs?.length) {
    let total = 0;
    for (const run of cell.runs) total += measure(run.text, run.font);
    return total;
  }
  let widest = 0;
  for (const line of autofitLines(cell.value)) {
    widest = Math.max(widest, measure(line, cell.font));
  }
  return widest;
}

/**
 * Painted line count of a cell at its current width.
 *
 * Wrapped cells use the same greedy per-character break as `wrapCellText`;
 * everything else (and rotated text, which Excel never wraps) is one line per
 * hard line break.
 */
export function countAutofitLines(
  cell: AutofitCellText,
  measure: MeasureAutofitText,
): number {
  const lines = autofitLines(cell.value);
  if (!cell.wrap || cell.rotation !== null) return Math.max(1, lines.length);
  const available = Number.isFinite(cell.availableWidth) ? Math.max(1, cell.availableWidth) : 1;
  let count = 0;
  for (const line of lines) {
    if (!line) {
      count += 1;
      continue;
    }
    let current = "";
    let breaks = 0;
    for (const character of line) {
      const next = current + character;
      if (current && measure(next, cell.font) > available) {
        breaks += 1;
        current = character;
      } else {
        current = next;
      }
    }
    count += breaks + 1;
  }
  return Math.max(1, count);
}

/**
 * Box a cell's text needs, rotation included.
 *
 * Rotation swaps the occupied width and height, so a 90° header only needs one
 * line of column width while its text needs the column's whole height.
 */
export function measureAutofitContent(
  cell: AutofitCellText,
  measure: MeasureAutofitText,
): AutofitContentBox {
  const width = measureAutofitTextWidth(cell, measure);
  const lineHeight = Number.isFinite(cell.lineHeight) && cell.lineHeight > 0 ? cell.lineHeight : 12;
  const height = Math.max(lineHeight, countAutofitLines(cell, measure) * lineHeight);
  return cell.rotation === null ? { height, width } : resolveRotatedTextBounds(width, height, cell.rotation);
}

function autofitNumber(value: number | undefined, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function autofitScale(value: number | undefined): number {
  return Number.isFinite(value) && (value as number) > 0 ? (value as number) : 1;
}

function clampAutofitSize(size: number, minimumPx: number, maximumPx: number): number {
  if (!Number.isFinite(size)) return minimumPx;
  return Math.max(minimumPx, Math.min(maximumPx, size));
}

/**
 * Column width that fits every measured cell, in unzoomed CSS pixels.
 *
 * The widths come from the same font, line height and rich-run measurement the
 * paint path uses, so the fitted column shows the value without shrinking or
 * spilling it into the next column. An empty column yields the minimum width;
 * callers that want to leave untouched columns alone skip the resize instead.
 */
export function resolveAutofitColumnWidth(
  cells: readonly AutofitCellText[],
  measure: MeasureAutofitText,
  options: AutofitOptions = {},
): number {
  const padding = Math.max(0, autofitNumber(options.paddingPx, AUTOFIT_CELL_PADDING_PX));
  const extra = Math.max(0, autofitNumber(options.extraPx, AUTOFIT_EXTRA_PX));
  const zoom = autofitScale(options.zoom);
  let width = 0;
  for (const cell of cells) {
    width = Math.max(width, measureAutofitContent(cell, measure).width);
  }
  const fitted = (width + padding * 2 + extra) / zoom;
  return clampAutofitSize(
    fitted,
    autofitNumber(options.minimumPx, AUTOFIT_MIN_COLUMN_WIDTH_PX),
    autofitNumber(options.maximumPx, AUTOFIT_MAX_COLUMN_WIDTH_PX),
  );
}

/**
 * Row height that fits every measured cell, in unzoomed CSS pixels.
 *
 * Wrapped cells count the lines they paint at their current column width, so a
 * row grows until its tallest wrapped value is fully visible.
 */
export function resolveAutofitRowHeight(
  cells: readonly AutofitCellText[],
  measure: MeasureAutofitText,
  options: AutofitOptions = {},
): number {
  const extra = Math.max(0, autofitNumber(options.extraPx, AUTOFIT_ROW_CHROME_PX));
  const zoom = autofitScale(options.zoom);
  let height = 0;
  for (const cell of cells) {
    height = Math.max(height, measureAutofitContent(cell, measure).height);
  }
  const fitted = (height + extra) / zoom;
  return clampAutofitSize(
    fitted,
    autofitNumber(options.minimumPx, AUTOFIT_MIN_ROW_HEIGHT_PX),
    autofitNumber(options.maximumPx, AUTOFIT_MAX_ROW_HEIGHT_PX),
  );
}
