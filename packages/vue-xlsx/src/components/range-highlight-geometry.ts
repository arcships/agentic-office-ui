import type { CSSProperties } from "vue";
import type { XlsxCellRange, XlsxImageRect, XlsxSheetData } from "@arcships/xlsx-core";
import { createGridPaneLayout, gridZoomFactor, type GridPane, type GridPaneKey } from "./grid-metrics";

const DEFAULT_COL_WIDTH_PX = 80;
const DEFAULT_ROW_HEIGHT_PX = 24;

export interface RangeHighlightViewport {
  height: number;
  width: number;
}

/** One range fragment positioned inside a freeze pane. */
export interface RangeHighlightSegment {
  pane: GridPaneKey;
  /** Pane-local pixel rect; width / height may be clipped at a pane edge. */
  rect: XlsxImageRect;
}

export interface RangeHighlightLayout {
  frozenHeight: number;
  frozenWidth: number;
  panes: GridPane[];
  /**
   * Resolves a zero-based range into one fragment per freeze pane it covers.
   * Fragments are pane-local: the scrolling panes subtract the current scroll
   * offset, frozen panes keep their content position, mirroring the grid.
   * Passing a viewport clips off-screen rows/columns so an enormous whole-row
   * or whole-column highlight never allocates a huge DOM box.
   */
  resolve: (range: XlsxCellRange, viewport?: RangeHighlightViewport) => RangeHighlightSegment[];
}

interface AxisGeometry {
  /** Workbook row/column index for each display slot. */
  indices: number[];
  /** Rendered size per display slot, zoom applied. */
  sizes: number[];
  /** Content offset per display slot plus the end offset. */
  offsets: number[];
  /** Number of leading display slots inside the frozen pane. */
  freezeCount: number;
  frozenExtent: number;
}

// Scroll ticks re-create the layout but the axis model only changes when the
// sheet object or zoom changes. Sheet objects are replaced on every edit, so a
// WeakMap keys the cache without retaining old workbooks.
const axisGeometryCache = new WeakMap<XlsxSheetData, Map<string, AxisGeometry>>();

function resolveAxisIndices(
  sheet: XlsxSheetData,
  axis: "row" | "column",
): number[] {
  const visible = axis === "row" ? sheet.visibleRows : sheet.visibleCols;
  if (visible?.length) return visible;
  const count = Math.max(
    axis === "row" ? sheet.rowHeights?.length ?? 0 : sheet.colWidths?.length ?? 0,
    axis === "row" ? sheet.rowCount : sheet.colCount,
    1,
  );
  return Array.from({ length: count }, (_, index) => index);
}

function resolveAxisSize(
  sheet: XlsxSheetData,
  axis: "row" | "column",
  displayIndex: number,
  actualIndex: number,
  zoom: number,
): number {
  const size = axis === "row"
    ? sheet.rowHeights?.[displayIndex]
      || sheet.rowHeightOverridesPx?.[actualIndex]
      || sheet.defaultRowHeightPx
      || DEFAULT_ROW_HEIGHT_PX
    : sheet.colWidths?.[displayIndex]
      || sheet.colWidthOverridesPx?.[actualIndex]
      || sheet.defaultColWidthPx
      || DEFAULT_COL_WIDTH_PX;
  return size * zoom;
}

function lowerBound(indices: readonly number[], value: number): number {
  let low = 0;
  let high = indices.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if ((indices[middle] ?? 0) < value) low = middle + 1;
    else high = middle;
  }
  return low;
}

function upperBound(indices: readonly number[], value: number): number {
  let low = 0;
  let high = indices.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if ((indices[middle] ?? 0) <= value) low = middle + 1;
    else high = middle;
  }
  return low;
}

function createAxisGeometry(
  sheet: XlsxSheetData,
  axis: "row" | "column",
  zoom: number,
): AxisGeometry {
  let cache = axisGeometryCache.get(sheet);
  if (!cache) {
    cache = new Map();
    axisGeometryCache.set(sheet, cache);
  }
  const cacheKey = `${axis}:${zoom}`;
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const indices = resolveAxisIndices(sheet, axis);
  const sizes = indices.map((actualIndex, displayIndex) =>
    resolveAxisSize(sheet, axis, displayIndex, actualIndex, zoom),
  );
  const offsets = new Array<number>(sizes.length + 1);
  offsets[0] = 0;
  for (let index = 0; index < sizes.length; index += 1) {
    offsets[index + 1] = offsets[index]! + sizes[index]!;
  }

  const freeze = axis === "row" ? sheet.freezePanes?.row ?? 0 : sheet.freezePanes?.col ?? 0;
  const freezeCount = freeze > 0 ? lowerBound(indices, freeze) : 0;
  const frozenExtent = offsets[freezeCount] ?? 0;
  const geometry = { freezeCount, frozenExtent, indices, offsets, sizes };
  cache.set(cacheKey, geometry);
  return geometry;
}

/** Display range covered by a workbook range, skipping hidden rows/columns. */
function resolveDisplayRange(
  axis: AxisGeometry,
  start: number,
  end: number,
): { end: number; start: number } | null {
  if (axis.indices.length === 0) return null;
  const startDisplay = lowerBound(axis.indices, Math.min(start, end));
  const endDisplay = upperBound(axis.indices, Math.max(start, end)) - 1;
  if (startDisplay >= axis.indices.length || endDisplay < startDisplay) return null;
  return { end: endDisplay, start: startDisplay };
}

/** Splits a display span at the freeze line so each fragment owns one pane. */
function splitAtFreeze(
  span: { end: number; start: number },
  freezeCount: number,
): Array<{ end: number; start: number }> {
  if (freezeCount > span.start && freezeCount <= span.end) {
    return [
      { end: freezeCount - 1, start: span.start },
      { end: span.end, start: freezeCount },
    ];
  }
  return [span];
}

function firstOffsetAbove(
  offsets: readonly number[],
  from: number,
  to: number,
  boundary: number,
): number {
  let low = from;
  let high = to + 1;
  while (low < high) {
    const middle = (low + high) >> 1;
    if ((offsets[middle + 1] ?? 0) > boundary) high = middle;
    else low = middle + 1;
  }
  return low;
}

function lastOffsetBelow(
  offsets: readonly number[],
  from: number,
  to: number,
  boundary: number,
): number {
  let low = from - 1;
  let high = to;
  while (low < high) {
    const middle = Math.floor((low + high + 1) / 2);
    if ((offsets[middle] ?? 0) < boundary) low = middle;
    else high = middle - 1;
  }
  return low;
}

/** Clips a display span to a content-space interval; `null` when fully off-screen. */
function clampDisplaySpan(
  span: { end: number; start: number },
  offsets: readonly number[],
  clipStart: number,
  clipEnd: number,
): { end: number; start: number } | null {
  if (!(clipEnd > clipStart)) return null;
  const start = firstOffsetAbove(offsets, span.start, span.end, clipStart);
  if (start > span.end) return null;
  const end = lastOffsetBelow(offsets, start, span.end, clipEnd);
  if (end < start) return null;
  return { end, start };
}

function paneFor(frozenRow: boolean, frozenColumn: boolean): GridPaneKey {
  if (frozenRow && frozenColumn) return "frozen-corner";
  if (frozenRow) return "frozen-rows";
  if (frozenColumn) return "frozen-cols";
  return "main";
}

function isFiniteViewport(viewport: RangeHighlightViewport | undefined): viewport is RangeHighlightViewport {
  return Boolean(viewport)
    && Number.isFinite(viewport!.width)
    && Number.isFinite(viewport!.height)
    && viewport!.width > 0
    && viewport!.height > 0;
}

/**
 * Freeze-pane geometry for the range-highlight overlay.
 *
 * The returned pane boxes and their local coordinates match
 * `createGridPaneLayout`: frozen strips are pinned at the grid origin, the
 * scrolling pane subtracts the scroll offset, and each pane clips its own
 * fragments. Because a range can cross a freeze line, ranges are split by the
 * freeze boundary before placement instead of drawing one oversized box.
 */
export function createRangeHighlightLayout(
  sheet: XlsxSheetData | null,
  zoomScale: number,
  scrollLeft = 0,
  scrollTop = 0,
): RangeHighlightLayout {
  const zoom = gridZoomFactor(zoomScale);
  const paneLayout = createGridPaneLayout(sheet, zoomScale, scrollLeft, scrollTop);
  if (!sheet) {
    return {
      frozenHeight: 0,
      frozenWidth: 0,
      panes: paneLayout.panes,
      resolve: () => [],
    };
  }

  const rows = createAxisGeometry(sheet, "row", zoom);
  const columns = createAxisGeometry(sheet, "column", zoom);
  const { frozenHeight, frozenWidth } = paneLayout;

  function resolve(
    range: XlsxCellRange,
    viewport?: RangeHighlightViewport,
  ): RangeHighlightSegment[] {
    const rowSpan = resolveDisplayRange(rows, range.start.row, range.end.row);
    const columnSpan = resolveDisplayRange(columns, range.start.col, range.end.col);
    if (!rowSpan || !columnSpan) return [];

    const hasViewport = isFiniteViewport(viewport);
    const result: RangeHighlightSegment[] = [];

    for (const rowSegment of splitAtFreeze(rowSpan, rows.freezeCount)) {
      const frozenRow = rows.freezeCount > 0 && rowSegment.start < rows.freezeCount;
      const rowClip = !hasViewport
        ? null
        : frozenRow
          ? { end: Math.min(frozenHeight, viewport!.height), start: 0 }
          : {
              end: scrollTop + viewport!.height,
              start: scrollTop + Math.min(frozenHeight, viewport!.height),
            };
      const clippedRows = rowClip
        ? clampDisplaySpan(rowSegment, rows.offsets, rowClip.start, rowClip.end)
        : rowSegment;
      if (!clippedRows) continue;

      for (const columnSegment of splitAtFreeze(columnSpan, columns.freezeCount)) {
        const frozenColumn = columns.freezeCount > 0 && columnSegment.start < columns.freezeCount;
        const columnClip = !hasViewport
          ? null
          : frozenColumn
            ? { end: Math.min(frozenWidth, viewport!.width), start: 0 }
            : {
                end: scrollLeft + viewport!.width,
                start: scrollLeft + Math.min(frozenWidth, viewport!.width),
              };
        const clippedColumns = columnClip
          ? clampDisplaySpan(columnSegment, columns.offsets, columnClip.start, columnClip.end)
          : columnSegment;
        if (!clippedColumns) continue;

        const left = columns.offsets[clippedColumns.start] ?? 0;
        const top = rows.offsets[clippedRows.start] ?? 0;
        const right = columns.offsets[clippedColumns.end + 1] ?? left;
        const bottom = rows.offsets[clippedRows.end + 1] ?? top;
        result.push({
          pane: paneFor(frozenRow, frozenColumn),
          rect: {
            height: Math.max(0, bottom - top),
            left: left - (frozenColumn ? 0 : frozenWidth + scrollLeft),
            top: top - (frozenRow ? 0 : frozenHeight + scrollTop),
            width: Math.max(0, right - left),
          },
        });
      }
    }

    return result;
  }

  return {
    frozenHeight,
    frozenWidth,
    panes: paneLayout.panes,
    resolve,
  };
}

function transparentColor(color: string, alpha: number): string {
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/iu.exec(color.trim());
  if (hex) {
    const value = hex[1]!.length === 3
      ? hex[1]!.split("").map((part) => part + part).join("")
      : hex[1]!;
    const red = Number.parseInt(value.slice(0, 2), 16);
    const green = Number.parseInt(value.slice(2, 4), 16);
    const blue = Number.parseInt(value.slice(4, 6), 16);
    return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
  }
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/iu.exec(color.trim());
  if (rgb) {
    return `rgba(${rgb[1]}, ${rgb[2]}, ${rgb[3]}, ${alpha})`;
  }
  return color;
}

/** Pane style for a range-highlight box; the parent pane owns clipping. */
export function rangeHighlightStyle(rect: XlsxImageRect, color: string): CSSProperties {
  return {
    backgroundColor: transparentColor(color, 0.18),
    border: `2px solid ${color}`,
    boxSizing: "border-box",
    height: `${rect.height}px`,
    left: `${rect.left}px`,
    pointerEvents: "none",
    position: "absolute",
    top: `${rect.top}px`,
    width: `${rect.width}px`,
  };
}
