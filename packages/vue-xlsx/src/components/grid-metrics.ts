import type { CSSProperties } from "vue";
import type { XlsxImageAnchor, XlsxImageRect, XlsxSheetData } from "@arcships/xlsx-core";
import { createAnchoredRectCache } from "@arcships/xlsx-core";
import { GRID_HEADER_HEIGHT, GRID_ROW_HEADER_WIDTH } from "../composables/internal";

/** Height of the grid's column-header strip, in CSS pixels. */
export const GRID_HEADER_HEIGHT_PX = GRID_HEADER_HEIGHT;
/** Width of the grid's row-header gutter, in CSS pixels. */
export const GRID_ROW_HEADER_WIDTH_PX = GRID_ROW_HEADER_WIDTH;
/** Grid defaults, matching the fallbacks XlsxGrid paints with. */
const GRID_DEFAULT_COL_WIDTH_PX = 80;
const GRID_DEFAULT_ROW_HEIGHT_PX = 24;

/**
 * Absolute box covering the grid body only (headers excluded), used by the
 * chart / image / drawing overlays. `overflow: hidden` on the layer keeps
 * anchored objects from painting over the headers or escaping the surface
 * into the surrounding UI while the grid is scrolled.
 */
export const gridBodyLayerStyle: CSSProperties = {
  height: `calc(100% - ${GRID_HEADER_HEIGHT_PX}px)`,
  left: `${GRID_ROW_HEADER_WIDTH_PX}px`,
  top: `${GRID_HEADER_HEIGHT_PX}px`,
  width: `calc(100% - ${GRID_ROW_HEADER_WIDTH_PX}px)`,
};

/** Zoom factor derived from the controller's percentage scale. */
export function gridZoomFactor(zoomScale: number): number {
  const scale = zoomScale / 100;
  return Number.isFinite(scale) ? Math.max(0.1, scale) : 1;
}

/** Pixel extent of the frozen columns or rows at the current zoom. */
export function frozenExtent(
  sheet: XlsxSheetData | null,
  axis: "column" | "row",
  zoomScale: number,
): number {
  if (!sheet?.freezePanes) return 0;
  const limit = axis === "column" ? sheet.freezePanes.col : sheet.freezePanes.row;
  const indices = axis === "column" ? sheet.visibleCols : sheet.visibleRows;
  const sizes = axis === "column" ? sheet.colWidths : sheet.rowHeights;
  const overrides = axis === "column" ? sheet.colWidthOverridesPx : sheet.rowHeightOverridesPx;
  const fallback = axis === "column" ? sheet.defaultColWidthPx : sheet.defaultRowHeightPx;
  const hardFallback = axis === "column" ? GRID_DEFAULT_COL_WIDTH_PX : GRID_DEFAULT_ROW_HEIGHT_PX;
  const zoom = gridZoomFactor(zoomScale);
  return indices.reduce((total, actualIndex, displayIndex) =>
    // Resolve sizes the way XlsxGrid does (resized columns live in the
    // override maps), otherwise overlay panes drift from the grid at the
    // freeze line.
    actualIndex < limit
      ? total + (sizes[displayIndex] || overrides?.[actualIndex] || fallback || hardFallback) * zoom
      : total,
  0);
}

const anchorRectCache = createAnchoredRectCache();

/** Content-space rect (grid-body coordinates, zoom applied) of an anchored object. */
export function resolveAnchorRect(
  anchor: XlsxImageAnchor,
  sheet: XlsxSheetData | null,
  zoomScale: number,
): XlsxImageRect {
  // Anchors are stable objects across scroll frames, so reuse their content
  // rect instead of walking the axis model on every scroll tick.
  return anchorRectCache.resolve(anchor, sheet, zoomScale);
}

export type GridPaneKey = "main" | "frozen-rows" | "frozen-cols" | "frozen-corner";

export interface GridPane {
  key: GridPaneKey;
  style: CSSProperties;
}

export interface GridPanePlacement {
  pane: GridPaneKey;
  style: CSSProperties;
}

export interface GridPaneLayout {
  panes: GridPane[];
  place: (rect: XlsxImageRect) => GridPanePlacement;
  frozenWidth: number;
  frozenHeight: number;
  /**
   * Whether an object in `pane` intersects the visible viewport (plus an
   * overscan margin). Used to cull anchored objects that scrolled away.
   */
  isVisible: (
    rect: XlsxImageRect,
    pane: GridPaneKey,
    viewportWidth: number,
    viewportHeight: number,
    overscan?: number,
  ) => boolean;
}

/**
 * Freeze-pane layout for anchored-object layers. `panes` are absolute,
 * `overflow: hidden` sub-boxes of the grid-body layer — Excel clips drawings
 * per pane — with frozen panes stacked above the main pane, mirroring how the
 * canvas paints frozen cells over the scrolling region. `place` assigns an
 * object to the pane owning its anchor origin and positions it relative to
 * that pane's box; only the scrolling axes subtract scroll.
 */
export function createGridPaneLayout(
  sheet: XlsxSheetData | null,
  zoomScale: number,
  scrollLeft: number,
  scrollTop: number,
): GridPaneLayout {
  const frozenWidth = frozenExtent(sheet, "column", zoomScale);
  const frozenHeight = frozenExtent(sheet, "row", zoomScale);
  const scrollingWidth = frozenWidth ? `calc(100% - ${frozenWidth}px)` : "100%";
  const scrollingHeight = frozenHeight ? `calc(100% - ${frozenHeight}px)` : "100%";

  const panes: GridPane[] = [
    { key: "main", style: { height: scrollingHeight, left: `${frozenWidth}px`, top: `${frozenHeight}px`, width: scrollingWidth, zIndex: 1 } },
  ];
  if (frozenHeight > 0) {
    panes.push({ key: "frozen-rows", style: { height: `${frozenHeight}px`, left: `${frozenWidth}px`, top: "0px", width: scrollingWidth, zIndex: 2 } });
  }
  if (frozenWidth > 0) {
    panes.push({ key: "frozen-cols", style: { height: scrollingHeight, left: "0px", top: `${frozenHeight}px`, width: `${frozenWidth}px`, zIndex: 2 } });
  }
  if (frozenWidth > 0 && frozenHeight > 0) {
    panes.push({ key: "frozen-corner", style: { height: `${frozenHeight}px`, left: "0px", top: "0px", width: `${frozenWidth}px`, zIndex: 3 } });
  }

  const place = (rect: XlsxImageRect): GridPanePlacement => {
    const colFrozen = rect.left < frozenWidth;
    const rowFrozen = rect.top < frozenHeight;
    const pane: GridPaneKey =
      colFrozen && rowFrozen ? "frozen-corner"
      : rowFrozen ? "frozen-rows"
      : colFrozen ? "frozen-cols"
      : "main";
    return {
      pane,
      style: {
        height: `${rect.height}px`,
        left: `${rect.left - (colFrozen ? 0 : frozenWidth + scrollLeft)}px`,
        top: `${rect.top - (rowFrozen ? 0 : frozenHeight + scrollTop)}px`,
        width: `${rect.width}px`,
      },
    };
  };

  const isVisible = (
    rect: XlsxImageRect,
    pane: GridPaneKey,
    viewportWidth: number,
    viewportHeight: number,
    overscan = 64,
  ) => {
    if (!(viewportWidth > 0) || !(viewportHeight > 0)) return true;
    let left = 0;
    let top = 0;
    let right = viewportWidth;
    let bottom = viewportHeight;

    if (pane === "main" || pane === "frozen-rows") {
      left = scrollLeft;
      right = scrollLeft + Math.max(0, viewportWidth - frozenWidth);
    }
    if (pane === "main" || pane === "frozen-cols") {
      top = scrollTop;
      bottom = scrollTop + Math.max(0, viewportHeight - frozenHeight);
    }
    if (pane === "frozen-rows" || pane === "frozen-corner") {
      top = 0;
      bottom = frozenHeight;
    }
    if (pane === "frozen-cols" || pane === "frozen-corner") {
      left = 0;
      right = frozenWidth;
    }

    return (
      rect.left + rect.width >= left - overscan
      && rect.left <= right + overscan
      && rect.top + rect.height >= top - overscan
      && rect.top <= bottom + overscan
    );
  };

  return { panes, place, frozenWidth, frozenHeight, isVisible };
}
