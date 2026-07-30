import type { CSSProperties } from "vue";
import type { XlsxImageAnchor, XlsxImageRect, XlsxSheetData } from "@arcships/xlsx-core";
import { anchorToAbsoluteRect, emuToPixels } from "@arcships/xlsx-core";
import { GRID_HEADER_HEIGHT, GRID_ROW_HEADER_WIDTH } from "../composables/internal";

/** Height of the grid's column-header strip, in CSS pixels. */
export const GRID_HEADER_HEIGHT_PX = GRID_HEADER_HEIGHT;
/** Width of the grid's row-header gutter, in CSS pixels. */
export const GRID_ROW_HEADER_WIDTH_PX = GRID_ROW_HEADER_WIDTH;

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
  const fallback = axis === "column" ? sheet.defaultColWidthPx : sheet.defaultRowHeightPx;
  const zoom = gridZoomFactor(zoomScale);
  return indices.reduce((total, actualIndex, displayIndex) =>
    actualIndex < limit ? total + (sizes[displayIndex] || fallback) * zoom : total,
  0);
}

/** Content-space rect (grid-body coordinates, zoom applied) of an anchored object. */
export function resolveAnchorRect(
  anchor: XlsxImageAnchor,
  sheet: XlsxSheetData | null,
  zoomScale: number,
): XlsxImageRect {
  const rect = anchorToAbsoluteRect(anchor, sheet);
  const zoom = gridZoomFactor(zoomScale);
  return {
    left: emuToPixels(rect.x) * zoom,
    top: emuToPixels(rect.y) * zoom,
    width: Math.max(1, emuToPixels(rect.cx) * zoom),
    height: Math.max(1, emuToPixels(rect.cy) * zoom),
  };
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

  return { panes, place };
}
