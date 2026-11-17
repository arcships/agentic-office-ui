import type { GridPaneKey } from "./grid-metrics";

/**
 * Rendering model shared by the grid's freeze-pane consumers.
 *
 * With a frozen row/column the grid body has two coordinate spaces: the
 * frozen strips are pinned at the grid origin, while the scrolling pane
 * subtracts the scroll offset. Every pointer hit test, header strip and
 * overlay that used to add `scrollLeft`/`scrollTop` unconditionally must
 * resolve the pane first — otherwise a pinned cell answers for the cell
 * that scrolled underneath it.
 *
 * Two directions are needed:
 * - pointer/box → sheet content (`paneContentOffset`, `resolvePaneContentPoint`)
 * - sheet content → viewport (`paneViewOffset`, `resolvePaneViewPoint`)
 */

/** Sheet-content offset of one axis for a pane-local offset: only scrolling adds scroll. */
export function paneContentOffset(
  localOffset: number,
  scrollOffset: number,
  frozen: boolean,
): number {
  return frozen ? localOffset : localOffset + scrollOffset;
}

/** Pane-local offset of one axis for a sheet-content offset: only scrolling subtracts scroll. */
export function paneViewOffset(
  contentOffset: number,
  scrollOffset: number,
  frozen: boolean,
): number {
  return frozen ? contentOffset : contentOffset - scrollOffset;
}

/** Whether a pointer/box offset sits inside the frozen strip of that axis. */
export function isFrozenPaneOffset(localOffset: number, frozenExtent: number): boolean {
  return frozenExtent > 0 && localOffset < frozenExtent;
}

/** Sheet-content offset of one axis for a grid-body-local offset. */
export function paneAxisContentOffset(
  localOffset: number,
  scrollOffset: number,
  frozenExtent: number,
): number {
  return paneContentOffset(localOffset, scrollOffset, isFrozenPaneOffset(localOffset, frozenExtent));
}

export interface PaneContentPoint {
  /** Pane owning the point. */
  pane: GridPaneKey;
  /** Horizontal sheet-content offset, scroll applied only in the scrolling pane. */
  contentX: number;
  /** Vertical sheet-content offset, scroll applied only in the scrolling pane. */
  contentY: number;
}

/**
 * Maps a grid-body-local point (headers excluded) to sheet content space and
 * reports the pane that owns it. Mirrors `createGridPaneLayout`: the frozen
 * strips keep content coordinates, the live pane adds the scroll offset.
 */
export function resolvePaneContentPoint(
  localX: number,
  localY: number,
  scrollLeft: number,
  scrollTop: number,
  frozenWidth: number,
  frozenHeight: number,
): PaneContentPoint {
  const frozenX = isFrozenPaneOffset(localX, frozenWidth);
  const frozenY = isFrozenPaneOffset(localY, frozenHeight);
  return {
    pane: frozenX && frozenY ? "frozen-corner" : frozenX ? "frozen-cols" : frozenY ? "frozen-rows" : "main",
    contentX: paneContentOffset(localX, scrollLeft, frozenX),
    contentY: paneContentOffset(localY, scrollTop, frozenY),
  };
}

export interface PaneViewPointInput {
  /** Left edge of the cell in sheet-content space. */
  contentX: number;
  /** Top edge of the cell in sheet-content space. */
  contentY: number;
  /** Display index of the cell's column/row (hidden columns and rows included). */
  displayCol: number;
  displayRow: number;
  /** Number of frozen display columns/rows. */
  frozenCols: number;
  frozenRows: number;
  scrollLeft: number;
  scrollTop: number;
}

/**
 * Inverse mapping for viewport-positioned overlays (inline editor, comment
 * popover): a cell inside a frozen strip keeps its content position, a cell in
 * the scrolling pane subtracts the scroll offset.
 */
export function resolvePaneViewPoint(input: PaneViewPointInput): { x: number; y: number } {
  return {
    x: paneViewOffset(input.contentX, input.scrollLeft, input.displayCol < input.frozenCols),
    y: paneViewOffset(input.contentY, input.scrollTop, input.displayRow < input.frozenRows),
  };
}

export interface PaneOverlayInput extends PaneViewPointInput {
  /** Overlay size in pixels; use 0 × 0 for a point anchor such as a popover. */
  width: number;
  height: number;
  /** Pixel size of the frozen column/row strips. */
  frozenWidth: number;
  frozenHeight: number;
  /** Pixel size of the scrolling viewport (grid body, headers excluded). */
  viewportWidth: number;
  viewportHeight: number;
}

export interface PaneOverlayRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * Places a cell-anchored overlay inside the pane that owns the cell.
 *
 * The anchor is mapped like `resolvePaneViewPoint` and the box is then clipped
 * to the pane's rectangle, so an overlay for a cell that is scrolled under a
 * frozen strip (or off the scrolling edge) sticks to the freeze line instead of
 * painting over the pinned cells or the headers. Clipping keeps the box aligned
 * with the still-visible part of the cell: the edge that runs under the freeze
 * line is the one that shrinks, never the opposite edge.
 */
export function resolvePaneOverlayRect(input: PaneOverlayInput): PaneOverlayRect {
  const point = resolvePaneViewPoint(input);
  const frozenWidth = Math.max(0, Math.min(input.frozenWidth, input.viewportWidth));
  const frozenHeight = Math.max(0, Math.min(input.frozenHeight, input.viewportHeight));
  const frozenX = input.displayCol < input.frozenCols;
  const frozenY = input.displayRow < input.frozenRows;
  const paneLeft = frozenX ? 0 : frozenWidth;
  const paneRight = frozenX ? frozenWidth : Math.max(frozenWidth, input.viewportWidth);
  const paneTop = frozenY ? 0 : frozenHeight;
  const paneBottom = frozenY ? frozenHeight : Math.max(frozenHeight, input.viewportHeight);
  const left = Math.max(paneLeft, Math.min(point.x, paneRight));
  const top = Math.max(paneTop, Math.min(point.y, paneBottom));
  const right = Math.max(left, Math.min(point.x + Math.max(0, input.width), paneRight));
  const bottom = Math.max(top, Math.min(point.y + Math.max(0, input.height), paneBottom));
  return { left, top, width: right - left, height: bottom - top };
}
