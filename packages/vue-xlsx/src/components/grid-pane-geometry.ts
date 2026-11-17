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
