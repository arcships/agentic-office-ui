import type {
  XlsxBorderPreset,
  XlsxBorderPresetInput,
  XlsxCellAddress,
  XlsxCellBorderEdgeInput,
  XlsxCellBorderStyleInput,
  XlsxCellRange,
  XlsxCellStyleInput
} from "@arcships/xlsx-core";
import { normalizeRange } from "./selection";

/** Line style and colour the ribbon falls back to when the host picks none. */
export const DEFAULT_BORDER_LINE_STYLE = "thin";
export const DEFAULT_BORDER_COLOR_HEX = "A1A1AA";

/** Presets in the order the ribbon lists them. */
export const XLSX_BORDER_PRESETS: readonly XlsxBorderPreset[] = [
  "all",
  "outside",
  "left",
  "right",
  "top",
  "bottom",
  "none"
];

/** One cell's share of a border command. */
export type XlsxSelectionBorderPatch = {
  cell: XlsxCellAddress;
  style: XlsxCellStyleInput;
};

/**
 * Border work for one command.
 *
 * Uniform presets only need the single engine range call; `outside` needs
 * per-cell patches because the inside of the range must stay untouched.
 */
export type XlsxSelectionBorderPlan =
  | { kind: "range"; style: XlsxCellStyleInput }
  | { kind: "cells"; patches: XlsxSelectionBorderPatch[] };

/** Resolve the command's line style and colour, filling in the ribbon defaults. */
export function resolveBorderEdge(input?: XlsxBorderPresetInput): XlsxCellBorderEdgeInput {
  return {
    style: input?.style ?? DEFAULT_BORDER_LINE_STYLE,
    color: input?.color ?? { colorType: "rgb", hex: DEFAULT_BORDER_COLOR_HEX }
  };
}

function buildOutsidePatches(
  range: XlsxCellRange,
  edge: XlsxCellBorderEdgeInput
): XlsxSelectionBorderPatch[] {
  const normalized = normalizeRange(range);
  const patches: XlsxSelectionBorderPatch[] = [];
  for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
    for (let col = normalized.start.col; col <= normalized.end.col; col += 1) {
      const border: XlsxCellBorderStyleInput = {};
      if (row === normalized.start.row) border.top = edge;
      if (row === normalized.end.row) border.bottom = edge;
      if (col === normalized.start.col) border.left = edge;
      if (col === normalized.end.col) border.right = edge;
      // A perimeter cell always owns at least one edge: with a single row or
      // column every cell is on both ends, and a single cell owns all four.
      if (Object.keys(border).length === 0) {
        continue;
      }
      patches.push({ cell: { row, col }, style: { border } });
    }
  }
  return patches;
}

/**
 * Turn a ribbon border preset into the style work for a selection.
 *
 * `all` borders every cell, `left`/`right`/`top`/`bottom` draw that edge on
 * every cell, `outside` walks the perimeter, and `none` clears every edge. The
 * command only writes the edges the preset names, so existing borders on the
 * other edges survive.
 */
export function buildSelectionBorderPlan(
  preset: XlsxBorderPreset,
  range: XlsxCellRange,
  edge: XlsxCellBorderEdgeInput
): XlsxSelectionBorderPlan {
  switch (preset) {
    case "outside":
      return { kind: "cells", patches: buildOutsidePatches(range, edge) };
    case "none": {
      const cleared = { style: "none" } as const;
      return {
        kind: "range",
        style: { border: { top: cleared, right: cleared, bottom: cleared, left: cleared } }
      };
    }
    case "all":
      return { kind: "range", style: { border: { top: edge, right: edge, bottom: edge, left: edge } } };
    case "left":
      return { kind: "range", style: { border: { left: edge } } };
    case "right":
      return { kind: "range", style: { border: { right: edge } } };
    case "top":
      return { kind: "range", style: { border: { top: edge } } };
    case "bottom":
      return { kind: "range", style: { border: { bottom: edge } } };
    default:
      throw new Error(`Unsupported border preset: ${String(preset)}`);
  }
}
