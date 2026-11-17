import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueXlsx = createRequire(
  new URL("../../packages/vue-xlsx/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromVueXlsx.resolve("@arcships/xlsx-core")).href
);
const { anchorToAbsoluteRect, emuToPixels } = core;

const PX_TO_EMU = 9525;

function absoluteRect(anchor, state) {
  const rect = anchorToAbsoluteRect(anchor, state);
  return {
    x: emuToPixels(rect.x),
    y: emuToPixels(rect.y),
    width: emuToPixels(rect.cx),
    height: emuToPixels(rect.cy),
  };
}

const stateWithHiddenColumn = {
  colWidthOverridesPx: { 0: 80, 1: 200, 2: 80, 3: 80 },
  rowHeightOverridesPx: { 0: 24, 1: 100, 2: 24 },
  defaultColWidthPx: 80,
  defaultRowHeightPx: 24,
  hiddenCols: [1],
  hiddenRows: [1],
  visibleCols: [0, 2, 3],
  visibleRows: [0, 2],
  colWidths: [80, 80, 80],
  rowHeights: [24, 24],
};

test("anchored objects skip hidden columns and rows when accumulating geometry", () => {
  const anchor = {
    kind: "one-cell",
    from: { col: 3, row: 2, colOffsetEmu: 0, rowOffsetEmu: 0 },
    sizeEmu: { cx: 10 * PX_TO_EMU, cy: 10 * PX_TO_EMU },
  };
  assert.deepEqual(
    absoluteRect(anchor, stateWithHiddenColumn),
    { height: 10, width: 10, x: 160, y: 24 },
    "hidden column 1 (200px) and hidden row 1 (100px) must not shift the anchor",
  );

  const twoCell = {
    kind: "two-cell",
    from: { col: 0, row: 0, colOffsetEmu: 0, rowOffsetEmu: 0 },
    to: { col: 3, row: 2, colOffsetEmu: 0, rowOffsetEmu: 0 },
  };
  assert.deepEqual(
    absoluteRect(twoCell, stateWithHiddenColumn),
    { height: 24, width: 160, x: 0, y: 0 },
    "two-cell anchors must span only the painted (visible) columns / rows",
  );
});

test("anchored objects prefer the grid's display-indexed widths and heights", () => {
  const state = {
    colWidthOverridesPx: { 0: 50, 1: 50, 2: 50 },
    rowHeightOverridesPx: { 0: 10, 1: 10, 2: 10 },
    defaultColWidthPx: 50,
    defaultRowHeightPx: 10,
    colWidths: [120, 120],
    rowHeights: [30, 30],
    visibleCols: [0, 1],
    visibleRows: [0, 1],
  };
  const anchor = {
    kind: "one-cell",
    from: { col: 2, row: 2, colOffsetEmu: 0, rowOffsetEmu: 0 },
    sizeEmu: { cx: PX_TO_EMU, cy: PX_TO_EMU },
  };
  assert.deepEqual(
    absoluteRect(anchor, state),
    { height: 1, width: 1, x: 240, y: 60 },
    "display-indexed sizes for the columns / rows before the anchor are used",
  );
});

test("hidden slots without visible lists still collapse to zero width", () => {
  const state = {
    colWidthOverridesPx: { 0: 80, 1: 200, 2: 80, 3: 80 },
    rowHeightOverridesPx: { 0: 24, 1: 100, 2: 24 },
    defaultColWidthPx: 80,
    defaultRowHeightPx: 24,
    hiddenCols: [1],
    hiddenRows: [1],
  };
  const anchor = {
    kind: "one-cell",
    from: { col: 3, row: 2, colOffsetEmu: 0, rowOffsetEmu: 0 },
    sizeEmu: { cx: PX_TO_EMU, cy: PX_TO_EMU },
  };
  assert.deepEqual(
    absoluteRect(anchor, state),
    { height: 1, width: 1, x: 160, y: 24 },
  );
});

test("absolute anchors and sheets without hidden metadata keep their old geometry", () => {
  const absolute = {
    kind: "absolute",
    positionEmu: { x: 30 * PX_TO_EMU, y: 40 * PX_TO_EMU },
    sizeEmu: { cx: 50 * PX_TO_EMU, cy: 60 * PX_TO_EMU },
  };
  assert.deepEqual(
    absoluteRect(absolute, stateWithHiddenColumn),
    { height: 60, width: 50, x: 30, y: 40 },
  );

  const legacy = {
    colWidthOverridesPx: { 0: 80, 1: 200, 2: 80, 3: 80 },
    rowHeightOverridesPx: { 0: 24, 1: 100, 2: 24 },
    defaultColWidthPx: 80,
    defaultRowHeightPx: 24,
  };
  const anchor = {
    kind: "one-cell",
    from: { col: 3, row: 2, colOffsetEmu: 0, rowOffsetEmu: 0 },
    sizeEmu: { cx: PX_TO_EMU, cy: PX_TO_EMU },
  };
  assert.deepEqual(
    absoluteRect(anchor, legacy),
    { height: 1, width: 1, x: 360, y: 124 },
  );
});
