// Regression coverage for issue #146 — DOCX table range selection.
//
// The issue bullets map to these tests:
//   1. anchor + focus + physical grid span (gridRange) .... normalization tests
//   2. multi-cell highlight includes merged cells .......... grid membership
//   3. commands/clipboard/clear use the whole rectangle .... component tests in
//      tests/component/docx-table-editing.test.mjs
//   4. Escape collapses to the anchor cell ................. component test

import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromDocxCore.resolve("@arcships/docx-core")).href
);

function cell(text, style = {}) {
  return {
    type: "table-cell",
    style: { widthTwips: 1000, ...style },
    nodes: [{ type: "paragraph", children: [{ type: "text", text }] }],
  };
}

function table(rows, widths = [1000, 1000, 1000]) {
  return {
    type: "table",
    style: {
      layout: "fixed",
      widthTwips: widths.reduce((sum, width) => sum + width, 0),
      columnWidthsTwips: widths,
    },
    rows,
  };
}

function range(overrides = {}) {
  return {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 0,
    focusCellIndex: 0,
    ...overrides,
  };
}

test("normalization stores the physical grid span of a rectangle", () => {
  const model = table([
    { type: "table-row", cells: [cell("A"), cell("B"), cell("C")] },
    { type: "table-row", cells: [cell("D"), cell("E"), cell("F")] },
  ]);

  const normalized = core.normalizeTableCellSelectionRange(
    model,
    range({
      anchorRowIndex: 0,
      anchorCellIndex: 1,
      focusRowIndex: 1,
      focusCellIndex: 2,
    }),
  );
  assert.deepEqual(normalized, {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 1,
    focusRowIndex: 1,
    focusCellIndex: 2,
    gridRange: { startColumnIndex: 1, endColumnIndex: 2 },
  });
  assert.equal(core.tableCellSelectionRangeHasMultipleCells(normalized), true);
  assert.deepEqual(core.tableCellSelectionAnchor(normalized), {
    rowIndex: 0,
    cellIndex: 1,
  });

  assert.deepEqual(
    core.tableCellSelectionRangeCells(model, normalized),
    [
      { rowIndex: 0, cellIndex: 1 },
      { rowIndex: 0, cellIndex: 2 },
      { rowIndex: 1, cellIndex: 1 },
      { rowIndex: 1, cellIndex: 2 },
    ],
  );
  assert.equal(core.isCellWithinTableSelectionRange(model, normalized, 1, 2), true);
  assert.equal(core.isCellWithinTableSelectionRange(model, normalized, 1, 0), false);
});

test("a horizontally merged cell joins a range that touches any of its columns", () => {
  const model = table([
    { type: "table-row", cells: [cell("A", { gridSpan: 2 }), cell("C")] },
    { type: "table-row", cells: [cell("D"), cell("E"), cell("F")] },
  ]);

  const normalized = core.normalizeTableCellSelectionRange(
    model,
    range({
      anchorRowIndex: 0,
      anchorCellIndex: 0,
      focusRowIndex: 0,
      focusCellIndex: 1,
    }),
  );
  assert.deepEqual(normalized.gridRange, {
    startColumnIndex: 0,
    endColumnIndex: 2,
  });
  assert.equal(
    core.isCellWithinTableSelectionRange(model, normalized, 0, 0),
    true,
    "the spanning cell is highlighted as a whole"
  );
  assert.deepEqual(
    core.tableCellSelectionRangeCells(model, normalized),
    [
      { rowIndex: 0, cellIndex: 0 },
      { rowIndex: 0, cellIndex: 1 },
    ]
  );

  // Selecting only the second physical column of the merged cell still
  // resolves the rectangle to the merged cell's full span.
  const columnOnly = core.normalizeTableCellSelectionRange(
    model,
    range({
      anchorRowIndex: 0,
      anchorCellIndex: 0,
      focusRowIndex: 0,
      focusCellIndex: 0,
      gridRange: { startColumnIndex: 1, endColumnIndex: 1 },
    }),
  );
  assert.deepEqual(columnOnly.gridRange, {
    startColumnIndex: 0,
    endColumnIndex: 1,
  });
  assert.equal(core.isCellWithinTableSelectionRange(model, columnOnly, 0, 0), true);
});

test("vertical merges expand the range and keep merged cells in the set", () => {
  const model = table([
    {
      type: "table-row",
      cells: [cell("A", { rowSpan: 2 }), cell("B")],
    },
    {
      type: "table-row",
      cells: [cell("", { vMergeContinuation: true }), cell("C")],
    },
  ], [1000, 1000]);

  const normalized = core.normalizeTableCellSelectionRange(
    model,
    range({
      anchorRowIndex: 1,
      anchorCellIndex: 0,
      focusRowIndex: 1,
      focusCellIndex: 0,
    }),
  );
  assert.deepEqual(normalized.gridRange, {
    startColumnIndex: 0,
    endColumnIndex: 0,
  });
  assert.deepEqual(
    core.tableCellSelectionRangeCells(model, normalized),
    [
      { rowIndex: 0, cellIndex: 0 },
      { rowIndex: 1, cellIndex: 0 },
    ],
    "the vertical merge pulls in the restart row and continuation cell"
  );
  assert.equal(core.isCellWithinTableSelectionRange(model, normalized, 0, 0), true);
  assert.equal(core.isCellWithinTableSelectionRange(model, normalized, 0, 1), false);
});

test("isCellSelected handles single cells and range selections", () => {
  const model = table([
    { type: "table-row", cells: [cell("A"), cell("B")] },
    { type: "table-row", cells: [cell("C"), cell("D")] },
  ], [1000, 1000]);
  const normalized = core.normalizeTableCellSelectionRange(
    model,
    range({
      anchorRowIndex: 0,
      anchorCellIndex: 0,
      focusRowIndex: 1,
      focusCellIndex: 1,
    }),
  );

  assert.equal(
    core.isCellSelected(
      {
        kind: "table-cell",
        tableIndex: 0,
        rowIndex: 1,
        cellIndex: 1,
        cellRange: normalized,
      },
      0,
      0,
      0,
      model
    ),
    true
  );
  assert.equal(
    core.isCellSelected(
      { kind: "table-cell", tableIndex: 0, rowIndex: 1, cellIndex: 1 },
      0,
      0,
      0,
      model
    ),
    false
  );
  assert.equal(
    core.isCellSelected(
      { kind: "table-cell", tableIndex: 0, rowIndex: 0, cellIndex: 1 },
      0,
      0,
      1,
      model
    ),
    true
  );
});

test("invalid range endpoints and empty tables resolve to no range", () => {
  const model = table([{ type: "table-row", cells: [cell("A")] }], [1000]);
  assert.equal(
    core.normalizeTableCellSelectionRange(
      model,
      range({ focusRowIndex: 4, focusCellIndex: 0 })
    ),
    undefined
  );
  assert.equal(
    core.normalizeTableCellSelectionRange(
      table([]),
      range({ anchorRowIndex: 0, focusRowIndex: 0 })
    ),
    undefined
  );
  assert.deepEqual(
    core.tableCellSelectionRangeCells(
      model,
      range({ focusRowIndex: 2, focusCellIndex: 2 })
    ),
    []
  );
  assert.equal(
    core.isCellWithinTableSelectionRange(
      model,
      range({ focusRowIndex: 2, focusCellIndex: 2 }),
      0,
      0
    ),
    false
  );
});
