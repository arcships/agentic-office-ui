import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);

const { XLSX_SELECTION_STATS_CELL_LIMIT, summarizeSelectionValues } = core;

test("summarizeSelectionValues matches Excel's count, sum and average", () => {
  const values = Array.from({ length: 10 }, (_, index) => index + 1);
  const summary = summarizeSelectionValues(values, { totalCells: 10 });

  assert.equal(summary.cellCount, 10);
  assert.equal(summary.count, 10);
  assert.equal(summary.numericCount, 10);
  assert.equal(summary.sum, 55);
  assert.equal(summary.average, 5.5);
  assert.equal(summary.min, 1);
  assert.equal(summary.max, 10);
  assert.equal(summary.truncated, false);
  assert.equal(summary.numberFormat, null);
});

test("text and errors count as cells but not as numbers", () => {
  const summary = summarizeSelectionValues([1, null, 3, null], { totalCells: 4 });

  assert.equal(summary.count, 4, "non-empty text cells count");
  assert.equal(summary.numericCount, 2);
  assert.equal(summary.sum, 4);
  assert.equal(summary.average, 2, "Excel averages over numeric cells only");
  assert.equal(summary.min, 1);
  assert.equal(summary.max, 3);
});

test("an empty selection reports no totals instead of zeroes", () => {
  const summary = summarizeSelectionValues([], { scannedCells: 4, totalCells: 4 });

  assert.equal(summary.cellCount, 4);
  assert.equal(summary.count, 0);
  assert.equal(summary.numericCount, 0);
  assert.equal(summary.sum, null);
  assert.equal(summary.average, null);
  assert.equal(summary.min, null);
  assert.equal(summary.max, null);
  assert.equal(summary.truncated, false);
});

test("a selection larger than the scan budget is marked truncated", () => {
  const summary = summarizeSelectionValues([5, 10], {
    scannedCells: XLSX_SELECTION_STATS_CELL_LIMIT,
    totalCells: XLSX_SELECTION_STATS_CELL_LIMIT + 5_000,
  });

  assert.equal(summary.cellCount, XLSX_SELECTION_STATS_CELL_LIMIT + 5_000);
  assert.equal(summary.count, 2);
  assert.equal(summary.truncated, true);
  assert.equal(XLSX_SELECTION_STATS_CELL_LIMIT, 10_000);
});

test("totals keep the first numeric cell's number format", () => {
  const numberFormat = { formatType: "custom", formatString: "$#,##0.00" };
  const summary = summarizeSelectionValues([1234.56, 2000], {
    numberFormat,
    totalCells: 2,
  });

  assert.deepEqual(summary.numberFormat, numberFormat);
});

test("compensated summation keeps decimal totals exact", () => {
  const summary = summarizeSelectionValues([0.1, 0.2, 0.3], { totalCells: 3 });

  assert.equal(summary.sum, 0.6);
  assert.ok(Math.abs((summary.average ?? 0) - 0.2) < 1e-15);
});

test("non-finite values are ignored like empty formula results", () => {
  const summary = summarizeSelectionValues([1, Number.NaN, Number.POSITIVE_INFINITY], {
    totalCells: 3,
  });

  assert.equal(summary.count, 3, "the cells are non-empty");
  assert.equal(summary.numericCount, 1);
  assert.equal(summary.sum, 1);
  assert.equal(summary.min, 1);
  assert.equal(summary.max, 1);
});
