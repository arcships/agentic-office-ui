import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";

const ROOT = path.resolve(import.meta.dirname, "../..");
const GRID_PAINT_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/grid-paint.ts");
const FILL_HANDLE_PATH = path.join(
  ROOT,
  "packages/vue-xlsx/src/components/grid-fill-handle.ts",
);

function loadTsModule(relativePath) {
  const output = ts.transpileModule(readFileSync(relativePath, "utf8"), {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
  return import(`data:text/javascript;base64,${Buffer.from(output).toString("base64")}`);
}

const paint = await loadTsModule(GRID_PAINT_PATH);
const fill = await loadTsModule(FILL_HANDLE_PATH);

const {
  AUTOFIT_MAX_COLUMN_WIDTH_PX,
  AUTOFIT_MIN_COLUMN_WIDTH_PX,
  countAutofitLines,
  estimateAutofitTextWidth,
  measureAutofitContent,
  measureAutofitTextWidth,
  resolveAutofitColumnWidth,
  resolveAutofitRowHeight,
} = paint;
const {
  FILL_HANDLE_SIZE_PX,
  fillRangesEqual,
  normalizeFillRange,
  resolveFillHandleBox,
  resolveFillTargetRange,
} = fill;

/** Ten pixels per character, independent of the font, keeps the maths exact. */
const TEN_PER_CHAR = (text) => text.length * 10;

function cell(overrides = {}) {
  return {
    availableWidth: 100,
    font: "normal 400 12px \"Segoe UI\"",
    lineHeight: 15,
    rotation: null,
    value: "",
    wrap: false,
    ...overrides,
  };
}

test("a fill drag grows the selection towards the pointer, one axis at a time", () => {
  const single = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };
  assert.deepEqual(
    resolveFillTargetRange(single, { row: 3, col: 0 }),
    { start: { row: 0, col: 0 }, end: { row: 3, col: 0 } },
    "dragging down from A1 fills A1:A4",
  );
  assert.deepEqual(
    resolveFillTargetRange(single, { row: 0, col: 2 }),
    { start: { row: 0, col: 0 }, end: { row: 0, col: 2 } },
    "dragging right fills A1:C1",
  );
  assert.deepEqual(
    resolveFillTargetRange(single, { row: 2, col: 2 }),
    { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } },
    "dragging diagonally grows both edges",
  );
  assert.deepEqual(
    resolveFillTargetRange(single, { row: 0, col: 0 }),
    single,
    "a pointer on the corner keeps the source range",
  );

  const block = { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } };
  assert.deepEqual(
    resolveFillTargetRange(block, { row: 3, col: 1 }),
    { start: { row: 0, col: 0 }, end: { row: 3, col: 1 } },
    "the pattern keeps its width while the fill grows downwards",
  );
  assert.deepEqual(
    resolveFillTargetRange(block, { row: 1, col: 1 }),
    block,
    "a pointer inside the pattern changes nothing",
  );

  const lower = { start: { row: 4, col: 1 }, end: { row: 5, col: 2 } };
  assert.deepEqual(
    resolveFillTargetRange(lower, { row: 1, col: 2 }),
    { start: { row: 1, col: 1 }, end: { row: 5, col: 2 } },
    "dragging above the pattern fills upwards",
  );
  assert.deepEqual(
    resolveFillTargetRange(lower, { row: 5, col: 0 }),
    { start: { row: 4, col: 0 }, end: { row: 5, col: 2 } },
    "dragging left of the pattern fills leftwards",
  );

  // A reversed selection (shift-clicking upwards) still fills downwards.
  assert.deepEqual(
    resolveFillTargetRange({ start: { row: 2, col: 0 }, end: { row: 0, col: 0 } }, { row: 4, col: 0 }),
    { start: { row: 0, col: 0 }, end: { row: 4, col: 0 } },
  );

  assert.equal(fillRangesEqual(single, { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }), true);
  assert.equal(fillRangesEqual(single, { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }), false);
  assert.deepEqual(
    normalizeFillRange({ start: { row: 5, col: 4 }, end: { row: 1, col: 2 } }),
    { start: { row: 1, col: 2 }, end: { row: 5, col: 4 } },
  );
});

test("the fill handle is a square centred on the selection corner", () => {
  assert.equal(FILL_HANDLE_SIZE_PX, 7);
  assert.deepEqual(resolveFillHandleBox(80, 24), { left: 76.5, top: 20.5 });
  assert.deepEqual(resolveFillHandleBox(10, 10, 4), { left: 8, top: 8 });
  assert.deepEqual(resolveFillHandleBox(Number.NaN, 0), { left: -3.5, top: -3.5 });
});

test("auto-fit measures text the way the paint path draws it", () => {
  assert.equal(measureAutofitTextWidth(cell({ value: "Hello" }), TEN_PER_CHAR), 50);
  assert.equal(measureAutofitTextWidth(cell({ value: "" }), TEN_PER_CHAR), 0);
  assert.equal(
    measureAutofitTextWidth(cell({ value: "short\nmuch longer line" }), TEN_PER_CHAR),
    160,
    "the widest hard line decides the width",
  );
  assert.equal(
    measureAutofitTextWidth(
      cell({
        runs: [
          { font: "bold 12px serif", text: "bold" },
          { font: "italic 20px serif", text: "run" },
        ],
        value: "boldrun",
      }),
      (text, font) => text.length * (font.includes("20px") ? 20 : 10),
    ),
    40 + 60,
    "rich-text runs are measured with their own font",
  );

  assert.equal(
    countAutofitLines(cell({ value: "one\ntwo\nthree" }), TEN_PER_CHAR),
    3,
    "unwrapped cells keep one painted line per hard break",
  );
  assert.equal(
    countAutofitLines(
      cell({ availableWidth: 40, value: "aaaa bbbb cccc", wrap: true }),
      TEN_PER_CHAR,
    ),
    4,
    "a wrapped cell breaks at the same width drawCellText wraps at (greedy, per character)",
  );
  assert.equal(
    countAutofitLines(
      cell({ availableWidth: 1, rotation: -45, value: "aaaa bbbb", wrap: true }),
      TEN_PER_CHAR,
    ),
    1,
    "rotated text never wraps",
  );

  const wrapped = measureAutofitContent(
    cell({ availableWidth: 40, lineHeight: 15, value: "aaaa bbbb cccc", wrap: true }),
    TEN_PER_CHAR,
  );
  assert.deepEqual(
    wrapped,
    { height: 60, width: 140 },
    "the natural width stops the value from wrapping once the column is fitted",
  );
});

test("auto-fit swaps the occupied box for rotated text", () => {
  const horizontal = measureAutofitContent(cell({ lineHeight: 12, value: "abcdefghij" }), TEN_PER_CHAR);
  assert.deepEqual(horizontal, { height: 12, width: 100 });

  const vertical = measureAutofitContent(
    cell({ lineHeight: 12, rotation: -90, value: "abcdefghij" }),
    TEN_PER_CHAR,
  );
  assert.ok(Math.abs(vertical.width - 12) < 1e-9, "90° text is one line high");
  assert.ok(Math.abs(vertical.height - 100) < 1e-9, "90° text is as tall as its text is wide");

  const diagonal = measureAutofitContent(
    cell({ lineHeight: 12, rotation: -45, value: "abcdefghij" }),
    TEN_PER_CHAR,
  );
  const expected = 100 * Math.SQRT1_2 + 12 * Math.SQRT1_2;
  assert.ok(Math.abs(diagonal.width - expected) < 1e-9);
  assert.ok(Math.abs(diagonal.height - expected) < 1e-9);
});

test("auto-fit column width fits the widest cell and stays inside Excel's limits", () => {
  assert.equal(
    resolveAutofitColumnWidth(
      [cell({ value: "Hello" }), cell({ value: "a much wider value" })],
      TEN_PER_CHAR,
    ),
    180 + 10 + 2,
    "widest text + both paddings + rounding room",
  );
  assert.equal(resolveAutofitColumnWidth([], TEN_PER_CHAR), AUTOFIT_MIN_COLUMN_WIDTH_PX);
  assert.equal(
    resolveAutofitColumnWidth([cell({ value: "x" })], TEN_PER_CHAR, { minimumPx: 50 }),
    50,
  );
  assert.equal(
    resolveAutofitColumnWidth(
      [cell({ value: "x".repeat(400) })],
      TEN_PER_CHAR,
    ),
    AUTOFIT_MAX_COLUMN_WIDTH_PX,
  );
  assert.equal(
    resolveAutofitColumnWidth([cell({ value: "Hello world" })], TEN_PER_CHAR, { zoom: 2 }),
    (110 + 10 + 2) / 2,
    "text measured at 200 % zoom fits half the unzoomed width",
  );
});

test("auto-fit row height fits the tallest wrapped cell", () => {
  assert.equal(
    resolveAutofitRowHeight([cell({ lineHeight: 15, value: "one line" })], TEN_PER_CHAR),
    15 + 4,
  );
  assert.equal(
    resolveAutofitRowHeight(
      [
        cell({ availableWidth: 40, lineHeight: 15, value: "aaaa bbbb cccc", wrap: true }),
        cell({ lineHeight: 15, value: "a" }),
      ],
      TEN_PER_CHAR,
    ),
    4 * 15 + 4,
    "the row grows until every wrapped line is visible",
  );
  assert.equal(resolveAutofitRowHeight([], TEN_PER_CHAR), 18);
  assert.equal(
    resolveAutofitRowHeight([cell({ lineHeight: 15, value: "x".repeat(200), rotation: -90 })], TEN_PER_CHAR),
    546,
    "a very long vertical header clamps at Excel's tallest row",
  );
});

test("the DOM-free text estimate still follows the cell's font size", () => {
  assert.equal(estimateAutofitTextWidth("abcd", "normal 400 12px \"Segoe UI\""), 4 * 12 * 0.55);
  assert.equal(estimateAutofitTextWidth("漢字", "normal 400 16px \"Segoe UI\""), 32);
  assert.equal(
    estimateAutofitTextWidth("abcd", "not a font"),
    4 * 12 * 0.55,
    "an unparsable font falls back to the 12 px default",
  );
});
