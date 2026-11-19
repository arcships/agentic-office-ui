import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";
import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { useXlsxViewerController, XlsxFormulaBar } = await importFromDemo(
  "@arcships/vue-xlsx",
);

const ROOT = path.resolve(import.meta.dirname, "../..");
const GRID_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxGrid.vue");
const GRID_SOURCE = readFileSync(GRID_PATH, "utf8");
const GRID_SCRIPT = [...GRID_SOURCE.matchAll(/<script(?:\s+setup)?[^>]*>([\s\S]*?)<\/script>/g)]
  .map((match) => match[1])
  .join("\n");
const GRID_AST = ts.createSourceFile(
  GRID_PATH,
  GRID_SCRIPT,
  ts.ScriptTarget.Latest,
  true,
  ts.ScriptKind.TS,
);

const FORMULA_CELL = { row: 0, col: 0 };

function configureWasm() {
  const bytes = readFileSync(
    new URL(
      "../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm",
      import.meta.url,
    ),
  );
  setWasmSource(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  );
}

function workbookBuffer() {
  const bytes = readFileSync(
    new URL("../../apps/demo/public/samples/sales-table.xlsx", import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function mountController() {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(),
        fileName: "cell-editing.xlsx",
        skipXmlParsing: true,
        useWorker: false,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "editable XLSX controller",
  );
  assert.ifError(controller.error);
  return { ...mounted, controller };
}

function descendants(node, predicate) {
  const matches = [];
  function visit(current) {
    if (predicate(current)) matches.push(current);
    ts.forEachChild(current, visit);
  }
  visit(node);
  return matches;
}

function functionNode(name) {
  return descendants(
    GRID_AST,
    (node) => ts.isFunctionDeclaration(node) && node.name?.text === name,
  )[0];
}

// XlsxGrid renders its inline editor over a canvas surface that the test
// renderer does not provide, so mount the real edit functions extracted from
// the component source against a real controller. This is the same technique
// xlsx-grid-scroll.test.mjs uses for grid internals.
async function loadInlineEditor(controller) {
  const required = [
    "onDblClick",
    "commitEdit",
    "commitEditFillingSelection",
    "startEdit",
    "commitAndTab",
    "commitEditingDraft",
    "hasRangeSelection",
  ];
  const bodies = required.map((name) => {
    const node = functionNode(name);
    assert.ok(node, `missing XlsxGrid.${name}`);
    return node.getText(GRID_AST);
  });
  // Older grids seeded and committed through the display value directly. Fall
  // back to that shape when the helper is missing so this test stays
  // behavioral: it must fail on the wrong behavior, not just on the refactor.
  const inputValueHelper = functionNode("getCellInputValue");
  if (inputValueHelper) bodies.push(inputValueHelper.getText(GRID_AST));
  const prelude = `
    const props = { controller: null, readOnly: false };
    const editingCell = { value: null };
    const editingValue = { value: "" };
    // The grid derives both from its props; the getters keep this harness in
    // step with the controller the test swaps in after loading.
    const selection = { get value() { return props.controller?.selection ?? null; } };
    const editingLocked = {
      get value() { return props.readOnly === true || props.controller?.isEditingLocked === true; },
    };
    // Formula-reference highlighting is exercised by its own component test;
    // this harness only needs the editing hooks to exist and stay the same
    // no-op places the real session would end/update.
    const editorCaret = { value: 0 };
    const FORMULA_REFERENCE_OWNER_GRID = "grid";
    const formulaReferenceEditing = {
      begin() {},
      update() {},
      end() {},
    };
    const isEditComposing = { value: false };
    let commitEditOnCompositionEnd = false;
    const editInputRef = { value: { focus() {}, select() {} } };
    const containerRef = { value: { focus() {} } };
    const emit = () => {};
    const nextTick = (task) => { task(); };
    let hitTestResult = null;
    function hitTestCell() { return hitTestResult; }
    function setHitTestResult(cell) { hitTestResult = cell; }
    // A double-click on a column/row boundary auto-fits that axis before the
    // cell editor is considered; this harness only drives the cell path.
    function hitTestAxis() { return null; }
    function autofitAxis() { return false; }
    let adjacentCell = null;
    function getAdjacentCell(cell) { return adjacentCell ?? { row: cell.row, col: cell.col + 1 }; }
    function setAdjacentCell(cell) { adjacentCell = cell; }
    // The editor brings its cell into the pane through the grid's scroll
    // surface, which this harness does not provide; the assertions below cover
    // seeding and committing the formula, not scrolling.
    function scrollToCell() {}
    ${inputValueHelper ? "" : "function getCellInputValue(cell) { return props.controller.getCellDisplayValue(cell); }"}
  `;
  const output = ts.transpileModule(
    `${prelude}\n${bodies.join("\n")}\nexport { props, editingCell, editingValue, onDblClick, commitEdit, commitEditFillingSelection, startEdit, commitAndTab, setHitTestResult, setAdjacentCell };`,
    {
      compilerOptions: {
        module: ts.ModuleKind.ESNext,
        target: ts.ScriptTarget.ES2022,
      },
    },
  ).outputText;
  const editor = await import(
    `data:text/javascript;base64,${Buffer.from(output).toString("base64")}`
  );
  editor.props.controller = controller;
  return editor;
}

async function seedFormulaCell(controller) {
  controller.setCellValue({ row: 0, col: 1 }, "2");
  controller.setCellValue({ row: 0, col: 2 }, "3");
  controller.setCellFormula(FORMULA_CELL, "=SUM(B1:C1)");
  controller.selectCell(FORMULA_CELL);
  await vue.nextTick();
  assert.equal(controller.getCellFormula(FORMULA_CELL), "=SUM(B1:C1)");
  assert.equal(controller.getCellDisplayValue(FORMULA_CELL), "5");
}

// Counts undo entries without leaving the workbook in an unwound state.
async function countUndoDepth(controller) {
  let count = 0;
  while (controller.canUndo && count < 100) {
    controller.undo();
    count += 1;
    await vue.nextTick();
  }
  while (controller.canRedo) {
    controller.redo();
    await vue.nextTick();
  }
  return count;
}

// Reads the workbook value behind a cell, including merged secondary cells
// that the display APIs hide.
function readRawCell(controller, cell) {
  const worksheet = controller.getActiveWorksheet();
  const value = worksheet.getCellAt(cell.row, cell.col);
  try {
    return value.toJs() ?? "";
  } finally {
    value.free();
  }
}

configureWasm();

test("inline editor seeds the formula and commits it back unchanged on blur", async () => {
  const { controller, app } = await mountController();
  try {
    await seedFormulaCell(controller);
    const editor = await loadInlineEditor(controller);
    const undoDepth = await countUndoDepth(controller);

    editor.setHitTestResult(FORMULA_CELL);
    editor.onDblClick({ clientX: 0, clientY: 0 });
    assert.equal(editor.editingCell.value?.row, FORMULA_CELL.row);
    assert.equal(
      editor.editingValue.value,
      "=SUM(B1:C1)",
      "the editor must be seeded with the formula, not the displayed value",
    );

    // Double-clicking and clicking away commits the untouched draft.
    editor.commitEdit();
    await vue.nextTick();
    assert.equal(controller.getCellFormula(FORMULA_CELL), "=SUM(B1:C1)");
    assert.equal(controller.getCellDisplayValue(FORMULA_CELL), "5");
    assert.equal(
      await countUndoDepth(controller),
      undoDepth,
      "committing an unchanged draft must not record undo history",
    );
  } finally {
    app.unmount();
  }
});

test("inline editing stores hard line breaks without flattening the cell", async () => {
  const { controller, app } = await mountController();
  try {
    const cell = { row: 0, col: 1 };
    const original = controller.getCellDisplayValue(cell);

    assert.equal(controller.commitCellInput(cell, "first line\nsecond line"), true);
    await vue.nextTick();
    assert.equal(
      controller.getCellDisplayValue(cell),
      "first line\nsecond line",
      "Alt+Enter's line break survives the commit",
    );

    controller.undo();
    await vue.nextTick();
    assert.equal(controller.getCellDisplayValue(cell), original, "one undo restores the previous cell");
  } finally {
    app.unmount();
  }
});

test("inline editor fills a multi-cell selection as one undo step", async () => {
  const { controller, app } = await mountController();
  try {
    const editor = await loadInlineEditor(controller);
    const cells = [
      { row: 30, col: 8 },
      { row: 30, col: 9 },
      { row: 31, col: 8 },
      { row: 31, col: 9 },
    ];
    const originals = cells.map((cell) => controller.getCellDisplayValue(cell));
    const undoDepth = await countUndoDepth(controller);

    controller.selectRange({ start: cells[0], end: cells[3] });
    await vue.nextTick();
    editor.startEdit(controller.activeCell);
    editor.editingValue.value = "filled";
    editor.commitEditFillingSelection();
    await vue.nextTick();
    assert.deepEqual(
      cells.map((cell) => controller.getCellDisplayValue(cell)),
      ["filled", "filled", "filled", "filled"],
      "an entry confirmed over a range lands in every cell",
    );
    assert.deepEqual(
      { start: controller.selection.start, end: controller.selection.end },
      { start: cells[0], end: cells[3] },
      "the filled range stays selected",
    );
    assert.equal(
      await countUndoDepth(controller),
      undoDepth + 1,
      "the whole fill is a single undo entry",
    );
    controller.undo();
    await vue.nextTick();
    assert.deepEqual(
      cells.map((cell) => controller.getCellDisplayValue(cell)),
      originals,
      "one undo restores every filled cell",
    );

    const formulaCells = [{ row: 30, col: 10 }, { row: 31, col: 10 }];
    controller.selectRange({ start: formulaCells[0], end: formulaCells[1] });
    await vue.nextTick();
    editor.startEdit(controller.activeCell);
    editor.editingValue.value = "=SUM(A1:A2)";
    editor.commitEditFillingSelection();
    await vue.nextTick();
    assert.equal(controller.getCellFormula(formulaCells[0]), "=SUM(A1:A2)");
    assert.equal(
      controller.getCellFormula(formulaCells[1]),
      "=SUM(A1:A2)",
      "a filled formula keeps its references instead of shifting per cell",
    );
  } finally {
    app.unmount();
  }
});

test("inline editor keeps formulas through F2/Enter, in-place formula edits and Tab", async () => {
  const { controller, app } = await mountController();
  try {
    await seedFormulaCell(controller);
    const editor = await loadInlineEditor(controller);

    // F2 / Enter entry point.
    editor.startEdit(FORMULA_CELL);
    assert.equal(editor.editingValue.value, "=SUM(B1:C1)");
    editor.commitEdit();
    await vue.nextTick();
    assert.equal(controller.getCellFormula(FORMULA_CELL), "=SUM(B1:C1)");

    // Editing the formula in place keeps the cell a formula.
    editor.startEdit(FORMULA_CELL);
    editor.editingValue.value = "=SUM(B1:C1)*2";
    editor.commitEdit();
    await vue.nextTick();
    assert.equal(controller.getCellFormula(FORMULA_CELL), "=SUM(B1:C1)*2");
    assert.equal(controller.getCellDisplayValue(FORMULA_CELL), "10");

    // Tab commits the formula and seeds the next cell with its own content.
    editor.setAdjacentCell({ row: 0, col: 1 });
    editor.startEdit(FORMULA_CELL);
    editor.commitAndTab();
    await vue.nextTick();
    assert.equal(controller.getCellFormula(FORMULA_CELL), "=SUM(B1:C1)*2");
    assert.equal(editor.editingCell.value?.col, 1);
    assert.equal(editor.editingValue.value, "2");

    // Typing a value over a formula cell replaces the formula, and undo
    // restores it.
    editor.setAdjacentCell(null);
    editor.startEdit(FORMULA_CELL);
    editor.editingValue.value = "42";
    editor.commitEdit();
    await vue.nextTick();
    assert.equal(controller.getCellFormula(FORMULA_CELL), "");
    assert.equal(controller.getCellDisplayValue(FORMULA_CELL), "42");

    controller.undo();
    await vue.nextTick();
    assert.equal(controller.getCellFormula(FORMULA_CELL), "=SUM(B1:C1)*2");
    assert.equal(controller.getCellDisplayValue(FORMULA_CELL), "10");
  } finally {
    app.unmount();
  }
});

test("formula bar writes plain text as values and keeps formula entry intact", async () => {
  const { controller, app } = await mountController();
  try {
    await seedFormulaCell(controller);
    const mounted = await mount(XlsxFormulaBar, { controller });
    try {
      const input = findByTestId(mounted.root, "xlsx-formula-input");
      assert.ok(input, "formula bar input is rendered");
      assert.equal(input.props.value, "=SUM(B1:C1)");
      input.blur = () => {};

      // Typing a plain number must take the value path, not the formula path.
      input.props.onFocus();
      input.props.onInput({ target: { value: "42" } });
      input.props.onKeydown({ key: "Enter", preventDefault() {} });
      await vue.nextTick();
      assert.equal(controller.getCellFormula(FORMULA_CELL), "");
      assert.equal(controller.getCellDisplayValue(FORMULA_CELL), "42");
      assert.equal(controller.canRedo, false);

      // Undo restores the formula that the value replaced.
      controller.undo();
      await vue.nextTick();
      assert.equal(controller.getCellFormula(FORMULA_CELL), "=SUM(B1:C1)");
      assert.equal(controller.getCellDisplayValue(FORMULA_CELL), "5");

      // Blurring an untouched draft neither rewrites the cell nor clears redo.
      assert.equal(controller.canRedo, true);
      input.props.onFocus();
      input.props.onBlur();
      await vue.nextTick();
      assert.equal(controller.getCellFormula(FORMULA_CELL), "=SUM(B1:C1)");
      assert.equal(controller.canRedo, true);

      // Typing a formula string still writes a formula.
      controller.selectCell(FORMULA_CELL);
      await vue.nextTick();
      input.props.onFocus();
      input.props.onInput({ target: { value: "=SUM(B1:C1)+1" } });
      input.props.onKeydown({ key: "Enter", preventDefault() {} });
      await vue.nextTick();
      assert.equal(controller.getCellFormula(FORMULA_CELL), "=SUM(B1:C1)+1");
      assert.equal(controller.getCellDisplayValue(FORMULA_CELL), "6");
    } finally {
      mounted.app.unmount();
    }
  } finally {
    app.unmount();
  }
});

test("fill shifts relative references and keeps absolute ones", async () => {
  const { controller, app } = await mountController();
  try {
    // The issue's case: =B1 filled down one row becomes =B2.
    controller.setCellFormula({ row: 0, col: 0 }, "=B1");
    controller.selectCell({ row: 0, col: 0 });
    await vue.nextTick();
    controller.fillSelection({ start: { row: 0, col: 0 }, end: { row: 2, col: 0 } });
    await vue.nextTick();
    assert.deepEqual(
      [0, 1, 2].map((row) => controller.getCellFormula({ row, col: 0 })),
      ["=B1", "=B2", "=B3"],
    );

    // Absolute parts stay anchored; mixed anchors only shift the relative half.
    controller.setCellFormula({ row: 10, col: 0 }, "=$A$1+B1+$C1+D$1");
    controller.selectCell({ row: 10, col: 0 });
    await vue.nextTick();
    controller.fillSelection({ start: { row: 10, col: 0 }, end: { row: 10, col: 2 } });
    await vue.nextTick();
    assert.deepEqual(
      [0, 1, 2].map((col) => controller.getCellFormula({ row: 10, col })),
      ["=$A$1+B1+$C1+D$1", "=$A$1+C1+$C1+E$1", "=$A$1+D1+$C1+F$1"],
    );

    // Quoted strings, quoted sheet names, structured table references and
    // function names are not rewritten.
    controller.setCellFormula({ row: 3, col: 0 }, "=LOG10(B1)&\"A1\"&'Sheet 1'!C1");
    controller.selectCell({ row: 3, col: 0 });
    await vue.nextTick();
    controller.fillSelection({ start: { row: 3, col: 0 }, end: { row: 4, col: 0 } });
    await vue.nextTick();
    assert.equal(
      controller.getCellFormula({ row: 4, col: 0 }),
      "=LOG10(B2)&\"A1\"&'Sheet 1'!C2",
    );

    controller.setCellFormula({ row: 6, col: 0 }, "=Table1[Amount]+B1+Name1");
    controller.selectCell({ row: 6, col: 0 });
    await vue.nextTick();
    controller.fillSelection({ start: { row: 6, col: 0 }, end: { row: 7, col: 0 } });
    await vue.nextTick();
    assert.equal(
      controller.getCellFormula({ row: 7, col: 0 }),
      "=Table1[Amount]+B2+Name1",
    );

    // Shifting a reference off the sheet produces Excel's #REF!.
    controller.setCellFormula({ row: 8, col: 0 }, "=A1");
    controller.selectCell({ row: 8, col: 0 });
    await vue.nextTick();
    controller.fillSelection({ start: { row: 8, col: 0 }, end: { row: 0, col: 0 } });
    await vue.nextTick();
    assert.equal(controller.getCellFormula({ row: 0, col: 0 }), "=#REF!");
  } finally {
    app.unmount();
  }
});

test("paste shifts relative references by the copy-to-paste distance", async () => {
  const { controller, app } = await mountController();
  try {
    controller.setCellFormula({ row: 0, col: 1 }, "=$A$1+B1");
    controller.selectCell({ row: 0, col: 1 });
    await vue.nextTick();
    const clipboard = controller.getClipboardData();
    assert.deepEqual(
      JSON.parse(clipboard.structured).origin,
      { col: 1, row: 0 },
      "the internal payload records where the range was copied from",
    );

    // One column right: the relative B1 moves, the anchored $A$1 does not.
    controller.selectCell({ row: 0, col: 3 });
    assert.equal(controller.pasteStructuredClipboardData(clipboard.structured), true);
    await vue.nextTick();
    assert.equal(controller.getCellFormula({ row: 0, col: 3 }), "=$A$1+D1");

    // Two rows down.
    controller.selectCell({ row: 0, col: 1 });
    const down = controller.getClipboardData();
    controller.selectCell({ row: 2, col: 1 });
    assert.equal(controller.pasteStructuredClipboardData(down.structured), true);
    await vue.nextTick();
    assert.equal(controller.getCellFormula({ row: 2, col: 1 }), "=$A$1+B3");

    // Payloads without an origin (other hosts, older builds) paste verbatim.
    const legacy = JSON.parse(down.structured);
    delete legacy.origin;
    controller.selectCell({ row: 6, col: 1 });
    assert.equal(controller.pasteStructuredClipboardData(JSON.stringify(legacy)), true);
    await vue.nextTick();
    assert.equal(controller.getCellFormula({ row: 6, col: 1 }), "=$A$1+B1");
  } finally {
    app.unmount();
  }
});

test("paste skips merged secondary cells and stays inside sheet bounds", async () => {
  const { controller, app } = await mountController();
  try {
    const worksheet = controller.getActiveWorksheet();
    const mergeStart = { row: 30, col: 10 };
    const mergeEnd = { row: 31, col: 11 };
    const secondaries = [
      { row: mergeStart.row, col: mergeEnd.col },
      { row: mergeEnd.row, col: mergeStart.col },
      { row: mergeEnd.row, col: mergeEnd.col },
    ];
    const assertMergeIntact = () => {
      assert.deepEqual(worksheet.getMergeSpan(mergeStart.row, mergeStart.col), {
        colSpan: 2,
        rowSpan: 2,
      });
      assert.deepEqual(
        secondaries.map((cell) => readRawCell(controller, cell)),
        ["", "", ""],
        "merged secondary cells must not receive pasted values",
      );
    };

    controller.selectRange({ start: mergeStart, end: mergeEnd });
    controller.mergeSelection();
    await vue.nextTick();
    assert.ok(secondaries.every((cell) => worksheet.isMergedSecondary(cell.row, cell.col)));
    assertMergeIntact();

    // Text paste writes the merge anchor and leaves the secondaries alone.
    controller.selectCell(mergeStart);
    assert.equal(controller.pasteText("one\ttwo\nthree\tfour"), true);
    await vue.nextTick();
    assert.equal(controller.getCellDisplayValue(mergeStart), "one");
    assertMergeIntact();

    // Structured (internal copy) paste behaves the same.
    const sourceStart = { row: 50, col: 10 };
    controller.setCellValue(sourceStart, "alpha");
    controller.setCellValue({ row: 50, col: 11 }, "beta");
    controller.setCellValue({ row: 51, col: 10 }, "gamma");
    controller.setCellValue({ row: 51, col: 11 }, "delta");
    controller.selectRange({ start: sourceStart, end: { row: 51, col: 11 } });
    await vue.nextTick();
    const block = controller.getClipboardData();
    controller.selectCell(mergeStart);
    assert.equal(controller.pasteStructuredClipboardData(block.structured), true);
    await vue.nextTick();
    assert.equal(controller.getCellDisplayValue(mergeStart), "alpha");
    assertMergeIntact();

    // Pasting past the last column stops at the sheet edge instead of writing
    // an out-of-bounds address, which the worksheet engine rejects.
    controller.selectCell({ row: 0, col: 16_383 });
    assert.equal(controller.pasteText("left\tright"), true);
    await vue.nextTick();
    assert.equal(controller.getCellDisplayValue({ row: 0, col: 16_383 }), "left");
    assert.equal(controller.selection?.end.col, 16_383, "the selection stays in bounds");

    controller.setCellValue({ row: 60, col: 10 }, "x");
    controller.setCellValue({ row: 60, col: 11 }, "y");
    controller.selectRange({ start: { row: 60, col: 10 }, end: { row: 60, col: 11 } });
    await vue.nextTick();
    const wide = controller.getClipboardData();
    controller.selectCell({ row: 60, col: 16_383 });
    assert.equal(controller.pasteStructuredClipboardData(wide.structured), true);
    await vue.nextTick();
    assert.equal(controller.getCellDisplayValue({ row: 60, col: 16_383 }), "x");
    assert.equal(controller.selection?.end.col, 16_383, "the selection stays in bounds");
  } finally {
    app.unmount();
  }
});
