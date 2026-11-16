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
  const required = ["onDblClick", "commitEdit", "startEdit", "commitAndTab"];
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
    const editInputRef = { value: { focus() {}, select() {} } };
    const containerRef = { value: { focus() {} } };
    const emit = () => {};
    const nextTick = (task) => { task(); };
    let hitTestResult = null;
    function hitTestCell() { return hitTestResult; }
    function setHitTestResult(cell) { hitTestResult = cell; }
    let adjacentCell = null;
    function getAdjacentCell(cell) { return adjacentCell ?? { row: cell.row, col: cell.col + 1 }; }
    function setAdjacentCell(cell) { adjacentCell = cell; }
    ${inputValueHelper ? "" : "function getCellInputValue(cell) { return props.controller.getCellDisplayValue(cell); }"}
  `;
  const output = ts.transpileModule(
    `${prelude}\n${bodies.join("\n")}\nexport { props, editingCell, editingValue, onDblClick, commitEdit, startEdit, commitAndTab, setHitTestResult, setAdjacentCell };`,
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
