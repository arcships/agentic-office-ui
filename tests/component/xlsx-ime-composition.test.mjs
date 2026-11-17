import assert from "node:assert/strict";
import path from "node:path";
import test from "node:test";
import { pathToFileURL } from "node:url";

import {
  findByTestId,
  mount,
  vue,
} from "./vue-test-renderer.mjs";

const ROOT = path.resolve(import.meta.dirname, "../..");
const XLSX_DIST = pathToFileURL(
  path.join(ROOT, "packages/vue-xlsx/dist/index.js"),
).href;

const { XlsxFormulaBar, XlsxGrid } = await import(XLSX_DIST);

function keyboardEvent(key, { isComposing = false, keyCode = 0 } = {}) {
  const event = {
    key,
    keyCode,
    isComposing,
    defaultPrevented: false,
    propagationStopped: false,
    preventDefault() {
      event.defaultPrevented = true;
    },
    stopPropagation() {
      event.propagationStopped = true;
    },
  };
  return event;
}

function installGridEnvironment() {
  const previousWindow = globalThis.window;
  const previousResizeObserver = globalThis.ResizeObserver;
  const previousDocument = globalThis.Document;
  const previousShadowRoot = globalThis.ShadowRoot;
  globalThis.window = {
    devicePixelRatio: 1,
    requestAnimationFrame: () => 1,
    cancelAnimationFrame() {},
  };
  globalThis.ResizeObserver = class ResizeObserver {
    observe() {}
    disconnect() {}
  };
  // v-model's runtime-dom directive probes `rootNode instanceof Document`.
  globalThis.Document = class Document {};
  globalThis.ShadowRoot = class ShadowRoot {};
  return () => {
    if (previousWindow === undefined) delete globalThis.window;
    else globalThis.window = previousWindow;
    if (previousResizeObserver === undefined) delete globalThis.ResizeObserver;
    else globalThis.ResizeObserver = previousResizeObserver;
    if (previousDocument === undefined) delete globalThis.Document;
    else globalThis.Document = previousDocument;
    if (previousShadowRoot === undefined) delete globalThis.ShadowRoot;
    else globalThis.ShadowRoot = previousShadowRoot;
  };
}

function createGridController() {
  const sheet = {
    name: "Sheet1",
    workbookSheetIndex: 0,
    visibleRows: Array.from({ length: 200 }, (_, index) => index),
    visibleCols: Array.from({ length: 50 }, (_, index) => index),
    mergedRegions: [],
    colWidths: [],
    rowHeights: [],
    colWidthOverridesPx: [],
    rowHeightOverridesPx: [],
    colStyleIds: [],
    rowStyleIds: [],
    styleById: new Map(),
    hyperlinks: [],
    comments: [],
    sparklines: [],
    conditionalFormatRules: [],
    maxUsedRow: 0,
    maxUsedCol: 0,
    showGridLines: true,
    getCellStyleAt: () => null,
  };
  const setCellValueCalls = [];
  return {
    setCellValueCalls,
    controller: {
      activeSheet: sheet,
      activeSheetIndex: 0,
      activeCell: { row: 0, col: 0 },
      activeCellAddress: "A1",
      selection: {
        start: { row: 0, col: 0 },
        end: { row: 0, col: 0 },
      },
      sheets: [sheet],
      tables: [],
      revision: 1,
      zoomScale: 100,
      isWorkerBacked: false,
      readOnly: false,
      selectedFormula: "",
      selectedFormulaTarget: null,
      selectedValue: "",
      getCellDisplayValue: () => "",
      getCellFormula: () => "",
      commitCellInput(cell, value) {
        setCellValueCalls.push({ cell, value });
      },
      getActiveWorksheet: () => null,
      getRowsBatchAsync: async () => [],
      setCellValue(cell, value) {
        setCellValueCalls.push({ cell, value });
      },
      selectCell() {},
      selectRange() {},
      clearSelectedCells() {},
      copySelectionToClipboard() {},
      pasteFromClipboard() {},
      resizeColumn() {},
      resizeRow() {},
      setActiveSheetIndex() {},
    },
  };
}

async function mountGrid(t) {
  const restoreEnvironment = installGridEnvironment();
  const { controller, setCellValueCalls } = createGridController();
  const mounted = await mount(XlsxGrid, { controller });
  const grid = findByTestId(mounted.root, "xlsx-grid");
  assert.ok(grid, "grid root renders");
  grid.getBoundingClientRect = () => ({ left: 0, top: 0, width: 640, height: 480 });
  grid.scrollLeft = 0;
  grid.scrollTop = 0;
  t.after(() => {
    mounted.app.unmount();
    restoreEnvironment();
  });
  return { ...mounted, grid, controller, setCellValueCalls };
}

async function openCellEditor(grid) {
  const enter = keyboardEvent("Enter");
  grid.props.onKeydown(enter);
  assert.equal(enter.defaultPrevented, true, "Enter starts editing a focused cell");
  await vue.nextTick();
  const editor = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
  assert.ok(editor, "inline editor renders");
  return editor;
}

async function flushUi() {
  // Deferred composition commits schedule a Vue nextTick callback that in turn
  // queues the editor unmount, so let two flushes pass before asserting.
  await vue.nextTick();
  await vue.nextTick();
}

test("XlsxGrid defers Enter commits until IME composition ends", async (t) => {
  const { grid, setCellValueCalls } = await mountGrid(t);
  const editor = await openCellEditor(grid);

  editor.props.onCompositionstart({});
  editor.props["onUpdate:modelValue"]("中文");
  const composingEnter = keyboardEvent("Enter", { isComposing: true });
  editor.props.onKeydown(composingEnter);

  assert.deepEqual(setCellValueCalls, [], "Enter during composition must not commit the cell");
  assert.equal(composingEnter.defaultPrevented, false, "the IME keeps ownership of Enter");
  assert.equal(composingEnter.propagationStopped, false);

  editor.props.onCompositionend({});
  await flushUi();

  assert.deepEqual(
    setCellValueCalls,
    [{ cell: { row: 0, col: 0 }, value: "中文" }],
    "the composed value is committed exactly once when composition ends",
  );
  assert.equal(findByTestId(grid.getRootNode(), "xlsx-cell-editor"), undefined);
});

test("XlsxGrid treats legacy IME keyCode 229 as composition", async (t) => {
  const { grid, setCellValueCalls } = await mountGrid(t);
  const editor = await openCellEditor(grid);

  editor.props.onCompositionstart({});
  editor.props["onUpdate:modelValue"]("にほん");
  const legacyEnter = keyboardEvent("Enter", { keyCode: 229 });
  editor.props.onKeydown(legacyEnter);
  assert.deepEqual(setCellValueCalls, [], "keyCode 229 during composition must not commit");

  editor.props.onCompositionend({});
  await flushUi();
  assert.deepEqual(setCellValueCalls, [
    { cell: { row: 0, col: 0 }, value: "にほん" },
  ]);
});

test("XlsxGrid still commits on plain Enter and ignores Escape while composing", async (t) => {
  const { grid, setCellValueCalls } = await mountGrid(t);
  const editor = await openCellEditor(grid);

  editor.props.onCompositionstart({});
  editor.props["onUpdate:modelValue"]("中文");
  const composingEscape = keyboardEvent("Escape", { isComposing: true });
  editor.props.onKeydown(composingEscape);
  assert.equal(composingEscape.defaultPrevented, false, "Escape cancels the IME composition");
  assert.ok(findByTestId(grid.getRootNode(), "xlsx-cell-editor"), "edit mode survives composing Escape");

  editor.props.onCompositionend({});
  await flushUi();
  assert.deepEqual(setCellValueCalls, [], "ending composition without Enter does not commit");

  const plainEnter = keyboardEvent("Enter");
  editor.props.onKeydown(plainEnter);
  assert.equal(plainEnter.defaultPrevented, true);
  assert.equal(plainEnter.propagationStopped, true);
  assert.deepEqual(setCellValueCalls, [
    { cell: { row: 0, col: 0 }, value: "中文" },
  ]);
  await flushUi();
  assert.equal(findByTestId(grid.getRootNode(), "xlsx-cell-editor"), undefined);
});

test("XlsxGrid does not turn composing keydown events into grid commands", async (t) => {
  const { grid, setCellValueCalls } = await mountGrid(t);

  const composingEnter = keyboardEvent("Enter", { isComposing: true });
  grid.props.onKeydown(composingEnter);
  await vue.nextTick();

  assert.equal(composingEnter.defaultPrevented, false);
  assert.equal(
    findByTestId(grid.getRootNode(), "xlsx-cell-editor"),
    undefined,
    "Enter owned by the IME must not open the inline editor",
  );
  assert.deepEqual(setCellValueCalls, []);
});

function createFormulaController() {
  const setSelectedFormulaCalls = [];
  return {
    setSelectedFormulaCalls,
    controller: {
      activeCellAddress: "A1",
      readOnly: false,
      selectedFormula: "",
      selectedFormulaTarget: null,
      selectedValue: "",
      setSelectedFormula(formula) {
        setSelectedFormulaCalls.push(formula);
      },
    },
  };
}

async function mountFormulaBar(t) {
  const { controller, setSelectedFormulaCalls } = createFormulaController();
  const mounted = await mount(XlsxFormulaBar, { controller });
  const input = findByTestId(mounted.root, "xlsx-formula-input");
  assert.ok(input, "formula input renders");
  t.after(() => mounted.app.unmount());
  return { ...mounted, input, setSelectedFormulaCalls };
}

test("XlsxFormulaBar defers Enter commits until IME composition ends", async (t) => {
  const { input, setSelectedFormulaCalls } = await mountFormulaBar(t);

  input.props.onFocus();
  input.props.onInput({ target: { value: "中文" } });
  input.props.onCompositionstart({});
  const composingEnter = keyboardEvent("Enter", { isComposing: true });
  input.props.onKeydown(composingEnter);

  assert.deepEqual(setSelectedFormulaCalls, [], "Enter during composition must not commit the formula");
  assert.equal(composingEnter.defaultPrevented, false, "the IME owns Enter until composition ends");

  input.props.onInput({ target: { value: "中文单元格" } });
  input.props.onCompositionend({});
  await vue.nextTick();

  assert.deepEqual(
    setSelectedFormulaCalls,
    ["中文单元格"],
    "the final composed formula is committed once composition ends",
  );
});

test("XlsxFormulaBar still commits on plain Enter", async (t) => {
  const { input, setSelectedFormulaCalls } = await mountFormulaBar(t);

  input.props.onFocus();
  input.props.onInput({ target: { value: "=SUM(A1:A2)" } });
  const plainEnter = keyboardEvent("Enter");
  input.props.onKeydown(plainEnter);

  assert.equal(plainEnter.defaultPrevented, true);
  assert.deepEqual(setSelectedFormulaCalls, ["=SUM(A1:A2)"]);
});
