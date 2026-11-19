import assert from "node:assert/strict";
import path from "node:path";
import test from "node:test";
import { pathToFileURL } from "node:url";

import { findByTestId, mount, vue, walk } from "./vue-test-renderer.mjs";

const ROOT = path.resolve(import.meta.dirname, "../..");
const XLSX_DIST = pathToFileURL(
  path.join(ROOT, "packages/vue-xlsx/dist/index.js"),
).href;

const { XlsxGrid } = await import(XLSX_DIST);

const ROW_HEIGHT = 24;
const GRID_WIDTH = 640;
const GRID_HEIGHT = 480;
// `onMounted` stores the ResizeObserver box, so the body is the grid minus the
// 24 px column-header strip.
const BODY_HEIGHT = GRID_HEIGHT - 24;
// One viewport of 24 px rows keeps a row of overlap, exactly like Excel.
const PAGE_STEP = Math.floor(BODY_HEIGHT / ROW_HEIGHT) - 1;
// The grid paints at least 200 rows and 50 columns so a viewport is always full.
const DISPLAY_ROWS = 200;
const DISPLAY_COLS = 50;

function keyboardEvent(key, {
  altKey = false,
  ctrlKey = false,
  metaKey = false,
  shiftKey = false,
} = {}) {
  const event = {
    key,
    keyCode: 0,
    altKey,
    ctrlKey,
    metaKey,
    shiftKey,
    isComposing: false,
    defaultPrevented: false,
    preventDefault() {
      event.defaultPrevented = true;
    },
    stopPropagation() {},
  };
  return event;
}

// Mounts share the process globals, so a test that mounts two grids would
// otherwise tear the stubs down when the first grid unmounts.
let environmentRefs = 0;
let restoreGridEnvironment = null;

function installGridEnvironment() {
  if (environmentRefs === 0) restoreGridEnvironment = applyGridEnvironment();
  environmentRefs += 1;
  let released = false;
  return () => {
    if (released) return;
    released = true;
    environmentRefs -= 1;
    if (environmentRefs === 0 && restoreGridEnvironment) {
      restoreGridEnvironment();
      restoreGridEnvironment = null;
    }
  };
}

function applyGridEnvironment() {
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
    constructor(callback) {
      this.callback = callback;
    }
    observe() {
      this.callback([{ contentRect: { width: GRID_WIDTH, height: GRID_HEIGHT } }]);
    }
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

function cellId(cell) {
  return `${cell.row}:${cell.col}`;
}

/**
 * Controller stub shaped like `XlsxViewerController` for the keys under test:
 * it records selection commands and keeps `activeCell`/`selection` symmetric
 * with the real controller (`selectRange` activates the range's end).
 */
function createKeyboardController({
  values = {},
  usedRow = 4,
  usedCol = 6,
  cancelCut = () => false,
  rightToLeft = false,
} = {}) {
  const valueMap = new Map(Object.entries(values));
  const sheet = {
    name: "Sheet1",
    workbookSheetIndex: 0,
    rightToLeft,
    visibleRows: Array.from({ length: DISPLAY_ROWS }, (_, index) => index),
    visibleCols: Array.from({ length: DISPLAY_COLS }, (_, index) => index),
    mergedRegions: [],
    colWidths: [],
    rowHeights: [],
    colWidthOverridesPx: {},
    rowHeightOverridesPx: {},
    colStyleIds: {},
    rowStyleIds: {},
    styleById: new Map(),
    hyperlinks: [],
    comments: [],
    sparklines: [],
    conditionalFormatRules: [],
    minUsedRow: usedRow >= 0 ? 0 : -1,
    minUsedCol: usedCol >= 0 ? 0 : -1,
    maxUsedRow: usedRow,
    maxUsedCol: usedCol,
    showGridLines: true,
    getCellStyleAt: () => null,
  };
  const commits = [];
  const fills = [];
  const selections = [];
  const controller = vue.reactive({
    activeSheet: sheet,
    activeSheetIndex: 0,
    activeCell: { row: 0, col: 0 },
    activeCellAddress: "A1",
    selection: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
    sheets: [sheet],
    tabs: [],
    tables: [],
    charts: [],
    images: [],
    shapes: [],
    formControls: [],
    revision: 1,
    zoomScale: 100,
    isLoading: false,
    error: null,
    isWorkerBacked: false,
    readOnly: false,
    isEditingLocked: false,
    selectedFormula: "",
    selectedFormulaTarget: null,
    selectedValue: "",
    canDownload: false,
    canExport: false,
    getCellDisplayValue(cell) {
      if (!cell) return "";
      return valueMap.get(cellId(cell)) ?? "";
    },
    getCellFormula() {
      return "";
    },
    getActiveWorksheet: () => null,
    getRowsBatchAsync: async () => [],
    setCellValue(cell, value) {
      valueMap.set(cellId(cell), value);
    },
    commitCellInput(cell, text) {
      commits.push({ cell: { ...cell }, text });
      valueMap.set(cellId(cell), text);
      return true;
    },
    commitSelectionInput(text) {
      fills.push(text);
      return 1;
    },
    cancelCut,
    cutSelection() {},
    clearSelectedCells() {},
    copySelectionToClipboard() {},
    pasteFromClipboard() {},
    resizeColumn() {},
    resizeRow() {},
    setActiveSheetIndex() {},
    dismissDataValidationFeedback() {},
    selectCell(cell) {
      selections.push({ kind: "cell", cell: { ...cell } });
      controller.activeCell = { ...cell };
      controller.selection = { start: { ...cell }, end: { ...cell } };
    },
    selectRange(range) {
      selections.push({
        kind: "range",
        range: {
          end: { ...range.end },
          start: { ...range.start },
        },
      });
      controller.selection = { start: { ...range.start }, end: { ...range.end } };
      controller.activeCell = { ...range.end };
    },
  });
  return { controller, commits, fills, selections, sheet, valueMap };
}

async function mountGrid(t, options = {}) {
  const restoreEnvironment = installGridEnvironment();
  const harness = createKeyboardController(options);
  const mounted = await mount(XlsxGrid, { controller: harness.controller });
  t.after(() => {
    mounted.app.unmount();
    restoreEnvironment();
  });
  const grid = findByTestId(mounted.root, "xlsx-grid");
  assert.ok(grid, "grid root renders");
  grid.getBoundingClientRect = () => ({ left: 0, top: 0, width: GRID_WIDTH, height: GRID_HEIGHT });
  grid.clientWidth = GRID_WIDTH;
  grid.clientHeight = GRID_HEIGHT;
  grid.scrollLeft = 0;
  grid.scrollTop = 0;
  grid.focus = () => {};
  for (const node of walk(mounted.root)) {
    if (node.type === "canvas") node.style = {};
  }
  await vue.nextTick();
  return { ...harness, grid, mounted };
}

/** Selects a cell or range the way the controller would, then lets Vue settle. */
async function select(harness, range) {
  harness.controller.selectRange(range);
  await vue.nextTick();
}

function pressGrid(grid, key, options) {
  const event = keyboardEvent(key, options);
  grid.props.onKeydown(event);
  return event;
}

function pressEditor(mounted, key, options) {
  const editor = findByTestId(mounted.root, "xlsx-cell-editor");
  assert.ok(editor, "inline editor renders");
  const event = keyboardEvent(key, options);
  editor.props.onKeydown(event);
  return { editor, event };
}

test("PageDown and PageUp move the active cell by one viewport of rows", async (t) => {
  const harness = await mountGrid(t);
  const { controller, grid } = harness;
  assert.equal(PAGE_STEP, 18, "the test viewport fits nineteen rows");

  const pageDown = pressGrid(grid, "PageDown");
  assert.equal(pageDown.defaultPrevented, true, "the grid consumes PageDown");
  assert.deepEqual(controller.activeCell, { row: PAGE_STEP, col: 0 }, "the column is kept");

  pressGrid(grid, "PageDown");
  assert.deepEqual(controller.activeCell, { row: PAGE_STEP * 2, col: 0 });

  pressGrid(grid, "PageUp");
  assert.deepEqual(controller.activeCell, { row: PAGE_STEP, col: 0 });

  await select(harness, { start: { row: 2, col: 3 }, end: { row: 2, col: 3 } });
  const extend = pressGrid(grid, "PageDown", { shiftKey: true });
  assert.equal(extend.defaultPrevented, true);
  assert.deepEqual(controller.selection, {
    start: { row: 2, col: 3 },
    end: { row: 2 + PAGE_STEP, col: 3 },
  }, "Shift+PageDown extends from the active cell");

  await select(harness, { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
  for (let index = 0; index < 20; index += 1) pressGrid(grid, "PageDown");
  assert.deepEqual(
    controller.activeCell,
    { row: DISPLAY_ROWS - 1, col: 0 },
    "paging past the last row stops on it",
  );

  // Ctrl/Cmd+PageUp/PageDown switches sheet tabs and is tracked separately
  // (#11); the grid must leave it to the viewer.
  const tabShortcut = pressGrid(grid, "PageDown", { ctrlKey: true });
  assert.equal(tabShortcut.defaultPrevented, false, "Ctrl+PageDown belongs to the sheet tabs");
  assert.deepEqual(
    controller.activeCell,
    { row: DISPLAY_ROWS - 1, col: 0 },
    "the active cell does not page on Ctrl+PageDown",
  );
});

test("Home and End jump to the row edges, Ctrl+Home and Ctrl+End to the sheet edges", async (t) => {
  const harness = await mountGrid(t, {
    values: { "1:0": "a", "1:3": "b", "3:2": "c" },
    usedRow: 3,
    usedCol: 5,
  });
  const { controller, grid } = harness;

  await select(harness, { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } });
  const home = pressGrid(grid, "Home");
  assert.equal(home.defaultPrevented, true);
  assert.deepEqual(controller.activeCell, { row: 1, col: 0 }, "Home is the row start");

  const end = pressGrid(grid, "End");
  assert.equal(end.defaultPrevented, true);
  assert.deepEqual(controller.activeCell, { row: 1, col: 3 }, "End is the row's last entry");

  await select(harness, { start: { row: 2, col: 2 }, end: { row: 2, col: 2 } });
  pressGrid(grid, "End");
  assert.deepEqual(
    controller.activeCell,
    { row: 2, col: DISPLAY_COLS - 1 },
    "an empty row falls back to the last displayed column",
  );

  const ctrlHome = pressGrid(grid, "Home", { ctrlKey: true });
  assert.equal(ctrlHome.defaultPrevented, true, "the browser's Home is replaced");
  assert.deepEqual(controller.activeCell, { row: 0, col: 0 }, "Ctrl+Home is the sheet start");

  pressGrid(grid, "End", { ctrlKey: true });
  assert.deepEqual(controller.activeCell, { row: 3, col: 5 }, "Ctrl+End is the last used cell");

  await select(harness, { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } });
  pressGrid(grid, "End", { shiftKey: true });
  assert.deepEqual(
    controller.selection,
    { start: { row: 1, col: 1 }, end: { row: 1, col: 3 } },
    "Shift+End extends the selection",
  );

  const empty = await mountGrid(t, { usedRow: -1, usedCol: -1 });
  pressGrid(empty.grid, "End", { ctrlKey: true });
  assert.deepEqual(
    empty.controller.activeCell,
    { row: 0, col: 0 },
    "an empty sheet ends where its used range starts",
  );
});

test("Ctrl+A selects the used range and the whole grid on a second press", async (t) => {
  const harness = await mountGrid(t, { usedRow: 3, usedCol: 5 });
  const { controller, grid } = harness;

  const first = pressGrid(grid, "a", { ctrlKey: true });
  assert.equal(first.defaultPrevented, true, "the browser's select-all is replaced");
  assert.deepEqual(controller.selection, {
    start: { row: 0, col: 0 },
    end: { row: 3, col: 5 },
  }, "the used range is selected");
  assert.deepEqual(controller.activeCell, { row: 3, col: 5 }, "the used end is active");

  pressGrid(grid, "a", { metaKey: true });
  assert.deepEqual(controller.selection, {
    start: { row: 0, col: 0 },
    end: { row: DISPLAY_ROWS - 1, col: DISPLAY_COLS - 1 },
  }, "a second Ctrl/Cmd+A behaves like the corner click");

  const empty = await mountGrid(t, { usedRow: -1, usedCol: -1 });
  pressGrid(empty.grid, "a", { ctrlKey: true });
  assert.deepEqual(empty.controller.selection, {
    start: { row: 0, col: 0 },
    end: { row: DISPLAY_ROWS - 1, col: DISPLAY_COLS - 1 },
  }, "an empty sheet selects everything on the first press");
});

test("Escape collapses a range selection to the active cell and still cancels a cut", async (t) => {
  const harness = await mountGrid(t, { values: { "2:2": "x" } });
  const { controller, grid } = harness;

  await select(harness, { start: { row: 1, col: 1 }, end: { row: 2, col: 2 } });
  const escape = pressGrid(grid, "Escape");
  assert.equal(escape.defaultPrevented, true, "the grid consumes the collapse");
  assert.deepEqual(controller.selection, {
    start: { row: 2, col: 2 },
    end: { row: 2, col: 2 },
  }, "the range collapses onto its active cell");
  assert.deepEqual(controller.activeCell, { row: 2, col: 2 });

  const single = pressGrid(grid, "Escape");
  assert.equal(single.defaultPrevented, false, "a single cell leaves Escape to the browser");

  let cancelled = 0;
  const cutting = await mountGrid(t, {
    values: {},
    cancelCut() {
      cancelled += 1;
      return true;
    },
  });
  await select(cutting, { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } });
  const cancel = pressGrid(cutting.grid, "Escape");
  assert.equal(cancel.defaultPrevented, true);
  assert.equal(cancelled, 1, "the pending cut still owns the first Escape");
  assert.deepEqual(
    cutting.controller.selection,
    { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
    "cancelling a cut leaves the selection alone",
  );
});

test("Ctrl+Arrow jumps to the data edge and stops on the axis ends", async (t) => {
  const harness = await mountGrid(t, {
    values: {
      "0:0": "a",
      "0:1": "b",
      "0:3": "c",
      "1:0": "d",
      "3:0": "e",
    },
    usedRow: 3,
    usedCol: 3,
  });
  const { controller, grid } = harness;

  const runEnd = pressGrid(grid, "ArrowRight", { ctrlKey: true });
  assert.equal(runEnd.defaultPrevented, true, "the grid consumes Ctrl+Arrow");
  assert.deepEqual(controller.activeCell, { row: 0, col: 1 }, "the run ends on B1");

  pressGrid(grid, "ArrowRight", { metaKey: true });
  assert.deepEqual(controller.activeCell, { row: 0, col: 3 }, "the gap jumps to the next entry");

  pressGrid(grid, "ArrowRight", { ctrlKey: true });
  assert.deepEqual(
    controller.activeCell,
    { row: 0, col: DISPLAY_COLS - 1 },
    "no entry further right lands on the last column",
  );

  pressGrid(grid, "ArrowLeft", { ctrlKey: true });
  assert.deepEqual(controller.activeCell, { row: 0, col: 3 }, "walking back finds the entry again");

  await select(harness, { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
  pressGrid(grid, "ArrowDown", { ctrlKey: true });
  assert.deepEqual(controller.activeCell, { row: 1, col: 0 }, "the contiguous run ends on A2");

  pressGrid(grid, "ArrowDown", { ctrlKey: true });
  assert.deepEqual(controller.activeCell, { row: 3, col: 0 }, "A4 holds the next entry");

  await select(harness, { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
  pressGrid(grid, "ArrowRight", { ctrlKey: true, shiftKey: true });
  assert.deepEqual(
    controller.selection,
    { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } },
    "Ctrl+Shift+Arrow extends the selection",
  );
});

test("Shift+Enter steps up, and in the editor it commits and steps up", async (t) => {
  const harness = await mountGrid(t, { values: { "2:2": "value" } });
  const { controller, commits, grid, mounted } = harness;

  await select(harness, { start: { row: 2, col: 2 }, end: { row: 2, col: 2 } });
  const shiftEnter = pressGrid(grid, "Enter", { shiftKey: true });
  assert.equal(shiftEnter.defaultPrevented, true);
  assert.deepEqual(controller.activeCell, { row: 1, col: 2 }, "Shift+Enter moves up one row");
  assert.deepEqual(commits, [], "moving up does not open the editor");

  await select(harness, { start: { row: 2, col: 2 }, end: { row: 2, col: 2 } });
  pressGrid(grid, "Enter");
  await vue.nextTick();
  const editor = findByTestId(mounted.root, "xlsx-cell-editor");
  assert.ok(editor, "Enter opens the editor on the active cell");
  editor.props["onUpdate:modelValue"]("edited");
  await vue.nextTick();

  const committed = pressEditor(mounted, "Enter", { shiftKey: true });
  await vue.nextTick();
  assert.equal(committed.event.defaultPrevented, true);
  assert.deepEqual(commits, [{ cell: { row: 2, col: 2 }, text: "edited" }]);
  assert.deepEqual(controller.activeCell, { row: 1, col: 2 }, "committing with Shift+Enter steps up");
  assert.equal(Boolean(findByTestId(mounted.root, "xlsx-cell-editor")), false, "the editor closes");
});

test("Ctrl+Arrow mirrors the column axis on a right-to-left sheet", async (t) => {
  const harness = await mountGrid(t, {
    values: { "0:0": "a", "0:2": "c" },
    usedRow: 0,
    usedCol: 2,
    rightToLeft: true,
  });
  const { controller, grid } = harness;

  pressGrid(grid, "ArrowLeft", { ctrlKey: true });
  assert.deepEqual(
    controller.activeCell,
    { row: 0, col: 2 },
    "Ctrl+Left crosses the gap leftwards on an RTL sheet",
  );

  pressGrid(grid, "ArrowRight", { ctrlKey: true });
  assert.deepEqual(
    controller.activeCell,
    { row: 0, col: 0 },
    "Ctrl+Right walks back towards the lower display index",
  );
});

test("Enter over a multi-cell selection fills the selection instead of one cell", async (t) => {
  // `selectRange` activates the range's end, which seeds the inline editor.
  const harness = await mountGrid(t, { values: { "2:2": "seed" }, usedRow: 2, usedCol: 2 });
  const { commits, fills, grid, mounted } = harness;

  await select(harness, { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } });
  pressGrid(grid, "Enter");
  await vue.nextTick();
  const editor = findByTestId(mounted.root, "xlsx-cell-editor");
  assert.ok(editor, "the editor opens on the active cell");
  assert.equal(editor.value, "seed", "the active cell seeds the draft");

  editor.props["onUpdate:modelValue"]("filled");
  await vue.nextTick();
  const committed = pressEditor(mounted, "Enter");
  await vue.nextTick();
  assert.equal(committed.event.defaultPrevented, true);
  assert.deepEqual(fills, ["filled"], "the whole selection takes the entry in one edit");
  assert.deepEqual(commits, [], "the active cell is not written twice");
  assert.deepEqual(
    harness.controller.selection,
    { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } },
    "the filled range stays selected",
  );
  assert.equal(Boolean(findByTestId(mounted.root, "xlsx-cell-editor")), false, "the editor closes");
});

test("Tab and blur commit only the active cell of a multi-cell selection", async (t) => {
  const harness = await mountGrid(t, { values: { "2:2": "seed" }, usedRow: 2, usedCol: 2 });
  const { commits, fills, grid, mounted } = harness;

  await select(harness, { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } });
  pressGrid(grid, "Enter");
  await vue.nextTick();
  const editor = findByTestId(mounted.root, "xlsx-cell-editor");
  assert.ok(editor, "the editor opens on the active cell");
  editor.props["onUpdate:modelValue"]("typed");
  await vue.nextTick();

  pressEditor(mounted, "Tab");
  await vue.nextTick();
  assert.deepEqual(
    commits,
    [{ cell: { row: 2, col: 2 }, text: "typed" }],
    "Tab commits the active cell only",
  );
  assert.deepEqual(fills, [], "Tab does not fill the range");

  // Close the editor Tab moved onto, then blur a fresh range edit.
  pressEditor(mounted, "Escape");
  await vue.nextTick();
  await select(harness, { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } });
  pressGrid(grid, "Enter");
  await vue.nextTick();
  const blurred = findByTestId(mounted.root, "xlsx-cell-editor");
  assert.ok(blurred, "the editor reopens over the range");
  blurred.props["onUpdate:modelValue"]("blurred");
  await vue.nextTick();
  blurred.props.onBlur();
  await vue.nextTick();
  assert.deepEqual(
    commits.at(-1),
    { cell: { row: 2, col: 2 }, text: "blurred" },
    "blur commits the active cell only",
  );
  assert.deepEqual(fills, [], "blur does not fill the range");
  assert.equal(Boolean(findByTestId(mounted.root, "xlsx-cell-editor")), false, "blur closes the editor");
});
