import assert from "node:assert/strict";
import path from "node:path";
import test from "node:test";
import { pathToFileURL } from "node:url";

import {
  findByTestId,
  mount,
  textContent,
  vue,
} from "./vue-test-renderer.mjs";

const ROOT = path.resolve(import.meta.dirname, "../..");
const XLSX_DIST = pathToFileURL(
  path.join(ROOT, "packages/vue-xlsx/dist/index.js"),
).href;

const { XlsxGrid } = await import(XLSX_DIST);

const ROW_HEIGHT = 24;
// 2px border + 2px vertical padding on each side of the textarea.
const EDITOR_CHROME_PX = 8;
const EDITOR_LINE_HEIGHT_PX = 15;

function keyboardEvent(key, {
  altKey = false,
  ctrlKey = false,
  isComposing = false,
  keyCode = 0,
  metaKey = false,
  shiftKey = false,
} = {}) {
  const event = {
    key,
    keyCode,
    altKey,
    ctrlKey,
    metaKey,
    shiftKey,
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

function pasteEvent(text) {
  const event = {
    clipboardData: {
      getData(type) {
        return type === "text/plain" || type === "text" ? text : "";
      },
    },
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
    constructor(callback) {
      this.callback = callback;
    }
    observe() {
      // The editor's box is clipped to the grid body; give it a real viewport.
      this.callback([{ contentRect: { width: 640, height: 480 } }]);
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

function cellKey(cell) {
  return `${cell.row}:${cell.col}`;
}

function createGridController() {
  const values = new Map();
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
  const commitCalls = [];
  const selectCalls = [];
  const controller = {
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
    isEditingLocked: false,
    selectedFormula: "",
    selectedFormulaTarget: null,
    selectedValue: "",
    getCellDisplayValue: (cell) => values.get(cellKey(cell)) ?? "",
    getCellFormula: () => "",
    commitCellInput(cell, value) {
      commitCalls.push({ cell: { ...cell }, value });
      values.set(cellKey(cell), value);
      return true;
    },
    getActiveWorksheet: () => null,
    getRowsBatchAsync: async () => [],
    setCellValue(cell, value) {
      values.set(cellKey(cell), value);
    },
    selectCell(cell) {
      selectCalls.push({ ...cell });
      controller.activeCell = { ...cell };
    },
    selectRange() {},
    clearSelectedCells() {},
    copySelectionToClipboard() {},
    pasteFromClipboard() {},
    resizeColumn() {},
    resizeRow() {},
    setActiveSheetIndex() {},
    dismissDataValidationFeedback() {},
  };
  return { commitCalls, controller: vue.reactive(controller), selectCalls, values };
}

async function mountGrid(t) {
  const restoreEnvironment = installGridEnvironment();
  const { commitCalls, controller, selectCalls, values } = createGridController();
  const mounted = await mount(XlsxGrid, { controller });
  const grid = findByTestId(mounted.root, "xlsx-grid");
  assert.ok(grid, "grid root renders");
  grid.getBoundingClientRect = () => ({ left: 0, top: 0, width: 640, height: 480 });
  grid.clientWidth = 640;
  grid.clientHeight = 480;
  grid.scrollLeft = 0;
  grid.scrollTop = 0;
  grid.focus = () => {};
  t.after(() => {
    mounted.app.unmount();
    restoreEnvironment();
  });
  return { ...mounted, commitCalls, controller, grid, selectCalls, values };
}

async function flushUi() {
  // The editor sizes itself from a post-flush watcher, then renders one more
  // time with the measured height.
  await vue.nextTick();
  await vue.nextTick();
}

async function openCellEditor(grid) {
  const enter = keyboardEvent("Enter");
  grid.props.onKeydown(enter);
  assert.equal(enter.defaultPrevented, true, "Enter starts editing a focused cell");
  await flushUi();
  const editor = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
  assert.ok(editor, "inline editor renders");
  return editor;
}

function setDraft(editor, value) {
  editor.props["onUpdate:modelValue"](value);
}

function editorHeight(editor) {
  return Number.parseFloat(String(editor.props.style?.height ?? ""));
}

test("the inline editor is a textarea with a hidden sizing mirror", async (t) => {
  const { grid } = await mountGrid(t);
  const editor = await openCellEditor(grid);

  assert.equal(editor.type, "textarea", "the editor must accept line breaks");
  assert.equal(editor.props.rows, "1", "the textarea grows through its style, not its rows");
  assert.equal(editor.props.wrap, "soft");
  assert.equal(editor.props.spellcheck, "false");
  assert.equal(String(editor.props["data-editor-lines"]), "1");

  const mirror = findByTestId(grid.getRootNode(), "xlsx-cell-editor-mirror");
  assert.ok(mirror, "the sizing mirror renders next to the editor");
  assert.equal(mirror.props["aria-hidden"], "true", "the mirror stays out of the accessibility tree");
  assert.equal(mirror.props.style.visibility, "hidden");
  assert.equal(mirror.props.style.width, editor.props.style.width, "the mirror measures the same box");
  assert.equal(mirror.props.style.fontSize, editor.props.style.fontSize);
  assert.equal(mirror.props.style.lineHeight, editor.props.style.lineHeight);
  assert.equal(mirror.props.style.whiteSpace, "pre-wrap");

  setDraft(editor, "mirror text");
  await flushUi();
  const measured = findByTestId(grid.getRootNode(), "xlsx-cell-editor-mirror");
  assert.equal(textContent(measured), "mirror text", "the mirror carries the current draft");
});

test("Alt+Enter inserts a hard line break and the editor row grows", async (t) => {
  const { commitCalls, grid } = await mountGrid(t);
  let editor = await openCellEditor(grid);
  setDraft(editor, "second line starts here");
  await flushUi();
  editor = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
  assert.equal(editorHeight(editor), ROW_HEIGHT, "a single line keeps the cell height");

  editor.selectionStart = "second line starts here".length;
  editor.selectionEnd = editor.selectionStart;
  const altEnter = keyboardEvent("Enter", { altKey: true });
  editor.props.onKeydown(altEnter);

  assert.equal(altEnter.defaultPrevented, true, "Alt+Enter owns the browser's shortcut");
  assert.equal(altEnter.propagationStopped, true);
  assert.deepEqual(commitCalls, [], "Alt+Enter must not commit the cell");

  await flushUi();
  editor = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
  assert.ok(editor, "the editor stays open for the next line");
  assert.equal(editor.value, "second line starts here\n", "the break is inserted at the caret");
  assert.equal(String(editor.props["data-editor-lines"]), "2");
  assert.equal(
    editorHeight(editor),
    2 * EDITOR_LINE_HEIGHT_PX + EDITOR_CHROME_PX,
    "the editor grows to show the new line instead of scrolling it away",
  );

  const enter = keyboardEvent("Enter");
  editor.props.onKeydown(enter);
  assert.equal(enter.defaultPrevented, true);
  assert.deepEqual(
    commitCalls,
    [{ cell: { row: 0, col: 0 }, value: "second line starts here\n" }],
    "Enter still commits the multi-line draft once",
  );
  await flushUi();
  assert.equal(findByTestId(grid.getRootNode(), "xlsx-cell-editor"), undefined);
});

test("Alt+Enter breaks the draft at the caret", async (t) => {
  const { grid } = await mountGrid(t);
  const editor = await openCellEditor(grid);
  setDraft(editor, "ab");
  await flushUi();
  const focused = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
  focused.selectionStart = 1;
  focused.selectionEnd = 1;
  focused.props.onKeydown(keyboardEvent("Enter", { altKey: true }));
  await flushUi();
  assert.equal(
    findByTestId(grid.getRootNode(), "xlsx-cell-editor").value,
    "a\nb",
    "the caret position decides where the hard break lands",
  );
});

test("Enter commits, Escape leaves the cell untouched and Tab commits and moves", async (t) => {
  const { commitCalls, grid, selectCalls, values } = await mountGrid(t);
  values.set("0:0", "original");
  let editor = await openCellEditor(grid);
  setDraft(editor, "changed");
  await flushUi();

  const escape = keyboardEvent("Escape");
  findByTestId(grid.getRootNode(), "xlsx-cell-editor").props.onKeydown(escape);
  assert.equal(escape.defaultPrevented, true);
  assert.deepEqual(commitCalls, [], "Escape drops the draft instead of writing it");
  await flushUi();
  assert.equal(findByTestId(grid.getRootNode(), "xlsx-cell-editor"), undefined);
  assert.equal(values.get("0:0"), "original", "the cell keeps its original value");

  // Tab still commits and seeds the next cell in the same editor flow.
  editor = await openCellEditor(grid);
  setDraft(editor, "committed");
  await flushUi();
  values.set("0:1", "next");
  const tab = keyboardEvent("Tab");
  findByTestId(grid.getRootNode(), "xlsx-cell-editor").props.onKeydown(tab);
  assert.equal(tab.defaultPrevented, true);
  assert.deepEqual(commitCalls, [{ cell: { row: 0, col: 0 }, value: "committed" }]);
  assert.deepEqual(selectCalls, [{ row: 0, col: 1 }]);
  await flushUi();
  const nextEditor = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
  assert.ok(nextEditor, "Tab keeps editing in the next cell");
  assert.equal(nextEditor.value, "next");
});

test("single-line ↑/↓ commit and move out while a multi-line draft keeps its arrows", async (t) => {
  const { commitCalls, grid, selectCalls } = await mountGrid(t);
  let editor = await openCellEditor(grid);
  setDraft(editor, "single line");
  await flushUi();

  const down = keyboardEvent("ArrowDown");
  findByTestId(grid.getRootNode(), "xlsx-cell-editor").props.onKeydown(down);
  assert.equal(down.defaultPrevented, true, "a single-line draft has no caret row to move to");
  assert.deepEqual(commitCalls, [{ cell: { row: 0, col: 0 }, value: "single line" }]);
  assert.deepEqual(selectCalls, [{ row: 1, col: 0 }], "the selection moves down one cell");
  await flushUi();
  assert.equal(findByTestId(grid.getRootNode(), "xlsx-cell-editor"), undefined);

  grid.props.onKeydown(keyboardEvent("Enter"));
  await flushUi();
  editor = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
  assert.ok(editor, "Enter opens the editor again");
  setDraft(editor, "first\nsecond");
  await flushUi();

  const multilineDown = keyboardEvent("ArrowDown");
  findByTestId(grid.getRootNode(), "xlsx-cell-editor").props.onKeydown(multilineDown);
  assert.equal(multilineDown.defaultPrevented, false, "the caret moves between the draft's lines");
  assert.equal(multilineDown.propagationStopped, false);
  assert.equal(commitCalls.length, 1, "a multi-line draft is not committed by an arrow key");
  assert.ok(findByTestId(grid.getRootNode(), "xlsx-cell-editor"), "the editor stays open");
});

test("blur still commits the multi-line draft", async (t) => {
  const { commitCalls, grid } = await mountGrid(t);
  const editor = await openCellEditor(grid);
  setDraft(editor, "first\nsecond");
  await flushUi();

  findByTestId(grid.getRootNode(), "xlsx-cell-editor").props.onBlur();
  assert.deepEqual(
    commitCalls,
    [{ cell: { row: 0, col: 0 }, value: "first\nsecond" }],
    "clicking away still commits through the existing draft/commit path",
  );
  await flushUi();
  assert.equal(findByTestId(grid.getRootNode(), "xlsx-cell-editor"), undefined);
});

test("pasted multi-line text keeps its line breaks and single-line paste stays native", async (t) => {
  const { commitCalls, grid } = await mountGrid(t);
  const editor = await openCellEditor(grid);

  const multiline = pasteEvent("alpha\r\nbeta");
  editor.props.onPaste(multiline);
  assert.equal(multiline.defaultPrevented, true, "the paste is inserted manually so it cannot be flattened");
  await flushUi();

  const pasted = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
  assert.equal(pasted.value, "alpha\nbeta", "CRLF is normalized to the draft's line separator");
  assert.equal(String(pasted.props["data-editor-lines"]), "2");

  const single = pasteEvent("plain");
  pasted.props.onPaste(single);
  assert.equal(single.defaultPrevented, false, "single-line pastes keep the browser's native caret and undo");
  assert.equal(pasted.value, "alpha\nbeta");

  pasted.props.onKeydown(keyboardEvent("Enter"));
  assert.deepEqual(commitCalls, [{ cell: { row: 0, col: 0 }, value: "alpha\nbeta" }]);
});

test("the hidden mirror's measured height drives the editor", async (t) => {
  const { grid } = await mountGrid(t);
  const editor = await openCellEditor(grid);
  const mirror = findByTestId(grid.getRootNode(), "xlsx-cell-editor-mirror");
  mirror.scrollHeight = 68;

  setDraft(editor, "a long line that wraps over several rows");
  await flushUi();

  const measured = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
  assert.equal(
    editorHeight(measured),
    72,
    "the textarea follows the mirror's box height (scrollHeight plus the border)",
  );
});

test("IME composition still owns the textarea editor", async (t) => {
  const { commitCalls, grid } = await mountGrid(t);
  const editor = await openCellEditor(grid);

  editor.props.onCompositionstart({});
  setDraft(editor, "中文");
  const composingEnter = keyboardEvent("Enter", { isComposing: true });
  editor.props.onKeydown(composingEnter);
  assert.equal(composingEnter.defaultPrevented, false, "the IME keeps the Enter");
  assert.deepEqual(commitCalls, [], "Enter during composition must not commit");

  const composingAltEnter = keyboardEvent("Enter", { altKey: true, isComposing: true });
  editor.props.onKeydown(composingAltEnter);
  assert.equal(composingAltEnter.defaultPrevented, false, "the IME owns Alt+Enter too");
  await flushUi();
  assert.equal(
    findByTestId(grid.getRootNode(), "xlsx-cell-editor").value,
    "中文",
    "a composing Alt+Enter does not inject a line break",
  );

  findByTestId(grid.getRootNode(), "xlsx-cell-editor").props.onCompositionend({});
  await flushUi();
  assert.deepEqual(
    commitCalls,
    [{ cell: { row: 0, col: 0 }, value: "中文" }],
    "composition still ends in exactly one commit",
  );
  assert.equal(findByTestId(grid.getRootNode(), "xlsx-cell-editor"), undefined);
});
