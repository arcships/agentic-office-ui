import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  walk,
  waitFor,
} from "./vue-test-renderer.mjs";

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const {
  useXlsxViewerController,
  XlsxFormulaBar,
  XlsxGrid,
} = await importFromDemo("@arcships/vue-xlsx");

const TARGET = { row: 5, col: 5 };
const FORMULA = "=SUM(A1:A3)+B2";

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
        fileName: "formula-highlight.xlsx",
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

async function seedFormula(controller, formula = FORMULA) {
  controller.setCellFormula(TARGET, formula);
  controller.selectCell(TARGET);
  await vue.nextTick();
  assert.equal(controller.getCellFormula(TARGET), formula);
  return controller;
}

function installGridEnvironment() {
  const previousWindow = globalThis.window;
  const previousResizeObserver = globalThis.ResizeObserver;
  const previousDocument = globalThis.Document;
  const previousShadowRoot = globalThis.ShadowRoot;
  globalThis.window = {
    devicePixelRatio: 1,
    // The paint callback is intentionally never invoked: this renderer has no
    // canvas implementation. Geometry and editor state stay observable.
    requestAnimationFrame: () => 1,
    cancelAnimationFrame() {},
  };
  globalThis.ResizeObserver = class ResizeObserver {
    constructor(callback) {
      this.callback = callback;
    }
    observe() {
      this.callback([{ contentRect: { width: 720, height: 520 } }]);
    }
    disconnect() {}
  };
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

function prepareGrid(root) {
  const grid = findByTestId(root, "xlsx-grid");
  assert.ok(grid, "grid root renders");
  grid.getBoundingClientRect = () => ({ left: 0, top: 0, width: 720, height: 520 });
  grid.clientWidth = 720;
  grid.clientHeight = 520;
  grid.scrollLeft = 0;
  grid.scrollTop = 0;
  grid.focus = () => {};
  return grid;
}

function keyboardEvent(key, extra = {}) {
  const event = {
    altKey: false,
    ctrlKey: false,
    defaultPrevented: false,
    isComposing: false,
    key,
    keyCode: 0,
    metaKey: false,
    shiftKey: false,
    propagationStopped: false,
    preventDefault() {
      event.defaultPrevented = true;
    },
    stopPropagation() {
      event.propagationStopped = true;
    },
    ...extra,
  };
  return event;
}

async function flushUi() {
  await vue.nextTick();
  await vue.nextTick();
}

function referenceSpans(layer) {
  return walk(layer).filter((node) => node.props?.["data-formula-reference"]);
}

function caretInside(formula, token, offset = 1) {
  const start = formula.indexOf(token);
  assert.notEqual(start, -1, `missing ${token} in ${formula}`);
  return start + offset;
}

async function typeIntoFormulaBar(input, value, caret = value.length) {
  input.value = value;
  input.selectionStart = caret;
  input.selectionEnd = caret;
  input.props.onInput({ target: input });
  await flushUi();
}

test("defined names typed in the formula bar resolve to their range", async () => {
  const { controller, app } = await mountController();
  try {
    await seedFormula(controller, "=0");
    assert.equal(controller.setNamedRange("Total", "$A$1:$A$3"), true);
    await vue.nextTick();

    const mounted = await mount(XlsxFormulaBar, { controller });
    try {
      const input = findByTestId(mounted.root, "xlsx-formula-input");
      input.blur = () => {};
      input.props.onFocus();
      await typeIntoFormulaBar(input, "=SUM(Total)", caretInside("=SUM(Total)", "Total", 2));

      const layer = findByTestId(mounted.root, "xlsx-formula-reference-highlight");
      assert.ok(layer, "the defined name is coloured like any other reference");
      assert.equal(layer.props["data-formula-reference-ranges"], "A1:A3");
      assert.equal(layer.props["data-formula-reference-active"], "A1:A3");
    } finally {
      mounted.app.unmount();
    }
  } finally {
    app.unmount();
  }
});

configureWasm();

test("the formula bar colours references and follows the caret", async () => {
  const { controller, app } = await mountController();
  try {
    await seedFormula(controller);
    const mounted = await mount(XlsxFormulaBar, { controller });
    try {
      const input = findByTestId(mounted.root, "xlsx-formula-input");
      assert.ok(input, "formula bar input renders");
      input.blur = () => {};

      input.props.onFocus();
      await typeIntoFormulaBar(input, FORMULA, caretInside(FORMULA, "A1:A3"));

      const layer = findByTestId(mounted.root, "xlsx-formula-reference-highlight");
      assert.ok(layer, "reference colour layer renders while editing");
      assert.equal(textContent(layer), FORMULA, "the layer paints the same formula text");
      assert.equal(layer.props["data-formula-reference-ranges"], "A1:A3 B2");
      assert.equal(layer.props["data-formula-reference-active"], "A1:A3");

      const spans = referenceSpans(layer);
      assert.deepEqual(
        spans.map((span) => span.props["data-formula-reference"]),
        ["A1:A3", "B2"],
      );
      assert.notEqual(spans[0].props.style.color, spans[1].props.style.color);
      assert.equal(spans[0].props["data-formula-reference-active"], "true");
      assert.equal(spans[1].props["data-formula-reference-active"], undefined);
      assert.equal(input.props.style.color, "transparent", "the native text yields to the layer");
      assert.ok(input.props.style.caretColor, "the caret stays visible over the layer");

      // Moving the caret from A1:A3 to B2 moves the active outline and the
      // emphasised reference text.
      input.selectionStart = caretInside(FORMULA, "B2");
      input.selectionEnd = input.selectionStart;
      input.props.onSelect({ target: input });
      await flushUi();
      const moved = findByTestId(mounted.root, "xlsx-formula-reference-highlight");
      assert.equal(moved.props["data-formula-reference-active"], "B2");
      const movedSpans = referenceSpans(moved);
      assert.equal(movedSpans[0].props["data-formula-reference-active"], undefined);
      assert.equal(movedSpans[1].props["data-formula-reference-active"], "true");

      // Escape drops the draft and the highlight, leaving the cell untouched.
      const cancelledDraft = `${FORMULA}+C3`;
      await typeIntoFormulaBar(input, cancelledDraft, caretInside(cancelledDraft, "B2"));
      assert.equal(
        controller.getCellFormula(TARGET),
        FORMULA,
        "highlighting the draft never mutates the cell before commit",
      );
      input.props.onKeydown(keyboardEvent("Escape"));
      await flushUi();
      assert.equal(
        findByTestId(mounted.root, "xlsx-formula-reference-highlight"),
        undefined,
        "Escape clears the coloured references",
      );
      assert.equal(controller.getCellFormula(TARGET), FORMULA);
      assert.equal(input.props.value, FORMULA, "the draft snaps back to the cell value");
    } finally {
      mounted.app.unmount();
    }
  } finally {
    app.unmount();
  }
});

test("Enter commits the formula and clears the formula-bar highlight", async () => {
  const { controller, app } = await mountController();
  try {
    await seedFormula(controller);
    const mounted = await mount(XlsxFormulaBar, { controller });
    try {
      const input = findByTestId(mounted.root, "xlsx-formula-input");
      input.blur = () => {};
      const nextFormula = "=SUM(A1:A3)+B2+C3";
      input.props.onFocus();
      await typeIntoFormulaBar(input, nextFormula, caretInside(nextFormula, "B2"));
      assert.ok(findByTestId(mounted.root, "xlsx-formula-reference-highlight"));

      input.props.onKeydown(keyboardEvent("Enter"));
      await flushUi();
      assert.equal(controller.getCellFormula(TARGET), nextFormula);
      assert.equal(
        findByTestId(mounted.root, "xlsx-formula-reference-highlight"),
        undefined,
        "Enter closes the edit and the reference layer",
      );
    } finally {
      mounted.app.unmount();
    }
  } finally {
    app.unmount();
  }
});

test("the inline grid editor colours references and outlines their cells", async (t) => {
  const restoreEnvironment = installGridEnvironment();
  t.after(restoreEnvironment);
  const { controller, app } = await mountController();
  try {
    await seedFormula(controller);
    const mounted = await mount(XlsxGrid, { controller });
    try {
      const grid = prepareGrid(mounted.root);
      const enter = keyboardEvent("Enter");
      grid.props.onKeydown(enter);
      assert.equal(enter.defaultPrevented, true, "Enter opens the inline editor");
      await flushUi();

      let editor = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
      assert.ok(editor, "inline editor renders");
      assert.equal(editor.value, FORMULA);
      assert.equal(grid.props["data-formula-reference-ranges"], "A1:A3 B2");
      assert.equal(
        grid.props["data-formula-reference-active"],
        "B2",
        "the seeded caret sits at the end of the formula, inside B2",
      );

      const mirror = findByTestId(grid.getRootNode(), "xlsx-cell-editor-mirror");
      assert.ok(mirror, "the mirror doubles as the reference colour layer");
      assert.equal(mirror.props.style.visibility, "visible");
      assert.equal(textContent(mirror), FORMULA);
      assert.equal(mirror.props["data-formula-reference-count"], 2);
      const spans = referenceSpans(mirror);
      assert.deepEqual(
        spans.map((span) => span.props["data-formula-reference"]),
        ["A1:A3", "B2"],
      );
      assert.notEqual(spans[0].props.style.color, spans[1].props.style.color);
      assert.equal(editor.props.style.color, "transparent");
      assert.equal(editor.props["data-formula-references"], "true");

      // Moving the caret inside A1:A3 activates that range's thicker outline.
      editor.selectionStart = caretInside(FORMULA, "A1:A3");
      editor.selectionEnd = editor.selectionStart;
      editor.props.onSelect({ target: editor });
      await flushUi();
      editor = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
      assert.equal(grid.props["data-formula-reference-active"], "A1:A3");
      const movedSpans = referenceSpans(
        findByTestId(grid.getRootNode(), "xlsx-cell-editor-mirror"),
      );
      assert.equal(movedSpans[0].props["data-formula-reference-active"], "true");
      assert.equal(movedSpans[1].props["data-formula-reference-active"], undefined);

      // Escape restores the cell and removes both the text colours and the
      // grid outlines.
      editor.props.onKeydown(keyboardEvent("Escape"));
      await flushUi();
      assert.equal(findByTestId(grid.getRootNode(), "xlsx-cell-editor"), undefined);
      assert.equal(grid.props["data-formula-reference-ranges"], undefined);
      assert.equal(grid.props["data-formula-reference-active"], undefined);
      assert.equal(controller.getCellFormula(TARGET), FORMULA, "Escape does not mutate the cell");

      // A plain value is not a formula: no reference layer, no outlines.
      grid.props.onKeydown(keyboardEvent("Enter"));
      await flushUi();
      editor = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
      editor.props["onUpdate:modelValue"]("plain text");
      await flushUi();
      const plainMirror = findByTestId(grid.getRootNode(), "xlsx-cell-editor-mirror");
      assert.equal(plainMirror.props.style.visibility, "hidden");
      assert.notEqual(
        findByTestId(grid.getRootNode(), "xlsx-cell-editor").props.style.color,
        "transparent",
        "the textarea keeps its normal text colour",
      );
      assert.equal(grid.props["data-formula-reference-ranges"], undefined);

      // Enter commits and clears the highlight as well.
      editor = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
      editor.props["onUpdate:modelValue"](FORMULA);
      await flushUi();
      assert.equal(grid.props["data-formula-reference-ranges"], "A1:A3 B2");
      findByTestId(grid.getRootNode(), "xlsx-cell-editor")
        .props.onKeydown(keyboardEvent("Enter"));
      await flushUi();
      assert.equal(grid.props["data-formula-reference-ranges"], undefined);
      assert.equal(controller.getCellFormula(TARGET), FORMULA);
    } finally {
      mounted.app.unmount();
    }
  } finally {
    app.unmount();
  }
});

test("formula-bar editing outlines referenced cells in the grid", async (t) => {
  const restoreEnvironment = installGridEnvironment();
  t.after(restoreEnvironment);
  const { controller, app } = await mountController();
  try {
    await seedFormula(controller);
    const Harness = vue.defineComponent({
      setup() {
        return () => vue.h("div", [
          vue.h(XlsxFormulaBar, { controller }),
          vue.h(XlsxGrid, { controller }),
        ]);
      },
    });
    const mounted = await mount(Harness);
    try {
      const grid = prepareGrid(mounted.root);
      const input = findByTestId(mounted.root, "xlsx-formula-input");
      input.blur = () => {};

      input.props.onFocus();
      await typeIntoFormulaBar(input, FORMULA, caretInside(FORMULA, "A1:A3"));
      assert.equal(
        grid.props["data-formula-reference-ranges"],
        "A1:A3 B2",
        "the grid outlines the references typed in the formula bar",
      );
      assert.equal(grid.props["data-formula-reference-active"], "A1:A3");
      assert.equal(
        findByTestId(grid.getRootNode(), "xlsx-cell-editor"),
        undefined,
        "the grid itself need not be editing",
      );

      // A sheet-qualified target is coloured in the bar but not outlined on a
      // sheet the grid is not showing.
      const crossSheet = "=SUM(A1:A3)+OtherSheet!B2";
      await typeIntoFormulaBar(input, crossSheet, caretInside(crossSheet, "A1:A3"));
      assert.equal(
        findByTestId(mounted.root, "xlsx-formula-reference-highlight")
          .props["data-formula-reference-ranges"],
        "A1:A3 B2",
        "the formula bar still colours the cross-sheet reference",
      );
      assert.equal(grid.props["data-formula-reference-ranges"], "A1:A3");
      assert.equal(grid.props["data-formula-reference-active"], "A1:A3");

      input.props.onKeydown(keyboardEvent("Escape"));
      await flushUi();
      assert.equal(grid.props["data-formula-reference-ranges"], undefined);
      assert.equal(grid.props["data-formula-reference-active"], undefined);
    } finally {
      mounted.app.unmount();
    }
  } finally {
    app.unmount();
  }
});
