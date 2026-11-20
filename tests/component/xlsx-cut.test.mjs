import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";
import {
  findByTestId,
  importFromDemo,
  keydownEvent,
  mount,
  teleportTarget,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

// The cut workflow is a state machine on top of the clipboard domain, so the
// tests drive the real controller and the real grid: only the browser plumbing
// the renderer cannot provide (animation frames, resize observers, canvas
// contexts) is stubbed. Painting itself is asserted against the component
// source because the host nodes have no 2D context.
globalThis.ResizeObserver ??= class ResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
};
globalThis.Document ??= class Document {};
globalThis.ShadowRoot ??= class ShadowRoot {};
// A parked animation frame: schedules are recorded but never run, so a pending
// cut marquee cannot turn a test into a busy repaint loop.
globalThis.window ??= {
  addEventListener() {},
  cancelAnimationFrame() {},
  devicePixelRatio: 1,
  removeEventListener() {},
  requestAnimationFrame: () => 1,
};
if (typeof Object.prototype.getContext !== "function") {
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value: () => null,
  });
}

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

function descendants(node, predicate) {
  const matches = [];
  function visit(current) {
    if (predicate(current)) matches.push(current);
    ts.forEachChild(current, visit);
  }
  visit(node);
  return matches;
}

function functionText(name) {
  const node = descendants(
    GRID_AST,
    (child) => ts.isFunctionDeclaration(child) && child.name?.text === name,
  )[0];
  assert.ok(node, `XlsxGrid is missing ${name}`);
  return node.getText(GRID_AST);
}

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const {
  XlsxContextMenu,
  XlsxGrid,
  XlsxSelectionOverlay,
  XlsxViewer,
  useXlsxViewerController,
} = await importFromDemo("@arcships/vue-xlsx");

function configureWasm() {
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}

function workbookBuffer(name = "sales-table.xlsx") {
  const bytes = readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function mountController(options = {}) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(options.sample),
        fileName: options.sample ?? "sales-table.xlsx",
        skipXmlParsing: true,
        useWorker: false,
        ...options.controller,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "XLSX cut controller ready",
  );
  assert.ifError(controller.error);
  return { ...mounted, controller };
}

async function mountGrid(options = {}) {
  let controller;
  let gridRef;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(options.sample),
        fileName: options.sample ?? "sales-table.xlsx",
        skipXmlParsing: true,
        useWorker: false,
      });
      gridRef = vue.ref(null);
      return () => vue.h("div", [
        vue.h(XlsxGrid, {
          ref: gridRef,
          controller,
          readOnly: options.readOnly === true,
        }),
      ]);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "XLSX grid controller ready",
  );
  assert.ifError(controller.error);
  return { ...mounted, controller, grid: findByTestId(mounted.root, "xlsx-grid") };
}

function setRangeValues(controller, col, values) {
  values.forEach((value, row) => {
    if (typeof value === "string" && value.startsWith("=")) {
      controller.setCellFormula({ row, col }, value);
    } else {
      controller.setCellValue({ row, col }, String(value));
    }
  });
}

function displayValues(controller, col, rows) {
  return rows.map((row) => controller.getCellDisplayValue({ row, col }));
}

configureWasm();

test("cut then paste moves values and formulas, clears the source and undoes in one step", async () => {
  const { app, controller } = await mountController();
  try {
    assert.equal(controller.cutSelection(), false, "an empty selection cannot be cut");

    // Scratch block in K1:K4 so the assertions are independent of the sample.
    setRangeValues(controller, 10, ["alpha", "beta", 2, "=K3*2"]);
    controller.selectRange({ start: { row: 0, col: 10 }, end: { row: 3, col: 10 } });

    assert.equal(controller.cutSelection(), true, "cutSelection reports a pending cut");
    assert.deepEqual(
      controller.cutState,
      {
        range: { start: { row: 0, col: 10 }, end: { row: 3, col: 10 } },
        sheetIndex: controller.activeSheet.workbookSheetIndex,
      },
      "the cut state names the source range and its sheet",
    );

    controller.selectCell({ row: 0, col: 12 });
    assert.equal(await controller.pasteFromClipboard(), true, "the cut payload pastes");

    assert.deepEqual(displayValues(controller, 12, [0, 1]), ["alpha", "beta"]);
    assert.equal(controller.getCellDisplayValue({ row: 2, col: 12 }), "2");
    assert.equal(
      controller.getCellFormula({ row: 3, col: 12 }),
      "=M3*2",
      "a moved formula shifts by the paste distance",
    );
    assert.deepEqual(
      displayValues(controller, 10, [0, 1]),
      ["", ""],
      "the cut source is cleared by the move",
    );
    assert.equal(controller.getCellDisplayValue({ row: 2, col: 10 }), "");
    assert.equal(controller.getCellFormula({ row: 3, col: 10 }), "", "the source formula is cleared too");
    assert.equal(controller.cutState, null, "a successful paste ends the cut state");

    controller.undo();
    await waitFor(
      () => controller.getCellDisplayValue({ row: 0, col: 10 }) === "alpha",
      "one undo restores the cut source",
    );
    assert.deepEqual(
      displayValues(controller, 10, [0, 1]),
      ["alpha", "beta"],
      "one undo restores every source value",
    );
    assert.equal(controller.getCellFormula({ row: 3, col: 10 }), "=K3*2");
    assert.deepEqual(
      displayValues(controller, 12, [0, 1]),
      ["", ""],
      "the same undo removes the pasted block",
    );

    // One history entry means one redo puts the whole move back, including the
    // cleared source rather than a half-applied paste.
    await waitFor(() => controller.canRedo, "the move is redoable");
    controller.redo();
    await waitFor(
      () => controller.getCellDisplayValue({ row: 0, col: 12 }) === "alpha",
      "one redo re-applies the move",
    );
    assert.equal(controller.getCellDisplayValue({ row: 0, col: 10 }), "");
    assert.equal(controller.getCellFormula({ row: 3, col: 12 }), "=M3*2");
  } finally {
    app.unmount();
  }
});

test("Ctrl/Cmd+X cuts from the grid and Escape cancels the pending move", async () => {
  const { app, controller, grid } = await mountGrid();
  try {
    assert.ok(grid, "the grid renders");
    controller.setCellValue({ row: 0, col: 10 }, "alpha");
    controller.selectRange({ start: { row: 0, col: 10 }, end: { row: 1, col: 10 } });
    await vue.nextTick();

    const cut = keydownEvent("x", { ctrlKey: true });
    grid.props.onKeydown(cut.event);
    await vue.nextTick();
    assert.equal(cut.preventedCount(), 1, "the grid owns Ctrl+X");
    assert.deepEqual(
      controller.cutState?.range,
      { start: { row: 0, col: 10 }, end: { row: 1, col: 10 } },
    );
    assert.equal(grid.props["data-cut-range"], "K1:K2", "the grid exposes the marquee range");

    const escape = keydownEvent("Escape");
    grid.props.onKeydown(escape.event);
    await vue.nextTick();
    assert.equal(escape.preventedCount(), 1, "Escape is consumed while a cut is pending");
    assert.equal(controller.cutState, null, "Escape drops the cut state");
    assert.equal(grid.props["data-cut-range"], undefined, "the marquee disappears");
    assert.equal(
      controller.getCellDisplayValue({ row: 0, col: 10 }),
      "alpha",
      "cancelling leaves the source untouched",
    );
    assert.equal(
      await controller.pasteFromClipboard(),
      false,
      "a cancelled cut has no internal payload left to paste",
    );

    // macOS uses the meta key for the same command.
    const metaCut = keydownEvent("x", { metaKey: true });
    grid.props.onKeydown(metaCut.event);
    await vue.nextTick();
    assert.ok(controller.cutState, "Cmd+X starts a cut");
    assert.equal(controller.cancelCut(), true);
    assert.equal(controller.cancelCut(), false, "cancelling twice is a no-op");

    // A plain Escape with no pending cut collapses a range onto its active
    // cell; a single-cell selection leaves the key to the browser.
    const collapseEscape = keydownEvent("Escape");
    grid.props.onKeydown(collapseEscape.event);
    await vue.nextTick();
    assert.equal(collapseEscape.preventedCount(), 1, "Escape collapses the range left by the cancelled cut");
    assert.deepEqual(
      controller.selection,
      { start: { row: 1, col: 10 }, end: { row: 1, col: 10 } },
      "the range collapses onto the active cell",
    );

    const idleEscape = keydownEvent("Escape");
    grid.props.onKeydown(idleEscape.event);
    assert.equal(idleEscape.preventedCount(), 0, "a single cell leaves Escape alone");
  } finally {
    app.unmount();
  }
});

test("a read-only grid refuses Ctrl+X", async () => {
  const { app, controller, grid } = await mountGrid({ readOnly: true });
  try {
    controller.selectCell({ row: 0, col: 10 });
    await vue.nextTick();
    const cut = keydownEvent("x", { ctrlKey: true });
    grid.props.onKeydown(cut.event);
    await vue.nextTick();
    assert.equal(cut.preventedCount(), 1, "the shortcut is still swallowed");
    assert.equal(controller.cutState, null, "read-only viewers cannot cut");
  } finally {
    app.unmount();
  }
});

test("the viewer cancels a pending cut on Escape even outside the grid", async () => {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(),
        fileName: "sales-table.xlsx",
        skipXmlParsing: true,
        useWorker: false,
      });
      return () => vue.h(XlsxViewer, {
        controller,
        showDefaultToolbar: false,
        showFormulaBar: false,
        showImages: false,
        showRibbon: false,
        showSheetTabs: false,
      });
    },
  });
  const mounted = await mount(Harness);
  try {
    await waitFor(
      () => controller.sourceState === "ready" || controller.error,
      "XLSX viewer ready",
    );
    assert.ifError(controller.error);
    const viewer = findByTestId(mounted.root, "xlsx-viewer");
    assert.ok(viewer, "the viewer renders");
    controller.selectCell({ row: 0, col: 10 });
    assert.equal(controller.cutSelection(), true);
    await vue.nextTick();

    const escape = keydownEvent("Escape", {
      target: { closest: (selector) => (selector.includes("input") ? {} : null) },
    });
    viewer.props.onKeydown(escape.event);
    await vue.nextTick();
    assert.equal(escape.preventedCount(), 0, "the formula bar keeps Escape while it is open");
    assert.ok(controller.cutState);

    const cancel = keydownEvent("Escape");
    viewer.props.onKeydown(cancel.event);
    await vue.nextTick();
    assert.equal(cancel.preventedCount(), 1, "the viewer consumes the cancel key");
    assert.equal(controller.cutState, null);
  } finally {
    mounted.app.unmount();
  }
});

test("a copy or a formats-only paste consumes the cut without moving anything", async () => {
  const { app, controller } = await mountController();
  try {
    setRangeValues(controller, 10, ["alpha", "beta"]);
    controller.selectRange({ start: { row: 0, col: 10 }, end: { row: 1, col: 10 } });

    assert.equal(controller.cutSelection(), true);
    assert.ok(controller.cutState);
    await controller.copySelectionToClipboard();
    assert.equal(controller.cutState, null, "a fresh copy ends the pending cut");

    assert.equal(controller.cutSelection(), true);
    controller.selectCell({ row: 0, col: 12 });
    assert.equal(
      await controller.pasteSpecialFromClipboard({ mode: "formats" }),
      true,
      "formats-only paste applies the styles",
    );
    assert.equal(controller.cutState, null, "formats-only paste consumes the cut");
    assert.deepEqual(
      displayValues(controller, 10, [0, 1]),
      ["alpha", "beta"],
      "a formats-only paste never clears the cut source",
    );
  } finally {
    app.unmount();
  }
});

test("a cut survives a sheet switch but only marks its own sheet", async () => {
  const { app, controller, grid } = await mountGrid({ sample: "financial-model.xlsx" });
  try {
    setRangeValues(controller, 10, ["alpha", "beta"]);
    controller.selectRange({ start: { row: 0, col: 10 }, end: { row: 1, col: 10 } });
    assert.equal(controller.cutSelection(), true);
    const sourceSheet = controller.activeSheet.workbookSheetIndex;
    await vue.nextTick();
    assert.equal(grid.props["data-cut-range"], "K1:K2");

    controller.setActiveSheetIndex(1);
    await vue.nextTick();
    assert.equal(
      controller.cutState?.sheetIndex,
      sourceSheet,
      "the pending cut still belongs to its source sheet",
    );
    assert.equal(grid.props["data-cut-range"], undefined, "another sheet shows no marquee");
    assert.notEqual(controller.activeSheet.workbookSheetIndex, sourceSheet);

    controller.selectCell({ row: 0, col: 12 });
    assert.equal(await controller.pasteFromClipboard(), true);
    assert.deepEqual(displayValues(controller, 12, [0, 1]), ["alpha", "beta"]);
    assert.equal(controller.cutState, null);

    controller.setActiveSheetIndex(0);
    await vue.nextTick();
    assert.deepEqual(
      displayValues(controller, 10, [0, 1]),
      ["", ""],
      "the move clears the source on the sheet it came from",
    );
  } finally {
    app.unmount();
  }
});

test("Escape cancels a cut that was made on another sheet", async () => {
  const { app, controller, grid } = await mountGrid({ sample: "financial-model.xlsx" });
  try {
    setRangeValues(controller, 10, ["alpha"]);
    controller.selectRange({ start: { row: 0, col: 10 }, end: { row: 0, col: 10 } });
    assert.equal(controller.cutSelection(), true);

    controller.setActiveSheetIndex(1);
    await vue.nextTick();
    assert.equal(grid.props["data-cut-range"], undefined, "the marquee stays on the source sheet");

    const escape = keydownEvent("Escape");
    grid.props.onKeydown(escape.event);
    await vue.nextTick();
    assert.equal(escape.preventedCount(), 1);
    assert.equal(controller.cutState, null, "Escape cancels the cut from any sheet");
  } finally {
    app.unmount();
  }
});

test("the selection overlay announces the pending cut only on its sheet", async () => {
  const controller = vue.reactive({
    activeSheet: { workbookSheetIndex: 0 },
    activeCellAddress: "A1",
    selectedRangeAddress: "K1:K2",
    cutState: { range: { start: { row: 0, col: 10 }, end: { row: 1, col: 10 } }, sheetIndex: 0 },
  });
  const mounted = await mount(XlsxSelectionOverlay, { controller });
  try {
    const overlay = () => walk(mounted.root)
      .find((node) => String(node.props?.class ?? "").includes("xlsx-selection-overlay"));
    assert.ok(overlay(), "the overlay renders");
    assert.equal(overlay().props["data-cut-address"], "K1:K2");
    assert.match(textContent(mounted.root), /Cut K1:K2/, "the cut is announced");

    controller.cutState = {
      range: { start: { row: 0, col: 10 }, end: { row: 1, col: 10 } },
      sheetIndex: 1,
    };
    await vue.nextTick();
    assert.equal(
      overlay().props["data-cut-address"],
      undefined,
      "a cut on another sheet is not announced here",
    );
    assert.match(textContent(mounted.root), /Current selection K1:K2/);
  } finally {
    mounted.app.unmount();
  }
});

test("the grid paints a moving dashed marquee for the pending cut", () => {
  const ants = functionText("drawCutAntsBorder");
  assert.match(ants, /CUT_ANTS_DASH/, "the cut outline uses the shared dash pattern");
  assert.match(ants, /dashOffset/, "the dash phase moves between frames");
  assert.match(ants, /#ffffff|#18181b/, "a halo keeps the dashes readable over fills");

  const border = functionText("drawRangeBorder");
  assert.match(border, /setLineDash/, "the outline is dashed like Excel's marquee");
  assert.match(border, /lineDashOffset/);

  const borders = functionText("paintSelectionBorders");
  assert.match(borders, /cutRange\.value/, "the marquee is gated on the cut state");
  assert.match(borders, /drawCutAntsBorder\(ctx, cut, color\)/);
  assert.match(borders, /ctx\.clip\(\)/, "the marquee is clipped per freeze pane");

  assert.match(functionText("scheduleCutAnts"), /requestAnimationFrame/);
  assert.match(functionText("scheduleCutAnts"), /CUT_ANTS_STEP_MS/);
  assert.match(GRID_SOURCE, /:data-cut-range="cutRangeAddress/, "the grid exposes the marquee range");
});
