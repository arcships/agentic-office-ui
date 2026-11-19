import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

// Only the browser plumbing the host renderer cannot provide is stubbed; the
// grid, the controller and every DOM node asserted below are real.
if (typeof Object.prototype.getContext !== "function") {
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value: () => null,
  });
}

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { XlsxGrid, useXlsxViewerController } = await importFromDemo(
  "@arcships/vue-xlsx",
);

// Grid chrome the assertions below are written against.
const ROW_HEADER_WIDTH = 48;
const HEADER_HEIGHT = 24;
const VIEWPORT = { height: 600, width: 1_600 };
const HANDLE = 7;
/** Ten pixels per character: the fake canvas measures text deterministically. */
const TEXT_WIDTH_PER_CHARACTER = 10;

function configureWasm() {
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}

function workbookBuffer() {
  const bytes = readFileSync(
    new URL("../../apps/demo/public/samples/sales-table.xlsx", import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function installGridEnvironment(t, { measureText = false } = {}) {
  const previous = {
    Document: globalThis.Document,
    ResizeObserver: globalThis.ResizeObserver,
    ShadowRoot: globalThis.ShadowRoot,
    document: globalThis.document,
    window: globalThis.window,
  };
  // Animation frames are queued so a test can run the grid's paint frame when
  // it needs the reactive scroll offsets it publishes.
  const frames = [];
  globalThis.window = {
    addEventListener() {},
    cancelAnimationFrame() {},
    devicePixelRatio: 1,
    removeEventListener() {},
    requestAnimationFrame: (callback) => frames.push(callback),
  };
  globalThis.ResizeObserver = class ResizeObserver {
    constructor(callback) {
      this.callback = callback;
    }
    observe() {
      this.callback([{ contentRect: { height: VIEWPORT.height, width: VIEWPORT.width } }]);
    }
    unobserve() {}
    disconnect() {}
  };
  // v-model's runtime-dom directive probes `rootNode instanceof Document`.
  globalThis.Document = class Document {};
  globalThis.ShadowRoot = class ShadowRoot {};
  if (measureText) {
    // The grid measures auto-fit text through a detached canvas; this one
    // reports a fixed advance so the fitted size is exact.
    const context = {
      font: "",
      measureText: (text) => ({ width: text.length * TEXT_WIDTH_PER_CHARACTER }),
    };
    globalThis.document = {
      createElement: () => ({ getContext: () => context }),
    };
  }
  const flushFrames = () => {
    while (frames.length) {
      const callback = frames.shift();
      callback(0);
    }
  };
  const restore = () => {
    for (const [name, value] of Object.entries(previous)) {
      if (name === "document") {
        if (value === undefined) delete globalThis.document;
        else globalThis.document = value;
        continue;
      }
      if (value === undefined) delete globalThis[name];
      else globalThis[name] = value;
    }
  };
  return { flushFrames, restore };
}

async function mountGrid(t, options = {}) {
  const { flushFrames, restore } = installGridEnvironment(t, {
    measureText: options.measureText === true,
  });
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(),
        fileName: "fill-handle.xlsx",
        skipXmlParsing: true,
        useWorker: false,
      });
      return () => vue.h(XlsxGrid, {
        controller,
        readOnly: options.readOnly === true,
      });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "XLSX fill-handle controller ready",
  );
  assert.ifError(controller.error);
  const grid = findByTestId(mounted.root, "xlsx-grid");
  assert.ok(grid, "grid root renders");
  grid.getBoundingClientRect = () => ({
    bottom: VIEWPORT.height,
    height: VIEWPORT.height,
    left: 0,
    right: VIEWPORT.width,
    top: 0,
    width: VIEWPORT.width,
  });
  grid.style = {};
  grid.scrollLeft = 0;
  grid.scrollTop = 0;
  grid.focus = () => {};
  grid.setPointerCapture = () => {};
  grid.hasPointerCapture = () => true;
  grid.releasePointerCapture = () => {};
  // Unmount before the DOM globals are torn down: the component's own
  // unmount hook touches `window`.
  t.after(() => mounted.app.unmount());
  t.after(restore);
  return { ...mounted, controller, flushFrames, grid };
}

// ── Sheet geometry ────────────────────────────────────────────────────
// The sample workbook ships its own column widths and row heights, so the
// pointer coordinates below come from the sheet state the grid paints from.

function axisState(controller, axis) {
  const sheet = controller.activeSheet;
  const indices = axis === "column" ? sheet.visibleCols : sheet.visibleRows;
  const sizes = axis === "column" ? sheet.colWidths : sheet.rowHeights;
  const fallback = (axis === "column" ? sheet.defaultColWidthPx : sheet.defaultRowHeightPx)
    || (axis === "column" ? 80 : 24);
  return { fallback, indices, sizes };
}

function displayIndex(controller, axis, actualIndex) {
  const { indices } = axisState(controller, axis);
  const index = indices.indexOf(actualIndex);
  return index >= 0 ? index : actualIndex;
}

function axisSize(controller, axis, actualIndex) {
  const { fallback, sizes } = axisState(controller, axis);
  return sizes[displayIndex(controller, axis, actualIndex)] || fallback;
}

function axisOffset(controller, axis, actualIndex) {
  const { fallback, sizes } = axisState(controller, axis);
  const last = displayIndex(controller, axis, actualIndex);
  let total = 0;
  for (let index = 0; index < last; index += 1) total += sizes[index] || fallback;
  return total;
}

/** Client coordinates of the centre of a sheet cell. */
function cellCenter(controller, row, col) {
  return {
    clientX: ROW_HEADER_WIDTH + axisOffset(controller, "column", col) + axisSize(controller, "column", col) / 2,
    clientY: HEADER_HEIGHT + axisOffset(controller, "row", row) + axisSize(controller, "row", row) / 2,
  };
}

/** Client coordinates of a cell's bottom-right corner. */
function cellCorner(controller, row, col) {
  return {
    x: ROW_HEADER_WIDTH + axisOffset(controller, "column", col) + axisSize(controller, "column", col),
    y: HEADER_HEIGHT + axisOffset(controller, "row", row) + axisSize(controller, "row", row),
  };
}

/** Records the resize requests the grid makes on the controller. */
function spyResize(controller) {
  const calls = [];
  for (const kind of ["resizeColumn", "resizeRow"]) {
    const original = controller[kind];
    controller[kind] = (index, size) => {
      calls.push({ index, kind, size });
      return original.call(controller, index, size);
    };
  }
  return calls;
}

function pointerEvent(init = {}) {
  return {
    button: 0,
    pointerId: 1,
    preventDefault() {
      this.defaultPrevented = true;
    },
    stopPropagation() {
      this.propagationStopped = true;
    },
    type: "pointermove",
    ...init,
  };
}

function fillHandle(root) {
  return findByTestId(root, "xlsx-fill-handle");
}

function displayValues(controller, cells) {
  return cells.map((cell) => controller.getCellDisplayValue(cell));
}

function columnWidth(controller, col) {
  const { sizes } = axisState(controller, "column");
  return sizes[displayIndex(controller, "column", col)] ?? null;
}

function rowHeight(controller, row) {
  const { sizes } = axisState(controller, "row");
  return sizes[displayIndex(controller, "row", row)] ?? null;
}

configureWasm();

test("the grid draws a fill handle at the selection's bottom-right corner", async (t) => {
  const { controller, flushFrames, grid } = await mountGrid(t);
  controller.setCellValue({ row: 0, col: 0 }, "1");
  controller.selectCell({ row: 0, col: 0 });
  await vue.nextTick();

  const a1 = cellCorner(controller, 0, 0);
  const handle = fillHandle(grid.getRootNode());
  assert.ok(handle, "a selection renders its fill handle");
  assert.equal(
    handle.props["aria-label"],
    "Fill handle: drag to fill the selection's pattern into more cells",
  );
  assert.equal(String(handle.props.type), "button");
  assert.equal(handle.props.style.width, `${HANDLE}px`);
  assert.equal(handle.props.style.height, `${HANDLE}px`);
  assert.equal(handle.props.style.left, `${a1.x - HANDLE / 2}px`, "centred on A1's right edge");
  assert.equal(handle.props.style.top, `${a1.y - HANDLE / 2}px`, "centred on A1's bottom edge");

  controller.selectRange({ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } });
  await vue.nextTick();
  const b2 = cellCorner(controller, 1, 1);
  const blockHandle = fillHandle(grid.getRootNode());
  assert.equal(blockHandle.props.style.left, `${b2.x - HANDLE / 2}px`);
  assert.equal(blockHandle.props.style.top, `${b2.y - HANDLE / 2}px`);

  // Scrolling the corner out of the body hides the handle instead of pinning
  // it to the viewport edge.
  grid.scrollLeft = VIEWPORT.width;
  grid.props.onScroll();
  flushFrames();
  await vue.nextTick();
  assert.equal(fillHandle(grid.getRootNode()), undefined);
});

test("a read-only grid has no fill handle", async (t) => {
  const { controller, grid } = await mountGrid(t, { readOnly: true });
  controller.setCellValue({ row: 0, col: 0 }, "1");
  controller.selectCell({ row: 0, col: 0 });
  await vue.nextTick();
  assert.equal(fillHandle(grid.getRootNode()), undefined);
});

test("dragging the handle fills the pattern in one undoable edit", async (t) => {
  const { controller, grid } = await mountGrid(t);
  // A1:A2 is the pattern; the drag fills it down over A3:A4.
  controller.setCellValue({ row: 0, col: 0 }, "1");
  controller.setCellValue({ row: 1, col: 0 }, "2");
  controller.selectRange({ start: { row: 0, col: 0 }, end: { row: 1, col: 0 } });
  await vue.nextTick();
  const before = displayValues(controller, [
    { row: 2, col: 0 },
    { row: 3, col: 0 },
  ]);

  const handle = fillHandle(grid.getRootNode());
  assert.ok(handle, "the selection's corner carries the handle");
  handle.props.onPointerdown(pointerEvent({ type: "pointerdown" }));

  const target = cellCenter(controller, 3, 0);
  grid.props.onPointermove(pointerEvent(target));
  await vue.nextTick();
  assert.equal(
    grid.props["data-fill-target"],
    "A1:A4",
    "the grid previews the range the release would fill",
  );
  assert.deepEqual(
    displayValues(controller, [
      { row: 2, col: 0 },
      { row: 3, col: 0 },
    ]),
    before,
    "dragging only previews; the cells change on release",
  );

  grid.props.onPointerup(pointerEvent({ ...target, type: "pointerup" }));
  await vue.nextTick();
  assert.equal(grid.props["data-fill-target"], undefined, "the preview ends with the drag");
  assert.deepEqual(
    displayValues(controller, [
      { row: 0, col: 0 },
      { row: 1, col: 0 },
      { row: 2, col: 0 },
      { row: 3, col: 0 },
    ]),
    ["1", "2", "1", "2"],
    "the fill tiles the source pattern down the dragged range",
  );
  assert.deepEqual(
    controller.selection,
    { start: { row: 0, col: 0 }, end: { row: 3, col: 0 } },
    "the selection grows to the filled range",
  );

  controller.undo();
  await vue.nextTick();
  assert.deepEqual(
    displayValues(controller, [
      { row: 2, col: 0 },
      { row: 3, col: 0 },
    ]),
    before,
    "one undo restores every filled cell, so the fill is a single history entry",
  );

  // The issue's own case: dragging A1's handle down fills A2:A3 with its value.
  controller.setCellValue({ row: 0, col: 0 }, "10");
  controller.selectCell({ row: 0, col: 0 });
  await vue.nextTick();
  fillHandle(grid.getRootNode()).props.onPointerdown(pointerEvent({ type: "pointerdown" }));
  grid.props.onPointermove(pointerEvent(cellCenter(controller, 2, 0)));
  await vue.nextTick();
  assert.equal(grid.props["data-fill-target"], "A1:A3");
  grid.props.onPointerup(pointerEvent({
    ...cellCenter(controller, 2, 0),
    type: "pointerup",
  }));
  await vue.nextTick();
  assert.deepEqual(
    displayValues(controller, [
      { row: 1, col: 0 },
      { row: 2, col: 0 },
    ]),
    ["10", "10"],
    "dragging A1 down copies its value into A2:A3",
  );
});

test("a fill drag that is cancelled or stays on the source writes nothing", async (t) => {
  const { controller, grid } = await mountGrid(t);
  controller.setCellValue({ row: 0, col: 0 }, "7");
  controller.selectCell({ row: 0, col: 0 });
  await vue.nextTick();
  const before = displayValues(controller, [
    { row: 1, col: 0 },
    { row: 2, col: 0 },
  ]);

  // A drag that ends on the source range is a no-op.
  fillHandle(grid.getRootNode()).props.onPointerdown(pointerEvent({ type: "pointerdown" }));
  grid.props.onPointerup(pointerEvent({ ...cellCenter(controller, 0, 0), type: "pointerup" }));
  await vue.nextTick();
  assert.equal(controller.getCellDisplayValue({ row: 1, col: 0 }), before[0]);

  // A cancelled pointer drops the preview without writing anything.
  fillHandle(grid.getRootNode()).props.onPointerdown(pointerEvent({ type: "pointerdown" }));
  grid.props.onPointermove(pointerEvent(cellCenter(controller, 2, 0)));
  await vue.nextTick();
  assert.equal(grid.props["data-fill-target"], "A1:A3");
  grid.props.onPointerup(pointerEvent({ ...cellCenter(controller, 2, 0), type: "pointercancel" }));
  await vue.nextTick();
  assert.equal(grid.props["data-fill-target"], undefined);
  assert.deepEqual(
    displayValues(controller, [
      { row: 1, col: 0 },
      { row: 2, col: 0 },
    ]),
    before,
    "cancelling a fill drag writes nothing",
  );
});

test("double-clicking a column boundary fits the column to its widest cell", async (t) => {
  const { controller, grid } = await mountGrid(t, { measureText: true });
  // Scratch column M, away from the sample's data, with one long value and one
  // short one: the fitted width must follow the long one.
  controller.setCellValue({ row: 0, col: 12 }, "Hello world");
  controller.setCellValue({ row: 2, col: 12 }, "ab");
  await waitFor(() => controller.activeSheet.maxUsedCol >= 12, "the scratch column is used");
  const original = columnWidth(controller, 12);
  assert.ok(original > 0);
  const resizes = spyResize(controller);

  const boundary = cellCorner(controller, 0, 12);
  grid.props.onDblclick({
    clientX: boundary.x,
    clientY: HEADER_HEIGHT / 2,
    preventDefault() {},
  });
  await vue.nextTick();

  // 11 characters × 10 px + both 5 px paddings + 2 px rounding room.
  assert.equal(resizes.length, 1, "one auto-fit asks for one resize");
  assert.equal(resizes[0].kind, "resizeColumn");
  assert.equal(resizes[0].index, 12, "the fitted column is the one whose boundary was hit");
  assert.equal(
    resizes[0].size,
    11 * TEXT_WIDTH_PER_CHARACTER + 12,
    "the requested width follows the painted text, not the short cell",
  );

  const fitted = columnWidth(controller, 12);
  assert.ok(fitted > original, `the column grows to fit the text (got ${fitted})`);
  assert.ok(
    fitted >= 11 * TEXT_WIDTH_PER_CHARACTER,
    "the fitted column is wide enough for the text itself",
  );

  controller.undo();
  await vue.nextTick();
  assert.equal(columnWidth(controller, 12), original, "one undo restores the previous width");
});

test("double-clicking a row boundary fits the row to its tallest cell", async (t) => {
  const { controller, grid } = await mountGrid(t, { measureText: true });
  controller.setCellValue({ row: 1, col: 12 }, "row text");
  await waitFor(() => controller.activeSheet.maxUsedCol >= 12, "the scratch cell is used");
  const original = rowHeight(controller, 1);
  const resizes = spyResize(controller);

  const boundary = cellCorner(controller, 1, 0);
  grid.props.onDblclick({
    clientX: ROW_HEADER_WIDTH / 2,
    clientY: boundary.y,
    preventDefault() {},
  });
  await vue.nextTick();

  // One 12 px line is 15 px tall plus the paint path's 4 px of vertical chrome.
  assert.equal(resizes.length, 1, "one auto-fit asks for one resize");
  assert.equal(resizes[0].kind, "resizeRow");
  assert.equal(resizes[0].index, 1, "the fitted row is the one whose boundary was hit");
  assert.ok(
    Math.abs(resizes[0].size - 19) <= 1,
    `the requested height fits one line of text (got ${resizes[0].size})`,
  );

  const fitted = rowHeight(controller, 1);
  assert.ok(fitted < original, `the row shrinks to its text (was ${original}, got ${fitted})`);

  controller.undo();
  await vue.nextTick();
  assert.equal(rowHeight(controller, 1), original, "one undo restores the previous height");
});

test("a header double-click away from a boundary resizes nothing and opens no editor", async (t) => {
  const { controller, grid } = await mountGrid(t);
  const original = columnWidth(controller, 0);

  grid.props.onDblclick({
    clientX: ROW_HEADER_WIDTH + axisSize(controller, "column", 0) / 2,
    clientY: HEADER_HEIGHT / 2,
    preventDefault() {},
  });
  await vue.nextTick();

  assert.equal(columnWidth(controller, 0), original, "a header click leaves the width alone");
  assert.equal(
    findByTestId(grid.getRootNode(), "xlsx-cell-editor"),
    undefined,
    "the header strip never opens the cell editor",
  );
});

test("double-clicking a cell body still opens the inline editor", async (t) => {
  const { controller, grid } = await mountGrid(t);
  controller.setCellValue({ row: 0, col: 0 }, "editable");
  controller.selectCell({ row: 1, col: 1 });
  await vue.nextTick();
  const width = columnWidth(controller, 0);

  grid.props.onDblclick({ ...cellCenter(controller, 0, 0), preventDefault() {} });
  await vue.nextTick();
  await vue.nextTick();

  const editor = findByTestId(grid.getRootNode(), "xlsx-cell-editor");
  assert.ok(editor, "a cell double-click still opens the editor");
  assert.equal(editor.value, "editable", "the editor is seeded with the cell's value");
  assert.equal(columnWidth(controller, 0), width, "editing does not resize the column");
});
