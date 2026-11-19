import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  walk,
  waitFor,
} from "./vue-test-renderer.mjs";

// Issue #22 acceptance: defined names stop being write-only. The controller
// lists them (workbook and sheet scope) and the formula bar's name box
// navigates to the range a name resolves to.

globalThis.ResizeObserver ??= class ResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
};
globalThis.Document ??= class Document {};
globalThis.ShadowRoot ??= class ShadowRoot {};
globalThis.window ??= {
  addEventListener() {},
  cancelAnimationFrame() {},
  clearTimeout: (id) => clearTimeout(id),
  devicePixelRatio: 1,
  removeEventListener() {},
  requestAnimationFrame: () => 1,
  setTimeout: (fn, ms) => setTimeout(fn, ms),
};
if (typeof Object.prototype.getContext !== "function") {
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value: () => null,
  });
}

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const xmldom = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= xmldom.DOMParser;
globalThis.Node ??= xmldom.Node;
globalThis.XMLSerializer ??= xmldom.XMLSerializer;

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { XlsxFormulaBar, useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

const wasmBytes = readFileSync(
  new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
);
setWasmSource(wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength));

function workbookBuffer() {
  const bytes = readFileSync(
    new URL("../../test-data/corpus/xlsx-pivot-slicers.xlsx", import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function mountNameBox(t, options = {}) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(),
        fileName: "xlsx-pivot-slicers.xlsx",
        skipXmlParsing: false,
        useWorker: false,
        ...options,
      });
      return () => vue.h(XlsxFormulaBar, { controller, readOnly: false });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "workbook loaded",
  );
  assert.ifError(controller.error);
  t.after(() => mounted.app.unmount());
  return {
    controller,
    input: findByTestId(mounted.root, "xlsx-name-box"),
    root: mounted.root,
  };
}

function keydown(target, key) {
  target.props.onKeydown({
    key,
    preventDefault() {},
    target: { blur() {} },
  });
}

test("the name box lists workbook and sheet-scoped defined names", async (t) => {
  const { input, root } = await mountNameBox(t);
  input.props.onFocus();
  await vue.nextTick();
  const list = findByTestId(root, "xlsx-name-box-list");
  assert.ok(list, "defined-name list opens on focus");
  const options = walk(list).filter((node) => node.props?.["data-name-box-option"] !== undefined);
  assert.deepEqual(options.map((node) => node.props["data-name"]), [
    "CurrentProduct",
    "RevenueTotal",
    "SalesData",
  ]);
  assert.deepEqual(options.map((node) => node.props["data-scope"]), ["sheet", "workbook", "workbook"]);
  assert.equal(options[0].props["data-sheet-name"], "Data");
});

test("choosing a defined name selects its range", async (t) => {
  const { controller, input, root } = await mountNameBox(t);
  input.props.onFocus();
  await vue.nextTick();
  const options = walk(findByTestId(root, "xlsx-name-box-list"))
    .filter((node) => node.props?.["data-name-box-option"] !== undefined);
  const salesData = options.find((node) => node.props["data-name"] === "SalesData");
  salesData.props.onClick();
  await vue.nextTick();
  assert.equal(controller.selectedRangeAddress, "A2:C6");
});

test("typing a defined name and pressing Enter navigates", async (t) => {
  const { controller, input, root } = await mountNameBox(t);
  input.props.onFocus();
  input.props.onInput({ target: { value: "CurrentProduct" } });
  await vue.nextTick();
  const options = walk(findByTestId(root, "xlsx-name-box-list"))
    .filter((node) => node.props?.["data-name-box-option"] !== undefined);
  assert.deepEqual(options.map((node) => node.props["data-name"]), ["CurrentProduct"]);
  keydown(input, "ArrowDown");
  keydown(input, "Enter");
  await vue.nextTick();
  assert.equal(controller.selectedRangeAddress, "B5");
});

test("the name box also accepts a typed A1 range", async (t) => {
  const { controller, input } = await mountNameBox(t);
  input.props.onFocus();
  input.props.onInput({ target: { value: "B2:D3" } });
  keydown(input, "Enter");
  await vue.nextTick();
  assert.equal(controller.selectedRangeAddress, "B2:D3");
});

test("Escape restores the active-cell address", async (t) => {
  const { controller, input } = await mountNameBox(t);
  controller.selectCell({ col: 1, row: 1 });
  await vue.nextTick();
  input.props.onFocus();
  input.props.onInput({ target: { value: "SalesData" } });
  keydown(input, "Escape");
  await vue.nextTick();
  assert.equal(input.props.value, "B2");
});
