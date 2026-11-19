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

// Issue #22 acceptance: `sales-table.xlsx` names TableStyleMedium2 with row
// stripes. The DOM grid shares its style resolution with the canvas painter, so
// asserting the computed cell styles proves the table style reaches both.

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
const { XlsxGrid, useXlsxViewerController } = await importFromDemo(
  "@arcships/vue-xlsx",
);

const VIEWPORT = { height: 600, width: 1_600 };

function configureWasm() {
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}

function sampleBuffer(name) {
  const bytes = new Uint8Array(
    readFileSync(new URL(`../../apps/demo/public/samples/${name}`, import.meta.url)),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function installGridEnvironment() {
  const previous = {
    Document: globalThis.Document,
    ResizeObserver: globalThis.ResizeObserver,
    ShadowRoot: globalThis.ShadowRoot,
    document: globalThis.document,
    window: globalThis.window,
  };
  const frames = [];
  globalThis.window = {
    addEventListener() {},
    cancelAnimationFrame() {},
    clearInterval: globalThis.clearInterval,
    clearTimeout: globalThis.clearTimeout,
    devicePixelRatio: 1,
    removeEventListener() {},
    requestAnimationFrame: (callback) => frames.push(callback),
    setInterval: globalThis.setInterval,
    setTimeout: globalThis.setTimeout,
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
  globalThis.Document = class Document {};
  globalThis.ShadowRoot = class ShadowRoot {};
  const flushFrames = () => {
    while (frames.length) {
      frames.shift()(0);
    }
  };
  const restore = () => {
    for (const [name, value] of Object.entries(previous)) {
      if (value === undefined) delete globalThis[name];
      else globalThis[name] = value;
    }
  };
  return { flushFrames, restore };
}

async function mountGrid(t, fileName = "sales-table.xlsx") {
  const { flushFrames, restore } = installGridEnvironment();
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: sampleBuffer(fileName),
        fileName,
        skipXmlParsing: false,
        useWorker: false,
      });
      return () => vue.h(XlsxGrid, {
        controller,
        readOnly: true,
        renderMode: "dom",
      });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    `${fileName} loaded`,
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
  t.after(() => mounted.app.unmount());
  t.after(restore);
  grid.props.onScroll({ type: "scroll" });
  flushFrames();
  await vue.nextTick();
  return { controller, root: mounted.root };
}

function cell(root, address) {
  return walk(root).find(
    (node) => node.props?.role === "gridcell" && node.props?.["data-cell"] === address,
  );
}

configureWasm();

test("TableStyleMedium2 paints its header row and row banding", async (t) => {
  const { controller, root } = await mountGrid(t);
  assert.equal(controller.tables.length, 1);
  assert.equal(controller.tables[0].styleInfo.name, "TableStyleMedium2");
  assert.equal(controller.tables[0].styleInfo.showRowStripes, true);

  const header = cell(root, "A1");
  assert.equal(header.props.style.backgroundColor, "#1c3a5e");
  assert.equal(header.props.style.color, "#ffffff");
  assert.equal(header.props.style.fontWeight, "700");

  // The first body row (row offset 0) is the first stripe, the second is not.
  assert.equal(cell(root, "A2").props.style.backgroundColor, "#dce6f2");
  assert.equal(cell(root, "A3").props.style.backgroundColor, "#ffffff");
  assert.equal(cell(root, "B2").props.style.backgroundColor, "#dce6f2");
  assert.equal(cell(root, "B3").props.style.backgroundColor, "#ffffff");
  assert.equal(cell(root, "H2").props.style.backgroundColor, "#dce6f2");
});

test("table banding follows the table start row and stops at its end", async (t) => {
  const { root } = await mountGrid(t);
  // The table starts at A1, so its body starts at row 2; the header keeps its
  // explicit fill and the second body row is unfilled.
  assert.equal(cell(root, "A4").props.style.backgroundColor, "#dce6f2");
  assert.equal(cell(root, "A5").props.style.backgroundColor, "#ffffff");
});
