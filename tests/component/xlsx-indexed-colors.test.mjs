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

// Issue #105 acceptance: indexed and custom-palette cells must paint the same
// colour in the viewer as Excel does. The unit test covers the parser and
// resolver; this test follows the palette through the controller into the DOM
// grid, whose style resolution is shared with the canvas renderer.

if (typeof Object.prototype.getContext !== "function") {
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value: () => null,
  });
}

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { XlsxGrid, useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

const VIEWPORT = { height: 600, width: 1_600 };

function configureWasm() {
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}

function corpusBuffer(name) {
  const bytes = readFileSync(new URL(`../../test-data/corpus/${name}`, import.meta.url));
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
    devicePixelRatio: 1,
    removeEventListener() {},
    requestAnimationFrame: (callback) => frames.push(callback),
    setInterval: globalThis.setInterval,
    setTimeout: globalThis.setTimeout,
    clearInterval: globalThis.clearInterval,
    clearTimeout: globalThis.clearTimeout,
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

async function mountGrid(t, fileName) {
  const { flushFrames, restore } = installGridEnvironment();
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: corpusBuffer(fileName),
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
  return mounted.root;
}

function cell(root, address) {
  return walk(root).find(
    (node) => node.props?.role === "gridcell" && node.props?.["data-cell"] === address,
  );
}

configureWasm();

test("standard indexed fills and fonts paint their Excel colours", async (t) => {
  const root = await mountGrid(t, "xlsx-indexed-colors.xlsx");

  assert.equal(cell(root, "A1").props.style.backgroundColor, "#ff0000");
  assert.equal(cell(root, "A2").props.style.backgroundColor, "#c0c0c0");
  assert.equal(cell(root, "A3").props.style.backgroundColor, "#969696");
  assert.equal(cell(root, "A4").props.style.backgroundColor, "#333333");
  assert.equal(cell(root, "A5").props.style.backgroundColor, "#ffffff");
  assert.equal(cell(root, "A6").props.style.backgroundColor, "#000000");
  assert.equal(cell(root, "A7").props.style.backgroundColor, "#ffffff");
  assert.equal(cell(root, "A8").props.style.backgroundColor, "#000000");
  assert.equal(cell(root, "C1").props.style.color, "#ff0000");
});

test("a workbook indexedColors override paints in the grid", async (t) => {
  const root = await mountGrid(t, "xlsx-indexed-custom-palette.xlsx");

  assert.equal(cell(root, "A1").props.style.backgroundColor, "#ff00ff");
  assert.equal(cell(root, "A2").props.style.backgroundColor, "#112233");
  assert.equal(
    cell(root, "A3").props.style.backgroundColor,
    "#ffffff",
    "indices the workbook does not override keep the standard palette",
  );
  assert.equal(cell(root, "B1").props.style.color, "#112233");
});
