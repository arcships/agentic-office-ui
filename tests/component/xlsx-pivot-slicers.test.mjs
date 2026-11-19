import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

// Issue #22 acceptance: the viewer exposes parsed pivot tables, slicers and
// defined names, and slicers (drawing objects) appear on the sheet instead of
// being silently dropped.

if (typeof Object.prototype.getContext !== "function") {
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value: () => null,
  });
}

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const { strFromU8, unzipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);
const xmldom = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= xmldom.DOMParser;
globalThis.Node ??= xmldom.Node;
globalThis.XMLSerializer ??= xmldom.XMLSerializer;

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { XlsxDrawingLayer, XlsxGrid, XlsxViewer, useXlsxViewerController } = await importFromDemo(
  "@arcships/vue-xlsx",
);

const VIEWPORT = { height: 700, width: 1_600 };

function configureWasm() {
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}

function workbookBuffer(fileName, directory) {
  const bytes = readFileSync(new URL(`../../${directory}/${fileName}`, import.meta.url));
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

async function mountWorkbook(t, fileName, directory = "test-data/corpus") {
  const { flushFrames, restore } = installGridEnvironment();
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(fileName, directory),
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

function sheetData(overrides = {}) {
  const visibleCols = Array.from({ length: 20 }, (_, index) => index);
  const visibleRows = Array.from({ length: 30 }, (_, index) => index);
  return {
    colWidthOverridesPx: {},
    colWidths: visibleCols.map(() => 80),
    defaultColWidthPx: 80,
    defaultRowHeightPx: 24,
    freezePanes: null,
    name: "Data",
    rowHeightOverridesPx: {},
    rowHeights: visibleRows.map(() => 24),
    showGridLines: true,
    visibleCols,
    visibleRows,
    ...overrides,
  };
}

function simpleController(overrides = {}) {
  return {
    activeSheet: sheetData(),
    charts: [],
    formControls: [],
    images: [],
    readOnly: true,
    selectedChartId: null,
    selectedImageId: null,
    selectedShapeId: null,
    shapes: [],
    slicers: [],
    zoomScale: 100,
    deleteShape() {},
    moveShapeBy() {},
    resizeShapeBy() {},
    selectShape() {},
    setCellValue() {},
    ...overrides,
  };
}

async function mountSlicerLayer(t, controller) {
  const { restore } = installGridEnvironment();
  t.after(restore);
  const mounted = await mount(vue.defineComponent({
    setup() {
      return () => vue.h(XlsxDrawingLayer, {
        controller,
        scrollLeft: 0,
        scrollTop: 0,
        viewportHeight: VIEWPORT.height,
        viewportWidth: VIEWPORT.width,
      });
    },
  }));
  t.after(() => mounted.app.unmount());
  await vue.nextTick();
  return mounted.root;
}

configureWasm();

test("the controller exposes pivot tables, slicers and defined names", async (t) => {
  const { controller, root } = await mountWorkbook(t, "xlsx-pivot-slicers.xlsx");

  assert.equal(controller.hasPivotTable, true);
  assert.equal(controller.pivotTables.length, 1);
  assert.equal(controller.pivotTables[0].name, "PivotTable1");
  assert.deepEqual(controller.pivotTables[0].rowFields, [0]);
  assert.equal(controller.pivotTables[0].dataFields[0].name, "Sum of Revenue");

  assert.equal(controller.slicers.length, 1);
  assert.equal(controller.slicers[0].name, "Region");
  assert.equal(controller.slicers[0].items[1].caption, "West");
  assert.equal(controller.slicers[0].items[1].selected, true);

  assert.deepEqual(controller.definedNames.map((entry) => entry.name), [
    "CurrentProduct",
    "RevenueTotal",
    "SalesData",
  ]);
  assert.equal(controller.definedNames[2].scope, "workbook");
  assert.equal(controller.definedNames[0].scope, "sheet");
  assert.equal(controller.definedNames[0].sheetName, "Data");

  // The cached pivot output is still painted as ordinary cells.
  assert.equal(
    textContent(walk(root).find((node) => node.props?.["data-cell"] === "J10")),
    "Sum of Revenue",
  );
  const grid = findByTestId(root, "xlsx-grid");
  assert.equal(grid.props["data-pivot-count"], 1);
  assert.equal(grid.props["data-pivot-names"], "PivotTable1");
  const badge = findByTestId(root, "xlsx-pivot-badge");
  assert.ok(badge, "the pivot table shows its static-preview badge");
  assert.equal(badge.props["data-pivot-name"], "PivotTable1");
  assert.match(badge.props.title, /static/i);
});

function installDownloadCapture(t) {
  const downloads = [];
  const blobs = new Map();
  const originalDocument = globalThis.document;
  const originalCreateObjectUrl = URL.createObjectURL;
  const originalRevokeObjectUrl = URL.revokeObjectURL;
  let nextObjectUrl = 0;
  globalThis.document = {
    body: { append() {} },
    createElement(tagName) {
      return {
        tagName,
        href: "",
        download: "",
        click() {
          downloads.push({ blob: blobs.get(this.href), fileName: this.download });
        },
        remove() {},
      };
    },
  };
  URL.createObjectURL = (blob) => {
    const objectUrl = `blob:test-${nextObjectUrl += 1}`;
    blobs.set(objectUrl, blob);
    return objectUrl;
  };
  URL.revokeObjectURL = () => {};
  t.after(() => {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    if (originalCreateObjectUrl === undefined) delete URL.createObjectURL;
    else URL.createObjectURL = originalCreateObjectUrl;
    if (originalRevokeObjectUrl === undefined) delete URL.revokeObjectURL;
    else URL.revokeObjectURL = originalRevokeObjectUrl;
  });
  return downloads;
}

test("workbooks without pivots report no pivot and no slicers", async (t) => {
  const { controller } = await mountWorkbook(t, "financial-model.xlsx", "apps/demo/public/samples");
  assert.equal(controller.hasPivotTable, false);
  assert.deepEqual(controller.pivotTables, []);
  assert.deepEqual(controller.slicers, []);
});

test("controller export keeps pivot, slicer and defined-name metadata on reload", async (t) => {
  const { controller } = await mountWorkbook(t, "xlsx-pivot-slicers.xlsx");
  const downloads = installDownloadCapture(t);
  controller.exportXlsx();
  await waitFor(() => downloads.length > 0, "workbook export");
  const bytes = new Uint8Array(await downloads[0].blob.arrayBuffer());
  const archive = unzipSync(bytes);
  for (const part of [
    "xl/pivotTables/pivotTable1.xml",
    "xl/pivotCache/pivotCacheDefinition1.xml",
    "xl/slicers/slicer1.xml",
    "xl/slicerCaches/slicerCache1.xml",
  ]) {
    assert.ok(archive[part], `${part} survives controller export`);
  }
  assert.match(strFromU8(archive["xl/workbook.xml"]), /name="SalesData"/);

  let reloaded;
  const mounted = await mount(vue.defineComponent({
    setup() {
      reloaded = useXlsxViewerController({
        file: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
        fileName: "reloaded.xlsx",
        skipXmlParsing: false,
        useWorker: false,
      });
      return () => vue.h("div");
    },
  }));
  await waitFor(() => reloaded.sourceState === "ready" || reloaded.error, "reloaded workbook");
  assert.ifError(reloaded.error);
  assert.equal(reloaded.hasPivotTable, true);
  assert.equal(reloaded.pivotTables[0].dataFields[0].name, "Sum of Revenue");
  assert.equal(reloaded.slicers[0].items[1].caption, "West");
  assert.equal(reloaded.slicers[0].items[1].selected, true);
  assert.deepEqual(reloaded.definedNames.map((entry) => entry.name), [
    "CurrentProduct",
    "RevenueTotal",
    "SalesData",
  ]);
  assert.ok(controller.pivotTables.length > 0);
  await mounted.app.unmount();
});

test("the full viewer mounts the pivot badge and slicer card", async (t) => {
  const { flushFrames, restore } = installGridEnvironment();
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer("xlsx-pivot-slicers.xlsx", "test-data/corpus"),
        fileName: "xlsx-pivot-slicers.xlsx",
        skipXmlParsing: false,
        useWorker: false,
      });
      return () => vue.h(XlsxViewer, { controller, readOnly: true, renderMode: "dom" });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "viewer loaded",
  );
  assert.ifError(controller.error);
  flushFrames();
  await vue.nextTick();
  try {
    assert.equal(findByTestId(mounted.root, "xlsx-pivot-badge").props["data-pivot-name"], "PivotTable1");
    assert.equal(findByTestId(mounted.root, "xlsx-slicer").props["data-slicer-name"], "Region");
  } finally {
    await mounted.app.unmount();
    restore();
  }
});

test("a slicer card renders with its caption, items and selection", async (t) => {
  const anchor = {
    from: { col: 13, colOffsetEmu: 0, row: 2, rowOffsetEmu: 0 },
    kind: "two-cell",
    to: { col: 16, colOffsetEmu: 0, row: 14, rowOffsetEmu: 0 },
  };
  const controller = simpleController({
    slicers: [{
      anchor,
      caption: "Region",
      cacheName: "Slicer_Region",
      columnCount: 1,
      items: [
        { caption: "East", key: "0", selected: false },
        { caption: "West", key: "1", selected: true },
        { caption: "North", key: "2", selected: false },
      ],
      name: "Region",
      pivotTables: ["PivotTable1"],
      sourceName: "Region",
      style: "SlicerStyleLight1",
      workbookSheetIndex: 0,
    }],
  });
  const root = await mountSlicerLayer(t, controller);
  const layer = findByTestId(root, "xlsx-drawing-layer");
  assert.equal(layer.props["data-slicer-count"], 1);
  const slicer = findByTestId(root, "xlsx-slicer");
  assert.ok(slicer, "slicer card renders");
  assert.equal(slicer.props["data-slicer-name"], "Region");
  assert.equal(slicer.props["data-slicer-count"], 3);
  assert.equal(slicer.props["aria-label"], "Region");
  const items = walk(slicer).filter((node) => node.props?.["data-slicer-item"] !== undefined);
  assert.deepEqual(items.map((node) => textContent(node)), ["East", "West", "North"]);
  assert.deepEqual(items.map((node) => node.props["data-selected"]), ["false", "true", "false"]);
});

test("a slicer without a drawing anchor is exposed but not painted", async (t) => {
  const controller = simpleController({
    slicers: [{ items: [], name: "Region", pivotTables: [], workbookSheetIndex: 0 }],
  });
  const root = await mountSlicerLayer(t, controller);
  assert.equal(findByTestId(root, "xlsx-slicer"), undefined);
  const layer = findByTestId(root, "xlsx-drawing-layer");
  assert.equal(layer.props["data-slicer-count"], 0);
});
