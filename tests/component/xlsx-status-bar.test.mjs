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
} from "./vue-test-renderer.mjs";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;
globalThis.Document ??= class Document {};
globalThis.ShadowRoot ??= class ShadowRoot {};
globalThis.ResizeObserver ??= class ResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
};
globalThis.requestAnimationFrame ??= (callback) => setTimeout(() => callback(Date.now()), 0);
globalThis.cancelAnimationFrame ??= (handle) => clearTimeout(handle);
// The grid paints its canvases in an animation frame; the host nodes have no
// 2D context and the status bar tests never assert painted pixels.
if (typeof Object.prototype.getContext !== "function") {
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value: () => null,
  });
}
globalThis.window ??= {
  addEventListener() {},
  cancelAnimationFrame: globalThis.cancelAnimationFrame,
  clearInterval: globalThis.clearInterval,
  clearTimeout: globalThis.clearTimeout,
  devicePixelRatio: 1,
  getComputedStyle: () => ({ getPropertyValue: () => "" }),
  removeEventListener() {},
  requestAnimationFrame: globalThis.requestAnimationFrame,
  setInterval: globalThis.setInterval,
  setTimeout: globalThis.setTimeout,
};

const { getSheetsWasmModule, setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { XlsxStatusBar, XlsxViewer, useXlsxViewerController } = await importFromDemo(
  "@arcships/vue-xlsx",
);

function configureWasm() {
  const bytes = readWasm();
  setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}

function readWasm() {
  return readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
}

/** Builds a real XLSX package from cell values plus optional added sheets. */
async function workbookBytes(cells, extraSheets = []) {
  const wasm = await getSheetsWasmModule();
  const workbook = new wasm.Workbook();
  const fill = (index, values) => {
    const worksheet = workbook.getSheet(index);
    try {
      for (const [address, value] of Object.entries(values)) {
        worksheet.setCell(address, value);
      }
    } finally {
      worksheet.free();
    }
  };
  fill(0, cells);
  for (const sheet of extraSheets) {
    fill(workbook.addSheet(sheet.name), sheet.cells ?? {});
  }
  const bytes = workbook.saveXlsxBytes();
  workbook.free();
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

/** Cells of a single column, e.g. `column([1, 2, 3])` becomes A1:A3. */
function column(values, columnName = "A") {
  return Object.fromEntries(values.map((value, index) => [
    `${columnName}${index + 1}`,
    value,
  ]));
}

/** Adds a worksheet-level `<autoFilter>` the engine has no setter for. */
function withAutoFilter(bytes, reference, sheetPath = "xl/worksheets/sheet1.xml") {
  const archive = unzipSync(new Uint8Array(bytes));
  const xml = strFromU8(archive[sheetPath]);
  archive[sheetPath] = strToU8(
    xml.replace("</worksheet>", `<autoFilter ref="${reference}"/></worksheet>`),
  );
  const patched = zipSync(archive);
  return patched.buffer.slice(patched.byteOffset, patched.byteOffset + patched.byteLength);
}

async function mountViewer(file, options = {}) {
  const {
    controllerOptions = {},
    viewerOptions = {},
    controller: providedController,
  } = options;
  let controller = providedController;
  const Harness = vue.defineComponent({
    setup() {
      controller ??= useXlsxViewerController({
        file,
        fileName: "status-bar.xlsx",
        useWorker: false,
        ...controllerOptions,
      });
      return () => vue.h(XlsxViewer, {
        controller,
        showDefaultToolbar: false,
        showFormulaBar: false,
        showImages: false,
        showRibbon: false,
        ...viewerOptions,
      });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "XLSX status bar controller",
  );
  assert.ifError(controller.error);
  await waitFor(() => controller.activeSheet, "an active worksheet");
  await vue.nextTick();
  return { ...mounted, controller };
}

function stat(root, testId) {
  const node = findByTestId(root, testId);
  return node ? textContent(node) : undefined;
}

function assertStats(root, expected) {
  for (const [testId, value] of Object.entries(expected)) {
    assert.equal(stat(root, testId), value, testId);
  }
}

async function selectRange(controller, start, end) {
  controller.selectRange({ start, end });
  await vue.nextTick();
}

configureWasm();

test("selecting A1:A10 shows Excel's count, sum, average, min and max", async (t) => {
  const file = await workbookBytes(column([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]));
  const viewer = await mountViewer(file);
  t.after(() => viewer.app.unmount());

  await selectRange(viewer.controller, { row: 0, col: 0 }, { row: 9, col: 0 });

  assert.equal(stat(viewer.root, "xlsx-status-range"), "A1:A10");
  assertStats(viewer.root, {
    "xlsx-status-count": "Count: 10",
    "xlsx-status-sum": "Sum: 55",
    "xlsx-status-average": "Average: 5.5",
    "xlsx-status-min": "Min: 1",
    "xlsx-status-max": "Max: 10",
    "xlsx-status-note": "Selected",
  });
  assert.equal(
    findByTestId(viewer.root, "xlsx-status-numeric-count"),
    undefined,
    "a fully numeric selection does not repeat the numeric count",
  );
  assert.equal(findByTestId(viewer.root, "xlsx-status-truncated"), undefined);
});

test("a single cell shows its address and its own totals", async (t) => {
  const file = await workbookBytes({ A1: 7 });
  const viewer = await mountViewer(file);
  t.after(() => viewer.app.unmount());

  await selectRange(viewer.controller, { row: 0, col: 0 }, { row: 0, col: 0 });

  assert.equal(stat(viewer.root, "xlsx-status-range"), "A1");
  assertStats(viewer.root, {
    "xlsx-status-count": "Count: 1",
    "xlsx-status-sum": "Sum: 7",
    "xlsx-status-average": "Average: 7",
    "xlsx-status-min": "Min: 7",
    "xlsx-status-max": "Max: 7",
  });
  assert.equal(findByTestId(viewer.root, "xlsx-status-note"), undefined);
});

test("text cells count but only numbers contribute to sum and average", async (t) => {
  const file = await workbookBytes({ A1: 1, B1: "alpha", C1: 3 });
  const viewer = await mountViewer(file);
  t.after(() => viewer.app.unmount());

  await selectRange(viewer.controller, { row: 0, col: 0 }, { row: 0, col: 2 });

  assert.equal(stat(viewer.root, "xlsx-status-range"), "A1:C1");
  assertStats(viewer.root, {
    "xlsx-status-count": "Count: 3",
    "xlsx-status-numeric-count": "Numerical count: 2",
    "xlsx-status-sum": "Sum: 4",
    "xlsx-status-average": "Average: 2",
    "xlsx-status-min": "Min: 1",
    "xlsx-status-max": "Max: 3",
  });
});

test("totals use the selected cells' number formats", async (t) => {
  const file = await workbookBytes({ A1: 1 });
  const viewer = await mountViewer(file);
  t.after(() => viewer.app.unmount());

  viewer.controller.setCellValue({ row: 0, col: 1 }, "$1,234.56");
  viewer.controller.setCellValue({ row: 1, col: 1 }, "$2,000.00");
  viewer.controller.setCellValue({ row: 0, col: 2 }, "50%");
  viewer.controller.setCellValue({ row: 1, col: 2 }, "25%");
  await vue.nextTick();

  await selectRange(viewer.controller, { row: 0, col: 1 }, { row: 1, col: 1 });
  assert.equal(stat(viewer.root, "xlsx-status-range"), "B1:B2");
  assertStats(viewer.root, {
    "xlsx-status-sum": "Sum: $3,234.56",
    "xlsx-status-average": "Average: $1,617.28",
    "xlsx-status-min": "Min: $1,234.56",
    "xlsx-status-max": "Max: $2,000.00",
  });

  await selectRange(viewer.controller, { row: 0, col: 2 }, { row: 1, col: 2 });
  assert.equal(stat(viewer.root, "xlsx-status-range"), "C1:C2");
  assertStats(viewer.root, {
    "xlsx-status-sum": "Sum: 75.00%",
    "xlsx-status-average": "Average: 37.50%",
    "xlsx-status-min": "Min: 25.00%",
    "xlsx-status-max": "Max: 50.00%",
  });
});

test("switching sheets resets the selection summary", async (t) => {
  const file = await workbookBytes(column([1, 2, 3]), [
    { name: "Second", cells: { A1: 9 } },
  ]);
  const viewer = await mountViewer(file);
  t.after(() => viewer.app.unmount());

  await selectRange(viewer.controller, { row: 0, col: 0 }, { row: 2, col: 0 });
  assert.equal(stat(viewer.root, "xlsx-status-sum"), "Sum: 6");

  viewer.controller.setActiveSheetIndex(1);
  await waitFor(
    () => viewer.controller.activeSheet?.name === "Second",
    "the second worksheet",
  );
  await vue.nextTick();

  assert.equal(stat(viewer.root, "xlsx-status-range"), undefined);
  assert.equal(stat(viewer.root, "xlsx-status-sum"), undefined);
  assert.equal(stat(viewer.root, "xlsx-status-ready"), "Ready");
});

test("filtered selections skip hidden rows and say Filtered", async (t) => {
  const file = withAutoFilter(
    await workbookBytes(column([10, 20, 30, 40, 50])),
    "A1:A5",
  );
  const viewer = await mountViewer(file);
  t.after(() => viewer.app.unmount());

  assert.equal(viewer.controller.autoFilterRanges.length, 1);
  assert.equal(viewer.controller.applyAutoFilterValues("A1:A5", 0, ["20"]), true);
  await vue.nextTick();
  assert.deepEqual(viewer.controller.hiddenAutoFilterRows, [2, 3, 4]);

  await selectRange(viewer.controller, { row: 0, col: 0 }, { row: 4, col: 0 });

  assert.equal(stat(viewer.root, "xlsx-status-range"), "A1:A5");
  assert.equal(stat(viewer.root, "xlsx-status-note"), "Filtered");
  assertStats(viewer.root, {
    "xlsx-status-count": "Count: 2",
    "xlsx-status-sum": "Sum: 30",
    "xlsx-status-average": "Average: 15",
    "xlsx-status-min": "Min: 10",
    "xlsx-status-max": "Max: 20",
  });
});

test("a selection beyond the cell budget is capped and marked", async (t) => {
  const file = await workbookBytes(column([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]));
  const viewer = await mountViewer(file);
  t.after(() => viewer.app.unmount());

  const originalGetWorksheet = viewer.controller.getActiveWorksheet;
  let cellsRead = 0;
  viewer.controller.getActiveWorksheet = function patchedGetActiveWorksheet() {
    const worksheet = originalGetWorksheet.call(viewer.controller);
    if (!worksheet) return worksheet;
    const getCalculatedValueAt = worksheet.getCalculatedValueAt.bind(worksheet);
    worksheet.getCalculatedValueAt = (...args) => {
      cellsRead += 1;
      return getCalculatedValueAt(...args);
    };
    return worksheet;
  };

  await selectRange(viewer.controller, { row: 0, col: 0 }, { row: 10_000, col: 0 });
  const marker = findByTestId(viewer.root, "xlsx-status-truncated");
  assert.ok(marker, "the truncated marker is rendered");
  assert.equal(marker.props.title, "Showing the first 10000 cells");
  assert.equal(cellsRead, 10_000, "the scan stops at the cell budget");
  assertStats(viewer.root, {
    "xlsx-status-count": "Count: 10",
    "xlsx-status-sum": "Sum: 55",
  });
});

test("showStatusBar:false removes the bar", async (t) => {
  const file = await workbookBytes({ A1: 1 });
  const viewer = await mountViewer(file, { viewerOptions: { showStatusBar: false } });
  t.after(() => viewer.app.unmount());

  await selectRange(viewer.controller, { row: 0, col: 0 }, { row: 0, col: 0 });
  assert.equal(findByTestId(viewer.root, "xlsx-status-bar"), undefined);
});

test("zh-CN locale translates the status labels", async (t) => {
  const file = await workbookBytes(column([1, 2]));
  const viewer = await mountViewer(file, {
    controllerOptions: { locale: "zh-CN" },
    viewerOptions: { locale: "zh-CN" },
  });
  t.after(() => viewer.app.unmount());

  await selectRange(viewer.controller, { row: 0, col: 0 }, { row: 1, col: 0 });
  assertStats(viewer.root, {
    "xlsx-status-count": "计数: 2",
    "xlsx-status-sum": "求和: 3",
    "xlsx-status-average": "平均值: 1.5",
    "xlsx-status-note": "已选定",
  });
});

// ── Worker-backed selections ───────────────────────────────────────────────

function workerSnapshot() {
  return {
    chartsByWorkbookSheetIndex: [[]],
    chartsheets: [],
    sheets: [{
      cachedFormulaValues: {},
      name: "Sheet1",
      visibility: "visible",
      workbookSheetIndex: 0,
      zoomScale: 100,
      maxUsedRow: 2,
      maxUsedCol: 0,
      minUsedRow: 0,
      minUsedCol: 0,
    }],
    tablesByWorkbookSheetIndex: [[]],
    tabs: [{
      id: "sheet-0",
      index: 0,
      kind: "sheet",
      name: "Sheet1",
      sheetIndex: 0,
      visibility: "visible",
      workbookSheetIndex: 0,
    }],
  };
}

class ManualWorker {
  constructor() {
    this.listeners = new Map();
    this.messages = [];
  }

  addEventListener(type, listener) {
    const listeners = this.listeners.get(type) ?? new Set();
    listeners.add(listener);
    this.listeners.set(type, listeners);
  }

  removeEventListener(type, listener) {
    this.listeners.get(type)?.delete(listener);
  }

  postMessage(message) {
    this.messages.push(message);
  }

  respond(message, result) {
    for (const listener of [...(this.listeners.get("message") ?? [])]) {
      listener({ data: { id: message.id, result, success: true } });
    }
  }

  terminate() {}
}

async function mountWorkerStatusBar(t) {
  const originalWorker = globalThis.Worker;
  globalThis.Worker = class Worker {};
  t.after(() => {
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
  });

  const file = await workbookBytes({ A1: 1, A2: 2, A3: 3 });
  const worker = new ManualWorker();
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        createWorker: () => worker,
        file,
        readOnly: true,
        skipXmlParsing: true,
        useWorker: true,
      });
      return () => vue.h(XlsxStatusBar, { controller });
    },
  });
  const mounted = await mount(Harness);
  t.after(() => mounted.app.unmount());

  await waitFor(() => worker.messages.length > 0, "worker load request");
  const load = worker.messages.find((message) => message.type === "load");
  assert.ok(load);
  worker.respond(load, workerSnapshot());
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "worker-backed controller",
  );
  assert.ifError(controller.error);
  await waitFor(() => controller.activeSheet, "worker sheet state");
  return { ...mounted, controller, worker };
}


test("Worker-backed selections request an aggregate and render it", async (t) => {
  const mounted = await mountWorkerStatusBar(t);
  mounted.controller.selectRange({
    start: { row: 0, col: 0 },
    end: { row: 2, col: 0 },
  });
  await vue.nextTick();

  await waitFor(
    () => mounted.worker.messages.some((message) => message.type === "getSelectionStats"),
    "worker selection stats request",
  );
  const request = mounted.worker.messages.find(
    (message) => message.type === "getSelectionStats",
  );
  assert.deepEqual(
    {
      cellCount: request.payload.cellCount,
      endCol: request.payload.endCol,
      endRow: request.payload.endRow,
      hiddenRows: request.payload.hiddenRows,
      maxCells: request.payload.maxCells,
      startCol: request.payload.startCol,
      startRow: request.payload.startRow,
      workbookSheetIndex: request.payload.workbookSheetIndex,
    },
    {
      cellCount: 3,
      endCol: 0,
      endRow: 2,
      hiddenRows: [],
      maxCells: 10_000,
      startCol: 0,
      startRow: 0,
      workbookSheetIndex: 0,
    },
  );

  mounted.worker.respond(request, {
    average: 2,
    cellCount: 3,
    count: 3,
    max: 3,
    min: 1,
    numberFormat: null,
    numericCount: 3,
    sum: 6,
    truncated: false,
  });

  await waitFor(
    () => stat(mounted.root, "xlsx-status-sum") === "Sum: 6",
    "worker-backed status totals",
  );
  assertStats(mounted.root, {
    "xlsx-status-range": "A1:A3",
    "xlsx-status-count": "Count: 3",
    "xlsx-status-average": "Average: 2",
    "xlsx-status-min": "Min: 1",
    "xlsx-status-max": "Max: 3",
    "xlsx-status-note": "Selected",
  });

  // A newer selection abandons the previous read instead of showing a stale
  // aggregate from the Worker.
  mounted.controller.selectCell({ row: 1, col: 0 });
  await vue.nextTick();
  await waitFor(
    () =>
      mounted.worker.messages.filter((message) => message.type === "getSelectionStats").length
        === 2,
    "second worker selection stats request",
  );
  const second = mounted.worker.messages.filter(
    (message) => message.type === "getSelectionStats",
  )[1];
  mounted.worker.respond(second, {
    average: 2,
    cellCount: 1,
    count: 1,
    max: 2,
    min: 2,
    numberFormat: null,
    numericCount: 1,
    sum: 2,
    truncated: false,
  });

  await waitFor(
    () => stat(mounted.root, "xlsx-status-sum") === "Sum: 2",
    "updated worker-backed totals",
  );
  assert.equal(stat(mounted.root, "xlsx-status-range"), "A2");
});
