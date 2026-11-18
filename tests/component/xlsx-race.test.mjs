import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

const { useXlsxViewerController, useXlsxViewerThumbnails } = await importFromDemo("@arcships/vue-xlsx");
const { createXlsxRuntime, setWasmSource } = await importFromDemo("@arcships/xlsx-core");

const workbookBytes = readFileSync(
  new URL("../../apps/demo/public/samples/sales-table.xlsx", import.meta.url),
);

function workbookBuffer() {
  return workbookBytes.buffer.slice(
    workbookBytes.byteOffset,
    workbookBytes.byteOffset + workbookBytes.byteLength,
  );
}

function sampleBuffer(name) {
  const bytes = readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

let mainThreadWasmConfigured = false;
function configureMainThreadWasm() {
  if (mainThreadWasmConfigured) return;
  const wasmBytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(
    wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
  );
  mainThreadWasmConfigured = true;
}

function deferred() {
  let reject;
  let resolve;
  const promise = new Promise((nextResolve, nextReject) => {
    reject = nextReject;
    resolve = nextResolve;
  });
  return { promise, reject, resolve };
}

function emptyWorkerSnapshot() {
  return {
    chartsByWorkbookSheetIndex: [],
    chartsheets: [],
    sheets: [],
    tablesByWorkbookSheetIndex: [],
    tabs: [],
  };
}

function workerSnapshotWithSheet() {
  return {
    chartsByWorkbookSheetIndex: [[]],
    chartsheets: [],
    sheets: [{
      cachedFormulaValues: {},
      name: "Sheet1",
      visibility: "visible",
      workbookSheetIndex: 0,
      zoomScale: 100,
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
    this.terminateCount = 0;
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

  fail(message, error) {
    for (const listener of [...(this.listeners.get("message") ?? [])]) {
      listener({ data: { id: message.id, error, success: false } });
    }
  }

  terminate() {
    this.terminateCount += 1;
  }
}

function enableWorkerSupport(t) {
  const originalWorker = globalThis.Worker;
  globalThis.Worker = class Worker {};
  t.after(() => {
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
  });
}

function mockRuntime(id, createWorkerClient) {
  return {
    createWorkerClient,
    dispose() {},
    id,
    parseOptions: {},
    wasmSource: undefined,
    workerUrl: undefined,
  };
}

async function mountController(options) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController(options);
      return () => vue.h(
        "div",
        { "data-state": controller.sourceState },
        controller.sourceState,
      );
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, controller };
}

test("deferred XLSX load resumes once and keeps one task id", async (t) => {
  enableWorkerSupport(t);
  const worker = new ManualWorker();
  const diagnostics = [];
  const mounted = await mountController({
    createWorker: () => worker,
    deferLoadingAboveBytes: 1,
    file: workbookBuffer(),
    onDiagnostic: (event) => diagnostics.push(event),
    readOnly: true,
    skipXmlParsing: true,
  });
  t.after(() => mounted.app.unmount());

  await waitFor(
    () => mounted.controller.isLoadDeferred,
    "deferred XLSX state",
  );
  assert.equal(mounted.controller.canLoadDeferred, true);
  assert.equal(worker.messages.length, 0);
  const deferredDiagnostic = diagnostics.find(
    (event) => event.type === "load-deferred",
  );
  assert.ok(deferredDiagnostic?.taskId);

  mounted.controller.continueDeferredLoad();
  mounted.controller.continueDeferredLoad();
  assert.equal(worker.messages.length, 1);
  assert.equal(worker.messages[0].type, "load");
  assert.equal(
    diagnostics.filter((event) => event.type === "load-resumed").length,
    1,
  );

  worker.respond(worker.messages[0], emptyWorkerSnapshot());
  await waitFor(
    () => mounted.controller.sourceState === "ready",
    "resumed XLSX ready state",
  );
  assert.deepEqual(
    diagnostics.map((event) => event.type),
    ["load-start", "load-deferred", "load-resumed", "load-success"],
  );
  assert.deepEqual(
    new Set(diagnostics.map((event) => event.taskId)),
    new Set([deferredDiagnostic.taskId]),
  );
  assert.deepEqual(mounted.warnings, []);
});

test("unmount aborts a resumed deferred load and ignores its late result", async (t) => {
  enableWorkerSupport(t);
  const diagnostics = [];
  const pendingLoad = deferred();
  const unhandledRejections = [];
  let observedSignal;
  let terminateCount = 0;
  const onUnhandledRejection = (reason) => {
    unhandledRejections.push(reason);
  };
  process.on("unhandledRejection", onUnhandledRejection);
  t.after(() => process.off("unhandledRejection", onUnhandledRejection));

  const client = {
    dispose() {
      terminateCount += 1;
    },
    loadWorkbook(_buffer, _skipXmlParsing, _showHiddenSheets, signal) {
      observedSignal = signal;
      signal.addEventListener("abort", () => {
        pendingLoad.reject(new DOMException("Aborted", "AbortError"));
      }, { once: true });
      return pendingLoad.promise;
    },
  };
  const mounted = await mountController({
    deferLoadingAboveBytes: 1,
    file: workbookBuffer(),
    onDiagnostic: (event) => diagnostics.push(event),
    readOnly: true,
    runtime: mockRuntime("xlsx-race-unmount", () => client),
    skipXmlParsing: true,
  });

  await waitFor(
    () => mounted.controller.isLoadDeferred,
    "deferred XLSX state before unmount",
  );
  mounted.controller.continueDeferredLoad();
  assert.equal(observedSignal.aborted, false);
  assert.equal(mounted.controller.sourceState, "loading");

  mounted.app.unmount();
  await vue.nextTick();
  assert.equal(observedSignal.aborted, true);
  assert.equal(terminateCount, 1);
  assert.equal(mounted.controller.sourceState, "idle");

  pendingLoad.resolve(emptyWorkerSnapshot());
  await new Promise((resolve) => setTimeout(resolve, 0));
  await vue.nextTick();
  assert.equal(mounted.controller.sourceState, "idle");
  assert.equal(
    diagnostics.some((event) => event.type === "load-success"),
    false,
  );
  assert.deepEqual(unhandledRejections, []);
  assert.deepEqual(mounted.warnings, []);
});

test("rapid worker-backed cell selection aborts the old snapshot and commits only the latest", async (t) => {
  enableWorkerSupport(t);
  const snapshotCalls = [];
  const client = {
    dispose() {},
    async loadWorkbook() {
      return workerSnapshotWithSheet();
    },
    getCellSnapshot(workbookSheetIndex, row, col, signal) {
      const pending = deferred();
      snapshotCalls.push({
        col,
        pending,
        row,
        signal,
        workbookSheetIndex,
      });
      return pending.promise;
    },
  };
  const mounted = await mountController({
    deferLoadingAboveBytes: 0,
    file: workbookBuffer(),
    readOnly: true,
    runtime: mockRuntime("xlsx-race-cell", () => client),
    skipXmlParsing: true,
  });
  t.after(() => mounted.app.unmount());

  await waitFor(
    () => mounted.controller.sourceState === "ready",
    "worker-backed XLSX ready state",
  );
  assert.equal(mounted.controller.isWorkerBacked, true);

  mounted.controller.selectCell({ col: 0, row: 0 });
  await waitFor(() => snapshotCalls.length === 1, "first cell snapshot");
  mounted.controller.selectCell({ col: 1, row: 0 });
  await waitFor(() => snapshotCalls.length === 2, "second cell snapshot");

  assert.deepEqual(
    snapshotCalls.map(({ col, row, workbookSheetIndex }) => ({
      col,
      row,
      workbookSheetIndex,
    })),
    [
      { col: 0, row: 0, workbookSheetIndex: 0 },
      { col: 1, row: 0, workbookSheetIndex: 0 },
    ],
  );
  assert.equal(snapshotCalls[0].signal.aborted, true);
  assert.equal(snapshotCalls[1].signal.aborted, false);

  snapshotCalls[1].pending.resolve({
    displayValue: "latest value",
    formula: "=LATEST()",
  });
  await waitFor(
    () => mounted.controller.getCellDisplayValue({ col: 1, row: 0 }) === "latest value",
    "latest cell snapshot cache",
  );
  snapshotCalls[0].pending.resolve({
    displayValue: "stale value",
    formula: "=STALE()",
  });
  await vue.nextTick();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(
    mounted.controller.getCellDisplayValue({ col: 0, row: 0 }),
    "",
  );
  assert.equal(
    mounted.controller.getCellDisplayValue({ col: 1, row: 0 }),
    "latest value",
  );
  assert.equal(
    mounted.controller.getCellFormula({ col: 1, row: 0 }),
    "=LATEST()",
  );
  assert.deepEqual(mounted.warnings, []);
});

test("public XLSX download always cleans up and avoids URL creation without document", async (t) => {
  enableWorkerSupport(t);
  const originalDocument = globalThis.document;
  const originalCreateObjectURL = URL.createObjectURL;
  const originalRevokeObjectURL = URL.revokeObjectURL;
  t.after(() => {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    URL.createObjectURL = originalCreateObjectURL;
    URL.revokeObjectURL = originalRevokeObjectURL;
  });
  const client = {
    dispose() {},
    async loadWorkbook() {
      return emptyWorkerSnapshot();
    },
  };
  const mounted = await mountController({
    deferLoadingAboveBytes: 0,
    file: workbookBuffer(),
    fileName: "report.xlsx",
    readOnly: true,
    runtime: mockRuntime("xlsx-race-download", () => client),
    skipXmlParsing: true,
  });
  t.after(() => mounted.app.unmount());
  await waitFor(
    () => mounted.controller.sourceState === "ready",
    "downloadable XLSX ready state",
  );

  let createCount = 0;
  let removeCount = 0;
  let revokeCount = 0;
  URL.createObjectURL = () => {
    createCount += 1;
    return "blob:xlsx-race";
  };
  URL.revokeObjectURL = (url) => {
    assert.equal(url, "blob:xlsx-race");
    revokeCount += 1;
  };
  globalThis.document = {
    body: { append() {} },
    createElement: () => ({
      click() {
        throw new Error("click failed");
      },
      remove() {
        removeCount += 1;
      },
    }),
  };

  assert.throws(
    () => mounted.controller.download(),
    /click failed/,
  );
  assert.equal(createCount, 1);
  assert.equal(removeCount, 1);
  assert.equal(revokeCount, 1);

  delete globalThis.document;
  assert.throws(
    () => mounted.controller.download(),
    /document/,
  );
  assert.equal(createCount, 1, "missing document must not create an object URL");
  assert.equal(removeCount, 1);
  assert.equal(revokeCount, 1);
});

test("unmount invalidates an asynchronous XLSX snapshot history restore", async (t) => {
  configureMainThreadWasm();
  const unhandledRejections = [];
  const onUnhandledRejection = (reason) => unhandledRejections.push(reason);
  process.on("unhandledRejection", onUnhandledRejection);
  t.after(() => process.off("unhandledRejection", onUnhandledRejection));

  const mounted = await mountController({
    file: workbookBuffer(),
    fileName: "history-race.xlsx",
    runtime: mockRuntime("xlsx-race-history", () => {
      throw new Error("editable history test must not create a Worker");
    }),
    skipXmlParsing: true,
    useWorker: false,
  });
  await waitFor(
    () => mounted.controller.sourceState === "ready" || mounted.controller.error,
    "editable XLSX terminal state",
  );
  assert.ifError(mounted.controller.error);
  assert.ok(mounted.controller.workbook);

  mounted.controller.selectRange({
    start: { col: 0, row: 0 },
    end: { col: 1, row: 0 },
  });
  mounted.controller.mergeSelection();
  assert.equal(mounted.controller.canUndo, true);

  mounted.controller.undo();
  mounted.app.unmount();
  await vue.nextTick();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(mounted.controller.workbook, null);
  assert.equal(mounted.controller.sourceState, "idle");
  assert.deepEqual(unhandledRejections, []);
  assert.deepEqual(mounted.warnings, []);
});

test("worker load failure falls back to the main thread", async (t) => {
  configureMainThreadWasm();
  enableWorkerSupport(t);
  const worker = new ManualWorker();
  const diagnostics = [];
  const mounted = await mountController({
    createWorker: () => worker,
    deferLoadingAboveBytes: 0,
    file: workbookBuffer(),
    onDiagnostic: (event) => diagnostics.push(event),
    readOnly: true,
    skipXmlParsing: true,
  });
  t.after(() => mounted.app.unmount());

  await waitFor(() => worker.messages.length === 1, "worker load request");
  assert.equal(worker.messages[0].type, "load");
  worker.fail(worker.messages[0], {
    code: "WORKER_FAILED",
    message: "worker exploded",
  });

  await waitFor(
    () => mounted.controller.sourceState === "ready"
      || mounted.controller.sourceState === "error",
    "main-thread fallback terminal state",
  );
  assert.equal(mounted.controller.sourceState, "ready");
  assert.equal(mounted.controller.sourceError, null);
  assert.equal(mounted.controller.isWorkerBacked, false);
  assert.ok(mounted.controller.workbook, "fallback must load a real workbook");
  assert.equal(
    diagnostics.filter((event) => event.type === "worker-error").length,
    1,
  );
  assert.equal(
    diagnostics.some((event) => event.type === "load-error"),
    false,
  );
  assert.deepEqual(mounted.warnings, []);
});

test("late worker chart hydration does not overwrite charts refreshed after an edit", async (t) => {
  configureMainThreadWasm();
  enableWorkerSupport(t);
  const worker = new ManualWorker();
  const mounted = await mountController({
    createWorker: () => worker,
    deferLoadingAboveBytes: 0,
    file: sampleBuffer("charts-images.xlsx"),
    readOnly: false,
  });
  t.after(() => mounted.app.unmount());

  await waitFor(
    () => mounted.controller.sourceState === "ready",
    "chart workbook ready state",
  );
  await waitFor(
    () => worker.messages.some((message) => message.type === "parseCharts"),
    "chart hydration request",
  );
  const chartMessage = worker.messages.find((message) => message.type === "parseCharts");
  const tabCountBefore = mounted.controller.tabs.length;
  const chartsheetCountBefore = mounted.controller.chartsheets.length;
  assert.ok(tabCountBefore > 0);
  assert.ok(chartsheetCountBefore > 0);

  mounted.controller.setCellValue({ col: 0, row: 0 }, "edited");
  assert.equal(mounted.controller.tabs.length, tabCountBefore);
  assert.equal(mounted.controller.chartsheets.length, chartsheetCountBefore);

  let quickChartCount = 0;
  for (let sheetIndex = 0; sheetIndex < mounted.controller.sheets.length; sheetIndex += 1) {
    quickChartCount += mounted.controller.getSheetCharts(sheetIndex).length;
  }
  worker.respond(chartMessage, {
    chartsByWorkbookSheetIndex: [
      Array.from({ length: Math.max(1, quickChartCount) }, () => ({
        chartType: "Line",
        series: [],
      })),
    ],
    chartsheets: [],
    tabs: [],
  });
  await new Promise((resolve) => setTimeout(resolve, 0));
  await vue.nextTick();

  assert.equal(
    mounted.controller.chartsheets.length,
    chartsheetCountBefore,
    "a stale worker result replaced charts refreshed after the edit",
  );
  assert.equal(mounted.controller.tabs.length, tabCountBefore);
  assert.deepEqual(mounted.warnings, []);
});

test("worker-backed charts refresh from the current Worker model after an edit", async (t) => {
  enableWorkerSupport(t);
  const worker = new ManualWorker();
  const mounted = await mountController({
    createWorker: () => worker,
    file: sampleBuffer("charts-images.xlsx"),
    readOnly: true,
  });
  t.after(() => mounted.app.unmount());

  await waitFor(
    () => worker.messages.some((message) => message.type === "load"),
    "worker load request",
  );
  const loadMessage = worker.messages.find((message) => message.type === "load");
  const firstChart = {
    axes: [],
    chartType: "Line",
    id: "chart-0-0",
    series: [{
      categories: ["A", "B"],
      dataPoints: [],
      id: "series-1",
      name: "Series 1",
      values: [10, 20],
    }],
    sheetIndex: 0,
    workbookSheetIndex: 0,
    zIndex: 0,
  };
  const snapshot = workerSnapshotWithSheet();
  worker.respond(loadMessage, {
    ...snapshot,
    chartsByWorkbookSheetIndex: [[firstChart]],
  });
  await waitFor(
    () => mounted.controller.sourceState === "ready" && mounted.controller.isWorkerBacked,
    "worker-backed chart workbook",
  );
  assert.deepEqual(mounted.controller.getSheetCharts(0)[0].series[0].values, [10, 20]);

  const refreshPromise = mounted.controller.refreshChartData();
  await waitFor(
    () => worker.messages.some((message) => message.type === "refreshCharts"),
    "worker chart refresh request",
  );
  const refreshMessage = worker.messages.find((message) => message.type === "refreshCharts");
  worker.respond(refreshMessage, {
    chartsByWorkbookSheetIndex: [[{
      ...firstChart,
      series: firstChart.series.map((series) => ({ ...series, values: [30, 40] })),
    }]],
    chartsheets: [],
    tabs: snapshot.tabs,
  });

  assert.equal(await refreshPromise, true);
  assert.deepEqual(
    mounted.controller.getSheetCharts(0)[0].series[0].values,
    [30, 40],
    "worker-backed charts keep the source cache after an in-worker edit",
  );
  assert.deepEqual(mounted.warnings, []);
});

test("worker-mode sheet thumbnails render snapshot values and widths", () => {
  const sheet = {
    cachedFormulaValues: { A1: "Revenue", B1: "1200" },
    colWidthOverridesPx: {},
    colWidths: [120, 90],
    defaultColWidthPx: 64,
    defaultRowHeightPx: 24,
    maxUsedCol: 1,
    maxUsedRow: 1,
    minUsedCol: 0,
    minUsedRow: 0,
    name: "Data",
    rowHeightOverridesPx: {},
    rowHeights: [30, 24],
    visibleCols: [0, 1],
    visibleRows: [0, 1],
    workbookSheetIndex: 0,
  };
  const thumbnails = useXlsxViewerThumbnails(() => ({
    sheets: [sheet],
    workbook: null,
  }));
  const thumbnail = thumbnails.thumbnails[0];
  assert.ok(thumbnail);
  assert.equal(thumbnail.contentWidth, 120 + 90 + 64 + 64);
  assert.equal(thumbnail.contentHeight, 30 + 24 + 24 + 24);

  const paintedText = [];
  const context = {
    beginPath() {},
    fillRect() {},
    fillText(text) { paintedText.push(text); },
    lineTo() {},
    moveTo() {},
    setTransform() {},
    stroke() {},
  };
  const canvas = {
    height: 0,
    style: {},
    width: 0,
    getContext: () => context,
  };
  assert.equal(thumbnails.paintThumbnail(0, canvas), true);
  assert.ok(paintedText.includes("Revenue"));
  assert.ok(paintedText.includes("1200"));
});

test("main-thread parsing enforces the maxParseMs budget", async (t) => {
  configureMainThreadWasm();
  const runtime = createXlsxRuntime({ limits: { maxParseMs: 1 } });
  t.after(() => runtime.dispose());
  const mounted = await mountController({
    file: sampleBuffer("large-grid.xlsx"),
    readOnly: false,
    runtime,
    skipXmlParsing: true,
    useWorker: false,
  });
  t.after(() => mounted.app.unmount());

  await waitFor(
    () => mounted.controller.sourceState === "ready"
      || mounted.controller.sourceState === "error",
    "main-thread parse terminal state",
  );
  assert.equal(mounted.controller.sourceState, "error");
  assert.equal(mounted.controller.sourceError?.code, "TIMEOUT");
  assert.equal(mounted.controller.sourceError?.limit, "maxParseMs");
  assert.deepEqual(mounted.warnings, []);
});

test("large main-thread workbook load keeps peak source-buffer memory under 1.5x", async (t) => {
  configureMainThreadWasm();
  const file = sampleBuffer("large-grid.xlsx");
  const fileSize = file.byteLength;
  assert.ok(fileSize > 64 * 1024, "fixture must be large enough for a meaningful budget");

  // Explicit byteLength accounting instead of a noisy process-wide heap delta:
  // every large ArrayBuffer clone made while loading counts against the peak.
  const originalSlice = ArrayBuffer.prototype.slice;
  let copiedBytes = 0;
  ArrayBuffer.prototype.slice = function trackedSlice(...args) {
    const result = originalSlice.apply(this, args);
    if (result instanceof ArrayBuffer && result.byteLength >= 1024) {
      copiedBytes += result.byteLength;
    }
    return result;
  };
  t.after(() => {
    ArrayBuffer.prototype.slice = originalSlice;
  });

  const mounted = await mountController({
    file,
    readOnly: false,
    skipXmlParsing: true,
    useWorker: false,
  });
  t.after(() => mounted.app.unmount());

  await waitFor(
    () => mounted.controller.sourceState === "ready"
      || mounted.controller.sourceState === "error",
    "large workbook load terminal state",
  );
  assert.ifError(mounted.controller.error);
  assert.equal(mounted.controller.sourceState, "ready");

  // The source itself is 1x; the whole load may keep at most another ~0.5x.
  const peakSourceBytes = fileSize + copiedBytes;
  const budget = Math.ceil(fileSize * 1.5);
  assert.ok(
    peakSourceBytes <= budget,
    `peak source-buffer memory ${peakSourceBytes} exceeded ${budget} `
      + `(fileSize=${fileSize}, copiedBytes=${copiedBytes})`,
  );
  assert.deepEqual(mounted.warnings, []);
});
