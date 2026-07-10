import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

const { useXlsxViewerController } = await importFromDemo("@extend-ai/vue-xlsx");
const { setWasmSource } = await importFromDemo("@extend-ai/xlsx-core");

const workbookBytes = readFileSync(
  new URL("../../apps/demo/public/samples/sales-table.xlsx", import.meta.url),
);

function workbookBuffer() {
  return workbookBytes.buffer.slice(
    workbookBytes.byteOffset,
    workbookBytes.byteOffset + workbookBytes.byteLength,
  );
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
  const wasmBytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(
    wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
  );
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
