import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

const { useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");
const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");

function configureMainThreadWasm() {
  const wasmBytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(
    wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
  );
}

function workbookBuffer() {
  const bytes = readFileSync(
    new URL("../../apps/demo/public/samples/sales-table.xlsx", import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
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

function enableWorkerSupport(t) {
  const originalWorker = globalThis.Worker;
  globalThis.Worker = class Worker {};
  t.after(() => {
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
  });
}

async function mountWorkerController(t, options = {}) {
  enableWorkerSupport(t);
  const worker = new ManualWorker();
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        createWorker: () => worker,
        file: workbookBuffer(),
        readOnly: true,
        skipXmlParsing: true,
        useWorker: true,
        ...options,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  t.after(() => mounted.app.unmount());

  await waitFor(() => worker.messages.length > 0, "worker load request");
  const loadMessage = worker.messages.find((message) => message.type === "load");
  assert.ok(loadMessage, "worker received a load request");
  worker.respond(loadMessage, workerSnapshotWithSheet());
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "worker-backed XLSX ready state",
  );
  assert.ifError(controller.error);
  assert.equal(controller.isWorkerBacked, true);

  return { ...mounted, controller, worker };
}

configureMainThreadWasm();

test("worker cell snapshots carry the number format so the formula bar localizes", async (t) => {
  const { controller, worker } = await mountWorkerController(t, { locale: "de-DE" });

  const cell = { col: 0, row: 0 };
  controller.selectCell(cell);
  await waitFor(
    () => worker.messages.some((message) => message.type === "getCellSnapshot"),
    "worker cell snapshot request",
  );

  const request = worker.messages.find((message) => message.type === "getCellSnapshot");
  assert.deepEqual(request.payload, { col: 0, row: 0, workbookSheetIndex: 0 });

  worker.respond(request, {
    displayValue: "1,234.56",
    formula: "",
    numberFormat: { formatType: "custom", formatString: "#,##0.00" },
  });

  await waitFor(
    () => controller.getCellDisplayValue(cell) === "1.234,56",
    "localized worker snapshot value",
  );
  assert.equal(controller.getCellDisplayValue(cell), "1.234,56");
  assert.equal(controller.getCellFormula(cell), "");
});

test("worker snapshots without a format keep the engine's display string", async (t) => {
  const { controller, worker } = await mountWorkerController(t, { locale: "de-DE" });

  const cell = { col: 0, row: 1 };
  controller.selectCell(cell);
  await waitFor(
    () => worker.messages.some((message) => message.type === "getCellSnapshot"),
    "worker cell snapshot request",
  );

  const request = worker.messages.find((message) => message.type === "getCellSnapshot");
  worker.respond(request, { displayValue: "1,234.56", formula: "" });

  await waitFor(
    () => controller.getCellDisplayValue(cell) === "1,234.56",
    "raw worker snapshot value",
  );
  // With no format information there is nothing to localize; the snapshot is
  // shown exactly as the worker produced it.
  assert.equal(controller.getCellDisplayValue(cell), "1,234.56");
});
