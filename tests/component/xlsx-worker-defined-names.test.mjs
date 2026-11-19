import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

import { importFromDemo, mount, vue, waitFor } from "./vue-test-renderer.mjs";

// Issue #22 audit: Worker-backed read-only loads replace the main-thread
// workbook with `null`, so `getNamedRanges()` cannot answer. The controller
// must fall back to the Worker's sheet snapshot for `definedNames`, otherwise
// the name box is empty even though the workbook has names.

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

/** `Worker`-shaped adapter that runs the real `dist/xlsx-worker.js` in-process. */
class WorkerScope {
  constructor() {
    this.listeners = new Map();
    this.owner = null;
  }

  addEventListener(type, listener) {
    this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
  }

  removeEventListener(type, listener) {
    this.listeners.set(
      type,
      (this.listeners.get(type) ?? []).filter((entry) => entry !== listener),
    );
  }

  postMessage(message) {
    this.owner?.receive(message);
  }

  dispatch(message) {
    for (const listener of [...(this.listeners.get("message") ?? [])]) {
      listener({ data: message });
    }
  }
}

class FakeWorker {
  constructor(scope) {
    this.scope = scope;
    this.listeners = new Map();
    scope.owner = this;
  }

  addEventListener(type, listener) {
    this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
  }

  removeEventListener(type, listener) {
    this.listeners.set(
      type,
      (this.listeners.get(type) ?? []).filter((entry) => entry !== listener),
    );
  }

  postMessage(message) {
    queueMicrotask(() => this.scope.dispatch(message));
  }

  receive(message) {
    for (const listener of [...(this.listeners.get("message") ?? [])]) {
      listener({ data: message });
    }
  }

  terminate() {}
}

const workerScope = new WorkerScope();
globalThis.self = workerScope;
globalThis.Worker ??= class Worker {};
await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core/worker")).href
);

const { createXlsxRuntime } = await importFromDemo("@arcships/xlsx-core");
const { useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

const wasmBytes = readFileSync(
  new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
);
const wasmSource = wasmBytes.buffer.slice(
  wasmBytes.byteOffset,
  wasmBytes.byteOffset + wasmBytes.byteLength,
);

function workbookBuffer() {
  const bytes = readFileSync(
    new URL("../../test-data/corpus/xlsx-pivot-slicers.xlsx", import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function mountWorkerBackedController(t) {
  const worker = new FakeWorker(workerScope);
  const runtime = createXlsxRuntime({
    createWorker: () => worker,
    wasmSource,
  });
  let controller;
  const mounted = await mount(vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        fileName: "xlsx-pivot-slicers.xlsx",
        file: workbookBuffer(),
        readOnly: true,
        runtime,
        useWorker: true,
      });
      return () => vue.h("div");
    },
  }));
  t.after(() => {
    mounted.app.unmount();
    runtime.dispose();
  });
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "Worker-backed workbook loaded",
  );
  assert.ifError(controller.error);
  return { controller };
}

test("Worker-backed loads expose defined names from the sheet snapshot", async (t) => {
  const { controller } = await mountWorkerBackedController(t);

  assert.equal(controller.isWorkerBacked, true);
  assert.equal(
    controller.workbook,
    null,
    "a Worker-backed load keeps no main-thread workbook",
  );
  assert.deepEqual(
    controller.activeSheet.definedNames.map((entry) => entry.name),
    ["SalesData", "RevenueTotal", "CurrentProduct"],
  );
  assert.deepEqual(
    controller.definedNames.map((entry) => entry.name),
    ["SalesData", "RevenueTotal", "CurrentProduct"],
  );
  assert.deepEqual(
    controller.definedNames.map((entry) => entry.scope),
    ["workbook", "workbook", "sheet"],
  );
  assert.equal(
    controller.definedNames[2].sheetName,
    "Data",
    "the snapshot keeps the sheet scope of a local name",
  );
});
