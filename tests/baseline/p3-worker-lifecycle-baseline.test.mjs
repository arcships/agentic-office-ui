import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  waitFor,
} from "../component/vue-test-renderer.mjs";

const {
  createBlankDocumentModel,
  createDocxRuntime,
} = await importFromDemo("@arcships/docx-core");
const {
  createXlsxRuntime,
  setWasmSource,
} = await importFromDemo("@arcships/xlsx-core");
const { useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

const TEST_PARSE_TIMEOUT_MS = 10;
const OBSERVATION_WINDOW_MS = 40;

class FakeWorker {
  constructor(onPost) {
    this.listeners = new Map();
    this.messages = [];
    this.onPost = onPost;
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
    this.onPost?.(message, this);
  }

  emit(type, event) {
    queueMicrotask(() => {
      for (const listener of [...(this.listeners.get(type) ?? [])]) {
        listener(event);
      }
    });
  }

  respond(data) {
    this.emit("message", { data });
  }

  terminate() {
    this.terminateCount += 1;
  }
}

function installWorkerGlobal(t) {
  const originalWorker = globalThis.Worker;
  globalThis.Worker = class Worker {};
  t.after(() => {
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
  });
}

function observePromise(promise, timeoutMs = OBSERVATION_WINDOW_MS) {
  return new Promise((resolve) => {
    let settled = false;
    const finish = (outcome) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve(outcome);
    };
    const timer = setTimeout(() => finish({ status: "pending" }), timeoutMs);
    promise.then(
      (value) => finish({ status: "resolved", value }),
      (error) => finish({ status: "rejected", error }),
    );
  });
}

function assertTimeoutOutcome(outcome, label) {
  assert.equal(outcome.status, "rejected", `${label} remained pending past its public timeout`);
  assert.equal(outcome.error?.code, "TIMEOUT", `${label} did not return the stable TIMEOUT code`);
}

function docxSuccess(message) {
  return {
    id: message.id,
    type: "success",
    model: createBlankDocumentModel(),
    package: {},
    timings: { buildModelMs: 1, parseMs: 1, totalMs: 2 },
  };
}

function xlsxSuccess(message) {
  return {
    id: message.id,
    success: true,
    result: {
      chartsByWorkbookSheetIndex: [],
      chartsheets: [],
      sheets: [],
      tablesByWorkbookSheetIndex: [],
      tabs: [],
    },
  };
}

function sampleWorkbookBuffer() {
  const bytes = readFileSync(
    new URL("../../apps/demo/public/samples/sales-table.xlsx", import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function configureXlsxMainThreadWasm() {
  const bytes = readFileSync(
    new URL(
      "../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm",
      import.meta.url,
    ),
  );
  setWasmSource(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  );
}

test("P3-WORKER-01 regression: DOCX timeout terminates Worker and the next load recovers", async (t) => {
  installWorkerGlobal(t);
  const hangingWorker = new FakeWorker();
  const succeedingWorker = new FakeWorker((message, worker) => {
    worker.respond(docxSuccess(message));
  });
  const workers = [hangingWorker, succeedingWorker];
  const runtime = createDocxRuntime({
    createWorker: () => workers.shift(),
    limits: { maxParseMs: TEST_PARSE_TIMEOUT_MS },
    wasmSource: new ArrayBuffer(8),
  });
  const loader = runtime.createLoader();

  const firstPromise = loader.load(new ArrayBuffer(16));
  const firstOutcome = await observePromise(firstPromise);
  const terminatedAtTimeout = hangingWorker.terminateCount;
  if (firstOutcome.status === "pending") loader.cancel();
  await firstPromise.catch(() => undefined);

  const recovered = await loader.load(new ArrayBuffer(16));
  loader.dispose();
  runtime.dispose();

  assert.equal(recovered.source, "worker", "the next load did not recover with a fresh Worker");
  assertTimeoutOutcome(firstOutcome, "DOCX Worker request");
  assert.equal(terminatedAtTimeout, 1, "DOCX timed-out Worker was not terminated immediately");
});

test("P3-WORKER-01 regression: DOCX cancellation terminates its Worker", async (t) => {
  installWorkerGlobal(t);
  const worker = new FakeWorker();
  const runtime = createDocxRuntime({
    createWorker: () => worker,
    wasmSource: new ArrayBuffer(8),
  });
  const loader = runtime.createLoader();
  const pending = loader.load(new ArrayBuffer(16));
  loader.cancel();

  await assert.rejects(
    pending,
    (error) => error?.code === "ABORTED" && error?.name === "AbortError",
  );
  const terminatedAtCancellation = worker.terminateCount;
  loader.dispose();
  runtime.dispose();
  assert.equal(terminatedAtCancellation, 1);
});

test("P3-WORKER-01 regression: DOCX Worker parse errors never use main-thread fallback", async (t) => {
  installWorkerGlobal(t);
  const diagnostics = [];
  const worker = new FakeWorker((message, currentWorker) => {
    currentWorker.respond({
      id: message.id,
      type: "error",
      error: { code: "PARSE_FAILED", message: "worker parse failed" },
    });
  });
  const runtime = createDocxRuntime({
    allowMainThreadFallback: true,
    createWorker: () => worker,
    mainThreadFallbackMaxBytes: 1024,
    onDiagnostic: (event) => diagnostics.push(event),
    wasmSource: new ArrayBuffer(8),
  });

  const rejection = assert.rejects(
    runtime.importDocxBuffer(new ArrayBuffer(16)),
    (error) => error?.code === "PARSE_FAILED",
  );
  await rejection;
  const usedFallback = diagnostics.some((event) => event.type === "main-thread-fallback");
  const terminateCount = worker.terminateCount;
  runtime.dispose();
  assert.equal(
    usedFallback,
    false,
  );
  assert.equal(terminateCount, 1);
});

test("P3-WORKER-01 regression: XLSX timeout terminates Worker and a new client recovers", async (t) => {
  installWorkerGlobal(t);
  const hangingWorker = new FakeWorker();
  const succeedingWorker = new FakeWorker((message, worker) => {
    worker.respond(xlsxSuccess(message));
  });
  const workers = [hangingWorker, succeedingWorker];
  const runtime = createXlsxRuntime({
    createWorker: () => workers.shift(),
    limits: { maxParseMs: TEST_PARSE_TIMEOUT_MS },
    wasmSource: "https://assets.example.test/xlsx.wasm",
  });
  const firstClient = runtime.createWorkerClient();
  const firstPromise = firstClient.loadWorkbook(new ArrayBuffer(16));
  const firstOutcome = await observePromise(firstPromise);
  const terminatedAtTimeout = hangingWorker.terminateCount;
  if (firstOutcome.status === "pending") firstClient.dispose();
  await firstPromise.catch(() => undefined);

  const secondClient = runtime.createWorkerClient();
  const recovered = await secondClient.loadWorkbook(new ArrayBuffer(16));
  secondClient.dispose();
  runtime.dispose();

  assert.deepEqual(recovered.sheets, [], "a new XLSX client did not recover after timeout");
  assertTimeoutOutcome(firstOutcome, "XLSX Worker request");
  assert.equal(terminatedAtTimeout, 1, "XLSX timed-out Worker was not terminated immediately");
});

test("P3-WORKER-01 regression: XLSX cancellation terminates its Worker", async (t) => {
  installWorkerGlobal(t);
  const worker = new FakeWorker();
  const runtime = createXlsxRuntime({
    createWorker: () => worker,
    wasmSource: "https://assets.example.test/xlsx.wasm",
  });
  const client = runtime.createWorkerClient();
  const controller = new AbortController();
  const pending = client.loadWorkbook(new ArrayBuffer(16), false, false, controller.signal);
  controller.abort();

  await assert.rejects(pending, (error) => error?.name === "AbortError");
  const terminatedAtCancellation = worker.terminateCount;
  runtime.dispose();
  assert.equal(terminatedAtCancellation, 1);
});

test("P3-WORKER-01 regression: XLSX Worker errors do not silently rerun on main thread", async (t) => {
  installWorkerGlobal(t);
  configureXlsxMainThreadWasm();
  const diagnostics = [];
  const worker = new FakeWorker((message, currentWorker) => {
    currentWorker.respond({
      id: message.id,
      success: false,
      error: "DOMParser is not defined",
    });
  });
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        createWorker: () => worker,
        file: sampleWorkbookBuffer(),
        onDiagnostic: (event) => diagnostics.push(event),
        readOnly: true,
        skipXmlParsing: true,
      });
      return () => vue.h("div", { "data-state": controller.sourceState });
    },
  });
  const mounted = await mount(Harness);

  try {
    await waitFor(
      () => controller.sourceState === "ready" || controller.sourceState === "error",
      "XLSX Worker failure terminal state",
    );
    assert.equal(
      controller.sourceState,
      "error",
      `XLSX silently reported ${controller.sourceState} after its Worker failed`,
    );
    assert.equal(
      diagnostics.some((event) => event.type === "load-success"),
      false,
      "XLSX emitted load-success after silently rerunning on the main thread",
    );
  } finally {
    mounted.app.unmount();
  }
});
