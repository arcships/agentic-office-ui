import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const samples = new URL("../../apps/demo/public/samples/", import.meta.url);
const DEFAULT_RESPONSE_TIMEOUT_MS = 20_000;

function sampleBuffer(name) {
  const bytes = readFileSync(new URL(name, samples));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

/**
 * Runs the real `dist/xlsx-worker.js` bundle inside this process. The Worker
 * protocol is modelled as synchronous message delivery plus asynchronous
 * responses, matching the browser event loop closely enough to exercise the
 * worker-side serial queue.
 */
class FakeWorkerGlobalScope {
  constructor() {
    this.listeners = new Map();
    this.nextRequestId = 1;
    this.queuedResponses = new Map();
    this.responseOrder = [];
    this.waiters = new Map();
  }

  addEventListener(type, listener) {
    this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
  }

  removeEventListener(type, listener) {
    this.listeners.set(
      type,
      (this.listeners.get(type) ?? []).filter((item) => item !== listener),
    );
  }

  postMessage(message) {
    if (!message || typeof message.id !== "number") return;
    this.responseOrder.push(message.id);
    const waiter = this.waiters.get(message.id);
    if (waiter) {
      this.waiters.delete(message.id);
      waiter.resolve(message);
      return;
    }
    this.queuedResponses.set(message.id, message);
  }

  dispatch(message) {
    for (const listener of [...(this.listeners.get("message") ?? [])]) {
      listener({ data: message });
    }
  }

  request(type, payload) {
    const id = this.nextRequestId;
    this.nextRequestId += 1;
    const response = this.waitForResponse(id);
    this.dispatch({ id, payload, type });
    return { id, response };
  }

  /** Dispatches a request whose response is deliberately not awaited. */
  send(type, payload) {
    const id = this.nextRequestId;
    this.nextRequestId += 1;
    this.dispatch({ id, payload, type });
    return id;
  }

  cancel(targetId) {
    this.dispatch({ payload: { targetId }, type: "cancel" });
  }

  waitForResponse(id, timeoutMs = DEFAULT_RESPONSE_TIMEOUT_MS) {
    const queued = this.queuedResponses.get(id);
    if (queued) {
      this.queuedResponses.delete(id);
      return Promise.resolve(queued);
    }
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.waiters.delete(id);
        reject(new Error(`worker request ${id} did not respond within ${timeoutMs}ms`));
      }, timeoutMs);
      this.waiters.set(id, {
        resolve: (response) => {
          clearTimeout(timer);
          resolve(response);
        },
      });
    });
  }

  hasResponse(id) {
    return this.queuedResponses.has(id);
  }
}

const workerScope = new FakeWorkerGlobalScope();
globalThis.self = workerScope;

const workerModuleUrl = pathToFileURL(
  requireFromXlsxCore.resolve("@arcships/xlsx-core/worker"),
).href;
await import(workerModuleUrl);
const core = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);
const sheetsWasm = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@dukelib/sheets-wasm")).href
);

const wasmBytes = readFileSync(
  requireFromXlsxCore.resolve("@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm"),
);
const wasmBuffer = wasmBytes.buffer.slice(
  wasmBytes.byteOffset,
  wasmBytes.byteOffset + wasmBytes.byteLength,
);
const limits = core.DEFAULT_XLSX_RUNTIME_LIMITS;

let createdWorkbooks = [];
let freedWorkbooks = new WeakSet();
const originalFromBytes = sheetsWasm.Workbook.fromBytes;
const originalFree = sheetsWasm.Workbook.prototype.free;
sheetsWasm.Workbook.fromBytes = function trackedFromBytes(...args) {
  const workbook = originalFromBytes.apply(this, args);
  createdWorkbooks.push(workbook);
  return workbook;
};
sheetsWasm.Workbook.prototype.free = function trackedFree(...args) {
  freedWorkbooks.add(this);
  return originalFree.apply(this, args);
};

function resetWorkbookTracking() {
  createdWorkbooks = [];
  freedWorkbooks = new WeakSet();
}

let wasmSourceSent = false;

function loadRequest(name) {
  const payload = { buffer: sampleBuffer(name), limits };
  if (!wasmSourceSent) {
    payload.wasmSource = wasmBuffer;
    wasmSourceSent = true;
  }
  return workerScope.request("load", payload);
}

async function loadWorkbook(name) {
  const response = await loadRequest(name).response;
  assert.equal(response.success, true, `load ${name} failed: ${JSON.stringify(response)}`);
  return response.result;
}

async function readCell(workbookSheetIndex = 0, row = 0, col = 0) {
  const response = await workerScope.request("getCellSnapshot", {
    col,
    row,
    workbookSheetIndex,
  }).response;
  assert.equal(response.success, true, `cell read failed: ${JSON.stringify(response)}`);
  return response.result;
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

test("XLSX worker: the latest load wins and a superseded load never commits state", async () => {
  resetWorkbookTracking();
  const first = loadRequest("sales-table.xlsx");
  const second = loadRequest("financial-model.xlsx");
  const [firstResponse, secondResponse] = await Promise.all([
    first.response,
    second.response,
  ]);

  assert.equal(firstResponse.success, false);
  assert.equal(firstResponse.error.code, "SUPERSEDED");
  assert.equal(secondResponse.success, true);
  assert.deepEqual(
    secondResponse.result.sheets.map((sheet) => sheet.name),
    ["Assumptions", "P&L", "Notes"],
  );
  assert.equal(
    createdWorkbooks.length,
    1,
    "the superseded load created a workbook it then had to release",
  );
  assert.equal(freedWorkbooks.has(createdWorkbooks[0]), false);

  const cell = await readCell();
  assert.equal(cell.displayValue, "Driver");
});

test("XLSX worker: replacing a workbook frees the previous one, never the current one", async () => {
  resetWorkbookTracking();
  await loadWorkbook("sales-table.xlsx");
  const salesWorkbook = createdWorkbooks.at(-1);
  assert.equal(freedWorkbooks.has(salesWorkbook), false);

  await loadWorkbook("financial-model.xlsx");
  const modelWorkbook = createdWorkbooks.at(-1);
  assert.notEqual(modelWorkbook, salesWorkbook);
  assert.equal(
    freedWorkbooks.has(salesWorkbook),
    true,
    "the replaced workbook was never freed",
  );
  assert.equal(
    freedWorkbooks.has(modelWorkbook),
    false,
    "the current workbook was freed by a later load",
  );
  assert.equal((await readCell()).displayValue, "Driver");

  await loadWorkbook("sales-table.xlsx");
  const reloadedWorkbook = createdWorkbooks.at(-1);
  assert.equal(
    freedWorkbooks.has(modelWorkbook),
    true,
    "the replaced workbook was never freed",
  );
  assert.equal(
    freedWorkbooks.has(reloadedWorkbook),
    false,
    "the current workbook was freed by a later load",
  );
  assert.equal((await readCell()).displayValue, "Region");
  assert.deepEqual(
    (await loadWorkbook("financial-model.xlsx")).sheets.map((sheet) => sheet.name),
    ["Assumptions", "P&L", "Notes"],
  );
});

test("XLSX worker: interactive reads wait for the running load instead of failing", async () => {
  await loadWorkbook("sales-table.xlsx");
  const load = loadRequest("financial-model.xlsx");
  const rows = workerScope.request("getRowsBatch", {
    rowCount: 5,
    startRow: 0,
    workbookSheetIndex: 0,
  });
  const [loadResponse, rowsResponse] = await Promise.all([load.response, rows.response]);

  assert.equal(loadResponse.success, true);
  assert.equal(rowsResponse.success, true);
  assert.ok(Array.isArray(rowsResponse.result), "rows request did not run after the load");
  assert.ok(rowsResponse.result.length > 0, "rows request returned no data after the load");
  assert.ok(
    workerScope.responseOrder.indexOf(load.id) < workerScope.responseOrder.indexOf(rows.id),
    "the read responded before the load it was queued behind",
  );
  assert.equal((await readCell()).displayValue, "Driver");
});

test("XLSX worker: a cancelled queued read is dropped and the worker keeps serving", async () => {
  await loadWorkbook("sales-table.xlsx");
  const load = loadRequest("financial-model.xlsx");
  const rowsId = workerScope.send("getRowsBatch", {
    rowCount: 5,
    startRow: 0,
    workbookSheetIndex: 0,
  });
  workerScope.cancel(rowsId);

  const loadResponse = await load.response;
  assert.equal(loadResponse.success, true);
  await delay(50);
  assert.equal(
    workerScope.hasResponse(rowsId),
    false,
    "the cancelled read still ran after the load",
  );

  const cell = await readCell();
  assert.equal(cell.displayValue, "Driver");
});

test("XLSX worker: a failed load keeps the previous workbook addressable", async () => {
  await loadWorkbook("financial-model.xlsx");
  const failed = workerScope.request("load", {
    buffer: sampleBuffer("corrupted.xlsx"),
    limits,
  });
  const failedResponse = await failed.response;
  assert.equal(failedResponse.success, false);
  assert.equal(failedResponse.error.code, "INVALID_SOURCE");
  assert.equal(
    (await readCell()).displayValue,
    "Driver",
    "a failed load left the worker without its previous workbook",
  );
});

test("XLSX worker: cell snapshots carry the resolved number format", async () => {
  const source = new sheetsWasm.Workbook();
  const sheet = source.getSheet(0);
  sheet.setCell("A1", 1234.56);
  sheet.setCellStyleAt(0, 0, {
    numberFormat: { formatType: "custom", formatString: "[$-407]#,##0.00" },
  });
  sheet.free();
  const bytes = source.saveXlsxBytes();
  source.free();
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);

  const payload = { buffer, limits };
  if (!wasmSourceSent) {
    payload.wasmSource = wasmBuffer;
    wasmSourceSent = true;
  }
  const loaded = await workerScope.request("load", payload).response;
  assert.equal(loaded.success, true, `load failed: ${JSON.stringify(loaded)}`);

  const snapshot = await readCell(0, 0, 0);
  assert.equal(snapshot.displayValue, "1,234.56");
  assert.deepEqual(snapshot.numberFormat, {
    formatType: "custom",
    formatString: "[$-407]#,##0.00",
  });
  assert.equal(snapshot.formula, "");
});

test("XLSX worker: AutoFilter rows expose formatted values and honour the row cap", async () => {
  const loaded = await loadWorkbook("sales-table.xlsx");
  assert.deepEqual(
    loaded.sheets[0].autoFilterRanges.map((range) => [range.reference, range.source, range.tableName]),
    [["A1:H181", "table", "SalesRecords"]],
  );

  const capped = await workerScope.request("getAutoFilterRows", {
    endCol: 2,
    endRow: 4,
    maxRows: 2,
    startCol: 0,
    startRow: 1,
    workbookSheetIndex: 0,
  }).response;
  assert.equal(capped.success, true, `AutoFilter rows failed: ${JSON.stringify(capped)}`);
  assert.equal(capped.result.truncated, true);
  assert.deepEqual(capped.result.rows, [
    { row: 1, values: ["North", "Rep 2", "XLSX Engine"] },
    { row: 2, values: ["South", "Rep 3", "Doc Extract"] },
  ]);

  const complete = await workerScope.request("getAutoFilterRows", {
    endCol: 0,
    endRow: 3,
    maxRows: 10,
    startCol: 0,
    startRow: 1,
    workbookSheetIndex: 0,
  }).response;
  assert.equal(complete.success, true);
  assert.equal(complete.result.truncated, false);
  assert.deepEqual(
    complete.result.rows.map((row) => row.values[0]),
    ["North", "South", "APAC"],
  );
});
