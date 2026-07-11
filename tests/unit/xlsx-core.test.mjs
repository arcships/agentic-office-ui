import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueXlsx = createRequire(
  new URL("../../packages/vue-xlsx/package.json", import.meta.url),
);
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
const core = await import(
  pathToFileURL(requireFromVueXlsx.resolve("@arcships/xlsx-core")).href
);
const runtimeEntry = await import(
  pathToFileURL(requireFromVueXlsx.resolve("@arcships/xlsx-core/runtime")).href
);
const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

function bytesFromFile(name) {
  return new Uint8Array(readFileSync(new URL(name, samples)));
}

const wasmPath = requireFromVueXlsx.resolve(
  "@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm",
);
const wasmBytes = readFileSync(wasmPath);
core.setWasmSource(
  wasmBytes.buffer.slice(
    wasmBytes.byteOffset,
    wasmBytes.byteOffset + wasmBytes.byteLength,
  ),
);
const wasm = await core.getSheetsWasmModule();

class ControlledXlsxWorker {
  constructor({ respond = true, throwOnPost = false } = {}) {
    this.respond = respond;
    this.throwOnPost = throwOnPost;
    this.listeners = new Map();
    this.terminated = false;
    this.terminateCount = 0;
    this.messages = [];
  }

  addEventListener(type, listener) {
    this.listeners.set(type, [...(this.listeners.get(type) || []), listener]);
  }

  removeEventListener(type, listener) {
    this.listeners.set(
      type,
      (this.listeners.get(type) || []).filter((item) => item !== listener),
    );
  }

  emit(type, event) {
    for (const listener of [...(this.listeners.get(type) || [])]) {
      listener(event);
    }
  }

  listenerCount(type) {
    return (this.listeners.get(type) || []).length;
  }

  postMessage(message) {
    if (this.throwOnPost) {
      throw new Error("postMessage failed");
    }
    this.messages.push(message);
    if (!this.respond) return;
    const delay = message.payload?.col === 0 ? 30 : 0;
    setTimeout(() => {
      if (this.terminated) return;
      const result = {
        displayValue: `column-${message.payload?.col}`,
        formula: `formula-${message.payload?.col}`,
      };
      for (const listener of this.listeners.get("message") || []) {
        listener({ data: { id: message.id, success: true, result } });
      }
    }, delay);
  }

  terminate() {
    if (!this.terminated) {
      this.terminateCount += 1;
    }
    this.terminated = true;
  }
}

class ControlledAbortSignal {
  constructor({ aborted = false } = {}) {
    this.aborted = aborted;
    this.listeners = new Set();
    this.addCount = 0;
    this.removeCount = 0;
  }

  addEventListener(type, listener) {
    if (type !== "abort") return;
    this.addCount += 1;
    this.listeners.add(listener);
  }

  removeEventListener(type, listener) {
    if (type !== "abort") return;
    this.removeCount += 1;
    this.listeners.delete(listener);
  }

  abort() {
    if (this.aborted) return;
    this.aborted = true;
    for (const listener of [...this.listeners]) {
      listener();
    }
  }

  get listenerCount() {
    return this.listeners.size;
  }
}

test("XLSX parses a normal workbook fixture", () => {
  const workbook = wasm.Workbook.fromBytes(bytesFromFile("financial-model.xlsx"));
  assert.ok(workbook.sheetCount > 0);
  assert.ok(workbook.sheetNames.length > 0);
  const sheet = workbook.getSheet(0);
  assert.ok(sheet.usedRange());
  workbook.free();
});

test("XLSX structure parser exposes real table definitions for sorting", () => {
  const assets = core.parseWorkbookStructureAssets(bytesFromFile("sales-table.xlsx"));
  const tables = core.normalizeWorkbookTableMetadata(
    assets.tableMetadataByWorkbookSheetIndex[0] ?? [],
  );
  assert.equal(tables.length, 1);
  assert.equal(tables[0].name, "SalesRecords");
  assert.equal(tables[0].reference, "A1:H181");
  assert.deepEqual(tables[0].start, { row: 0, col: 0 });
  assert.deepEqual(tables[0].end, { row: 180, col: 7 });
  assert.deepEqual(tables[0].columns.map((column) => column.name), [
    "Region", "Rep", "Product", "Month", "Units", "Unit Price", "Revenue", "Notes",
  ]);
});

test("XLSX resolves a chartsheet drawing into a renderable chart", () => {
  const bytes = bytesFromFile("charts-images.xlsx");
  const structure = core.parseWorkbookStructureAssets(bytes);
  assert.deepEqual(structure.sheetStatesByWorkbookSheetIndex[0]?.mergedRegions, [{
    start: { row: 6, col: 0 },
    end: { row: 7, col: 2 },
  }]);
  const workbook = wasm.Workbook.fromBytes(bytes);
  try {
    const imageAssets = core.parseWorkbookChartStyleAssets(bytes);
    const assets = core.loadWorkbookChartAssets(
      workbook,
      imageAssets,
      new Map([[0, 0]]),
    );
    assert.equal(assets.chartsheets.length, 1);
    const chartsheet = assets.chartsheets[0];
    assert.equal(chartsheet.name, "Revenue Chart");
    assert.equal(chartsheet.chartPath, "xl/charts/chart2.xml");
    assert.equal(chartsheet.charts?.length, 1);
    const chart = chartsheet.charts[0];
    assert.deepEqual(chartsheet.chartIds, [chart.id]);
    assert.equal(chart.chartType, "ColumnClustered");
    assert.equal(chart.title, "Revenue by Quarter");
    assert.equal(chart.editable, false);
    assert.deepEqual(chart.series[0].categories, ["Q1", "Q2", "Q3", "Q4"]);
    assert.deepEqual(chart.series[0].values, [120000, 145000, 168000, 210000]);
  } finally {
    workbook.free();
  }
});

test("XLSX rejects empty and corrupted input", () => {
  assert.throws(
    () => wasm.Workbook.fromBytes(new Uint8Array()),
    /Unable to detect file format/,
  );
  assert.throws(
    () => wasm.Workbook.fromBytes(bytesFromFile("corrupted.xlsx")),
    /Unable to detect file format/,
  );
});

test("XLSX roundtrip preserves values, formula, calculation and cell style", () => {
  const workbook = new wasm.Workbook();
  const sheetIndex = workbook.addSheet("Unit Roundtrip");
  const sheet = workbook.getSheet(sheetIndex);
  sheet.setCell("A1", "roundtrip");
  sheet.setCell("B1", 2);
  sheet.setFormula("C1", "=B1*3");
  sheet.setCellStyle("B1", {
    font: {
      bold: true,
      color: { colorType: "rgb", hex: "FF0000" },
    },
    numberFormat: { formatType: "custom", formatString: "0.00" },
  });
  workbook.calculate();

  const exported = workbook.saveXlsxBytes();
  assert.ok(exported.byteLength > 0);
  const reloaded = wasm.Workbook.fromBytes(exported);
  const reloadedSheet = reloaded.getSheetByName("Unit Roundtrip");
  assert.equal(reloadedSheet.getCell("A1").toString(), "roundtrip");
  assert.equal(reloadedSheet.getFormulaAt(0, 2), "=B1*3");
  assert.equal(reloadedSheet.getCalculatedValue("C1").toString(), "6");
  assert.equal(reloadedSheet.getFormattedValue("B1"), "2.00");
  const style = reloadedSheet.getCellStyle("B1");
  assert.equal(style.font.bold, true);
  assert.equal(style.font.color.hex, "FFFF0000");
  assert.equal(style.numberFormat.formatString, "0.00");

  reloaded.free();
  workbook.free();
});

test("XLSX Worker client matches out-of-order responses to their requests", async () => {
  const worker = new ControlledXlsxWorker();
  const client = new core.XlsxWorkerClient({ createWorker: () => worker });
  const oldRequest = client.getCellSnapshot(0, 0, 0);
  const latestRequest = client.getCellSnapshot(0, 0, 1);
  const [oldResult, latestResult] = await Promise.all([oldRequest, latestRequest]);
  assert.equal(oldResult.displayValue, "column-0");
  assert.equal(latestResult.displayValue, "column-1");
  client.dispose();
  assert.equal(worker.terminated, true);
});

test("XLSX Worker client does not post any pre-cancelled request", async () => {
  const worker = new ControlledXlsxWorker({ respond: false });
  const client = new core.XlsxWorkerClient({
    createWorker: () => worker,
    wasmSource: "https://assets.example/worker.wasm",
  });
  const signals = Array.from(
    { length: 4 },
    () => new ControlledAbortSignal({ aborted: true }),
  );

  await Promise.all([
    assert.rejects(
      client.loadWorkbook(new ArrayBuffer(4), false, false, signals[0]),
      (error) => error?.name === "AbortError",
    ),
    assert.rejects(
      client.parseCharts(new ArrayBuffer(4), false, false, signals[1]),
      (error) => error?.name === "AbortError",
    ),
    assert.rejects(
      client.getCellSnapshot(0, 0, 0, signals[2]),
      (error) => error?.name === "AbortError",
    ),
    assert.rejects(
      client.getRowsBatch(0, 0, 10, signals[3]),
      (error) => error?.name === "AbortError",
    ),
  ]);
  assert.equal(worker.messages.length, 0);
  for (const signal of signals) {
    assert.equal(signal.addCount, 0);
    assert.equal(signal.removeCount, 0);
  }

  const next = client.loadWorkbook(new ArrayBuffer(4));
  assert.equal(
    worker.messages[0].payload.wasmSource,
    "https://assets.example/worker.wasm",
  );
  worker.emit("message", {
    data: { id: worker.messages[0].id, success: true, result: {} },
  });
  await next;
  client.dispose();
});

test("XLSX Worker cancellation terminates the client and rejects every pending request", async () => {
  const worker = new ControlledXlsxWorker({ respond: false });
  const client = new core.XlsxWorkerClient({ createWorker: () => worker });
  const cancelledSignal = new ControlledAbortSignal();
  const activeSignal = new ControlledAbortSignal();
  const cancelled = client.getCellSnapshot(0, 0, 0, cancelledSignal);
  const active = client.getCellSnapshot(0, 0, 1, activeSignal);
  const cancelledMessage = worker.messages[0];
  const activeMessage = worker.messages[1];
  const cancelledRejection = assert.rejects(
    cancelled,
    (error) => error?.name === "AbortError",
  );
  const activeRejection = assert.rejects(
    active,
    (error) => error?.name === "AbortError",
  );

  cancelledSignal.abort();
  await Promise.all([cancelledRejection, activeRejection]);
  assert.equal(cancelledSignal.listenerCount, 0);
  assert.equal(cancelledSignal.removeCount, 1);
  assert.equal(activeSignal.listenerCount, 0);
  assert.equal(activeSignal.removeCount, 1);
  assert.equal(worker.terminated, true);
  assert.equal(worker.terminateCount, 1);

  worker.emit("message", {
    data: {
      id: cancelledMessage.id,
      success: true,
      result: { displayValue: "late", formula: "late" },
    },
  });
  worker.emit("message", {
    data: {
      id: activeMessage.id,
      success: true,
      result: { displayValue: "active", formula: "active" },
    },
  });
  await assert.rejects(
    client.getRowsBatch(0, 0, 1),
    (error) => error?.name === "AbortError",
  );
  client.dispose();
});

test("XLSX Worker client cleans AbortSignal on success and response error", async () => {
  const worker = new ControlledXlsxWorker({ respond: false });
  const client = new core.XlsxWorkerClient({ createWorker: () => worker });
  const successSignal = new ControlledAbortSignal();
  const success = client.getRowsBatch(0, 0, 1, successSignal);
  worker.emit("message", {
    data: { id: worker.messages[0].id, success: true, result: ["row"] },
  });
  assert.deepEqual(await success, ["row"]);
  assert.equal(successSignal.addCount, 1);
  assert.equal(successSignal.removeCount, 1);
  assert.equal(successSignal.listenerCount, 0);

  const errorSignal = new ControlledAbortSignal();
  const failed = client.getRowsBatch(0, 1, 1, errorSignal);
  worker.emit("message", {
    data: { id: worker.messages[1].id, success: false, error: "bad response" },
  });
  await assert.rejects(failed, /bad response/);
  assert.equal(errorSignal.addCount, 1);
  assert.equal(errorSignal.removeCount, 1);
  assert.equal(errorSignal.listenerCount, 0);
  assert.equal(worker.terminated, true);
  client.dispose();
});

test("XLSX Worker client cleans AbortSignal when postMessage throws", async () => {
  const worker = new ControlledXlsxWorker({ respond: false, throwOnPost: true });
  const client = new core.XlsxWorkerClient({
    createWorker: () => worker,
    wasmSource: "https://assets.example/retry.wasm",
  });
  const signal = new ControlledAbortSignal();
  await assert.rejects(
    client.loadWorkbook(new ArrayBuffer(4), false, false, signal),
    /postMessage failed/,
  );
  assert.equal(signal.addCount, 1);
  assert.equal(signal.removeCount, 1);
  assert.equal(signal.listenerCount, 0);
  assert.equal(worker.messages.length, 0);
  assert.equal(worker.terminated, true);
  await assert.rejects(
    client.loadWorkbook(new ArrayBuffer(4)),
    /postMessage failed/,
  );
  client.dispose();
});

test("XLSX Worker disposal cancels every pending request and cleans once", async () => {
  const worker = new ControlledXlsxWorker({ respond: false });
  let disposeCount = 0;
  const client = new core.XlsxWorkerClient({
    createWorker: () => worker,
    onDispose: () => { disposeCount += 1; },
  });
  const firstSignal = new ControlledAbortSignal();
  const secondSignal = new ControlledAbortSignal();
  const first = client.getRowsBatch(0, 0, 10, firstSignal);
  const second = client.getCellSnapshot(0, 0, 0, secondSignal);
  const firstRejection = assert.rejects(
    first,
    (error) => error?.name === "AbortError",
  );
  const secondRejection = assert.rejects(
    second,
    (error) => error?.name === "AbortError",
  );
  client.dispose();
  client.dispose();
  await Promise.all([firstRejection, secondRejection]);
  assert.equal(worker.terminated, true);
  assert.equal(worker.terminateCount, 1);
  assert.equal(disposeCount, 1);
  assert.equal(firstSignal.removeCount, 1);
  assert.equal(secondSignal.removeCount, 1);
  assert.equal(firstSignal.listenerCount, 0);
  assert.equal(secondSignal.listenerCount, 0);
  assert.equal(worker.listenerCount("message"), 0);
  assert.equal(worker.listenerCount("error"), 0);
  assert.equal(worker.listenerCount("messageerror"), 0);

  const messageCount = worker.messages.length;
  await assert.rejects(
    client.getRowsBatch(0, 0, 1),
    (error) => error?.name === "AbortError",
  );
  assert.equal(worker.messages.length, messageCount);
});

test("XLSX Worker error and messageerror permanently close the client", async (t) => {
  for (const faultType of ["error", "messageerror"]) {
    await t.test(faultType, async () => {
      const worker = new ControlledXlsxWorker({ respond: false });
      let disposeCount = 0;
      const client = new core.XlsxWorkerClient({
        createWorker: () => worker,
        onDispose: () => { disposeCount += 1; },
      });
      const firstSignal = new ControlledAbortSignal();
      const secondSignal = new ControlledAbortSignal();
      const first = client.getRowsBatch(0, 0, 1, firstSignal);
      const second = client.getCellSnapshot(0, 0, 0, secondSignal);
      const expectedMessage = faultType === "error"
        ? "worker exploded"
        : "Worker returned an unreadable response.";
      const firstRejection = assert.rejects(
        first,
        (error) => error?.name === "XlsxWorkerError" &&
          error?.message === expectedMessage,
      );
      const secondRejection = assert.rejects(
        second,
        (error) => error?.name === "XlsxWorkerError" &&
          error?.message === expectedMessage,
      );

      worker.emit(faultType, faultType === "error" ? { message: expectedMessage } : {});
      await Promise.all([firstRejection, secondRejection]);
      assert.equal(firstSignal.removeCount, 1);
      assert.equal(secondSignal.removeCount, 1);
      assert.equal(firstSignal.listenerCount, 0);
      assert.equal(secondSignal.listenerCount, 0);
      assert.equal(worker.terminated, true);
      assert.equal(worker.terminateCount, 1);
      assert.equal(disposeCount, 1);
      assert.equal(worker.listenerCount("message"), 0);
      assert.equal(worker.listenerCount("error"), 0);
      assert.equal(worker.listenerCount("messageerror"), 0);

      const messageCount = worker.messages.length;
      await assert.rejects(
        client.getRowsBatch(0, 0, 1),
        (error) => error?.name === "XlsxWorkerError" &&
          error?.message === expectedMessage,
      );
      assert.equal(worker.messages.length, messageCount);
      client.dispose();
      assert.equal(worker.terminateCount, 1);
      assert.equal(disposeCount, 1);
    });
  }
});

test("XLSX URL loading uses shared policy, controlled fetch and safe diagnostics URL", async () => {
  const bytes = bytesFromFile("sales-table.xlsx");
  const calls = [];
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/viewer/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async (url, init) => {
      calls.push({ url, init });
      return {
        ok: true,
        status: 200,
        url,
        headers: { get: () => null },
        arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
      };
    },
  };
  const result = await core.loadVerifiedXlsxSource("sales-table.xlsx?token=secret", policy);
  assert.equal(result.fileName, "sales-table.xlsx");
  assert.equal(result.resolvedUrl, "https://app.example/viewer/sales-table.xlsx");
  assert.equal(calls[0].init.redirect, "error");
  assert.equal(calls[0].init.credentials, "omit");
  assert.equal(JSON.stringify(result).includes("secret"), false);
});

test("XLSX URL loading rejects an oversized Content-Length before reading the body", async () => {
  let arrayBufferCalls = 0;
  let bodyCancelCalls = 0;
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async (url) => ({
      ok: true,
      status: 200,
      url,
      headers: { get: (name) => name.toLowerCase() === "content-length" ? "8" : null },
      body: { cancel: async () => { bodyCancelCalls += 1; } },
      arrayBuffer: async () => {
        arrayBufferCalls += 1;
        return new ArrayBuffer(8);
      },
    }),
  };

  await assert.rejects(
    core.loadVerifiedXlsxSource("book.xlsx", policy, undefined, { maxInputBytes: 4 }),
    (error) => error?.name === "XlsxSourceError" &&
      error?.code === "LIMIT_EXCEEDED" &&
      error?.phase === "input" &&
      error?.limit === "maxInputBytes" &&
      error?.actual === 8 &&
      error?.allowed === 4,
  );
  assert.equal(arrayBufferCalls, 0);
  assert.equal(bodyCancelCalls, 1);
});

test("XLSX URL loading cancels a streaming response as soon as maxInputBytes is exceeded", async () => {
  const chunks = [
    new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
    new Uint8Array([1, 2, 3, 4]),
    new Uint8Array([5, 6, 7, 8]),
  ];
  let reads = 0;
  let cancelCalls = 0;
  let releaseCalls = 0;
  let fallbackArrayBufferCalls = 0;
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async (url) => ({
      ok: true,
      status: 200,
      url,
      headers: { get: () => null },
      body: {
        getReader: () => ({
          read: async () => {
            const value = chunks[reads];
            reads += 1;
            return value ? { done: false, value } : { done: true, value: undefined };
          },
          cancel: async () => { cancelCalls += 1; },
          releaseLock: () => { releaseCalls += 1; },
        }),
      },
      arrayBuffer: async () => {
        fallbackArrayBufferCalls += 1;
        return new ArrayBuffer(0);
      },
    }),
  };

  await assert.rejects(
    core.loadVerifiedXlsxSource("book.xlsx", policy, undefined, { maxInputBytes: 5 }),
    (error) => error?.name === "XlsxSourceError" &&
      error?.code === "LIMIT_EXCEEDED" &&
      error?.phase === "input" &&
      error?.limit === "maxInputBytes" &&
      error?.actual === 8 &&
      error?.allowed === 5,
  );
  assert.equal(reads, 2);
  assert.equal(cancelCalls, 1);
  assert.equal(releaseCalls, 1);
  assert.equal(fallbackArrayBufferCalls, 0);
});

test("XLSX dangerous URL is rejected before fetch with the stable shared code", async () => {
  let requestCount = 0;
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async () => { requestCount += 1; throw new Error("must not fetch"); },
  };
  await assert.rejects(
    core.loadVerifiedXlsxSource("https://other.example/book.xlsx?token=secret", policy),
    (error) => error?.name === "XlsxSourceError" &&
      error?.code === "SOURCE_NOT_ALLOWED" &&
      !JSON.stringify(error).includes("secret"),
  );
  assert.equal(requestCount, 0);
});

test("XLSX runtimes isolate Worker and WASM configuration from legacy globals", async () => {
  assert.equal(typeof core.createXlsxRuntime, "function");
  const workersA = [];
  const workersB = [];
  const configA = {
    wasmSource: "https://assets.example/a.wasm",
    parseOptions: { showHiddenSheets: true, skipXmlParsing: true },
    limits: { maxArchiveEntries: 7, maxWorksheets: 3 },
    createWorker: () => {
      const worker = new ControlledXlsxWorker();
      workersA.push(worker);
      return worker;
    },
  };
  const runtimeA = core.createXlsxRuntime(configA);
  const runtimeB = core.createXlsxRuntime({
    wasmSource: "https://assets.example/b.wasm",
    createWorker: () => {
      const worker = new ControlledXlsxWorker();
      workersB.push(worker);
      return worker;
    },
  });
  configA.wasmSource = "https://mutated.invalid/a.wasm";
  configA.parseOptions.showHiddenSheets = false;
  configA.limits.maxArchiveEntries = 1;

  const clientA = runtimeA.createWorkerClient();
  const clientB = runtimeB.createWorkerClient();
  await Promise.all([
    clientA.loadWorkbook(new ArrayBuffer(4)),
    clientB.loadWorkbook(new ArrayBuffer(4)),
  ]);
  assert.notEqual(runtimeA.id, runtimeB.id);
  assert.equal(workersA[0].messages[0].payload.wasmSource, "https://assets.example/a.wasm");
  assert.equal(workersB[0].messages[0].payload.wasmSource, "https://assets.example/b.wasm");
  assert.deepEqual(runtimeA.parseOptions, { showHiddenSheets: true, skipXmlParsing: true });
  assert.equal(Object.isFrozen(runtimeA.limits), true);
  assert.equal(runtimeA.limits.maxArchiveEntries, 7);
  assert.equal(workersA[0].messages[0].payload.limits.maxWorksheets, 3);
  await clientA.parseCharts(new ArrayBuffer(4));
  assert.equal(workersA[0].messages[1].payload.wasmSource, undefined);
  assert.throws(
    () => core.setWasmSource("https://legacy.invalid/late.wasm"),
    /before the default main-thread WASM module initializes/,
  );

  runtimeA.dispose();
  runtimeB.dispose();
  assert.equal(workersA[0].terminated, true);
  assert.equal(workersB[0].terminated, true);
  let disposedError;
  assert.throws(
    () => runtimeA.createWorkerClient(),
    (error) => {
      disposedError = error;
      return error instanceof runtimeEntry.XlsxRuntimeError &&
        error.name === "XlsxRuntimeError" &&
        error.code === "RUNTIME_DISPOSED";
    },
  );
  assert.deepEqual(disposedError.toJSON(), {
    name: "XlsxRuntimeError",
    code: "RUNTIME_DISPOSED",
    message: "XLSX 运行实例已经销毁。",
    format: "xlsx",
    runtimeId: runtimeA.id,
  });
});

test("XLSX Worker preserves structured budget errors", async () => {
  const worker = new ControlledXlsxWorker({ respond: false });
  const client = new core.XlsxWorkerClient({ createWorker: () => worker });
  const pending = client.loadWorkbook(new ArrayBuffer(4));
  worker.emit("message", {
    data: {
      id: worker.messages[0].id,
      success: false,
      error: {
        code: "LIMIT_EXCEEDED",
        message: "shared strings exceeded",
        phase: "xml",
        limit: "maxSharedStrings",
        actual: 3,
        allowed: 2,
      },
    },
  });
  await assert.rejects(
    pending,
    (error) => error instanceof core.XlsxWorkerError &&
      error.code === "LIMIT_EXCEEDED" &&
      error.phase === "xml" &&
      error.limit === "maxSharedStrings" &&
      error.actual === 3 &&
      error.allowed === 2,
  );
  client.dispose();
});

test("XLSX Worker failures keep a stable public load error", () => {
  const error = core.toXlsxLoadError(
    new core.XlsxWorkerError("WORKER_FAILED", "worker initialization failed", {
      phase: "parse",
    }),
    "file",
  );
  assert.deepEqual(error, {
    code: "WORKER_UNAVAILABLE",
    message: "worker initialization failed",
    phase: "parse",
    sourceKind: "file",
  });
});

test("XLSX Worker timeout terminates the client for clean recovery", async () => {
  const worker = new ControlledXlsxWorker({ respond: false });
  const client = new core.XlsxWorkerClient({
    createWorker: () => worker,
    limits: { maxParseMs: 5 },
  });
  await assert.rejects(
    client.loadWorkbook(new ArrayBuffer(4)),
    (error) => error instanceof core.XlsxWorkerError &&
      error.code === "TIMEOUT" &&
      error.actual > error.allowed,
  );
  assert.equal(worker.terminated, true);
  client.dispose();
});
