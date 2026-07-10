import assert from "node:assert/strict";
import { readFileSync, readdirSync } from "node:fs";
import { test } from "node:test";

const packageRoot = new URL("../../packages/office-runtime/", import.meta.url);
const runtime = await import(new URL("dist/index.js", packageRoot).href);

function sourceFiles(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const url = new URL(`${entry.name}${entry.isDirectory() ? "/" : ""}`, directory);
    return entry.isDirectory() ? sourceFiles(url) : entry.name.endsWith(".ts") ? [url] : [];
  });
}

function createCountedAbortSignal() {
  const controller = new AbortController();
  const signal = controller.signal;
  const originalAdd = signal.addEventListener.bind(signal);
  const originalRemove = signal.removeEventListener.bind(signal);
  const activeAbortListeners = new Set();
  let addCalls = 0;
  let removeCalls = 0;

  signal.addEventListener = (type, listener, options) => {
    if (type === "abort") {
      addCalls += 1;
      activeAbortListeners.add(listener);
    }
    return originalAdd(type, listener, options);
  };
  signal.removeEventListener = (type, listener, options) => {
    if (type === "abort") {
      removeCalls += 1;
      activeAbortListeners.delete(listener);
    }
    return originalRemove(type, listener, options);
  };

  return {
    controller,
    signal,
    counts: () => ({
      addCalls,
      removeCalls,
      active: activeAbortListeners.size,
    }),
  };
}

function assertStale(task) {
  assert.equal(task.isCurrent(), false);
  assert.throws(
    () => task.assertCurrent(),
    (error) => error?.name === "OfficeLoadError" && error?.code === "STALE_RESULT",
  );
}

test("office-runtime is private, framework-free and has no UI DOM dependency", () => {
  const manifest = JSON.parse(readFileSync(new URL("package.json", packageRoot), "utf8"));
  assert.equal(manifest.private, true);
  assert.equal(manifest.publishConfig, undefined);
  assert.equal(JSON.stringify(manifest).includes("vue"), false);

  for (const file of sourceFiles(new URL("src/", packageRoot))) {
    const source = readFileSync(file, "utf8");
    assert.doesNotMatch(source, /(?:from\s+["']vue["']|\.vue["'])/);
    assert.doesNotMatch(source, /\b(?:window|document|HTMLElement|HTMLCanvasElement|ImageBitmap)\b/);
  }
});

test("load context snapshots source, resources, limits and stable task identity", () => {
  const resources = { wasmUrl: "https://assets.example/docx.wasm", workerUrl: "https://assets.example/docx-worker.js" };
  const limits = { maxInputBytes: 1024 };
  const sequence = runtime.createOfficeTaskSequence("runtime-a");
  const context = runtime.createOfficeLoadContext({
    runtimeId: "runtime-a",
    taskId: sequence.next(),
    source: { kind: "bytes", bytes: new ArrayBuffer(8), name: "sample.docx" },
    resources,
    limits,
    signal: new AbortController().signal,
  });

  resources.wasmUrl = "https://mutated.invalid/wasm";
  limits.maxInputBytes = 1;
  assert.equal(context.runtimeId, "runtime-a");
  assert.equal(context.taskId, "runtime-a:1");
  assert.equal(context.sourceKind, "bytes");
  assert.equal(context.resources.wasmUrl, "https://assets.example/docx.wasm");
  assert.equal(context.limits.maxInputBytes, 1024);
  assert.equal(Object.isFrozen(context), true);
  assert.equal(Object.isFrozen(context.resources), true);
  assert.equal(Object.isFrozen(context.limits), true);
});

test("task sequences are instance-owned and shared by loaders of one runtime", () => {
  const sequenceA = runtime.createOfficeTaskSequence("runtime-a");
  const sequenceB = runtime.createOfficeTaskSequence("runtime-b");
  const loaderA = runtime.createLatestTaskCoordinator(sequenceA);
  const loaderB = runtime.createLatestTaskCoordinator(sequenceA);
  const otherRuntimeLoader = runtime.createLatestTaskCoordinator(sequenceB);

  const first = loaderA.start({ kind: "bytes", bytes: new ArrayBuffer(1) });
  const second = loaderB.start({ kind: "bytes", bytes: new ArrayBuffer(1) });
  const isolated = otherRuntimeLoader.start({ kind: "bytes", bytes: new ArrayBuffer(1) });
  assert.equal(first.context.taskId, "runtime-a:1");
  assert.equal(second.context.taskId, "runtime-a:2");
  assert.equal(isolated.context.taskId, "runtime-b:1");
  assert.equal(first.isCurrent(), true);
  assert.equal(second.isCurrent(), true);

  const replacement = loaderA.start({ kind: "bytes", bytes: new ArrayBuffer(1) });
  assert.equal(replacement.context.taskId, "runtime-a:3");
  assert.equal(first.signal.aborted, true);
  assert.equal(first.isCurrent(), false);
  assert.throws(() => first.assertCurrent(), (error) => error?.code === "STALE_RESULT");

  loaderA.dispose();
  loaderB.dispose();
  otherRuntimeLoader.dispose();
  sequenceA.dispose();
  sequenceB.dispose();
});

test("external and internal cancellation combine without mutating the caller signal", () => {
  const sequence = runtime.createOfficeTaskSequence("cancel-runtime");
  const loader = runtime.createLatestTaskCoordinator(sequence);
  const external = new AbortController();
  const task = loader.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: external.signal },
  );
  loader.cancel();
  assert.equal(task.signal.aborted, true);
  assert.equal(external.signal.aborted, false);

  const alreadyAborted = new AbortController();
  alreadyAborted.abort();
  const cancelledAtStart = loader.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: alreadyAborted.signal },
  );
  assert.equal(cancelledAtStart.signal.aborted, true);
  loader.dispose();
  assert.throws(
    () => loader.start({ kind: "bytes", bytes: new ArrayBuffer(1) }),
    (error) => error?.code === "RUNTIME_DISPOSED",
  );
  sequence.dispose();
});

test("finished tasks stop being current and remove the external abort listener once", () => {
  const sequence = runtime.createOfficeTaskSequence("finish-lifecycle");
  const coordinator = runtime.createLatestTaskCoordinator(sequence);
  const external = createCountedAbortSignal();
  const task = coordinator.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: external.signal },
  );

  assert.equal(task.isCurrent(), true);
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 0, active: 1 });
  task.finish();
  assertStale(task);
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 1, active: 0 });

  task.finish();
  coordinator.cancel();
  coordinator.dispose();
  coordinator.dispose();
  sequence.dispose();
  sequence.dispose();
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
});

test("replacing a task removes its external abort listener and makes it stale", () => {
  const sequence = runtime.createOfficeTaskSequence("replace-lifecycle");
  const coordinator = runtime.createLatestTaskCoordinator(sequence);
  const oldExternal = createCountedAbortSignal();
  const nextExternal = createCountedAbortSignal();
  const oldTask = coordinator.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: oldExternal.signal },
  );
  const nextTask = coordinator.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: nextExternal.signal },
  );

  assert.equal(oldTask.signal.aborted, true);
  assertStale(oldTask);
  assert.deepEqual(oldExternal.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
  assert.equal(nextTask.isCurrent(), true);
  assert.deepEqual(nextExternal.counts(), { addCalls: 1, removeCalls: 0, active: 1 });

  oldTask.finish();
  oldTask.finish();
  nextTask.finish();
  nextTask.finish();
  coordinator.dispose();
  sequence.dispose();
  assert.deepEqual(oldExternal.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
  assert.deepEqual(nextExternal.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
});

test("cancel is idempotent and removes the external abort listener once", () => {
  const sequence = runtime.createOfficeTaskSequence("cancel-lifecycle");
  const coordinator = runtime.createLatestTaskCoordinator(sequence);
  const external = createCountedAbortSignal();
  const task = coordinator.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: external.signal },
  );

  coordinator.cancel();
  coordinator.cancel();
  assert.equal(task.signal.aborted, true);
  assert.equal(external.signal.aborted, false);
  assertStale(task);
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 1, active: 0 });

  task.finish();
  task.finish();
  coordinator.dispose();
  sequence.dispose();
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
});

test("dispose is idempotent and removes the external abort listener once", () => {
  const sequence = runtime.createOfficeTaskSequence("dispose-lifecycle");
  const coordinator = runtime.createLatestTaskCoordinator(sequence);
  const external = createCountedAbortSignal();
  const task = coordinator.start(
    { kind: "bytes", bytes: new ArrayBuffer(1) },
    { signal: external.signal },
  );

  coordinator.dispose();
  coordinator.dispose();
  assert.equal(task.signal.aborted, true);
  assertStale(task);
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
  assert.throws(
    () => coordinator.start({ kind: "bytes", bytes: new ArrayBuffer(1) }),
    (error) => error?.code === "RUNTIME_DISPOSED",
  );

  task.finish();
  task.finish();
  sequence.dispose();
  sequence.dispose();
  assert.deepEqual(external.counts(), { addCalls: 1, removeCalls: 1, active: 0 });
});

test("URL policy rejects dangerous sources before fetch and strips secrets", async () => {
  let requestCount = 0;
  const policy = {
    enabled: true,
    baseUrl: "https://app.example/viewer/",
    allowRelativeUrl: true,
    allowedProtocols: ["https:"],
    allowedOrigins: ["https://app.example"],
    fetch: async () => {
      requestCount += 1;
      throw new Error("must not fetch");
    },
  };

  await assert.rejects(
    runtime.loadOfficeSource({ kind: "url", url: "javascript:alert(1)" }, { urlPolicy: policy }),
    (error) => error?.name === "OfficeLoadError" && error?.code === "SOURCE_NOT_ALLOWED",
  );
  await assert.rejects(
    runtime.loadOfficeSource({ kind: "url", url: "https://other.example/a.xlsx?token=secret#x" }, { urlPolicy: policy }),
    (error) => error?.code === "SOURCE_NOT_ALLOWED" && !JSON.stringify(error).includes("secret"),
  );
  assert.equal(requestCount, 0);
  assert.equal(
    runtime.sanitizeOfficeUrl("https://app.example/a.xlsx?token=secret#x"),
    "https://app.example/a.xlsx",
  );
});

test("source loading uses controlled fetch and enforces the input byte limit", async () => {
  const calls = [];
  const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);
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
        headers: { get: (name) => name.toLowerCase() === "content-length" ? String(bytes.byteLength) : null },
        arrayBuffer: async () => bytes.buffer.slice(0),
      };
    },
  };

  const loaded = await runtime.loadOfficeSource(
    { kind: "url", url: "book.xlsx?token=secret" },
    { urlPolicy: policy, limits: { maxInputBytes: bytes.byteLength } },
  );
  assert.deepEqual([...new Uint8Array(loaded.buffer)], [...bytes]);
  assert.equal(loaded.sourceKind, "url");
  assert.equal(loaded.resolvedUrl, "https://app.example/viewer/book.xlsx");
  assert.equal(calls[0].init.redirect, "error");
  assert.equal(calls[0].init.credentials, "omit");

  await assert.rejects(
    runtime.loadOfficeSource(
      { kind: "url", url: "book.xlsx" },
      { urlPolicy: policy, limits: { maxInputBytes: bytes.byteLength - 1 } },
    ),
    (error) => error?.code === "LIMIT_EXCEEDED" && error?.limit === "maxInputBytes",
  );
});

test("all formats receive the same stable abort error and safe diagnostics", () => {
  const abort = new Error("cancelled");
  abort.name = "AbortError";
  for (const format of ["docx", "xlsx", "pdf"]) {
    const error = runtime.toOfficeLoadError(abort, {
      fallbackCode: "INVALID_SOURCE",
      format,
      sourceKind: "url",
      taskId: `${format}:1`,
      url: "https://app.example/file?token=secret",
    });
    assert.equal(error.name, "OfficeLoadError");
    assert.equal(error.code, "ABORTED");
    assert.equal(error.format, format);
    assert.equal(error.taskId, `${format}:1`);
    assert.equal(error.url, "https://app.example/file");
    assert.equal(JSON.stringify(error).includes("secret"), false);
    assert.equal(JSON.stringify(error).includes("stack"), false);
  }

  const event = runtime.createOfficeDiagnostic({
    type: "load-error",
    runtimeId: "safe-runtime",
    taskId: "safe-runtime:1",
    sourceKind: "url",
    url: "https://app.example/file?token=secret#fragment",
  });
  assert.equal(event.url, "https://app.example/file");
  assert.equal(JSON.stringify(event).includes("secret"), false);
  assert.equal(Object.isFrozen(event), true);
});
