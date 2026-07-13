import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);
const runtimeEntry = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core/runtime")).href
);
const vueDocx = await import(
  new URL("../../packages/vue-docx/dist/index.js", import.meta.url).href
);
const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

function arrayBufferFromFile(name) {
  const bytes = readFileSync(new URL(name, samples));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function withWordMainContentType(input, contentType) {
  const archive = unzipSync(new Uint8Array(input));
  const contentTypes = strFromU8(archive["[Content_Types].xml"]);
  archive["[Content_Types].xml"] = strToU8(contentTypes.replace(
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    contentType,
  ));
  const bytes = zipSync(archive);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function collectText(value, result = []) {
  if (!value || typeof value !== "object") return result;
  if (value.type === "text" && typeof value.text === "string") {
    result.push(value.text);
  }
  for (const child of Array.isArray(value) ? value : Object.values(value)) {
    collectText(child, result);
  }
  return result;
}

function workerResult(request, marker) {
  return {
    data: {
      id: request.id,
      type: "success",
      package: { parts: new Map(), binaryAssets: new Map() },
      model: {
        nodes: [{ id: marker, type: "paragraph", children: [] }],
        metadata: {
          sourceParts: 0,
          warnings: [],
          headerSections: [],
          footerSections: [],
          paragraphStyles: [],
        },
      },
      timings: { totalMs: 1, parseMs: 0.5, buildModelMs: 0.5 },
    },
  };
}

test("DOCX surface search returns exact UTF-16 ranges for body and direct table paragraphs", () => {
  const paragraph = (text) => ({
    type: "paragraph",
    children: [{ type: "text", text }],
  });
  const model = {
    nodes: [
      paragraph("Alpha alpha ALPHA"),
      {
        type: "table",
        rows: [{
          type: "table-row",
          cells: [{
            type: "table-cell",
            nodes: [
              paragraph("literal a.b and A.B"),
              {
                type: "table",
                rows: [{
                  type: "table-row",
                  cells: [{ type: "table-cell", nodes: [paragraph("nested a.b")] }],
                }],
              },
            ],
          }],
        }],
      },
    ],
    metadata: {
      sourceParts: 0,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
    },
  };

  const bodyMatches = vueDocx.findDocxSearchMatches(model, "alpha");
  assert.deepEqual(
    bodyMatches.map((match) => [match.range.start.offset, match.range.end.offset]),
    [[0, 5], [6, 11], [12, 17]],
  );
  assert.deepEqual(bodyMatches[0].range.start.location, { kind: "paragraph", nodeIndex: 0 });

  const tableMatches = vueDocx.findDocxSearchMatches(model, "a.b");
  assert.equal(tableMatches.length, 2);
  assert.deepEqual(tableMatches[0].range.start.location, {
    kind: "table-cell",
    tableIndex: 1,
    rowIndex: 0,
    cellIndex: 0,
    paragraphIndex: 0,
  });
  assert.equal(tableMatches.every((match) => match.nodeIndex === 1), true);
});

class ControlledDocxWorker {
  constructor({ delay = 0, marker = "worker", respond = true } = {}) {
    this.delay = delay;
    this.marker = marker;
    this.respond = respond;
    this.listeners = new Map();
    this.messages = [];
    this.terminated = false;
    this.timer = null;
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

  postMessage(request) {
    this.messages.push(request);
    if (!this.respond) return;
    this.timer = setTimeout(() => {
      if (this.terminated) return;
      for (const listener of this.listeners.get("message") || []) {
        listener(workerResult(request, this.marker));
      }
    }, this.delay);
  }

  terminate() {
    this.terminated = true;
    if (this.timer) clearTimeout(this.timer);
  }
}

async function withWorkerGlobal(run) {
  const original = globalThis.Worker;
  globalThis.Worker = ControlledDocxWorker;
  try {
    return await run();
  } finally {
    if (original === undefined) delete globalThis.Worker;
    else globalThis.Worker = original;
  }
}

test("DOCX parses a normal fixture and preserves content through roundtrip", async () => {
  const input = arrayBufferFromFile("invoice-table.docx");
  const first = await core.wasmBuildDocModelFromBytes(input);
  const firstText = collectText(first.model).join(" ");
  assert.match(firstText, /INVOICE INV-2026-0705/);
  assert.ok(first.model.nodes.length > 0);

  const exported = await core.wasmSerializeDocx(first.model, first.package);
  assert.ok(exported.byteLength > 0);
  assert.deepEqual([...new Uint8Array(exported).slice(0, 2)], [0x50, 0x4b]);

  const second = await core.wasmBuildDocModelFromBytes(exported);
  const secondText = collectText(second.model).join(" ");
  assert.match(secondText, /INVOICE INV-2026-0705/);
  assert.equal(second.model.nodes.length, first.model.nodes.length);
});

test("DOCX importer accepts macro-enabled documents and Word templates as read-only inputs", async () => {
  const source = arrayBufferFromFile("invoice-table.docx");
  for (const [format, contentType] of [
    ["docm", "application/vnd.ms-word.document.macroEnabled.main+xml"],
    ["dotx", "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"],
    ["dotm", "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"],
  ]) {
    const result = await core.importDocxBuffer(withWordMainContentType(source, contentType), {
      useWorker: false,
    });
    assert.ok(result.model.nodes.length > 0, `${format} should produce a document model`);
  }
});

test("DOCX rejects empty and corrupted input with a stable parse error", async () => {
  for (const input of [new ArrayBuffer(0), arrayBufferFromFile("corrupted.docx")]) {
    await assert.rejects(
      core.importDocxBuffer(input, { useWorker: false }),
      (error) => error?.name === "DocxImportError" && error?.code === "PARSE_FAILED",
    );
  }
});

test("DOCX runtime entry exposes its public import error constructor", () => {
  assert.equal(runtimeEntry.DocxImportError, core.DocxImportError);
  const error = new runtimeEntry.DocxImportError("PARSE_FAILED", "broken document");
  assert.equal(error.name, "DocxImportError");
  assert.equal(error.code, "PARSE_FAILED");
});

test("DOCX page margin styles include CSS pixel units", () => {
  assert.deepEqual(
    core.pageMarginPaddingStyle({ top: 56, right: 48, bottom: 40, left: 32 }),
    {
      paddingTop: "56px",
      paddingRight: "48px",
      paddingBottom: "40px",
      paddingLeft: "32px",
    },
  );
});

test("DOCX abort terminates the active Worker and rejects as ABORTED", async () => {
  await withWorkerGlobal(async () => {
    const worker = new ControlledDocxWorker({ respond: false });
    const runtime = core.createDocxRuntime({
      createWorker: () => worker,
      wasmUrl: "https://unit.invalid/docx.wasm",
    });
    const controller = new AbortController();
    const pending = runtime.importDocxBuffer(new ArrayBuffer(8), {
      signal: controller.signal,
    });
    controller.abort();
    await assert.rejects(
      pending,
      (error) => error?.name === "AbortError" && error?.code === "ABORTED",
    );
    assert.equal(worker.terminated, true);
    runtime.dispose();
  });
});

test("DOCX loader cancels the older request and accepts only the latest result", async () => {
  await withWorkerGlobal(async () => {
    const workers = [];
    const runtime = core.createDocxRuntime({
      createWorker: () => {
        const index = workers.length;
        const worker = new ControlledDocxWorker({
          delay: index === 0 ? 40 : 0,
          marker: index === 0 ? "old" : "latest",
        });
        workers.push(worker);
        return worker;
      },
      wasmUrl: "https://unit.invalid/docx.wasm",
    });
    const loader = runtime.createLoader();
    const oldRequest = loader.load(new ArrayBuffer(8));
    const latestRequest = loader.load(new ArrayBuffer(8));

    await assert.rejects(
      oldRequest,
      (error) => error?.name === "AbortError" && error?.code === "ABORTED",
    );
    const latest = await latestRequest;
    assert.equal(latest.source, "worker");
    assert.equal(latest.model.nodes[0].id, "latest");
    assert.equal(workers[0].terminated, true);

    loader.dispose();
    runtime.dispose();
  });
});

test("DOCX runtime loads URL sources through shared policy and one task identity", async () => {
  await withWorkerGlobal(async () => {
    const diagnostics = [];
    const requests = [];
    const runtime = core.createDocxRuntime({
      createWorker: () => new ControlledDocxWorker({ marker: "url-source" }),
      wasmUrl: "https://assets.example/docx.wasm",
      urlPolicy: {
        enabled: true,
        baseUrl: "https://app.example/viewer/",
        allowRelativeUrl: true,
        allowedProtocols: ["https:"],
        allowedOrigins: ["https://app.example"],
        fetch: async (url, init) => {
          requests.push({ url, init });
          const bytes = new Uint8Array(arrayBufferFromFile("invoice-table.docx"));
          return {
            ok: true,
            status: 200,
            url,
            headers: { get: (name) => name.toLowerCase() === "content-length" ? String(bytes.byteLength) : null },
            arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
          };
        },
      },
      limits: { maxInputBytes: 1024 * 1024 },
      onDiagnostic: (event) => diagnostics.push(event),
    });
    const result = await runtime.loadSource({ kind: "url", url: "invoice-table.docx?token=secret" });
    assert.equal(result.model.nodes[0].id, "url-source");
    assert.equal(requests.length, 1);
    assert.equal(requests[0].init.redirect, "error");
    assert.equal(requests[0].init.credentials, "omit");
    const taskIds = new Set(diagnostics.map((event) => event.requestId));
    assert.equal(taskIds.size, 1);
    assert.match([...taskIds][0], new RegExp(`^${runtime.id}:1$`));
    assert.equal(JSON.stringify(diagnostics).includes("secret"), false);
    runtime.dispose();
  });
});

test("DOCX runtime rejects a dangerous URL before fetch", async () => {
  let requestCount = 0;
  const runtime = core.createDocxRuntime({
    urlPolicy: {
      enabled: true,
      baseUrl: "https://app.example/",
      allowRelativeUrl: true,
      allowedProtocols: ["https:"],
      allowedOrigins: ["https://app.example"],
      fetch: async () => { requestCount += 1; throw new Error("must not fetch"); },
    },
  });
  await assert.rejects(
    runtime.loadSource({ kind: "url", url: "javascript:alert(1)" }),
    (error) => error?.name === "DocxImportError" && error?.code === "SOURCE_NOT_ALLOWED",
  );
  assert.equal(requestCount, 0);
  runtime.dispose();
});

test("DOCX runtime snapshots limits instead of reading later caller mutations", async () => {
  await withWorkerGlobal(async () => {
    const limits = { maxInputBytes: 8, maxArchiveEntries: 7, maxDocxPages: 3 };
    const worker = new ControlledDocxWorker({ marker: "snapshotted-config" });
    const runtime = core.createDocxRuntime({
      limits,
      createWorker: () => worker,
      wasmUrl: "https://assets.example/docx.wasm",
    });
    limits.maxInputBytes = 1;
    limits.maxArchiveEntries = 1;
    limits.maxDocxPages = 1;
    const result = await runtime.importDocxBuffer(new ArrayBuffer(4));
    assert.equal(result.model.nodes[0].id, "snapshotted-config");
    assert.equal(Object.isFrozen(runtime.limits), true);
    assert.equal(runtime.limits.maxArchiveEntries, 7);
    assert.equal(runtime.limits.maxDocxPages, 3);
    assert.equal(worker.messages[0].limits.maxArchiveEntries, 7);
    runtime.dispose();
  });
});

test("DOCX Worker timeout is structured and terminates the blocked Worker", async () => {
  await withWorkerGlobal(async () => {
    const worker = new ControlledDocxWorker({ respond: false });
    const runtime = core.createDocxRuntime({
      createWorker: () => worker,
      wasmUrl: "https://assets.example/docx.wasm",
      limits: { maxInputBytes: 8, maxParseMs: 5 },
    });
    await assert.rejects(
      runtime.importDocxBuffer(new ArrayBuffer(4)),
      (error) => error?.code === "TIMEOUT" &&
        error?.phase === "parse" &&
        error?.limit === "maxParseMs" &&
        error?.actual > error?.allowed,
    );
    assert.equal(worker.terminated, true);
    runtime.dispose();
  });
});
