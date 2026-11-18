import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

import { importFromDemo, mount, vue } from "./vue-test-renderer.mjs";

const core = await importFromDemo("@arcships/docx-core");
const { useDocxEditor } = await importFromDemo("@arcships/vue-docx");
const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const { strFromU8, unzipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);

const nativeSetTimeout = globalThis.setTimeout;

async function waitUntil(predicate, message, timeoutMs = 5_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    await vue.nextTick();
    if (predicate()) return;
    await new Promise((resolve) => nativeSetTimeout(resolve, 5));
  }
  throw new Error(`Timed out waiting for ${message}`);
}

async function mountController(options = {}) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({
        starterModel: core.createBlankDocumentModel(),
        ...options,
      });
      return () => vue.h("div", { "data-testid": "docx-export-bytes-harness" }, controller.status);
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

function installWorker(t, WorkerClass) {
  const originalWorker = globalThis.Worker;
  if (Array.isArray(WorkerClass.instances)) WorkerClass.instances.length = 0;
  globalThis.Worker = WorkerClass;
  t.after(() => {
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
  });
}

function installDownloadEnvironment(t) {
  const originalDocument = globalThis.document;
  const originalCreateObjectURL = URL.createObjectURL;
  const originalRevokeObjectURL = URL.revokeObjectURL;
  const originalSetTimeout = globalThis.setTimeout;
  const originalClearTimeout = globalThis.clearTimeout;
  const actions = [];
  const blobs = [];
  const downloads = [];
  const exportTimers = new Set();

  globalThis.document = {
    body: {
      append: () => actions.push("append"),
    },
    createElement(tagName) {
      assert.equal(tagName, "a");
      actions.push("anchor");
      return {
        style: {},
        download: "",
        href: "",
        click() {
          actions.push("click");
          downloads.push(this.download);
        },
        remove() {
          actions.push("remove");
        },
      };
    },
  };
  URL.createObjectURL = (blob) => {
    blobs.push(blob);
    actions.push("create-url");
    return "blob:docx-export-bytes-test";
  };
  URL.revokeObjectURL = (url) => actions.push(`revoke:${url}`);
  globalThis.setTimeout = (callback, delay, ...args) => {
    if (delay !== 1000) return originalSetTimeout(callback, delay, ...args);
    const handle = { exportTimer: true };
    exportTimers.add(handle);
    actions.push("set-timer");
    return handle;
  };
  globalThis.clearTimeout = (handle) => {
    if (exportTimers.delete(handle)) {
      actions.push("clear-timer");
      return;
    }
    return originalClearTimeout(handle);
  };

  let restored = false;
  const restore = () => {
    if (restored) return;
    restored = true;
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    URL.createObjectURL = originalCreateObjectURL;
    URL.revokeObjectURL = originalRevokeObjectURL;
    globalThis.setTimeout = originalSetTimeout;
    globalThis.clearTimeout = originalClearTimeout;
  };
  t.after(restore);
  return { actions, blobs, downloads, exportTimers, restore };
}

class RecordingExportWorker {
  static instances = [];

  constructor() {
    this.listeners = new Map();
    this.messages = [];
    this.terminated = false;
    RecordingExportWorker.instances.push(this);
  }

  addEventListener(type, listener) {
    this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
  }

  removeEventListener(type, listener) {
    this.listeners.set(
      type,
      (this.listeners.get(type) ?? []).filter((candidate) => candidate !== listener),
    );
  }

  postMessage(message) {
    const safeMessage = structuredClone(message);
    this.messages.push(safeMessage);
    nativeSetTimeout(async () => {
      if (this.terminated) return;
      this.emit("message", {
        id: safeMessage.id,
        type: "progress",
        phase: "serializing",
        percent: 10,
      });
      try {
        const bytes = await core.serializeDocx(safeMessage.model, safeMessage.basePackage);
        if (this.terminated) return;
        this.emit("message", {
          id: safeMessage.id,
          type: "progress",
          phase: "finalizing",
          percent: 90,
        });
        this.emit("message", { id: safeMessage.id, type: "success", bytes });
      } catch (error) {
        this.emit("message", {
          id: safeMessage.id,
          type: "error",
          error: { message: error instanceof Error ? error.message : String(error) },
        });
      }
    }, 0);
  }

  terminate() {
    this.terminated = true;
  }

  emit(type, data) {
    for (const listener of this.listeners.get(type) ?? []) {
      listener({ data });
    }
  }
}

class HangingExportWorker {
  static instances = [];

  constructor() {
    this.listeners = new Map();
    this.messages = [];
    this.terminated = false;
    HangingExportWorker.instances.push(this);
  }

  addEventListener(type, listener) {
    this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
  }

  removeEventListener(type, listener) {
    this.listeners.set(
      type,
      (this.listeners.get(type) ?? []).filter((candidate) => candidate !== listener),
    );
  }

  postMessage(message) {
    this.messages.push(structuredClone(message));
  }

  terminate() {
    this.terminated = true;
  }

  emit(type, data) {
    for (const listener of this.listeners.get(type) ?? []) {
      listener({ data });
    }
  }
}

class FailingExportWorker extends HangingExportWorker {
  postMessage(message) {
    const safeMessage = structuredClone(message);
    this.messages.push(safeMessage);
    queueMicrotask(() => {
      this.emit("message", {
        id: safeMessage.id,
        type: "error",
        error: { code: "EXPORT_FAILED", message: "synthetic serializer failure" },
      });
    });
  }
}

function docxParts(bytes) {
  return unzipSync(new Uint8Array(bytes));
}

function documentXml(bytes) {
  return strFromU8(docxParts(bytes)["word/document.xml"]);
}

test("getDocxBytes returns a Word-readable DOCX without downloading", async (t) => {
  installWorker(t, RecordingExportWorker);
  const mounted = await mountController();
  const environment = installDownloadEnvironment(t);

  try {
    mounted.controller.appendParagraph("Programmatic export paragraph");
    const bytes = await mounted.controller.getDocxBytes();

    assert.ok(bytes instanceof ArrayBuffer, "getDocxBytes must resolve an ArrayBuffer");
    assert.deepEqual(
      [...new Uint8Array(bytes).slice(0, 4)],
      [0x50, 0x4b, 0x03, 0x04],
      "the bytes must be a ZIP/OOXML package",
    );
    const archive = docxParts(bytes);
    assert.ok(archive["[Content_Types].xml"], "the package must declare its content types");
    assert.match(documentXml(bytes), /Programmatic export paragraph/);

    const reopened = await core.wasmBuildDocModelFromBytes(bytes);
    assert.match(
      JSON.stringify(reopened.model.nodes),
      /Programmatic export paragraph/,
      "the returned bytes must reopen as a document model",
    );

    assert.equal(environment.actions.length, 0, "getDocxBytes must not create a download");
    assert.equal(mounted.controller.status, "Ready", "a programmatic read must not report an export");
    assert.equal(mounted.controller.isDirty, true, "a programmatic read must not mark the document saved");

    assert.equal(RecordingExportWorker.instances.length, 1);
    const worker = RecordingExportWorker.instances[0];
    assert.equal(worker.messages.length, 1);
    assert.equal(worker.messages[0].type, "export-docx");
    assert.equal(worker.terminated, true, "the export Worker must be released");
  } finally {
    mounted.app.unmount();
    environment.restore();
  }
});

test("getDocxBlob returns the export MIME type without downloading", async (t) => {
  installWorker(t, RecordingExportWorker);
  const mounted = await mountController({ initialFileName: "template.dotx" });
  const environment = installDownloadEnvironment(t);

  try {
    mounted.controller.appendParagraph("Blob export paragraph");
    const blob = await mounted.controller.getDocxBlob();

    assert.ok(blob instanceof Blob);
    assert.equal(
      blob.type,
      "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
    );
    const bytes = await blob.arrayBuffer();
    assert.match(documentXml(bytes), /Blob export paragraph/);
    assert.equal(environment.actions.length, 0, "getDocxBlob must not create a download");
  } finally {
    mounted.app.unmount();
    environment.restore();
  }
});

test("getDocxBytes applies the pending export model transformer", async (t) => {
  installWorker(t, RecordingExportWorker);
  const mounted = await mountController();

  try {
    mounted.controller.appendParagraph("original paragraph");
    mounted.controller.registerPendingExportModelTransformer((model) =>
      core.updateParagraphText(model, 0, "transformed for export")
    );

    const bytes = await mounted.controller.getDocxBytes();
    assert.match(documentXml(bytes), /transformed for export/);
    assert.match(
      JSON.stringify(mounted.controller.model.nodes[0]),
      /transformed for export/,
      "the transformed model must become the editor model",
    );
  } finally {
    mounted.app.unmount();
  }
});

test("getDocxBytes posts a clone-safe base package", async (t) => {
  const basePackage = core.createMinimalDocxPackage();
  const importedModel = await core.buildDocModel(basePackage);
  const originalWorker = globalThis.Worker;

  class SuccessfulImportWorker {
    constructor() {
      this.listeners = new Map();
    }
    addEventListener(type, listener) {
      this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
    }
    removeEventListener() {}
    postMessage(request) {
      queueMicrotask(() => {
        for (const listener of this.listeners.get("message") ?? []) {
          listener({
            data: {
              id: request.id,
              type: "success",
              package: basePackage,
              model: importedModel,
              timings: { totalMs: 1, parseMs: 1, buildModelMs: 0 },
            },
          });
        }
      });
    }
    terminate() {}
  }

  globalThis.Worker = SuccessfulImportWorker;
  const mounted = await mountController();
  try {
    await mounted.controller.importDocxFile(
      new File([new Uint8Array([1, 2, 3])], "package-export.docx"),
    );
    assert.equal(mounted.controller.fileName, "package-export.docx");

    if (Array.isArray(RecordingExportWorker.instances)) {
      RecordingExportWorker.instances.length = 0;
    }
    globalThis.Worker = RecordingExportWorker;
    const bytes = await mounted.controller.getDocxBytes();
    assert.deepEqual(
      [...new Uint8Array(bytes).slice(0, 4)],
      [0x50, 0x4b, 0x03, 0x04],
    );
    assert.equal(RecordingExportWorker.instances.length, 1);
    assert.ok(
      RecordingExportWorker.instances[0].messages[0].basePackage.parts instanceof Map,
      "the Worker message must contain a plain Map, not a reactive proxy",
    );
  } finally {
    mounted.app.unmount();
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
  }
});

test("exportDocx still downloads the same serialization getDocxBytes returns", async (t) => {
  installWorker(t, RecordingExportWorker);
  const mounted = await mountController({ initialFileName: "shared.docx" });
  const environment = installDownloadEnvironment(t);

  try {
    mounted.controller.appendParagraph("Shared serialization path");
    const bytes = await mounted.controller.getDocxBytes();

    mounted.controller.exportDocx();
    await waitUntil(
      () => mounted.controller.status === "Exported DOCX",
      "worker-backed DOCX export",
    );

    assert.deepEqual(environment.downloads, ["shared-edited.docx"]);
    assert.equal(environment.blobs.length, 1);
    const downloaded = await environment.blobs[0].arrayBuffer();
    assert.equal(
      documentXml(downloaded),
      documentXml(bytes),
      "download and programmatic read must share one serialization path",
    );
    assert.equal(mounted.controller.isDirty, false, "an export still marks the document saved");
  } finally {
    mounted.app.unmount();
    environment.restore();
  }
});

test("getDocxBytes rejects with ABORTED when the document changes", async (t) => {
  installWorker(t, HangingExportWorker);
  const mounted = await mountController();

  try {
    const pending = mounted.controller.getDocxBytes();
    await waitUntil(
      () => HangingExportWorker.instances[0]?.messages.length > 0,
      "the export Worker request",
    );

    mounted.controller.appendParagraph("changed during export");

    await assert.rejects(
      pending,
      (error) => error?.name === "AbortError" && error?.code === "ABORTED",
    );
    assert.equal(HangingExportWorker.instances[0].terminated, true);
    assert.equal(mounted.controller.status, "Export cancelled: document changed");
  } finally {
    mounted.app.unmount();
  }
});

test("getDocxBytes rejects with ABORTED when the editor is disposed", async (t) => {
  installWorker(t, HangingExportWorker);
  const mounted = await mountController();
  const pending = mounted.controller.getDocxBytes();
  await waitUntil(
    () => HangingExportWorker.instances[0]?.messages.length > 0,
    "the export Worker request",
  );

  mounted.app.unmount();

  await assert.rejects(
    pending,
    (error) => error?.name === "AbortError" && error?.code === "ABORTED",
  );
  assert.equal(HangingExportWorker.instances[0].terminated, true);
});

test("getDocxBytes propagates structured export failures", async (t) => {
  installWorker(t, FailingExportWorker);
  const mounted = await mountController();

  try {
    await assert.rejects(
      mounted.controller.getDocxBytes(),
      (error) => error?.name === "DocxExportError" &&
        error?.code === "EXPORT_FAILED" &&
        /synthetic serializer failure/.test(error?.message),
    );
  } finally {
    mounted.app.unmount();
  }
});

test("getDocxBytes refuses to export an empty required form field", async (t) => {
  installWorker(t, HangingExportWorker);
  const starterModel = core.createBlankDocumentModel();
  starterModel.nodes[0].children = [
    {
      type: "form-field",
      fieldType: "text",
      required: true,
      widget: { name: "Employee" },
    },
  ];
  const mounted = await mountController({ starterModel });

  try {
    await assert.rejects(
      mounted.controller.getDocxBytes(),
      (error) => error?.name === "DocxExportError" &&
        error?.code === "EXPORT_FAILED" &&
        /required form field "Employee" is empty/.test(error?.message),
    );
    assert.equal(HangingExportWorker.instances.length, 0, "validation must stop before serialization");

    mounted.controller.exportDocx();
    await waitUntil(
      () => /required form field "Employee" is empty/.test(mounted.controller.status),
      "the required-field export status",
    );
    assert.equal(HangingExportWorker.instances.length, 0);
  } finally {
    mounted.app.unmount();
  }
});

test("getDocxBytes refuses to run while a document import is in flight", async (t) => {
  installWorker(t, HangingExportWorker);
  const mounted = await mountController();
  const previousWorker = globalThis.Worker;
  class HangingImportWorker {
    constructor() {
      this.listeners = new Map();
    }
    addEventListener(type, listener) {
      this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
    }
    removeEventListener() {}
    postMessage() {}
    terminate() {}
  }

  globalThis.Worker = HangingImportWorker;
  let loading;
  try {
    loading = mounted.controller.importDocxFile(
      new File([new Uint8Array([1, 2, 3])], "loading.docx"),
    );
    assert.equal(mounted.controller.isImporting, true);

    await assert.rejects(
      mounted.controller.getDocxBytes(),
      (error) => error?.name === "DocxExportError" &&
        error?.code === "EXPORT_FAILED" &&
        /Wait for the document to finish loading/.test(error?.message),
    );
    assert.equal(HangingExportWorker.instances.length, 0);
  } finally {
    mounted.app.unmount();
    await loading?.catch(() => undefined);
    globalThis.Worker = previousWorker;
  }
});

test("getDocxBytes rejects with WORKER_UNAVAILABLE in a UI host without a Worker", async (t) => {
  const originalWorker = globalThis.Worker;
  const originalDocument = globalThis.document;
  delete globalThis.Worker;
  globalThis.document = {};
  t.after(() => {
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
  });

  const mounted = await mountController();
  try {
    await assert.rejects(
      mounted.controller.getDocxBytes(),
      (error) => error?.name === "DocxExportError" && error?.code === "WORKER_UNAVAILABLE",
    );
  } finally {
    mounted.app.unmount();
  }
});

test("getDocxBytes falls back to the inline serializer in a headless host", async (t) => {
  const originalWorker = globalThis.Worker;
  const originalDocument = globalThis.document;
  delete globalThis.Worker;
  delete globalThis.document;
  t.after(() => {
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
  });

  const mounted = await mountController({ initialFileName: "headless.docx" });

  try {
    mounted.controller.appendParagraph("Headless main-thread paragraph");
    const bytes = await mounted.controller.getDocxBytes();

    assert.ok(bytes instanceof ArrayBuffer);
    assert.match(documentXml(bytes), /Headless main-thread paragraph/);

    const reopened = await core.wasmBuildDocModelFromBytes(bytes);
    assert.match(JSON.stringify(reopened.model.nodes), /Headless main-thread paragraph/);

    const blob = await mounted.controller.getDocxBlob();
    assert.equal(blob.type, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    assert.match(documentXml(await blob.arrayBuffer()), /Headless main-thread paragraph/);
  } finally {
    mounted.app.unmount();
  }
});
