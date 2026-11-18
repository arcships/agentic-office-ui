import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
} from "./vue-test-renderer.mjs";

const { DocxEditor, useDocxEditor } = await importFromDemo("@arcships/vue-docx");
const { serializeDocx } = await importFromDemo("@arcships/docx-core");
const nativeSetTimeout = globalThis.setTimeout;

class ImmediateExportWorker {
  constructor() {
    this.listeners = new Map();
    this.terminated = false;
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
    if (!message || message.type !== "export-docx") return;
    nativeSetTimeout(async () => {
      if (this.terminated) return;
      this.emit("message", {
        id: message.id,
        type: "progress",
        phase: "serializing",
        percent: 10,
      });
      try {
        const bytes = await serializeDocx(message.model, message.basePackage);
        if (this.terminated) return;
        this.emit("message", {
          id: message.id,
          type: "progress",
          phase: "finalizing",
          percent: 90,
        });
        this.emit("message", { id: message.id, type: "success", bytes });
      } catch (error) {
        this.emit("message", {
          id: message.id,
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

async function waitUntil(predicate, message, timeoutMs = 5_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    await vue.nextTick();
    if (predicate()) return;
    await new Promise((resolve) => nativeSetTimeout(resolve, 5));
  }
  throw new Error(`Timed out waiting for ${message}`);
}

async function mountEditorController(fileName = "race.docx") {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ initialFileName: fileName });
      return () => vue.h("div", { "data-testid": "editor-harness" }, controller.status);
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

function installDownloadEnvironment(t, { clickError, documentAvailable = true } = {}) {
  const originalDocument = globalThis.document;
  const originalWorker = globalThis.Worker;
  const originalCreateObjectURL = URL.createObjectURL;
  const originalRevokeObjectURL = URL.revokeObjectURL;
  const originalSetTimeout = globalThis.setTimeout;
  const originalClearTimeout = globalThis.clearTimeout;
  const actions = [];
  const blobs = [];
  const downloads = [];
  const exportTimers = new Set();
  let nextTimerId = 0;

  if (documentAvailable) {
    globalThis.document = {
      body: {
        append: () => actions.push("append"),
      },
      createElement: (tagName) => {
        actions.push(`anchor:${tagName}`);
        return {
          style: {},
          click() {
            actions.push("click");
            downloads.push(this.download);
            if (clickError) throw clickError;
          },
          remove() {
            actions.push("remove");
          },
        };
      },
    };
  } else {
    delete globalThis.document;
  }
  globalThis.Worker = ImmediateExportWorker;

  URL.createObjectURL = (blob) => {
    blobs.push(blob);
    actions.push("create-url");
    return "blob:docx-export-test";
  };
  URL.revokeObjectURL = (url) => actions.push(`revoke:${url}`);
  globalThis.setTimeout = (callback, delay, ...args) => {
    if (delay !== 1000) return originalSetTimeout(callback, delay, ...args);
    const handle = { exportTimerId: ++nextTimerId };
    exportTimers.add(handle);
    actions.push("set-timer");
    return handle;
  };
  globalThis.clearTimeout = (handle) => {
    if (exportTimers.has(handle)) {
      exportTimers.delete(handle);
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
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
    URL.createObjectURL = originalCreateObjectURL;
    URL.revokeObjectURL = originalRevokeObjectURL;
    globalThis.setTimeout = originalSetTimeout;
    globalThis.clearTimeout = originalClearTimeout;
  };
  t.after(restore);

  return { actions, blobs, downloads, exportTimers, restore };
}

function count(actions, value) {
  return actions.filter((action) => action === value).length;
}

class HangingImportWorker {
  constructor() {
    this.listeners = new Map();
    this.terminated = false;
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

  postMessage() {}

  terminate() {
    this.terminated = true;
  }
}

test("DOCX export reports worker progress before downloading", async (t) => {
  const mounted = await mountEditorController();
  const environment = installDownloadEnvironment(t);

  mounted.controller.exportDocx();
  assert.match(mounted.controller.status, /^Exporting DOCX/);
  await waitUntil(
    () => mounted.controller.status === "Exported DOCX",
    "worker-backed DOCX export",
  );
  assert.equal(environment.downloads.length, 1);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX export refuses to run while a document import is in flight", async (t) => {
  const mounted = await mountEditorController("loading.docx");
  const environment = installDownloadEnvironment(t);
  const previousWorker = globalThis.Worker;
  globalThis.Worker = HangingImportWorker;
  let loading;

  try {
    loading = mounted.controller.importDocxFile(
      new File([new Uint8Array([1, 2, 3])], "loading.docx"),
    );
    assert.equal(mounted.controller.isImporting, true);

    mounted.controller.exportDocx();
    assert.match(mounted.controller.status, /Wait for the document to finish loading/);
    await new Promise((resolve) => nativeSetTimeout(resolve, 20));
    assert.equal(environment.downloads.length, 0);
    assert.equal(environment.blobs.length, 0);
  } finally {
    mounted.app.unmount();
    await loading?.catch(() => undefined);
    globalThis.Worker = previousWorker;
    environment.restore();
  }
  assert.deepEqual(mounted.warnings, []);
});

test("DOCX rapid exports allow only the latest task to download and clean once", async (t) => {
  const mounted = await mountEditorController();
  const { actions, exportTimers } = installDownloadEnvironment(t);

  mounted.controller.exportDocx();
  mounted.controller.exportDocx();

  await waitUntil(
    () => mounted.controller.status === "Exported DOCX",
    "latest DOCX export",
  );
  assert.equal(count(actions, "create-url"), 1);
  assert.equal(count(actions, "append"), 1);
  assert.equal(count(actions, "click"), 1);
  assert.equal(count(actions, "set-timer"), 1);
  assert.equal(exportTimers.size, 1);

  mounted.app.unmount();
  await vue.nextTick();
  assert.equal(count(actions, "remove"), 1);
  assert.equal(count(actions, "revoke:blob:docx-export-test"), 1);
  assert.equal(count(actions, "clear-timer"), 1);
  assert.equal(exportTimers.size, 0);
  assert.deepEqual(mounted.warnings, []);
});

test("DOCX model replacement and unmount invalidate pending exports", async (t) => {
  const modelChanged = await mountEditorController();
  const modelEnvironment = installDownloadEnvironment(t);
  modelChanged.controller.exportDocx();
  modelChanged.controller.appendParagraph("invalidate pending export");
  await new Promise((resolve) => nativeSetTimeout(resolve, 50));
  assert.equal(count(modelEnvironment.actions, "create-url"), 0);
  assert.equal(count(modelEnvironment.actions, "click"), 0);
  assert.equal(modelChanged.controller.status, "Export cancelled: document changed");
  modelChanged.app.unmount();

  const unmounted = await mountEditorController();
  const beforeUnmount = modelEnvironment.actions.length;
  unmounted.controller.exportDocx();
  unmounted.app.unmount();
  await new Promise((resolve) => nativeSetTimeout(resolve, 50));
  assert.equal(
    modelEnvironment.actions.slice(beforeUnmount).filter((action) => action === "click").length,
    0,
  );
  assert.deepEqual(modelChanged.warnings, []);
  assert.deepEqual(unmounted.warnings, []);
});

test("DOCX export keeps macro and template container names and MIME types", async (t) => {
  for (const [inputName, expectedName, expectedType] of [
    ["macro.docm", "macro-edited.docm", "application/vnd.ms-word.document.macroEnabled.12"],
    ["template.dotx", "template-edited.dotx", "application/vnd.openxmlformats-officedocument.wordprocessingml.template"],
    ["template.dotm", "template-edited.dotm", "application/vnd.ms-word.template.macroEnabled.12"],
    ["legacy.doc", "edited.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
  ]) {
    const mounted = await mountEditorController(inputName);
    const environment = installDownloadEnvironment(t);
    try {
      mounted.controller.exportDocx();
      await waitUntil(
        () => mounted.controller.status === "Exported DOCX",
        `${inputName} export`,
      );
      assert.deepEqual(environment.downloads, [expectedName]);
      assert.equal(environment.blobs.length, 1);
      assert.equal(environment.blobs[0].type, expectedType.toLowerCase());
      assert.deepEqual(mounted.warnings, []);
    } finally {
      mounted.app.unmount();
      environment.restore();
    }
  }
});

test("DOCX export checks document before URL creation", async (t) => {
  const mounted = await mountEditorController();
  const { actions } = installDownloadEnvironment(t, { documentAvailable: false });

  mounted.controller.exportDocx();
  await waitUntil(
    () => mounted.controller.status.startsWith("Export failed:"),
    "missing document export failure",
  );

  assert.match(mounted.controller.status, /no document available/);
  assert.equal(count(actions, "create-url"), 0);
  assert.equal(count(actions, "click"), 0);
  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("DOCX click failure removes anchor, revokes URL and clears timer once", async (t) => {
  const mounted = await mountEditorController();
  const { actions, exportTimers } = installDownloadEnvironment(t, {
    clickError: new Error("synthetic click failure"),
  });

  mounted.controller.exportDocx();
  await waitUntil(
    () => mounted.controller.status.startsWith("Export failed:"),
    "click export failure",
  );

  assert.match(mounted.controller.status, /synthetic click failure/);
  assert.equal(count(actions, "create-url"), 1);
  assert.equal(count(actions, "append"), 1);
  assert.equal(count(actions, "set-timer"), 1);
  assert.equal(count(actions, "click"), 1);
  assert.equal(count(actions, "remove"), 1);
  assert.equal(count(actions, "revoke:blob:docx-export-test"), 1);
  assert.equal(count(actions, "clear-timer"), 1);
  assert.equal(exportTimers.size, 0);

  mounted.app.unmount();
  assert.equal(count(actions, "remove"), 1, "unmount must not clean the failed task twice");
  assert.equal(count(actions, "revoke:blob:docx-export-test"), 1);
  assert.deepEqual(mounted.warnings, []);
});

test("DocxEditor clears an active file import when file becomes undefined", async (t) => {
  const originalWorker = globalThis.Worker;
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  const originalResizeObserver = globalThis.ResizeObserver;
  const originalRect = Object.getOwnPropertyDescriptor(Object.prototype, "getBoundingClientRect");

  class HangingWorker {
    static instances = [];

    constructor() {
      this.listeners = new Map();
      this.terminated = false;
      this.messages = [];
      HangingWorker.instances.push(this);
    }

    addEventListener(type, listener) {
      this.listeners.set(type, [...(this.listeners.get(type) || []), listener]);
    }

    removeEventListener(type, listener) {
      this.listeners.set(
        type,
        (this.listeners.get(type) || []).filter((candidate) => candidate !== listener),
      );
    }

    postMessage(message) {
      this.messages.push(message);
    }

    terminate() {
      this.terminated = true;
    }
  }

  globalThis.Worker = HangingWorker;
  globalThis.window = {
    addEventListener() {},
    removeEventListener() {},
  };
  globalThis.document = {
    body: { append() {} },
    createElement() {
      return {
        style: { setProperty() {} },
        appendChild() {},
        setAttribute() {},
        getContext() {
          return { measureText: (text) => ({ width: String(text).length * 8 }) };
        },
        innerHTML: "",
      };
    },
    createTextNode(text) {
      return { textContent: String(text) };
    },
  };
  globalThis.ResizeObserver = class ResizeObserver {
    observe() {}
    disconnect() {}
  };
  Object.defineProperty(Object.prototype, "getBoundingClientRect", {
    configurable: true,
    value: () => ({ height: 1056, width: 816, top: 0, left: 0, right: 816, bottom: 1056 }),
  });

  t.after(() => {
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
    if (originalWindow === undefined) delete globalThis.window;
    else globalThis.window = originalWindow;
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    if (originalResizeObserver === undefined) delete globalThis.ResizeObserver;
    else globalThis.ResizeObserver = originalResizeObserver;
    if (originalRect) Object.defineProperty(Object.prototype, "getBoundingClientRect", originalRect);
    else delete Object.prototype.getBoundingClientRect;
  });

  const file = vue.ref(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4]).buffer);
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(DocxEditor, {
        editable: false,
        file: file.value,
        showThumbnails: false,
        showToolbar: false,
      });
    },
  });
  const mounted = await mount(Harness);
  await waitUntil(
    () => HangingWorker.instances.some((worker) => worker.messages.length > 0),
    "DOCX import Worker request",
  );

  file.value = undefined;
  await waitUntil(
    () => HangingWorker.instances.every((worker) => worker.terminated),
    "DOCX import cancellation after file clear",
  );
  assert.match(textContent(findByTestId(mounted.root, "editor-status")), /Ready/);
  assert.equal(HangingWorker.instances.length, 1);
  assert.equal(HangingWorker.instances[0].terminated, true);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});
