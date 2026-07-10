import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
} from "./vue-test-renderer.mjs";

const { DocxEditor, useDocxEditor } = await importFromDemo("@extend-ai/vue-docx");
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

async function mountEditorController() {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ initialFileName: "race.docx" });
      return () => vue.h("div", { "data-testid": "editor-harness" }, controller.status);
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

function installDownloadEnvironment(t, { clickError, documentAvailable = true } = {}) {
  const originalDocument = globalThis.document;
  const originalCreateObjectURL = URL.createObjectURL;
  const originalRevokeObjectURL = URL.revokeObjectURL;
  const originalSetTimeout = globalThis.setTimeout;
  const originalClearTimeout = globalThis.clearTimeout;
  const actions = [];
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

  URL.createObjectURL = () => {
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

  t.after(() => {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    URL.createObjectURL = originalCreateObjectURL;
    URL.revokeObjectURL = originalRevokeObjectURL;
    globalThis.setTimeout = originalSetTimeout;
    globalThis.clearTimeout = originalClearTimeout;
  });

  return { actions, exportTimers };
}

function count(actions, value) {
  return actions.filter((action) => action === value).length;
}

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
  assert.notEqual(modelChanged.controller.status, "Exported DOCX");
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
