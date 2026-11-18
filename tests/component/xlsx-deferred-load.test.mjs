import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

const { XlsxViewer, useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

const DEFERRED_BYTES = 3 * 1024 * 1024;

function createDeferredController() {
  const state = vue.reactive({
    activeSheet: null,
    activeTab: null,
    canLoadDeferred: true,
    deferredLoadFileSize: DEFERRED_BYTES,
    displayFileName: "large-workbook.xlsx",
    error: null,
    isLoadDeferred: true,
    isLoading: false,
    readOnly: false,
    tabs: [],
  });
  const calls = [];
  return {
    calls,
    state,
    controller: {
      get activeSheet() { return state.activeSheet; },
      get activeTab() { return state.activeTab; },
      get canLoadDeferred() { return state.canLoadDeferred; },
      cancelDeferredLoad() {
        calls.push("cancel");
        state.isLoadDeferred = false;
        state.deferredLoadFileSize = null;
      },
      continueDeferredLoad() {
        calls.push("continue");
      },
      get deferredLoadFileSize() { return state.deferredLoadFileSize; },
      get displayFileName() { return state.displayFileName; },
      get error() { return state.error; },
      get isLoadDeferred() { return state.isLoadDeferred; },
      get isLoading() { return state.isLoading; },
      get readOnly() { return state.readOnly; },
      get tabs() { return state.tabs; },
    },
  };
}

function viewerProps(controller, overrides = {}) {
  return {
    controller,
    showDefaultToolbar: false,
    showFormulaBar: false,
    showImages: false,
    showRibbon: false,
    locale: "zh-CN",
    ...overrides,
  };
}

test("XlsxViewer shows a deferred-load prompt with file name, size and actions", async () => {
  const { calls, controller } = createDeferredController();
  const events = [];
  const mounted = await mount(XlsxViewer, viewerProps(controller, {
    onLoadDeferred: (payload) => events.push(["load-deferred", payload]),
    onDeferredCancel: (payload) => events.push(["deferred-cancel", payload]),
    onDeferredResumed: (payload) => events.push(["deferred-resumed", payload]),
  }));
  try {
    assert.equal(
      findByTestId(mounted.root, "xlsx-viewer").props["data-state"],
      "deferred",
      "deferred loads must expose their own state instead of the empty state",
    );
    const prompt = findByTestId(mounted.root, "xlsx-deferred-load");
    assert.ok(prompt, "the component itself must render the deferred-load prompt");
    const text = textContent(prompt);
    assert.match(text, /large-workbook\.xlsx/);
    assert.match(text, /3\.0 MB/);
    assert.match(text, /继续解析/);
    assert.match(text, /取消/);
    assert.doesNotMatch(textContent(mounted.root), /打开一个 XLSX 文件开始查看/);
    assert.deepEqual(events, [
      ["load-deferred", { fileName: "large-workbook.xlsx", bytes: DEFERRED_BYTES }],
    ]);

    const continueButton = findByTestId(mounted.root, "xlsx-deferred-continue");
    assert.ok(continueButton);
    assert.equal(continueButton.props.disabled, false);
    continueButton.props.onClick();
    await vue.nextTick();
    assert.deepEqual(calls, ["continue"]);
    assert.deepEqual(events.at(-1), [
      "deferred-resumed",
      { fileName: "large-workbook.xlsx", bytes: DEFERRED_BYTES },
    ]);

    findByTestId(mounted.root, "xlsx-deferred-cancel").props.onClick();
    await vue.nextTick();
    assert.deepEqual(calls, ["continue", "cancel"]);
    assert.deepEqual(events.at(-1), [
      "deferred-cancel",
      { fileName: "large-workbook.xlsx", bytes: DEFERRED_BYTES },
    ]);
    assert.equal(findByTestId(mounted.root, "xlsx-viewer").props["data-state"], "empty");
    assert.equal(findByTestId(mounted.root, "xlsx-deferred-load"), undefined);
    assert.match(textContent(mounted.root), /打开一个 XLSX 文件开始查看/);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("XlsxViewer exposes a deferredLoad slot with resume, cancel and file metadata", async () => {
  const { calls, controller } = createDeferredController();
  const slotPayloads = [];
  const Wrapper = vue.defineComponent({
    setup() {
      return () => vue.h(XlsxViewer, viewerProps(controller), {
        deferredLoad: (slotProps) => {
          slotPayloads.push(slotProps);
          return vue.h("div", { "data-testid": "xlsx-deferred-custom" }, [
            vue.h("span", { "data-testid": "xlsx-deferred-custom-name" }, slotProps.fileName),
            vue.h("span", { "data-testid": "xlsx-deferred-custom-size" }, slotProps.sizeLabel),
            vue.h(
              "button",
              {
                "data-testid": "xlsx-deferred-custom-resume",
                onClick: () => slotProps.resume(),
              },
              "resume",
            ),
            vue.h(
              "button",
              {
                "data-testid": "xlsx-deferred-custom-cancel",
                onClick: () => slotProps.cancel(),
              },
              "cancel",
            ),
          ]);
        },
      });
    },
  });
  const mounted = await mount(Wrapper);
  try {
    assert.ok(findByTestId(mounted.root, "xlsx-deferred-custom"));
    assert.equal(findByTestId(mounted.root, "xlsx-deferred-continue"), undefined);
    const payload = slotPayloads[0];
    assert.equal(payload.fileName, "large-workbook.xlsx");
    assert.equal(payload.bytes, DEFERRED_BYTES);
    assert.equal(payload.sizeLabel, "3.0 MB");
    assert.equal(typeof payload.resume, "function");
    assert.equal(typeof payload.cancel, "function");

    findByTestId(mounted.root, "xlsx-deferred-custom-resume").props.onClick();
    findByTestId(mounted.root, "xlsx-deferred-custom-cancel").props.onClick();
    await vue.nextTick();
    assert.deepEqual(calls, ["continue", "cancel"]);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("deferred XLSX controller prompt cancels back to the empty viewer", async () => {
  const bytes = readFileSync(
    new URL("../../apps/demo/public/samples/sales-table.xlsx", import.meta.url),
  );
  const file = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const diagnostics = [];
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        deferLoadingAboveBytes: 1,
        file,
        fileName: "sales-table.xlsx",
        onDiagnostic: (event) => diagnostics.push(event),
        skipXmlParsing: true,
        useWorker: false,
      });
      return () => vue.h(XlsxViewer, viewerProps(controller));
    },
  });
  const mounted = await mount(Harness);
  try {
    await waitFor(() => controller.isLoadDeferred, "deferred XLSX prompt");
    await vue.nextTick();
    assert.equal(findByTestId(mounted.root, "xlsx-viewer").props["data-state"], "deferred");
    const prompt = findByTestId(mounted.root, "xlsx-deferred-load");
    assert.ok(prompt);
    const text = textContent(prompt);
    assert.match(text, /sales-table\.xlsx/);
    assert.match(text, /文件大小 [\d.]+ (B|KB|MB)/);
    assert.equal(controller.deferredLoadFileSize, bytes.byteLength);
    assert.ok(diagnostics.some((event) => event.type === "load-deferred"));

    findByTestId(mounted.root, "xlsx-deferred-cancel").props.onClick();
    await waitFor(
      () => controller.sourceState === "idle" && !controller.isLoadDeferred,
      "cancelled XLSX state",
    );
    await vue.nextTick();
    assert.equal(controller.activeSheet, null);
    assert.equal(controller.deferredLoadFileSize, null);
    assert.equal(controller.error, null);
    assert.equal(
      diagnostics.filter((event) => event.type === "load-cancelled").length,
      1,
    );
    assert.equal(findByTestId(mounted.root, "xlsx-deferred-load"), undefined);
    assert.equal(findByTestId(mounted.root, "xlsx-viewer").props["data-state"], "empty");
    assert.match(textContent(mounted.root), /打开一个 XLSX 文件开始查看/);

    controller.cancelDeferredLoad();
    assert.equal(
      diagnostics.filter((event) => event.type === "load-cancelled").length,
      1,
      "cancelling again must be a no-op",
    );
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});
