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
  walk,
} from "./vue-test-renderer.mjs";

const { DocxViewer } = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel, wasmBuildDocModelFromBytes } = await importFromDemo("@arcships/docx-core");
const { XlsxViewer, useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");
const { PdfViewer } = await importFromDemo("@arcships/vue-pdf");
const { FileUpload } = await importFromDemo("@arcships/vue-ui");

// Vue's native v-model directive checks these browser constructors during an
// update. The custom component renderer intentionally has no browser DOM.
if (globalThis.Document === undefined) globalThis.Document = class Document {};
if (globalThis.ShadowRoot === undefined) globalThis.ShadowRoot = class ShadowRoot {};

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((nextResolve, nextReject) => {
    resolve = nextResolve;
    reject = nextReject;
  });
  return { promise, reject, resolve };
}

function minimalDocxResult() {
  return {
    model: createBlankDocumentModel(),
    package: {},
    source: "worker",
  };
}

function xlsxController(overrides = {}) {
  return {
    activeSheet: null,
    error: null,
    isLoading: false,
    readOnly: false,
    redo() {},
    tabs: [],
    undo() {},
    ...overrides,
  };
}

function createPdfRuntime({
  pageCount = 3,
  onOpenDocument,
  onRenderPage,
  onRenderThumbnail,
  onSearch,
} = {}) {
  const calls = {
    closeDocument: [],
    dispose: [],
    openDocument: [],
    renderPage: [],
    renderThumbnail: [],
    search: [],
  };
  const document = {
    id: "component-pdf-document",
    pageCount,
    pages: Array.from({ length: pageCount }, (_, index) => ({
      index,
      width: 600 + index,
      height: 800 + index,
      rotation: 0,
    })),
  };
  const runtime = {
    runtimeId: "component-pdf-runtime",
    async openDocument(bytes, options = {}) {
      calls.openDocument.push({ bytes, options });
      return onOpenDocument ? onOpenDocument(bytes, options, document) : document;
    },
    async renderPage(currentDocument, pageIndex, options) {
      calls.renderPage.push({ document: currentDocument, options, pageIndex });
      return onRenderPage
        ? onRenderPage(currentDocument, pageIndex, options)
        : new Blob([`page:${pageIndex}`], { type: "image/png" });
    },
    async renderThumbnail(currentDocument, pageIndex, options = {}) {
      calls.renderThumbnail.push({ document: currentDocument, options, pageIndex });
      return onRenderThumbnail
        ? onRenderThumbnail(currentDocument, pageIndex, options)
        : new Blob([`thumbnail:${pageIndex}`], { type: "image/png" });
    },
    async search(currentDocument, query, options = {}) {
      calls.search.push({ document: currentDocument, options, query });
      return onSearch ? onSearch(currentDocument, query, options) : [];
    },
    async closeDocument(currentDocument) {
      calls.closeDocument.push(currentDocument);
    },
    async dispose() {
      calls.dispose.push(true);
    },
  };
  return { calls, document, runtime };
}

test("DocxViewer exposes empty, loading, success and error states with public events", async () => {
  const empty = await mount(DocxViewer, { emptyState: "DOCX empty" });
  assert.equal(findByTestId(empty.root, "docx-viewer").props["data-state"], "idle");
  assert.match(textContent(empty.root), /DOCX empty/);
  assert.deepEqual(empty.warnings, []);
  empty.app.unmount();

  const pending = deferred();
  const events = [];
  const success = await mount(DocxViewer, {
    file: new Uint8Array([1, 2, 3]).buffer,
    runtime: {
      createLoader: () => ({ load: () => pending.promise, cancel() {}, dispose() {} }),
    },
    onLoadStart: () => events.push("start"),
    onLoadSuccess: (result) => events.push(["success", result.source]),
  });
  assert.equal(findByTestId(success.root, "docx-viewer").props["data-state"], "loading");
  pending.resolve(minimalDocxResult());
  await waitFor(
    () => findByTestId(success.root, "docx-viewer")?.props["data-state"] === "ready",
    "DOCX ready state",
  );
  assert.deepEqual(events, ["start", ["success", "worker"]]);
  assert.ok(findByTestId(success.root, "docx-page"));
  assert.deepEqual(success.warnings, []);
  success.app.unmount();

  const failure = new Error("broken DOCX");
  failure.code = "PARSE_FAILED";
  const errors = [];
  const failed = await mount(DocxViewer, {
    file: new Uint8Array([4]).buffer,
    runtime: {
      createLoader: () => ({ load: async () => { throw failure; }, cancel() {}, dispose() {} }),
    },
    onLoadError: (error) => errors.push(error),
  });
  await waitFor(
    () => findByTestId(failed.root, "docx-viewer")?.props["data-state"] === "error",
    "DOCX error state",
  );
  assert.equal(findByTestId(failed.root, "load-error").props["data-error-code"], "PARSE_FAILED");
  assert.equal(errors[0], failure);
  assert.deepEqual(failed.warnings, []);
  failed.app.unmount();
});

test("DocxViewer displays model-derived tracked changes and comments without an editor controller", async () => {
  const file = readFileSync(new URL("../../apps/demo/public/samples/review-comments.docx", import.meta.url));
  const bytes = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength);
  const { model } = await wasmBuildDocModelFromBytes(bytes);
  const mounted = await mount(DocxViewer, { model, fileName: "review-comments.docx" });
  await vue.nextTick();

  const viewer = findByTestId(mounted.root, "docx-viewer");
  const trackedToggle = findByTestId(mounted.root, "docx-show-tracked-changes");
  const commentToggle = findByTestId(mounted.root, "docx-show-comments");
  assert.equal(viewer.props["data-show-tracked-changes"], "true");
  assert.equal(viewer.props["data-show-comments"], "true");
  assert.equal(trackedToggle.props.disabled, false);
  assert.equal(commentToggle.props.disabled, false);
  assert.equal(trackedToggle.props["aria-pressed"], true);
  assert.equal(commentToggle.props["aria-pressed"], true);
  assert.match(textContent(mounted.root), /Alex Reviewer/);
  assert.match(textContent(mounted.root), /Morgan Editor/);
  assert.match(textContent(mounted.root), /Priya Reviewer/);
  assert.match(textContent(mounted.root), /Please confirm this report includes the service-level summary/);

  commentToggle.props.onClick();
  await vue.nextTick();
  assert.equal(findByTestId(mounted.root, "docx-viewer").props["data-show-comments"], "false");
  assert.doesNotMatch(textContent(mounted.root), /Please confirm this report includes the service-level summary/);
  assert.match(textContent(mounted.root), /Alex Reviewer/);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DocxViewer aborts an active runtime load when unmounted", async () => {
  let observedSignal;
  let errorEvents = 0;
  const runtime = {
    createLoader() {
      const controller = new AbortController();
      return {
        load() {
          observedSignal = controller.signal;
          return new Promise((_resolve, reject) => {
            controller.signal.addEventListener("abort", () => {
              reject(new DOMException("Aborted", "AbortError"));
            }, { once: true });
          });
        },
        cancel() { controller.abort(); },
        dispose() { controller.abort(); },
      };
    },
  };
  const mounted = await mount(DocxViewer, {
    file: new Uint8Array([1]).buffer,
    runtime,
    onLoadError: () => { errorEvents += 1; },
  });
  assert.equal(observedSignal.aborted, false);
  mounted.app.unmount();
  await vue.nextTick();
  assert.equal(observedSignal.aborted, true);
  assert.equal(errorEvents, 0, "unmount cancellation must not emit a stale load error");
});

test("XlsxViewer renders loading, error and empty states and handles undo/redo shortcuts", async () => {
  for (const [controller, expected] of [
    [xlsxController({ isLoading: true }), "加载中"],
    [xlsxController({ error: new Error("XLSX failed") }), "XLSX failed"],
    [xlsxController(), "打开一个 XLSX 文件开始查看"],
  ]) {
    const current = await mount(XlsxViewer, {
      controller,
      height: "320px",
      showDefaultToolbar: false,
      showFormulaBar: false,
      showImages: false,
      showRibbon: false,
    });
    assert.match(textContent(current.root), new RegExp(expected));
    assert.deepEqual(current.warnings, []);
    current.app.unmount();
  }

  const calls = [];
  const keyboard = await mount(XlsxViewer, {
    controller: xlsxController({
      undo: () => calls.push("undo"),
      redo: () => calls.push("redo"),
    }),
    rounded: false,
    showDefaultToolbar: false,
    showFormulaBar: false,
    showRibbon: false,
  });
  const viewer = walk(keyboard.root).find((node) => node.props?.class === "xlsx-viewer");
  assert.equal(viewer.props.style.borderRadius, "0px", "rounded prop reaches the public root style");
  let prevented = 0;
  viewer.props.onKeydown({ ctrlKey: true, metaKey: false, key: "z", preventDefault: () => { prevented += 1; } });
  viewer.props.onKeydown({ ctrlKey: true, metaKey: false, key: "y", preventDefault: () => { prevented += 1; } });
  assert.deepEqual(calls, ["undo", "redo"]);
  assert.equal(prevented, 2);
  assert.deepEqual(keyboard.warnings, []);
  keyboard.app.unmount();
});

test("XlsxViewer renders and selects a resolved chartsheet chart", async () => {
  const selected = [];
  const chart = {
    anchor: {
      kind: "absolute",
      positionEmu: { x: 0, y: 0 },
      sizeEmu: { cx: 6096000, cy: 3429000 },
    },
    axes: [],
    chartType: "ColumnClustered",
    editable: false,
    id: "chartsheet-0-chart-0",
    series: [],
    sheetIndex: -1,
    title: "Revenue by Quarter",
    workbookSheetIndex: -1,
    zIndex: 200,
  };
  const mounted = await mount(XlsxViewer, {
    controller: xlsxController({
      activeTab: { chartsheetIndex: 0, id: "chartsheet-0", index: 1, kind: "chartsheet", name: "Revenue Chart" },
      charts: [chart],
      isChartsLoading: false,
      selectChart: (id) => selected.push(id),
      selectedChartElement: null,
      selectedChartId: null,
    }),
    showDefaultToolbar: false,
    showFormulaBar: false,
    showImages: false,
    showRibbon: false,
  });
  assert.equal(findByTestId(mounted.root, "xlsx-viewer").props["data-state"], "ready");
  assert.ok(findByTestId(mounted.root, "xlsx-chartsheet"));
  const renderedChart = findByTestId(mounted.root, "xlsx-chartsheet-chart");
  assert.ok(renderedChart, "已解析的 chartsheet 图表必须进入正式渲染容器");
  renderedChart.props.onPointerdown();
  assert.deepEqual(selected, [chart.id]);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("useXlsxViewerController owns and terminates its Worker on unmount", async (t) => {
  const originalWorker = globalThis.Worker;
  globalThis.Worker = class Worker {};
  t.after(() => {
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
  });

  class FakeWorker {
    listeners = new Map();
    messages = [];
    terminated = false;

    addEventListener(type, listener) {
      this.listeners.set(type, listener);
    }

    removeEventListener(type, listener) {
      if (this.listeners.get(type) === listener) this.listeners.delete(type);
    }

    postMessage(message) {
      this.messages.push(message);
      queueMicrotask(() => {
        this.listeners.get("message")?.({
          data: {
            id: message.id,
            success: true,
            result: {
              chartsByWorkbookSheetIndex: [],
              chartsheets: [],
              sheets: [],
              tablesByWorkbookSheetIndex: [],
              tabs: [],
            },
          },
        });
      });
    }

    terminate() {
      this.terminated = true;
    }
  }

  const bytes = readFileSync(new URL("../../apps/demo/public/samples/sales-table.xlsx", import.meta.url));
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const worker = new FakeWorker();
  const diagnostics = [];
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        createWorker: () => worker,
        file: buffer,
        onDiagnostic: (event) => diagnostics.push(event),
        readOnly: true,
        skipXmlParsing: true,
      });
      return () => vue.h("div", { "data-state": controller.sourceState }, controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(() => controller.sourceState === "ready", "XLSX Worker-backed ready state");
  assert.equal(worker.messages[0]?.type, "load");
  assert.equal(controller.isWorkerBacked, true);
  assert.deepEqual(diagnostics.map((event) => event.type), ["load-start", "load-success"]);
  mounted.app.unmount();
  assert.equal(worker.terminated, true);
  assert.deepEqual(mounted.warnings, []);
});

test("PdfViewer becomes ready only after the injected runtime opens and renders the PDF", async () => {
  const empty = await mount(PdfViewer, { showToolbar: false });
  assert.equal(findByTestId(empty.root, "pdf-status").props["data-state"], "idle");
  assert.match(textContent(empty.root), /No PDF loaded/);
  empty.app.unmount();

  const pdfBytes = new TextEncoder().encode("%PDF-1.4\n1 0 obj\n<< /Type /Page >>\nendobj\n%%EOF");
  const openGate = deferred();
  const firstRenderGate = deferred();
  let isFirstRender = true;
  const { calls, document, runtime } = createPdfRuntime({
    onOpenDocument: () => openGate.promise,
    onRenderPage: (_document, pageIndex) => {
      if (isFirstRender) {
        isFirstRender = false;
        return firstRenderGate.promise;
      }
      return new Blob([`page:${pageIndex}`], { type: "image/png" });
    },
    onSearch: () => [{
      pageIndex: 2,
      before: "before ",
      match: "needle",
      after: " after",
      rects: [],
    }],
  });
  const diagnostics = [];
  const pages = [];
  const ready = await mount(PdfViewer, {
    source: { kind: "blob", blob: new Blob([pdfBytes], { type: "application/pdf" }), fileName: "safe.pdf" },
    runtime,
    onActivePageChange: (page) => pages.push(page),
    onDiagnostic: (event) => diagnostics.push(event),
    onDocumentLoadSuccess: (count) => pages.push(`count:${count}`),
  });
  await waitFor(() => calls.openDocument.length === 1, "PDF runtime open");
  assert.equal(findByTestId(ready.root, "pdf-status").props["data-state"], "loading");
  assert.equal(calls.renderPage.length, 0, "render waits for openDocument");

  openGate.resolve(document);
  await waitFor(() => calls.renderPage.length === 1, "first PDF page render");
  assert.equal(findByTestId(ready.root, "pdf-status").props["data-state"], "loading");
  assert.equal(findByTestId(ready.root, "pdf-page-image"), undefined, "ready requires rendered pixels");

  firstRenderGate.resolve(new Blob(["first-page"], { type: "image/png" }));
  await waitFor(
    () => findByTestId(ready.root, "pdf-status")?.props["data-state"] === "ready",
    "PDF ready state",
  );
  assert.deepEqual(diagnostics.slice(0, 4).map((event) => event.type), [
    "load-start",
    "render-start",
    "render-success",
    "load-success",
  ]);
  assert.deepEqual(pages, ["count:3", 1]);
  assert.equal(findByTestId(ready.root, "pdf-document").props["data-page-count"], 3);
  assert.equal(findByTestId(ready.root, "pdf-page-count").children[0].text, "3");
  assert.equal(findByTestId(ready.root, "pdf-page").props["data-render-state"], "rendered");
  assert.ok(findByTestId(ready.root, "pdf-page-image"));

  findByTestId(ready.root, "pdf-next-page").props.onClick();
  await waitFor(
    () => calls.renderPage.length >= 2 &&
      calls.renderPage.at(-1).pageIndex === 1 &&
      findByTestId(ready.root, "pdf-page")?.props["data-render-state"] === "rendered",
    "second PDF page render",
  );
  assert.equal(findByTestId(ready.root, "pdf-document").props["data-active-page"], 2);

  findByTestId(ready.root, "pdf-zoom-in").props.onClick();
  await waitFor(
    () => calls.renderPage.at(-1)?.options.scale === 1.25 &&
      findByTestId(ready.root, "pdf-page")?.props["data-render-state"] === "rendered",
    "zoomed PDF page render",
  );
  assert.equal(findByTestId(ready.root, "pdf-document").props["data-zoom"], 1.25);

  findByTestId(ready.root, "pdf-rotate").props.onClick();
  await waitFor(
    () => calls.renderPage.at(-1)?.options.rotation === 90 &&
      findByTestId(ready.root, "pdf-page")?.props["data-render-state"] === "rendered",
    "rotated PDF page render",
  );
  assert.equal(findByTestId(ready.root, "pdf-document").props["data-rotation"], 90);

  const searchInput = findByTestId(ready.root, "pdf-search-input");
  searchInput.value = "needle";
  searchInput.dispatchEvent({ type: "input", target: searchInput });
  await vue.nextTick();
  const searchForm = walk(ready.root).find((node) => node.type === "form" && node.props.role === "search");
  await searchForm.props.onSubmit({ preventDefault() {} });
  await waitFor(
    () => calls.search.length === 1 &&
      calls.renderPage.at(-1)?.pageIndex === 2 &&
      findByTestId(ready.root, "pdf-page")?.props["data-render-state"] === "rendered",
    "PDF search result navigation",
  );
  assert.equal(calls.search[0].query, "needle");
  assert.equal(calls.search[0].document.id, document.id);
  assert.equal(findByTestId(ready.root, "pdf-document").props["data-active-page"], 3);
  assert.match(textContent(findByTestId(ready.root, "pdf-search-result")), /第 1 \/ 1 项，第 3 页/);
  assert.deepEqual(pages, ["count:3", 1, 2, 3]);
  assert.deepEqual(ready.warnings, []);
  ready.app.unmount();

  await waitFor(() => calls.closeDocument.length === 1, "external PDF document close");
  assert.equal(calls.closeDocument[0].id, document.id);
  assert.equal(calls.dispose.length, 0, "the viewer must not dispose a caller-owned runtime");
});

test("PdfViewer exposes source and renderer failures without reporting a false success", async () => {
  const invalidRuntime = createPdfRuntime();
  const errors = [];
  const invalid = await mount(PdfViewer, {
    source: { kind: "blob", blob: new Blob(["not pdf"], { type: "text/plain" }) },
    runtime: invalidRuntime.runtime,
    onDocumentLoadError: (error) => errors.push(error),
  });
  await waitFor(
    () => findByTestId(invalid.root, "pdf-status")?.props["data-state"] === "error",
    "PDF error state",
  );
  assert.equal(findByTestId(invalid.root, "load-error").props["data-error-code"], "INVALID_PDF");
  assert.equal(errors[0].code, "INVALID_PDF");
  assert.equal(invalidRuntime.calls.openDocument.length, 0, "invalid bytes never reach the renderer");
  assert.deepEqual(invalid.warnings, []);
  invalid.app.unmount();

  const rendererFailure = new Error("renderer rejected the document");
  const failedRuntime = createPdfRuntime({
    onRenderPage: async () => { throw rendererFailure; },
  });
  const rendererErrors = [];
  let successEvents = 0;
  const failed = await mount(PdfViewer, {
    source: {
      kind: "blob",
      blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }),
    },
    runtime: failedRuntime.runtime,
    onDocumentLoadError: (error) => rendererErrors.push(error),
    onDocumentLoadSuccess: () => { successEvents += 1; },
  });
  await waitFor(
    () => findByTestId(failed.root, "pdf-status")?.props["data-state"] === "error",
    "PDF renderer error state",
  );
  assert.equal(rendererErrors.length, 1);
  assert.equal(rendererErrors[0].code, "INVALID_PDF");
  assert.match(rendererErrors[0].message, /无法解析或渲染/);
  assert.equal(successEvents, 0);
  assert.equal(findByTestId(failed.root, "pdf-page-image"), undefined);
  await waitFor(() => failedRuntime.calls.closeDocument.length === 1, "failed PDF document close");
  failed.app.unmount();
  assert.equal(failedRuntime.calls.dispose.length, 0);
});

test("PdfViewer rejects an oversized PDF before reading bytes or opening a runtime", async () => {
  let sourceReads = 0;
  const oversizedBlob = {
    size: 9,
    type: "application/pdf",
    async arrayBuffer() {
      sourceReads += 1;
      throw new Error("oversized source must not be read");
    },
  };
  const fake = createPdfRuntime();
  const diagnostics = [];
  const errors = [];
  const mounted = await mount(PdfViewer, {
    maxFileSize: 8,
    source: { kind: "blob", blob: oversizedBlob },
    runtime: fake.runtime,
    onDiagnostic: (event) => diagnostics.push(event),
    onDocumentLoadError: (error) => errors.push(error),
  });
  await waitFor(
    () => findByTestId(mounted.root, "pdf-status")?.props["data-state"] === "error",
    "oversized PDF error state",
  );
  assert.equal(findByTestId(mounted.root, "load-error").props["data-error-code"], "PDF_TOO_LARGE");
  assert.match(textContent(findByTestId(mounted.root, "load-error")), /9.*8/);
  assert.equal(errors.length, 1);
  assert.equal(errors[0].code, "PDF_TOO_LARGE");
  assert.equal(errors[0].actual, 9);
  assert.equal(errors[0].allowed, 8);
  assert.equal(sourceReads, 0, "Blob size is checked before arrayBuffer");
  assert.equal(fake.calls.openDocument.length, 0, "oversized bytes never reach the PDF runtime");
  assert.equal(fake.calls.renderPage.length, 0);
  assert.equal(fake.calls.renderThumbnail.length, 0);
  assert.equal(diagnostics.at(-1).type, "load-error");
  assert.equal(diagnostics.at(-1).error.code, "PDF_TOO_LARGE");
  mounted.app.unmount();
  assert.equal(fake.calls.dispose.length, 0);
});

test("FileUpload reports accepted and rejected files through public events", async () => {
  const acceptedEvents = [];
  const rejectedEvents = [];
  const mounted = await mount(FileUpload, {
    accept: ".docx",
    maxSize: 8,
    multiple: true,
    onFilesAccepted: (files) => acceptedEvents.push(files),
    onFilesRejected: (rejections) => rejectedEvents.push(rejections),
  });
  const input = walk(mounted.root).find((node) => node.type === "input" && node.props.type === "file");
  assert.ok(input);
  const accepted = { name: "ok.docx", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", size: 4 };
  const wrongType = { name: "wrong.xlsx", type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", size: 4 };
  const tooLarge = { name: "large.docx", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", size: 12 };
  const target = { files: [accepted, wrongType, tooLarge], value: "selected" };
  input.props.onChange({ target });
  await vue.nextTick();

  assert.deepEqual(acceptedEvents, [[accepted]]);
  assert.deepEqual(
    rejectedEvents[0].map(({ code, file }) => ({ code, name: file.name })),
    [
      { code: "FILE_TYPE_NOT_ACCEPTED", name: "wrong.xlsx" },
      { code: "FILE_TOO_LARGE", name: "large.docx" },
    ],
  );
  assert.match(textContent(mounted.root), /wrong\.xlsx.*not an accepted file type/);
  assert.match(textContent(mounted.root), /large\.docx.*size limit/);
  assert.equal(target.value, "");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfViewer aborts URL loading when unmounted", async () => {
  let observedSignal;
  let errorEvents = 0;
  const fake = createPdfRuntime();
  const fetch = (_url, init) => {
    observedSignal = init.signal;
    return new Promise((_resolve, reject) => {
      init.signal.addEventListener("abort", () => reject(new DOMException("Aborted", "AbortError")), { once: true });
    });
  };
  const mounted = await mount(PdfViewer, {
    source: { kind: "url", url: "https://files.example/report.pdf" },
    runtime: fake.runtime,
    urlPolicy: {
      allowedOrigins: ["https://files.example"],
      allowedProtocols: ["https:"],
      baseUrl: "https://files.example/",
      fetch,
    },
    onDocumentLoadError: () => { errorEvents += 1; },
  });
  assert.equal(observedSignal.aborted, false);
  mounted.app.unmount();
  await vue.nextTick();
  assert.equal(observedSignal.aborted, true);
  assert.equal(errorEvents, 0);
  assert.equal(fake.calls.openDocument.length, 0);
  assert.equal(fake.calls.dispose.length, 0);
});

test("PdfViewer aborts an active runtime render and suppresses stale errors on unmount", async () => {
  let observedSignal;
  let errorEvents = 0;
  const fake = createPdfRuntime({
    onRenderPage: (_document, _pageIndex, options) => {
      observedSignal = options.signal;
      return new Promise((_resolve, reject) => {
        options.signal.addEventListener(
          "abort",
          () => reject(new DOMException("Aborted", "AbortError")),
          { once: true },
        );
      });
    },
  });
  const mounted = await mount(PdfViewer, {
    source: {
      kind: "blob",
      blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }),
    },
    runtime: fake.runtime,
    onDocumentLoadError: () => { errorEvents += 1; },
  });
  await waitFor(() => observedSignal, "active PDF render signal");
  assert.equal(observedSignal.aborted, false);
  mounted.app.unmount();
  await waitFor(() => observedSignal.aborted, "PDF render abort on unmount");
  await waitFor(() => fake.calls.closeDocument.length >= 1, "PDF close on render cancellation");
  assert.equal(errorEvents, 0, "unmount cancellation must not emit a stale load error");
  assert.equal(fake.calls.closeDocument[0].id, fake.document.id);
  assert.equal(fake.calls.dispose.length, 0, "caller-owned runtime remains usable");
});

test("PdfViewer downloads verified bytes without reading the source again and revokes its URL", async (t) => {
  const originalDocument = globalThis.document;
  const originalCreateObjectURL = URL.createObjectURL;
  const originalRevokeObjectURL = URL.revokeObjectURL;
  const actions = [];
  globalThis.document = {
    body: { append: () => actions.push("append") },
    createElement: () => ({
      click: () => actions.push("click"),
      remove: () => actions.push("remove"),
    }),
  };
  let imageSequence = 0;
  URL.createObjectURL = (blob) => {
    if (blob.type === "application/pdf") {
      actions.push("create");
      return "blob:component-test-download";
    }
    return `blob:component-test-image-${++imageSequence}`;
  };
  URL.revokeObjectURL = (url) => {
    if (url === "blob:component-test-download") actions.push(`revoke:${url}`);
  };
  t.after(() => {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    URL.createObjectURL = originalCreateObjectURL;
    URL.revokeObjectURL = originalRevokeObjectURL;
  });

  const diagnostics = [];
  const pdfBytes = new TextEncoder().encode("%PDF-1.4\n%%EOF");
  const sourceBlob = new Blob([pdfBytes], { type: "application/pdf" });
  const readSource = sourceBlob.arrayBuffer.bind(sourceBlob);
  let sourceReads = 0;
  sourceBlob.arrayBuffer = () => {
    sourceReads += 1;
    return readSource();
  };
  const fake = createPdfRuntime({ pageCount: 1 });
  const mounted = await mount(PdfViewer, {
    source: { kind: "blob", blob: sourceBlob },
    runtime: fake.runtime,
    onDiagnostic: (event) => diagnostics.push(event),
  });
  await waitFor(
    () => findByTestId(mounted.root, "pdf-status")?.props["data-state"] === "ready",
    "downloadable PDF ready state",
  );
  assert.equal(sourceReads, 1);
  assert.equal(fake.calls.openDocument.length, 1);
  await findByTestId(mounted.root, "pdf-download").props.onClick();
  await new Promise((resolve) => setTimeout(resolve, 10));
  assert.deepEqual(actions, ["create", "append", "click", "remove", "revoke:blob:component-test-download"]);
  assert.equal(sourceReads, 1, "download reuses the verified in-memory copy");
  assert.ok(diagnostics.some((event) => event.type === "download" && event.bytes === pdfBytes.byteLength));
  mounted.app.unmount();
  await waitFor(() => fake.calls.closeDocument.length === 1, "downloaded PDF close");
  assert.equal(fake.calls.dispose.length, 0);
});
