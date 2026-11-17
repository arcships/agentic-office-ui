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

const { DocxViewer, renderParagraphRuns } = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel, wasmBuildDocModelFromBytes } = await importFromDemo("@arcships/docx-core");
const { XlsxViewer, useXlsxSearch, useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");
const { PdfSurface, PdfViewer } = await importFromDemo("@arcships/vue-pdf");
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
  onGetPageAnnotations,
  onGetPageTextGeometry,
  onGetTextSlices,
  onRenderPage,
  onRenderThumbnail,
  onSearch,
  onSetFormFieldValue,
  onUpdateAnnotationAppearance,
} = {}) {
  const calls = {
    closeDocument: [],
    dispose: [],
    getPageAnnotations: [],
    getPageTextGeometry: [],
    getTextSlices: [],
    openDocument: [],
    renderPage: [],
    renderThumbnail: [],
    search: [],
    setFormFieldValue: [],
    updateAnnotationAppearance: [],
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
    async getPageTextGeometry(currentDocument, pageIndex, options = {}) {
      calls.getPageTextGeometry.push({ document: currentDocument, options, pageIndex });
      return onGetPageTextGeometry
        ? onGetPageTextGeometry(currentDocument, pageIndex, options)
        : { runs: [] };
    },
    ...(onGetPageAnnotations
      ? {
          async getPageAnnotations(currentDocument, pageIndex, options = {}) {
            calls.getPageAnnotations.push({ document: currentDocument, options, pageIndex });
            return onGetPageAnnotations(currentDocument, pageIndex, options);
          },
        }
      : {}),
    ...(onSetFormFieldValue
      ? {
          async setFormFieldValue(currentDocument, pageIndex, annotationId, value, options = {}) {
            calls.setFormFieldValue.push({ document: currentDocument, pageIndex, annotationId, value, options });
            return onSetFormFieldValue(currentDocument, pageIndex, annotationId, value, options);
          },
        }
      : {}),
    ...(onUpdateAnnotationAppearance
      ? {
          async updateAnnotationAppearance(currentDocument, pageIndex, annotationId, patch, options = {}) {
            calls.updateAnnotationAppearance.push({ document: currentDocument, pageIndex, annotationId, patch, options });
            return onUpdateAnnotationAppearance(currentDocument, pageIndex, annotationId, patch, options);
          },
        }
      : {}),
    async getTextSlices(currentDocument, slices, options = {}) {
      calls.getTextSlices.push({ document: currentDocument, options, slices });
      return onGetTextSlices
        ? onGetTextSlices(currentDocument, slices, options)
        : slices.map(() => "");
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

function pdfTextGeometry(text) {
  return {
    runs: [{
      charStart: 0,
      rect: { x: 0, y: 0, width: text.length * 10, height: 12 },
      glyphs: Array.from(text, (character, index) => ({
        x: index * 10,
        y: 0,
        width: 8,
        height: 12,
        flags: character === " " ? 1 : 0,
      })),
    }],
  };
}

function preparePdfSurfaceInteraction(root) {
  const surface = findByTestId(root, "pdf-surface");
  const interaction = walk(root).find((node) => node.props?.class === "pdf-surface__interaction");
  assert.ok(surface);
  assert.ok(interaction);
  const captures = new Set();
  surface.focus = () => {};
  interaction.getBoundingClientRect = () => ({ left: 0, top: 0, width: 600, height: 800 });
  interaction.setPointerCapture = (pointerId) => captures.add(pointerId);
  interaction.hasPointerCapture = (pointerId) => captures.has(pointerId);
  interaction.releasePointerCapture = (pointerId) => captures.delete(pointerId);
  return interaction;
}

function pdfPointerEvent(interaction, { clientX, clientY = 6, pointerId = 1 }) {
  return {
    button: 0,
    clientX,
    clientY,
    currentTarget: interaction,
    isPrimary: true,
    pointerId,
    pointerType: "mouse",
    preventDefault() {},
  };
}

function pdfInteractionLayers(root) {
  const layers = new Map();
  for (const node of walk(root)) {
    if (node.props?.class !== "pdf-surface__interaction") continue;
    const slot = node.parent;
    const pageIndex = Number(slot?.props?.["data-page-index"]);
    if (!Number.isSafeInteger(pageIndex)) continue;
    const captures = new Set();
    node.getBoundingClientRect = () => ({ left: 0, top: 0, width: 600, height: 800 });
    node.setPointerCapture = (pointerId) => captures.add(pointerId);
    node.hasPointerCapture = (pointerId) => captures.has(pointerId);
    node.releasePointerCapture = (pointerId) => captures.delete(pointerId);
    layers.set(pageIndex, { layer: node, slot });
  }
  return layers;
}

/** Route pointer hit-testing to a specific rendered page, like a real document. */
function routePdfPointerToPage(surface, page) {
  page.slot.dataset = { pageIndex: String(Number(page.slot.props["data-page-index"])) };
  page.slot.querySelector = (selector) =>
    selector === ".pdf-surface__interaction" ? page.layer : null;
  surface.ownerDocument = {
    elementFromPoint: (clientX, clientY) =>
      clientY >= 0 && clientY <= 800 ? { closest: () => page.slot } : null,
  };
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

test("DocxViewer hides upload and download entry points through showUpload and showDownload", async () => {
  const shown = await mount(DocxViewer, { emptyState: "DOCX empty" });
  assert.ok(findByTestId(shown.root, "docx-upload"), "upload button visible by default");
  assert.ok(findByTestId(shown.root, "docx-download"), "download button visible by default");
  assert.match(textContent(shown.root), /Choose Word file/);
  assert.deepEqual(shown.warnings, []);
  shown.app.unmount();

  const hidden = await mount(DocxViewer, {
    emptyState: "DOCX empty",
    showUpload: false,
    showDownload: false,
  });
  assert.ok(findByTestId(hidden.root, "docx-viewer-toolbar"), "toolbar still renders");
  assert.equal(findByTestId(hidden.root, "docx-upload"), undefined, "upload button hidden");
  assert.equal(findByTestId(hidden.root, "docx-download"), undefined, "download button hidden");
  assert.doesNotMatch(textContent(hidden.root), /Choose Word file/);
  assert.deepEqual(hidden.warnings, []);
  hidden.app.unmount();

  const failure = new Error("broken DOCX");
  failure.code = "PARSE_FAILED";
  const failed = await mount(DocxViewer, {
    file: new Uint8Array([9]).buffer,
    showUpload: false,
    runtime: {
      createLoader: () => ({ load: async () => { throw failure; }, cancel() {}, dispose() {} }),
    },
  });
  await waitFor(
    () => findByTestId(failed.root, "docx-viewer")?.props["data-state"] === "error",
    "DOCX error state",
  );
  assert.doesNotMatch(textContent(failed.root), /Choose another Word file/);
  assert.ok(findByTestId(failed.root, "docx-download"), "download button retained when only upload is hidden");
  assert.deepEqual(failed.warnings, []);
  failed.app.unmount();
});

test("DocxViewer resolves locale and host message overrides for visible copy and aria-labels", async () => {
  const zh = await mount(DocxViewer, { locale: "zh-CN" });
  assert.match(textContent(zh.root), /未打开文档/);
  assert.equal(
    findByTestId(zh.root, "docx-sidebar-toggle").props["aria-label"],
    "显示或隐藏页面缩略图",
  );
  assert.deepEqual(zh.warnings, []);
  zh.app.unmount();

  const custom = await mount(DocxViewer, {
    messages: { docx: { viewer: { emptyTitle: "Nothing open" } } },
  });
  assert.match(textContent(custom.root), /Nothing open/);
  assert.deepEqual(custom.warnings, []);
  custom.app.unmount();
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

function flattenVNodes(value, result = []) {
  if (Array.isArray(value)) {
    for (const item of value) flattenVNodes(item, result);
    return result;
  }
  if (!value || typeof value !== "object") return result;
  result.push(value);
  flattenVNodes(value.children, result);
  return result;
}

test("DOCX inline revisions render underlines for insertions and strikethrough for deletions", async () => {
  const file = readFileSync(new URL("../../apps/demo/public/samples/review-comments.docx", import.meta.url));
  const bytes = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength);
  const { model } = await wasmBuildDocModelFromBytes(bytes);
  const paragraph = model.nodes.find(
    (node) => node.type === "paragraph"
      && /<w:ins\b/.test(node.sourceXml ?? "")
      && /<w:del\b/.test(node.sourceXml ?? ""),
  );
  assert.ok(paragraph, "fixture contains a paragraph with both an insertion and a deletion");

  const vnodes = renderParagraphRuns(
    paragraph,
    "tracked-revision",
    "light",
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    { showTrackedChanges: true, trackedMarkupMode: "inline" },
  );
  const trackedSpans = flattenVNodes(vnodes).filter(
    (node) => node.props?.["data-docx-tracked-change"],
  );
  const insertionSpan = trackedSpans.find(
    (node) => node.props["data-docx-tracked-change"] === "insertion",
  );
  const deletionSpan = trackedSpans.find(
    (node) => node.props["data-docx-tracked-change"] === "deletion",
  );
  assert.ok(insertionSpan, "insertion run is rendered");
  assert.ok(deletionSpan, "deletion run is rendered");

  const decorationTokens = (node) => new Set(
    String(node.props.style?.textDecoration ?? "")
      .split(/\s+/)
      .filter(Boolean),
  );
  const insertionDecoration = decorationTokens(insertionSpan);
  const deletionDecoration = decorationTokens(deletionSpan);
  assert.ok(insertionDecoration.has("underline"), "insertions are underlined");
  assert.equal(insertionDecoration.has("line-through"), false);
  assert.equal(insertionDecoration.has("none"), false);
  assert.ok(deletionDecoration.has("line-through"), "deletions are struck through");
  assert.equal(deletionDecoration.has("underline"), false);
  assert.equal(deletionDecoration.has("none"), false);
  assert.notEqual(
    insertionSpan.props.style.color,
    deletionSpan.props.style.color,
    "insertions and deletions use different colours",
  );
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
    [xlsxController({ isLoading: true }), "Loading"],
    [xlsxController({ error: new Error("XLSX failed") }), "XLSX failed"],
    [xlsxController(), "Open an XLSX file"],
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

test("XlsxViewer resolves locale and host message overrides for visible copy and aria-labels", async () => {
  const controller = xlsxController({
    canZoomIn: false,
    canZoomOut: false,
    displayFileName: "book.xlsx",
    zoomScale: 100,
  });
  const zh = await mount(XlsxViewer, {
    controller,
    locale: "zh-CN",
    showFormulaBar: false,
    showRibbon: false,
  });
  assert.match(textContent(zh.root), /打开一个 XLSX 文件开始查看/);
  assert.equal(findByTestId(zh.root, "xlsx-search-toggle").props["aria-label"], "搜索工作簿");
  assert.deepEqual(zh.warnings, []);
  zh.app.unmount();

  const custom = await mount(XlsxViewer, {
    controller,
    messages: {
      xlsx: {
        toolbar: { search: "Find cells" },
        viewer: { empty: "Nothing to show" },
      },
    },
    showFormulaBar: false,
    showRibbon: false,
  });
  assert.match(textContent(custom.root), /Nothing to show/);
  assert.equal(findByTestId(custom.root, "xlsx-search-toggle").props["aria-label"], "Find cells");
  assert.deepEqual(custom.warnings, []);
  custom.app.unmount();
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

test("XLSX surface search returns exact value/formula ranges and activates a merged-cell anchor", async () => {
  const activated = [];
  const scrolled = [];
  let worksheetFreeCount = 0;
  const sheet = {
    name: "Data",
    workbookSheetIndex: 4,
    minUsedRow: 2,
    maxUsedRow: 2,
    minUsedCol: 2,
    maxUsedCol: 4,
    visibleRows: [2],
    visibleCols: [3, 4],
    mergedRegions: [{ start: { row: 2, col: 2 }, end: { row: 2, col: 3 } }],
  };
  const worksheet = {
    getFormattedValueAt: (_row, col) => col === 3 ? "Merged Needle Value" : "",
    getFormulaAt: (_row, col) => col === 4 ? "=NEEDLE(1)" : "",
    getCalculatedValueAt: () => ({ is_empty: true, toString: () => "", free() {} }),
    free: () => { worksheetFreeCount += 1; },
  };
  const controller = {
    activeSheetIndex: 0,
    isWorkerBacked: false,
    sheets: [sheet],
    workbook: { getSheet: () => worksheet },
    setActiveSheetIndex(index) { this.activeSheetIndex = index; },
    selectCell: (cell) => activated.push(cell),
  };

  const scope = vue.effectScope();
  let search;
  scope.run(() => {
    search = useXlsxSearch(() => controller, {
      scrollToCell: (cell) => scrolled.push(cell),
    });
  });
  const state = await search.search("needle");

  assert.equal(state.status, "ready");
  assert.equal(state.activeIndex, 0);
  assert.equal(state.matches.length, 2);
  assert.deepEqual(state.matches[0].cell, { row: 2, col: 2 });
  assert.equal(state.matches[0].address, "C3");
  assert.equal(state.matches[0].start, 7);
  assert.equal(state.matches[0].end, 13);
  assert.equal(state.matches[1].text, "=NEEDLE(1)");
  assert.equal(state.matches[1].start, 1);
  assert.equal(state.matches[1].end, 7);
  assert.deepEqual(scrolled, [{ row: 2, col: 2 }]);
  assert.deepEqual(activated, [{ row: 2, col: 2 }]);
  assert.equal(worksheetFreeCount, 1);
  scope.stop();
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

test("PdfSurface finalizes a drag released before page geometry resolves", async () => {
  const geometryGate = deferred();
  const text = "abcd";
  const fake = createPdfRuntime({
    pageCount: 1,
    onGetPageTextGeometry: () => geometryGate.promise,
    onGetTextSlices: (_document, slices) => slices.map((slice) =>
      text.slice(slice.charIndex, slice.charIndex + slice.charCount)),
  });
  const selections = [];
  const mounted = await mount(PdfSurface, {
    source: { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) },
    runtime: fake.runtime,
    onSelectionChange: (selection) => selections.push(selection),
  });
  await waitFor(
    () => fake.calls.getPageTextGeometry.length === 1 &&
      walk(mounted.root).some((node) => node.props?.class === "pdf-surface__interaction"),
    "PDF selection layer and prefetched geometry request",
  );
  const interaction = preparePdfSurfaceInteraction(mounted.root);

  interaction.props.onPointerdown(pdfPointerEvent(interaction, { clientX: 2 }));
  interaction.props.onPointerup(pdfPointerEvent(interaction, { clientX: 38 }));
  assert.equal(fake.calls.getPageTextGeometry.length, 1, "pointerdown reuses the prefetched request");

  geometryGate.resolve(pdfTextGeometry(text));
  await waitFor(() => selections.at(-1)?.kind === "text", "late PDF drag selection");
  assert.equal(selections.at(-1).text, text);
  assert.deepEqual(selections.at(-1).range, { pageIndex: 0, charIndex: 0, charCount: 4 });
  mounted.app.unmount();
});

test("PdfSurface uses the native second click to select a Unicode word", async () => {
  const text = "hello world";
  const fake = createPdfRuntime({
    pageCount: 1,
    onGetPageTextGeometry: () => pdfTextGeometry(text),
    onGetTextSlices: (_document, slices) => slices.map((slice) =>
      text.slice(slice.charIndex, slice.charIndex + slice.charCount)),
  });
  const selections = [];
  const mounted = await mount(PdfSurface, {
    source: { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) },
    runtime: fake.runtime,
    onSelectionChange: (selection) => selections.push(selection),
  });
  await waitFor(
    () => fake.calls.getPageTextGeometry.length === 1 &&
      walk(mounted.root).some((node) => node.props?.class === "pdf-surface__interaction"),
    "PDF word selection layer",
  );
  const interaction = preparePdfSurfaceInteraction(mounted.root);
  const point = { clientX: 65, pointerId: 1 };

  interaction.props.onPointerdown(pdfPointerEvent(interaction, point));
  interaction.props.onPointerup(pdfPointerEvent(interaction, point));
  await interaction.props.onClick({
    button: 0,
    clientX: point.clientX,
    clientY: 6,
    currentTarget: interaction,
    detail: 1,
    preventDefault() {},
  });
  interaction.props.onPointerdown(pdfPointerEvent(interaction, { ...point, pointerId: 2 }));
  interaction.props.onPointerup(pdfPointerEvent(interaction, { ...point, pointerId: 2 }));
  await interaction.props.onClick({
    button: 0,
    clientX: point.clientX,
    clientY: 6,
    currentTarget: interaction,
    detail: 2,
    preventDefault() {},
  });

  await waitFor(() => selections.at(-1)?.kind === "text", "double-click PDF word selection");
  assert.equal(selections.at(-1).text, "world");
  assert.deepEqual(selections.at(-1).range, { pageIndex: 0, charIndex: 6, charCount: 5 });
  mounted.app.unmount();
});

test("PdfSurface exposes the unified controlled object/region reference contract", async () => {
  const fake = createPdfRuntime({
    pageCount: 1,
    onGetPageTextGeometry: () => pdfTextGeometry("reference surface"),
  });
  const source = { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) };
  const confirmations = [];
  const resolutions = [];
  let api;
  let mode;
  const Harness = vue.defineComponent({
    setup() {
      const surface = vue.ref(null);
      mode = vue.ref("object");
      vue.watch(surface, (value) => { api = value; }, { flush: "sync" });
      return () => vue.h(PdfSurface, {
        ref: surface,
        source,
        runtime: fake.runtime,
        documentId: "reference.pdf",
        selectionMode: mode.value,
        emitReferenceCandidates: true,
        onReferenceConfirm: (event) => confirmations.push(event),
        onReferenceResolve: (event) => resolutions.push(event),
      });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(() => api && fake.calls.getPageTextGeometry.length === 1, "PDF unified reference surface");
  let interaction = preparePdfSurfaceInteraction(mounted.root);

  await interaction.props.onClick({
    button: 0,
    clientX: 20,
    clientY: 20,
    currentTarget: interaction,
    detail: 1,
    shiftKey: false,
  });
  assert.equal(confirmations.length, 1);
  assert.equal(confirmations[0].reference.kind, "page");
  assert.equal(confirmations[0].reference.document.documentId, "reference.pdf");
  assert.equal(api.getDocumentRevision().format, "pdf");
  const result = await api.resolveReference(confirmations[0].reference);
  assert.equal(result.status, "exact");
  assert.equal(resolutions.at(-1).referenceId, confirmations[0].reference.referenceId);

  mode.value = "region";
  await vue.nextTick();
  interaction = preparePdfSurfaceInteraction(mounted.root);
  interaction.props.onPointerdown(pdfPointerEvent(interaction, { clientX: 10, clientY: 10, pointerId: 7 }));
  interaction.props.onPointermove(pdfPointerEvent(interaction, { clientX: 110, clientY: 210, pointerId: 7 }));
  interaction.props.onPointerup(pdfPointerEvent(interaction, { clientX: 110, clientY: 210, pointerId: 7 }));
  assert.equal(confirmations.at(-1).reference.kind, "region");
  assert.equal(confirmations.at(-1).reference.locator.value.space, "page");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfSurface drag selection spans pages and copies the joined text", async () => {
  const pageTexts = ["abcdefgh", "ijklmnop"];
  const fake = createPdfRuntime({
    pageCount: 2,
    onGetPageTextGeometry: (_document, pageIndex) => pdfTextGeometry(pageTexts[pageIndex]),
    onGetTextSlices: (_document, slices) => slices.map((slice) =>
      pageTexts[slice.pageIndex].slice(slice.charIndex, slice.charIndex + slice.charCount)),
  });
  const selections = [];
  const mounted = await mount(PdfSurface, {
    source: { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) },
    runtime: fake.runtime,
    onSelectionChange: (selection) => selections.push(selection),
  });
  await waitFor(
    () => fake.calls.getPageTextGeometry.length === 2 && pdfInteractionLayers(mounted.root).size === 2,
    "PDF selection layers for both pages",
  );
  const surface = findByTestId(mounted.root, "pdf-surface");
  surface.focus = () => {};
  const layers = pdfInteractionLayers(mounted.root);
  const pageOne = layers.get(0);
  const pageTwo = layers.get(1);
  routePdfPointerToPage(surface, pageTwo);

  pageOne.layer.props.onPointerdown(pdfPointerEvent(pageOne.layer, { clientX: 2, pointerId: 11 }));
  pageOne.layer.props.onPointermove(pdfPointerEvent(pageOne.layer, { clientX: 35, pointerId: 11 }));
  pageOne.layer.props.onPointerup(pdfPointerEvent(pageOne.layer, { clientX: 35, pointerId: 11 }));

  await waitFor(() => selections.at(-1)?.kind === "text", "cross-page PDF drag selection");
  const selection = selections.at(-1);
  assert.equal(selection.text, "abcdefgh\nijkl");
  assert.deepEqual(selection.range, { pageIndex: 0, charIndex: 0, charCount: 8 });
  assert.deepEqual(selection.segments.map((segment) => segment.range), [
    { pageIndex: 0, charIndex: 0, charCount: 8 },
    { pageIndex: 1, charIndex: 0, charCount: 4 },
  ]);
  assert.deepEqual(selection.segments.map((segment) => segment.text), ["abcdefgh", "ijkl"]);
  assert.equal(selection.segments[0].rects.length > 0, true);
  assert.equal(selection.segments[1].rects.length > 0, true);
  assert.equal(
    walk(mounted.root).filter((node) => node.props?.class === "pdf-surface__selection-hit").length,
    selection.segments[0].rects.length + selection.segments[1].rects.length,
    "selection rectangles render on both pages",
  );

  const copied = [];
  surface.props.onCopy({
    clipboardData: { setData: (type, value) => copied.push({ type, value }) },
    preventDefault() {},
  });
  assert.deepEqual(copied, [{ type: "text/plain", value: "abcdefgh\nijkl" }]);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfSurface forwards Escape from the region layer exactly once", async () => {
  const fake = createPdfRuntime({ pageCount: 1, onGetPageTextGeometry: () => pdfTextGeometry("region") });
  const cancellations = [];
  const mounted = await mount(PdfSurface, {
    source: { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) },
    runtime: fake.runtime,
    selectionMode: "region",
    onSelectionCancel: (event) => cancellations.push(event),
  });
  await waitFor(
    () => walk(mounted.root).some((node) => node.props?.class === "pdf-surface__interaction"),
    "PDF region interaction layer",
  );
  const surface = findByTestId(mounted.root, "pdf-surface");
  const interaction = preparePdfSurfaceInteraction(mounted.root);
  let stopped = false;
  interaction.props.onKeydown({
    key: "Escape",
    stopPropagation() { stopped = true; },
    preventDefault() {},
  });
  assert.equal(stopped, true, "region layer stops the Escape bubble");
  assert.deepEqual(cancellations, [{ mode: "region", reason: "escape" }]);

  surface.props.onKeydown({ key: "Escape", defaultPrevented: true });
  assert.equal(cancellations.length, 1, "an already-handled Escape is not emitted twice");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfSurface exposes link and widget hit rectangles and forwards activation", async () => {
  const annotations = [
    { id: "link-1", pageIndex: 0, kind: "link", rect: { x: 10, y: 10, width: 50, height: 20 }, url: "https://example.com/a" },
    { id: "widget-1", pageIndex: 0, kind: "widget", rect: { x: 100, y: 10, width: 40, height: 20 }, fieldName: "name", fieldValue: "Ada" },
  ];
  const fake = createPdfRuntime({
    pageCount: 1,
    onGetPageAnnotations: () => annotations,
    onGetPageTextGeometry: () => pdfTextGeometry("annotated"),
  });
  const links = [];
  const activations = [];
  const mounted = await mount(PdfSurface, {
    source: { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) },
    runtime: fake.runtime,
    onLinkActivate: (event) => links.push(event),
    onAnnotationActivate: (event) => activations.push(event),
  });
  await waitFor(
    () => fake.calls.getPageAnnotations.length === 1 &&
      walk(mounted.root).some((node) => node.props?.["data-annotation-kind"] === "link"),
    "PDF annotation hit layer",
  );
  const hits = walk(mounted.root).filter((node) => node.props?.["data-annotation-kind"]);
  assert.deepEqual(hits.map((node) => node.props["data-annotation-kind"]).sort(), ["link", "widget"]);
  assert.equal(hits[0].props.style.left, "10px");

  const interaction = preparePdfSurfaceInteraction(mounted.root);
  const click = async (clientX, clientY) => {
    interaction.props.onPointerdown(pdfPointerEvent(interaction, { clientX, clientY }));
    interaction.props.onPointerup(pdfPointerEvent(interaction, { clientX, clientY }));
    await interaction.props.onClick({
      button: 0,
      clientX,
      clientY,
      currentTarget: interaction,
      detail: 1,
      preventDefault() {},
    });
  };

  await click(30, 15);
  assert.deepEqual(links, [{
    pageIndex: 0,
    annotationId: "link-1",
    clientX: 30,
    clientY: 15,
    url: "https://example.com/a",
  }]);
  assert.deepEqual(activations, []);

  await click(110, 15);
  assert.deepEqual(activations, [{
    pageIndex: 0,
    annotationId: "widget-1",
    kind: "widget",
    clientX: 110,
    clientY: 15,
    fieldName: "name",
    fieldValue: "Ada",
  }]);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfSurface writes AcroForm values through the runtime and refreshes the page", async () => {
  const widgets = [
    {
      id: "widget-name",
      pageIndex: 0,
      kind: "widget",
      rect: { x: 10, y: 10, width: 120, height: 20 },
      fieldName: "name",
      fieldValue: "",
      fieldType: "text",
    },
    {
      id: "widget-agree",
      pageIndex: 0,
      kind: "widget",
      rect: { x: 10, y: 40, width: 16, height: 16 },
      fieldName: "agree",
      fieldValue: "Off",
      fieldType: "checkbox",
      exportValue: "Yes",
    },
    {
      id: "widget-plan",
      pageIndex: 0,
      kind: "widget",
      rect: { x: 10, y: 70, width: 90, height: 20 },
      fieldName: "plan",
      fieldValue: "Alpha",
      fieldType: "combobox",
      options: [
        { label: "Alpha", selected: true },
        { label: "Beta", selected: false },
      ],
    },
  ];
  const fake = createPdfRuntime({
    pageCount: 1,
    onGetPageAnnotations: () => widgets,
    onGetPageTextGeometry: () => pdfTextGeometry("form"),
    onSetFormFieldValue: (_document, pageIndex, annotationId, value) => ({
      ...widgets.find((widget) => widget.id === annotationId),
      fieldValue: value === true ? "Yes" : String(value),
    }),
  });
  const changes = [];
  const mounted = await mount(PdfSurface, {
    source: { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) },
    runtime: fake.runtime,
    onFormFieldChange: (event) => changes.push(event),
  });
  await waitFor(
    () => walk(mounted.root).some((node) => node.props?.["data-form-field-id"] === "widget-name"),
    "PDF AcroForm controls",
  );
  const controls = walk(mounted.root).filter((node) => node.props?.["data-form-field-id"]);
  assert.deepEqual(
    controls.map((node) => node.props["data-form-field-type"]).sort(),
    ["checkbox", "combobox", "text"],
  );
  const planControl = controls.find((node) => node.props["data-form-field-id"] === "widget-plan");
  assert.equal(planControl.type, "select");
  assert.deepEqual(planControl.options.map((option) => option.props.value), ["Alpha", "Beta"]);

  const rendersBefore = fake.calls.renderPage.length;
  await controls.find((node) => node.props["data-form-field-id"] === "widget-name").props.onChange({
    target: { value: "Ada Lovelace" },
  });
  assert.deepEqual(fake.calls.setFormFieldValue.at(-1).annotationId, "widget-name");
  assert.equal(fake.calls.setFormFieldValue.at(-1).value, "Ada Lovelace");
  assert.equal(changes.at(-1).fieldName, "name");
  assert.equal(changes.at(-1).value, "Ada Lovelace");
  await waitFor(() => fake.calls.renderPage.length > rendersBefore, "form write page refresh");

  await controls.find((node) => node.props["data-form-field-id"] === "widget-agree").props.onChange({
    target: { checked: true, value: "Yes" },
  });
  assert.equal(fake.calls.setFormFieldValue.at(-1).value, true);
  assert.equal(changes.at(-1).value, "Yes");

  await planControl.props.onChange({ target: { value: "Beta" } });
  assert.equal(fake.calls.setFormFieldValue.at(-1).annotationId, "widget-plan");
  assert.equal(fake.calls.setFormFieldValue.at(-1).value, "Beta");
  assert.equal(changes.at(-1).value, "Beta");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfSurface updates annotation appearance through the exposed API", async () => {
  const annotations = [
    { id: "note-1", pageIndex: 0, kind: "annotation", rect: { x: 10, y: 10, width: 40, height: 20 }, contents: "old", color: "#ffff00", opacity: 1 },
  ];
  const fake = createPdfRuntime({
    pageCount: 1,
    onGetPageAnnotations: () => annotations,
    onGetPageTextGeometry: () => pdfTextGeometry("note"),
    onUpdateAnnotationAppearance: (_document, pageIndex, annotationId, patch) => ({
      ...annotations.find((annotation) => annotation.id === annotationId),
      ...patch,
    }),
  });
  const changes = [];
  let api;
  const Harness = vue.defineComponent({
    setup() {
      const surface = vue.ref(null);
      vue.watch(surface, (value) => { api = value; }, { flush: "sync" });
      return () => vue.h(PdfSurface, {
        ref: surface,
        source: { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) },
        runtime: fake.runtime,
        onAnnotationChange: (event) => changes.push(event),
      });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(() => api && fake.calls.getPageAnnotations.length === 1, "PDF annotation appearance surface");

  const rendersBefore = fake.calls.renderPage.length;
  const updated = await api.updateAnnotationAppearance(0, "note-1", { contents: "reviewed", color: "#ff0000" });
  assert.equal(updated.contents, "reviewed");
  assert.deepEqual(fake.calls.updateAnnotationAppearance.at(-1).patch, { contents: "reviewed", color: "#ff0000" });
  assert.equal(changes.at(-1).annotation.contents, "reviewed");
  await waitFor(() => fake.calls.renderPage.length > rendersBefore, "annotation update page refresh");

  const pageAnnotations = await api.getPageAnnotations(0);
  assert.equal(pageAnnotations.find((annotation) => annotation.id === "note-1").contents, "reviewed");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfViewer exposes localized toolbar aria-labels and message overrides", async () => {
  const zh = await mount(PdfViewer, { locale: "zh-CN" });
  assert.equal(findByTestId(zh.root, "pdf-sidebar-toggle").props["aria-label"], "显示或隐藏缩略图");
  assert.equal(findByTestId(zh.root, "pdf-page-input").props["aria-label"], "页码");
  assert.deepEqual(zh.warnings, []);
  zh.app.unmount();

  const custom = await mount(PdfViewer, {
    messages: { pdf: { viewer: { toggleThumbnails: "Toggle panel" } } },
  });
  assert.equal(findByTestId(custom.root, "pdf-sidebar-toggle").props["aria-label"], "Toggle panel");
  assert.deepEqual(custom.warnings, []);
  custom.app.unmount();
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
  assert.match(textContent(findByTestId(ready.root, "pdf-search-result")), /Result 1 \/ 1, page 3/);
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
  assert.match(rendererErrors[0].message, /could not be parsed or rendered/);
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

test("PdfViewer forwards password credentials to the runtime open call", async () => {
  const fake = createPdfRuntime();
  const onPassword = async () => "resolved-password";
  const mounted = await mount(PdfViewer, {
    source: { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) },
    runtime: fake.runtime,
    password: "initial-password",
    onPassword,
  });
  await waitFor(() => fake.calls.openDocument.length === 1, "PDF runtime open with password options");
  assert.equal(fake.calls.openDocument[0].options.password, "initial-password");
  assert.equal(fake.calls.openDocument[0].options.onPassword, onPassword);
  mounted.app.unmount();
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

test("PdfViewer sanitizes load-start diagnostics instead of leaking query tokens", async () => {
  const diagnostics = [];
  const fetch = (_url, init) => new Promise((_resolve, reject) => {
    init.signal.addEventListener(
      "abort",
      () => reject(new DOMException("Aborted", "AbortError")),
      { once: true },
    );
  });
  const fake = createPdfRuntime();
  const mounted = await mount(PdfViewer, {
    source: { kind: "url", url: "https://files.example/report.pdf?token=secret#page=2" },
    runtime: fake.runtime,
    urlPolicy: {
      allowedOrigins: ["https://files.example"],
      allowedProtocols: ["https:"],
      baseUrl: "https://files.example/",
      fetch,
    },
    onDiagnostic: (event) => diagnostics.push(event),
  });
  await waitFor(
    () => diagnostics.some((event) => event.type === "load-start"),
    "PDF load-start diagnostic",
  );
  const loadStart = diagnostics.find((event) => event.type === "load-start");
  assert.equal(loadStart.url, "https://files.example/report.pdf");
  assert.equal(JSON.stringify(diagnostics).includes("secret"), false);
  assert.equal(JSON.stringify(diagnostics).includes("page=2"), false);
  mounted.app.unmount();
});

test("PdfViewer applies token sanitization to the legacy src path", async () => {
  const diagnostics = [];
  const fetch = (_url, init) => new Promise((_resolve, reject) => {
    init.signal.addEventListener(
      "abort",
      () => reject(new DOMException("Aborted", "AbortError")),
      { once: true },
    );
  });
  const mounted = await mount(PdfViewer, {
    src: "https://files.example/report.pdf?token=legacy-secret#page=3",
    runtime: createPdfRuntime().runtime,
    urlPolicy: {
      allowedOrigins: ["https://files.example"],
      allowedProtocols: ["https:"],
      baseUrl: "https://files.example/",
      fetch,
    },
    onDiagnostic: (event) => diagnostics.push(event),
  });
  await waitFor(
    () => diagnostics.some((event) => event.type === "load-start"),
    "legacy PDF load-start diagnostic",
  );
  assert.equal(
    diagnostics.find((event) => event.type === "load-start").url,
    "https://files.example/report.pdf",
  );
  assert.equal(JSON.stringify(diagnostics).includes("legacy-secret"), false);
  assert.equal(JSON.stringify(diagnostics).includes("page=3"), false);
  mounted.app.unmount();
});

test("PdfViewer follows same-origin redirects and reports only the sanitized final URL", async () => {
  const bytes = new TextEncoder().encode("%PDF-1.4\n%%EOF");
  const diagnostics = [];
  const fetch = async () => ({
    ok: true,
    status: 200,
    url: "https://files.example/cdn/final.pdf?token=redirect-secret#fragment",
    headers: { get: (name) => name.toLowerCase() === "content-type" ? "application/pdf" : null },
    arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  });
  const mounted = await mount(PdfViewer, {
    source: { kind: "url", url: "https://files.example/report.pdf?token=request-secret" },
    runtime: createPdfRuntime().runtime,
    urlPolicy: {
      allowedOrigins: ["https://files.example"],
      allowedProtocols: ["https:"],
      baseUrl: "https://files.example/",
      fetch,
    },
    onDiagnostic: (event) => diagnostics.push(event),
  });
  await waitFor(
    () => diagnostics.some((event) => event.type === "load-success"),
    "redirected PDF load-success",
  );
  assert.equal(
    diagnostics.find((event) => event.type === "load-success").url,
    "https://files.example/cdn/final.pdf",
  );
  assert.equal(JSON.stringify(diagnostics).includes("secret"), false);
  mounted.app.unmount();
});

test("PdfViewer rejects cross-origin redirects with token-free load errors", async () => {
  const errors = [];
  const diagnostics = [];
  const fetch = async () => ({
    ok: true,
    status: 200,
    url: "https://blocked.invalid/sample.pdf?token=redirect-secret",
    headers: { get: () => null },
    arrayBuffer: async () => new ArrayBuffer(0),
  });
  const mounted = await mount(PdfViewer, {
    source: { kind: "url", url: "https://files.example/report.pdf?token=request-secret" },
    runtime: createPdfRuntime().runtime,
    urlPolicy: {
      allowedOrigins: ["https://files.example"],
      allowedProtocols: ["https:"],
      baseUrl: "https://files.example/",
      fetch,
    },
    onDocumentLoadError: (error) => errors.push(error),
    onDiagnostic: (event) => diagnostics.push(event),
  });
  await waitFor(() => errors.length === 1, "redirected PDF load error");
  assert.equal(errors[0].code, "SOURCE_NOT_ALLOWED");
  assert.equal(errors[0].url, "https://blocked.invalid/sample.pdf");
  assert.equal(JSON.stringify({ errors, diagnostics }).includes("secret"), false);
  assert.equal(
    diagnostics.find((event) => event.type === "load-start").url,
    "https://files.example/report.pdf",
  );
  assert.equal(
    diagnostics.find((event) => event.type === "load-error").error.url,
    "https://blocked.invalid/sample.pdf",
  );
  mounted.app.unmount();
});

test("PdfSurface rejects cross-origin redirects with token-free load errors for source and src", async () => {
  for (const sourceProps of [
    { source: { kind: "url", url: "https://files.example/report.pdf?token=request-secret" } },
    { src: "https://files.example/report.pdf?token=request-secret" },
  ]) {
    const errors = [];
    const fetch = async () => ({
      ok: true,
      status: 200,
      url: "https://blocked.invalid/sample.pdf?token=redirect-secret",
      headers: { get: () => null },
      arrayBuffer: async () => new ArrayBuffer(0),
    });
    const mounted = await mount(PdfSurface, {
      ...sourceProps,
      runtime: createPdfRuntime().runtime,
      urlPolicy: {
        allowedOrigins: ["https://files.example"],
        allowedProtocols: ["https:"],
        baseUrl: "https://files.example/",
        fetch,
      },
      onDocumentLoadError: (error) => errors.push(error),
    });
    await waitFor(() => errors.length === 1, "surface redirected PDF load error");
    assert.equal(errors[0].code, "SOURCE_NOT_ALLOWED");
    assert.equal(errors[0].url, "https://blocked.invalid/sample.pdf");
    assert.equal(JSON.stringify(errors).includes("secret"), false);
    mounted.app.unmount();
  }
});

test("PdfViewer sanitizes the source URL in render-error diagnostics", async () => {
  const bytes = new TextEncoder().encode("%PDF-1.4\n%%EOF");
  const diagnostics = [];
  const fake = createPdfRuntime({
    onRenderPage: () => { throw new Error("render exploded"); },
  });
  const fetch = async () => ({
    ok: true,
    status: 200,
    url: "https://files.example/report.pdf?token=request-secret",
    headers: { get: (name) => name.toLowerCase() === "content-type" ? "application/pdf" : null },
    arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  });
  const mounted = await mount(PdfViewer, {
    source: { kind: "url", url: "https://files.example/report.pdf?token=request-secret" },
    runtime: fake.runtime,
    urlPolicy: {
      allowedOrigins: ["https://files.example"],
      allowedProtocols: ["https:"],
      baseUrl: "https://files.example/",
      fetch,
    },
    onDiagnostic: (event) => diagnostics.push(event),
  });
  await waitFor(
    () => diagnostics.some((event) => event.type === "render-error"),
    "PDF render-error diagnostic",
  );
  const renderError = diagnostics.find((event) => event.type === "render-error");
  assert.equal(renderError.error.url, "https://files.example/report.pdf");
  assert.equal(JSON.stringify(diagnostics).includes("secret"), false);
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

test("PdfSurface finishes its load task and releases abort listeners across renders", async (t) => {
  const liveAbortListeners = new Set();
  const originalAdd = AbortSignal.prototype.addEventListener;
  const originalRemove = AbortSignal.prototype.removeEventListener;
  AbortSignal.prototype.addEventListener = function (type, listener, options) {
    if (type === "abort") liveAbortListeners.add(listener);
    return originalAdd.call(this, type, listener, options);
  };
  AbortSignal.prototype.removeEventListener = function (type, listener, options) {
    if (type === "abort") liveAbortListeners.delete(listener);
    return originalRemove.call(this, type, listener, options);
  };
  t.after(() => {
    AbortSignal.prototype.addEventListener = originalAdd;
    AbortSignal.prototype.removeEventListener = originalRemove;
  });

  const source = { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) };
  const fake = createPdfRuntime({ pageCount: 2 });
  const zoom = vue.ref(1);
  let surfaceApi;
  const Harness = vue.defineComponent({
    setup() {
      const surface = vue.ref(null);
      vue.watch(surface, (value) => { surfaceApi = value; }, { flush: "sync" });
      return () => vue.h(PdfSurface, { ref: surface, source, runtime: fake.runtime, zoom: zoom.value });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => surfaceApi && fake.calls.renderPage.length === 2,
    "initial surface page renders",
  );
  await waitFor(() => liveAbortListeners.size === 0, "initial load task abort listener cleanup");

  zoom.value = 2;
  await vue.nextTick();
  await waitFor(() => fake.calls.renderPage.length === 4, "zoomed surface page renders");
  assert.equal(fake.calls.renderPage.at(-1).options.scale, 2);
  await waitFor(() => liveAbortListeners.size === 0, "zoomed load task abort listener cleanup");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
  assert.equal(fake.calls.dispose.length, 0);
});

test("PdfSurface disposes its load coordinator so unmount aborts URL loading", async () => {
  let observedSignal;
  let errorEvents = 0;
  const fake = createPdfRuntime();
  const fetch = (_url, init) => {
    observedSignal = init.signal;
    return new Promise((_resolve, reject) => {
      init.signal.addEventListener("abort", () => reject(new DOMException("Aborted", "AbortError")), { once: true });
    });
  };
  const mounted = await mount(PdfSurface, {
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
  await waitFor(() => observedSignal.aborted, "surface URL abort on unmount");
  assert.equal(errorEvents, 0, "unmount cancellation must not emit a stale load error");
  assert.equal(fake.calls.openDocument.length, 0);
  assert.equal(fake.calls.dispose.length, 0, "caller-owned runtime remains usable");
});

test("PdfSurface closes a document opened after unmount instead of adopting it", async () => {
  const openGate = deferred();
  const fake = createPdfRuntime({ onOpenDocument: () => openGate.promise });
  const successes = [];
  const errors = [];
  const mounted = await mount(PdfSurface, {
    source: { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) },
    runtime: fake.runtime,
    onDocumentLoadSuccess: (count) => successes.push(count),
    onDocumentLoadError: (error) => errors.push(error),
  });
  await waitFor(() => fake.calls.openDocument.length === 1, "surface document open");
  mounted.app.unmount();
  openGate.resolve(fake.document);
  await waitFor(() => fake.calls.closeDocument.length === 1, "late surface document close");
  assert.equal(fake.calls.closeDocument[0], fake.document);
  assert.deepEqual(successes, []);
  assert.deepEqual(errors, []);
  assert.equal(fake.calls.renderPage.length, 0, "an unmounted surface must not render pages");
});

test("PdfSurface reloads when urlPolicy, maxFileSize or runtime changes", async () => {
  const source = { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) };
  const first = createPdfRuntime({ pageCount: 1 });
  const second = createPdfRuntime({ pageCount: 1 });
  const maxFileSize = vue.ref(1024 * 1024);
  const runtime = vue.shallowRef(first.runtime);
  const urlPolicy = vue.shallowRef({ allowedOrigins: ["https://files.example"] });
  const successes = [];
  const errors = [];
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(PdfSurface, {
        source,
        runtime: runtime.value,
        maxFileSize: maxFileSize.value,
        urlPolicy: urlPolicy.value,
        onDocumentLoadSuccess: (count) => successes.push(count),
        onDocumentLoadError: (error) => errors.push(error),
      });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => first.calls.openDocument.length === 1 && first.calls.renderPage.length === 1,
    "initial surface load",
  );
  assert.deepEqual(successes, [1]);

  urlPolicy.value = { allowedOrigins: ["https://other.example"] };
  await vue.nextTick();
  await waitFor(() => first.calls.openDocument.length === 2, "urlPolicy surface reload");
  assert.equal(first.calls.closeDocument.length, 1, "the previous document is closed before reloading");

  maxFileSize.value = 4;
  await waitFor(() => errors.length === 1, "maxFileSize surface reload");
  assert.equal(errors[0].code, "PDF_TOO_LARGE");
  assert.equal(first.calls.openDocument.length, 2, "an oversized reload never opens the source");
  assert.equal(first.calls.closeDocument.length, 2);

  maxFileSize.value = 1024 * 1024;
  await waitFor(() => first.calls.openDocument.length === 3, "restored maxFileSize surface reload");

  runtime.value = second.runtime;
  await waitFor(() => second.calls.openDocument.length === 1, "runtime swap surface reload");
  assert.equal(first.calls.closeDocument.length, 3, "the previous runtime document is closed on swap");
  assert.equal(second.calls.closeDocument.length, 0);
  assert.deepEqual(mounted.warnings, []);

  mounted.app.unmount();
  await waitFor(() => second.calls.closeDocument.length === 1, "runtime swap surface unmount close");
  assert.equal(first.calls.dispose.length, 0, "caller-owned runtimes remain usable");
  assert.equal(second.calls.dispose.length, 0);
});

function preparePdfSelectionLayer(root) {
  const layer = walk(root).find((node) => node.props?.["data-testid"] === "pdf-selection-layer");
  assert.ok(layer, "PDF viewer selection layer");
  const captures = new Set();
  layer.getBoundingClientRect = () => ({ left: 0, top: 0, width: 600, height: 800 });
  layer.setPointerCapture = (pointerId) => captures.add(pointerId);
  layer.hasPointerCapture = (pointerId) => captures.has(pointerId);
  layer.releasePointerCapture = (pointerId) => captures.delete(pointerId);
  return layer;
}

function viewerPointerEvent(layer, { clientX, clientY = 6, pointerId = 1 }) {
  return {
    button: 0,
    clientX,
    clientY,
    currentTarget: layer,
    isPrimary: true,
    pointerId,
    pointerType: "mouse",
    preventDefault() {},
  };
}

test("PdfViewer selects text on the rendered page, copies it and clears on Escape", async () => {
  const text = "hello world";
  const fake = createPdfRuntime({
    pageCount: 1,
    onGetPageTextGeometry: () => pdfTextGeometry(text),
    onGetTextSlices: (_document, slices) => slices.map((slice) =>
      text.slice(slice.charIndex, slice.charIndex + slice.charCount)),
  });
  const selections = [];
  const mounted = await mount(PdfViewer, {
    source: { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) },
    runtime: fake.runtime,
    showToolbar: false,
    onSelectionChange: (selection) => selections.push(selection),
  });
  await waitFor(
    () => findByTestId(mounted.root, "pdf-status")?.props["data-state"] === "ready" &&
      walk(mounted.root).some((node) => node.props?.["data-testid"] === "pdf-selection-layer"),
    "PDF viewer page and selection layer",
  );
  const layer = preparePdfSelectionLayer(mounted.root);

  layer.props.onPointerdown(viewerPointerEvent(layer, { clientX: 2, pointerId: 21 }));
  layer.props.onPointermove(viewerPointerEvent(layer, { clientX: 65, pointerId: 21 }));
  layer.props.onPointerup(viewerPointerEvent(layer, { clientX: 65, pointerId: 21 }));

  await waitFor(() => selections.at(-1)?.kind === "text", "PDF viewer drag selection");
  assert.equal(selections.at(-1).text, "hello w");
  assert.deepEqual(selections.at(-1).range, { pageIndex: 0, charIndex: 0, charCount: 7 });
  assert.equal(
    walk(mounted.root).filter((node) => node.props?.class === "pdf-viewer__selection-rect").length > 0,
    true,
    "selection rectangles render over the page",
  );

  const copied = [];
  findByTestId(mounted.root, "pdf-viewer").props.onCopy({
    clipboardData: { setData: (type, value) => copied.push({ type, value }) },
    preventDefault() {},
  });
  assert.deepEqual(copied, [{ type: "text/plain", value: "hello w" }]);

  findByTestId(mounted.root, "pdf-viewer").props.onKeydown({ key: "Escape" });
  assert.equal(selections.at(-1).kind, "none");
  await vue.nextTick();
  assert.equal(
    walk(mounted.root).filter((node) => node.props?.class === "pdf-viewer__selection-rect").length,
    0,
    "Escape clears the selection rectangles",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfViewer uses the native second click to select a word", async () => {
  const text = "hello world";
  const fake = createPdfRuntime({
    pageCount: 1,
    onGetPageTextGeometry: () => pdfTextGeometry(text),
    onGetTextSlices: (_document, slices) => slices.map((slice) =>
      text.slice(slice.charIndex, slice.charIndex + slice.charCount)),
  });
  const selections = [];
  const mounted = await mount(PdfViewer, {
    source: { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) },
    runtime: fake.runtime,
    showToolbar: false,
    onSelectionChange: (selection) => selections.push(selection),
  });
  await waitFor(
    () => findByTestId(mounted.root, "pdf-status")?.props["data-state"] === "ready" &&
      walk(mounted.root).some((node) => node.props?.["data-testid"] === "pdf-selection-layer"),
    "PDF viewer word selection layer",
  );
  const layer = preparePdfSelectionLayer(mounted.root);
  const point = { clientX: 65, pointerId: 31 };

  layer.props.onPointerdown(viewerPointerEvent(layer, point));
  layer.props.onPointerup(viewerPointerEvent(layer, point));
  await layer.props.onClick({
    button: 0,
    clientX: point.clientX,
    clientY: 6,
    currentTarget: layer,
    detail: 1,
    preventDefault() {},
  });
  layer.props.onPointerdown(viewerPointerEvent(layer, { ...point, pointerId: 32 }));
  layer.props.onPointerup(viewerPointerEvent(layer, { ...point, pointerId: 32 }));
  await layer.props.onClick({
    button: 0,
    clientX: point.clientX,
    clientY: 6,
    currentTarget: layer,
    detail: 2,
    preventDefault() {},
  });

  await waitFor(() => selections.at(-1)?.kind === "text", "PDF viewer double-click word selection");
  assert.equal(selections.at(-1).text, "world");
  assert.deepEqual(selections.at(-1).range, { pageIndex: 0, charIndex: 6, charCount: 5 });
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});
