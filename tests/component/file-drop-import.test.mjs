// Issue #109 — drag-and-drop file import for XLSX, PDF and PPTX.
//
// Coverage map:
//   1. shared `useFileDrop` composable: accept list, drag depth, outline ....... composable tests
//   2. XLSX viewer drops a workbook through `controller.loadFile()` .......... xlsx test
//   3. PDF viewer drops a document through its `load()` pipeline ............. pdf test
//   4. PPTX viewer drops a deck through `usePptxDocument` .................... pptx test
//   5. a wrong file type reports a visible message ........................... per-format tests
//   6. the drag-over outline clears after drop/leave ......................... per-format tests
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

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
globalThis.Document ??= class Document {};
globalThis.ShadowRoot ??= class ShadowRoot {};
// Reka UI primitives check refs against these constructors even when the
// custom renderer has no real DOM; the stubs keep mounting DOM-free.
if (globalThis.Element === undefined) globalThis.Element = class Element {};
if (globalThis.DocumentFragment === undefined) {
  globalThis.DocumentFragment = class DocumentFragment {
    constructor() {
      this.children = [];
      this.parent = null;
    }
  };
}
// The composite viewers mount real surfaces; only the browser APIs those
// surfaces touch are stubbed. The renderer supplies the host nodes.
globalThis.ResizeObserver ??= class ResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
};
globalThis.requestAnimationFrame ??= (callback) => setTimeout(() => callback(Date.now()), 0);
globalThis.cancelAnimationFrame ??= (handle) => clearTimeout(handle);
if (typeof Object.prototype.getContext !== "function") {
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    writable: true,
    value: () => null,
  });
}

// A window stub that records the drag-end listeners `useFileDrop` installs, so
// the "outline clears when a drag ends outside the viewer" path is testable.
const windowListeners = new Map();
globalThis.window = {
  addEventListener(type, listener) {
    const listeners = windowListeners.get(type) ?? new Set();
    listeners.add(listener);
    windowListeners.set(type, listeners);
  },
  removeEventListener(type, listener) {
    windowListeners.get(type)?.delete(listener);
  },
  cancelAnimationFrame: globalThis.cancelAnimationFrame,
  clearInterval: globalThis.clearInterval,
  clearTimeout: globalThis.clearTimeout,
  devicePixelRatio: 1,
  getComputedStyle: () => ({ getPropertyValue: () => "" }),
  requestAnimationFrame: globalThis.requestAnimationFrame,
  setInterval: globalThis.setInterval,
  setTimeout: globalThis.setTimeout,
};

function fireWindowEvent(type) {
  for (const listener of [...windowListeners.get(type) ?? []]) {
    listener({ type, preventDefault() {} });
  }
}

const { FILE_DROP_OUTLINE, useFileDrop } = await importFromDemo("@arcships/vue-ui");
const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { XlsxViewer, useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");
const { PdfViewer } = await importFromDemo("@arcships/vue-pdf");
const { PptxViewer } = await importFromDemo("@arcships/vue-pptx");

function configureWasm() {
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}

function workbookBuffer(name) {
  const bytes = readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function sampleFile(name, bytes = [1, 2, 3], type = "") {
  return new File([new Uint8Array(bytes)], name, type ? { type } : undefined);
}

/** A `DragEvent`-shaped object the component renderer can hand to handlers. */
function dropEvent(files, overrides = {}) {
  return {
    preventDefault() {},
    dataTransfer: {
      dropEffect: "none",
      files,
      items: files.map(() => ({ kind: "file" })),
      types: ["Files"],
    },
    currentTarget: null,
    relatedTarget: null,
    ...overrides,
  };
}

function leaveEvent(relatedTarget = null) {
  return { currentTarget: null, relatedTarget, preventDefault() {} };
}

function dropTargetOf(mounted, testId) {
  const target = findByTestId(mounted.root, testId);
  assert.ok(target, `${testId} must render`);
  return target;
}

// ── 1. shared composable ───────────────────────────────────────────────────

function mountDropHarness(accept = { extensions: [".xlsx"] }) {
  const accepted = [];
  const rejected = [];
  let drop;
  const Harness = vue.defineComponent({
    setup() {
      drop = useFileDrop({
        accept,
        onFile: (file) => accepted.push(file.name),
        onReject: (files) => rejected.push(files.map((file) => file.name)),
      });
      return () => vue.h("div", {
        "data-active": drop.isDragOver.value ? "true" : "false",
        "data-testid": "drop-target",
        onDragenter: drop.onDragEnter,
        onDragleave: drop.onDragLeave,
        onDragover: drop.onDragOver,
        onDrop: drop.onDrop,
        style: { outline: drop.isDragOver.value ? FILE_DROP_OUTLINE : "none" },
      });
    },
  });
  return {
    get accepted() { return accepted; },
    get drop() { return drop; },
    get rejected() { return rejected; },
    mount: () => mount(Harness),
  };
}

test("useFileDrop marks the outline on dragover without inflating the depth", async () => {
  const harness = mountDropHarness();
  const mounted = await harness.mount();
  const target = findByTestId(mounted.root, "drop-target");
  const fileEvent = () => dropEvent([sampleFile("book.xlsx")]);

  try {
    // `dragover` repeats forever while the pointer stays inside; it must not
    // accumulate depth or a single `dragleave` would leave the outline stuck.
    target.props.onDragover(fileEvent());
    target.props.onDragover(fileEvent());
    await vue.nextTick();
    assert.equal(target.props.style.outline, FILE_DROP_OUTLINE);
    assert.equal(target.props["data-active"], "true");

    target.props.onDragleave(leaveEvent());
    await vue.nextTick();
    assert.equal(target.props.style.outline, "none", "one dragleave must clear the outline");
    assert.equal(target.props["data-active"], "false");
  } finally {
    mounted.app.unmount();
  }
});

test("useFileDrop keeps the outline until every nested dragenter is left", async () => {
  const harness = mountDropHarness();
  const mounted = await harness.mount();
  const target = findByTestId(mounted.root, "drop-target");

  try {
    target.props.onDragenter(dropEvent([]));
    target.props.onDragenter(dropEvent([]));
    await vue.nextTick();
    assert.equal(target.props.style.outline, FILE_DROP_OUTLINE);

    target.props.onDragleave(leaveEvent());
    await vue.nextTick();
    assert.equal(target.props.style.outline, FILE_DROP_OUTLINE, "one level still holds the drag");

    target.props.onDragleave(leaveEvent());
    await vue.nextTick();
    assert.equal(target.props.style.outline, "none");

    // A drag that ends outside the target clears through the window listener.
    target.props.onDragenter(dropEvent([]));
    await vue.nextTick();
    assert.equal(target.props.style.outline, FILE_DROP_OUTLINE);
    fireWindowEvent("dragend");
    await vue.nextTick();
    assert.equal(target.props.style.outline, "none", "dragend outside the viewer must reset the outline");

    // Leaving for a child element is not a dragleave of the target itself.
    target.props.onDragenter(dropEvent([]));
    const child = { child: true };
    target.props.onDragleave({
      currentTarget: { contains: (node) => node === child },
      relatedTarget: child,
      preventDefault() {},
    });
    await vue.nextTick();
    assert.equal(target.props.style.outline, FILE_DROP_OUTLINE);
  } finally {
    mounted.app.unmount();
  }
});

test("useFileDrop picks the first accepted file and reports rejected drops", async () => {
  const harness = mountDropHarness({ extensions: [".xlsx"], mimeTypes: ["text/csv"] });
  const mounted = await harness.mount();
  const target = findByTestId(mounted.root, "drop-target");

  try {
    target.props.onDrop(dropEvent([sampleFile("notes.txt")]));
    await vue.nextTick();
    assert.deepEqual(harness.rejected, [["notes.txt"]]);
    assert.deepEqual(harness.accepted, []);

    target.props.onDrop(dropEvent([sampleFile("notes.txt"), sampleFile("book.xlsx")]));
    await vue.nextTick();
    assert.deepEqual(harness.accepted, ["book.xlsx"], "the first supported file wins");
    assert.equal(harness.rejected.length, 1);

    // The MIME list accepts a file whose extension is not on the extension list.
    target.props.onDrop(dropEvent([sampleFile("export", [1], "text/csv")]));
    await vue.nextTick();
    assert.deepEqual(harness.accepted, ["book.xlsx", "export"]);
    assert.equal(harness.rejected.length, 1);
  } finally {
    mounted.app.unmount();
  }
});

// ── 2. XLSX viewer ─────────────────────────────────────────────────────────

configureWasm();

test("dropping a workbook on XlsxViewer loads it through the controller loader", async () => {
  const droppedBytes = workbookBuffer("sales-table.xlsx");
  const dropped = [];
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer("financial-model.xlsx"),
        fileName: "financial-model.xlsx",
        useWorker: false,
      });
      return () => vue.h(XlsxViewer, {
        controller,
        onFileDrop: (file) => dropped.push(file.name),
        showDefaultToolbar: false,
        showFormulaBar: false,
        showImages: false,
        showRibbon: false,
      });
    },
  });
  const mounted = await mount(Harness);
  const target = dropTargetOf(mounted, "xlsx-viewer");

  try {
    await waitFor(
      () => controller.sourceState === "ready" || controller.error,
      "initial workbook",
      10_000,
    );
    assert.ifError(controller.error);
    assert.equal(controller.displayFileName, "financial-model.xlsx");

    target.props.onDragover(dropEvent([]));
    await vue.nextTick();
    assert.equal(target.props.style.outline, FILE_DROP_OUTLINE);
    assert.equal(target.props["data-file-drop-active"], "true");

    target.props.onDragleave(leaveEvent());
    await vue.nextTick();
    assert.equal(target.props.style.outline, "none");
    assert.equal(target.props["data-file-drop-active"], undefined);

    target.props.onDrop(dropEvent([
      sampleFile("sales-table.xlsx", new Uint8Array(droppedBytes)),
    ]));
    await waitFor(
      () => controller.displayFileName === "sales-table.xlsx" && controller.sourceState === "ready",
      "dropped workbook",
      15_000,
    );
    assert.ifError(controller.error);
    assert.deepEqual(dropped, ["sales-table.xlsx"], "the drop is reported to the host");
    assert.equal(target.props.style.outline, "none", "the drop clears the outline");

    target.props.onDrop(dropEvent([sampleFile("notes.txt")]));
    await waitFor(
      () => findByTestId(mounted.root, "xlsx-drop-error"),
      "unsupported-drop message",
    );
    assert.match(
      textContent(findByTestId(mounted.root, "xlsx-drop-error")),
      /Only \.xlsx, \.xls, \.xlsm, \.xlsb, \.xltx, \.xltm and \.csv files are supported/,
    );
    assert.equal(controller.displayFileName, "sales-table.xlsx", "the rejected file is not loaded");
    assert.deepEqual(dropped, ["sales-table.xlsx"], "rejected files are not reported as accepted");
  } finally {
    mounted.app.unmount();
  }
});

// ── 3. PDF viewer ──────────────────────────────────────────────────────────

const PDF_SOURCE = { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) };

function createPdfRuntime({ pageCount = 2 } = {}) {
  const calls = { closeDocument: [], openDocument: [], renderPage: [], renderThumbnail: [] };
  const document = {
    id: "pdf-drop-document",
    pageCount,
    pages: Array.from({ length: pageCount }, (_, index) => ({
      index,
      width: 600,
      height: 800,
      rotation: 0,
    })),
  };
  return {
    calls,
    document,
    runtime: {
      runtimeId: "pdf-drop-runtime",
      async openDocument() {
        calls.openDocument.push(true);
        return document;
      },
      async renderPage(currentDocument, pageIndex) {
        calls.renderPage.push(pageIndex);
        return new Blob([`page:${pageIndex}`], { type: "image/png" });
      },
      async renderThumbnail(_document, pageIndex) {
        calls.renderThumbnail.push(pageIndex);
        return new Blob(["thumbnail"], { type: "image/png" });
      },
      async getPageTextGeometry() {
        return { runs: [] };
      },
      async getTextSlices(_document, slices) {
        return slices.map(() => "");
      },
      async search() {
        return [];
      },
      async closeDocument(currentDocument) {
        calls.closeDocument.push(currentDocument);
      },
      async dispose() {},
    },
  };
}

test("dropping a PDF on PdfViewer loads it and rejects other files", async () => {
  const fake = createPdfRuntime();
  const dropped = [];
  const mounted = await mount(PdfViewer, {
    source: PDF_SOURCE,
    runtime: fake.runtime,
    onFileDrop: (file) => dropped.push(file.name),
  });
  const target = dropTargetOf(mounted, "pdf-viewer");
  const viewerState = () => findByTestId(mounted.root, "pdf-status")?.props["data-state"];

  try {
    await waitFor(() => viewerState() === "ready", "initial PDF", 10_000);
    assert.equal(fake.calls.openDocument.length, 1);

    target.props.onDragover(dropEvent([]));
    await vue.nextTick();
    assert.equal(target.props.style.outline, FILE_DROP_OUTLINE);
    assert.equal(target.props["data-file-drop-active"], "true");

    target.props.onDragleave(leaveEvent());
    await vue.nextTick();
    assert.equal(target.props.style?.outline, undefined, "the PDF outline must not stick");

    const pdfBytes = new TextEncoder().encode("%PDF-1.7\n%%EOF");
    target.props.onDrop(dropEvent([sampleFile("dropped.pdf", pdfBytes, "application/pdf")]));
    await waitFor(
      () => fake.calls.openDocument.length === 2 && viewerState() === "ready",
      "dropped PDF",
      10_000,
    );
    assert.deepEqual(dropped, ["dropped.pdf"], "the drop is reported to the host");
    assert.equal(target.props.style?.outline, undefined, "the drop clears the outline");

    target.props.onDrop(dropEvent([sampleFile("notes.txt")]));
    await waitFor(
      () => findByTestId(mounted.root, "pdf-drop-error"),
      "unsupported-drop message",
    );
    assert.match(
      textContent(findByTestId(mounted.root, "pdf-drop-error")),
      /Only PDF files are supported/,
    );
    assert.equal(fake.calls.openDocument.length, 2, "the rejected file is not loaded");
    assert.deepEqual(dropped, ["dropped.pdf"]);
  } finally {
    mounted.app.unmount();
  }
});

// ── 4. PPTX viewer ─────────────────────────────────────────────────────────

function pptxDocumentWithSlides(count) {
  return {
    fileName: "preview.pptx",
    height: 540,
    width: 960,
    slides: Array.from({ length: count }, (_, index) => ({
      hidden: false,
      index,
      number: index + 1,
    })),
    warnings: [],
  };
}

function createPptxSession({ document = pptxDocumentWithSlides(2) } = {}) {
  const calls = { dispose: 0, open: [], renderSlide: [], renderThumbnail: [] };
  let currentSlideIndex = 0;
  const session = {
    document: null,
    get currentSlideIndex() { return currentSlideIndex; },
    get zoomPercent() { return 100; },
    async open(source, options) {
      calls.open.push({ options, source });
      currentSlideIndex = options?.initialSlide ?? 0;
      return document;
    },
    async renderSlide(index) {
      calls.renderSlide.push(index);
      currentSlideIndex = index;
    },
    renderThumbnail(index, target, width) {
      calls.renderThumbnail.push({ index, target, width });
      return null;
    },
    searchText() { return []; },
    async highlightSearchResult() { return null; },
    clearSearchHighlights() {},
    async setZoom() {},
    cancel() {},
    dispose() { calls.dispose += 1; },
  };
  return { calls, session };
}

test("dropping a presentation on PptxViewer opens it and rejects other files", async () => {
  const fake = createPptxSession();
  const dropped = [];
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([1, 2, 3]),
    sessionFactory: () => fake.session,
    showSearch: false,
    showSidebar: false,
    onFileDrop: (file) => dropped.push(file.name),
  });
  const target = dropTargetOf(mounted, "pptx-viewer");

  try {
    await waitFor(
      () => target.props["data-state"] === "ready",
      "initial PPTX",
      10_000,
    );
    assert.equal(fake.calls.open.length, 1);

    target.props.onDragover(dropEvent([]));
    await vue.nextTick();
    assert.equal(target.props.style.outline, FILE_DROP_OUTLINE);
    assert.equal(target.props["data-file-drop-active"], "true");

    target.props.onDragleave(leaveEvent());
    await vue.nextTick();
    assert.equal(target.props.style?.outline, undefined, "the PPTX outline must not stick");

    target.props.onDrop(dropEvent([sampleFile("dropped.pptx")]));
    await waitFor(
      () => fake.calls.open.length === 2,
      "dropped presentation",
      10_000,
    );
    assert.deepEqual(dropped, ["dropped.pptx"], "the drop is reported to the host");
    assert.ok(fake.calls.open[1].source instanceof File, "the dropped File is the new source");
    assert.equal(target.props.style?.outline, undefined, "the drop clears the outline");

    target.props.onDrop(dropEvent([sampleFile("notes.txt")]));
    await waitFor(
      () => findByTestId(mounted.root, "pptx-drop-error"),
      "unsupported-drop message",
    );
    assert.match(
      textContent(findByTestId(mounted.root, "pptx-drop-error")),
      /Only \.pptx, \.pptm, \.ppsx, \.ppsm, \.potx and \.potm files are supported/,
    );
    assert.equal(fake.calls.open.length, 2, "the rejected file is not opened");
    assert.deepEqual(dropped, ["dropped.pptx"]);
  } finally {
    mounted.app.unmount();
  }
});
