import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  keydownEvent,
  mount,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

// Vue's native v-model directive checks these browser constructors during an
// update. The custom component renderer intentionally has no browser DOM.
if (globalThis.Document === undefined) globalThis.Document = class Document {};
if (globalThis.ShadowRoot === undefined) globalThis.ShadowRoot = class ShadowRoot {};

const { DocxViewer } = await importFromDemo("@arcships/vue-docx");
const { PdfSurface, PdfViewer } = await importFromDemo("@arcships/vue-pdf");
const { cloneDocModel, defaultStarterModel } = await importFromDemo("@arcships/docx-core");

const VIEWER_LAYOUT = {
  pageWidth: 200,
  pageHeight: 300,
  margin: 10,
  minLineHeight: 30,
  paragraphSpacing: 10,
};

function keydown(node, key, init = {}) {
  const { event, preventedCount } = keydownEvent(key, init);
  const handlers = Array.isArray(node.props?.onKeydown)
    ? node.props.onKeydown
    : [node.props?.onKeydown];
  for (const handler of handlers) handler?.(event);
  return { event, prevented: preventedCount() };
}

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function paginatedDocxModel() {
  const model = cloneDocModel(defaultStarterModel);
  model.nodes = Array.from({ length: 30 }, (_, index) => paragraph(`Viewer item ${index + 1}`));
  model.metadata.sectionPropertiesXml =
    '<w:sectPr><w:pgSz w:w="3000" w:h="4500"/>' +
    '<w:pgMar w:top="150" w:right="150" w:bottom="150" w:left="150"/></w:sectPr>';
  return model;
}

async function waitForDocxPages(root, total) {
  await waitFor(
    () => walk(root).some((node) => node.props?.["data-docx-page-count"] === total),
    `DOCX ${total}-page layout`,
  );
}

function createPdfRuntime({ pageCount = 3 } = {}) {
  const document = {
    id: `keyboard-pdf-${pageCount}`,
    pageCount,
    pages: Array.from({ length: pageCount }, (_, index) => ({
      index,
      width: 600,
      height: 800,
      rotation: 0,
    })),
  };
  return {
    runtimeId: "keyboard-pdf-runtime",
    async openDocument() { return document },
    async renderPage() { return new Blob(["page"], { type: "image/png" }) },
    async renderThumbnail() { return new Blob(["thumb"], { type: "image/png" }) },
    async getPageTextGeometry() { return { runs: [] } },
    async getTextSlices(_document, slices) { return slices.map(() => "") },
    async search() { return [] },
    async closeDocument() {},
    async dispose() {},
  };
}

const PDF_SOURCE = {
  kind: "blob",
  blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }),
};

async function mountPdfViewer(extra = {}) {
  const mounted = await mount(PdfViewer, {
    source: PDF_SOURCE,
    runtime: createPdfRuntime({ pageCount: 3 }),
    ...extra,
  });
  await waitFor(
    () => findByTestId(mounted.root, "pdf-status")?.props["data-state"] === "ready",
    "PDF viewer ready",
  );
  return mounted;
}

test("DocxViewer moves pages with PageUp/PageDown and Home/End", async () => {
  const mounted = await mount(DocxViewer, {
    model: paginatedDocxModel(),
    layoutOptions: VIEWER_LAYOUT,
  });
  await waitForDocxPages(mounted.root, 2);

  const viewer = findByTestId(mounted.root, "docx-viewer");
  const pageInput = () => findByTestId(mounted.root, "docx-page-current");
  assert.equal(pageInput().value, "1");

  const down = keydown(viewer, "PageDown");
  await vue.nextTick();
  assert.equal(pageInput().value, "2", "PageDown moves one page");
  assert.equal(down.prevented, 1, "the handled key is prevented");

  keydown(viewer, "PageUp");
  await vue.nextTick();
  assert.equal(pageInput().value, "1", "PageUp moves one page back");

  keydown(viewer, "End");
  await vue.nextTick();
  assert.equal(pageInput().value, "2", "End jumps to the last page");

  keydown(viewer, "Home");
  await vue.nextTick();
  assert.equal(pageInput().value, "1", "Home jumps to the first page");

  const boundary = keydown(viewer, "PageUp");
  await vue.nextTick();
  assert.equal(pageInput().value, "1");
  assert.equal(boundary.prevented, 0, "a boundary key stays with the browser");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("DocxViewer zooms with Ctrl/Cmd +/-/0 and focuses search with Ctrl/Cmd+F", async () => {
  const mounted = await mount(DocxViewer, {
    model: paginatedDocxModel(),
    layoutOptions: VIEWER_LAYOUT,
  });
  await waitForDocxPages(mounted.root, 2);

  const viewer = findByTestId(mounted.root, "docx-viewer");
  const zoom = () => findByTestId(mounted.root, "docx-zoom-select").props.value;
  assert.equal(zooms(mounted.root), 100);

  keydown(viewer, "+", { ctrlKey: true });
  await vue.nextTick();
  assert.equal(zooms(mounted.root), 125, "Ctrl++ zooms in");

  keydown(viewer, "=", { metaKey: true });
  await vue.nextTick();
  assert.equal(zooms(mounted.root), 150, "Cmd+= zooms in");

  keydown(viewer, "-", { ctrlKey: true });
  await vue.nextTick();
  assert.equal(zooms(mounted.root), 125, "Ctrl+- zooms out");

  keydown(viewer, "0", { ctrlKey: true });
  await vue.nextTick();
  assert.equal(zooms(mounted.root), 100, "Ctrl+0 resets the zoom");

  // The default row hides the find field (upstream keeps find in the actions
  // menu); Ctrl/Cmd+F opens the same find popover and focuses its input.
  assert.equal(
    findByTestId(mounted.root, "docx-search-input"),
    undefined,
    "the default toolbar row must not render an inline find field",
  );
  keydown(viewer, "f", { ctrlKey: true });
  await vue.nextTick();
  await vue.nextTick();
  let search = findByTestId(mounted.root, "docx-search-input");
  assert.ok(search, "Ctrl+F opens the find popover");
  let focused = 0;
  search.focus = () => { focused += 1; };
  keydown(viewer, "f", { ctrlKey: true });
  await vue.nextTick();
  assert.ok(focused >= 1, "Ctrl+F focuses the find field");

  search = findByTestId(mounted.root, "docx-search-input");
  search.value = "needle";
  search.props.onInput({ target: search });
  await vue.nextTick();
  keydown(viewer, "Escape");
  await vue.nextTick();
  assert.equal(
    findByTestId(mounted.root, "docx-search-input").value,
    "",
    "Escape cancels the active search",
  );

  assert.equal(zoom(), 100);
  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

function zooms(root) {
  return Number(findByTestId(root, "docx-zoom-select").props.value);
}

test("viewer shortcuts do not fire from editable fields or in a sibling viewer", async () => {
  const first = await mount(DocxViewer, {
    model: paginatedDocxModel(),
    layoutOptions: VIEWER_LAYOUT,
  });
  const second = await mount(DocxViewer, {
    model: paginatedDocxModel(),
    layoutOptions: VIEWER_LAYOUT,
  });
  await waitForDocxPages(first.root, 2);
  await waitForDocxPages(second.root, 2);

  const firstViewer = findByTestId(first.root, "docx-viewer");
  const firstInput = () => findByTestId(first.root, "docx-page-current");
  const secondInput = () => findByTestId(second.root, "docx-page-current");

  keydown(firstViewer, "PageDown", { target: { tagName: "INPUT" } });
  await vue.nextTick();
  assert.equal(firstInput().value, "1", "PageDown inside a text field is left untouched");
  assert.equal(secondInput().value, "1", "a sibling viewer does not receive the key");

  keydown(firstViewer, "PageDown");
  await vue.nextTick();
  assert.equal(firstInput().value, "2");
  assert.equal(secondInput().value, "1", "shortcuts stay scoped to the focused viewer root");

  first.app.unmount();
  second.app.unmount();
  assert.deepEqual(first.warnings, []);
  assert.deepEqual(second.warnings, []);
});

test("PdfViewer moves pages, zooms, searches and cancels search with Escape", async () => {
  const activePages = [];
  const mounted = await mountPdfViewer({
    defaultZoom: 1,
    onActivePageChange: (page) => activePages.push(page),
  });
  const viewer = findByTestId(mounted.root, "pdf-viewer");
  const pageInput = () => findByTestId(mounted.root, "pdf-page-input");
  const document = () => findByTestId(mounted.root, "pdf-document");
  assert.equal(pageInput().value, "1");

  const down = keydown(viewer, "PageDown");
  await vue.nextTick();
  assert.equal(pageInput().value, "2", "PageDown moves one page");
  assert.equal(document().props["data-active-page"], 2);
  assert.deepEqual(activePages.at(-1), 2, "active-page-change follows the keyboard page");
  assert.equal(down.prevented, 1);

  keydown(viewer, "End");
  await vue.nextTick();
  assert.equal(pageInput().value, "3", "End jumps to the last page");

  keydown(viewer, "Home");
  await vue.nextTick();
  assert.equal(pageInput().value, "1", "Home jumps to the first page");

  keydown(viewer, "PageUp");
  await vue.nextTick();
  assert.equal(pageInput().value, "1", "PageUp stops at the first page");

  keydown(viewer, "=", { ctrlKey: true });
  await vue.nextTick();
  assert.equal(document().props["data-zoom"], 1.25, "Ctrl+= zooms in");

  keydown(viewer, "-", { ctrlKey: true });
  await vue.nextTick();
  assert.equal(document().props["data-zoom"], 1, "Ctrl+- zooms out");

  keydown(viewer, "=", { metaKey: true });
  await vue.nextTick();
  assert.equal(document().props["data-zoom"], 1.25, "Cmd+= zooms in");

  keydown(viewer, "0", { metaKey: true });
  await vue.nextTick();
  assert.equal(document().props["data-zoom"], 1, "Cmd+0 resets the zoom");

  keydown(viewer, "f", { ctrlKey: true });
  await vue.nextTick();
  const search = findByTestId(mounted.root, "pdf-search-input");
  assert.ok(search, "Ctrl+F opens the PDF search field");
  let focused = 0;
  search.focus = () => { focused += 1; };
  keydown(viewer, "f", { ctrlKey: true });
  await vue.nextTick();
  assert.equal(focused, 1, "Ctrl+F focuses the open PDF search field");

  search.value = "needle";
  search.dispatchEvent({ type: "input", target: search });
  await vue.nextTick();
  keydown(viewer, "Escape");
  await vue.nextTick();
  assert.equal(
    findByTestId(mounted.root, "pdf-search-input"),
    undefined,
    "Escape closes the PDF search field",
  );
  keydown(viewer, "f", { ctrlKey: true });
  await vue.nextTick();
  assert.equal(
    findByTestId(mounted.root, "pdf-search-input").value,
    "",
    "Escape cancels the PDF search",
  );

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("PdfViewer Ctrl/Cmd+P follows the editable-target guard but still prints outside fields", async () => {
  const previousWindow = globalThis.window;
  let printed = 0;
  globalThis.window = {
    print: () => { printed += 1; },
  };
  try {
    const mounted = await mountPdfViewer();
    const viewer = findByTestId(mounted.root, "pdf-viewer");
    findByTestId(mounted.root, "pdf-search-toggle").props.onClick();
    await vue.nextTick();
    const search = findByTestId(mounted.root, "pdf-search-input");
    const pageInput = findByTestId(mounted.root, "pdf-page-input");

    const fromSearch = keydown(viewer, "p", { ctrlKey: true, target: search });
    const fromPageInput = keydown(viewer, "p", { metaKey: true, target: pageInput });
    const alreadyHandled = keydown(viewer, "p", {
      ctrlKey: true,
      target: viewer,
      defaultPrevented: true,
    });
    await new Promise((resolve) => setTimeout(resolve, 20));

    assert.equal(fromSearch.prevented, 0, "Ctrl+P in the search field is not prevented");
    assert.equal(fromPageInput.prevented, 0, "Cmd+P in the page field is not prevented");
    assert.equal(alreadyHandled.prevented, 0, "a child-handled Ctrl+P is not re-handled");
    assert.equal(printed, 0, "printing must not start from an editable field");

    const outside = keydown(viewer, "p", { ctrlKey: true, target: viewer });
    assert.equal(outside.prevented, 1, "Ctrl+P outside fields is handled");
    await waitFor(() => printed === 1, "PDF print invocation");

    mounted.app.unmount();
    assert.deepEqual(mounted.warnings, []);
  } finally {
    if (previousWindow === undefined) delete globalThis.window;
    else globalThis.window = previousWindow;
  }
});

test("PdfSurface pages with PageUp/PageDown and Home/End and zooms with Ctrl keys", async () => {
  const zoomState = vue.ref(1);
  const cancels = [];
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(PdfSurface, {
        source: PDF_SOURCE,
        runtime: createPdfRuntime({ pageCount: 4 }),
        selectionMode: "region",
        zoom: zoomState.value,
        "onUpdate:zoom": (value) => { zoomState.value = value; },
        onSelectionCancel: (event) => cancels.push(event),
      });
    },
  });
  const mounted = await mount(Harness);
  const surface = findByTestId(mounted.root, "pdf-surface");
  await waitFor(
    () => walk(mounted.root).some((node) => node.props?.class === "pdf-surface__pages"),
    "PDF surface pages",
  );

  const scrolled = [];
  surface.querySelector = (selector) => ({
    scrollIntoView: (options) => scrolled.push({ selector, options }),
  });
  surface.clientHeight = 600;

  keydown(surface, "PageDown");
  await vue.nextTick();
  assert.deepEqual(scrolled.at(-1)?.selector, '[data-page-index="1"]', "PageDown moves one page");

  keydown(surface, "End");
  await vue.nextTick();
  assert.deepEqual(scrolled.at(-1)?.selector, '[data-page-index="3"]', "End jumps to the last page");

  keydown(surface, "Home");
  await vue.nextTick();
  assert.deepEqual(scrolled.at(-1)?.selector, '[data-page-index="0"]', "Home jumps to the first page");

  keydown(surface, "=", { ctrlKey: true });
  await vue.nextTick();
  assert.equal(zoomState.value, 1.25, "Ctrl+= zooms the surface in");
  keydown(surface, "-", { ctrlKey: true });
  await vue.nextTick();
  assert.equal(zoomState.value, 1, "Ctrl+- zooms the surface out");
  keydown(surface, "=", { ctrlKey: true });
  await vue.nextTick();
  keydown(surface, "0", { ctrlKey: true });
  await vue.nextTick();
  assert.equal(zoomState.value, 1, "Ctrl+0 resets the surface zoom");

  const before = scrolled.length;
  keydown(surface, "PageDown", { target: { tagName: "INPUT" } });
  await vue.nextTick();
  assert.equal(scrolled.length, before, "surface page keys do not fire inside inputs");

  keydown(surface, "Escape");
  assert.deepEqual(cancels, [{ mode: "region", reason: "escape" }], "Escape cancels the region");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("viewer roots expose a visible keyboard focus ring", () => {
  const source = (path) => readFileSync(new URL(`../../${path}`, import.meta.url), "utf8");
  assert.match(source("packages/vue-docx/src/components/DocxViewer.vue"), /\.docx-viewer:focus-visible\s*\{[^}]*outline:/);
  assert.match(source("packages/vue-pdf/src/components/PdfViewer.vue"), /\.pdf-viewer:focus-visible\s*\{[^}]*outline:/);
  assert.match(source("packages/vue-pdf/src/components/PdfSurface.vue"), /\.pdf-surface:focus-visible\s*\{[^}]*outline:/);
  assert.match(source("packages/vue-pptx/src/style.css"), /\.pptx-viewer:focus-visible\s*\{[^}]*outline:/);
  assert.match(source("packages/vue-xlsx/src/components/XlsxViewer.vue"), /\.xlsx-viewer:focus-visible\s*\{[^}]*outline:/);
});
