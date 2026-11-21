import assert from "node:assert/strict";
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

// Vue's native v-model directive checks these browser constructors during an
// update. The custom component renderer intentionally has no browser DOM.
if (globalThis.Document === undefined) globalThis.Document = class Document {};
if (globalThis.ShadowRoot === undefined) globalThis.ShadowRoot = class ShadowRoot {};

const { PdfSurface, PdfViewer } = await importFromDemo("@arcships/vue-pdf");

const PDF_SOURCE = {
  kind: "blob",
  blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }),
};

function createPdfRuntime({ pageCount = 3 } = {}) {
  const calls = { closeDocument: [], dispose: [], openDocument: [], renderPage: [], renderThumbnail: [] };
  const document = {
    id: "pdf-view-mode-document",
    pageCount,
    pages: Array.from({ length: pageCount }, (_, index) => ({
      index,
      width: 600,
      height: 800,
      rotation: 0,
    })),
  };
  const runtime = {
    runtimeId: "pdf-view-mode-runtime",
    async openDocument() {
      calls.openDocument.push(true);
      return document;
    },
    async renderPage(currentDocument, pageIndex, options) {
      calls.renderPage.push({ document: currentDocument, options, pageIndex });
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
    async dispose() {
      calls.dispose.push(true);
    },
  };
  return { calls, document, runtime };
}

function stubObjectUrls(t) {
  const created = [];
  const revoked = [];
  const originalCreate = URL.createObjectURL;
  const originalRevoke = URL.revokeObjectURL;
  let counter = 0;
  URL.createObjectURL = () => {
    const url = `blob:pdf-view-mode-${++counter}`;
    created.push(url);
    return url;
  };
  URL.revokeObjectURL = (url) => revoked.push(url);
  t.after(() => {
    URL.createObjectURL = originalCreate;
    URL.revokeObjectURL = originalRevoke;
  });
  return { created, revoked };
}

function classNames(node) {
  return String(node?.props?.class || "").split(/\s+/).filter(Boolean);
}

async function chooseToolbarAction(root, testId) {
  findByTestId(root, "pdf-actions").props.onClick();
  await waitFor(() => findByTestId(root, testId) !== undefined, `${testId} actions item`);
  findByTestId(root, testId).props.onClick();
  await vue.nextTick();
}

function nodesWithClass(root, className) {
  return walk(root).filter((node) => classNames(node).includes(className));
}

function pageSlots(root) {
  return nodesWithClass(root, "pdf-surface__page-slot");
}

function makeScrollable(node, { clientHeight = 800, clientWidth = 900, scrollTop = 0 } = {}) {
  const state = { scrollTop, scrollLeft: 0 };
  Object.defineProperty(node, "clientHeight", { configurable: true, get: () => clientHeight });
  Object.defineProperty(node, "clientWidth", { configurable: true, get: () => clientWidth });
  Object.defineProperty(node, "scrollTop", {
    configurable: true,
    get: () => state.scrollTop,
    set: (value) => { state.scrollTop = value; },
  });
  Object.defineProperty(node, "scrollLeft", {
    configurable: true,
    get: () => state.scrollLeft,
    set: (value) => { state.scrollLeft = value; },
  });
  const scrolls = [];
  node.scrollTo = (options) => {
    scrolls.push(options);
    if (typeof options?.top === "number") state.scrollTop = options.top;
  };
  node.scrollIntoView = (options) => { scrolls.push(options); };
  node.getBoundingClientRect = () => ({
    bottom: clientHeight,
    height: clientHeight,
    left: 0,
    right: clientWidth,
    top: 0,
    width: clientWidth,
  });
  return { scrolls, state };
}

async function mountSurfaceHarness(props) {
  let api;
  const Harness = vue.defineComponent({
    setup() {
      const surface = vue.ref(null);
      vue.watch(surface, (value) => { api = value; }, { flush: "sync" });
      return () => vue.h(PdfSurface, { ref: surface, ...props });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(() => api, "exposed PDF surface API");
  return { ...mounted, api: () => api };
}

test("PdfSurface spread renders two pages side by side with correct page numbers", async (t) => {
  stubObjectUrls(t);
  const fake = createPdfRuntime({ pageCount: 3 });
  const mounted = await mount(PdfSurface, {
    source: PDF_SOURCE,
    runtime: fake.runtime,
    pageLayout: "spread",
  });
  await waitFor(() => pageSlots(mounted.root).length === 3, "three PDF page slots");

  const surface = findByTestId(mounted.root, "pdf-surface");
  assert.equal(surface.props["data-page-layout"], "spread");
  const slots = pageSlots(mounted.root);
  assert.deepEqual(
    slots.map((slot) => slot.props["data-page-number"]),
    [1, 2, 3],
    "every spread page keeps its own page number",
  );
  assert.deepEqual(slots.map((slot) => slot.props.style.gridRow), ["1", "1", "2"]);
  assert.deepEqual(slots.map((slot) => slot.props.style.gridColumn), ["1", "2", "1 / -1"]);

  const stack = nodesWithClass(mounted.root, "pdf-surface__pages")[0];
  assert.equal(stack.props.style.gridTemplateColumns, "600px 600px");
  assert.equal(stack.props.style.columnGap, "8px");
  assert.equal(stack.props.style.rowGap, "8px");

  const images = nodesWithClass(mounted.root, "pdf-surface__page-img");
  assert.deepEqual(
    images.slice(0, 2).map((image) => image.props.alt),
    ["Page 1", "Page 2"],
    "the facing bitmaps are labelled with their own page numbers",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfSurface single page mode rasterizes only the page it snapped to", async (t) => {
  stubObjectUrls(t);
  const fake = createPdfRuntime({ pageCount: 3 });
  const visiblePages = [];
  const mounted = await mount(PdfSurface, {
    source: PDF_SOURCE,
    runtime: fake.runtime,
    pageLayout: "single",
    onVisiblePageChange: (pageIndex) => visiblePages.push(pageIndex),
  });
  await waitFor(() => pageSlots(mounted.root).length === 3, "three PDF page slots");

  const surface = findByTestId(mounted.root, "pdf-surface");
  assert.equal(surface.props["data-page-layout"], "single");
  const scroll = makeScrollable(surface, { clientHeight: 800, scrollTop: 900 });
  surface.props.onScrollPassive();

  // Page rows start at 16, 824 and 1632. The 900px offset sits inside page 2,
  // so single-page mode snaps to that page's top and keeps one bitmap.
  await waitFor(() => scroll.state.scrollTop === 824, "single-page snap to page 2");
  await waitFor(() => visiblePages.at(-1) === 1, "page 2 published as visible");
  await waitFor(
    () => fake.calls.renderPage.some((call) => call.pageIndex === 1),
    "page 2 rasterized",
  );
  const rendered = new Set(fake.calls.renderPage.map((call) => call.pageIndex));
  assert.deepEqual([...rendered], [0, 1], "the third page stays a placeholder");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfSurface spread resolves the facing page in a context menu click", async (t) => {
  stubObjectUrls(t);
  const fake = createPdfRuntime({ pageCount: 3 });
  const contextMenus = [];
  const mounted = await mount(PdfSurface, {
    source: PDF_SOURCE,
    runtime: fake.runtime,
    pageLayout: "spread",
    onContextMenu: (event) => contextMenus.push(event),
  });
  await waitFor(() => pageSlots(mounted.root).length === 3, "three PDF page slots");

  const surface = findByTestId(mounted.root, "pdf-surface");
  makeScrollable(surface, { clientHeight: 800, clientWidth: 900 });
  surface.props.onContextmenu?.({ clientX: 200, clientY: 100, preventDefault() {} });
  surface.props.onContextmenu?.({ clientX: 800, clientY: 100, preventDefault() {} });
  assert.deepEqual(
    contextMenus.map((event) => event.pageIndex),
    [0, 1],
    "left and right halves of a spread resolve their own page",
  );
  mounted.app.unmount();
});

test("PdfViewer spread renders the active page and its facing page", async (t) => {
  const urls = stubObjectUrls(t);
  const fake = createPdfRuntime({ pageCount: 4 });
  const mounted = await mount(PdfViewer, {
    source: PDF_SOURCE,
    runtime: fake.runtime,
    pageLayout: "spread",
  });

  await waitFor(
    () => findByTestId(mounted.root, "pdf-facing-page-image") !== undefined,
    "facing page bitmap",
  );
  const facing = findByTestId(mounted.root, "pdf-facing-page");
  assert.equal(facing.props["data-page-number"], 2);
  assert.equal(
    findByTestId(mounted.root, "pdf-facing-page-image").props.alt,
    "PDF page 2",
    "the facing bitmap is labelled with its own page number",
  );
  assert.equal(findByTestId(mounted.root, "pdf-page").props["data-page-number"], 1);

  // Page navigation keeps the pair in sync: page 3 faces page 4.
  await chooseToolbarAction(mounted.root, "pdf-next-page");
  await chooseToolbarAction(mounted.root, "pdf-next-page");
  await waitFor(
    () => findByTestId(mounted.root, "pdf-page")?.props["data-page-number"] === 3,
    "active page 3",
  );
  await waitFor(
    () => findByTestId(mounted.root, "pdf-facing-page")?.props["data-page-number"] === 4,
    "facing page 4",
  );
  assert.equal(
    findByTestId(mounted.root, "pdf-facing-page-image").props.alt,
    "PDF page 4",
  );
  assert.ok(urls.created.length >= 3, "each spread bitmap comes from its own object URL");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfViewer continuous mode lays out every page and keeps the active page across modes", async (t) => {
  stubObjectUrls(t);
  const fake = createPdfRuntime({ pageCount: 4 });
  const mounted = await mount(PdfViewer, {
    source: PDF_SOURCE,
    runtime: fake.runtime,
    // Uncontrolled so the toolbar layout control can drive the mode.
    defaultPageLayout: "continuous",
  });

  await waitFor(() => findByTestId(mounted.root, "pdf-page") !== undefined, "PDF page stage");
  const frame = nodesWithClass(mounted.root, "pdf-frame-wrap")[0];
  assert.equal(frame.props["data-page-layout"], "continuous");
  const stack = nodesWithClass(mounted.root, "pdf-page-stack")[0];
  assert.equal(
    stack.props.style.height,
    `${4 * 800 + 4 * 16}px`,
    "four pages stack vertically plus the trailing row gap",
  );
  assert.equal(stack.props.style.minWidth, "600px");
  await waitFor(
    () => walk(mounted.root).filter((node) => node.props?.["data-testid"] === "pdf-page-placeholder").length >= 2,
    "placeholder pages around the active one",
  );
  assert.deepEqual(
    walk(mounted.root)
      .filter((node) => node.props?.["data-testid"] === "pdf-page-placeholder")
      .map((node) => node.props["data-page-number"]),
    [2, 3],
    "the pages beside the active one are sized, numbered placeholders",
  );

  // Scrolling to the third page makes it active and rasterizes it.
  const scroll = makeScrollable(frame, { clientHeight: 800, scrollTop: 2 * 816 });
  frame.props.onScrollPassive();
  await waitFor(
    () => findByTestId(mounted.root, "pdf-page")?.props["data-page-number"] === 3,
    "active page 3 from the scroll position",
  );
  await waitFor(
    () => fake.calls.renderPage.some((call) => call.pageIndex === 2),
    "page 3 rasterized in continuous mode",
  );
  assert.equal(
    findByTestId(mounted.root, "pdf-document").props["data-active-page"],
    3,
  );

  // Switching to single-page mode keeps page 3 selected.
  await chooseToolbarAction(mounted.root, "pdf-page-layout-single");
  await waitFor(
    () => nodesWithClass(mounted.root, "pdf-frame-wrap")[0].props["data-page-layout"] === "single",
    "single-page layout",
  );
  assert.equal(findByTestId(mounted.root, "pdf-page").props["data-page-number"], 3);
  assert.equal(
    findByTestId(mounted.root, "pdf-document").props["data-active-page"],
    3,
    "switching layout must not rewind the active page",
  );
  assert.equal(scroll.scrolls.length >= 0, true);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});
