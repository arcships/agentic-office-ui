import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const { PdfSurface, PdfViewer } = await importFromDemo("@arcships/vue-pdf");

// Vue's native v-model directive checks these browser constructors during an
// update. The custom component renderer intentionally has no browser DOM.
if (globalThis.Document === undefined) globalThis.Document = class Document {};
if (globalThis.ShadowRoot === undefined) globalThis.ShadowRoot = class ShadowRoot {};

const PDF_SOURCE = { kind: "blob", blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }) };

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((nextResolve, nextReject) => {
    resolve = nextResolve;
    reject = nextReject;
  });
  return { promise, reject, resolve };
}

function createPdfRuntime({ pageCount = 1, onGetPageTextLayer, onRenderPage, rotations = [] } = {}) {
  const calls = { closeDocument: [], dispose: [], getPageTextLayer: [], openDocument: [], renderPage: [] };
  const document = {
    id: "pdf-lifecycle-document",
    pageCount,
    pages: Array.from({ length: pageCount }, (_, index) => ({
      index,
      width: 600,
      height: 800,
      rotation: rotations[index] ?? 0,
    })),
  };
  const runtime = {
    runtimeId: "pdf-lifecycle-runtime",
    async openDocument() {
      calls.openDocument.push(true);
      return document;
    },
    async renderPage(currentDocument, pageIndex, options) {
      calls.renderPage.push({ document: currentDocument, options, pageIndex });
      return onRenderPage
        ? onRenderPage(currentDocument, pageIndex, options)
        : new Blob([`page:${pageIndex}`], { type: "image/png" });
    },
    async renderThumbnail() {
      return new Blob(["thumbnail"], { type: "image/png" });
    },
    async getPageTextGeometry() {
      return { runs: [] };
    },
    ...(onGetPageTextLayer
      ? {
          async getPageTextLayer(_currentDocument, pageIndex) {
            calls.getPageTextLayer.push(pageIndex);
            return onGetPageTextLayer(pageIndex);
          },
        }
      : {}),
    async getTextSlices(_currentDocument, slices) {
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
    const url = `blob:pdf-lifecycle-${++counter}`;
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

/** Fake `window.matchMedia` that reports display-scale changes the way a monitor swap does. */
function stubDisplayScale(t, initialDpr) {
  const listeners = new Set();
  const queries = [];
  const previousWindow = globalThis.window;
  const previousDpr = globalThis.devicePixelRatio;
  const media = {
    addEventListener(type, listener) {
      if (type === "change") listeners.add(listener);
    },
    removeEventListener(type, listener) {
      if (type === "change") listeners.delete(listener);
    },
  };
  globalThis.devicePixelRatio = initialDpr;
  globalThis.window = {
    matchMedia(query) {
      if (!query.startsWith("(resolution:")) {
        return {
          matches: false,
          media: query,
          addEventListener() {},
          removeEventListener() {},
        };
      }
      queries.push(query);
      return { matches: true, media: query, ...media };
    },
  };
  t.after(() => {
    if (previousWindow === undefined) delete globalThis.window;
    else globalThis.window = previousWindow;
    if (previousDpr === undefined) delete globalThis.devicePixelRatio;
    else globalThis.devicePixelRatio = previousDpr;
  });
  return {
    queries,
    change(nextDpr) {
      globalThis.devicePixelRatio = nextDpr;
      for (const listener of [...listeners]) listener({ matches: false, media: `(resolution: ${nextDpr}dppx)` });
    },
  };
}

function stubResizeObserver(t) {
  const observers = [];
  const previous = globalThis.ResizeObserver;
  globalThis.ResizeObserver = class ResizeObserver {
    constructor(callback) {
      this.callback = callback;
      observers.push(this);
    }
    observe() {}
    disconnect() {
      this.disconnected = true;
    }
  };
  t.after(() => {
    if (previous === undefined) delete globalThis.ResizeObserver;
    else globalThis.ResizeObserver = previous;
  });
  return observers;
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

function surfacePageImages(root) {
  return walk(root).filter((node) => node.props?.class === "pdf-surface__page-img");
}

test("PdfSurface swaps page bitmaps in place and revokes replaced object URLs on zoom", async (t) => {
  const urls = stubObjectUrls(t);
  const fake = createPdfRuntime({ pageCount: 2 });
  const mounted = await mountSurfaceHarness({ source: PDF_SOURCE, runtime: fake.runtime });

  await waitFor(() => surfacePageImages(mounted.root).length === 2, "two rendered PDF pages");
  const firstImages = surfacePageImages(mounted.root);
  const firstSources = firstImages.map((node) => node.props.src);
  assert.deepEqual(firstSources, urls.created.slice(0, 2));
  assert.deepEqual(urls.revoked, []);

  mounted.api().zoom = 2;
  await waitFor(
    () => surfacePageImages(mounted.root).every((node, index) => node.props.src !== firstSources[index]),
    "re-rendered PDF pages",
  );

  const secondImages = surfacePageImages(mounted.root);
  assert.deepEqual(
    secondImages.map((node) => node.props.src),
    urls.created.slice(2, 4),
    "each page swaps to the freshly rendered bitmap",
  );
  for (const [index, node] of secondImages.entries()) {
    assert.equal(node, firstImages[index], "the page <img> is patched in place instead of remounted");
  }
  assert.deepEqual(urls.revoked, firstSources, "replaced bitmaps are revoked before the swap");
  assert.deepEqual(
    fake.calls.renderPage.filter((call) => call.options.scale === 2).map((call) => call.pageIndex),
    [0, 1],
    "the visible window re-renders at the new zoom instead of reusing the stale slots",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
  await waitFor(() => urls.revoked.length >= 4, "bitmaps revoked on unmount");
});

test("PdfSurface maps scroll and context-menu positions in scroll space and re-measures per layout", async (t) => {
  const fake = createPdfRuntime({ pageCount: 3 });
  const visiblePages = [];
  const contextMenus = [];
  const mounted = await mountSurfaceHarness({
    source: PDF_SOURCE,
    runtime: fake.runtime,
    onVisiblePageChange: (pageIndex) => visiblePages.push(pageIndex),
    onContextMenu: (event) => contextMenus.push(event),
  });
  await waitFor(() => surfacePageImages(mounted.root).length === 3, "three rasterized PDF pages");

  const surface = findByTestId(mounted.root, "pdf-surface");
  // The stack is 3 pages of 600x800 at zoom 1 with an 8px gap and 16px padding:
  // page tops are 16, 824 and 1632 inside the scroll content.
  let scrollTop = 700;
  surface.clientHeight = 800;
  // The surface itself sits 400px down the host page. Only a scroll-relative
  // computation can be right here; `offsetTop` would fold this offset in.
  const surfaceTop = 400;
  surface.getBoundingClientRect = () => ({ bottom: surfaceTop + 800, height: 800, left: 0, right: 900, top: surfaceTop, width: 900 });
  Object.defineProperty(surface, "scrollTop", {
    configurable: true,
    get: () => scrollTop,
    set: (value) => { scrollTop = value },
  });

  surface.props.onScrollPassive();
  await waitFor(() => visiblePages.at(-1) === 1, "visible page inside the second page");

  scrollTop = 1200;
  surface.props.onScrollPassive();
  assert.equal(visiblePages.at(-1), 1, "the viewport center stays on the second page");

  surface.props.onContextmenu({ clientX: 300, clientY: 600, preventDefault() {} });
  assert.equal(contextMenus.at(-1).pageIndex, 1, "the context menu resolves the page under the pointer");
  assert.equal(contextMenus.at(-1).containerY, 600 - surfaceTop);

  // Zooming doubles every page height, so the same 700px scroll offset now
  // lands inside the first page.
  mounted.api().zoom = 2;
  await vue.nextTick();
  await waitFor(() => mounted.api().zoom === 2, "zoom applied");
  surface.props.onScrollPassive();
  await waitFor(() => visiblePages.at(-1) === 0, "re-measured page offsets after the layout change");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfSurface fits width to the rotated page edge", async (t) => {
  const observers = stubResizeObserver(t);
  const fake = createPdfRuntime({ pageCount: 1 });
  const mounted = await mountSurfaceHarness({
    source: PDF_SOURCE,
    runtime: fake.runtime,
    defaultZoom: 0.5,
    fitWidth: true,
  });
  await waitFor(() => surfacePageImages(mounted.root).length === 1, "fit-width PDF page");
  assert.ok(observers.length >= 1, "fit width observes the surface width");

  const surface = findByTestId(mounted.root, "pdf-surface");
  surface.clientWidth = 632;
  await waitFor(() => mounted.api().zoom === 1, "portrait fit width");

  mounted.api().rotation = 90;
  await waitFor(() => mounted.api().zoom === 0.75, "fit width follows the rotated page edge");
  await waitFor(
    () => fake.calls.renderPage.at(-1)?.options.rotation === 90 && fake.calls.renderPage.at(-1)?.options.scale === 0.75,
    "rotated fit-width render",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfSurface re-renders pages when the display pixel ratio changes", async (t) => {
  const display = stubDisplayScale(t, 2);
  stubObjectUrls(t);
  const fake = createPdfRuntime({ pageCount: 1 });
  const mounted = await mountSurfaceHarness({ source: PDF_SOURCE, runtime: fake.runtime });

  await waitFor(() => fake.calls.renderPage.length === 1, "initial pixel ratio render");
  assert.equal(fake.calls.renderPage[0].options.dpr, 2);
  assert.deepEqual(display.queries, ["(resolution: 2dppx)"]);

  display.change(3);
  await waitFor(() => fake.calls.renderPage.length === 2, "re-render after the display scale changed");
  assert.equal(fake.calls.renderPage[1].options.dpr, 3, "3x displays are no longer down-sampled to 2x");
  assert.equal(display.queries.at(-1), "(resolution: 3dppx)");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfViewer keeps the current page bitmap mounted while the next page renders", async (t) => {
  stubObjectUrls(t);
  const gate = deferred();
  let renders = 0;
  const fake = createPdfRuntime({
    pageCount: 2,
    onRenderPage: () => {
      renders += 1;
      return renders === 1 ? new Blob(["page-1"], { type: "image/png" }) : gate.promise;
    },
  });
  const mounted = await mount(PdfViewer, { source: PDF_SOURCE, runtime: fake.runtime });
  await waitFor(() => findByTestId(mounted.root, "pdf-status")?.props["data-state"] === "ready", "first PDF page");
  const firstImage = findByTestId(mounted.root, "pdf-page-image");
  assert.ok(firstImage);
  const firstSource = firstImage.props.src;

  findByTestId(mounted.root, "pdf-next-page").props.onClick();
  await waitFor(
    () => findByTestId(mounted.root, "pdf-page")?.props["data-render-state"] === "rendering",
    "second page render start",
  );

  const duringRender = findByTestId(mounted.root, "pdf-page-image");
  assert.ok(duringRender, "the stage keeps the previous bitmap instead of collapsing to white");
  assert.equal(duringRender, firstImage, "the placeholder is the already decoded image");
  assert.equal(duringRender.props.src, firstSource);

  gate.resolve(new Blob(["page-2"], { type: "image/png" }));
  await waitFor(
    () => findByTestId(mounted.root, "pdf-page-image")?.props.src !== firstSource,
    "second page bitmap",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfViewer restores the scroll anchor after a zoom", async (t) => {
  stubObjectUrls(t);
  let pageWidth = 600;
  let pageHeight = 800;
  const fake = createPdfRuntime({
    pageCount: 1,
    onRenderPage: (_document, _pageIndex, options) => {
      pageWidth = 600 * options.scale;
      pageHeight = 800 * options.scale;
      return new Blob(["zoomed-page"], { type: "image/png" });
    },
  });
  const mounted = await mount(PdfViewer, { source: PDF_SOURCE, runtime: fake.runtime });
  await waitFor(() => findByTestId(mounted.root, "pdf-status")?.props["data-state"] === "ready", "rendered PDF page");

  const frame = walk(mounted.root).find((node) => node.props?.class === "pdf-frame-wrap");
  const image = findByTestId(mounted.root, "pdf-page-image");
  assert.ok(frame);
  assert.ok(image);
  frame.getBoundingClientRect = () => ({ bottom: 600, height: 600, left: 0, right: 800, top: 0, width: 800 });
  image.getBoundingClientRect = () => ({ bottom: 20 + pageHeight, height: pageHeight, left: 20, right: 20 + pageWidth, top: 20, width: pageWidth });

  findByTestId(mounted.root, "pdf-zoom-in").props.onClick();
  await waitFor(() => fake.calls.renderPage.at(-1)?.options.scale === 1.25, "zoom re-render");

  // Anchor captured at the viewport center (400, 300) of the 600x800 page:
  // x = 20 + 750 * 0.6333 - 400 = 95, y = 20 + 1000 * 0.35 - 300 = 70.
  await waitFor(() => frame.scrollLeft === 95 && frame.scrollTop === 70, "scroll anchor restoration");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfViewer re-renders the page when the display pixel ratio changes", async (t) => {
  const display = stubDisplayScale(t, 1);
  stubObjectUrls(t);
  const fake = createPdfRuntime({ pageCount: 1 });
  const mounted = await mount(PdfViewer, { source: PDF_SOURCE, runtime: fake.runtime });
  await waitFor(() => findByTestId(mounted.root, "pdf-status")?.props["data-state"] === "ready", "rendered PDF page");
  assert.equal(fake.calls.renderPage[0].options.dpr, 1);

  display.change(3);
  await waitFor(() => fake.calls.renderPage.length === 2, "re-render after the display scale changed");
  assert.equal(fake.calls.renderPage[1].options.dpr, 3);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfViewer exposes intrinsic page rotation separately from the user rotation", async (t) => {
  stubObjectUrls(t);
  const fake = createPdfRuntime({ pageCount: 2, rotations: [0, 90] });
  const mounted = await mount(PdfViewer, { source: PDF_SOURCE, runtime: fake.runtime });
  await waitFor(() => findByTestId(mounted.root, "pdf-status")?.props["data-state"] === "ready", "first PDF page");
  assert.equal(findByTestId(mounted.root, "pdf-document")?.props["data-rotation"], 0);
  assert.equal(findByTestId(mounted.root, "pdf-page")?.props["data-page-rotation"], 0);

  findByTestId(mounted.root, "pdf-next-page").props.onClick();
  await waitFor(
    () => findByTestId(mounted.root, "pdf-page")?.props["data-page-number"] === 2 &&
      findByTestId(mounted.root, "pdf-page")?.props["data-page-rotation"] === 90,
    "second page intrinsic /Rotate metadata",
  );
  // The user rotation control stays untouched: the 90 degrees come from the
  // PDF page dictionary, and the runtime receives only the user rotation.
  assert.equal(findByTestId(mounted.root, "pdf-document")?.props["data-rotation"], 0);
  await waitFor(() => fake.calls.renderPage.at(-1)?.pageIndex === 1, "second page render");
  assert.equal(fake.calls.renderPage.at(-1)?.options.rotation, 0);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PdfViewer and PdfSurface hide the text layer for image-only pages", async (t) => {
  stubObjectUrls(t);
  const emptyLayer = (pageIndex) => ({ pageIndex, width: 600, height: 800, glyphs: [] });

  const viewerRuntime = createPdfRuntime({ pageCount: 1, onGetPageTextLayer: emptyLayer });
  const viewer = await mount(PdfViewer, { source: PDF_SOURCE, runtime: viewerRuntime.runtime });
  await waitFor(
    () => findByTestId(viewer.root, "pdf-status")?.props["data-state"] === "ready" &&
      viewerRuntime.calls.getPageTextLayer.length > 0,
    "image-only viewer text-layer probe",
  );
  await vue.nextTick();
  assert.equal(
    findByTestId(viewer.root, "pdf-text-layer"),
    undefined,
    "an image-only page must not expose an empty selectable text layer",
  );
  assert.ok(
    findByTestId(viewer.root, "pdf-selection-layer"),
    "the selection layer still reports that selection is unavailable",
  );
  assert.deepEqual(viewer.warnings, []);
  viewer.app.unmount();

  const surfaceRuntime = createPdfRuntime({ pageCount: 1, onGetPageTextLayer: emptyLayer });
  const mounted = await mountSurfaceHarness({ source: PDF_SOURCE, runtime: surfaceRuntime.runtime });
  await waitFor(
    () => surfacePageImages(mounted.root).length === 1 && surfaceRuntime.calls.getPageTextLayer.length > 0,
    "image-only surface text-layer probe",
  );
  await vue.nextTick();
  assert.equal(
    walk(mounted.root).find((node) => node.props?.["data-text-page"] === 0),
    undefined,
    "the surface must not mount an empty text layer over an image-only page",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});
