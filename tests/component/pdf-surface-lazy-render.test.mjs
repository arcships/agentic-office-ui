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

const { PdfSurface } = await importFromDemo("@arcships/vue-pdf");

// Mirrors the page stack the surface lays out: each slot is the page display
// height plus the flex gap, starting after the container's top padding.
const PAGE_BASE_HEIGHT = 800;
const PAGE_GAP_PX = 8;
const PAGES_PADDING_TOP_PX = 16;
const VIEWPORT_HEIGHT_PX = 600;
/** Trailing scroll/rasterization debounce used by PdfSurface. */
const SETTLE_MS = 300;

function pageOffset(pageIndex) {
  let top = PAGES_PADDING_TOP_PX;
  for (let index = 0; index < pageIndex; index++) {
    top += PAGE_BASE_HEIGHT + index + PAGE_GAP_PX;
  }
  return top;
}

function pageHeight(pageIndex) {
  return PAGE_BASE_HEIGHT + pageIndex;
}

function createLazyPdfRuntime({ pageCount = 40 } = {}) {
  const document = {
    id: `lazy-pdf-document-${pageCount}`,
    pageCount,
    pages: Array.from({ length: pageCount }, (_, index) => ({
      index,
      width: 600,
      height: pageHeight(index),
      rotation: 0,
    })),
  };
  const calls = {
    closeDocument: [],
    dispose: [],
    getPageTextGeometry: [],
    getTextSlices: [],
    openDocument: [],
    renderPage: [],
    renderThumbnail: [],
    search: [],
  };
  const runtime = {
    runtimeId: "lazy-pdf-runtime",
    async openDocument() {
      calls.openDocument.push(document);
      return document;
    },
    async renderPage(currentDocument, pageIndex, options) {
      calls.renderPage.push({ document: currentDocument, options, pageIndex });
      return new Blob([`page:${pageIndex}`], { type: "image/png" });
    },
    async renderThumbnail(_currentDocument, pageIndex) {
      calls.renderThumbnail.push(pageIndex);
      return new Blob([`thumbnail:${pageIndex}`], { type: "image/png" });
    },
    async getPageTextGeometry(_currentDocument, pageIndex) {
      calls.getPageTextGeometry.push(pageIndex);
      return { runs: [] };
    },
    async getTextSlices(_currentDocument, slices) {
      calls.getTextSlices.push(slices.length);
      return slices.map(() => "");
    },
    async search(_currentDocument, query) {
      calls.search.push(query);
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

function trackObjectUrls() {
  const created = [];
  const revoked = [];
  const originalCreate = URL.createObjectURL;
  const originalRevoke = URL.revokeObjectURL;
  URL.createObjectURL = (value) => {
    const url = originalCreate.call(URL, value);
    created.push(url);
    return url;
  };
  URL.revokeObjectURL = (url) => {
    revoked.push(url);
    if (typeof originalRevoke === "function") originalRevoke.call(URL, url);
  };
  return {
    created,
    revoked,
    restore() {
      URL.createObjectURL = originalCreate;
      URL.revokeObjectURL = originalRevoke;
    },
  };
}

function settle(ms = SETTLE_MS) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function mountLazySurface({ pageCount = 40, zoom = undefined } = {}) {
  const fake = createLazyPdfRuntime({ pageCount });
  const visiblePages = [];
  const loadSuccess = [];
  const zoomState = vue.ref(zoom);
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(PdfSurface, {
        source: {
          kind: "blob",
          blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }),
        },
        runtime: fake.runtime,
        zoom: zoomState.value,
        "onUpdate:zoom": (value) => { zoomState.value = value; },
        onDocumentLoadSuccess: (count) => loadSuccess.push(count),
        onVisiblePageChange: (pageIndex) => visiblePages.push(pageIndex),
      });
    },
  });
  const mounted = await mount(Harness);
  const surface = findByTestId(mounted.root, "pdf-surface");
  assert.ok(surface, "PDF surface root");
  return { ...mounted, fake, loadSuccess, surface, visiblePages, zoom: zoomState };
}

function renderedPageIndexes(fake) {
  return fake.calls.renderPage.map((call) => call.pageIndex);
}

function pageImages(root) {
  return walk(root).filter((node) => node.props?.class === "pdf-surface__page-img");
}

function pagePlaceholders(root) {
  return walk(root).filter((node) => node.props?.class === "pdf-surface__page-placeholder");
}

function scrollSurfaceTo(surface, scrollTop) {
  surface.clientHeight = VIEWPORT_HEIGHT_PX;
  surface.scrollTop = scrollTop;
  surface.props.onScrollPassive();
}

test("PdfSurface rasterizes only the pages around the viewport", async () => {
  const urls = trackObjectUrls();
  const mounted = await mountLazySurface({ pageCount: 40 });
  try {
    await waitFor(() => mounted.loadSuccess.length === 1, "PDF document load");
    assert.deepEqual(mounted.loadSuccess, [40]);
    await waitFor(() => renderedPageIndexes(mounted.fake).length === 3, "leading render window");
    await settle();

    assert.deepEqual(
      renderedPageIndexes(mounted.fake),
      [0, 1, 2],
      "an unmeasured viewport rasterizes a bounded leading window, not all 40 pages",
    );
    assert.equal(urls.created.length, 3, "one bitmap per rendered page");
    assert.equal(urls.revoked.length, 0);
    assert.equal(pageImages(mounted.root).length, 3, "only the rendered window shows pixels");
    assert.equal(pagePlaceholders(mounted.root).length, 37, "every other page keeps its slot");

    mounted.zoom.value = 1.25;
    await waitFor(
      () => renderedPageIndexes(mounted.fake).length === 6,
      "zoom re-render of the visible window",
    );
    await settle();

    const zoomedCalls = mounted.fake.calls.renderPage.slice(3);
    assert.deepEqual(
      zoomedCalls.map((call) => call.pageIndex),
      [0, 1, 2],
      "zoom re-rasterizes the visible window instead of every page",
    );
    assert.equal(zoomedCalls.every((call) => call.options.scale === 1.25), true);
    assert.equal(urls.created.length, 6);
    assert.equal(urls.revoked.length, 3, "stale bitmaps are revoked when the scale changes");
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
    urls.restore();
  }
});

test("PdfSurface moves its render window with the scroll instead of querying every slot", async () => {
  const urls = trackObjectUrls();
  const mounted = await mountLazySurface({ pageCount: 40 });
  try {
    await waitFor(() => renderedPageIndexes(mounted.fake).length === 3, "leading render window");
    await settle();

    // The scroll path must use the cached page offsets: any per-page DOM query
    // would hit this trap.
    mounted.surface.querySelector = () => {
      throw new Error("PDF scroll handling must not query the DOM per page");
    };
    scrollSurfaceTo(mounted.surface, pageOffset(20));
    await waitFor(
      () => renderedPageIndexes(mounted.fake).length === 6,
      "render window after scrolling",
    );
    await settle();

    assert.deepEqual(
      renderedPageIndexes(mounted.fake).slice(3),
      [19, 20, 21],
      "scrolling rasterizes the destination window plus one page of overscan",
    );
    assert.equal(mounted.visiblePages.at(-1), 20, "the visible page follows the cached offsets");
    assert.equal(urls.revoked.length, 3, "the pages that left the viewport release their bitmaps");
    assert.equal(pageImages(mounted.root).length, 3);

    scrollSurfaceTo(mounted.surface, pageOffset(20));
    await settle();
    assert.equal(renderedPageIndexes(mounted.fake).length, 6, "the settled window is not re-rendered");
    assert.equal(mounted.visiblePages.at(-1), 20);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
    urls.restore();
  }
});

test("PdfSurface debounces a fast scroll into a single rasterization pass", async () => {
  const urls = trackObjectUrls();
  const mounted = await mountLazySurface({ pageCount: 40 });
  try {
    await waitFor(() => renderedPageIndexes(mounted.fake).length === 3, "leading render window");
    await settle();

    for (const pageIndex of [5, 15, 30]) scrollSurfaceTo(mounted.surface, pageOffset(pageIndex));
    await waitFor(
      () => renderedPageIndexes(mounted.fake).length > 3,
      "rasterization after the scroll settled",
    );
    await settle();

    assert.deepEqual(
      renderedPageIndexes(mounted.fake).slice(3),
      [29, 30, 31],
      "only the window the scroll stopped on is rasterized",
    );
    assert.equal(urls.created.length, 6);
    assert.equal(urls.revoked.length, 3);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
    urls.restore();
  }
});

test("PdfSurface revokes every page bitmap when the document is released", async () => {
  const urls = trackObjectUrls();
  const mounted = await mountLazySurface({ pageCount: 12 });
  try {
    await waitFor(() => renderedPageIndexes(mounted.fake).length === 3, "leading render window");
    await settle();
    assert.equal(urls.created.length, 3);
    mounted.app.unmount();
    await waitFor(() => urls.revoked.length === 3, "object URL release on unmount");
    assert.equal(urls.revoked.length, urls.created.length);
    assert.equal(mounted.fake.calls.closeDocument.length, 1);
  } finally {
    mounted.app.unmount();
    urls.restore();
  }
});
