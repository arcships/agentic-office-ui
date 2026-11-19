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
const PAGE_BASE_WIDTH = 600;
const PAGE_GAP_PX = 8;
const PAGES_PADDING_TOP_PX = 16;
const VIEWPORT_HEIGHT_PX = 600;
/** Trailing scroll/rasterization debounce used by PdfSurface (100 ms) plus slack. */
const SETTLE_MS = 300;

function pageOffset(pageIndex) {
  let top = PAGES_PADDING_TOP_PX;
  for (let index = 0; index < pageIndex; index++) {
    top += pageHeight(index) + PAGE_GAP_PX;
  }
  return top;
}

function pageHeight(pageIndex) {
  return PAGE_BASE_HEIGHT + pageIndex;
}

/** RGBA footprint the surface must account for when it rasterizes a page. */
function pageBitmapBytes(pageIndex, dpr = 1) {
  return PAGE_BASE_WIDTH * dpr * pageHeight(pageIndex) * dpr * 4;
}

function createLazyPdfRuntime({ pageCount = 40 } = {}) {
  const document = {
    id: `lazy-pdf-document-${pageCount}`,
    pageCount,
    pages: Array.from({ length: pageCount }, (_, index) => ({
      index,
      width: PAGE_BASE_WIDTH,
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

async function mountLazySurface({
  pageCount = 40,
  zoom = undefined,
  maxBitmapBytes = undefined,
  maxPagePixels = undefined,
} = {}) {
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
        // These tests measure page bitmaps; the shared thumbnail strip has its
        // own navigation-aids coverage.
        showThumbnails: false,
        zoom: zoomState.value,
        maxBitmapBytes,
        maxPagePixels,
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

test("a 200-page PDF rasterizes only the pages around the viewport", async () => {
  const urls = trackObjectUrls();
  const mounted = await mountLazySurface({ pageCount: 200 });
  try {
    await waitFor(() => mounted.loadSuccess.length === 1, "PDF document load");
    assert.deepEqual(mounted.loadSuccess, [200]);
    await waitFor(() => renderedPageIndexes(mounted.fake).length === 3, "leading render window");
    await settle();

    assert.deepEqual(
      renderedPageIndexes(mounted.fake),
      [0, 1, 2],
      "an unmeasured viewport rasterizes a bounded leading window, not all 200 pages",
    );
    assert.equal(urls.created.length, 3, "one bitmap per rendered page");
    assert.equal(urls.revoked.length, 0);
    assert.equal(pageImages(mounted.root).length, 3, "only the rendered window shows pixels");
    assert.equal(pagePlaceholders(mounted.root).length, 197, "every other page keeps its slot");
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
    urls.restore();
  }
});

test("zoom 100 → 150 → 200 re-renders at most the visible range per step", async () => {
  const urls = trackObjectUrls();
  const mounted = await mountLazySurface({ pageCount: 200, zoom: 1 });
  try {
    await waitFor(() => renderedPageIndexes(mounted.fake).length === 3, "leading render window");
    await settle();

    for (const [step, nextZoom] of [1.5, 2].entries()) {
      const before = renderedPageIndexes(mounted.fake).length;
      mounted.zoom.value = nextZoom;
      await waitFor(
        () => renderedPageIndexes(mounted.fake).length > before,
        `zoom step ${step + 1} re-render`,
      );
      await settle();
      const stepCalls = mounted.fake.calls.renderPage.slice(before);
      assert.ok(
        stepCalls.length <= 3,
        `zoom to ${nextZoom * 100}% rasterized ${stepCalls.length} pages instead of the visible range`,
      );
      assert.deepEqual(stepCalls.map((call) => call.pageIndex), [0, 1, 2]);
      assert.equal(stepCalls.every((call) => call.options.scale === nextZoom), true);
    }
    assert.equal(urls.created.length, 9, "three window bitmaps per zoom step");
    assert.equal(urls.revoked.length, 6, "the previous scale's bitmaps are revoked");
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
    urls.restore();
  }
});

test("scrolling keeps off-screen bitmaps in the LRU and reuses them on the way back", async () => {
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
    assert.equal(urls.revoked.length, 0, "default budget keeps the first window cached");
    assert.equal(pageImages(mounted.root).length, 6, "cached pages keep their bitmap mounted");

    scrollSurfaceTo(mounted.surface, pageOffset(20));
    await settle();
    assert.equal(renderedPageIndexes(mounted.fake).length, 6, "the settled window is not re-rendered");

    scrollSurfaceTo(mounted.surface, 0);
    await settle();
    assert.equal(
      renderedPageIndexes(mounted.fake).length,
      6,
      "scrolling back reuses the cached bitmaps instead of rasterizing again",
    );
    assert.equal(urls.revoked.length, 0);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
    urls.restore();
  }
});

test("a byte budget smaller than the document evicts the least recently used page", async () => {
  const urls = trackObjectUrls();
  // Two of the three leading pages fit; the oldest one has to go.
  const maxBitmapBytes = Math.floor(pageBitmapBytes(1) + pageBitmapBytes(2) + 1);
  const mounted = await mountLazySurface({ pageCount: 40, maxBitmapBytes });
  try {
    await waitFor(() => renderedPageIndexes(mounted.fake).length === 3, "leading render window");
    await settle();

    scrollSurfaceTo(mounted.surface, pageOffset(20));
    await waitFor(
      () => renderedPageIndexes(mounted.fake).length === 6,
      "render window after scrolling",
    );
    await settle();
    assert.equal(urls.revoked.length, 1, "the LRU evicted exactly the oldest bitmap");

    scrollSurfaceTo(mounted.surface, 0);
    await waitFor(
      () => renderedPageIndexes(mounted.fake).length === 7,
      "evicted page rasterizes again on the way back",
    );
    await settle();
    assert.equal(
      renderedPageIndexes(mounted.fake).at(-1),
      0,
      "only the evicted page 0 rasterized; pages 1 and 2 came from the cache",
    );
    assert.ok(urls.revoked.length >= 1);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
    urls.restore();
  }
});

test("maxBitmapBytes 0 releases every bitmap that leaves the viewport", async () => {
  const urls = trackObjectUrls();
  const mounted = await mountLazySurface({ pageCount: 40, maxBitmapBytes: 0 });
  try {
    await waitFor(() => renderedPageIndexes(mounted.fake).length === 3, "leading render window");
    await settle();

    scrollSurfaceTo(mounted.surface, pageOffset(20));
    await waitFor(
      () => renderedPageIndexes(mounted.fake).length === 6,
      "render window after scrolling",
    );
    await settle();
    assert.equal(urls.revoked.length, 3, "retention disabled: the old window is released");
    assert.equal(pageImages(mounted.root).length, 3);

    scrollSurfaceTo(mounted.surface, 0);
    await waitFor(
      () => renderedPageIndexes(mounted.fake).length === 8,
      "released pages rasterize again on the way back",
    );
    assert.deepEqual(renderedPageIndexes(mounted.fake).slice(6), [0, 1]);
  } finally {
    mounted.app.unmount();
    urls.restore();
  }
});

test("the per-page pixel cap downsizes the raster, not the layout", async () => {
  const urls = trackObjectUrls();
  const maxPagePixels = 120_000;
  const mounted = await mountLazySurface({ pageCount: 12, maxPagePixels });
  try {
    await waitFor(() => renderedPageIndexes(mounted.fake).length === 3, "leading render window");
    await settle();

    const firstCall = mounted.fake.calls.renderPage[0];
    assert.ok(firstCall.options.scale < 1, "the raster scale is reduced to fit the budget");
    const rasterWidth = PAGE_BASE_WIDTH * firstCall.options.scale * firstCall.options.dpr;
    const rasterHeight = pageHeight(0) * firstCall.options.scale * firstCall.options.dpr;
    assert.ok(
      rasterWidth * rasterHeight <= maxPagePixels,
      `rendered ${rasterWidth * rasterHeight} px against a ${maxPagePixels} px budget`,
    );

    const images = pageImages(mounted.root);
    assert.equal(images[0].props.style.width, `${PAGE_BASE_WIDTH}px`, "the slot keeps its CSS size");
    assert.equal(images[0].props.style.height, `${pageHeight(0)}px`, "CSS downscales the bitmap");
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
    urls.restore();
  }
});

test("a resize burst re-renders once after the debounce window", async () => {
  const urls = trackObjectUrls();
  const originalResizeObserver = globalThis.ResizeObserver;
  const observers = [];
  globalThis.ResizeObserver = class FakeResizeObserver {
    constructor(callback) {
      this.callback = callback;
      observers.push(this);
    }
    observe() {}
    disconnect() {}
  };
  const mounted = await mountLazySurface({ pageCount: 40 });
  try {
    await waitFor(() => renderedPageIndexes(mounted.fake).length === 3, "leading render window");
    await settle();
    assert.equal(observers.length, 1, "the surface observes its scroll container");

    // A taller viewport pulls one more page into the window.
    mounted.surface.clientHeight = PAGE_BASE_HEIGHT * 3;
    for (let tick = 0; tick < 6; tick++) observers[0].callback();
    await settle(60);
    assert.equal(
      renderedPageIndexes(mounted.fake).length,
      3,
      "nothing rasterizes while the resize is still in flight",
    );

    await waitFor(
      () => renderedPageIndexes(mounted.fake).length === 4,
      "single rasterization pass after the resize debounce",
    );
    await settle();
    assert.equal(
      renderedPageIndexes(mounted.fake).length,
      4,
      "six resize ticks coalesced into one pass, not one pass each",
    );
    assert.deepEqual(renderedPageIndexes(mounted.fake).slice(3), [3], "only the newly visible page");
    assert.deepEqual(mounted.warnings, []);
  } finally {
    globalThis.ResizeObserver = originalResizeObserver;
    mounted.app.unmount();
    urls.restore();
  }
});

test("scroll frames stay flat as the page count grows", async () => {
  async function measureScrollFrames(pageCount) {
    const urls = trackObjectUrls();
    const mounted = await mountLazySurface({ pageCount });
    try {
      await waitFor(() => renderedPageIndexes(mounted.fake).length === 3, "leading render window");
      await settle();
      // Frames must not touch the DOM: the cached offsets decide the window.
      mounted.surface.querySelector = () => {
        throw new Error("scroll frames must not query page slots");
      };
      mounted.surface.querySelectorAll = () => {
        throw new Error("scroll frames must not query page slots");
      };
      mounted.surface.getBoundingClientRect = () => {
        throw new Error("scroll frames must not measure layout");
      };
      const frames = 300;
      const startedAt = process.hrtime.bigint();
      for (let frame = 0; frame < frames; frame++) {
        scrollSurfaceTo(mounted.surface, pageOffset(Math.floor((frame / frames) * (pageCount - 4))));
      }
      const elapsedMs = Number(process.hrtime.bigint() - startedAt) / 1e6;
      await waitFor(
        () => renderedPageIndexes(mounted.fake).length === 6,
        "one rasterization pass after the scroll burst",
      );
      await settle();
      return {
        elapsedMs,
        frames,
        perFrameMs: elapsedMs / frames,
        renderCalls: renderedPageIndexes(mounted.fake).length,
      };
    } finally {
      mounted.app.unmount();
      urls.restore();
    }
  }

  const small = await measureScrollFrames(40);
  const large = await measureScrollFrames(400);

  for (const measured of [small, large]) {
    assert.ok(
      measured.perFrameMs < 2,
      `scroll frame took ${measured.perFrameMs.toFixed(3)} ms; a 60 fps frame budget is 16.7 ms`,
    );
    assert.equal(
      measured.renderCalls,
      6,
      "a scroll burst rasterizes the settled window once, independent of the page count",
    );
  }
  assert.ok(
    large.perFrameMs < small.perFrameMs * 4 + 0.5,
    `per-frame cost grew with the page count: ${small.perFrameMs.toFixed(3)} → ${large.perFrameMs.toFixed(3)} ms`,
  );
});

test("PdfSurface revokes every page bitmap when the document is released", async () => {
  const urls = trackObjectUrls();
  const mounted = await mountLazySurface({ pageCount: 12 });
  try {
    await waitFor(() => renderedPageIndexes(mounted.fake).length === 3, "leading render window");
    await settle();
    scrollSurfaceTo(mounted.surface, pageOffset(6));
    await waitFor(
      () => renderedPageIndexes(mounted.fake).length === 6,
      "second render window",
    );
    await settle();
    assert.equal(urls.created.length, 6);
    const createdBeforeUnmount = urls.created.length;
    mounted.app.unmount();
    await waitFor(
      () => urls.revoked.length === createdBeforeUnmount,
      "object URL release on unmount",
    );
    assert.equal(mounted.fake.calls.closeDocument.length, 1);
  } finally {
    mounted.app.unmount();
    urls.restore();
  }
});
