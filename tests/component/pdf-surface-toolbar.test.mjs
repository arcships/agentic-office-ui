// Issue #172 — `PdfSurface` can opt into the shared PDF toolbar while its
// default footprint stays toolbar-free. The toolbar is the same
// `PdfViewerToolbar` used by `PdfViewer`, so page/zoom/rotate/search/download/
// print keep their viewer test ids and keyboard behaviour. Layout is the one
// command with no Surface equivalent: `PdfSurface.pageLayout` is a controlled
// prop, so the toolbar renders that group disabled instead of inventing a
// second source of truth.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  keydownEvent,
  mount,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const { PdfSurface } = await importFromDemo("@arcships/vue-pdf");

// Vue's native v-model directive checks these browser constructors during an
// update. The custom component renderer intentionally has no browser DOM.
if (globalThis.Document === undefined) globalThis.Document = class Document {};
if (globalThis.ShadowRoot === undefined) globalThis.ShadowRoot = class ShadowRoot {};

const PDF_SOURCE = {
  kind: "blob",
  blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }),
};

function createRuntime({ pageCount = 3, searchHits = [] } = {}) {
  const calls = { search: [], renderPage: [], renderThumbnail: [] };
  const document = {
    id: "pdf-surface-toolbar-document",
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
      runtimeId: "pdf-surface-toolbar-runtime",
      async openDocument() {
        return document;
      },
      async renderPage(_document, pageIndex) {
        calls.renderPage.push(pageIndex);
        return new Blob([`page:${pageIndex}`], { type: "image/png" });
      },
      async renderThumbnail(_document, pageIndex) {
        calls.renderThumbnail.push(pageIndex);
        return new Blob([`thumbnail:${pageIndex}`], { type: "image/png" });
      },
      async getPageLabels() {
        return ["1", "2", "3"].slice(0, pageCount);
      },
      async getOutline() {
        return [];
      },
      async getPageTextGeometry() {
        return { runs: [] };
      },
      async getTextSlices(_document, slices) {
        return slices.map(() => "");
      },
      async search(_document, query) {
        calls.search.push(query);
        return searchHits;
      },
      async closeDocument() {},
      async dispose() {},
    },
  };
}

function stubObjectUrls(t) {
  const originalCreate = URL.createObjectURL;
  const originalRevoke = URL.revokeObjectURL;
  let counter = 0;
  URL.createObjectURL = () => `blob:pdf-surface-toolbar-${++counter}`;
  URL.revokeObjectURL = () => {};
  t.after(() => {
    URL.createObjectURL = originalCreate;
    URL.revokeObjectURL = originalRevoke;
  });
}

function classNames(node) {
  return String(node?.props?.class || "").split(/\s+/).filter(Boolean);
}

function nodesWithClass(root, className) {
  return walk(root).filter((node) => classNames(node).includes(className));
}

/** The Surface calls `scrollToPage` through a DOM query; the host tree is fake. */
function stubSurfacePageScroll(root) {
  const scrolled = [];
  const surface = findByTestId(root, "pdf-surface");
  surface.querySelector = (selector) => {
    const match = /data-page-index="(\d+)"/u.exec(String(selector));
    return {
      querySelector: () => null,
      scrollIntoView: (options) => scrolled.push({
        pageIndex: Number(match?.[1]),
        options,
      }),
    };
  };
  return { scrolled };
}

async function waitForToolbarReady(root) {
  await waitFor(
    () => findByTestId(root, "pdf-page-count")?.children[0]?.text === "3",
    "surface toolbar with a three-page document",
  );
}

test("PdfSurface keeps its toolbar-free default footprint", async (t) => {
  stubObjectUrls(t);
  const mounted = await mount(PdfSurface, {
    source: PDF_SOURCE,
    runtime: createRuntime().runtime,
  });
  try {
    await waitFor(
      () => nodesWithClass(mounted.root, "pdf-surface__page-img").length > 0,
      "rendered surface page",
    );

    assert.equal(
      findByTestId(mounted.root, "pdf-toolbar"),
      undefined,
      "showToolbar defaults to false and the minimal surface stays toolbar-free",
    );
    assert.equal(findByTestId(mounted.root, "pdf-page-input"), undefined);
    assert.equal(findByTestId(mounted.root, "pdf-surface-status-bar"), undefined);
    assert.ok(findByTestId(mounted.root, "pdf-surface"), "the scroll surface is still mounted");
    assert.ok(findByTestId(mounted.root, "pdf-surface-nav"), "the default thumbnails rail is unchanged");
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("PdfSurface toolbar page, zoom, rotate and sidebar stay on the Surface state", async (t) => {
  stubObjectUrls(t);
  const fake = createRuntime();
  const mounted = await mount(PdfSurface, {
    source: PDF_SOURCE,
    runtime: fake.runtime,
    showToolbar: true,
  });
  try {
    await waitForToolbarReady(mounted.root);
    const { scrolled } = stubSurfacePageScroll(mounted.root);
    const toolbar = findByTestId(mounted.root, "pdf-toolbar");
    assert.equal(toolbar.props.role, "toolbar");
    assert.equal(toolbar.props["aria-label"], "Document toolbar");
    assert.equal(findByTestId(mounted.root, "pdf-page-input").value, "1");
    assert.equal(findByTestId(mounted.root, "pdf-page-count").children[0].text, "3");
    assert.equal(
      findByTestId(mounted.root, "pdf-zoom-select").value,
      1,
      "the toolbar starts at the Surface zoom",
    );

    findByTestId(mounted.root, "pdf-actions").props.onClick();
    await vue.nextTick();
    findByTestId(mounted.root, "pdf-next-page").props.onClick();
    await vue.nextTick();
    assert.equal(findByTestId(mounted.root, "pdf-page-input").value, "2");
    assert.deepEqual(scrolled.at(-1).pageIndex, 1, "toolbar pages are 1-based, Surface pages 0-based");
    assert.equal(
      findByTestId(mounted.root, "pdf-surface").props["data-page-layout"],
      "continuous",
      "the controlled layout is unchanged",
    );

    findByTestId(mounted.root, "pdf-zoom-in").props.onClick();
    await vue.nextTick();
    assert.equal(findByTestId(mounted.root, "pdf-zoom-select").value, 1.25);
    findByTestId(mounted.root, "pdf-zoom-out").props.onClick();
    await vue.nextTick();
    assert.equal(findByTestId(mounted.root, "pdf-zoom-select").value, 1);

    const slotWidth = () => nodesWithClass(mounted.root, "pdf-surface__page-slot")[0].props.style.width;
    assert.equal(slotWidth(), "600px");
    findByTestId(mounted.root, "pdf-rotate").props.onClick();
    await vue.nextTick();
    assert.equal(slotWidth(), "800px", "rotation drives the Surface page geometry");

    const nav = () => findByTestId(mounted.root, "pdf-surface-nav");
    assert.equal(nav().props["data-sidebar-open"], "true", "the Surface rail starts open as before");
    findByTestId(mounted.root, "pdf-sidebar-toggle").props.onClick();
    await vue.nextTick();
    assert.equal(nav().props["data-sidebar-open"], "false");
    findByTestId(mounted.root, "pdf-sidebar-toggle").props.onClick();
    await vue.nextTick();
    assert.equal(nav().props["data-sidebar-open"], "true");
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("PdfSurface toolbar search uses the Surface search session and Escape closes it", async (t) => {
  stubObjectUrls(t);
  const fake = createRuntime({
    searchHits: [
      { kind: "pdf-text", pageIndex: 1, charIndex: 0, charCount: 6, before: "", match: "needle", after: "", rects: [] },
      { kind: "pdf-text", pageIndex: 2, charIndex: 4, charCount: 6, before: "", match: "needle", after: "", rects: [] },
    ],
  });
  const mounted = await mount(PdfSurface, {
    source: PDF_SOURCE,
    runtime: fake.runtime,
    showToolbar: true,
  });
  try {
    await waitForToolbarReady(mounted.root);
    stubSurfacePageScroll(mounted.root);

    findByTestId(mounted.root, "pdf-search-toggle").props.onClick();
    await vue.nextTick();
    const input = findByTestId(mounted.root, "pdf-search-input");
    assert.ok(input, "search opens on the toolbar");
    input.value = "needle";
    input.dispatchEvent({ type: "input", target: input });
    await vue.nextTick();
    const searchForm = walk(mounted.root).find(
      (node) => node.type === "form" && node.props?.role === "search",
    );
    assert.ok(searchForm, "the toolbar search form is rendered");
    searchForm.props.onSubmit({ preventDefault() {} });
    await waitFor(
      () => /Result 1 \/ 2, page 2/u.test(textContent(findByTestId(mounted.root, "pdf-search-result"))),
      "first search result",
    );
    assert.deepEqual(fake.calls.search, ["needle"], "the Surface runtime owns the search");

    findByTestId(mounted.root, "pdf-search-next").props.onClick();
    await waitFor(
      () => /Result 2 \/ 2, page 3/u.test(textContent(findByTestId(mounted.root, "pdf-search-result"))),
      "next search result",
    );
    findByTestId(mounted.root, "pdf-search-prev").props.onClick();
    await waitFor(
      () => /Result 1 \/ 2, page 2/u.test(textContent(findByTestId(mounted.root, "pdf-search-result"))),
      "previous search result",
    );

    const escape = keydownEvent("Escape");
    let focused = 0;
    findByTestId(mounted.root, "pdf-surface").focus = () => { focused += 1; };
    findByTestId(mounted.root, "pdf-toolbar").props.onKeydown(escape.event);
    await vue.nextTick();
    assert.equal(findByTestId(mounted.root, "pdf-search-input"), undefined);
    assert.equal(findByTestId(mounted.root, "pdf-search-result"), undefined);
    assert.ok(focused >= 1, "closing search returns focus to the Surface");

    // Ctrl/Cmd + F is enabled by the opt-in toolbar and is handled by the
    // shell, so it also works while a toolbar control has focus.
    const searchKey = keydownEvent("f", { ctrlKey: true, target: { tagName: "DIV" } });
    findByTestId(mounted.root, "pdf-surface-shell").props.onKeydown(searchKey.event);
    await vue.nextTick();
    assert.ok(findByTestId(mounted.root, "pdf-search-input"), "Ctrl+F opens the Surface toolbar search");
    const reopen = findByTestId(mounted.root, "pdf-search-input");
    assert.equal(reopen.value, "", "closing search cleared the draft");
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("PdfSurface toolbar layout is disabled because pageLayout is host-controlled", async (t) => {
  stubObjectUrls(t);
  const mounted = await mount(PdfSurface, {
    source: PDF_SOURCE,
    runtime: createRuntime().runtime,
    showToolbar: true,
    pageLayout: "continuous",
  });
  try {
    await waitForToolbarReady(mounted.root);
    findByTestId(mounted.root, "pdf-actions").props.onClick();
    await vue.nextTick();
    const group = findByTestId(mounted.root, "pdf-page-layout");
    assert.equal(group.props["data-layout-disabled"], "true");
    for (const layout of ["continuous", "single", "spread"]) {
      assert.equal(
        findByTestId(mounted.root, `pdf-page-layout-${layout}`).props.disabled,
        true,
        `${layout} is explicitly disabled`,
      );
    }
    findByTestId(mounted.root, "pdf-page-layout-single").props.onClick();
    await vue.nextTick();
    assert.equal(
      findByTestId(mounted.root, "pdf-surface").props["data-page-layout"],
      "continuous",
      "a disabled layout command cannot diverge from the pageLayout prop",
    );
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("PdfSurface toolbar print and download use the loaded Surface bytes", async (t) => {
  stubObjectUrls(t);
  const previousWindow = globalThis.window;
  const previousDocument = globalThis.document;
  const downloads = [];
  let printed = 0;
  const mounted = await mount(PdfSurface, {
    source: PDF_SOURCE,
    runtime: createRuntime().runtime,
    showToolbar: true,
  });
  try {
    await waitForToolbarReady(mounted.root);
    findByTestId(mounted.root, "pdf-actions").props.onClick();
    await vue.nextTick();

    globalThis.window = { print: () => { printed += 1; } };
    findByTestId(mounted.root, "pdf-print").props.onClick();
    await waitFor(() => printed === 1, "browser print dialog");

    globalThis.document = {
      createElement: () => ({
        href: "",
        download: "",
        rel: "",
        click() { downloads.push(this.download); },
        remove() {},
      }),
      body: { append() {} },
    };
    findByTestId(mounted.root, "pdf-actions").props.onClick();
    await vue.nextTick();
    findByTestId(mounted.root, "pdf-download").props.onClick();
    await waitFor(() => downloads.length === 1, "download anchor click");
    assert.deepEqual(downloads, ["document.pdf"]);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
    if (previousWindow === undefined) delete globalThis.window;
    else globalThis.window = previousWindow;
    if (previousDocument === undefined) delete globalThis.document;
    else globalThis.document = previousDocument;
  }
});

test("PdfSurface toolbar wraps at 390px and the sidebar uses the overlay Escape path", async (t) => {
  stubObjectUrls(t);
  const previousWindow = globalThis.window;
  const previousDocument = globalThis.document;
  const fakeDocument = new EventTarget();
  globalThis.document = fakeDocument;
  globalThis.window = {
    innerWidth: 390,
    addEventListener() {},
    removeEventListener() {},
    matchMedia(query) {
      return {
        media: query,
        matches: false,
        addEventListener() {},
        removeEventListener() {},
      };
    },
  };
  const mounted = await mount(PdfSurface, {
    source: PDF_SOURCE,
    runtime: createRuntime().runtime,
    showToolbar: true,
  });
  try {
    await waitForToolbarReady(mounted.root);
    const toolbar = findByTestId(mounted.root, "pdf-toolbar");
    assert.ok(toolbar, "the toolbar is mounted at 390px");
    assert.ok(walk(toolbar).some((node) => classNames(node).includes("pdf-toolbar__leading")));
    assert.ok(walk(toolbar).some((node) => classNames(node).includes("pdf-toolbar__trailing")));
    // The host renderer has no layout engine. The wrapping contract lives in
    // the shared toolbar CSS; every fixed-size control sits inside one of the
    // two flex groups that move to a new line instead of widening the row.
    const source = readFileSync(
      new URL("../../packages/vue-pdf/src/components/PdfViewerToolbar.vue", import.meta.url),
      "utf8",
    );
    assert.match(source, /\.pdf-toolbar\s*\{[^}]*flex-wrap:\s*wrap/s);
    assert.match(source, /box-sizing:\s*border-box/);

    const nav = () => findByTestId(mounted.root, "pdf-surface-nav");
    assert.equal(nav().props["data-sidebar-mode"], "overlay", "390px resolves the rail to overlay");
    assert.equal(nav().props["data-sidebar-open"], "true");

    const escape = keydownEvent("Escape");
    findByTestId(mounted.root, "pdf-toolbar").props.onKeydown(escape.event);
    await vue.nextTick();
    assert.equal(nav().props["data-sidebar-open"], "false", "toolbar Escape closes the overlay rail");

    findByTestId(mounted.root, "pdf-sidebar-toggle").props.onClick();
    await vue.nextTick();
    assert.equal(nav().props["data-sidebar-open"], "true");
    const documentEscape = new Event("keydown", { cancelable: true });
    Object.defineProperty(documentEscape, "key", { configurable: true, value: "Escape" });
    fakeDocument.dispatchEvent(documentEscape);
    await vue.nextTick();
    assert.equal(nav().props["data-sidebar-open"], "false", "document Escape also closes the overlay rail");
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
    if (previousWindow === undefined) delete globalThis.window;
    else globalThis.window = previousWindow;
    if (previousDocument === undefined) delete globalThis.document;
    else globalThis.document = previousDocument;
  }
});
