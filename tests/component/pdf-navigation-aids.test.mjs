import assert from "node:assert/strict";
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

const { PdfSurface, PdfViewer } = await importFromDemo("@arcships/vue-pdf");

const PDF_SOURCE = {
  kind: "blob",
  blob: new Blob(["%PDF-1.4\n%%EOF"], { type: "application/pdf" }),
};

const OUTLINE = [
  {
    title: "Chapter 1",
    pageIndex: 0,
    children: [
      {
        title: "Section 1.1",
        pageIndex: 1,
        children: [{ title: "Sub 1.1.1", pageIndex: 2, children: [] }],
      },
      { title: "External", url: "https://example.com/a", children: [] },
    ],
  },
  { title: "Chapter 2", pageIndex: 3, children: [] },
];

const PAGE_LABELS = ["i", "ii", "iii", "iv"];

function createNavigationRuntime({ pageCount = 4, outline = OUTLINE, labels = PAGE_LABELS, withOutline = true, withLabels = true } = {}) {
  const calls = { closeDocument: [], dispose: [], getOutline: 0, getPageLabels: 0, openDocument: [], renderPage: [], renderThumbnail: [] };
  const document = {
    id: "pdf-navigation-document",
    pageCount,
    pages: Array.from({ length: pageCount }, (_, index) => ({
      index,
      width: 600,
      height: 800,
      rotation: 0,
    })),
  };
  const runtime = {
    runtimeId: "pdf-navigation-runtime",
    async openDocument() {
      calls.openDocument.push(true);
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
    ...(withLabels
      ? {
          async getPageLabels() {
            calls.getPageLabels += 1;
            return labels;
          },
        }
      : {}),
    ...(withOutline
      ? {
          async getOutline() {
            calls.getOutline += 1;
            return outline;
          },
        }
      : {}),
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
    const url = `blob:pdf-navigation-${++counter}`;
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

function byTestId(root, testId) {
  return walk(root).filter((node) => node.props?.["data-testid"] === testId);
}

function outlineItems(root) {
  return byTestId(root, "pdf-outline-item");
}

function outlineItem(root, id) {
  return outlineItems(root).find((node) => node.props?.["data-outline-id"] === id);
}

function activeThumbnail(root, testId = "pdf-surface-thumbnail") {
  return byTestId(root, testId).find((node) => classNames(node).includes("active"));
}

function makeScrollable(node, { clientHeight = 600, clientWidth = 900 } = {}) {
  const state = { scrollTop: 0, scrollLeft: 0 };
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
  return state;
}

function stubPageScroll(root) {
  const scrolled = [];
  const surface = findByTestId(root, "pdf-surface");
  surface.querySelector = (selector) => {
    const match = /data-page-index="(\d+)"/u.exec(String(selector));
    return {
      scrollIntoView: (options) => scrolled.push({ pageIndex: Number(match?.[1]), options }),
    };
  };
  return { scrolled, surface };
}

async function mountSurface(options = {}) {
  const fake = createNavigationRuntime(options);
  const outlineActivations = [];
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(PdfSurface, {
        source: PDF_SOURCE,
        runtime: fake.runtime,
        showThumbnails: options.showThumbnails ?? true,
        showOutline: options.showOutline ?? true,
        onOutlineActivate: (event) => outlineActivations.push(event),
      });
    },
  });
  const mounted = await mount(Harness);
  const surface = findByTestId(mounted.root, "pdf-surface");
  assert.ok(surface, "PDF surface root");
  makeScrollable(surface);
  return { ...mounted, fake, outlineActivations, surface };
}

test("PdfSurface renders a bookmark tree and hides it when the runtime has no outline", async (t) => {
  stubObjectUrls(t);
  const mounted = await mountSurface();
  try {
    await waitFor(() => outlineItems(mounted.root).length > 0, "outline tree");
    assert.equal(findByTestId(mounted.root, "pdf-surface-nav") !== undefined, true, "navigation rail");
    assert.equal(findByTestId(mounted.root, "pdf-surface-outline") !== undefined, true, "outline panel");
    assert.equal(findByTestId(mounted.root, "pdf-surface-thumbnails") !== undefined, true, "thumbnails render by default");
    assert.equal(outlineItems(mounted.root).length, 4, "root entries and their children are visible");
    assert.equal(outlineItem(mounted.root, "0-0-0"), undefined, "collapsed grandchildren stay hidden");
    assert.equal(outlineItem(mounted.root, "0").props["aria-expanded"], true);
    assert.equal(outlineItem(mounted.root, "0-0").props["aria-expanded"], false);
    assert.equal(outlineItem(mounted.root, "0-0").props["aria-level"], 2);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }

  const withoutOutline = await mountSurface({ withOutline: false });
  try {
    await waitFor(() => withoutOutline.fake.calls.getPageLabels === 1, "page labels");
    assert.equal(findByTestId(withoutOutline.root, "pdf-surface-outline"), undefined);
    assert.equal(findByTestId(withoutOutline.root, "pdf-surface-nav") !== undefined, true, "thumbnail rail stays");
    assert.equal(findByTestId(withoutOutline.root, "pdf-surface-thumbnails") !== undefined, true);
  } finally {
    withoutOutline.app.unmount();
  }

  const withoutLabels = await mountSurface({ withLabels: false, showThumbnails: true });
  try {
    await waitFor(() => byTestId(withoutLabels.root, "pdf-surface-thumbnail-image").length === 4, "fallback thumbnails");
    assert.equal(withoutLabels.fake.calls.getPageLabels, 0, "runtime without page labels is not asked");
    assert.deepEqual(
      byTestId(withoutLabels.root, "pdf-surface-thumbnail").map((node) => textContent(node)),
      ["1", "2", "3", "4"],
      "thumbnails fall back to decimal page numbers",
    );
  } finally {
    withoutLabels.app.unmount();
  }
});

test("outline keyboard navigation expands, collapses and moves the roving focus", async (t) => {
  stubObjectUrls(t);
  const mounted = await mountSurface();
  try {
    await waitFor(() => outlineItems(mounted.root).length > 0, "outline tree");
    assert.equal(outlineItem(mounted.root, "0").props.tabindex, 0, "first outline entry starts focused");

    const down = keydownEvent("ArrowDown");
    outlineItem(mounted.root, "0").props.onKeydown(down.event);
    await vue.nextTick();
    assert.equal(down.preventedCount(), 1);
    assert.equal(outlineItem(mounted.root, "0-0").props.tabindex, 0);
    assert.equal(outlineItem(mounted.root, "0").props.tabindex, -1);

    const right = keydownEvent("ArrowRight");
    outlineItem(mounted.root, "0-0").props.onKeydown(right.event);
    await vue.nextTick();
    assert.equal(right.preventedCount(), 1);
    assert.equal(outlineItem(mounted.root, "0-0").props["aria-expanded"], true);
    assert.ok(outlineItem(mounted.root, "0-0-0"), "Right arrow expands the collapsed branch");

    const left = keydownEvent("ArrowLeft");
    outlineItem(mounted.root, "0-0").props.onKeydown(left.event);
    await vue.nextTick();
    assert.equal(left.preventedCount(), 1);
    assert.equal(outlineItem(mounted.root, "0-0").props["aria-expanded"], false);
    assert.equal(outlineItem(mounted.root, "0-0-0"), undefined);
    assert.equal(outlineItem(mounted.root, "0-0").props.tabindex, 0, "collapse keeps the parent focused");

    const toggle = walk(outlineItem(mounted.root, "0-0")).find(
      (node) => node.props?.["data-testid"] === "pdf-outline-toggle",
    );
    toggle.props.onClick({ stopPropagation() {} });
    await vue.nextTick();
    assert.ok(outlineItem(mounted.root, "0-0-0"), "clicking the chevron expands without navigating");
    toggle.props.onClick({ stopPropagation() {} });
    await vue.nextTick();
    assert.equal(outlineItem(mounted.root, "0-0-0"), undefined, "clicking the chevron again collapses");

    const end = keydownEvent("End");
    outlineItem(mounted.root, "0-0").props.onKeydown(end.event);
    await vue.nextTick();
    assert.equal(outlineItem(mounted.root, "1").props.tabindex, 0, "End focuses the last visible entry");

    const home = keydownEvent("Home");
    outlineItem(mounted.root, "1").props.onKeydown(home.event);
    await vue.nextTick();
    assert.equal(outlineItem(mounted.root, "0").props.tabindex, 0, "Home focuses the first entry");
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("outline activation navigates to the target page and reports link bookmarks", async (t) => {
  stubObjectUrls(t);
  const mounted = await mountSurface({ showThumbnails: true });
  try {
    await waitFor(() => outlineItems(mounted.root).length > 0, "outline tree");
    await waitFor(() => byTestId(mounted.root, "pdf-surface-thumbnail-image").length === 4, "thumbnails");
    const { scrolled } = stubPageScroll(mounted.root);

    const enter = keydownEvent("Enter");
    outlineItem(mounted.root, "1").props.onKeydown(enter.event);
    await vue.nextTick();
    assert.equal(enter.preventedCount(), 1);
    assert.deepEqual(scrolled.at(-1)?.pageIndex, 3, "Enter navigates to Chapter 2's page");
    assert.equal(
      activeThumbnail(mounted.root)?.props?.["data-page-number"],
      4,
      "the thumbnail strip follows the outline target",
    );
    assert.equal(mounted.outlineActivations.at(-1)?.pageIndex, 3);
    assert.equal(mounted.outlineActivations.at(-1)?.title, "Chapter 2");

    outlineItem(mounted.root, "0-1").props.onClick();
    await vue.nextTick();
    assert.deepEqual(mounted.outlineActivations.at(-1), {
      title: "External",
      url: "https://example.com/a",
      blocked: false,
    });
    assert.deepEqual(scrolled.length, 1, "a URI bookmark does not scroll the document");
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("thumbnail strip follows the visible page and clicking a thumbnail scrolls to it", async (t) => {
  stubObjectUrls(t);
  const mounted = await mountSurface({ showThumbnails: true });
  try {
    await waitFor(() => byTestId(mounted.root, "pdf-surface-thumbnail-image").length === 4, "thumbnails");
    assert.equal(byTestId(mounted.root, "pdf-surface-thumbnail").length, 4);
    assert.equal(
      activeThumbnail(mounted.root)?.props?.["data-page-number"],
      1,
      "the first page starts active",
    );
    const labels = byTestId(mounted.root, "pdf-surface-thumbnail").map((node) => textContent(node));
    assert.deepEqual(labels, ["i", "ii", "iii", "iv"]);

    // Scroll to the third page: 8 px page gap plus the 16 px stack padding.
    mounted.surface.scrollTop = 16 + 2 * (800 + 8);
    mounted.surface.props.onScrollPassive();
    await waitFor(
      () => activeThumbnail(mounted.root)?.props?.["data-page-number"] === 3,
      "thumbnail follows the visible page",
    );
    assert.equal(
      outlineItem(mounted.root, "0-0").props["aria-selected"],
      "true",
      "the outline highlights the entry for the visible page",
    );

    const { scrolled } = stubPageScroll(mounted.root);
    byTestId(mounted.root, "pdf-surface-thumbnail")
      .find((node) => node.props["data-page-number"] === 4)
      .props.onClick();
    await vue.nextTick();
    assert.equal(scrolled.at(-1)?.pageIndex, 3, "clicking page 4 scrolls to index 3");
    assert.equal(activeThumbnail(mounted.root)?.props?.["data-page-number"], 4);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("PdfSurface hides the outline panel when showOutline is false", async (t) => {
  stubObjectUrls(t);
  const mounted = await mountSurface({ showOutline: false });
  try {
    await waitFor(() => mounted.fake.calls.getOutline === 1, "outline request");
    assert.equal(findByTestId(mounted.root, "pdf-surface-outline"), undefined);
    assert.equal(findByTestId(mounted.root, "pdf-surface-nav") !== undefined, true, "thumbnails keep the rail");
    assert.equal(findByTestId(mounted.root, "pdf-surface-thumbnails") !== undefined, true);
  } finally {
    mounted.app.unmount();
  }
});

test("PdfViewer shares the thumbnail strip and labels pages from the runtime", async (t) => {
  stubObjectUrls(t);
  const fake = createNavigationRuntime();
  const mounted = await mount(PdfViewer, {
    source: PDF_SOURCE,
    runtime: fake.runtime,
  });
  try {
    await waitFor(() => byTestId(mounted.root, "pdf-thumbnail-image").length === 4, "viewer thumbnails");
    assert.equal(findByTestId(mounted.root, "pdf-thumbnails") !== undefined, true);
    const labels = byTestId(mounted.root, "pdf-thumbnail").map((node) => textContent(node));
    assert.deepEqual(labels, ["i", "ii", "iii", "iv"]);
    assert.equal(
      byTestId(mounted.root, "pdf-thumbnail").find((node) => classNames(node).includes("active"))
        ?.props?.["data-page-number"],
      1,
    );
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});
