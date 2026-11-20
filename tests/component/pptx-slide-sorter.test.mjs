import assert from "node:assert/strict";
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
// Reka UI primitives check refs against these constructors even when the
// custom renderer has no real DOM; the stubs keep mounting DOM-free.
if (globalThis.Element === undefined) globalThis.Element = class Element {};
if (globalThis.DocumentFragment === undefined) globalThis.DocumentFragment = class DocumentFragment {};

const { PptxViewer } = await importFromDemo("@arcships/vue-pptx");

function documentWithSlides(count) {
  return {
    fileName: "sorter.pptx",
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

function createSession({ slideCount = 5 } = {}) {
  const calls = { dispose: 0, open: [], renderSlide: [], renderThumbnail: [], setZoom: [] };
  let currentSlideIndex = 0;
  let zoomPercent = 100;
  const session = {
    document: null,
    get currentSlideIndex() { return currentSlideIndex; },
    get zoomPercent() { return zoomPercent; },
    async open(_source, options) {
      calls.open.push(options);
      session.document = documentWithSlides(slideCount);
      currentSlideIndex = options?.initialSlide ?? 0;
      return session.document;
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
    async setZoom(value) {
      calls.setZoom.push(value);
      zoomPercent = value;
    },
    cancel() {},
    dispose() { calls.dispose += 1; },
  };
  return { calls, session };
}

function classNames(node) {
  return String(node?.props?.class || "").split(/\s+/).filter(Boolean);
}

function nodesWithClass(root, className) {
  return walk(root).filter((node) => classNames(node).includes(className));
}

function sorterItems(root) {
  return walk(root).filter((node) => node.props?.["data-testid"] === "pptx-sorter-item");
}

function activeSorterItem(root) {
  return sorterItems(root).find((node) => node.props["data-active"] === "true");
}

async function mountViewer(props = {}) {
  const fake = createSession(props.slideCount === undefined ? {} : { slideCount: props.slideCount });
  const { slideCount, ...rest } = props;
  void slideCount;
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([1, 2, 3]),
    sessionFactory: () => fake.session,
    showSearch: false,
    showSidebar: false,
    ...rest,
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX viewer ready",
  );
  return { ...mounted, calls: fake.calls };
}

test("PPTX sorter renders a numbered slide grid in browse mode", async () => {
  const mounted = await mountViewer({ slideCount: 5 });
  const sorter = findByTestId(mounted.root, "pptx-sorter");
  assert.equal(sorter, undefined, "single view is the default browse layout");

  const sorterButton = findByTestId(mounted.root, "pptx-view-mode-sorter");
  assert.notEqual(sorterButton.props["aria-pressed"], "true", "single view is pressed, not the sorter");
  sorterButton.props.onClick();
  await waitFor(() => findByTestId(mounted.root, "pptx-sorter") !== undefined, "slide sorter grid");

  const grid = findByTestId(mounted.root, "pptx-sorter-grid") ?? nodesWithClass(mounted.root, "pptx-viewer__sorter-grid")[0];
  assert.equal(grid.props.style.gridTemplateColumns, "repeat(3, minmax(0, 1fr))");
  assert.deepEqual(
    sorterItems(mounted.root).map((item) => item.props["data-slide-index"]),
    [0, 1, 2, 3, 4],
    "every slide gets a numbered sorter cell",
  );
  const thumbnails = nodesWithClass(mounted.root, "pptx-thumbnail")
    .filter((node) => node.type === "button")
    .map((node) => node.props["data-pptx-thumbnail-index"]);
  assert.deepEqual(thumbnails, [0, 1, 2, 3, 4]);
  assert.equal(
    activeSorterItem(mounted.root)?.props["data-slide-index"],
    0,
    "the current slide stays selected when the sorter opens",
  );
  assert.equal(
    findByTestId(mounted.root, "pptx-view-mode-sorter").props["aria-pressed"],
    "true",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PPTX sorter selection stays in sync with the slide and the single view", async () => {
  const mounted = await mountViewer({ slideCount: 5 });
  findByTestId(mounted.root, "pptx-view-mode-sorter").props.onClick();
  await waitFor(() => findByTestId(mounted.root, "pptx-sorter") !== undefined, "slide sorter grid");

  // Selecting a sorter cell moves the deck to that slide without closing the grid.
  const third = sorterItems(mounted.root).find((item) => item.props["data-slide-index"] === 3);
  const thumbnail = walk(third).find(
    (node) => node.type === "button" && node.props?.class?.includes("pptx-thumbnail"),
  );
  thumbnail.props.onClick();
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer").props["data-slide-index"] === 3,
    "the sorter selection drives the active slide",
  );
  assert.equal(activeSorterItem(mounted.root)?.props["data-slide-index"], 3);
  assert.ok(findByTestId(mounted.root, "pptx-sorter"), "the grid stays open after selecting a slide");

  // Switching back to the single view keeps the same slide selected.
  findByTestId(mounted.root, "pptx-view-mode-single").props.onClick();
  await waitFor(() => findByTestId(mounted.root, "pptx-sorter") === undefined, "sorter closed");
  assert.equal(
    findByTestId(mounted.root, "pptx-viewer").props["data-slide-index"],
    3,
    "switching browse layouts keeps the current slide",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PPTX sorter only mounts the slides around the visible rows", async () => {
  const mounted = await mountViewer({ slideCount: 40 });
  findByTestId(mounted.root, "pptx-view-mode-sorter").props.onClick();
  await waitFor(() => findByTestId(mounted.root, "pptx-sorter") !== undefined, "slide sorter grid");

  const sorter = findByTestId(mounted.root, "pptx-sorter");
  const initial = sorterItems(mounted.root);
  assert.ok(initial.length < 40, "a long deck does not mount every sorter cell");
  assert.equal(initial[0].props["data-slide-index"], 0);

  // Scrolling the grid moves the mounted window without changing the slide.
  sorter.scrollTop = 600;
  sorter.clientHeight = 640;
  sorter.clientWidth = 500;
  sorter.props.onScrollPassive({ target: sorter });
  await vue.nextTick();
  const scrolled = sorterItems(mounted.root);
  assert.ok(
    scrolled[0].props["data-slide-index"] > 0,
    "the sorter window follows the scroll offset",
  );
  assert.equal(
    findByTestId(mounted.root, "pptx-viewer").props["data-slide-index"],
    0,
    "scrolling the sorter does not change the active slide",
  );
  assert.equal(sorter.scrollTop, 600);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PPTX sorter keeps browse keys scrolling instead of changing slides", async () => {
  const mounted = await mountViewer({ slideCount: 30 });
  findByTestId(mounted.root, "pptx-view-mode-sorter").props.onClick();
  await waitFor(() => findByTestId(mounted.root, "pptx-sorter") !== undefined, "slide sorter grid");

  const viewer = findByTestId(mounted.root, "pptx-viewer");
  const sorter = findByTestId(mounted.root, "pptx-sorter");
  sorter.scrollTop = 0;
  sorter.scrollHeight = 4000;
  sorter.clientHeight = 640;
  const { event, preventedCount } = keydownEvent("PageDown");
  for (const handler of Array.isArray(viewer.props.onKeydown) ? viewer.props.onKeydown : [viewer.props.onKeydown]) {
    handler?.(event);
  }
  assert.equal(preventedCount(), 1, "the sorter consumes the page-scroll key");
  assert.equal(
    findByTestId(mounted.root, "pptx-viewer").props["data-slide-index"],
    0,
    "PageDown scrolls the grid instead of moving to the next slide",
  );
  assert.ok(sorter.scrollTop > 0, "the grid scrolled by a page");
  mounted.app.unmount();
});
