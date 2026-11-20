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
// Reka UI primitives check refs against these constructors even when the
// custom renderer has no real DOM; the stubs keep mounting DOM-free.
if (globalThis.Element === undefined) globalThis.Element = class Element {};
if (globalThis.DocumentFragment === undefined) globalThis.DocumentFragment = class DocumentFragment {};

const { PptxViewer } = await importFromDemo("@arcships/vue-pptx");

const FIRST_SLIDE_NOTES = "Open with the quarter.\nThen pause for the chart.";
const THIRD_SLIDE_NOTES = "Close on the ask.";

function presenterDocument({ hiddenNext = false, secondSlideNotes = "Second slide notes." } = {}) {
  return {
    fileName: "presenter.pptx",
    width: 960,
    height: 540,
    slides: [
      { hidden: false, index: 0, number: 1, notes: FIRST_SLIDE_NOTES },
      {
        hidden: hiddenNext,
        index: 1,
        number: 2,
        ...(hiddenNext || secondSlideNotes === null ? {} : { notes: secondSlideNotes }),
      },
      { hidden: false, index: 2, number: 3, notes: THIRD_SLIDE_NOTES },
    ],
    warnings: [],
  };
}

function capability() {
  return { approximate: 0, discovered: 0, features: [], static: 0, strict: 0, unparsed: 0 };
}

function createSession({ document = presenterDocument() } = {}) {
  const calls = { dispose: 0, open: [], renderSlide: [], renderThumbnail: [], setZoom: [] };
  let currentSlideIndex = 0;
  let zoomPercent = 100;
  let snapshot = {
    activeNodeIds: [],
    blockedMediaIds: [],
    clickBoundary: 0,
    generation: 1,
    positionMs: 0,
    slideIndex: 0,
    status: "ready",
  };
  const listeners = new Set();
  const controller = {
    get snapshot() { return snapshot; },
    async next() {},
    async activateObject() { return false; },
    async previous() {},
    async play() {},
    pause() {},
    async resume() {},
    async reset() {},
    async goToSlide(index) {
      snapshot = { ...snapshot, slideIndex: index };
      currentSlideIndex = index;
    },
    async resumeBlockedMedia() {},
    subscribe(listener) {
      listeners.add(listener);
      listener({ type: "statechange", snapshot });
      listener({ type: "capability", report: capability() });
      return () => listeners.delete(listener);
    },
    dispose() {},
  };
  const session = {
    document: null,
    get currentSlideIndex() { return currentSlideIndex; },
    get zoomPercent() { return zoomPercent; },
    async open(_source, options) {
      calls.open.push(options);
      session.document = document;
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
    async setZoom(value) {
      calls.setZoom.push(value);
      zoomPercent = value;
    },
    cancel() {},
    dispose() { calls.dispose += 1; },
    playbackDocument: { capability: capability(), slides: [] },
    capabilityReport: capability(),
    createPlaybackController(options) {
      snapshot = { ...snapshot, slideIndex: options?.initialSlide ?? 0 };
      return controller;
    },
  };
  return {
    calls,
    controller,
    session,
    emitState(slideIndex) {
      snapshot = { ...snapshot, slideIndex };
      for (const listener of listeners) listener({ type: "statechange", snapshot });
    },
  };
}

function presenters(root) {
  return walk(root).filter((node) => node.props?.["data-testid"] === "pptx-presenter");
}

function previewedSlides(root) {
  return walk(root)
    .filter((node) => node.props?.["data-pptx-thumbnail-index"] !== undefined)
    .map((node) => node.props["data-pptx-thumbnail-index"]);
}

function notesParagraphs(root) {
  return walk(root)
    .filter((node) => node.props?.["data-testid"] === "pptx-presenter-paragraph")
    .map((node) => textContent(node).trim());
}

async function mountViewer(props = {}) {
  const fake = createSession({ document: props.document ?? presenterDocument() });
  const { document: _document, ...rest } = props;
  void _document;
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([1, 2, 3]),
    documentSessionFactory: () => fake.session,
    sessionFactory: () => fake.session,
    showSearch: false,
    showSidebar: false,
    autoplay: false,
    ...rest,
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX viewer ready",
  );
  return { ...mounted, fake };
}

test("present mode with showNotes renders the presenter rail with notes, next slide and clock", async () => {
  const mounted = await mountViewer({ mode: "present", showNotes: true });
  try {
    const rail = presenters(mounted.root)[0];
    assert.ok(rail, "the presenter rail is mounted");
    assert.equal(rail.props["data-presenter-mode"], "present");
    assert.match(rail.props["aria-label"], /Presenter view/);
    assert.ok(
      walk(mounted.root).some((node) => String(node.props?.class || "").includes("pptx-viewer--presenter")),
      "the viewer root marks the presenter layout",
    );
    assert.deepEqual(notesParagraphs(mounted.root), [
      "Open with the quarter.",
      "Then pause for the chart.",
    ]);
    assert.equal(textContent(findByTestId(mounted.root, "pptx-presenter-timer")).trim(), "00:00");
    assert.deepEqual(previewedSlides(mounted.root), [1], "the next slide is previewed");
    assert.equal(findByTestId(mounted.root, "pptx-presenter-last"), undefined);

    // A finished deck previews no next slide and says so.
    mounted.fake.emitState(2);
    await vue.nextTick();
    assert.deepEqual(notesParagraphs(mounted.root), ["Close on the ask."]);
    assert.ok(findByTestId(mounted.root, "pptx-presenter-last"), "the last slide has no preview");
    assert.deepEqual(previewedSlides(mounted.root), []);
  } finally {
    mounted.app.unmount();
  }
  assert.deepEqual(mounted.warnings, []);
});

test("the presenter clock counts while presenting", async (t) => {
  const realSetInterval = globalThis.setInterval;
  const realClearInterval = globalThis.clearInterval;
  const realNow = Date.now;
  const tickers = [];
  let now = 1_000_000;
  globalThis.setInterval = (callback) => {
    tickers.push(callback);
    return tickers.length;
  };
  globalThis.clearInterval = () => {};
  Date.now = () => now;
  t.after(() => {
    globalThis.setInterval = realSetInterval;
    globalThis.clearInterval = realClearInterval;
    Date.now = realNow;
  });

  const mounted = await mountViewer({ mode: "present", showNotes: true });
  try {
    assert.equal(tickers.length, 1, "one clock interval runs while presenting");
    now += 65_000;
    tickers.at(-1)();
    await vue.nextTick();
    assert.equal(textContent(findByTestId(mounted.root, "pptx-presenter-timer")).trim(), "01:05");
  } finally {
    mounted.app.unmount();
  }
  assert.deepEqual(mounted.warnings, []);
});

test("the presenter rail skips hidden slides unless they are shown", async () => {
  const skipping = await mountViewer({
    mode: "present",
    showNotes: true,
    document: presenterDocument({ hiddenNext: true }),
  });
  try {
    assert.deepEqual(previewedSlides(skipping.root), [2], "a hidden next slide is skipped");
  } finally {
    skipping.app.unmount();
  }

  const showing = await mountViewer({
    mode: "present",
    showNotes: true,
    showHiddenSlides: true,
    document: presenterDocument({ hiddenNext: true }),
  });
  try {
    assert.deepEqual(previewedSlides(showing.root), [1], "showHiddenSlides previews the hidden slide");
  } finally {
    showing.app.unmount();
  }
});

test("browse mode shows only the current slide's notes", async () => {
  const mounted = await mountViewer({ showNotes: true });
  try {
    const rail = presenters(mounted.root)[0];
    assert.ok(rail, "the notes panel is mounted in browse mode");
    assert.equal(rail.props["data-presenter-mode"], "browse");
    assert.match(rail.props["aria-label"], /Speaker notes/);
    assert.deepEqual(notesParagraphs(mounted.root), [
      "Open with the quarter.",
      "Then pause for the chart.",
    ]);
    assert.equal(findByTestId(mounted.root, "pptx-presenter-timer"), undefined);
    assert.deepEqual(previewedSlides(mounted.root), []);
  } finally {
    mounted.app.unmount();
  }
  assert.deepEqual(mounted.warnings, []);
});

test("showNotes is opt-in and a slide without notes says so", async () => {
  const hidden = await mountViewer({ mode: "present" });
  try {
    assert.deepEqual(presenters(hidden.root), [], "notes stay off without showNotes");
  } finally {
    hidden.app.unmount();
  }
  assert.deepEqual(hidden.warnings, []);

  const mounted = await mountViewer({
    mode: "present",
    showNotes: true,
    document: presenterDocument({ secondSlideNotes: null }),
  });
  try {
    mounted.fake.emitState(1);
    await vue.nextTick();
    assert.deepEqual(notesParagraphs(mounted.root), []);
    assert.ok(findByTestId(mounted.root, "pptx-presenter-empty"), "a slide without notes shows the empty state");
  } finally {
    mounted.app.unmount();
  }
  assert.deepEqual(mounted.warnings, []);
});

test("the presenter rail is localized", async () => {
  const mounted = await mountViewer({ mode: "present", showNotes: true, locale: "zh-CN" });
  try {
    const rail = presenters(mounted.root)[0];
    assert.equal(rail.props["aria-label"], "演讲者视图");
    assert.match(textContent(findByTestId(mounted.root, "pptx-presenter-notes")), /备注/);
    assert.equal(textContent(findByTestId(mounted.root, "pptx-presenter-timer")).trim(), "00:00");
  } finally {
    mounted.app.unmount();
  }
  assert.deepEqual(mounted.warnings, []);
});
