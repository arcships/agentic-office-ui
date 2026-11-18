import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
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

// Vue's native v-model directive checks these browser constructors during an
// update. The custom component renderer intentionally has no browser DOM.
if (globalThis.Document === undefined) globalThis.Document = class Document {};
if (globalThis.ShadowRoot === undefined) globalThis.ShadowRoot = class ShadowRoot {};

const {
  PptxStage,
  PptxViewer,
  usePptxDocument,
  usePptxPlayback,
} = await importFromDemo("@arcships/vue-pptx");

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((nextResolve, nextReject) => {
    resolve = nextResolve;
    reject = nextReject;
  });
  return { promise, reject, resolve };
}

function documentWithSlides(count, fileName = "preview.pptx", warnings = []) {
  return {
    fileName,
    height: 540,
    width: 960,
    slides: Array.from({ length: count }, (_, index) => ({
      hidden: index === 1,
      index,
      number: index + 1,
    })),
    warnings,
  };
}

function createSession({ document = documentWithSlides(3), open, renderSlide, withRefresh = false } = {}) {
  const calls = {
    clearSearchHighlights: 0,
    dispose: 0,
    open: [],
    refresh: 0,
    renderSlide: [],
    renderThumbnail: [],
    setZoom: [],
  };
  let currentSlideIndex = 0;
  let zoomPercent = 100;
  const session = {
    document: null,
    get currentSlideIndex() { return currentSlideIndex; },
    get zoomPercent() { return zoomPercent; },
    async open(source, options) {
      calls.open.push(options);
      const loaded = open ? await open(source, options) : document;
      session.document = loaded;
      currentSlideIndex = options?.initialSlide ?? 0;
      return loaded;
    },
    async renderSlide(index) {
      calls.renderSlide.push(index);
      if (renderSlide) await renderSlide(index, session);
      currentSlideIndex = index;
    },
    renderThumbnail(index, target, width) {
      calls.renderThumbnail.push({ index, target, width });
      return null;
    },
    ...(withRefresh ? { async refresh() { calls.refresh += 1; } } : {}),
    searchText() { return []; },
    async highlightSearchResult() { return null; },
    clearSearchHighlights() { calls.clearSearchHighlights += 1; },
    async setZoom(value) {
      calls.setZoom.push(value);
      zoomPercent = value;
    },
    cancel() {},
    dispose() { calls.dispose += 1; },
  };
  return { calls, session };
}

function createPlaybackSession({ document = documentWithSlides(3), open, renderSlide } = {}) {
  const preview = createSession({ document, open, renderSlide });
  const listeners = new Set();
  const calls = {
    ...preview.calls,
    activateObject: [],
    createPlaybackController: 0,
    createPlaybackControllerOptions: [],
    disposeController: 0,
    goToSlide: [],
    next: 0,
    pause: 0,
    play: 0,
    previous: 0,
    reset: 0,
    resume: 0,
    resumeBlockedMedia: [],
  };
  const capability = {
    approximate: 1,
    discovered: 5,
    features: [],
    static: 1,
    strict: 3,
    unparsed: 0,
  };
  let snapshot = {
    activeNodeIds: [],
    blockedMediaIds: [],
    clickBoundary: 0,
    generation: 1,
    positionMs: 0,
    slideIndex: 0,
    status: "ready",
  };
  const controller = {
    get snapshot() { return snapshot; },
    async next() { calls.next += 1; },
    async activateObject(key) { calls.activateObject.push(key); return false; },
    async previous() { calls.previous += 1; },
    async play() { calls.play += 1; },
    pause() { calls.pause += 1; },
    async resume() { calls.resume += 1; },
    async reset() { calls.reset += 1; },
    async goToSlide(index, options) { calls.goToSlide.push([index, options]); },
    async resumeBlockedMedia(mediaId) { calls.resumeBlockedMedia.push(mediaId); },
    subscribe(listener) {
      listeners.add(listener);
      listener({ type: "statechange", snapshot });
      listener({ type: "capability", report: capability });
      return () => listeners.delete(listener);
    },
    dispose() { calls.disposeController += 1; },
  };
  const session = {
    ...preview.session,
    playbackDocument: { capability, slides: [] },
    capabilityReport: capability,
    createPlaybackController(options) {
      calls.createPlaybackController += 1;
      calls.createPlaybackControllerOptions.push(options);
      snapshot = { ...snapshot, slideIndex: options?.initialSlide ?? 0 };
      return controller;
    },
    dispose() {
      calls.dispose += 1;
      preview.session.dispose();
    },
  };
  return {
    calls,
    capability,
    controller,
    emit(event) {
      if (event.type === "statechange") snapshot = event.snapshot;
      for (const listener of listeners) listener(event);
    },
    session,
  };
}

function button(root, label) {
  return walk(root).find((node) => node.type === "button" && node.props?.["aria-label"] === label);
}

function pressKey(node, key, init = {}) {
  const { event, preventedCount } = keydownEvent(key, init);
  const handlers = Array.isArray(node.props?.onKeydown)
    ? node.props.onKeydown
    : [node.props?.onKeydown];
  for (const handler of handlers) handler?.(event);
  return { event, prevented: preventedCount() };
}

function classNames(node) {
  return String(node?.props?.class || "").split(/\s+/).filter(Boolean);
}

function nodesWithClass(root, className) {
  return walk(root).filter((node) => classNames(node).includes(className));
}

function thumbnails(root) {
  return nodesWithClass(root, "pptx-thumbnail").filter((node) => node.type === "button");
}

function installIntersectionObserver(t) {
  const original = globalThis.IntersectionObserver;
  const instances = [];
  class StubIntersectionObserver {
    constructor(callback, options) {
      this.callback = callback;
      this.options = options;
      this.targets = new Set();
      this.disconnected = false;
      instances.push(this);
    }

    observe(target) {
      this.targets.add(target);
    }

    unobserve(target) {
      this.targets.delete(target);
    }

    disconnect() {
      this.disconnected = true;
      this.targets.clear();
    }

    intersect(target) {
      this.targets.delete(target);
      this.callback([{ target, isIntersecting: true }]);
    }
  }
  globalThis.IntersectionObserver = StubIntersectionObserver;
  t.after(() => {
    if (original) globalThis.IntersectionObserver = original;
    else delete globalThis.IntersectionObserver;
  });
  return instances;
}

function installFontFaceSet(t) {
  const originalDocument = globalThis.document;
  const listeners = new Set();
  let resolveReady;
  const ready = new Promise((resolve) => { resolveReady = resolve; });
  const fonts = {
    ready,
    status: "loading",
    addEventListener(type, listener) {
      if (type === "loadingdone") listeners.add(listener);
    },
    removeEventListener(type, listener) {
      if (type === "loadingdone") listeners.delete(listener);
    },
  };
  globalThis.document = { fonts };
  t.after(() => {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
  });
  return {
    emitLoadingDone() {
      for (const listener of [...listeners]) listener({ type: "loadingdone" });
    },
    listenerCount: () => listeners.size,
    resolveReady: () => resolveReady(fonts),
  };
}

test("PptxStage and composables expose the minimal document and playback controls", async () => {
  const playback = createPlaybackSession({ document: documentWithSlides(2) });
  let documentApi;
  let playbackApi;
  let opening;
  const Harness = vue.defineComponent({
    setup() {
      const stage = vue.ref(null);
      const element = vue.computed(() => stage.value?.element ?? null);
      documentApi = usePptxDocument(element, {
        factory: () => playback.session,
      });
      playbackApi = usePptxPlayback(documentApi, { autoplay: false });
      opening = documentApi.open(new Uint8Array([1, 2, 3]));
      return () => vue.h(PptxStage, { ref: stage, class: "custom-stage", tabindex: 0 });
    },
  });

  const mounted = await mount(Harness);
  await opening;
  await waitFor(() => documentApi.state.value === "ready", "headless PPTX document ready");
  await waitFor(() => playbackApi.controller.value !== null, "headless PPTX playback ready");

  const stage = findByTestId(mounted.root, "pptx-stage");
  assert.match(stage.props.class, /pptx-stage/);
  assert.match(stage.props.class, /custom-stage/);
  assert.equal(documentApi.document.value.slides.length, 2);
  assert.equal(documentApi.getSession(), playback.session);

  await documentApi.goTo(1);
  assert.equal(documentApi.activeIndex.value, 1);
  await playbackApi.next();
  assert.equal(playback.calls.next, 1);

  mounted.app.unmount();
  assert.equal(playback.calls.dispose, 1);
  assert.deepEqual(mounted.warnings, []);
});

test("PptxStage emits complete slide, object and context-menu interactions", async () => {
  const selections = [];
  const objects = [];
  const menus = [];
  const mounted = await mount(PptxStage, {
    onSelectionChange: (selection) => selections.push(selection),
    onObjectClick: (object) => objects.push(object),
    onContextMenu: (menu) => menus.push(menu),
  });
  const stage = findByTestId(mounted.root, "pptx-stage");
  stage.getBoundingClientRect = () => ({ left: 10, top: 20 });
  const slide = { dataset: { slideIndex: "2" } };
  const object = { dataset: { pptxObjectKey: "slide-3/object-7" } };
  const slideTarget = {
    closest: (selector) => selector === "[data-slide-index]" ? slide : null,
  };
  const objectTarget = {
    closest: (selector) => selector === "[data-slide-index]" ? slide : object,
  };

  stage.props.onClick({ target: slideTarget });
  stage.props.onClick({ target: objectTarget });
  stage.props.onContextmenu({
    target: objectTarget,
    clientX: 110,
    clientY: 220,
    preventDefault() {},
  });

  assert.deepEqual(selections, [
    { kind: "slide", slideIndex: 2 },
    { kind: "slide", slideIndex: 2 },
  ]);
  assert.deepEqual(objects, [{
    kind: "object",
    slideIndex: 2,
    objectKey: "slide-3/object-7",
  }]);
  assert.deepEqual(menus, [{
    kind: "object",
    slideIndex: 2,
    objectKey: "slide-3/object-7",
    clientX: 110,
    clientY: 220,
    containerX: 100,
    containerY: 200,
  }]);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("PptxStage only consumes controlled gesture zoom wheels", async () => {
  const zooms = [];
  const mounted = await mount(PptxStage, {
    zoom: 1,
    "onUpdate:zoom": (zoom) => zooms.push(zoom),
  });
  const stage = findByTestId(mounted.root, "pptx-stage");
  let prevented = 0;
  const wheel = (ctrlKey, deltaY) => stage.dispatchEvent({
    type: "wheel",
    ctrlKey,
    deltaY,
    deltaMode: 0,
    clientX: 10,
    clientY: 20,
    preventDefault: () => { prevented += 1; },
  });

  wheel(false, -20);
  assert.equal(prevented, 0);
  assert.deepEqual(zooms, []);

  wheel(true, -20);
  assert.equal(prevented, 1);
  assert.equal(zooms.length, 1);
  assert.ok(zooms[0] > 1);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer defaults to English copy and exposes translated aria-labels", async () => {
  const mounted = await mount(PptxViewer);
  assert.match(textContent(mounted.root), /Open a PPTX file to get started/);
  assert.ok(button(mounted.root, "Next slide"));
  assert.ok(button(mounted.root, "Previous slide"));
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PptxViewer exposes loading, ready navigation, hidden-page marker and cleanup", async () => {
  const pending = deferred();
  const { calls, session } = createSession({ open: () => pending.promise });
  const events = [];
  let sessionOptions;
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([1, 2, 3]),
    sessionFactory: (_element, options) => {
      sessionOptions = options;
      return session;
    },
    showSearch: false,
    showSidebar: false,
    locale: "zh-CN",
    onLoadStart: () => events.push("start"),
    onLoadSuccess: (document) => events.push(["success", document.slides.length]),
    onSlideChange: (index) => events.push(["slide", index]),
  });

  assert.equal(findByTestId(mounted.root, "pptx-viewer").props["data-state"], "loading");
  pending.resolve(documentWithSlides(3));
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX ready state",
  );
  assert.match(textContent(findByTestId(mounted.root, "pptx-page-counter")), /1 \/ 3/);
  assert.deepEqual(events, ["start", ["success", 3], ["slide", 0]]);
  assert.equal(sessionOptions.renderMode, "list");
  assert.equal(sessionOptions.listOptions.windowed, true);

  sessionOptions.onSlideChange(1);
  await vue.nextTick();
  assert.match(textContent(findByTestId(mounted.root, "pptx-page-counter")), /2 \/ 3/);
  assert.match(textContent(mounted.root), /隐藏页/);

  button(mounted.root, "下一页").props.onClick();
  await waitFor(() => calls.renderSlide.includes(2), "PPTX next slide scroll");
  assert.match(textContent(findByTestId(mounted.root, "pptx-page-counter")), /3 \/ 3/);

  button(mounted.root, "放大").props.onClick();
  await waitFor(() => calls.setZoom.includes(125), "PPTX zoom");
  assert.match(textContent(mounted.root), /125%/);

  mounted.app.unmount();
  assert.equal(calls.dispose, 1);
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer keeps only the newest file result", async () => {
  const first = deferred();
  const second = deferred();
  const sessions = [
    createSession({ open: () => first.promise }),
    createSession({ open: () => second.promise }),
  ];
  const source = vue.ref(new Uint8Array([1]));
  const loaded = [];
  let factoryIndex = 0;
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(PptxViewer, {
        source: source.value,
        sessionFactory: () => sessions[factoryIndex++].session,
        showSearch: false,
        showSidebar: false,
        onLoadSuccess: (document) => loaded.push(document.fileName),
      });
    },
  });
  const mounted = await mount(Harness);

  source.value = new Uint8Array([2]);
  await waitFor(() => factoryIndex === 2, "second PPTX load");
  second.resolve(documentWithSlides(2, "second.pptx"));
  await waitFor(() => loaded.includes("second.pptx"), "newest PPTX result");
  first.resolve(documentWithSlides(1, "first.pptx"));
  await new Promise((resolve) => setTimeout(resolve, 20));

  assert.deepEqual(loaded, ["second.pptx"]);
  assert.equal(sessions[0].calls.dispose, 1);
  assert.equal(findByTestId(mounted.root, "pptx-viewer").props["data-state"], "ready");
  mounted.app.unmount();
  assert.equal(sessions[1].calls.dispose, 1);
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer renders a load failure without a stale success", async () => {
  const failure = new Error("broken PPTX");
  const { session } = createSession({ open: async () => { throw failure; } });
  const errors = [];
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([9]),
    sessionFactory: () => session,
    showSearch: false,
    showSidebar: false,
    onLoadError: (error) => errors.push(error),
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "error",
    "PPTX error state",
  );
  assert.match(textContent(findByTestId(mounted.root, "pptx-error")), /broken PPTX/);
  assert.equal(errors[0].code, "PARSE_FAILED");
  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer present mode wires playback controls and forwards playback events", async () => {
  const playback = createPlaybackSession();
  const events = [];
  let sessionOptions;
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([7, 8, 9]),
    mode: "present",
    autoplay: false,
    locale: "zh-CN",
    documentSessionFactory: (_element, options) => {
      sessionOptions = options;
      return playback.session;
    },
    onPlaybackReady: (controller) => events.push(["ready", controller === playback.controller]),
    onPlaybackStateChange: (snapshot) => events.push(["state", snapshot.status]),
    onStepChange: (slideIndex, boundary) => events.push(["step", slideIndex, boundary]),
    onPlaybackWarning: (warning) => events.push(["warning", warning.code]),
    onCapability: (report) => events.push(["capability", report.strict]),
    onMediaRequest: (mediaId) => events.push(["media", mediaId]),
    onAction: (action) => events.push(["action", action.kind]),
    onPlaybackError: (error) => events.push(["error", error.code]),
  });

  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX playback ready state",
  );
  assert.equal(playback.calls.createPlaybackController, 1);
  assert.equal(sessionOptions.renderMode, "slide");
  assert.deepEqual(events.slice(0, 3), [
    ["ready", true],
    ["state", "ready"],
    ["capability", 3],
  ]);
  assert.match(textContent(mounted.root), /精确 3\/5/);
  assert.equal(button(mounted.root, "下一页"), undefined);

  button(mounted.root, "下一步").props.onClick();
  button(mounted.root, "上一步").props.onClick();
  button(mounted.root, "播放").props.onClick();
  button(mounted.root, "重新播放").props.onClick();
  await waitFor(
    () => playback.calls.next === 1 && playback.calls.previous === 1 && playback.calls.play === 1 && playback.calls.reset === 1,
    "PPTX playback control calls",
  );

  playback.emit({ type: "stepchange", slideIndex: 0, boundary: 2 });
  playback.emit({
    type: "warning",
    warning: { code: "UNSUPPORTED", message: "test", recoverable: true, slideIndex: 0 },
  });
  playback.emit({ type: "mediarequest", mediaId: "media-1", reason: "autoplay-blocked" });
  playback.emit({
    type: "statechange",
    snapshot: {
      ...playback.controller.snapshot,
      blockedMediaIds: ["media-1"],
      status: "blocked",
    },
  });
  playback.emit({
    type: "action",
    action: { kind: "open-url", sourceObjectKey: "slide:1/object:2", url: "https://example.test" },
  });
  await vue.nextTick();
  assert.deepEqual(events.filter((event) => event[0] !== "state").slice(-4), [
    ["step", 0, 2],
    ["warning", "UNSUPPORTED"],
    ["media", "media-1"],
    ["action", "open-url"],
  ]);
  button(mounted.root, "继续播放媒体").props.onClick();
  await waitFor(() => playback.calls.resumeBlockedMedia.length === 1, "blocked media resume control");
  button(mounted.root, "全屏").props.onClick();
  await waitFor(() => events.some((event) => event[0] === "error"), "fullscreen rejection event");
  assert.ok(events.some((event) => event[0] === "error" && event[1] === "FULLSCREEN_REJECTED"));

  mounted.app.unmount();
  assert.equal(playback.calls.dispose, 1);
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer keeps the deck reachable when a single slide fails to render", async () => {
  const broken = new Set([2]);
  const { calls, session } = createSession({
    document: documentWithSlides(4),
    renderSlide(index) {
      if (broken.has(index)) throw new Error(`slide ${index} broke`);
    },
  });
  const failures = [];
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([1, 2, 3]),
    sessionFactory: () => session,
    showSearch: false,
    showSidebar: false,
    locale: "zh-CN",
    onSlideError: (error, index) => failures.push([index, error.message]),
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX ready state",
  );
  button(mounted.root, "下一页").props.onClick();
  await waitFor(() => calls.renderSlide.includes(1), "second slide render");
  button(mounted.root, "下一页").props.onClick();
  await waitFor(() => failures.length === 1, "slide failure event");

  const viewer = findByTestId(mounted.root, "pptx-viewer");
  assert.equal(viewer.props["data-state"], "ready");
  assert.equal(viewer.props["data-slide-error"], 2);
  assert.deepEqual(failures, [[2, "页面渲染失败：slide 2 broke"]]);
  assert.match(textContent(findByTestId(mounted.root, "pptx-slide-error")), /第 3 页渲染失败/);
  // The previous slide stays on screen instead of being replaced by a fatal error.
  assert.match(textContent(findByTestId(mounted.root, "pptx-page-counter")), /2 \/ 4/);
  assert.equal(findByTestId(mounted.root, "pptx-stage").props["aria-hidden"], false);
  assert.equal(findByTestId(mounted.root, "pptx-error"), undefined);

  // Retry fails again and keeps the deck usable.
  findByTestId(mounted.root, "pptx-slide-retry").props.onClick();
  await waitFor(() => failures.length === 2, "slide retry failure");
  assert.deepEqual(failures, [
    [2, "页面渲染失败：slide 2 broke"],
    [2, "页面渲染失败：slide 2 broke"],
  ]);

  // Once the slide renders, retry clears the banner and navigation continues.
  broken.delete(2);
  findByTestId(mounted.root, "pptx-slide-retry").props.onClick();
  await waitFor(() => findByTestId(mounted.root, "pptx-slide-error") === undefined, "slide retry success");
  assert.equal(findByTestId(mounted.root, "pptx-viewer").props["data-slide-index"], 2);

  // Skipping the failed slide jumps to the next one instead of stalling.
  broken.add(3);
  button(mounted.root, "下一页").props.onClick();
  await waitFor(() => findByTestId(mounted.root, "pptx-slide-error") !== undefined, "second slide failure");
  assert.equal(findByTestId(mounted.root, "pptx-viewer").props["data-slide-error"], 3);
  findByTestId(mounted.root, "pptx-slide-skip").props.onClick();
  await waitFor(() => findByTestId(mounted.root, "pptx-slide-error") === undefined, "skipped slide");
  assert.equal(findByTestId(mounted.root, "pptx-viewer").props["data-state"], "ready");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer keeps the current slide when switching between browse and present", async () => {
  const playback = createPlaybackSession({ document: documentWithSlides(4) });
  const source = new Uint8Array([4, 5, 6]);
  const mode = vue.ref("browse");
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(PptxViewer, {
        source,
        mode: mode.value,
        autoplay: false,
        sessionFactory: () => playback.session,
        documentSessionFactory: () => playback.session,
        showSearch: false,
        showSidebar: false,
        locale: "zh-CN",
      });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "browse ready state",
  );

  button(mounted.root, "下一页").props.onClick();
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-slide-index"] === 1,
    "browse slide 2",
  );
  button(mounted.root, "下一页").props.onClick();
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-slide-index"] === 2,
    "browse slide 3",
  );

  mode.value = "present";
  await waitFor(
    () => playback.calls.createPlaybackControllerOptions.length === 1,
    "playback controller for present mode",
  );
  assert.equal(playback.calls.open.at(-1).initialSlide, 2);
  assert.equal(playback.calls.createPlaybackControllerOptions[0].initialSlide, 2);
  assert.equal(findByTestId(mounted.root, "pptx-viewer").props["data-slide-index"], 2);
  // Present mode letterboxes using the deck aspect ratio instead of clipping.
  await waitFor(
    () => findByTestId(mounted.root, "pptx-stage").props.style?.["--pptx-slide-aspect"] === "960 / 540",
    "letterboxed present stage",
  );
  assert.deepEqual(findByTestId(mounted.root, "pptx-stage").props.style, {
    "--pptx-slide-aspect": "960 / 540",
    "--pptx-slide-ratio": "1.7778",
  });

  mode.value = "browse";
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready"
      && !findByTestId(mounted.root, "pptx-viewer").props.class.includes("pptx-viewer--present"),
    "browse mode restored",
  );
  assert.equal(playback.calls.open.at(-1).initialSlide, 2);
  assert.equal(findByTestId(mounted.root, "pptx-viewer").props["data-slide-index"], 2);
  assert.equal(
    findByTestId(mounted.root, "pptx-stage").props.style?.["--pptx-slide-aspect"],
    undefined,
  );

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer surfaces a cancelled load instead of spinning forever", async () => {
  const { PptxPreviewError } = await importFromDemo("@arcships/pptx-core");
  let attempts = 0;
  const { session } = createSession({
    open: async () => {
      attempts += 1;
      if (attempts === 1) throw new PptxPreviewError("ABORTED", "PPTX 加载已取消。");
      return documentWithSlides(2);
    },
  });
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([7]),
    sessionFactory: () => session,
    showSearch: false,
    showSidebar: false,
    locale: "zh-CN",
  });

  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "cancelled",
    "cancelled state",
  );
  assert.equal(findByTestId(mounted.root, "pptx-loading"), undefined);
  assert.match(textContent(findByTestId(mounted.root, "pptx-cancelled")), /已取消/);

  findByTestId(mounted.root, "pptx-cancelled-retry").props.onClick();
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "cancelled load retried",
  );
  assert.equal(attempts, 2);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer stops loading when a load times out", async () => {
  const { calls, session } = createSession({ open: () => new Promise(() => {}) });
  const errors = [];
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([8]),
    sessionFactory: () => session,
    openTimeoutMs: 20,
    showSearch: false,
    showSidebar: false,
    locale: "zh-CN",
    onLoadError: (error) => errors.push(error.code),
  });

  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "error",
    "load timeout error state",
  );
  assert.deepEqual(errors, ["TIMEOUT"]);
  assert.match(textContent(findByTestId(mounted.root, "pptx-error")), /超过 20 毫秒/);
  assert.equal(findByTestId(mounted.root, "pptx-loading"), undefined);
  assert.equal(calls.dispose, 1);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer reports document command failures instead of leaking rejections", async () => {
  const { session } = createSession();
  session.setZoom = async () => { throw new Error("zoom failed"); };
  const failures = [];
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([9, 9]),
    sessionFactory: () => session,
    showSearch: false,
    showSidebar: false,
    locale: "zh-CN",
    onDocumentError: (error) => failures.push(error.message),
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX ready state",
  );

  button(mounted.root, "放大").props.onClick();
  await waitFor(() => failures.length === 1, "zoom failure event");
  assert.deepEqual(failures, ["zoom failed"]);
  assert.equal(findByTestId(mounted.root, "pptx-viewer").props["data-state"], "ready");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer clears the inline height while fullscreen", async () => {
  const listeners = new Map();
  const fakeDocument = {
    fullscreenElement: null,
    addEventListener(type, listener) { listeners.set(type, listener); },
    removeEventListener(type, listener) {
      if (listeners.get(type) === listener) listeners.delete(type);
    },
  };
  globalThis.document = fakeDocument;
  try {
    const { session } = createSession();
    const mounted = await mount(PptxViewer, {
      source: new Uint8Array([1]),
      height: "calc(100vh - 190px)",
      sessionFactory: () => session,
      showSearch: false,
      showSidebar: false,
      locale: "zh-CN",
    });
    await waitFor(
      () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
      "PPTX ready state",
    );
    assert.deepEqual(findByTestId(mounted.root, "pptx-viewer").props.style, {
      height: "calc(100vh - 190px)",
      maxHeight: "calc(100vh - 190px)",
    });

    // The test renderer hands template refs a copy of the host node, so the
    // fullscreen element stands in for a container that holds the viewer.
    fakeDocument.fullscreenElement = { contains: () => true };
    listeners.get("fullscreenchange")();
    await vue.nextTick();
    assert.equal(findByTestId(mounted.root, "pptx-viewer").props.style, undefined);

    fakeDocument.fullscreenElement = null;
    listeners.get("fullscreenchange")();
    await vue.nextTick();
    assert.deepEqual(findByTestId(mounted.root, "pptx-viewer").props.style, {
      height: "calc(100vh - 190px)",
      maxHeight: "calc(100vh - 190px)",
    });

    mounted.app.unmount();
    assert.deepEqual(mounted.warnings, []);
  } finally {
    delete globalThis.document;
  }
});

test("PptxViewer renders document warnings and lets the host replace them", async () => {
  const warnings = [{
    code: "NOTES_UNSUPPORTED",
    message: "当前静态底座尚未解析演讲者备注。",
  }];
  const { session } = createSession({ document: documentWithSlides(2, "notes.pptx", warnings) });

  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([2]),
    sessionFactory: () => session,
    showSearch: false,
    showSidebar: false,
    locale: "zh-CN",
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX ready state",
  );
  assert.match(textContent(findByTestId(mounted.root, "pptx-warnings")), /演讲者备注/);
  mounted.app.unmount();

  const hidden = await mount(PptxViewer, {
    source: new Uint8Array([2]),
    sessionFactory: () => createSession({ document: documentWithSlides(2, "notes.pptx", warnings) }).session,
    showSearch: false,
    showSidebar: false,
    showWarnings: false,
    locale: "zh-CN",
  });
  await waitFor(
    () => findByTestId(hidden.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX ready state without warnings",
  );
  assert.equal(findByTestId(hidden.root, "pptx-warnings"), undefined);
  hidden.app.unmount();

  const slotted = await mount(vue.defineComponent({
    setup() {
      return () => vue.h(
        PptxViewer,
        {
          source: new Uint8Array([3]),
          sessionFactory: () => createSession({ document: documentWithSlides(2, "notes.pptx", warnings) }).session,
          showSearch: false,
          showSidebar: false,
        },
        { warnings: ({ warnings: list }) => vue.h("p", { class: "custom-warning" }, `${list.length} warnings`) },
      );
    },
  }));
  await waitFor(
    () => findByTestId(slotted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX ready state with warning slot",
  );
  assert.match(textContent(findByTestId(slotted.root, "pptx-warnings")), /1 warnings/);
  slotted.app.unmount();
});

test("PptxViewer rebuilds playback when showHiddenSlides changes", async () => {
  const playback = createPlaybackSession();
  const source = new Uint8Array([5]);
  const showHiddenSlides = vue.ref(false);
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(PptxViewer, {
        source,
        mode: "present",
        autoplay: false,
        showHiddenSlides: showHiddenSlides.value,
        documentSessionFactory: () => playback.session,
        showSearch: false,
        showSidebar: false,
      });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => playback.calls.createPlaybackControllerOptions.length >= 1,
    "playback controller",
  );
  assert.equal(playback.calls.createPlaybackControllerOptions[0].skipHiddenSlides, true);

  showHiddenSlides.value = true;
  await waitFor(
    () => playback.calls.createPlaybackControllerOptions.length >= 2,
    "playback controller rebuilt for hidden slides",
  );
  assert.equal(playback.calls.createPlaybackControllerOptions.at(-1).skipHiddenSlides, false);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("pptx-viewer stylesheet keeps the fullscreen and letterbox rules", async () => {
  const { readFileSync } = await import("node:fs");
  const css = readFileSync(
    new URL("../../packages/vue-pptx/src/style.css", import.meta.url),
    "utf8",
  );
  assert.match(css, /\.pptx-viewer:fullscreen\s*\{/);
  assert.match(css, /aspect-ratio:\s*var\(--pptx-slide-aspect/);
  assert.match(css, /--pptx-slide-ratio/);
});

test("PptxViewer windows the sidebar and shares one thumbnail observer", async (t) => {
  const observers = installIntersectionObserver(t);
  const { calls, session } = createSession({ document: documentWithSlides(240) });
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([1]),
    sessionFactory: () => session,
    showSearch: false,
    locale: "zh-CN",
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX ready state",
  );
  // 960x540 slide at the default 148px thumbnail width => 132px rows.
  const sidebar = findByTestId(mounted.root, "pptx-sidebar");
  assert.equal(sidebar.props.style["--pptx-thumbnail-item-height"], "132px");

  const initial = thumbnails(mounted.root);
  assert.equal(initial.length, 12, "only the visible window plus overscan may mount");
  assert.equal(initial[0].props["data-pptx-thumbnail-index"], 0);
  assert.equal(initial.at(-1).props["data-pptx-thumbnail-index"], 11);
  const initialSpacers = nodesWithClass(mounted.root, "pptx-viewer__sidebar-spacer");
  assert.equal(initialSpacers.length, 1);
  assert.equal(initialSpacers[0].props.style.height, `${(240 - 12) * 132}px`);

  assert.equal(observers.length, 1, "the whole sidebar must share a single observer");
  assert.equal(observers[0].options.rootMargin, "240px");
  assert.equal(observers[0].targets.size, 12);
  assert.deepEqual(calls.renderThumbnail, [], "thumbnails wait until they approach the viewport");

  const observed = [...observers[0].targets].find((node) => node.props["data-pptx-thumbnail-index"] === 2);
  assert.ok(observed, "the third thumbnail must be observed");
  observers[0].intersect(observed);
  await vue.nextTick();
  assert.deepEqual(calls.renderThumbnail.map((call) => call.index), [2]);
  assert.equal(calls.renderThumbnail[0].width, 148);

  sidebar.props.onScrollPassive({ target: { scrollTop: 30 * 132, clientHeight: 660 } });
  await vue.nextTick();
  const scrolled = thumbnails(mounted.root);
  assert.equal(scrolled.length, 12);
  assert.equal(scrolled[0].props["data-pptx-thumbnail-index"], 27);
  assert.equal(scrolled.at(-1).props["data-pptx-thumbnail-index"], 38);
  const scrolledSpacers = nodesWithClass(mounted.root, "pptx-viewer__sidebar-spacer");
  assert.equal(scrolledSpacers.length, 2);
  assert.equal(scrolledSpacers[0].props.style.height, `${27 * 132}px`);
  assert.equal(observers[0].targets.size, 12, "scrolled-out thumbnails must unobserve");

  mounted.app.unmount();
  assert.equal(observers[0].disconnected, true);
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer re-measures the stage and thumbnails once per loaded font batch", async (t) => {
  const fontSet = installFontFaceSet(t);
  const { calls, session } = createSession({ document: documentWithSlides(3), withRefresh: true });
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([1]),
    sessionFactory: () => session,
    showSearch: false,
    locale: "zh-CN",
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX ready state",
  );
  assert.deepEqual(calls.renderThumbnail.map((call) => call.index), [0, 1, 2]);
  assert.equal(calls.refresh, 0);

  fontSet.emitLoadingDone();
  fontSet.emitLoadingDone();
  fontSet.emitLoadingDone();
  await waitFor(() => calls.refresh === 1, "PPTX refresh after fonts load");
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(calls.refresh, 1, "one font batch must coalesce into one refresh");

  await waitFor(() => calls.renderThumbnail.length === 6, "thumbnail repaint after fonts load");
  assert.deepEqual(
    calls.renderThumbnail.slice(3).map((call) => call.index).sort((left, right) => left - right),
    [0, 1, 2],
    "thumbnails must render again with the loaded font metrics",
  );

  fontSet.resolveReady();
  await new Promise((resolve) => setTimeout(resolve, 5));
  assert.equal(calls.refresh, 1, "fonts.ready must not repeat a batch already reported by loadingdone");
  assert.equal(calls.renderThumbnail.length, 6);

  mounted.app.unmount();
  assert.equal(fontSet.listenerCount(), 0);
  const refreshed = calls.refresh;
  fontSet.emitLoadingDone();
  await new Promise((resolve) => setTimeout(resolve, 5));
  assert.equal(calls.refresh, refreshed, "the font listener must be removed on unmount");
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer refreshes once when webfonts settle through fonts.ready", async (t) => {
  const fontSet = installFontFaceSet(t);
  const { calls, session } = createSession({ document: documentWithSlides(2), withRefresh: true });
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([1]),
    sessionFactory: () => session,
    showSearch: false,
    locale: "zh-CN",
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX ready state",
  );
  assert.equal(calls.refresh, 0);
  assert.deepEqual(calls.renderThumbnail.map((call) => call.index), [0, 1]);

  fontSet.resolveReady();
  await waitFor(() => calls.refresh === 1, "fonts.ready refresh");
  await waitFor(() => calls.renderThumbnail.length === 4, "thumbnail repaint after fonts.ready");
  await new Promise((resolve) => setTimeout(resolve, 5));
  assert.equal(calls.refresh, 1);
  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer tolerates sessions without a refresh hook", async (t) => {
  const fontSet = installFontFaceSet(t);
  const { session } = createSession({ document: documentWithSlides(2) });
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([1]),
    sessionFactory: () => session,
    showSearch: false,
    locale: "zh-CN",
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX ready state",
  );

  fontSet.emitLoadingDone();
  fontSet.resolveReady();
  await new Promise((resolve) => setTimeout(resolve, 5));
  assert.equal(findByTestId(mounted.root, "pptx-viewer").props["data-state"], "ready");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("PPTX viewer styles keep a CJK-aware font stack", async () => {
  const style = await readFile(
    new URL("../../packages/vue-pptx/src/style.css", import.meta.url),
    "utf8",
  );
  const stack = /--pptx-font-family:\s*([^;]+);/u.exec(style)?.[1] ?? "";
  for (const family of [
    "PingFang SC",
    "Hiragino Sans GB",
    "Microsoft YaHei",
    "Noto Sans CJK SC",
    "Source Han Sans SC",
  ]) {
    assert.ok(stack.includes(`"${family}"`), `font stack must include ${family}`);
  }
  assert.match(style, /font-family:\s*var\(--pptx-font-family\)/u);
});

test("PptxViewer browse keys scroll the slide list instead of changing slides", async () => {
  const { calls, session } = createSession({ document: documentWithSlides(3) });
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([1, 2, 3]),
    sessionFactory: () => session,
    showSidebar: false,
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX ready state",
  );

  const viewer = findByTestId(mounted.root, "pptx-viewer");
  const stageWrap = nodesWithClass(mounted.root, "pptx-viewer__stage-wrap")[0];
  assert.ok(stageWrap, "browse mode owns a scroll container");
  stageWrap.clientHeight = 400;
  stageWrap.scrollHeight = 1600;
  stageWrap.scrollTop = 0;
  const rendersBefore = calls.renderSlide.length;

  const pageDown = pressKey(viewer, "PageDown");
  await vue.nextTick();
  assert.equal(stageWrap.scrollTop, 360, "PageDown scrolls one viewport of the list");
  assert.equal(pageDown.prevented, 1, "the handled scroll key is prevented");
  assert.equal(calls.renderSlide.length, rendersBefore, "browse PageDown must not navigate slides");
  assert.equal(viewer.props["data-slide-index"], 0);

  pressKey(viewer, "ArrowDown");
  await vue.nextTick();
  assert.equal(stageWrap.scrollTop, 408, "ArrowDown scrolls one line");

  pressKey(viewer, "End");
  await vue.nextTick();
  assert.equal(stageWrap.scrollTop, 1200, "End scrolls to the bottom of the list");

  pressKey(viewer, "Home");
  await vue.nextTick();
  assert.equal(stageWrap.scrollTop, 0, "Home scrolls to the top of the list");

  pressKey(viewer, "PageDown", { target: { tagName: "INPUT" } });
  await vue.nextTick();
  assert.equal(stageWrap.scrollTop, 0, "browse keys do not fire inside an input");

  pressKey(viewer, "=", { ctrlKey: true });
  await waitFor(() => calls.setZoom.includes(125), "PPTX browse Ctrl+= zoom");
  pressKey(viewer, "-", { ctrlKey: true });
  await waitFor(() => calls.setZoom.includes(100), "PPTX browse Ctrl+- zoom");
  pressKey(viewer, "=", { ctrlKey: true });
  await waitFor(() => calls.setZoom.filter((value) => value === 125).length === 2, "PPTX browse zoom in again");
  pressKey(viewer, "0", { ctrlKey: true });
  await waitFor(
    () => calls.setZoom.filter((value) => value === 100).length === 2,
    "PPTX browse Ctrl+0 zoom reset",
  );
  assert.equal(calls.renderSlide.length, rendersBefore, "zoom does not switch slides");

  const search = walk(mounted.root).find(
    (node) => node.type === "input" && node.props?.type === "search",
  );
  assert.ok(search, "search input is available in browse mode");
  let focused = 0;
  search.focus = () => { focused += 1; };
  pressKey(viewer, "f", { ctrlKey: true });
  assert.equal(focused, 1, "Ctrl+F focuses the PPTX search field");

  search.value = "deck";
  search.dispatchEvent({ type: "input", target: search });
  await vue.nextTick();
  pressKey(viewer, "Escape");
  await vue.nextTick();
  assert.equal(
    walk(mounted.root).find((node) => node.type === "input" && node.props?.type === "search")?.value,
    "",
    "Escape cancels the PPTX search",
  );

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("PptxViewer present keys navigate slides and ignore editable/interactive targets", async () => {
  const playback = createPlaybackSession();
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([7, 8, 9]),
    mode: "present",
    autoplay: false,
    showSidebar: false,
    documentSessionFactory: () => playback.session,
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX playback ready state",
  );
  await waitFor(
    () => playback.calls.createPlaybackController === 1,
    "PPTX playback controller",
  );
  const viewer = findByTestId(mounted.root, "pptx-viewer");

  pressKey(viewer, "PageDown");
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(playback.calls.next, 1, "present PageDown advances one slide");

  pressKey(viewer, "ArrowRight");
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(playback.calls.next, 2, "present ArrowRight advances one slide");

  pressKey(viewer, "PageUp");
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(playback.calls.previous, 1, "present PageUp goes back one slide");

  pressKey(viewer, "Home");
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.deepEqual(playback.calls.goToSlide.at(-1), [0, { includeHidden: false }]);

  pressKey(viewer, "End");
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.deepEqual(playback.calls.goToSlide.at(-1), [2, { includeHidden: false }]);

  const nextBeforeEditable = playback.calls.next;
  pressKey(viewer, "ArrowDown", { target: { tagName: "INPUT" } });
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(playback.calls.next, nextBeforeEditable, "present keys stay out of inputs");

  const nextBeforeButton = playback.calls.next;
  pressKey(viewer, " ", {
    target: { tagName: "BUTTON", closest: (selector) => (selector.includes("button") ? {} : null) },
  });
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(playback.calls.next, nextBeforeButton, "Space on a focused button activates the button");

  const previousDocument = globalThis.document;
  let exitedFullscreen = 0;
  globalThis.document = {
    fullscreenElement: {},
    exitFullscreen: async () => { exitedFullscreen += 1; },
  };
  try {
    pressKey(viewer, "Escape");
    await new Promise((resolve) => setTimeout(resolve, 0));
    assert.equal(exitedFullscreen, 1, "present Escape leaves fullscreen");
  } finally {
    if (previousDocument === undefined) delete globalThis.document;
    else globalThis.document = previousDocument;
  }

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});
