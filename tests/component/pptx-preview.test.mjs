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

function documentWithSlides(count, fileName = "preview.pptx") {
  return {
    fileName,
    height: 540,
    width: 960,
    slides: Array.from({ length: count }, (_, index) => ({
      hidden: index === 1,
      index,
      number: index + 1,
    })),
    warnings: [],
  };
}

function createSession({ document = documentWithSlides(3), open } = {}) {
  const calls = {
    clearSearchHighlights: 0,
    dispose: 0,
    renderSlide: [],
    setZoom: [],
  };
  let currentSlideIndex = 0;
  let zoomPercent = 100;
  const session = {
    document: null,
    get currentSlideIndex() { return currentSlideIndex; },
    get zoomPercent() { return zoomPercent; },
    async open(source, options) {
      const loaded = open ? await open(source, options) : document;
      session.document = loaded;
      currentSlideIndex = options?.initialSlide ?? 0;
      return loaded;
    },
    async renderSlide(index) {
      calls.renderSlide.push(index);
      currentSlideIndex = index;
    },
    renderThumbnail() { return null; },
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

function createPlaybackSession({ document = documentWithSlides(3) } = {}) {
  const preview = createSession({ document });
  const listeners = new Set();
  const calls = {
    ...preview.calls,
    activateObject: [],
    createPlaybackController: 0,
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
    dispose() {},
  };
  const session = {
    ...preview.session,
    playbackDocument: { capability, slides: [] },
    capabilityReport: capability,
    createPlaybackController() {
      calls.createPlaybackController += 1;
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

test("PptxViewer exposes loading, ready navigation, hidden-page marker and cleanup", async () => {
  const pending = deferred();
  const { calls, session } = createSession({ open: () => pending.promise });
  const events = [];
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([1, 2, 3]),
    sessionFactory: () => session,
    showSearch: false,
    showSidebar: false,
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

  button(mounted.root, "下一页").props.onClick();
  await waitFor(() => calls.renderSlide.includes(1), "PPTX next slide");
  assert.match(textContent(mounted.root), /隐藏页/);
  assert.match(textContent(findByTestId(mounted.root, "pptx-page-counter")), /2 \/ 3/);

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
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([7, 8, 9]),
    mode: "present",
    autoplay: false,
    documentSessionFactory: () => playback.session,
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
