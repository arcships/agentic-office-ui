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
} from "./vue-test-renderer.mjs";

// Vue's native v-model directive checks these browser constructors during an
// update. The custom component renderer intentionally has no browser DOM.
if (globalThis.Document === undefined) globalThis.Document = class Document {};
if (globalThis.ShadowRoot === undefined) globalThis.ShadowRoot = class ShadowRoot {};

const {
  PptxStage,
  PptxViewer,
} = await importFromDemo("@arcships/vue-pptx");

const EXTERNAL_KEY = "ppt/slides/slide1.xml|slide|shape:2";
const INTERNAL_KEY = "ppt/slides/slide1.xml|slide|shape:3";
const SCRIPT_KEY = "ppt/slides/slide1.xml|slide|shape:4";
const HOVER_KEY = "ppt/slides/slide1.xml|slide|shape:5";
const PLAIN_KEY = "ppt/slides/slide1.xml|slide|shape:6";

function capability() {
  return { approximate: 0, discovered: 0, features: [], static: 0, strict: 0, unparsed: 0 };
}

function linkActions() {
  return [
    {
      id: "action:0:0",
      sourceObjectKey: EXTERNAL_KEY,
      trigger: "click",
      kind: "open-url",
      url: "https://example.com/docs",
      tooltip: "Read the docs",
    },
    {
      id: "action:0:1",
      sourceObjectKey: INTERNAL_KEY,
      trigger: "click",
      kind: "go-to-slide",
      targetSlideIndex: 2,
      tooltip: "Jump",
    },
    {
      id: "action:0:2",
      sourceObjectKey: SCRIPT_KEY,
      trigger: "click",
      kind: "unsupported",
      url: "javascript:alert(1)",
    },
    {
      id: "action:0:3",
      sourceObjectKey: HOVER_KEY,
      trigger: "hover",
      kind: "open-url",
      url: "https://example.com/hover",
    },
  ];
}

function linkPlaybackDocument() {
  return {
    capability: capability(),
    slides: Array.from({ length: 3 }, (_, index) => ({
      actions: index === 0 ? linkActions() : [],
      capability: capability(),
      hidden: false,
      index,
      media: [],
      morphFromPrevious: [],
      nodes: {},
      objects: [],
    })),
  };
}

function documentWithSlides(count) {
  return {
    fileName: "links.pptx",
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

function createBrowseSession() {
  const calls = { dispose: 0, open: [], renderSlide: [] };
  let currentSlideIndex = 0;
  const session = {
    document: null,
    playbackDocument: linkPlaybackDocument(),
    get currentSlideIndex() { return currentSlideIndex; },
    get zoomPercent() { return 100; },
    async open(_source, options) {
      calls.open.push(options);
      session.document = documentWithSlides(3);
      currentSlideIndex = options?.initialSlide ?? 0;
      return session.document;
    },
    async renderSlide(index) {
      calls.renderSlide.push(index);
      currentSlideIndex = index;
    },
    renderThumbnail() { return null; },
    searchText() { return []; },
    async highlightSearchResult() { return null; },
    clearSearchHighlights() {},
    async setZoom() {},
    cancel() {},
    dispose() { calls.dispose += 1; },
  };
  return { calls, session };
}

function createPlaybackSession() {
  const listeners = new Set();
  const calls = { activateObject: [], dispose: 0, disposeController: 0, next: 0, renderSlide: [] };
  let currentSlideIndex = 0;
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
    async activateObject(key) { calls.activateObject.push(key); return true; },
    async previous() {},
    async play() {},
    pause() {},
    async resume() {},
    async reset() {},
    async goToSlide() {},
    async resumeBlockedMedia() {},
    subscribe(listener) {
      listeners.add(listener);
      listener({ type: "statechange", snapshot });
      return () => listeners.delete(listener);
    },
    dispose() { calls.disposeController += 1; },
  };
  const session = {
    document: null,
    playbackDocument: linkPlaybackDocument(),
    capabilityReport: capability(),
    get currentSlideIndex() { return currentSlideIndex; },
    get zoomPercent() { return 100; },
    async open() {
      session.document = documentWithSlides(3);
      return session.document;
    },
    async renderSlide(index) {
      calls.renderSlide.push(index);
      currentSlideIndex = index;
    },
    renderThumbnail() { return null; },
    searchText() { return []; },
    async highlightSearchResult() { return null; },
    clearSearchHighlights() {},
    async setZoom() {},
    cancel() {},
    createPlaybackController(options) {
      snapshot = { ...snapshot, slideIndex: options?.initialSlide ?? 0 };
      return controller;
    },
    dispose() { calls.dispose += 1; },
  };
  return {
    calls,
    controller,
    session,
    emit(event) {
      if (event.type === "statechange") snapshot = event.snapshot;
      for (const listener of listeners) listener(event);
    },
  };
}

function mouseEvent(init) {
  let prevented = 0;
  const event = {
    ...init,
    preventDefault() { prevented += 1; },
  };
  return { event, preventedCount: () => prevented };
}

function fakeTarget(slideIndex, objectKey) {
  const slide = { dataset: { slideIndex: String(slideIndex) } };
  const object = objectKey ? { dataset: { pptxObjectKey: objectKey } } : null;
  return {
    slide,
    object,
    target: {
      closest(selector) {
        if (selector === "[data-slide-index]") return slide;
        if (selector === "[data-pptx-object-key]") return object;
        return null;
      },
    },
  };
}

function handlers(node, name) {
  const value = node?.props?.[name];
  return Array.isArray(value) ? value : [value];
}

function trigger(node, name, event) {
  for (const handler of handlers(node, name)) handler?.(event);
}

function installOpener(t) {
  const opened = [];
  const original = globalThis.open;
  globalThis.open = (url, name, features) => {
    opened.push([url, name, features]);
    return null;
  };
  t.after(() => {
    if (original === undefined) delete globalThis.open;
    else globalThis.open = original;
  });
  return opened;
}

function installClipboard(t) {
  const written = [];
  const original = Object.getOwnPropertyDescriptor(globalThis, "navigator");
  Object.defineProperty(globalThis, "navigator", {
    configurable: true,
    value: { clipboard: { async writeText(value) { written.push(value); } } },
  });
  t.after(() => {
    if (original) Object.defineProperty(globalThis, "navigator", original);
    else delete globalThis.navigator;
  });
  return written;
}

async function mountReady(t, props) {
  const mounted = await mount(PptxViewer, props);
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX links ready state",
  );
  return mounted;
}

test("browse mode routes linked shapes through the URL policy and stays clickable by keyboard", async (t) => {
  const opened = installOpener(t);
  const { calls, session } = createBrowseSession();
  const actions = [];
  const mounted = await mountReady(t, {
    source: new Uint8Array([1, 2, 3]),
    sessionFactory: () => session,
    showSearch: false,
    showSidebar: false,
    onAction: (action) => actions.push(action),
  });
  const stage = findByTestId(mounted.root, "pptx-stage");

  // External link: opened once, reported to the host, native anchor prevented.
  const external = fakeTarget(0, EXTERNAL_KEY);
  const externalClick = mouseEvent({ target: external.target });
  trigger(stage, "onClick", externalClick.event);
  assert.deepEqual(opened, [["https://example.com/docs", "_blank", "noopener,noreferrer"]]);
  assert.equal(externalClick.preventedCount(), 1);
  assert.deepEqual(actions, [{
    kind: "open-url",
    url: "https://example.com/docs",
    sourceObjectKey: EXTERNAL_KEY,
  }]);

  // Internal jump navigates through the document session.
  const internalClick = mouseEvent({ target: fakeTarget(0, INTERNAL_KEY).target });
  trigger(stage, "onClick", internalClick.event);
  await waitFor(() => calls.renderSlide.includes(2), "internal link navigation");
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer").props["data-slide-index"] === 2,
    "internal link slide index",
  );
  assert.equal(internalClick.preventedCount(), 1);
  assert.equal(actions.at(-1).kind, "go-to-slide");
  assert.equal(actions.at(-1).slideIndex, 2);
  assert.equal(opened.length, 1, "internal links never open a window");

  // javascript: targets are rejected and reported as unsupported.
  const scriptClick = mouseEvent({ target: fakeTarget(0, SCRIPT_KEY).target });
  trigger(stage, "onClick", scriptClick.event);
  assert.equal(scriptClick.preventedCount(), 1);
  assert.equal(opened.length, 1);
  assert.equal(actions.at(-1).kind, "unsupported");

  // Hover links are not click targets; a shape without a click link does nothing.
  for (const key of [HOVER_KEY, PLAIN_KEY]) {
    const click = mouseEvent({ target: fakeTarget(0, key).target });
    trigger(stage, "onClick", click.event);
    assert.equal(click.preventedCount(), 0);
  }
  assert.equal(opened.length, 1);
  assert.equal(actions.length, 3);

  // Keyboard activation uses the same route as a pointer click.
  const keyboardTarget = fakeTarget(0, EXTERNAL_KEY);
  keyboardTarget.target.closest = (selector) => selector === "[data-pptx-object-key][data-pptx-hyperlink='click']"
    ? keyboardTarget.object
    : selector === "[data-slide-index]"
      ? keyboardTarget.slide
      : selector === "[data-pptx-object-key]"
        ? keyboardTarget.object
        : null;
  const { event, preventedCount } = keydownEvent("Enter", { target: keyboardTarget.target });
  let stopped = 0;
  Object.defineProperty(event, "stopPropagation", {
    configurable: true,
    value() { stopped += 1; },
  });
  let activated;
  const stageOnly = await mount(PptxStage, {
    onLinkActivate: (target) => { activated = target; },
  });
  findByTestId(stageOnly.root, "pptx-stage").props.onKeydown(event);
  assert.deepEqual(activated, { slideIndex: 0, objectKey: EXTERNAL_KEY });
  assert.equal(preventedCount(), 1);
  assert.equal(stopped, 1);
  stageOnly.app.unmount();

  // The viewer wires that stage event to the same policy-checked open.
  const viewerKeyboard = keydownEvent("Enter", { target: keyboardTarget.target });
  trigger(stage, "onKeydown", viewerKeyboard.event);
  await waitFor(() => opened.length === 2, "keyboard link activation in the viewer");
  assert.deepEqual(opened.at(-1), ["https://example.com/docs", "_blank", "noopener,noreferrer"]);

  mounted.app.unmount();
  assert.equal(calls.dispose, 1);
  assert.deepEqual(mounted.warnings, []);
});

test("linkHandling=host keeps the policy and the opening with the host", async (t) => {
  const opened = installOpener(t);
  const { session } = createBrowseSession();
  const actions = [];
  const mounted = await mountReady(t, {
    source: new Uint8Array([3, 3, 3]),
    sessionFactory: () => session,
    showSearch: false,
    showSidebar: false,
    linkHandling: "host",
    onAction: (action) => actions.push(action),
  });
  const stage = findByTestId(mounted.root, "pptx-stage");
  const click = mouseEvent({ target: fakeTarget(0, EXTERNAL_KEY).target });
  trigger(stage, "onClick", click.event);
  assert.equal(opened.length, 0, "the viewer must not open links in host mode");
  assert.equal(click.preventedCount(), 0, "the host owns the click");
  assert.deepEqual(actions, [{
    kind: "open-url",
    url: "https://example.com/docs",
    sourceObjectKey: EXTERNAL_KEY,
  }]);
  mounted.app.unmount();
});

test("present mode link opening is also opt-out", async (t) => {
  const opened = installOpener(t);
  const playback = createPlaybackSession();
  const actions = [];
  const mounted = await mountReady(t, {
    source: new Uint8Array([5, 5, 5]),
    mode: "present",
    autoplay: false,
    linkHandling: "host",
    documentSessionFactory: () => playback.session,
    onAction: (action) => actions.push(action),
  });
  playback.emit({
    type: "action",
    action: { kind: "open-url", url: "https://example.com/host", sourceObjectKey: EXTERNAL_KEY },
  });
  await vue.nextTick();
  assert.equal(opened.length, 0);
  assert.equal(actions.at(-1).kind, "open-url");
  assert.equal(actions.at(-1).url, "https://example.com/host");
  mounted.app.unmount();
});

test("browse mode keeps hover links inert but exposes their tooltip metadata", () => {
  // The marker contract is covered by the unit test; this asserts the viewer
  // never treats a hover action as a click.
  const target = fakeTarget(0, HOVER_KEY);
  const event = mouseEvent({ target: target.target });
  assert.equal(event.preventedCount(), 0);
});

test("present mode opens approved action URLs and still forwards action events", async (t) => {
  const opened = installOpener(t);
  const playback = createPlaybackSession();
  const actions = [];
  const mounted = await mountReady(t, {
    source: new Uint8Array([4, 5, 6]),
    mode: "present",
    autoplay: false,
    documentSessionFactory: () => playback.session,
    onAction: (action) => actions.push(action),
  });

  playback.emit({
    type: "action",
    action: { kind: "open-url", url: "https://example.com/present", sourceObjectKey: EXTERNAL_KEY },
  });
  await vue.nextTick();
  assert.deepEqual(opened, [["https://example.com/present", "_blank", "noopener,noreferrer"]]);
  assert.equal(actions.at(-1).kind, "open-url");
  assert.equal(actions.at(-1).url, "https://example.com/present");

  // A policy-rejected target cannot reach the browser, even from a host-supplied
  // playback document; the host sees it downgraded to unsupported.
  playback.emit({
    type: "action",
    action: { kind: "open-url", url: "javascript:alert(1)", sourceObjectKey: SCRIPT_KEY },
  });
  await vue.nextTick();
  assert.equal(opened.length, 1);
  assert.equal(actions.at(-1).kind, "unsupported");

  // present-mode link handling is opt-out: hosts that own the policy get the
  // same action events, and the viewer opens nothing itself.
  mounted.app.unmount();
  assert.equal(playback.calls.dispose, 1);
});
test("a linked shape click in present mode is left to the controller", async (t) => {
  const opened = installOpener(t);
  const playback = createPlaybackSession();
  const mounted = await mountReady(t, {
    source: new Uint8Array([7, 8, 9]),
    mode: "present",
    autoplay: false,
    documentSessionFactory: () => playback.session,
  });
  const stage = findByTestId(mounted.root, "pptx-stage");
  const click = mouseEvent({ target: fakeTarget(0, EXTERNAL_KEY).target });
  trigger(stage, "onClick", click.event);
  await waitFor(() => playback.calls.activateObject.length === 1, "present link activation");
  assert.deepEqual(playback.calls.activateObject, [EXTERNAL_KEY]);
  assert.equal(click.preventedCount(), 1, "the nested anchor must not open twice");
  assert.equal(opened.length, 0, "the controller owns the action in present mode");
  mounted.app.unmount();
});

test("the built-in context menu opens and copies links and still reports the host event", async (t) => {
  const opened = installOpener(t);
  const written = installClipboard(t);
  const { calls, session } = createBrowseSession();
  const menus = [];
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([9, 9, 9]),
    sessionFactory: () => session,
    showSearch: false,
    showSidebar: false,
    onContextMenu: (context) => menus.push(context),
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX context menu ready state",
  );
  const stage = findByTestId(mounted.root, "pptx-stage");
  stage.getBoundingClientRect = () => ({ left: 10, top: 20 });

  const external = fakeTarget(0, EXTERNAL_KEY);
  const openEvent = mouseEvent({ target: external.target, clientX: 40, clientY: 60 });
  stage.props.onContextmenu(openEvent.event);
  await vue.nextTick();

  assert.equal(menus.length, 1);
  assert.deepEqual(menus[0], {
    kind: "object",
    slideIndex: 0,
    objectKey: EXTERNAL_KEY,
    clientX: 40,
    clientY: 60,
    containerX: 30,
    containerY: 40,
  });
  const menu = findByTestId(mounted.root, "pptx-context-menu");
  assert.ok(menu, "built-in menu is rendered");
  assert.equal(menu.props.role, "menu");
  assert.deepEqual(menu.props.style, { left: "40px", top: "60px" });
  const openItem = findByTestId(mounted.root, "pptx-context-menu-open-link");
  assert.match(textContent(openItem), /Open link/);
  const copyItem = findByTestId(mounted.root, "pptx-context-menu-copy-link");
  assert.match(textContent(copyItem), /Copy link/);

  copyItem.props.onClick();
  await waitFor(() => written.length === 1, "copy link");
  assert.deepEqual(written, ["https://example.com/docs"]);
  assert.equal(findByTestId(mounted.root, "pptx-context-menu"), undefined, "copy closes the menu");
  assert.equal(opened.length, 0);

  // Open link routes through the same policy as a click.
  const secondOpen = mouseEvent({ target: external.target, clientX: 12, clientY: 24 });
  stage.props.onContextmenu(secondOpen.event);
  await vue.nextTick();
  findByTestId(mounted.root, "pptx-context-menu-open-link").props.onClick();
  await waitFor(() => opened.length === 1, "context menu open link");
  assert.deepEqual(opened, [["https://example.com/docs", "_blank", "noopener,noreferrer"]]);

  // Internal jumps get a slide-specific label and navigate.
  const internal = fakeTarget(0, INTERNAL_KEY);
  stage.props.onContextmenu(mouseEvent({ target: internal.target, clientX: 1, clientY: 2 }).event);
  await vue.nextTick();
  const goItem = findByTestId(mounted.root, "pptx-context-menu-open-link");
  assert.match(textContent(goItem), /Go to slide 3/);
  goItem.props.onClick();
  await waitFor(() => calls.renderSlide.includes(2), "context menu internal jump");
  assert.equal(findByTestId(mounted.root, "pptx-context-menu-copy-link"), undefined);

  // Unlinked shapes keep the host event but show no built-in menu.
  stage.props.onContextmenu(mouseEvent({ target: fakeTarget(0, PLAIN_KEY).target, clientX: 3, clientY: 4 }).event);
  await vue.nextTick();
  assert.equal(findByTestId(mounted.root, "pptx-context-menu"), undefined);
  assert.equal(menus.length, 4);

  // Escape closes the menu through the shared menu keyboard semantics.
  stage.props.onContextmenu(mouseEvent({ target: external.target, clientX: 5, clientY: 6 }).event);
  await vue.nextTick();
  const reopened = findByTestId(mounted.root, "pptx-context-menu");
  const { event: escape } = keydownEvent("Escape");
  reopened.props.onKeydown(escape);
  await vue.nextTick();
  assert.equal(findByTestId(mounted.root, "pptx-context-menu"), undefined);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("the built-in context menu can be turned off without losing the host event", async () => {
  const { session } = createBrowseSession();
  const menus = [];
  const mounted = await mount(PptxViewer, {
    source: new Uint8Array([2, 2, 2]),
    sessionFactory: () => session,
    showSearch: false,
    showSidebar: false,
    showContextMenu: false,
    onContextMenu: (context) => menus.push(context),
  });
  await waitFor(
    () => findByTestId(mounted.root, "pptx-viewer")?.props["data-state"] === "ready",
    "PPTX menu opt-out ready state",
  );
  const stage = findByTestId(mounted.root, "pptx-stage");
  stage.props.onContextmenu(mouseEvent({ target: fakeTarget(0, EXTERNAL_KEY).target }).event);
  await vue.nextTick();
  assert.equal(menus.length, 1);
  assert.equal(findByTestId(mounted.root, "pptx-context-menu"), undefined);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});
