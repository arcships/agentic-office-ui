import assert from "node:assert/strict";
import test from "node:test";

import { importFromDemo, keydownEvent, vue } from "../component/vue-test-renderer.mjs";

const {
  VIEWER_KEYBOARD_SHORTCUTS,
  isEditableViewerTarget,
  isInteractiveViewerTarget,
  resolveViewerKeyCommand,
  useViewerKeys,
} = await importFromDemo("@arcships/vue-ui");

function keyEvent(key, init = {}) {
  const { event, preventedCount } = keydownEvent(key, init);
  Object.defineProperty(event, "prevented", { configurable: true, get: preventedCount });
  return event;
}

function resolve(key, options = {}, init = {}) {
  return resolveViewerKeyCommand(keyEvent(key, init), options);
}

test("viewer key map documents the page, zoom, search and escape contract", () => {
  const keys = VIEWER_KEYBOARD_SHORTCUTS.map((entry) => entry.keys);
  for (const expected of [
    "PageDown",
    "PageUp",
    "Home",
    "End",
    "Ctrl/Cmd + + or =",
    "Ctrl/Cmd + -",
    "Ctrl/Cmd + 0",
    "Ctrl/Cmd + F",
    "Escape",
  ]) {
    assert.ok(keys.includes(expected), `shared map documents ${expected}`);
  }
  assert.equal(
    VIEWER_KEYBOARD_SHORTCUTS.filter((entry) => entry.scope === "present").length > 0,
    true,
    "present-only arrow/space overrides stay documented",
  );
});

test("page keys resolve only when page navigation is enabled", () => {
  assert.equal(resolve("PageDown"), null);
  assert.equal(resolve("PageDown", { pageNavigation: true }), "page-next");
  assert.equal(resolve("PageUp", { pageNavigation: true }), "page-previous");
  assert.equal(resolve("Home", { pageNavigation: true }), "page-first");
  assert.equal(resolve("End", { pageNavigation: true }), "page-last");
  assert.equal(resolve("ArrowDown", { pageNavigation: true }), null);
});

test("present mode claims arrows and space, browse mode leaves them alone", () => {
  for (const key of ["ArrowDown", "ArrowRight", " "]) {
    assert.equal(resolve(key, { mode: "browse" }), null);
    assert.equal(resolve(key, { mode: "present", pageNavigation: true }), "page-next");
  }
  for (const key of ["ArrowUp", "ArrowLeft"]) {
    assert.equal(resolve(key, { mode: "browse" }), null);
    assert.equal(resolve(key, { mode: "present", pageNavigation: true }), "page-previous");
  }
  // Page keys keep their page meaning in presentation mode as well.
  assert.equal(resolve("PageDown", { mode: "present", pageNavigation: true }), "page-next");
  // A focused playback button owns Space; the viewer must not advance a slide.
  assert.equal(
    resolve(" ", { mode: "present", pageNavigation: true }, {
      target: { tagName: "BUTTON", closest: (selector) => ({ selector }) },
    }),
    null,
  );
  assert.equal(
    resolve("ArrowDown", { mode: "present", pageNavigation: true }, {
      target: { tagName: "BUTTON", closest: (selector) => ({ selector }) },
    }),
    "page-next",
    "arrow keys still navigate when a present-mode button has focus",
  );
});

test("zoom resolves only with Ctrl/Cmd and the matching keys", () => {
  assert.equal(resolve("+", { zoom: true }), null, "plain + is not a viewer shortcut");
  assert.equal(resolve("+", { zoom: true }, { ctrlKey: true }), "zoom-in");
  assert.equal(resolve("=", { zoom: true }, { metaKey: true }), "zoom-in");
  assert.equal(resolve("-", { zoom: true }, { ctrlKey: true }), "zoom-out");
  assert.equal(resolve("0", { zoom: true }, { ctrlKey: true }), "zoom-reset");
  assert.equal(resolve("f", { zoom: true }, { ctrlKey: true }), null);
});

test("search and escape resolve as separate commands", () => {
  assert.equal(resolve("f", { search: true }, { ctrlKey: true }), "search");
  assert.equal(resolve("F", { search: true }, { metaKey: true }), "search");
  assert.equal(resolve("Escape", { escape: true }), "escape");
  assert.equal(resolve("Escape", {}), null);
  assert.equal(resolve("f", { search: true }, { ctrlKey: true, defaultPrevented: true }), null);
});

test("shift/alt navigation and editable targets keep their native meaning", () => {
  assert.equal(resolve("PageDown", { pageNavigation: true }, { shiftKey: true }), null);
  assert.equal(resolve("End", { pageNavigation: true }, { altKey: true }), null);
  for (const target of [
    { tagName: "INPUT" },
    { tagName: "TEXTAREA" },
    { tagName: "SELECT" },
    { isContentEditable: true },
    { tagName: "DIV", closest: (selector) => (selector.includes("contenteditable") ? {} : null) },
  ]) {
    assert.equal(isEditableViewerTarget(target), true, `${JSON.stringify(target)} is editable`);
  }
  assert.equal(isEditableViewerTarget({ tagName: "DIV" }), false);
  assert.equal(isEditableViewerTarget(null), false);
  assert.equal(
    isInteractiveViewerTarget({ tagName: "BUTTON", closest: () => ({}) }),
    true,
  );
  assert.equal(isInteractiveViewerTarget({ tagName: "DIV" }), false);
});

test("useViewerKeys handles a command once, prevents default and respects editable targets", () => {
  const events = [];
  const keyboard = useViewerKeys({
    pageNavigation: true,
    zoom: true,
    search: () => false,
    escape: true,
    onCommand: (command) => {
      events.push(command);
      return command !== "page-first";
    },
  });

  const pageDown = keyEvent("PageDown");
  assert.equal(keyboard.onKeydown(pageDown), true);
  assert.equal(pageDown.prevented, 1);
  assert.deepEqual(events, ["page-next"]);

  const home = keyEvent("Home");
  assert.equal(keyboard.onKeydown(home), false, "an unhandled boundary keeps the browser key");
  assert.equal(home.prevented, 0);

  const editable = keyEvent("PageDown", { target: { tagName: "INPUT" } });
  assert.equal(keyboard.onKeydown(editable), false);
  assert.equal(editable.prevented, 0);

  const escape = keyEvent("Escape");
  assert.equal(keyboard.onKeydown(escape), true);
  assert.deepEqual(events, ["page-next", "page-first", "escape"]);
});

test("useViewerKeys reads reactive switches and mode", async () => {
  const searchEnabled = vue.ref(false);
  const mode = vue.ref("browse");
  const seen = [];
  const keyboard = useViewerKeys({
    search: searchEnabled,
    mode,
    pageNavigation: () => mode.value === "present",
    onCommand: (command) => {
      seen.push(command);
      return true
    },
  });

  assert.equal(keyboard.onKeydown(keyEvent("f", { ctrlKey: true })), false);
  searchEnabled.value = true;
  assert.equal(keyboard.onKeydown(keyEvent("f", { ctrlKey: true })), true);

  assert.equal(keyboard.onKeydown(keyEvent("ArrowDown")), false);
  mode.value = "present";
  assert.equal(keyboard.onKeydown(keyEvent("ArrowDown")), true);
  await vue.nextTick();
  assert.deepEqual(seen, ["search", "page-next"]);
});
