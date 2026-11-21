import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";

import {
  importFromDemo,
  keydownEvent,
  mount,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;
globalThis.Document ??= class Document {};
globalThis.ShadowRoot ??= class ShadowRoot {};

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { XlsxSheetTabs, XlsxViewer, useXlsxViewerController } = await importFromDemo(
  "@arcships/vue-xlsx",
);
const { officeMessagesEn } = await import(
  new URL("../../packages/office-interaction/src/i18n.ts", import.meta.url)
);

const messages = officeMessagesEn.xlsx;
const SHEET_TABS_SOURCE = readFileSync(
  new URL("../../packages/vue-xlsx/src/components/XlsxSheetTabs.vue", import.meta.url),
  "utf8",
);
const VIEWER_SOURCE = readFileSync(
  new URL("../../packages/vue-xlsx/src/components/XlsxViewer.vue", import.meta.url),
  "utf8",
);

function configureWasm() {
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}

function workbookBuffer(name) {
  const bytes = readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function tabButtons(root) {
  return walk(root).filter((node) => node.props?.["data-testid"] === "xlsx-sheet-tab");
}

function tabButton(root, sheetName) {
  return tabButtons(root).find((node) => node.props?.["data-sheet-name"] === sheetName);
}

function tabWrapper(root, sheetName) {
  const tab = tabButton(root, sheetName);
  assert.ok(tab, `${sheetName} tab is rendered`);
  return tab.parent;
}

function tabList(root) {
  return walk(root).find((node) => node.props?.role === "tablist");
}

function dispatchKeydown(node, event) {
  const handler = node.props.onKeydown;
  for (const entry of Array.isArray(handler) ? handler : [handler]) entry?.(event);
}

async function mountTabs(file, options = {}) {
  const { panelId, readOnly, restoredTabIndex, ...controllerOptions } = options;
  const buffer = workbookBuffer(file);
  let controller;
  const showTabs = vue.ref(restoredTabIndex === undefined);
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: buffer,
        fileName: file,
        useWorker: false,
        ...controllerOptions,
      });
      if (restoredTabIndex !== undefined) {
        vue.watch(
          () => controller.sourceState,
          (state) => {
            if (state !== "ready") return;
            controller.setActiveTabIndex(restoredTabIndex);
            showTabs.value = true;
          },
          { flush: "sync", immediate: true },
        );
      }
      return () =>
        controller.sourceState === "ready" && showTabs.value
          ? vue.h(XlsxSheetTabs, {
              controller,
              readOnly: readOnly === true,
              panelId,
            })
          : vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    `${file} tabs controller`,
  );
  assert.ifError(controller.error);
  await waitFor(
    () => tabButtons(mounted.root).length > 0,
    `${file} tab strip`,
  );
  return { ...mounted, controller };
}

function viewerRoot(root) {
  const viewer = walk(root).find((node) => node.props?.class === "xlsx-viewer");
  assert.ok(viewer, "the viewer root is rendered");
  return viewer;
}

/**
 * Minimal controller for the viewer-level keyboard contract: switching tabs
 * does not need a parsed workbook or a grid.
 */
function keyboardController(overrides = {}) {
  const calls = { redo: 0, setActiveTabIndex: [], undo: 0 };
  const controller = {
    activeSheet: null,
    activeTab: null,
    activeTabIndex: 0,
    calculationStatus: null,
    canLoadDeferred: false,
    deferredLoadFileSize: 0,
    displayFileName: "keyboard.xlsx",
    error: null,
    hiddenSheets: [],
    isEditingLocked: false,
    isLoading: false,
    isLoadDeferred: false,
    isPrintPreviewOpen: false,
    readOnly: false,
    redo() {
      calls.redo += 1;
    },
    setActiveTabIndex(index) {
      calls.setActiveTabIndex.push(index);
      this.activeTabIndex = index;
    },
    sheets: [],
    sourceError: null,
    tabs: [
      { id: "tab-one", kind: "sheet", name: "One", sheetIndex: 0, visibility: "visible" },
      { id: "tab-two", kind: "sheet", name: "Two", sheetIndex: 1, visibility: "visible" },
    ],
    undo() {
      calls.undo += 1;
    },
    zoomScale: 100,
    ...overrides,
  };
  return { calls, controller: vue.reactive(controller) };
}

async function mountKeyboardViewer(overrides = {}) {
  const { calls, controller } = keyboardController(overrides);
  const mounted = await mount(XlsxViewer, {
    controller,
    showDefaultToolbar: false,
    showFormulaBar: false,
    showImages: false,
    showRibbon: false,
  });
  return { ...mounted, calls, controller, viewer: viewerRoot(mounted.root) };
}

configureWasm();

// ── Active tab visibility and scrolling ────────────────────────────────────

test("the restored active tab is scrolled into view on mount", async () => {
  // The harness applies the restored active tab before it renders the strip,
  // matching a host that opens a workbook on a saved sheet.
  const { app, root, controller } = await mountTabs("financial-model.xlsx", { restoredTabIndex: 2 });
  try {
    assert.equal(controller.activeTabIndex, 2);
    const wrappers = ["Assumptions", "P&L", "Notes"].map((name) => tabWrapper(root, name));
    await waitFor(
      () => wrappers[2].scrollIntoViewCalls.length > 0,
      "the restored active tab is scrolled into view on mount",
    );
    assert.deepEqual(wrappers[2].scrollIntoViewCalls.at(-1), {
      block: "nearest",
      inline: "nearest",
    });
    assert.deepEqual(wrappers[0].scrollIntoViewCalls, [], "inactive tabs are not scrolled");
    assert.deepEqual(wrappers[1].scrollIntoViewCalls, [], "inactive tabs are not scrolled");
  } finally {
    app.unmount();
  }
});

test("switching tabs scrolls only the new active tab, inside the strip", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    const wrappers = ["Assumptions", "P&L", "Notes"].map((name) => tabWrapper(root, name));
    await waitFor(
      () => wrappers[0].scrollIntoViewCalls.length > 0,
      "the initial active tab is scrolled into view",
    );

    // Replace the recorder with per-tab spies so the switch is the only call.
    const revealed = new Map(wrappers.map((wrapper) => [wrapper, []]));
    for (const wrapper of wrappers) {
      wrapper.scrollIntoView = (options) => revealed.get(wrapper).push(options);
    }
    const pageScrolls = [];
    root.scrollIntoView = (options) => pageScrolls.push(options);

    controller.setActiveTabIndex(2);
    await waitFor(
      () => revealed.get(wrappers[2]).length > 0,
      "the newly active tab is scrolled into view",
    );

    assert.deepEqual(revealed.get(wrappers[2]).at(-1), {
      block: "nearest",
      inline: "nearest",
    });
    assert.deepEqual(revealed.get(wrappers[0]), [], "the old active tab is not scrolled");
    assert.deepEqual(revealed.get(wrappers[1]), [], "sibling tabs are not scrolled");
    assert.deepEqual(pageScrolls, [], "the page is never scrolled");

    assert.match(
      SHEET_TABS_SOURCE,
      /\.xlsx-sheettabs__list\s*\{[^}]*scroll-behavior:\s*smooth/,
      "the strip scrolls smoothly",
    );
  } finally {
    app.unmount();
  }
});

// ── Tablist semantics ───────────────────────────────────────────────────────

test("the strip is a labelled tablist with a roving tab stop and active state", async () => {
  const { app, root } = await mountTabs("financial-model.xlsx");
  try {
    const list = tabList(root);
    assert.ok(list, "the tab list has role=tablist");
    assert.equal(list.props["aria-label"], messages.sheetTabs.tabListLabel);
    assert.equal(list.props["aria-orientation"], "horizontal");

    const tabs = tabButtons(root);
    assert.deepEqual(tabs.map((tab) => tab.props["data-sheet-name"]), [
      "Assumptions",
      "P&L",
      "Notes",
    ]);
    for (const tab of tabs) {
      assert.equal(tab.props.role, "tab");
      assert.equal(tab.props["data-testid"], "xlsx-sheet-tab", "the testid contract is kept");
      assert.equal(typeof tab.props.id, "string");
    }
    assert.equal(new Set(tabs.map((tab) => tab.props.id)).size, tabs.length, "tab ids are unique");
    assert.deepEqual(tabs.map((tab) => tab.props["aria-selected"]), [true, false, false]);
    assert.deepEqual(tabs.map((tab) => tab.props.tabindex), [0, -1, -1], "one roving tab stop");
    assert.match(tabs[0].props.class, /xlsx-sheettabs__tab--active/);
    assert.doesNotMatch(tabs[1].props.class, /xlsx-sheettabs__tab--active/);
    assert.equal(
      tabs.every((tab) => tab.props["aria-controls"] === undefined),
      true,
      "a standalone strip has no panel to point at",
    );
  } finally {
    app.unmount();
  }
});

test("tabs point aria-controls at the worksheet panel id they control", async () => {
  const { app, root } = await mountTabs("financial-model.xlsx", { panelId: "xlsx-grid-panel" });
  try {
    assert.deepEqual(
      tabButtons(root).map((tab) => tab.props["aria-controls"]),
      ["xlsx-grid-panel", "xlsx-grid-panel", "xlsx-grid-panel"],
    );
    // The viewer owns the id on both sides, so the pointer cannot drift.
    assert.match(VIEWER_SOURCE, /:id="gridId"/, "the grid carries the panel id");
    assert.match(
      VIEWER_SOURCE,
      /:panel-id="controller\.activeSheet \? gridId : undefined"/,
      "the strip receives the same panel id",
    );
  } finally {
    app.unmount();
  }
});

test("ArrowLeft/ArrowRight/Home/End move focus without switching tabs", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    const focused = [];
    for (const tab of tabButtons(root)) {
      tab.focus = () => focused.push(tab.props["data-sheet-name"]);
    }
    const [first, second, third] = tabButtons(root);

    const right = keydownEvent("ArrowRight");
    dispatchKeydown(first, right.event);
    assert.equal(right.preventedCount(), 1, "the arrow key is claimed");
    assert.deepEqual(focused, ["P&L"]);
    assert.equal(controller.activeTabIndex, 0, "moving focus does not change the sheet");
    await vue.nextTick();
    assert.deepEqual(
      tabButtons(root).map((tab) => tab.props.tabindex),
      [-1, 0, -1],
      "the roving tab stop follows focus",
    );

    dispatchKeydown(second, keydownEvent("End").event);
    assert.deepEqual(focused, ["P&L", "Notes"]);
    dispatchKeydown(third, keydownEvent("Home").event);
    assert.deepEqual(focused, ["P&L", "Notes", "Assumptions"]);
    dispatchKeydown(first, keydownEvent("ArrowLeft").event);
    assert.deepEqual(focused, ["P&L", "Notes", "Assumptions", "Notes"], "arrows wrap");
    assert.equal(controller.activeTabIndex, 0, "focus movement never activates");
  } finally {
    app.unmount();
  }
});

test("Enter and Space activate the focused tab and focus the active tab", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    const focused = [];
    for (const tab of tabButtons(root)) {
      tab.focus = () => focused.push(tab.props["data-sheet-name"]);
    }

    const enter = keydownEvent("Enter");
    dispatchKeydown(tabButtons(root)[1], enter.event);
    await vue.nextTick();
    assert.equal(enter.preventedCount(), 1);
    assert.equal(controller.activeTabIndex, 1);
    assert.deepEqual(tabButtons(root).map((tab) => tab.props["aria-selected"]), [false, true, false]);
    assert.deepEqual(tabButtons(root).map((tab) => tab.props.tabindex), [-1, 0, -1]);
    assert.ok(focused.includes("P&L"), "the activated tab keeps focus");

    dispatchKeydown(tabButtons(root)[2], keydownEvent(" ").event);
    await vue.nextTick();
    assert.equal(controller.activeTabIndex, 2);
    assert.deepEqual(tabButtons(root).map((tab) => tab.props["aria-selected"]), [false, false, true]);
    assert.match(tabButtons(root)[2].props.class, /xlsx-sheettabs__tab--active/);

    dispatchKeydown(tabButtons(root)[0], keydownEvent("Space").event);
    await vue.nextTick();
    assert.equal(controller.activeTabIndex, 0);
  } finally {
    app.unmount();
  }
});

test("focus follows the active tab after a switch without stealing grid focus", async () => {
  const focusedStrip = await mountTabs("financial-model.xlsx");
  try {
    const focused = [];
    for (const tab of tabButtons(focusedStrip.root)) {
      tab.focus = () => focused.push(tab.props["data-sheet-name"]);
    }
    tabWrapper(focusedStrip.root, "Assumptions").props.onFocusin();
    focusedStrip.controller.setActiveTabIndex(2);
    await waitFor(() => focused.includes("Notes"), "focus follows the active tab");
    assert.equal(focused.at(-1), "Notes");
    await vue.nextTick();
    assert.deepEqual(
      tabButtons(focusedStrip.root).map((tab) => tab.props.tabindex),
      [-1, -1, 0],
      "the active tab owns the tab stop",
    );
  } finally {
    focusedStrip.app.unmount();
  }

  // A switch driven from the grid (or a host restore) must not pull focus out
  // of the sheet just because the strip re-rendered.
  const unfocusedStrip = await mountTabs("financial-model.xlsx");
  try {
    const focused = [];
    for (const tab of tabButtons(unfocusedStrip.root)) {
      tab.focus = () => focused.push(tab.props["data-sheet-name"]);
    }
    unfocusedStrip.controller.setActiveTabIndex(2);
    await waitFor(
      () => tabWrapper(unfocusedStrip.root, "Notes").scrollIntoViewCalls.length > 0,
      "the switch is applied",
    );
    assert.deepEqual(focused, [], "the strip does not steal focus from the grid");
    await vue.nextTick();
    assert.deepEqual(
      tabButtons(unfocusedStrip.root).map((tab) => tab.props.tabindex),
      [-1, -1, 0],
      "the active tab remains the tab stop",
    );
  } finally {
    unfocusedStrip.app.unmount();
  }
});

// ── Viewer keyboard shortcuts ──────────────────────────────────────────────

test("Ctrl/Cmd+PageUp/PageDown switch sheets and clamp at the ends", async () => {
  const mounted = await mountKeyboardViewer();
  try {
    const { calls, controller, viewer } = mounted;
    assert.equal(controller.activeTabIndex, 0);

    const down = keydownEvent("PageDown", { ctrlKey: true });
    viewer.props.onKeydown(down.event);
    assert.equal(down.preventedCount(), 1, "the browser must not also scroll the page");
    assert.deepEqual(calls.setActiveTabIndex, [1]);
    assert.equal(controller.activeTabIndex, 1);
    await vue.nextTick();

    const last = keydownEvent("PageDown", { ctrlKey: true });
    viewer.props.onKeydown(last.event);
    assert.equal(last.preventedCount(), 1, "the shortcut is claimed even at the last tab");
    assert.deepEqual(calls.setActiveTabIndex, [1], "the index is clamped at the end");

    const up = keydownEvent("PageUp", { metaKey: true });
    viewer.props.onKeydown(up.event);
    assert.equal(up.preventedCount(), 1);
    assert.deepEqual(calls.setActiveTabIndex, [1, 0]);

    const first = keydownEvent("PageUp", { ctrlKey: true });
    viewer.props.onKeydown(first.event);
    assert.equal(first.preventedCount(), 1, "the shortcut is claimed even at the first tab");
    assert.deepEqual(calls.setActiveTabIndex, [1, 0], "the index is clamped at the start");
  } finally {
    mounted.app.unmount();
  }
});

test("sheet switching stays available in read-only viewers", async () => {
  const mounted = await mountKeyboardViewer({ readOnly: true });
  try {
    const down = keydownEvent("PageDown", { ctrlKey: true });
    mounted.viewer.props.onKeydown(down.event);
    assert.equal(down.preventedCount(), 1);
    assert.equal(mounted.controller.activeTabIndex, 1);
    assert.deepEqual(mounted.calls.setActiveTabIndex, [1]);
  } finally {
    mounted.app.unmount();
  }
});

test("sheet shortcuts and undo stay out of an open cell editor", async () => {
  const mounted = await mountKeyboardViewer();
  try {
    const editor = {
      closest(selector) {
        return selector.includes("xlsx-cell-editor") ? this : null;
      },
      tagName: "INPUT",
    };

    const down = keydownEvent("PageDown", { ctrlKey: true, target: editor });
    mounted.viewer.props.onKeydown(down.event);
    assert.equal(down.preventedCount(), 0, "the editor keeps the key");
    assert.deepEqual(mounted.calls.setActiveTabIndex, []);

    const up = keydownEvent("PageUp", { metaKey: true, target: editor });
    mounted.viewer.props.onKeydown(up.event);
    assert.equal(up.preventedCount(), 0);
    assert.deepEqual(mounted.calls.setActiveTabIndex, []);

    const undo = keydownEvent("z", { ctrlKey: true, target: editor });
    mounted.viewer.props.onKeydown(undo.event);
    assert.equal(mounted.calls.undo, 0, "the editor owns undo while it is open");
    assert.equal(undo.preventedCount(), 0);

    const redo = keydownEvent("y", { ctrlKey: true, target: editor });
    mounted.viewer.props.onKeydown(redo.event);
    assert.equal(mounted.calls.redo, 0);
    assert.equal(redo.preventedCount(), 0);

    const plainUndo = keydownEvent("z", { ctrlKey: true });
    mounted.viewer.props.onKeydown(plainUndo.event);
    assert.equal(mounted.calls.undo, 1);
    assert.equal(plainUndo.preventedCount(), 1);
  } finally {
    mounted.app.unmount();
  }
});

// Regression for the #191 CI flake: the strip swaps its 40px viewport for 50px
// when the tabs overflow, and that measurement used to run once on mount. A web
// font swap changes the tab widths without resizing the scroll element, so the
// stale `overflowing` value survived and the same cell was captured at 684x162
// on one page load and 684x182 on the next. The strip now re-measures when the
// tab list resizes and once the fonts are ready.
test("the overflow viewport re-measures after the tab list or fonts settle", async () => {
  const originalResizeObserver = globalThis.ResizeObserver;
  const originalDocument = globalThis.document;
  const observers = [];
  let resolveFonts;
  const fontsReady = new Promise((resolve) => {
    resolveFonts = resolve;
  });

  class FakeResizeObserver {
    constructor(callback) {
      this.callback = callback;
      observers.push(callback);
    }
    observe() {}
    disconnect() {}
  }

  globalThis.ResizeObserver = FakeResizeObserver;
  globalThis.document = {
    ...(originalDocument || {}),
    fonts: { ready: fontsReady },
  };

  const { app, root } = await mountTabs("financial-model.xlsx");
  try {
    const scroll = walk(root).find((node) =>
      String(node.props?.class || "").includes("xlsx-sheettabs__scroll"),
    );
    assert.ok(scroll, "the scroll viewport exists");
    assert.equal(scroll.props["data-has-overflow-x"], undefined, "40px by default");
    assert.equal(observers.length, 2, "the scroller and the tab list are observed");

    Object.defineProperty(scroll, "scrollWidth", { configurable: true, value: 400 });
    Object.defineProperty(scroll, "clientWidth", { configurable: true, value: 300 });
    observers[1]([]);
    await vue.nextTick();
    assert.equal(scroll.props["data-has-overflow-x"], "", "overflow grows the viewport");

    // Font swap: the tabs now fit, but the scroll element's own box did not
    // change, so only the fonts.ready re-measure can catch it.
    Object.defineProperty(scroll, "scrollWidth", { configurable: true, value: 300 });
    resolveFonts();
    await fontsReady;
    await vue.nextTick();
    await vue.nextTick();
    assert.equal(
      scroll.props["data-has-overflow-x"],
      undefined,
      "the 50px state is not left behind by a stale pre-font measurement",
    );
  } finally {
    app.unmount();
    if (originalResizeObserver === undefined) delete globalThis.ResizeObserver;
    else globalThis.ResizeObserver = originalResizeObserver;
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
  }
});
