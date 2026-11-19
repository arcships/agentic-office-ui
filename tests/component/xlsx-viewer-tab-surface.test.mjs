import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
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
// The composite viewer renders the real grid, and the controller runtime reads
// browser timers. Only what those paths touch is stubbed; nothing here pretends
// to be a full DOM (the component renderer supplies the host nodes).
globalThis.ResizeObserver ??= class ResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
};
globalThis.requestAnimationFrame ??= (callback) => setTimeout(() => callback(Date.now()), 0);
globalThis.cancelAnimationFrame ??= (handle) => clearTimeout(handle);
// The grid paints its canvases in an animation frame after mount. The component
// renderer's host nodes have no 2D context, and painting is not part of the tab
// surface under test, so the paint paths bail out instead of drawing.
if (typeof Object.prototype.getContext !== "function") {
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value: () => null,
  });
}
globalThis.window ??= {
  addEventListener() {},
  cancelAnimationFrame: globalThis.cancelAnimationFrame,
  clearInterval: globalThis.clearInterval,
  clearTimeout: globalThis.clearTimeout,
  devicePixelRatio: 1,
  getComputedStyle: () => ({ getPropertyValue: () => "" }),
  removeEventListener() {},
  requestAnimationFrame: globalThis.requestAnimationFrame,
  setInterval: globalThis.setInterval,
  setTimeout: globalThis.setTimeout,
};

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { XlsxViewer, useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

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

function tablist(root) {
  return walk(root).find((node) => node.props?.role === "tablist");
}

function tabButtons(root) {
  return walk(root).filter((node) => node.props?.["data-testid"] === "xlsx-sheet-tab");
}

function tabButton(root, sheetName) {
  return tabButtons(root).find((node) => node.props?.["data-sheet-name"] === sheetName);
}

function clickTab(root, sheetName) {
  const button = tabButton(root, sheetName);
  assert.ok(button, `${sheetName} tab is rendered`);
  button.props.onClick?.({ type: "click" });
}

async function createController(file) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(file),
        fileName: file,
        useWorker: false,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    `${file} controller`,
  );
  assert.ifError(controller.error);
  return { ...mounted, controller };
}

/**
 * Mounts `XlsxViewer` around an in-test controller and records the tab surface
 * events. The viewer mounts while the workbook is still loading, which is what
 * a restored `activeTabIndex` meets in a real host.
 */
async function mountViewer(file, options = {}) {
  const {
    activeTabIndex,
    controller: providedController,
    showFormulaBar = false,
    showRibbon = false,
    showSheetTabs,
  } = options;
  const controlled = activeTabIndex !== undefined;
  const tabIndex = vue.ref(activeTabIndex);
  const tabChanges = [];
  const updates = [];
  let controller = providedController;
  const Harness = vue.defineComponent({
    setup() {
      controller ??= useXlsxViewerController({
        file: workbookBuffer(file),
        fileName: file,
        useWorker: false,
      });
      return () => vue.h(XlsxViewer, {
        controller,
        activeTabIndex: controlled ? tabIndex.value : undefined,
        showSheetTabs,
        showDefaultToolbar: false,
        showFormulaBar,
        showImages: false,
        showRibbon,
        onTabChange: (tab) => { tabChanges.push(tab); },
        "onUpdate:activeTabIndex": (index) => {
          updates.push(index);
          if (controlled) tabIndex.value = index;
        },
      });
    },
  });
  const mounted = await mount(Harness);
  if (!providedController) {
    await waitFor(
      () => controller.sourceState === "ready" || controller.error,
      `${file} controller`,
    );
    assert.ifError(controller.error);
  }
  await waitFor(() => controller.tabs.length > 0, `${file} tabs`);
  await vue.nextTick();
  return { ...mounted, controller, tabChanges, tabIndex, updates };
}

configureWasm();

// ── showSheetTabs ──────────────────────────────────────────────────────────

test("the sheet tab strip renders by default and showSheetTabs hides only the strip", async () => {
  const open = await mountViewer("financial-model.xlsx", {
    showFormulaBar: true,
    showRibbon: true,
  });
  try {
    assert.ok(tablist(open.root), "the strip is rendered by default");
    assert.deepEqual(
      tabButtons(open.root).map((node) => node.props["data-sheet-name"]),
      ["Assumptions", "P&L", "Notes"],
    );
  } finally {
    open.app.unmount();
  }

  const hidden = await mountViewer("financial-model.xlsx", {
    showFormulaBar: true,
    showRibbon: true,
    showSheetTabs: false,
  });
  try {
    assert.equal(tablist(hidden.root), undefined, "no tablist is rendered");
    assert.deepEqual(tabButtons(hidden.root), [], "no tab buttons are rendered");
    assert.ok(findByTestId(hidden.root, "xlsx-grid"), "the grid stays mounted");
    assert.ok(findByTestId(hidden.root, "xlsx-format-toolbar"), "the ribbon stays mounted");
    assert.ok(findByTestId(hidden.root, "xlsx-name-box"), "the formula bar stays mounted");
    assert.deepEqual(
      hidden.controller.tabs.map((tab) => tab.name),
      ["Assumptions", "P&L", "Notes"],
      "hiding the strip does not drop the controller's tabs",
    );
    assert.equal(hidden.controller.activeTabIndex, 0);
    assert.equal(hidden.controller.activeSheet?.name, "Assumptions");
  } finally {
    hidden.app.unmount();
  }
});

test("a hidden strip still switches sheets through the controller and reports the switch", async () => {
  const viewer = await mountViewer("financial-model.xlsx", { showSheetTabs: false });
  try {
    viewer.controller.setActiveTabIndex(2);
    await waitFor(
      () => viewer.controller.activeSheet?.name === "Notes",
      "the controller switches to Notes without a strip",
    );
    assert.equal(viewer.controller.activeTab?.kind, "sheet");
    assert.deepEqual(viewer.tabChanges.map((tab) => tab.name), ["Notes"]);
    assert.deepEqual(viewer.updates, [2], "the index is reported for v-model hosts");
    assert.equal(tablist(viewer.root), undefined, "the strip stays hidden");
  } finally {
    viewer.app.unmount();
  }
});

// ── Uncontrolled (controller-owned) state ──────────────────────────────────

test("the opening tab is not a switch and re-renders do not repeat the event", async () => {
  const viewer = await mountViewer("financial-model.xlsx");
  try {
    assert.deepEqual(viewer.tabChanges, [], "loading a workbook is not a tab switch");
    assert.deepEqual(viewer.updates, [], "the opening index is not echoed");

    viewer.controller.selectCell({ row: 1, col: 1 });
    await vue.nextTick();
    viewer.controller.selectCell({ row: 3, col: 2 });
    await vue.nextTick();
    assert.deepEqual(viewer.tabChanges, [], "internal re-renders never report a switch");
  } finally {
    viewer.app.unmount();
  }
});

test("a tab click reports update:activeTabIndex and tabChange exactly once", async () => {
  const viewer = await mountViewer("financial-model.xlsx");
  try {
    clickTab(viewer.root, "P&L");
    await waitFor(
      () => viewer.controller.activeTabIndex === 1,
      "the clicked tab becomes active",
    );
    await vue.nextTick();

    assert.deepEqual(viewer.updates, [1]);
    assert.equal(viewer.tabChanges.length, 1);
    const [tab] = viewer.tabChanges;
    assert.equal(tab.id, "sheet-1");
    assert.equal(tab.index, 1);
    assert.equal(tab.kind, "sheet", "the event carries the full tab");
    assert.equal(tab.name, "P&L");
    assert.equal(tab.sheetIndex, 1);
    assert.equal(tab.visibility, "visible");

    clickTab(viewer.root, "P&L");
    await vue.nextTick();
    assert.deepEqual(viewer.updates, [1], "activating the active tab is not a switch");
    assert.equal(viewer.tabChanges.length, 1);

    clickTab(viewer.root, "Notes");
    await waitFor(
      () => viewer.controller.activeTabIndex === 2,
      "the second switch lands",
    );
    await vue.nextTick();
    assert.deepEqual(viewer.updates, [1, 2]);
    assert.deepEqual(viewer.tabChanges.map((entry) => entry.name), ["P&L", "Notes"]);
  } finally {
    viewer.app.unmount();
  }
});

test("an undefined activeTabIndex leaves a controller-restored tab untouched", async () => {
  const owner = await createController("financial-model.xlsx");
  try {
    owner.controller.setActiveTabIndex(2);
    assert.equal(owner.controller.activeTab?.name, "Notes");

    const viewer = await mountViewer("financial-model.xlsx", { controller: owner.controller });
    try {
      assert.equal(viewer.controller.activeTabIndex, 2, "the viewer does not reset the tab");
      assert.equal(viewer.controller.activeSheet?.name, "Notes");
      assert.deepEqual(viewer.tabChanges, [], "the restored tab is the opening tab");

      clickTab(viewer.root, "Assumptions");
      await waitFor(
        () => owner.controller.activeTabIndex === 0,
        "the uncontrolled viewer keeps switching",
      );
      assert.deepEqual(viewer.updates, [0]);
      assert.deepEqual(viewer.tabChanges.map((tab) => tab.name), ["Assumptions"]);
    } finally {
      viewer.app.unmount();
    }
  } finally {
    owner.app.unmount();
  }
});

// ── Controlled activeTabIndex ──────────────────────────────────────────────

test("a controlled activeTabIndex is applied after an asynchronous load without echoing", async () => {
  const viewer = await mountViewer("financial-model.xlsx", { activeTabIndex: 2 });
  try {
    assert.equal(viewer.controller.activeTabIndex, 2);
    assert.equal(viewer.controller.activeTab?.name, "Notes");
    assert.equal(viewer.controller.activeSheet?.name, "Notes");
    assert.deepEqual(viewer.tabChanges, [], "the host's own restore is not a switch");
    assert.deepEqual(viewer.updates, [], "the host already holds the index");

    viewer.tabIndex.value = 0;
    await waitFor(
      () => viewer.controller.activeTabIndex === 0,
      "the controlled value is mirrored onto the controller",
    );
    assert.equal(viewer.controller.activeSheet?.name, "Assumptions");
    assert.deepEqual(viewer.tabChanges, []);
    assert.deepEqual(viewer.updates, []);
  } finally {
    viewer.app.unmount();
  }
});

test("v-model:active-tab-index round-trips a user switch exactly once", async () => {
  const viewer = await mountViewer("financial-model.xlsx", { activeTabIndex: 0 });
  try {
    clickTab(viewer.root, "Notes");
    await waitFor(
      () => viewer.controller.activeTabIndex === 2,
      "Notes becomes active",
    );
    await vue.nextTick();

    assert.deepEqual(viewer.updates, [2]);
    assert.equal(viewer.tabChanges.length, 1);
    assert.equal(viewer.tabChanges[0].name, "Notes");
    assert.equal(viewer.tabIndex.value, 2, "the host state follows the update event");

    await vue.nextTick();
    assert.deepEqual(viewer.updates, [2], "the v-model write-back is not echoed");
    assert.equal(viewer.tabChanges.length, 1, "the write-back does not switch again");
    assert.equal(viewer.controller.activeTabIndex, 2);
    assert.equal(viewer.controller.activeTab?.name, "Notes");
  } finally {
    viewer.app.unmount();
  }
});

test("an out-of-range controlled index is ignored until it becomes valid", async () => {
  const viewer = await mountViewer("financial-model.xlsx", { activeTabIndex: 9 });
  try {
    assert.equal(viewer.controller.activeTabIndex, 0, "the controller keeps a valid tab");
    assert.deepEqual(viewer.tabChanges, []);

    viewer.tabIndex.value = 2;
    await waitFor(
      () => viewer.controller.activeTabIndex === 2,
      "a valid controlled index is applied",
    );
    assert.equal(viewer.controller.activeTab?.name, "Notes");
    assert.deepEqual(viewer.tabChanges, []);
  } finally {
    viewer.app.unmount();
  }
});

// ── Chartsheet tabs ────────────────────────────────────────────────────────

test("tabChange carries chartsheets, whose tab index is not a sheet index", async () => {
  const viewer = await mountViewer("charts-images.xlsx");
  try {
    const chartsheetTab = viewer.controller.tabs[1];
    assert.equal(chartsheetTab.kind, "chartsheet", "the sample interleaves a chartsheet");
    assert.equal(chartsheetTab.sheetIndex, undefined);
    assert.equal(chartsheetTab.chartsheetIndex, 0);
    assert.equal(viewer.controller.sheets.length, 1, "there is only one worksheet");

    viewer.controller.setActiveTabIndex(1);
    await waitFor(
      () => viewer.controller.activeTab?.kind === "chartsheet",
      "the chartsheet tab becomes active",
    );
    await vue.nextTick();

    assert.deepEqual(viewer.updates, [1]);
    assert.equal(viewer.tabChanges.length, 1);
    const [tab] = viewer.tabChanges;
    assert.equal(tab.kind, "chartsheet");
    assert.equal(tab.name, "Revenue Chart");
    assert.equal(tab.index, 1);
    assert.equal(tab.chartsheetIndex, 0);
    assert.equal(tab.sheetIndex, undefined);
    assert.equal(
      viewer.controller.sheets[tab.index],
      undefined,
      "a tab index cannot be used as a sheet index",
    );

    viewer.controller.setActiveTabIndex(0);
    await waitFor(
      () => viewer.controller.activeTab?.kind === "sheet",
      "the worksheet tab becomes active again",
    );
    await vue.nextTick();
    assert.deepEqual(viewer.updates, [1, 0]);
    assert.deepEqual(
      viewer.tabChanges.map((entry) => [entry.kind, entry.name]),
      [["chartsheet", "Revenue Chart"], ["sheet", "Dashboard"]],
    );
  } finally {
    viewer.app.unmount();
  }
});

// ── Rebuilds that are not switches ─────────────────────────────────────────

test("renames and reorders rebuild the strip without reporting a tab switch", async () => {
  const viewer = await mountViewer("financial-model.xlsx");
  try {
    clickTab(viewer.root, "P&L");
    await waitFor(
      () => viewer.controller.activeTabIndex === 1,
      "P&L becomes active",
    );
    await vue.nextTick();
    assert.equal(viewer.tabChanges.length, 1);

    assert.equal(viewer.controller.renameSheet(1, "Profit"), true, "the sheet is renamed");
    await waitFor(
      () => viewer.controller.tabs[1]?.name === "Profit",
      "the strip rebuilds with the new name",
    );
    await vue.nextTick();
    assert.equal(
      viewer.tabChanges.length,
      1,
      "a new tab object for the same id is not a switch",
    );
    assert.equal(viewer.controller.activeTab?.name, "Profit");

    assert.equal(viewer.controller.moveSheet(1, 0), true, "the active tab is reordered");
    await waitFor(
      () => viewer.controller.activeTabIndex === 0,
      "the moved tab keeps its index in the strip",
    );
    await vue.nextTick();
    assert.equal(viewer.controller.activeTab?.name, "Profit");
    assert.equal(
      viewer.tabChanges.length,
      1,
      "moving the active tab changes its index, not the active tab",
    );
    assert.deepEqual(viewer.updates, [1, 0], "v-model hosts still follow the index move");
  } finally {
    viewer.app.unmount();
  }
});
