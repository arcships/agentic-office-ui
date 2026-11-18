import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

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
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
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
const { officeMessagesEn } = await import(
  new URL("../../packages/office-interaction/src/i18n.ts", import.meta.url)
);
const { XlsxSheetTabs, useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

const messages = officeMessagesEn.xlsx;

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

/** Rewrites the workbook part so one named worksheet gets `state`. */
function visibilityBuffer(name, sheetName, visibility) {
  const source = new Uint8Array(readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  ));
  const archive = unzipSync(source);
  const workbookPath = "xl/workbook.xml";
  const document = new XmlDomParser().parseFromString(
    strFromU8(archive[workbookPath]),
    "application/xml",
  );
  const sheet = Array.from(document.getElementsByTagName("*"))
    .find((node) => node.localName === "sheet" && node.getAttribute("name") === sheetName);
  assert.ok(sheet, `${sheetName} exists in ${name}`);
  if (visibility === "visible") {
    sheet.removeAttribute("state");
  } else {
    sheet.setAttribute("state", visibility);
  }
  archive[workbookPath] = strToU8(new XMLSerializer().serializeToString(document));
  const bytes = zipSync(archive, { level: 6 });
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function installDownloadCapture(t) {
  const downloads = [];
  const blobs = new Map();
  const originalDocument = globalThis.document;
  const originalCreateObjectUrl = URL.createObjectURL;
  const originalRevokeObjectUrl = URL.revokeObjectURL;
  let nextObjectUrl = 0;

  globalThis.document = {
    body: { append() {} },
    createElement() {
      return {
        href: "",
        download: "",
        click() {
          downloads.push({ blob: blobs.get(this.href), fileName: this.download });
        },
        remove() {},
      };
    },
  };
  URL.createObjectURL = (blob) => {
    const objectUrl = `blob:sheet-visibility-test-${nextObjectUrl += 1}`;
    blobs.set(objectUrl, blob);
    return objectUrl;
  };
  URL.revokeObjectURL = () => {};

  t.after(() => {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    URL.createObjectURL = originalCreateObjectUrl;
    URL.revokeObjectURL = originalRevokeObjectUrl;
  });

  return downloads;
}

async function mountController(file, options = {}) {
  const input = typeof file === "string" ? workbookBuffer(file) : file;
  const fileName = typeof file === "string" ? file : "exported.xlsx";
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({ file: input, fileName, useWorker: false, ...options });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    `${fileName} controller`,
  );
  assert.ifError(controller.error);
  return { ...mounted, controller };
}

async function mountTabs(file, options = {}) {
  const input = typeof file === "string" ? workbookBuffer(file) : file;
  const fileName = typeof file === "string" ? file : "visibility.xlsx";
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: input,
        fileName,
        useWorker: false,
        ...options,
      });
      return () =>
        controller.sourceState === "ready"
          ? vue.h(XlsxSheetTabs, { controller, readOnly: options.readOnly === true })
          : vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    `${fileName} tabs controller`,
  );
  assert.ifError(controller.error);
  await waitFor(
    () => walk(mounted.root).some((node) => node.props?.["data-testid"] === "xlsx-sheet-tab"),
    `${fileName} tab strip`,
  );
  return { ...mounted, controller };
}

function tabButton(root, sheetName) {
  return walk(root).find(
    (node) =>
      node.props?.["data-testid"] === "xlsx-sheet-tab"
      && node.props["data-sheet-name"] === sheetName,
  );
}

function tabWrapper(root, sheetName) {
  const tab = tabButton(root, sheetName);
  assert.ok(tab, `${sheetName} tab is rendered`);
  return tab.parent;
}

function tabOrder(root) {
  return walk(root)
    .filter((node) => node.props?.["data-testid"] === "xlsx-sheet-tab")
    .map((node) => node.props["data-sheet-name"]);
}

function classNames(node) {
  return String(node?.props?.class ?? "").split(/\s+/).filter(Boolean);
}

function openTabMenu(root, sheetName, overrides = {}) {
  const wrapper = tabWrapper(root, sheetName);
  wrapper.props.onContextmenu({
    clientX: 40,
    clientY: 20,
    preventDefault() {},
    ...overrides,
  });
}

function openUnhideList(root) {
  const unhide = findByTestId(root, "xlsx-unhide-menu");
  assert.ok(unhide, "the Unhide menu entry is rendered");
  unhide.props.onClick();
}

function unhideItems(root) {
  return walk(root)
    .filter((node) => node.props?.["data-testid"] === "xlsx-unhide-sheet")
    .map((node) => node.props["data-sheet-name"]);
}

configureWasm();

// ── Controller API ───────────────────────────────────────────────────────────

test("setSheetVisibility hides the active sheet, selects the nearest visible tab and is undoable", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    controller.setActiveTabIndex(1);
    assert.equal(controller.activeTab?.name, "P&L");

    assert.equal(controller.setSheetVisibility(1, "hidden"), true);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "Notes"]);
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), ["Assumptions", "Notes"]);
    assert.equal(controller.activeTab?.name, "Notes", "the next visible tab takes over");
    assert.equal(controller.activeSheet?.name, "Notes");
    assert.equal(controller.workbook.getSheet(1).visibility, "hidden");
    assert.equal(controller.getSheetVisibility(1), "hidden");
    assert.deepEqual(controller.hiddenSheets, [
      { name: "P&L", visibility: "hidden", workbookSheetIndex: 1 },
    ]);
    assert.equal(controller.canUndo, true, "the hide is one undo entry");

    controller.undo();
    await waitFor(() => controller.tabs.length === 3, "undo restores the hidden tab");
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "P&L", "Notes"]);
    assert.equal(controller.activeTab?.name, "P&L", "undo restores the active tab");
    assert.equal(controller.getSheetVisibility(1), "visible");
    assert.deepEqual(controller.hiddenSheets, []);

    await waitFor(() => controller.canRedo, "the hide moves to the redo stack");
    controller.redo();
    await waitFor(() => controller.tabs.length === 2, "redo hides the tab again");
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "Notes"]);
    assert.equal(controller.activeTab?.name, "Notes");
  } finally {
    app.unmount();
  }
});

test("hiding the last tab falls back to the previous visible sheet", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    controller.setActiveTabIndex(2);
    assert.equal(controller.activeTab?.name, "Notes");

    assert.equal(controller.setSheetVisibility(2, "hidden"), true);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "P&L"]);
    assert.equal(controller.activeTab?.name, "P&L", "there is no next tab, so the previous one wins");
    assert.equal(controller.activeSheet?.name, "P&L");
  } finally {
    app.unmount();
  }
});

test("hiding an inactive sheet keeps the active sheet and refuses the last visible sheet", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    controller.setActiveTabIndex(2);
    const cell = { row: 1, col: 1 };
    controller.selectCell(cell);
    assert.equal(controller.setSheetVisibility(1, "hidden"), true);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "Notes"]);
    assert.equal(controller.activeTab?.name, "Notes", "the active tab did not move");
    assert.equal(controller.activeTabIndex, 1, "the tab index follows the shortened strip");
    assert.deepEqual(controller.activeCell, cell, "hiding another sheet keeps the selection");
    assert.deepEqual(controller.selection, { end: cell, start: cell });

    assert.equal(controller.setSheetVisibility(0, "hidden"), true);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Notes"]);
    assert.equal(controller.activeTab?.name, "Notes");

    assert.equal(
      controller.setSheetVisibility(2, "hidden"),
      false,
      "the last visible sheet cannot be hidden",
    );
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Notes"]);
    assert.equal(controller.getSheetVisibility(2), "visible");
  } finally {
    app.unmount();
  }
});

test("setSheetVisibility rejects invalid indices, values, no-ops and read-only viewers", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    for (const index of [-1, 3, 0.5, Number.NaN]) {
      assert.equal(controller.setSheetVisibility(index, "hidden"), false, `index ${index}`);
    }
    assert.equal(controller.setSheetVisibility(0, "gone"), false);
    assert.equal(controller.setSheetVisibility(0, "visible"), false, "visible is already a no-op");
    assert.equal(controller.getSheetVisibility(-1), null);
    assert.equal(controller.getSheetVisibility(3), null);
    assert.equal(controller.canUndo, false);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "P&L", "Notes"]);
  } finally {
    app.unmount();
  }

  const readOnly = await mountController("financial-model.xlsx", { readOnly: true });
  try {
    assert.equal(readOnly.controller.setSheetVisibility(0, "hidden"), false);
    assert.deepEqual(readOnly.controller.tabs.map((tab) => tab.name), [
      "Assumptions",
      "P&L",
      "Notes",
    ]);
    assert.equal(readOnly.controller.canUndo, false);
  } finally {
    readOnly.app.unmount();
  }
});

test("unhide restores the worksheet, shows it and makes it active", async () => {
  const { app, controller } = await mountController(
    visibilityBuffer("financial-model.xlsx", "P&L", "hidden"),
  );
  try {
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "Notes"]);
    assert.deepEqual(controller.hiddenSheets, [
      { name: "P&L", visibility: "hidden", workbookSheetIndex: 1 },
    ]);

    assert.equal(controller.setSheetVisibility(1, "visible"), true);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "P&L", "Notes"]);
    assert.equal(controller.activeTab?.name, "P&L", "unhiding selects the restored sheet");
    assert.equal(controller.activeSheet?.name, "P&L");
    assert.equal(controller.getSheetVisibility(1), "visible");
    assert.deepEqual(controller.hiddenSheets, []);
    assert.equal(controller.canUndo, true);

    controller.undo();
    await waitFor(() => controller.tabs.length === 2, "undo hides the sheet again");
    assert.deepEqual(controller.hiddenSheets.map((sheet) => sheet.name), ["P&L"]);
  } finally {
    app.unmount();
  }
});

test("veryHidden sheets stay out of the tab strip and Unhide list but the API can restore them", async () => {
  const { app, controller } = await mountController(
    visibilityBuffer("financial-model.xlsx", "P&L", "veryHidden"),
    { showHiddenSheets: true },
  );
  try {
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "Notes"]);
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), ["Assumptions", "Notes"]);
    assert.equal(controller.getSheetVisibility(1), "veryHidden");
    assert.deepEqual(controller.hiddenSheets, [], "veryHidden is not an Unhide candidate");

    assert.equal(controller.setSheetVisibility(1, "visible"), true);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "P&L", "Notes"]);
    assert.equal(controller.activeTab?.name, "P&L");
  } finally {
    app.unmount();
  }
});

test("showHiddenSheets includes hidden tabs as hidden entries and toggles at runtime", async () => {
  const { app, controller } = await mountController(
    visibilityBuffer("financial-model.xlsx", "P&L", "hidden"),
  );
  try {
    assert.equal(controller.showHiddenSheets, false);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "Notes"]);
    controller.setActiveTabIndex(1);
    const cell = { row: 2, col: 2 };
    controller.selectCell(cell);

    assert.equal(controller.setShowHiddenSheets(true), true);
    assert.equal(controller.showHiddenSheets, true);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "P&L", "Notes"]);
    assert.equal(controller.tabs[1].visibility, "hidden");
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), ["Assumptions", "P&L", "Notes"]);
    assert.equal(controller.activeTab?.name, "Notes", "the active sheet stays active");
    assert.deepEqual(controller.activeCell, cell, "a view toggle keeps the selection");
    assert.equal(controller.canUndo, false, "showing hidden tabs is a view option, not an edit");

    assert.equal(controller.setShowHiddenSheets(false), true);
    assert.equal(controller.showHiddenSheets, false);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "Notes"]);
  } finally {
    app.unmount();
  }
});

test("turning hidden tabs off while a hidden tab is active selects the nearest visible sheet", async () => {
  const { app, controller } = await mountController(
    visibilityBuffer("financial-model.xlsx", "P&L", "hidden"),
    { showHiddenSheets: true },
  );
  try {
    controller.setActiveTabIndex(1);
    assert.equal(controller.activeTab?.name, "P&L");
    assert.equal(controller.activeTab?.visibility, "hidden");

    assert.equal(controller.setShowHiddenSheets(false), true);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "Notes"]);
    assert.equal(controller.activeTab?.name, "Notes");
    assert.equal(controller.activeSheet?.name, "Notes");
    assert.deepEqual(controller.activeCell, null);
  } finally {
    app.unmount();
  }
});

test("charts, images and shapes stay with their sheet across hide and unhide", async () => {
  const { app, controller } = await mountController("charts-images.xlsx");
  try {
    controller.addSheet("Data");
    await vue.nextTick();
    assert.deepEqual(
      controller.tabs.map((tab) => `${tab.kind}:${tab.name}`),
      ["sheet:Dashboard", "chartsheet:Revenue Chart", "sheet:Data"],
    );

    assert.equal(controller.setSheetVisibility(0, "hidden"), true);
    assert.equal(controller.activeTab?.name, "Data");
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), ["Data"]);
    assert.equal(controller.sheets[0].workbookSheetIndex, 1);
    assert.equal(controller.getSheetCharts(0).length, 0);
    assert.equal(controller.getSheetImages(0).length, 0);
    assert.equal(controller.getSheetShapes(0).length, 0);

    assert.equal(controller.setSheetVisibility(0, "visible"), true);
    assert.equal(controller.activeTab?.name, "Dashboard");
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), ["Dashboard", "Data"]);
    assert.equal(controller.getSheetCharts(0).length, 1, "Dashboard keeps its chart");
    assert.ok(controller.getSheetImages(0).length > 0, "Dashboard keeps its image");
    assert.ok(controller.getSheetShapes(0).length > 0, "Dashboard keeps its shape");
    assert.equal(controller.chartsheets.length, 1);
    assert.ok(controller.chartsheets[0].charts?.length > 0);
  } finally {
    app.unmount();
  }
});

test("hiding a worksheet keeps an active chartsheet tab selected", async () => {
  const { app, controller } = await mountController("charts-images.xlsx");
  try {
    controller.addSheet("Data");
    await vue.nextTick();
    controller.setActiveTabIndex(1);
    assert.equal(controller.activeTab?.kind, "chartsheet");

    assert.equal(controller.setSheetVisibility(0, "hidden"), true);
    assert.deepEqual(
      controller.tabs.map((tab) => `${tab.kind}:${tab.name}`),
      ["chartsheet:Revenue Chart", "sheet:Data"],
    );
    assert.equal(controller.activeTabIndex, 0, "the chartsheet tab index is remapped");
    assert.equal(controller.activeTab?.kind, "chartsheet");
    assert.equal(controller.activeTab?.name, "Revenue Chart");
    assert.equal(controller.activeSheet, null);
  } finally {
    app.unmount();
  }
});

test("an exported hidden worksheet survives export and reload", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController("financial-model.xlsx");
  let exportedBytes;
  try {
    controller.setActiveTabIndex(1);
    assert.equal(controller.setSheetVisibility(1, "hidden"), true);
    controller.exportXlsx();
    await waitFor(() => downloads.length > 0, "XLSX export download");
    const { blob } = downloads.pop();
    assert.ok(blob, "export produced a blob");
    exportedBytes = new Uint8Array(await blob.arrayBuffer());
    assert.equal(new TextDecoder().decode(exportedBytes.slice(0, 2)), "PK", "export is a zip package");
  } finally {
    app.unmount();
  }

  const reopened = await mountController(
    exportedBytes.buffer.slice(exportedBytes.byteOffset, exportedBytes.byteOffset + exportedBytes.byteLength),
  );
  try {
    assert.deepEqual(reopened.controller.tabs.map((tab) => tab.name), ["Assumptions", "Notes"]);
    assert.equal(reopened.controller.workbook.getSheet(1).visibility, "hidden");
    assert.equal(reopened.controller.getSheetVisibility(1), "hidden");
    assert.deepEqual(reopened.controller.hiddenSheets.map((sheet) => sheet.name), ["P&L"]);

    assert.equal(reopened.controller.setSheetVisibility(1, "visible"), true);
    assert.deepEqual(reopened.controller.tabs.map((tab) => tab.name), [
      "Assumptions",
      "P&L",
      "Notes",
    ]);
  } finally {
    reopened.app.unmount();
  }
});

// ── Tab UI ──────────────────────────────────────────────────────────────────

test("the tab context menu hides a sheet and lists it under Unhide", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    openTabMenu(root, "P&L");
    await vue.nextTick();
    assert.ok(findByTestId(root, "xlsx-sheet-tab-menu"), "the tab menu opens");
    const hideItem = findByTestId(root, "xlsx-hide-sheet");
    assert.equal(hideItem.children?.[0]?.text, messages.sheetTabs.hide);
    assert.equal(hideItem.props.disabled, false, "P&L is not the last visible sheet");

    hideItem.props.onClick();
    await vue.nextTick();
    assert.equal(findByTestId(root, "xlsx-sheet-tab-menu"), undefined, "the menu closes");
    assert.deepEqual(tabOrder(root), ["Assumptions", "Notes"]);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "Notes"]);
    assert.equal(controller.activeTab?.name, "Notes");

    openTabMenu(root, "Assumptions");
    await vue.nextTick();
    openUnhideList(root);
    await vue.nextTick();
    assert.ok(findByTestId(root, "xlsx-unhide-sheet-menu"), "the Unhide list opens");
    assert.deepEqual(unhideItems(root), ["P&L"]);

    const entry = walk(root).find(
      (node) => node.props?.["data-testid"] === "xlsx-unhide-sheet"
        && node.props["data-sheet-name"] === "P&L",
    );
    entry.props.onClick();
    await vue.nextTick();
    assert.deepEqual(tabOrder(root), ["Assumptions", "P&L", "Notes"]);
    assert.equal(controller.activeTab?.name, "P&L", "the restored sheet is active");
    assert.equal(controller.getSheetVisibility(1), "visible");
  } finally {
    app.unmount();
  }
});

test("Hide is disabled when the menu targets the last visible sheet", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    assert.equal(controller.setSheetVisibility(0, "hidden"), true);
    assert.equal(controller.setSheetVisibility(1, "hidden"), true);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Notes"]);

    await vue.nextTick();
    openTabMenu(root, "Notes");
    await vue.nextTick();
    const hideItem = findByTestId(root, "xlsx-hide-sheet");
    assert.equal(hideItem.props.disabled, true);
    hideItem.props.onClick();
    await vue.nextTick();
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Notes"]);
  } finally {
    app.unmount();
  }
});

test("hidden tabs render dimmed with an unhide affordance", async () => {
  const { app, root, controller } = await mountTabs(
    visibilityBuffer("financial-model.xlsx", "P&L", "hidden"),
    { showHiddenSheets: true },
  );
  try {
    const tab = tabButton(root, "P&L");
    assert.ok(tab, "the hidden tab is listed when showHiddenSheets is on");
    assert.equal(tab.props["data-visibility"], "hidden");
    assert.ok(classNames(tab).includes("xlsx-sheettabs__tab--hidden"));

    const unhide = walk(root).find(
      (node) => node.props?.["data-testid"] === "xlsx-unhide-sheet-tab"
        && node.props["data-sheet-name"] === "P&L",
    );
    assert.ok(unhide, "the hidden tab offers an unhide control");
    assert.match(unhide.props["aria-label"], /P&L/);

    unhide.props.onClick({ stopPropagation() {} });
    await vue.nextTick();
    assert.equal(controller.getSheetVisibility(1), "visible");
    assert.equal(tabButton(root, "P&L").props["data-visibility"], "visible");
    assert.equal(controller.activeTab?.name, "P&L");
  } finally {
    app.unmount();
  }
});

test("veryHidden sheets never render a tab or an unhide affordance", async () => {
  const { app, root } = await mountTabs(
    visibilityBuffer("financial-model.xlsx", "P&L", "veryHidden"),
    { showHiddenSheets: true },
  );
  try {
    assert.equal(tabButton(root, "P&L"), undefined);
    assert.equal(
      walk(root).find((node) => node.props?.["data-testid"] === "xlsx-unhide-sheet-tab"),
      undefined,
    );

    openTabMenu(root, "Assumptions");
    await vue.nextTick();
    const unhide = findByTestId(root, "xlsx-unhide-menu");
    assert.equal(unhide.props.disabled, true, "there is no Unhide candidate");
    openUnhideList(root);
    await vue.nextTick();
    assert.deepEqual(unhideItems(root), []);
    assert.equal(
      findByTestId(root, "xlsx-unhide-sheet-menu").children?.[0]?.props?.["data-testid"],
      undefined,
      "the list contains only the empty-state text",
    );
  } finally {
    app.unmount();
  }
});

test("a read-only viewer exposes no hide or unhide controls", async () => {
  const { app, root } = await mountTabs(
    visibilityBuffer("financial-model.xlsx", "P&L", "hidden"),
    { showHiddenSheets: true, readOnly: true },
  );
  try {
    assert.ok(tabButton(root, "P&L"), "the hidden tab is still shown");
    assert.equal(
      walk(root).find((node) => node.props?.["data-testid"] === "xlsx-unhide-sheet-tab"),
      undefined,
    );
    openTabMenu(root, "Assumptions");
    await vue.nextTick();
    assert.equal(findByTestId(root, "xlsx-sheet-tab-menu"), undefined);
  } finally {
    app.unmount();
  }
});
