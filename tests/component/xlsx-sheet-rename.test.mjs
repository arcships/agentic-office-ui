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
    const objectUrl = `blob:rename-test-${nextObjectUrl += 1}`;
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
  const buffer = workbookBuffer(file);
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: buffer,
        fileName: file,
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
    `${file} tabs controller`,
  );
  assert.ifError(controller.error);
  await waitFor(
    () => walk(mounted.root).some((node) => node.props?.["data-testid"] === "xlsx-sheet-tab"),
    `${file} tab strip`,
  );
  return { ...mounted, controller };
}

function tabButton(root, sheetName) {
  return walk(root).find(
    (node) =>
      node.props?.["data-testid"] === "xlsx-sheet-tab"
      && node.props?.["data-sheet-name"] === sheetName,
  );
}

function keydown(key) {
  return {
    key,
    preventDefault() {},
    stopPropagation() {},
  };
}

// Vue merges the Enter and Escape bindings into one keydown prop (an array of
// handlers); the host renderer has to run them the way the DOM would.
function dispatchKeydown(node, event) {
  const handler = node.props.onKeydown;
  for (const entry of Array.isArray(handler) ? handler : [handler]) {
    entry?.(event);
  }
}

function tabWrapper(tab) {
  return tab.parent;
}

async function beginRename(root, sheetName) {
  const tab = tabButton(root, sheetName);
  assert.ok(tab, `${sheetName} tab is rendered`);
  tab.props.onDblclick({});
  await vue.nextTick();
  const input = findByTestId(root, "xlsx-sheet-rename-input");
  assert.ok(input, "double-click opens the inline input");
  return input;
}

configureWasm();

test("renameSheet renames in place, keeps the tab index and the cross-sheet formulas", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), [
      "Assumptions",
      "P&L",
      "Notes",
    ]);
    controller.setActiveTabIndex(1);
    const activeTabIndex = controller.activeTabIndex;
    const activeSheetIndex = controller.activeSheetIndex;
    const originalProfit = controller.getCellDisplayValue({ row: 1, col: 1 });
    assert.equal(controller.setNamedRange("Rates", "Assumptions!$B$2:$D$2"), true);

    assert.equal(controller.renameSheet(0, "  Inputs & Rates  "), true);
    await vue.nextTick();

    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), [
      "Inputs & Rates",
      "P&L",
      "Notes",
    ]);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), [
      "Inputs & Rates",
      "P&L",
      "Notes",
    ]);
    assert.deepEqual(controller.workbook.sheetNames, ["Inputs & Rates", "P&L", "Notes"]);
    assert.equal(controller.activeTabIndex, activeTabIndex, "the active tab does not move");
    assert.equal(controller.activeSheetIndex, activeSheetIndex, "worksheet order is unchanged");
    assert.equal(
      controller.getCellDisplayValue({ row: 1, col: 1 }),
      originalProfit,
      "a formula that points at the renamed sheet still calculates",
    );
    assert.equal(controller.getSheetNameError("inputs & rates", 0), null, "case-only rename is legal");
    assert.equal(
      controller.getNamedRanges().find((entry) => entry.name === "Rates")?.refersTo,
      "'Inputs & Rates'!$B$2:$D$2",
      "a defined name follows the renamed sheet",
    );

    // A second rename works on the rebuilt workbook and keeps every reference.
    assert.equal(controller.renameSheet(0, "Inputs"), true);
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), ["Inputs", "P&L", "Notes"]);
    assert.equal(
      controller.getNamedRanges().find((entry) => entry.name === "Rates")?.refersTo,
      "Inputs!$B$2:$D$2",
    );
    assert.equal(
      controller.getCellDisplayValue({ row: 1, col: 1 }),
      originalProfit,
      "the second rename keeps the cross-sheet formula working",
    );
  } finally {
    app.unmount();
  }
});

test("renameSheet rejects Excel's invalid names and leaves the workbook untouched", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    const before = controller.sheets.map((sheet) => sheet.name);
    for (const [name, issue] of [
      ["", "empty"],
      ["   ", "empty"],
      ["x".repeat(32), "too-long"],
      ["A/B", "invalid-characters"],
      ["Q1 [final]", "invalid-characters"],
      ["'Quoted", "apostrophe"],
      ["P&L", "duplicate"],
      ["p&l", "duplicate"],
    ]) {
      assert.equal(controller.getSheetNameError(name, 0), issue, JSON.stringify(name));
      assert.equal(controller.renameSheet(0, name), false, JSON.stringify(name));
    }
    assert.equal(controller.renameSheet(9, "Data"), false, "an unknown index is rejected");
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), before);
    assert.equal(controller.canUndo, false, "a rejected rename records no history");
  } finally {
    app.unmount();
  }
});

test("renameSheet checks chartsheet names too and cannot rename a chartsheet tab", async () => {
  const { app, controller } = await mountController("charts-images.xlsx");
  try {
    assert.deepEqual(controller.tabs.map((tab) => `${tab.kind}:${tab.name}`), [
      "sheet:Dashboard",
      "chartsheet:Revenue Chart",
    ]);
    assert.equal(controller.getSheetNameError("revenue chart", 0), "duplicate");
    assert.equal(controller.renameSheet(0, "Revenue Chart"), false);
    assert.equal(controller.sheets[0].name, "Dashboard");

    controller.setActiveTabIndex(1);
    assert.equal(controller.renameActiveSheet("Charts"), false, "a chartsheet is not a worksheet");

    controller.setActiveTabIndex(0);
    assert.equal(controller.renameActiveSheet("Overview"), true);
    assert.equal(controller.activeTab?.name, "Overview");

    // The image/chart assets are re-read from the rewritten package, so they
    // stay attached to the sheet and follow the new name.
    assert.ok(controller.images.length > 0, "images survive the rename");
    assert.equal(controller.getSheetCharts(0).length, 1, "charts stay on the sheet");
    const seriesFormula = controller.getChartSeriesFormula(controller.charts[0].id, 0);
    assert.match(seriesFormula, /SERIES\(Overview!/);
    assert.doesNotMatch(seriesFormula, /Dashboard/);
  } finally {
    app.unmount();
  }
});

test("a worksheet rename is one undo step and redo restores the new name", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    assert.equal(controller.renameSheet(2, "Notes 2026"), true);
    assert.equal(controller.canUndo, true);
    assert.equal(controller.tabs[2].name, "Notes 2026");

    controller.undo();
    await waitFor(() => controller.tabs[2]?.name === "Notes", "undo restores the old name");
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), [
      "Assumptions",
      "P&L",
      "Notes",
    ]);

    await waitFor(() => controller.canRedo, "the snapshot moves to the redo stack");
    controller.redo();
    await waitFor(() => controller.tabs[2]?.name === "Notes 2026", "redo restores the rename");
  } finally {
    app.unmount();
  }
});

test("an exported rename survives export and reload with its formulas intact", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController("financial-model.xlsx");
  let exportedBytes;
  try {
    const originalProfit = (() => {
      controller.setActiveTabIndex(1);
      return controller.getCellDisplayValue({ row: 1, col: 1 });
    })();
    assert.equal(controller.renameSheet(0, "Inputs"), true);

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
    assert.deepEqual(reopened.controller.sheets.map((sheet) => sheet.name), [
      "Inputs",
      "P&L",
      "Notes",
    ]);
    reopened.controller.setActiveTabIndex(1);
    assert.equal(
      reopened.controller.getCellDisplayValue({ row: 1, col: 1 }),
      "$50,400",
      "the reloaded workbook still resolves the renamed sheet",
    );
  } finally {
    reopened.app.unmount();
  }
});

test("a read-only viewer refuses to rename", async () => {
  const { app, controller } = await mountController("financial-model.xlsx", { readOnly: true });
  try {
    assert.equal(controller.renameSheet(0, "Inputs"), false);
    assert.equal(controller.renameActiveSheet("Inputs"), false);
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), [
      "Assumptions",
      "P&L",
      "Notes",
    ]);
  } finally {
    app.unmount();
  }
});

test("double-click opens an inline tab editor: Enter commits and updates the strip", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    const input = await beginRename(root, "Assumptions");
    assert.equal(input.props.value, "Assumptions", "the editor starts with the current name");
    assert.equal(input.props["aria-label"], messages.sheetTabs.renameLabel);

    input.props.onInput({ target: { value: "  Inputs  " } });
    await vue.nextTick();
    dispatchKeydown(input, keydown("Enter"));
    await vue.nextTick();

    assert.equal(findByTestId(root, "xlsx-sheet-rename-input"), undefined, "Enter closes the editor");
    assert.equal(controller.sheets[0].name, "Inputs", "surrounding whitespace is trimmed");
    assert.ok(tabButton(root, "Inputs"), "the strip renders the new label");
  } finally {
    app.unmount();
  }
});

test("Esc cancels the inline rename and leaves the sheet name alone", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    const input = await beginRename(root, "P&L");
    input.props.onInput({ target: { value: "Profit" } });
    await vue.nextTick();
    dispatchKeydown(input, keydown("Escape"));
    await vue.nextTick();

    assert.equal(findByTestId(root, "xlsx-sheet-rename-input"), undefined);
    assert.equal(controller.sheets[1].name, "P&L");
    assert.ok(tabButton(root, "P&L"));
  } finally {
    app.unmount();
  }
});

test("blur commits a valid inline rename", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    const input = await beginRename(root, "Notes");
    input.props.onInput({ target: { value: "Notes 2026" } });
    await vue.nextTick();
    input.props.onBlur({});
    await vue.nextTick();

    assert.equal(findByTestId(root, "xlsx-sheet-rename-input"), undefined);
    assert.equal(controller.sheets[2].name, "Notes 2026");
  } finally {
    app.unmount();
  }
});

test("an empty or duplicate inline name keeps the editor open with a visible error", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    const input = await beginRename(root, "Assumptions");

    input.props.onInput({ target: { value: "   " } });
    await vue.nextTick();
    dispatchKeydown(input, keydown("Enter"));
    await vue.nextTick();
    let error = findByTestId(root, "xlsx-sheet-rename-error");
    assert.ok(error, "the editor reports the rejection");
    assert.equal(error.children?.[0]?.text, messages.sheetTabs.renameEmpty);
    assert.ok(findByTestId(root, "xlsx-sheet-rename-input"), "the editor stays open");
    assert.equal(controller.sheets[0].name, "Assumptions");

    input.props.onInput({ target: { value: "p&l" } });
    await vue.nextTick();
    assert.equal(findByTestId(root, "xlsx-sheet-rename-error"), undefined, "typing clears the error");
    dispatchKeydown(input, keydown("Enter"));
    await vue.nextTick();
    error = findByTestId(root, "xlsx-sheet-rename-error");
    assert.equal(error.children?.[0]?.text, messages.sheetTabs.renameDuplicate);
    assert.ok(findByTestId(root, "xlsx-sheet-rename-input"), "a duplicate keeps the editor open");
    assert.equal(controller.sheets[0].name, "Assumptions");

    input.props.onInput({ target: { value: "A/B" } });
    await vue.nextTick();
    input.props.onBlur({});
    await vue.nextTick();
    error = findByTestId(root, "xlsx-sheet-rename-error");
    assert.equal(error.children?.[0]?.text, messages.diagnostics.sheetNameInvalid);
    assert.ok(findByTestId(root, "xlsx-sheet-rename-input"), "blur keeps an invalid editor open");
  } finally {
    app.unmount();
  }
});

test("the tab context menu offers Rename and closes after it opens the editor", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    const tab = tabButton(root, "P&L");
    tabWrapper(tab).props.onContextmenu({
      clientX: 40,
      clientY: 20,
      preventDefault() {},
    });
    await vue.nextTick();

    assert.ok(findByTestId(root, "xlsx-sheet-tab-menu"), "the tab menu opens");
    assert.equal(controller.activeTabIndex, 1, "right-click selects the tab it targets");
    const renameItem = findByTestId(root, "xlsx-rename-sheet");
    assert.equal(renameItem.children?.[0]?.text, messages.sheetTabs.rename);

    // The backdrop closes the menu without touching the sheet.
    findByTestId(root, "xlsx-sheet-tab-menu-backdrop").props.onPointerdown();
    await vue.nextTick();
    assert.equal(findByTestId(root, "xlsx-sheet-tab-menu"), undefined);

    tab.props.onKeydown({
      key: "ContextMenu",
      preventDefault() {},
      currentTarget: { getBoundingClientRect: () => ({ left: 8, bottom: 20 }) },
    });
    await vue.nextTick();
    assert.ok(findByTestId(root, "xlsx-sheet-tab-menu"), "the ContextMenu key opens the same menu");

    findByTestId(root, "xlsx-rename-sheet").props.onClick();
    await vue.nextTick();
    assert.equal(findByTestId(root, "xlsx-sheet-tab-menu"), undefined, "the menu closes");
    const input = findByTestId(root, "xlsx-sheet-rename-input");
    assert.ok(input, "Rename opens the inline editor");
    assert.equal(input.props.value, "P&L");
  } finally {
    app.unmount();
  }
});

test("chartsheet tabs and read-only viewers expose no rename affordance", async () => {
  const charts = await mountTabs("charts-images.xlsx");
  try {
    const chartTab = tabButton(charts.root, "Revenue Chart");
    chartTab.props.onDblclick({});
    await vue.nextTick();
    assert.equal(findByTestId(charts.root, "xlsx-sheet-rename-input"), undefined);
    tabWrapper(chartTab).props.onContextmenu({ clientX: 1, clientY: 1, preventDefault() {} });
    await vue.nextTick();
    assert.equal(findByTestId(charts.root, "xlsx-sheet-tab-menu"), undefined);
  } finally {
    charts.app.unmount();
  }

  const readOnly = await mountTabs("financial-model.xlsx", { readOnly: true });
  try {
    const tab = tabButton(readOnly.root, "Assumptions");
    tab.props.onDblclick({});
    await vue.nextTick();
    assert.equal(findByTestId(readOnly.root, "xlsx-sheet-rename-input"), undefined);
    tabWrapper(tab).props.onContextmenu({ clientX: 1, clientY: 1, preventDefault() {} });
    await vue.nextTick();
    assert.equal(findByTestId(readOnly.root, "xlsx-sheet-tab-menu"), undefined);
  } finally {
    readOnly.app.unmount();
  }
});

test("the tab menu keeps sheet insert/delete while the strip buttons stay opt-in", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    // #160: the default strip renders tabs only (upstream geometry); the
    // Excel add/remove buttons are behind `showSheetActions`.
    assert.equal(findByTestId(root, "xlsx-add-sheet"), undefined, "no add button by default");
    assert.equal(findByTestId(root, "xlsx-remove-sheet"), undefined, "no remove button by default");

    const openMenu = async (name) => {
      const tab = tabButton(root, name);
      tabWrapper(tab).props.onContextmenu({ clientX: 40, clientY: 20, preventDefault() {} });
      await vue.nextTick();
      assert.ok(findByTestId(root, "xlsx-sheet-tab-menu"), "the tab menu opens");
    };

    await openMenu("P&L");
    findByTestId(root, "xlsx-insert-sheet").props.onClick();
    await vue.nextTick();
    assert.equal(findByTestId(root, "xlsx-sheet-tab-menu"), undefined, "the menu closes");
    assert.equal(controller.tabs.length, 4, "Insert sheet adds a worksheet tab");

    await openMenu("Notes");
    const before = controller.tabs.length;
    findByTestId(root, "xlsx-delete-sheet").props.onClick();
    await vue.nextTick();
    assert.equal(findByTestId(root, "xlsx-sheet-tab-menu"), undefined, "the menu closes");
    assert.equal(controller.tabs.length, before - 1, "Delete sheet removes the active worksheet");
  } finally {
    app.unmount();
  }
});

test("showSheetActions restores Excel's arrows, add and remove buttons", async () => {
  const buffer = workbookBuffer("financial-model.xlsx");
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({ file: buffer, fileName: "financial-model.xlsx", useWorker: false });
      return () =>
        controller.sourceState === "ready"
          ? vue.h(XlsxSheetTabs, { controller, showSheetActions: true })
          : vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  try {
    await waitFor(() => controller.sourceState === "ready", "controller ready");
    assert.ok(findByTestId(mounted.root, "xlsx-add-sheet"), "add button");
    assert.ok(findByTestId(mounted.root, "xlsx-remove-sheet"), "remove button");
    assert.ok(
      walk(mounted.root).some((node) => String(node.props?.class ?? "").includes("xlsx-sheettabs--actions")),
      "the opt-in strip switches to the actions layout",
    );
  } finally {
    mounted.app.unmount();
  }
});
