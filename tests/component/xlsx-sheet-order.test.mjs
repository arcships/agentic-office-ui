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

function hiddenSheetBuffer(name, sheetName) {
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
  sheet.setAttribute("state", "hidden");
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
    const objectUrl = `blob:sheet-order-test-${nextObjectUrl += 1}`;
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

function dropIndicator(root) {
  return findByTestId(root, "xlsx-sheet-drop-indicator");
}

function dragEvent({ clientX = 0, left = 0, width = 100, dataTransfer } = {}) {
  const transfer = dataTransfer ?? {
    dropEffect: "",
    effectAllowed: "",
    setData() {},
  };
  return {
    clientX,
    currentTarget: {
      getBoundingClientRect: () => ({ bottom: 24, height: 24, left, right: left + width, top: 0, width }),
    },
    dataTransfer: transfer,
    preventDefault() {
      this.defaultPrevented = true;
    },
    stopPropagation() {
      this.propagationStopped = true;
    },
  };
}

function dispatchKeydown(node, event) {
  const handler = node.props.onKeydown;
  for (const entry of Array.isArray(handler) ? handler : [handler]) {
    entry?.(event);
  }
}

configureWasm();

// ── Controller API ───────────────────────────────────────────────────────────

test("moveSheet reorders worksheets and keeps the moved sheet active", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "P&L", "Notes"]);
    controller.setActiveTabIndex(1);
    const originalProfit = controller.getCellDisplayValue({ row: 1, col: 1 });

    assert.equal(controller.moveSheet(0, 2), true);
    await vue.nextTick();

    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["P&L", "Notes", "Assumptions"]);
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), ["P&L", "Notes", "Assumptions"]);
    assert.deepEqual(controller.workbook.sheetNames, ["P&L", "Notes", "Assumptions"]);
    assert.equal(controller.activeTabIndex, 2, "the moved tab is active");
    assert.equal(controller.activeTab?.name, "Assumptions");
    assert.equal(controller.activeSheetIndex, 2);
    assert.equal(controller.activeSheet?.name, "Assumptions");
    assert.equal(
      controller.tabs[0].sheetIndex,
      0,
      "a worksheet tab keeps its visible sheet index after the move",
    );

    controller.setActiveTabIndex(0);
    assert.equal(
      controller.getCellDisplayValue({ row: 1, col: 1 }),
      originalProfit,
      "a cross-sheet formula still calculates after the move",
    );

    assert.equal(controller.moveSheet(2, 0), true);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "P&L", "Notes"]);
    assert.equal(controller.activeTab?.name, "Assumptions");
    assert.equal(controller.activeSheetIndex, 0);
  } finally {
    app.unmount();
  }
});

test("moveSheet keeps the selection when the moved tab was already active", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    const cell = { row: 2, col: 1 };
    controller.selectCell(cell);
    assert.deepEqual(controller.activeCell, cell);

    assert.equal(controller.moveSheet(0, 1), true);
    await vue.nextTick();
    assert.equal(controller.activeTabIndex, 1);
    assert.deepEqual(controller.activeCell, cell, "the move does not clear the selection");
    assert.deepEqual(controller.selection, { end: cell, start: cell });
  } finally {
    app.unmount();
  }
});

test("moveSheet rejects out-of-range indices, a same-slot drop and read-only viewers", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    for (const [from, to] of [[-1, 0], [0, 3], [3, 0], [0.5, 1], [0, 1.5], [1, 1]]) {
      assert.equal(controller.moveSheet(from, to), false, `${from} -> ${to}`);
    }
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "P&L", "Notes"]);
    assert.equal(controller.canUndo, false, "a rejected move records no history");
  } finally {
    app.unmount();
  }

  const readOnly = await mountController("financial-model.xlsx", { readOnly: true });
  try {
    assert.equal(readOnly.controller.moveSheet(0, 1), false);
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

test("moveSheet maps tab indices across a workbook with a chartsheet", async () => {
  const { app, controller } = await mountController("charts-images.xlsx");
  try {
    assert.deepEqual(
      controller.tabs.map((tab) => `${tab.kind}:${tab.name}`),
      ["sheet:Dashboard", "chartsheet:Revenue Chart"],
    );
    assert.equal(controller.setFreezePanes({ col: 1, row: 1 }), true);

    assert.equal(controller.moveSheet(0, 1), true);
    await vue.nextTick();
    assert.deepEqual(
      controller.tabs.map((tab) => `${tab.kind}:${tab.name}`),
      ["chartsheet:Revenue Chart", "sheet:Dashboard"],
    );
    assert.equal(controller.activeTabIndex, 1);
    assert.equal(controller.activeTab?.kind, "sheet", "the moved worksheet is active");
    assert.equal(controller.activeSheet?.name, "Dashboard");
    assert.equal(controller.chartsheets.length, 1, "the chartsheet survives the move");
    assert.ok(controller.chartsheets[0].charts?.length > 0, "the chartsheet chart survives the move");
    assert.equal(controller.getSheetCharts(0).length, 1, "charts stay on Dashboard");
    assert.ok(controller.getSheetImages(0).length > 0, "images stay on Dashboard");
    assert.ok(controller.getSheetShapes(0).length > 0, "shapes stay on Dashboard");
    assert.deepEqual(
      controller.freezePanes,
      { col: 1, row: 1 },
      "the sheet state follows the worksheet, not its old chart slot",
    );

    assert.equal(controller.moveSheet(0, 1), true);
    await vue.nextTick();
    assert.deepEqual(
      controller.tabs.map((tab) => `${tab.kind}:${tab.name}`),
      ["sheet:Dashboard", "chartsheet:Revenue Chart"],
    );
    assert.equal(controller.activeTabIndex, 1);
    assert.equal(controller.activeTab?.kind, "chartsheet", "the moved chartsheet is active");
  } finally {
    app.unmount();
  }
});

test("images, charts and sheet state follow their sheet across the reorder", async () => {
  const { app, controller } = await mountController("charts-images.xlsx");
  try {
    controller.addSheet("Data");
    await vue.nextTick();
    assert.deepEqual(controller.tabs.map((tab) => tab.name), [
      "Dashboard",
      "Revenue Chart",
      "Data",
    ]);

    // Move the empty sheet to the front: Dashboard becomes worksheet 1.
    assert.equal(controller.moveSheet(2, 0), true);
    await vue.nextTick();
    assert.deepEqual(controller.tabs.map((tab) => tab.name), [
      "Data",
      "Dashboard",
      "Revenue Chart",
    ]);
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), ["Data", "Dashboard"]);
    assert.equal(controller.sheets[0].workbookSheetIndex, 0);
    assert.equal(controller.sheets[1].workbookSheetIndex, 1);
    assert.equal(controller.getSheetImages(0).length, 0, "the new sheet has no images");
    assert.ok(controller.getSheetImages(1).length > 0, "Dashboard keeps its image");
    assert.equal(controller.getSheetCharts(0).length, 0);
    assert.equal(controller.getSheetCharts(1).length, 1, "Dashboard keeps its chart");
    assert.equal(controller.getSheetShapes(0).length, 0);
    assert.ok(controller.getSheetShapes(1).length > 0, "Dashboard keeps its shape");
  } finally {
    app.unmount();
  }
});

test("hidden sheets keep their slot when a visible tab moves", async () => {
  const { app, controller } = await mountController(
    hiddenSheetBuffer("financial-model.xlsx", "P&L"),
  );
  try {
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "Notes"]);
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), ["Assumptions", "Notes"]);
    assert.equal(controller.workbook.getSheet(1).visibility, "hidden");

    controller.selectCell({ col: 1, row: 1 });
    assert.equal(controller.setFreezePanes({ col: 1, row: 1 }), true);

    assert.equal(controller.moveSheet(0, 1), true);
    await vue.nextTick();
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Notes", "Assumptions"]);
    assert.deepEqual(controller.workbook.sheetNames, ["Notes", "P&L", "Assumptions"]);
    assert.equal(controller.workbook.getSheet(1).name, "P&L", "the hidden sheet keeps its slot");
    assert.equal(controller.workbook.getSheet(1).visibility, "hidden");
    assert.equal(controller.activeTab?.name, "Assumptions");
    assert.deepEqual(controller.freezePanes, { col: 1, row: 1 }, "state follows the visible sheet");

    controller.setActiveTabIndex(0);
    assert.equal(controller.activeTab?.name, "Notes");
    assert.equal(controller.freezePanes, null, "the pane did not stay on the old slot");
  } finally {
    app.unmount();
  }
});

test("showHiddenSheets exposes hidden tabs to the reorder API", async () => {
  const { app, controller } = await mountController(
    hiddenSheetBuffer("financial-model.xlsx", "P&L"),
    { showHiddenSheets: true },
  );
  try {
    assert.deepEqual(controller.tabs.map((tab) => tab.name), [
      "Assumptions",
      "P&L",
      "Notes",
    ]);
    assert.equal(controller.moveSheet(0, 2), true);
    await vue.nextTick();
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["P&L", "Notes", "Assumptions"]);
    assert.equal(controller.activeTab?.name, "Assumptions");
    assert.equal(controller.workbook.getSheet(0).visibility, "hidden");
  } finally {
    app.unmount();
  }
});

test("sheet-scoped authoring state follows its sheet across the reorder", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    controller.setActiveTabIndex(1);
    assert.equal(controller.setFreezePanes({ col: 1, row: 1 }), true);
    assert.deepEqual(controller.freezePanes, { col: 1, row: 1 });

    assert.equal(controller.moveSheet(1, 0), true);
    await vue.nextTick();
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), ["P&L", "Assumptions", "Notes"]);
    assert.equal(controller.activeTab?.name, "P&L");
    assert.deepEqual(controller.freezePanes, { col: 1, row: 1 }, "the pane stays with P&L");

    controller.setActiveTabIndex(1);
    assert.equal(controller.activeTab?.name, "Assumptions");
    assert.equal(controller.freezePanes, null, "the pane does not follow the old slot");
  } finally {
    app.unmount();
  }
});

test("a reorder is one undo step and redo restores the new order", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    assert.equal(controller.moveSheet(0, 2), true);
    assert.equal(controller.canUndo, true);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["P&L", "Notes", "Assumptions"]);
    assert.equal(controller.activeTabIndex, 2);

    controller.undo();
    await waitFor(
      () => controller.tabs[0]?.name === "Assumptions",
      "undo restores the original order",
    );
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "P&L", "Notes"]);
    assert.equal(controller.activeTabIndex, 0, "undo restores the old active tab");

    await waitFor(() => controller.canRedo, "the snapshot moves to the redo stack");
    controller.redo();
    await waitFor(
      () => controller.tabs[0]?.name === "P&L",
      "redo restores the moved order",
    );
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["P&L", "Notes", "Assumptions"]);
    assert.equal(controller.activeTabIndex, 2, "redo restores the moved active tab");
  } finally {
    app.unmount();
  }
});

test("undo restores a chartsheet that was active before the move", async () => {
  const { app, controller } = await mountController("charts-images.xlsx");
  try {
    controller.setActiveTabIndex(1);
    assert.equal(controller.activeTab?.kind, "chartsheet");

    assert.equal(controller.moveSheet(1, 0), true);
    await vue.nextTick();
    assert.equal(controller.activeTab?.kind, "chartsheet");
    assert.equal(controller.activeTabIndex, 0);

    controller.undo();
    await waitFor(
      () => controller.tabs[0]?.kind === "sheet",
      "undo restores the original tab order",
    );
    assert.equal(controller.activeTabIndex, 1);
    assert.equal(controller.activeTab?.kind, "chartsheet", "the chartsheet is still active");

    await waitFor(() => controller.canRedo, "the chartsheet move moves to the redo stack");
    controller.redo();
    await waitFor(
      () => controller.tabs[0]?.kind === "chartsheet",
      "redo restores the moved chartsheet order",
    );
    assert.equal(controller.activeTabIndex, 0);
    assert.equal(controller.activeTab?.kind, "chartsheet");
  } finally {
    app.unmount();
  }
});

test("an exported reorder survives export and reload with its formulas intact", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController("financial-model.xlsx");
  let exportedBytes;
  try {
    assert.equal(controller.moveSheet(0, 2), true);
    await vue.nextTick();
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
    exportedBytes.buffer.slice(
      exportedBytes.byteOffset,
      exportedBytes.byteOffset + exportedBytes.byteLength,
    ),
  );
  try {
    assert.deepEqual(reopened.controller.sheets.map((sheet) => sheet.name), [
      "P&L",
      "Notes",
      "Assumptions",
    ]);
    assert.deepEqual(reopened.controller.tabs.map((tab) => tab.name), [
      "P&L",
      "Notes",
      "Assumptions",
    ]);
    reopened.controller.setActiveTabIndex(0);
    assert.equal(
      reopened.controller.getCellDisplayValue({ row: 1, col: 1 }),
      "$50,400",
      "the reloaded workbook still resolves the cross-sheet formula",
    );
  } finally {
    reopened.app.unmount();
  }
});

test("an exported chartsheet move survives export and reload", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController("charts-images.xlsx");
  let exportedBytes;
  try {
    assert.equal(controller.moveSheet(0, 1), true);
    await vue.nextTick();
    assert.deepEqual(
      controller.tabs.map((tab) => `${tab.kind}:${tab.name}`),
      ["chartsheet:Revenue Chart", "sheet:Dashboard"],
    );
    controller.exportXlsx();
    await waitFor(() => downloads.length > 0, "XLSX export download");
    const { blob } = downloads.pop();
    exportedBytes = new Uint8Array(await blob.arrayBuffer());
  } finally {
    app.unmount();
  }

  const reopened = await mountController(
    exportedBytes.buffer.slice(
      exportedBytes.byteOffset,
      exportedBytes.byteOffset + exportedBytes.byteLength,
    ),
  );
  try {
    assert.deepEqual(
      reopened.controller.tabs.map((tab) => `${tab.kind}:${tab.name}`),
      ["chartsheet:Revenue Chart", "sheet:Dashboard"],
    );
    assert.equal(reopened.controller.chartsheets.length, 1);
    assert.ok(
      reopened.controller.chartsheets[0].charts?.length > 0,
      "the chartsheet chart survives export and reload",
    );
    assert.equal(reopened.controller.getSheetCharts(0).length, 1);
  } finally {
    reopened.app.unmount();
  }
});

// ── Drag interaction ────────────────────────────────────────────────────────

test("dragging a tab past another tab reorders the strip and shows an insertion indicator", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    const source = tabWrapper(root, "Assumptions");
    const target = tabWrapper(root, "P&L");

    source.props.onDragstart(dragEvent());
    await vue.nextTick();
    assert.equal(source.props["data-dragging"], "true");
    assert.ok(dropIndicator(root) === undefined, "no indicator before a dragover");

    target.props.onDragover(dragEvent({ clientX: 90 }));
    await vue.nextTick();
    const indicator = dropIndicator(root);
    assert.ok(indicator, "the insertion indicator appears between tabs");
    assert.equal(indicator.props["data-drop-slot"], 2, "the right half targets the slot after P&L");

    target.props.onDrop(dragEvent({ clientX: 90 }));
    await vue.nextTick();

    assert.equal(dropIndicator(root), undefined, "the indicator disappears after the drop");
    assert.deepEqual(tabOrder(root), ["P&L", "Assumptions", "Notes"]);
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["P&L", "Assumptions", "Notes"]);
    assert.equal(controller.activeTabIndex, 1);
    assert.equal(controller.activeTab?.name, "Assumptions");
    const activeTab = tabButton(root, "Assumptions");
    assert.equal(activeTab.props["aria-selected"], true, "the moved tab is marked selected");
    assert.equal(activeTab.props["data-visibility"], "visible");
    assert.match(activeTab.props.class, /xlsx-sheettabs__tab--active/);
    assert.equal(activeTab.props.tabindex, 0, "the active tab is the roving tab stop");
    assert.equal(tabButton(root, "P&L").props["aria-selected"], false);
  } finally {
    app.unmount();
  }
});

test("dropping on a tab's left half inserts before it and brings the moved tab into view", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    const wrappers = ["Assumptions", "P&L", "Notes"].map((name) => tabWrapper(root, name));
    const revealed = [];
    for (const wrapper of wrappers) {
      wrapper.scrollIntoView = (options) => revealed.push(options);
    }

    const movedWrapper = tabWrapper(root, "Notes");

    movedWrapper.props.onDragstart(dragEvent());
    const first = tabWrapper(root, "Assumptions");
    first.props.onDragover(dragEvent({ clientX: 10 }));
    await vue.nextTick();
    assert.equal(dropIndicator(root)?.props["data-drop-slot"], 0);

    first.props.onDrop(dragEvent({ clientX: 10 }));
    await waitFor(() => revealed.length > 0, "the moved tab is scrolled into view");
    assert.deepEqual(tabOrder(root), ["Notes", "Assumptions", "P&L"]);
    assert.equal(controller.activeTabIndex, 0);
    assert.equal(controller.activeTab?.name, "Notes");
    assert.deepEqual(revealed[0], { block: "nearest", inline: "nearest" });
  } finally {
    app.unmount();
  }
});

test("dropping a tab on its own slot is a no-op", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    const source = tabWrapper(root, "P&L");
    source.props.onDragstart(dragEvent());
    source.props.onDragover(dragEvent({ clientX: 10 }));
    await vue.nextTick();
    assert.equal(dropIndicator(root)?.props["data-drop-slot"], 1, "the left half targets its own slot");

    source.props.onDrop(dragEvent({ clientX: 10 }));
    await vue.nextTick();
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "P&L", "Notes"]);
    assert.equal(controller.canUndo, false, "a same-slot drop records no history");
    assert.equal(dropIndicator(root), undefined);
  } finally {
    app.unmount();
  }
});

test("Esc during a drag cancels it without changing the order", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    const source = tabWrapper(root, "Assumptions");
    const target = tabWrapper(root, "P&L");
    source.props.onDragstart(dragEvent());
    target.props.onDragover(dragEvent({ clientX: 90 }));
    await vue.nextTick();
    assert.ok(dropIndicator(root), "the drag is in progress");

    dispatchKeydown(target, {
      preventDefault() {
        this.defaultPrevented = true;
      },
      key: "Escape",
      stopPropagation() {
        this.propagationStopped = true;
      },
    });
    await vue.nextTick();
    assert.equal(dropIndicator(root), undefined, "Esc removes the indicator");
    assert.equal(source.props["data-dragging"], undefined);

    target.props.onDrop(dragEvent({ clientX: 90 }));
    await vue.nextTick();
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "P&L", "Notes"]);
    assert.equal(controller.canUndo, false, "a cancelled drag records no history");
  } finally {
    app.unmount();
  }
});

test("a drag registers a global Escape listener and removes it on drop", async () => {
  const { app, root } = await mountTabs("financial-model.xlsx");
  const previousDocument = globalThis.document;
  const listeners = [];
  globalThis.document = {
    addEventListener(event, listener, capture) {
      listeners.push({ capture, event, listener });
    },
    removeEventListener(event, listener, capture) {
      const index = listeners.findIndex(
        (entry) => entry.event === event && entry.listener === listener && entry.capture === capture,
      );
      if (index >= 0) listeners.splice(index, 1);
    },
  };
  try {
    const source = tabWrapper(root, "Assumptions");
    const target = tabWrapper(root, "P&L");
    source.props.onDragstart(dragEvent());
    assert.deepEqual(
      listeners.map((entry) => [entry.event, entry.capture]),
      [["keydown", true]],
      "the drag listens for Escape outside the tab strip",
    );
    target.props.onDragover(dragEvent({ clientX: 90 }));
    target.props.onDrop(dragEvent({ clientX: 90 }));
    await vue.nextTick();
    assert.equal(listeners.length, 0, "the listener is removed after the drop");
    assert.equal(dropIndicator(root), undefined);
  } finally {
    if (previousDocument === undefined) delete globalThis.document;
    else globalThis.document = previousDocument;
    app.unmount();
  }
});

test("read-only viewers expose no drag affordance and ignore drops", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx", { readOnly: true });
  try {
    const source = tabWrapper(root, "Assumptions");
    const target = tabWrapper(root, "P&L");
    assert.equal(source.props.draggable, false, "read-only tabs are not draggable");

    source.props.onDragstart(dragEvent());
    target.props.onDragover(dragEvent({ clientX: 90 }));
    await vue.nextTick();
    assert.equal(dropIndicator(root), undefined, "no drag starts in read-only mode");
    target.props.onDrop(dragEvent({ clientX: 90 }));
    assert.deepEqual(controller.tabs.map((tab) => tab.name), ["Assumptions", "P&L", "Notes"]);
  } finally {
    app.unmount();
  }
});

test("a chartsheet tab can be dragged and keeps its chartsheet content", async () => {
  const { app, root, controller } = await mountTabs("charts-images.xlsx");
  try {
    const source = tabWrapper(root, "Revenue Chart");
    const target = tabWrapper(root, "Dashboard");
    source.props.onDragstart(dragEvent());
    target.props.onDragover(dragEvent({ clientX: 10 }));
    await vue.nextTick();
    assert.equal(dropIndicator(root)?.props["data-drop-slot"], 0);

    target.props.onDrop(dragEvent({ clientX: 10 }));
    await vue.nextTick();
    assert.deepEqual(
      controller.tabs.map((tab) => `${tab.kind}:${tab.name}`),
      ["chartsheet:Revenue Chart", "sheet:Dashboard"],
    );
    assert.equal(controller.activeTabIndex, 0);
    assert.equal(controller.activeTab?.kind, "chartsheet");
    assert.equal(controller.chartsheets.length, 1);
    assert.equal(controller.getSheetCharts(0).length, 1);
  } finally {
    app.unmount();
  }
});

test("a failed move leaves the tab strip unchanged and reports the reason", async () => {
  const { app, root, controller } = await mountTabs("financial-model.xlsx");
  try {
    const originalMoveSheet = controller.moveSheet;
    controller.moveSheet = () => false;
    const source = tabWrapper(root, "Assumptions");
    const target = tabWrapper(root, "P&L");
    source.props.onDragstart(dragEvent());
    target.props.onDragover(dragEvent({ clientX: 90 }));
    target.props.onDrop(dragEvent({ clientX: 90 }));
    await vue.nextTick();

    const error = findByTestId(root, "xlsx-sheet-move-error");
    assert.ok(error, "the failure is visible");
    assert.equal(error.children?.[0]?.text, messages.sheetTabs.moveFailed);
    assert.deepEqual(tabOrder(root), ["Assumptions", "P&L", "Notes"]);
    controller.moveSheet = originalMoveSheet;
  } finally {
    app.unmount();
  }
});
