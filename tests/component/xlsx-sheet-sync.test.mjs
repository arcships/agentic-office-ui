import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
// Table definitions only exist once the structure parser can read sheet XML.
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

function workbookBuffer(name) {
  const bytes = readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function configureWasm() {
  const bytes = readFileSync(
    new URL(
      "../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm",
      import.meta.url,
    ),
  );
  setWasmSource(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  );
}

async function mountController(fileName, options = {}) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(fileName),
        fileName,
        useWorker: false,
        ...options,
      });
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

/**
 * The grid renders `activeSheet` (which comes from the active tab), while the
 * chart/image overlays read `activeSheetIndex`. Both must describe one sheet.
 */
function assertActiveSheetIsConsistent(controller, expectedName) {
  const activeTab = controller.tabs[controller.activeTabIndex];
  assert.equal(controller.activeSheet?.name, expectedName, "grid sheet");
  assert.equal(activeTab?.name, expectedName, "tab strip sheet");
  assert.equal(activeTab?.sheetIndex, controller.activeSheetIndex, "overlay sheet index");
}

configureWasm();

test("XLSX cross-sheet undo realigns the grid, the tab strip and the overlays", async () => {
  const { app, controller } = await mountController("financial-model.xlsx", {
    skipXmlParsing: true,
  });
  const cell = { col: 0, row: 1 };

  try {
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), [
      "Assumptions",
      "P&L",
      "Notes",
    ]);

    controller.setActiveSheetIndex(0);
    await vue.nextTick();
    controller.selectCell(cell);
    const originalValue = controller.getCellDisplayValue(cell);
    controller.setCellValue(cell, "UNDO-ME");
    await vue.nextTick();
    assert.equal(controller.getCellDisplayValue(cell), "UNDO-ME");

    controller.setActiveTabIndex(1);
    await vue.nextTick();
    assertActiveSheetIsConsistent(controller, "P&L");
    assert.equal(controller.selection, null, "a user tab switch still clears the selection");

    controller.undo();
    await vue.nextTick();
    await vue.nextTick();

    assertActiveSheetIsConsistent(controller, "Assumptions");
    assert.equal(controller.activeSheetIndex, 0);
    assert.equal(controller.getCellDisplayValue(cell), originalValue);
    assert.deepEqual(controller.selection, { start: cell, end: cell });
    assert.deepEqual(controller.activeCell, cell);
  } finally {
    app.unmount();
  }
});

test("XLSX cross-sheet range undo restores A's selection onto A, not onto the visible sheet", async () => {
  const { app, controller } = await mountController("financial-model.xlsx", {
    skipXmlParsing: true,
  });
  const range = { start: { col: 0, row: 0 }, end: { col: 1, row: 1 } };

  try {
    controller.setActiveSheetIndex(0);
    await vue.nextTick();
    controller.selectRange(range);
    controller.setRangeStyle(range, { bold: true });
    await vue.nextTick();

    controller.setActiveTabIndex(2);
    await vue.nextTick();
    assertActiveSheetIsConsistent(controller, "Notes");

    controller.undo();
    await vue.nextTick();
    await vue.nextTick();

    assertActiveSheetIsConsistent(controller, "Assumptions");
    assert.equal(controller.activeTabIndex, 0);
    assert.deepEqual(controller.selection, range, "the range is painted on the sheet it belongs to");

    // The preserved switch must not leak into the next user tab change.
    controller.setActiveTabIndex(1);
    await vue.nextTick();
    assert.equal(controller.selection, null);
    assert.equal(controller.activeCell, null);
  } finally {
    app.unmount();
  }
});

test("XLSX undo from a chartsheet tab returns to the sheet tab that owns the edit", async () => {
  const { app, controller } = await mountController("charts-images.xlsx");
  const cell = { col: 0, row: 0 };

  try {
    const chartsheetTabIndex = controller.tabs.findIndex((tab) => tab.kind === "chartsheet");
    assert.ok(chartsheetTabIndex > 0, "fixture exposes a chartsheet tab");

    controller.setActiveSheetIndex(0);
    await vue.nextTick();
    controller.selectCell(cell);
    controller.setCellValue(cell, "UNDO-ME");
    await vue.nextTick();

    controller.setActiveTabIndex(chartsheetTabIndex);
    await vue.nextTick();
    assert.equal(controller.activeSheet, null, "a chartsheet tab renders no worksheet");

    controller.undo();
    await vue.nextTick();
    await vue.nextTick();

    assertActiveSheetIsConsistent(controller, "Dashboard");
    assert.equal(controller.activeTabIndex, 0);
    assert.equal(controller.getSheetCharts().length, 1, "chart overlays follow the restored sheet");
    assert.equal(controller.getSheetImages().length, 1, "image overlays follow the restored sheet");
    assert.deepEqual(controller.selection, { start: cell, end: cell });
  } finally {
    app.unmount();
  }
});

test("XLSX structural undo restores the selection it took the snapshot with", async () => {
  const { app, controller } = await mountController("financial-model.xlsx", {
    skipXmlParsing: true,
  });
  const range = { start: { col: 2, row: 2 }, end: { col: 3, row: 3 } };

  try {
    controller.setActiveTabIndex(1);
    await vue.nextTick();
    controller.selectRange(range);
    await vue.nextTick();

    controller.addSheet("Extra");
    await vue.nextTick();
    assert.equal(controller.sheets.length, 4);
    assertActiveSheetIsConsistent(controller, "Extra");

    controller.undo();
    await waitFor(() => controller.canRedo, "snapshot undo");

    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), [
      "Assumptions",
      "P&L",
      "Notes",
    ]);
    assertActiveSheetIsConsistent(controller, "P&L");
    assert.deepEqual(controller.selection, range);
    assert.deepEqual(controller.activeCell, range.end);
  } finally {
    app.unmount();
  }
});

test("XLSX deleting a middle sheet clears the selection and follows the tab strip", async () => {
  const { app, controller } = await mountController("financial-model.xlsx", {
    skipXmlParsing: true,
  });

  try {
    controller.setActiveTabIndex(1);
    await vue.nextTick();
    controller.selectRange({ start: { col: 3, row: 3 }, end: { col: 4, row: 4 } });
    await vue.nextTick();

    controller.removeActiveSheet();
    await vue.nextTick();

    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), ["Assumptions", "Notes"]);
    assertActiveSheetIsConsistent(controller, "Notes");
    assert.equal(controller.activeSheetIndex, 1);
    assert.equal(controller.selection, null, "the deleted sheet's selection must not survive");
    assert.equal(controller.activeCell, null);

    controller.undo();
    await waitFor(() => controller.canRedo, "snapshot undo of the removal");
    assert.deepEqual(controller.sheets.map((sheet) => sheet.name), [
      "Assumptions",
      "P&L",
      "Notes",
    ]);
    assertActiveSheetIsConsistent(controller, "P&L");
  } finally {
    app.unmount();
  }
});

test("XLSX read-only controllers refuse to sort a table", async () => {
  const readOnly = await mountController("sales-table.xlsx", { readOnly: true });

  try {
    assert.equal(readOnly.controller.readOnly, true);
    assert.deepEqual(
      readOnly.controller.tables.map((table) => table.name),
      ["SalesRecords"],
      "fixture exposes a sortable table",
    );
    const before = [1, 2, 3].map((row) =>
      readOnly.controller.getCellDisplayValue({ col: 0, row }),
    );

    readOnly.controller.sortTable("SalesRecords", 0, "descending");
    await vue.nextTick();

    assert.deepEqual(
      [1, 2, 3].map((row) => readOnly.controller.getCellDisplayValue({ col: 0, row })),
      before,
      "a read-only controller must not reorder cells",
    );
    assert.equal(readOnly.controller.sortState, null);
    assert.equal(readOnly.controller.canUndo, false);
  } finally {
    readOnly.app.unmount();
  }
});

test("XLSX writable controllers still sort tables in both directions", async () => {
  const editable = await mountController("sales-table.xlsx");

  try {
    assert.deepEqual(editable.controller.tables.map((table) => table.name), ["SalesRecords"]);
    const firstRowBefore = editable.controller.getCellDisplayValue({ col: 0, row: 1 });

    editable.controller.sortTable("SalesRecords", 0, "descending");
    await vue.nextTick();

    assert.deepEqual(editable.controller.sortState, {
      columnIndex: 0,
      direction: "descending",
      tableName: "SalesRecords",
    });
    assert.notEqual(editable.controller.getCellDisplayValue({ col: 0, row: 1 }), firstRowBefore);
    assert.equal(editable.controller.canUndo, true);

    editable.controller.undo();
    await vue.nextTick();
    assert.equal(editable.controller.getCellDisplayValue({ col: 0, row: 1 }), firstRowBefore);
  } finally {
    editable.app.unmount();
  }
});
