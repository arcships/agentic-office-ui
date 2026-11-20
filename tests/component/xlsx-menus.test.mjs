// Issue #160 — XLSX cell/header/AutoFilter/chart/drawing menus on the shared
// `OfficeDropdownMenu` (Reka).
//
// The menus are real primitive-layer menus now, so they need a browser-like DOM:
// this file boots happy-dom, mounts the components and drives them through real
// pointer/keyboard events. It replaces the fake-renderer menu assertions that
// could only work with the previous hand-rolled markup.

import assert from "node:assert/strict";
import test, { after } from "node:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register({ url: "http://localhost/" });

const { createRequire } = await import("node:module");
const { pathToFileURL } = await import("node:url");

const requireFromDemo = createRequire(
  new URL("../../apps/demo/package.json", import.meta.url),
);
const vue = await import(pathToFileURL(requireFromDemo.resolve("vue")).href);
const xlsx = await import(pathToFileURL(requireFromDemo.resolve("@arcships/vue-xlsx")).href);

const { createApp, h, nextTick } = vue;
const {
  XlsxAutoFilterMenu,
  XlsxContextMenu,
  XlsxHeaderMenu,
} = xlsx;

after(() => {
  GlobalRegistrator.unregister();
});

const SHEET = {
  colStyleIds: [],
  colWidths: [80, 80, 80, 80],
  defaultColWidthPx: 80,
  defaultRowHeightPx: 24,
  maxUsedCol: 3,
  maxUsedRow: 5,
  name: "Sheet1",
  rightToLeft: false,
  rowHeights: [24, 24, 24, 24, 24, 24],
  rowStyleIds: [],
  styleById: {},
  visibleCols: [0, 1, 2, 3],
  visibleRows: [0, 1, 2, 3, 4, 5],
};

function mountRoot(render) {
  const host = document.createElement("div");
  document.body.append(host);
  const app = createApp({ render });
  app.mount(host);
  const cleanup = () => {
    app.unmount();
    host.remove();
    document.body.innerHTML = "";
  };
  return { app, host, cleanup };
}

async function flush(ms = 60) {
  await nextTick();
  await new Promise((resolve) => setTimeout(resolve, ms));
  await nextTick();
}

function byTestId(testId) {
  return document.querySelector(`[data-testid="${testId}"]`);
}

function byTestIds(testId) {
  return [...document.querySelectorAll(`[data-testid="${testId}"]`)];
}

function pointer(el, type = "pointerdown") {
  el.dispatchEvent(new window.MouseEvent(type, { bubbles: true, cancelable: true }));
}

/** Reka menu items select on click (top level and inside submenus). */
async function chooseItem(testId) {
  const item = byTestId(testId);
  assert.ok(item, `${testId} is rendered`);
  item.click();
  await flush();
  return item;
}

function press(key, target = document.activeElement ?? document.body) {
  target.dispatchEvent(new window.KeyboardEvent("keydown", {
    bubbles: true,
    cancelable: true,
    key,
  }));
}

/** A zero-size element standing in for the grid canvas. */
function makeGridTarget() {
  const target = document.createElement("div");
  document.body.append(target);
  target.getBoundingClientRect = () => ({
    bottom: 600,
    height: 600,
    left: 0,
    right: 800,
    top: 0,
    width: 800,
    x: 0,
    y: 0,
    toJSON() {},
  });
  return target;
}

function contextMenuController(overrides = {}) {
  const calls = [];
  return {
    calls,
    controller: {
      activeCell: { row: 0, col: 0 },
      activeCellAddress: "A1",
      activeSheet: SHEET,
      cutSelection: () => {
        calls.push(["cut"]);
        return true;
      },
      clearSelectedCells: () => calls.push(["clear"]),
      deleteColumns: (col) => calls.push(["deleteColumns", col]),
      deleteRows: (row) => calls.push(["deleteRows", row]),
      editCell: () => calls.push(["editCell"]),
      freezePanes: null,
      getCellComment: () => null,
      getCellHyperlink: () => null,
      getTableAtCell: () => null,
      insertColumns: (col) => calls.push(["insertColumns", col]),
      insertRows: (row) => calls.push(["insertRows", row]),
      isEditingLocked: false,
      mergeAcrossSelection: () => calls.push(["mergeAcross"]),
      mergeAndCenterSelection: () => calls.push(["mergeCenter"]),
      mergeSelection: () => calls.push(["merge"]),
      pasteFromClipboard: async () => true,
      pasteSpecialFromClipboard: async (options) => {
        calls.push(["pasteSpecial", options]);
        return true;
      },
      readOnly: false,
      redo: () => calls.push(["redo"]),
      revision: 0,
      selectAll: () => calls.push(["selectAll"]),
      selectCell: (cell) => calls.push(["selectCell", cell.row, cell.col]),
      selectColumn: (col) => calls.push(["selectColumn", col]),
      selectRow: (row) => calls.push(["selectRow", row]),
      selection: { end: { row: 1, col: 1 }, start: { row: 0, col: 0 } },
      sortTable: (name, columnIndex, direction) => calls.push(["sortTable", name, columnIndex, direction]),
      tables: [],
      toggleFreezePanesAtActiveCell: () => calls.push(["freeze"]),
      undo: () => calls.push(["undo"]),
      unmergeSelection: () => calls.push(["unmerge"]),
      zoomScale: 100,
      ...overrides,
    },
  };
}

async function mountCellMenu(controller, target = makeGridTarget()) {
  let api;
  const mounted = mountRoot(() => h(XlsxContextMenu, {
    controller,
    ref: (value) => { api = value; },
    targetElement: target,
  }));
  await flush();
  return { ...mounted, api: () => api, target };
}

test("the cell menu renders the shared menu surface with every command", async () => {
  const { controller } = contextMenuController();
  const mounted = await mountCellMenu(controller);
  try {
    assert.equal(mounted.api().openAt(100, 60), true, "openAt hits a cell");
    await flush(120);

    const menu = byTestId("xlsx-context-menu");
    assert.ok(menu, "the menu content carries the public test id");
    assert.equal(menu.getAttribute("role"), "menu", "Reka renders a real menu");
    assert.equal(menu.getAttribute("data-menu-kind"), "cell");

    for (const testId of [
      "xlsx-context-cut",
      "xlsx-context-copy",
      "xlsx-context-paste",
      "xlsx-context-paste-special",
      "xlsx-context-edit-cell",
      "xlsx-context-clear",
      "xlsx-context-number-format",
      "xlsx-context-names",
      "xlsx-context-insert-row-above",
      "xlsx-context-insert-row-below",
      "xlsx-context-delete-row",
      "xlsx-context-insert-column-left",
      "xlsx-context-insert-column-right",
      "xlsx-context-delete-column",
      "xlsx-context-merge-center",
      "xlsx-context-merge-across",
      "xlsx-context-merge",
      "xlsx-context-unmerge",
      "xlsx-context-select-row",
      "xlsx-context-select-column",
      "xlsx-context-select-all",
      "xlsx-context-comment",
      "xlsx-context-hyperlink",
      "xlsx-context-chart",
      "xlsx-context-conditional",
      "xlsx-context-validation",
      "xlsx-context-table",
      "xlsx-context-freeze",
      "xlsx-context-undo",
      "xlsx-context-redo",
    ]) {
      assert.equal(byTestIds(testId).length, 1, `${testId} is rendered once`);
    }
    assert.match(menu.textContent, /Ctrl\+X/, "cut keeps its shortcut hint");
    assert.match(menu.textContent, /Ctrl\+V/, "paste keeps its shortcut hint");
    assert.equal(
      byTestId("xlsx-context-select-row").getAttribute("role"),
      "menuitemcheckbox",
      "selection state is a checkable menu row",
    );
  } finally {
    mounted.cleanup();
  }
});

test("cell menu commands reach the controller and close the menu", async () => {
  const { calls, controller } = contextMenuController();
  const mounted = await mountCellMenu(controller);
  try {
    mounted.api().openAt(100, 60);
    await flush(120);
    await chooseItem("xlsx-context-cut");
    assert.deepEqual(calls.at(-1), ["cut"], "cut reaches the controller");
    assert.equal(byTestId("xlsx-context-menu"), null, "the menu closes after a command");

    mounted.api().openAt(100, 60);
    await flush(120);
    assert.equal(byTestId("xlsx-context-paste-all"), null, "the submenu starts closed");
    const pasteSpecial = byTestId("xlsx-context-paste-special");
    pointer(pasteSpecial, "pointerenter");
    pasteSpecial.click();
    await flush(200);
    assert.ok(byTestId("xlsx-context-paste-values"), "the paste-special submenu opens");
    await chooseItem("xlsx-context-paste-values");
    assert.deepEqual(calls.at(-1), ["pasteSpecial", { mode: "values" }]);
    assert.equal(byTestId("xlsx-context-menu"), null);
  } finally {
    mounted.cleanup();
  }
});

test("structural, merge and selection commands stay wired", async () => {
  const { calls, controller } = contextMenuController();
  const mounted = await mountCellMenu(controller);
  try {
    for (const [testId, expected] of [
      ["xlsx-context-insert-row-above", ["insertRows", 0]],
      ["xlsx-context-insert-row-below", ["insertRows", 1]],
      ["xlsx-context-delete-row", ["deleteRows", 0]],
      ["xlsx-context-insert-column-left", ["insertColumns", 0]],
      ["xlsx-context-insert-column-right", ["insertColumns", 1]],
      ["xlsx-context-delete-column", ["deleteColumns", 0]],
      ["xlsx-context-merge-center", ["mergeCenter"]],
      ["xlsx-context-merge-across", ["mergeAcross"]],
      ["xlsx-context-merge", ["merge"]],
      ["xlsx-context-unmerge", ["unmerge"]],
      ["xlsx-context-select-row", ["selectRow", 0]],
      ["xlsx-context-select-column", ["selectColumn", 0]],
      ["xlsx-context-select-all", ["selectAll"]],
    ]) {
      mounted.api().openAt(100, 60);
      await flush(120);
      await chooseItem(testId);
      assert.deepEqual(calls.at(-1), expected, `${testId} reaches the controller`);
      assert.equal(byTestId("xlsx-context-menu"), null, `${testId} closes the menu`);
    }
  } finally {
    mounted.cleanup();
  }
});

test("a read-only workbook disables editing commands and keeps copy", async () => {
  const { controller } = contextMenuController({ isEditingLocked: true, readOnly: true });
  const mounted = await mountCellMenu(controller);
  try {
    mounted.api().openAt(100, 60);
    await flush(120);
    assert.equal(byTestId("xlsx-context-copy").getAttribute("data-disabled"), null);
    for (const testId of [
      "xlsx-context-cut",
      "xlsx-context-chart",
      "xlsx-context-conditional",
      "xlsx-context-validation",
    ]) {
      const item = byTestId(testId);
      assert.ok(item, `${testId} stays visible`);
      assert.equal(item.getAttribute("data-disabled"), "", `${testId} is disabled`);
    }
    assert.equal(byTestId("xlsx-context-paste"), null, "paste is not offered");
    assert.equal(byTestId("xlsx-context-insert-row-above"), null, "row edits are not offered");
  } finally {
    mounted.cleanup();
  }
});

test("menus take focus, walk with arrows and return focus to the grid", async () => {
  const { controller } = contextMenuController();
  const mounted = await mountCellMenu(controller);
  try {
    mounted.api().openAt(100, 60);
    await flush(150);
    press("ArrowDown");
    await flush(80);
    assert.equal(
      document.activeElement?.getAttribute("data-testid"),
      "xlsx-context-cut",
      "ArrowDown focuses the first item",
    );
    press("ArrowDown");
    await flush(80);
    assert.equal(
      document.activeElement?.getAttribute("data-testid"),
      "xlsx-context-copy",
      "ArrowDown moves to the next item",
    );
    press("Escape");
    await flush(150);
    assert.equal(byTestId("xlsx-context-menu"), null, "Escape closes the menu");
    assert.equal(document.activeElement, mounted.target, "focus returns to the grid element");
  } finally {
    mounted.cleanup();
  }
});

test("chart and drawing menus replace the cell actions", async () => {
  // Full chart shape: `chartTypeUnavailable` inspects series/anchor/editable.
  const chart = {
    anchor: {
      kind: "one-cell",
      from: { col: 0, colOffsetEmu: 0, row: 0, rowOffsetEmu: 0 },
      sizeEmu: { cx: 914400, cy: 914400 },
    },
    chartPath: "xl/charts/chart1.xml",
    chartType: "ColumnClustered",
    editable: true,
    id: "chart-1",
    series: [
      { categories: ["A"], dataPoints: [], id: "series-1", name: "Series 1", values: [1] },
      { categories: ["A"], dataPoints: [], id: "series-2", name: "Series 2", values: [2] },
    ],
    title: "Revenue",
    workbookSheetIndex: 0,
    zIndex: 0,
  };
  const calls = [];
  const { controller } = contextMenuController({
    changeChartType: (id, type) => {
      calls.push(["changeChartType", id, type]);
      return true;
    },
    deleteSelectedDrawing: () => calls.push(["deleteDrawing"]),
    duplicateDrawing: () => calls.push(["duplicateDrawing"]),
    getChartById: (id) => (id === chart.id ? chart : null),
    getShapeById: () => ({ id: "shape-1" }),
    selectChart: (id) => calls.push(["selectChart", id]),
    selectShape: (id) => calls.push(["selectShape", id]),
  });
  const mounted = await mountCellMenu(controller);
  try {
    assert.equal(mounted.api().openForChart(chart.id, 40, 60), true);
    await flush(150);
    assert.equal(byTestId("xlsx-context-menu").getAttribute("data-menu-kind"), "cell");
    assert.equal(byTestId("xlsx-context-comment"), null, "cell actions are replaced");
    assert.deepEqual(calls.at(-1), ["selectChart", chart.id]);

    const trigger = byTestId("xlsx-context-chart-type");
    assert.ok(trigger, "Change Chart Type is offered");
    pointer(trigger, "pointerenter");
    trigger.click();
    await flush(250);
    assert.ok(byTestId("xlsx-chart-type-menu"), "the family submenu opens");
    const family = byTestId("xlsx-chart-family-Pie");
    assert.ok(family, "the Pie family is listed");
    pointer(family, "pointerenter");
    family.click();
    await flush(250);
    assert.ok(byTestId("xlsx-chart-type-Pie3D"), "the type submenu opens");
    await chooseItem("xlsx-chart-type-Pie3D");
    assert.deepEqual(calls.at(-1), ["changeChartType", chart.id, "Pie3D"]);
    assert.equal(byTestId("xlsx-context-menu"), null);

    mounted.api().openForDrawing("shape", "shape-1", 20, 20);
    await flush(150);
    assert.deepEqual(calls.at(-1), ["selectShape", "shape-1"]);
    assert.ok(byTestId("xlsx-context-drawing-duplicate"), "drawing actions are offered");
    await chooseItem("xlsx-context-drawing-duplicate");
    assert.deepEqual(calls.at(-1), ["duplicateDrawing"]);
    assert.equal(byTestId("xlsx-context-menu"), null);
  } finally {
    mounted.cleanup();
  }
});

test("the column header menu routes its commands and closes", async () => {
  const calls = [];
  const controller = contextMenuController().controller;
  controller.insertColumns = (col) => calls.push(["insertColumns", col]);
  controller.deleteColumns = (col) => calls.push(["deleteColumns", col]);
  controller.selectColumn = (col) => calls.push(["selectColumn", col]);
  const closes = [];
  const mounted = mountRoot(() => h(XlsxHeaderMenu, {
    controller,
    index: 1,
    kind: "column",
    x: 100,
    y: 100,
    onClose: () => closes.push("close"),
  }));
  await flush(120);
  try {
    assert.ok(byTestId("xlsx-column-menu-list"), "the column menu renders");
    for (const testId of [
      "xlsx-column-menu-insert-left",
      "xlsx-column-menu-insert-right",
      "xlsx-column-menu-delete",
      "xlsx-column-menu-select",
      "xlsx-column-menu-clear",
    ]) {
      assert.ok(byTestId(testId), `${testId} is rendered`);
    }
    await chooseItem("xlsx-column-menu-insert-right");
    assert.deepEqual(calls.at(-1), ["insertColumns", 2]);
    assert.equal(closes.length, 1, "the menu reports its close");
  } finally {
    mounted.cleanup();
  }
});

test("the row header menu refuses commands in a read-only workbook", async () => {
  const calls = [];
  const controller = contextMenuController().controller;
  controller.readOnly = true;
  controller.insertRows = (row) => calls.push(["insertRows", row]);
  const closes = [];
  const mounted = mountRoot(() => h(XlsxHeaderMenu, {
    controller,
    index: 2,
    kind: "row",
    x: 40,
    y: 100,
    onClose: () => closes.push("close"),
  }));
  await flush(120);
  try {
    assert.ok(byTestId("xlsx-row-menu-list"), "the row menu renders");
    await chooseItem("xlsx-row-menu-insert-above");
    assert.deepEqual(calls, [], "no controller command runs");
    assert.ok(closes.length >= 1, "the menu still closes");
  } finally {
    mounted.cleanup();
  }
});

function autoFilterController(overrides = {}) {
  const calls = [];
  return {
    calls,
    controller: {
      applyAutoFilterValues: (reference, columnIndex, values, options) =>
        calls.push(["apply", reference, columnIndex, values, options]),
      autoFilterStates: [],
      clearAutoFilter: (reference) => calls.push(["clear", reference]),
      getAutoFilterColumnValues: () => [
        { blank: false, value: "Alpha" },
        { blank: false, value: "Beta" },
        { blank: true, value: "" },
      ],
      isEditingLocked: false,
      readOnly: false,
      revision: 0,
      sortAutoFilter: (reference, columnIndex, direction) =>
        calls.push(["sort", reference, columnIndex, direction]),
      ...overrides,
    },
  };
}

function mountAutoFilter(controller) {
  const closes = [];
  const mounted = mountRoot(() => h(XlsxAutoFilterMenu, {
    controller,
    range: { reference: "A1:B4" },
    columnIndex: 1,
    columnName: "Region",
    x: 120,
    y: 90,
    onClose: () => closes.push("close"),
  }));
  return { ...mounted, closes };
}

test("the AutoFilter panel filters values, keeps the menu open and applies", async () => {
  const { calls, controller } = autoFilterController();
  const mounted = mountAutoFilter(controller);
  await flush(120);
  try {
    const menu = byTestId("xlsx-autofilter-menu");
    assert.ok(menu, "the panel renders on the shared menu surface");
    assert.equal(menu.getAttribute("role"), "menu");
    assert.ok(byTestId("xlsx-autofilter-sort-ascending"), "sort is offered");
    assert.ok(byTestId("xlsx-autofilter-select-all"), "select all is offered");
    assert.ok(byTestId("xlsx-autofilter-value-0"), "a value row is offered");
    assert.equal(byTestId("xlsx-autofilter-value-0").getAttribute("aria-checked"), "true");

    await chooseItem("xlsx-autofilter-value-0");
    assert.ok(byTestId("xlsx-autofilter-menu"), "toggling a value keeps the panel open");
    assert.equal(byTestId("xlsx-autofilter-value-0").getAttribute("aria-checked"), "false");

    await chooseItem("xlsx-autofilter-apply");
    assert.equal(calls.at(-1)[0], "apply");
    assert.deepEqual(calls.at(-1)[1], "A1:B4");
    assert.equal(calls.at(-1)[2], 1);
    assert.deepEqual(calls.at(-1)[3], ["Beta"], "only the still-checked values apply");
    assert.deepEqual(mounted.closes, ["close"], "Apply closes the panel");
  } finally {
    mounted.cleanup();
  }
});

test("the AutoFilter panel sorts, clears, closes on Escape and locks in read-only", async () => {
  const { calls, controller } = autoFilterController();
  const mounted = mountAutoFilter(controller);
  await flush(120);
  try {
    await chooseItem("xlsx-autofilter-sort-descending");
    assert.deepEqual(calls.at(-1), ["sort", "A1:B4", 1, "descending"]);
    assert.deepEqual(mounted.closes, ["close"]);
  } finally {
    mounted.cleanup();
  }

  const locked = autoFilterController({ readOnly: true });
  const readOnly = mountAutoFilter(locked.controller);
  await flush(120);
  try {
    const sort = byTestId("xlsx-autofilter-sort-ascending");
    assert.equal(sort.getAttribute("data-disabled"), "", "sorting is disabled in read-only");
    press("Escape", byTestId("xlsx-autofilter-menu"));
    await flush(150);
    assert.equal(byTestId("xlsx-autofilter-menu"), null, "Escape closes the panel");
    assert.ok(readOnly.closes.length >= 1);
  } finally {
    readOnly.cleanup();
  }
});

test("the menus are real shared menu primitives, not local markup", async () => {
  const { readFileSync } = await import("node:fs");
  const read = (name) => readFileSync(
    new URL(`../../packages/vue-xlsx/src/components/${name}.vue`, import.meta.url),
    "utf8",
  );
  for (const name of ["XlsxContextMenu", "XlsxHeaderMenu", "XlsxAutoFilterMenu"]) {
    const source = read(name);
    assert.match(source, /OfficeDropdownMenu/, `${name} renders OfficeDropdownMenu`);
    assert.match(source, /OfficeMenuEntry/, `${name} builds the shared item model`);
    assert.equal(
      /class="office-menu__item/.test(source),
      false,
      `${name} no longer hand-rolls menu rows`,
    );
    assert.equal(
      /<Teleport/.test(source),
      false,
      `${name} no longer ships a custom portal/backdrop`,
    );
  }
});
