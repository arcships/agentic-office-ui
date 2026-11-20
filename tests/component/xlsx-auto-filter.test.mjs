import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import test from "node:test";
import { pathToFileURL } from "node:url";

import {
  findByTestId,
  importFromDemo,
  mount,
  teleportTarget,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

// Real controller + real grid; only the browser plumbing the renderer cannot
// provide is stubbed. Canvas painting is asserted through the grid's
// data-* contract because the host nodes have no 2D context.
globalThis.ResizeObserver ??= class ResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
};
globalThis.Document ??= class Document {};
globalThis.ShadowRoot ??= class ShadowRoot {};
globalThis.window ??= {
  addEventListener() {},
  cancelAnimationFrame() {},
  clearTimeout: (id) => clearTimeout(id),
  devicePixelRatio: 1,
  removeEventListener() {},
  requestAnimationFrame: () => 1,
  setTimeout: (fn, ms) => setTimeout(fn, ms),
};
if (typeof Object.prototype.getContext !== "function") {
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value: () => null,
  });
}

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const { unzipSync, zipSync, strFromU8, strToU8 } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);
const xmldom = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= xmldom.DOMParser;
globalThis.Node ??= xmldom.Node;
globalThis.XMLSerializer ??= xmldom.XMLSerializer;

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { useXlsxViewerController, XlsxGrid, XlsxViewer } = await importFromDemo(
  "@arcships/vue-xlsx",
);

function configureWasm() {
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}
configureWasm();

function toBuffer(bytes) {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function sampleBytes(name) {
  return new Uint8Array(
    readFileSync(new URL(`../../apps/demo/public/samples/${name}`, import.meta.url)),
  );
}

function patchSheetParts(bytes, replace) {
  const archive = unzipSync(bytes);
  for (const part of Object.keys(archive)) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(part)) continue;
    archive[part] = strToU8(replace(strFromU8(archive[part])));
  }
  return zipSync(archive);
}

const WORKSHEET_FILTER =
  '<autoFilter ref="A1:D5"><filterColumn colId="0"><filters><filter val="ARPU"/></filters></filterColumn></autoFilter>';

function worksheetFilterWorkbook({ frozen = false, merged = false } = {}) {
  return patchSheetParts(sampleBytes("financial-model.xlsx"), (xml) => {
    let next = xml;
    if (frozen) {
      next = next.replace(
        /<sheetView([^>]*)>/,
        '<sheetView$1><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
      );
    }
    const extra = `${WORKSHEET_FILTER}${merged ? '<mergeCells count="1"><mergeCell ref="C4:D4"/></mergeCells>' : ""}`;
    return next.replace("</sheetData>", `</sheetData>${extra}`);
  });
}

function worksheetFilterBytes(options) {
  const bytes = worksheetFilterWorkbook(options);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function mountController(file, options = {}) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file,
        fileName: options.fileName ?? "auto-filter.xlsx",
        readOnly: options.readOnly === true,
        skipXmlParsing: false,
        useWorker: false,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "XLSX AutoFilter controller ready",
  );
  assert.ifError(controller.error);
  return { ...mounted, controller };
}

async function mountGrid(controller) {
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(XlsxGrid, { controller });
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, grid: findByTestId(mounted.root, "xlsx-grid") };
}

function autoFilterTriggers(root) {
  return walk(root).filter((node) => node.props?.["data-testid"] === "xlsx-autofilter-trigger");
}

function click(node, init = {}) {
  assert.equal(typeof node?.props?.onClick, "function", "expected a clickable node");
  node.props.onClick({
    preventDefault() {},
    stopPropagation() {},
    currentTarget: node,
    ...init,
  });
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
    const objectUrl = `blob:test-${nextObjectUrl += 1}`;
    blobs.set(objectUrl, blob);
    return objectUrl;
  };
  URL.revokeObjectURL = () => {};

  t.after(() => {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    if (originalCreateObjectUrl === undefined) delete URL.createObjectURL;
    else URL.createObjectURL = originalCreateObjectUrl;
    if (originalRevokeObjectUrl === undefined) delete URL.revokeObjectURL;
    else URL.revokeObjectURL = originalRevokeObjectUrl;
  });

  return downloads;
}

async function exportedBytes(controller, downloads) {
  const before = downloads.length;
  controller.exportXlsx();
  await vue.nextTick();
  assert.equal(downloads.length, before + 1, "export must produce one download");
  return new Uint8Array(await downloads.at(-1).blob.arrayBuffer());
}

test("parses worksheet and table AutoFilters and seeds the file's filter", async () => {
  const { app, controller } = await mountController(worksheetFilterBytes());
  try {
    controller.setActiveSheetIndex(0);
    assert.equal(controller.autoFilterRanges.length, 1);
    const range = controller.autoFilterRanges[0];
    assert.equal(range.reference, "A1:D5");
    assert.equal(range.source, "worksheet");
    assert.deepEqual(
      range.columns.map((column) => [column.columnIndex, column.values, column.includeBlank]),
      [[0, ["ARPU"], false]],
    );

    // The criteria stored in the package hide every row except "ARPU".
    assert.deepEqual(
      controller.autoFilterStates.map((state) => state.columns),
      [[{ columnIndex: 0, includeBlank: false, values: ["ARPU"] }]],
    );
    assert.deepEqual(controller.hiddenAutoFilterRows, [1, 3, 4]);
    assert.deepEqual(
      controller.getAutoFilterColumnValues(range, 0).map((option) => option.value),
      ["ARPU", "Churn", "Gross Margin", "Starting Customers"],
    );

    assert.equal(controller.clearAutoFilter(), true);
    assert.deepEqual(controller.hiddenAutoFilterRows, [], "clearing restores the rows");
    assert.deepEqual(controller.autoFilterStates, []);

    assert.equal(controller.applyAutoFilterValues(range.reference, 0, ["Churn"]), true);
    assert.deepEqual(controller.hiddenAutoFilterRows, [1, 2, 3]);
    assert.equal(
      controller.applyAutoFilterValues("Z9:Z10", 0, ["x"]),
      false,
      "an unknown range is refused",
    );
    assert.equal(
      controller.applyAutoFilterValues(range.reference, 9, ["x"]),
      false,
      "a column outside the range is refused",
    );
  } finally {
    app.unmount();
  }
});

test("sorting a worksheet AutoFilter range is undoable and records its range", async () => {
  const { app, controller } = await mountController(worksheetFilterBytes());
  try {
    controller.setActiveSheetIndex(0);
    controller.clearAutoFilter();
    assert.equal(controller.getCellDisplayValue({ row: 1, col: 0 }), "Starting Customers");
    assert.equal(controller.sortAutoFilter("A1:D5", 0, "ascending"), true);
    assert.equal(controller.getCellDisplayValue({ row: 1, col: 0 }), "ARPU");
    assert.deepEqual(controller.sortState, {
      columnIndex: 0,
      direction: "ascending",
      rangeReference: "A1:D5",
      tableName: "A1:D5",
    });

    controller.undo();
    await waitFor(
      () => controller.getCellDisplayValue({ row: 1, col: 0 }) === "Starting Customers",
      "one undo restores the worksheet rows",
    );
  } finally {
    app.unmount();
  }
});

test("filtering leaves merges and frozen panes untouched and sorting refuses to break them", async () => {
  const { app, controller } = await mountController(
    worksheetFilterBytes({ frozen: true, merged: true }),
  );
  const gridMount = await mountGrid(controller);
  try {
    controller.setActiveSheetIndex(0);
    await vue.nextTick();
    const sheet = controller.activeSheet;
    const mergedRegions = JSON.parse(JSON.stringify(sheet.mergedRegions));
    const freezePanes = JSON.parse(JSON.stringify(sheet.freezePanes));
    assert.equal(mergedRegions.length, 1);
    assert.deepEqual(freezePanes, { col: 0, row: 1 });
    assert.equal(gridMount.grid.props["data-merged-region-count"], 1);

    controller.clearAutoFilter();
    assert.equal(controller.applyAutoFilterValues("A1:D5", 0, ["ARPU"]), true);
    assert.deepEqual(controller.hiddenAutoFilterRows, [1, 3, 4]);
    assert.equal(
      controller.sortAutoFilter("A1:D5", 1, "ascending"),
      false,
      "a merged data body cannot be sorted",
    );
    assert.equal(controller.sortState, null);
    assert.deepEqual(JSON.parse(JSON.stringify(controller.activeSheet.mergedRegions)), mergedRegions);
    assert.deepEqual(JSON.parse(JSON.stringify(controller.activeSheet.freezePanes)), freezePanes);
    await vue.nextTick();
    assert.equal(gridMount.grid.props["data-merged-region-count"], 1);
    assert.equal(gridMount.grid.props["data-auto-filter-trigger-count"], 4);

    controller.clearAutoFilter();
    assert.deepEqual(JSON.parse(JSON.stringify(controller.activeSheet.mergedRegions)), mergedRegions);
    assert.deepEqual(JSON.parse(JSON.stringify(controller.activeSheet.freezePanes)), freezePanes);
  } finally {
    gridMount.app.unmount();
    app.unmount();
  }
});

test("export -> reload keeps applied and cleared worksheet filters", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController(worksheetFilterBytes());
  let exported;
  try {
    controller.setActiveSheetIndex(0);
    controller.clearAutoFilter();
    exported = await exportedBytes(controller, downloads);
  } finally {
    app.unmount();
  }

  const cleared = await mountController(toBuffer(exported));
  try {
    cleared.controller.setActiveSheetIndex(0);
    assert.deepEqual(cleared.controller.hiddenAutoFilterRows, [], "a cleared filter stays cleared");
    assert.deepEqual(cleared.controller.autoFilterRanges[0].columns, []);
  } finally {
    cleared.app.unmount();
  }

  const appliedMount = await mountController(worksheetFilterBytes());
  let applied;
  try {
    appliedMount.controller.setActiveSheetIndex(0);
    appliedMount.controller.clearAutoFilter();
    appliedMount.controller.applyAutoFilterValues("A1:D5", 0, ["Churn", "ARPU"]);
    applied = await exportedBytes(appliedMount.controller, downloads);
  } finally {
    appliedMount.app.unmount();
  }

  const reloaded = await mountController(toBuffer(applied));
  try {
    reloaded.controller.setActiveSheetIndex(0);
    assert.deepEqual(reloaded.controller.hiddenAutoFilterRows, [1, 3]);
    assert.deepEqual(reloaded.controller.autoFilterStates[0].columns, [
      { columnIndex: 0, includeBlank: false, values: ["Churn", "ARPU"] },
    ]);
  } finally {
    reloaded.app.unmount();
  }
});

test("export -> reload keeps a table filter inside the table part", async (t) => {
  const downloads = installDownloadCapture(t);
  const bytes = sampleBytes("sales-table.xlsx");
  const { app, controller } = await mountController(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    { fileName: "sales-table.xlsx" },
  );
  let exported;
  try {
    assert.equal(controller.applyAutoFilterValues("A1:H181", 0, ["North"]), true);
    assert.ok(controller.hiddenAutoFilterRows.length > 0);
    exported = await exportedBytes(controller, downloads);
  } finally {
    app.unmount();
  }

  const reloaded = await mountController(toBuffer(exported), { fileName: "sales-filtered.xlsx" });
  try {
    const range = reloaded.controller.autoFilterRanges.find((entry) => entry.source === "table");
    assert.equal(range.reference, "A1:H181");
    assert.deepEqual(range.columns[0].values, ["North"]);
    assert.ok(reloaded.controller.hiddenAutoFilterRows.length > 0);
    for (const row of reloaded.controller.activeSheet.visibleRows) {
      if (row === 0 || reloaded.controller.hiddenAutoFilterRows.includes(row)) continue;
      assert.equal(
        reloaded.controller.getCellDisplayValue({ row, col: 0 }),
        "North",
        "only the matching region survives the reload",
      );
    }
  } finally {
    reloaded.app.unmount();
  }
});

test("the table-header-menu slot replaces the trigger and receives filter context", async () => {
  const { app, controller } = await mountController(worksheetFilterBytes());
  let slotProps = null;
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(
        XlsxViewer,
        {
          controller,
          showDefaultToolbar: false,
          showFormulaBar: false,
          showRibbon: false,
          showSheetTabs: false,
        },
        {
          "table-header-menu": (props) => {
            slotProps = props;
            return vue.h(
              "button",
              {
                "data-testid": "custom-table-header-menu",
                onClick: props.triggerProps.onClick,
              },
              props.triggerIcon,
            );
          },
        },
      );
    },
  });
  const mounted = await mount(Harness);
  try {
    controller.setActiveSheetIndex(0);
    await vue.nextTick();
    assert.equal(autoFilterTriggers(mounted.root).length, 0, "the built-in trigger is replaced");
    const custom = walk(mounted.root).filter(
      (node) => node.props?.["data-testid"] === "custom-table-header-menu",
    );
    assert.equal(custom.length, 4);
    assert.ok(slotProps);
    assert.equal(slotProps.autoFilter.reference, "A1:D5");
    assert.equal(slotProps.table, null);
    assert.equal(slotProps.column.index, 3);
    assert.equal(slotProps.hasFilter, true);
    assert.equal(slotProps.hasValueFilter, false, "only column A carries the file's filter");
    assert.equal(slotProps.direction, null);
    assert.deepEqual(
      slotProps.getValueOptions().map((option) => option.value),
      ["0.06", "0.76", "255", "492"],
    );

    // Opening the Reka panel itself is covered by the happy-dom suite
    // (xlsx-menus.test.mjs); this test owns the slot/host contract only.
    assert.equal(typeof slotProps.triggerProps.onClick, "function");
  } finally {
    mounted.app.unmount();
    app.unmount();
  }
});
