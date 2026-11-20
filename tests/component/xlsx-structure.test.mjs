import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
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

const { applyWorkbookDefinedNames, setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);
const {
  XlsxAuthoringDialog,
  XlsxContextMenu,
  XlsxRibbon,
  useXlsxViewerController,
} = await importFromDemo("@arcships/vue-xlsx");

function configureWasm() {
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}
configureWasm();

function sampleBuffer(name = "charts-images.xlsx") {
  const bytes = readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function mountController(options = {}) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({ useWorker: false, ...options });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "XLSX controller ready state",
  );
  assert.ifError(controller.error);
  assert.equal(controller.sourceState, "ready");
  return { ...mounted, controller };
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

async function exportedBytes(downloads, controller) {
  controller.exportXlsx();
  await waitFor(() => downloads.length > 0, "XLSX export download");
  const { blob } = downloads.pop();
  assert.ok(blob, "export produced a blob");
  return new Uint8Array(await blob.arrayBuffer());
}

function cell(controller, row, col) {
  return controller.getCellDisplayValue({ row, col });
}

function formula(controller, row, col) {
  return controller.getCellFormula({ row, col });
}

test("insertRows shifts values, formulas, merges, notes, links, ranges and pans", async (t) => {
  const { app, controller } = await mountController({ file: sampleBuffer() });
  try {
    // Scratch block below the sample's own content so the assertions are not
    // coupled to the fixture's values.
    controller.setCellValue({ row: 10, col: 0 }, "alpha");
    controller.setCellValue({ row: 11, col: 0 }, "beta");
    controller.setCellValue({ row: 12, col: 0 }, "gamma");
    controller.setCellFormula({ row: 11, col: 1 }, "=A12");
    controller.setRangeStyle(
      { start: { row: 11, col: 0 }, end: { row: 11, col: 0 } },
      { font: { bold: true }, numberFormat: { formatType: "custom", formatString: "0.00" } },
    );
    controller.selectRange({ start: { row: 11, col: 0 }, end: { row: 12, col: 1 } });
    controller.mergeSelection();
    controller.selectCell({ row: 11, col: 0 });
    controller.setCellHyperlink({ row: 11, col: 0 }, { target: "https://example.com/structure", tooltip: "docs" });
    controller.setCellComment({ row: 11, col: 0 }, { text: "beta note", author: "Dana" });
    controller.addConditionalFormatRule({
      colors: ["#63be7b"],
      kind: "dataBar",
      ranges: [{ start: { row: 10, col: 0 }, end: { row: 12, col: 0 } }],
    });
    controller.selectRange({ start: { row: 10, col: 0 }, end: { row: 12, col: 0 } });
    controller.setSelectionDataValidation({ listSource: "a,b,c", validationType: "list" });
    controller.setFreezePanes({ col: 0, row: 5 });
    controller.selectCell({ row: 11, col: 0 });

    assert.equal(controller.insertRows(11), true);

    assert.equal(cell(controller, 11, 0), "");
    assert.equal(cell(controller, 12, 0), "beta");
    assert.equal(cell(controller, 13, 0), "gamma");
    assert.equal(formula(controller, 12, 1), "=A13");
    assert.equal(controller.getCellHyperlink({ row: 12, col: 0 })?.target, "https://example.com/structure");
    assert.equal(controller.getCellHyperlink({ row: 11, col: 0 }), null);
    assert.equal(controller.getCellComment({ row: 12, col: 0 })?.text, "beta note");
    assert.deepEqual(
      controller.activeSheet.mergedRegions.map((range) => [range.start.row, range.start.col, range.end.row, range.end.col]),
      [[6, 0, 7, 2], [12, 0, 13, 1]],
      "the scratch merge followed the inserted row",
    );
    const rule = controller.getConditionalFormatRules().at(-1);
    assert.deepEqual(
      rule.ranges.map((range) => [range.start.row, range.end.row]),
      [[10, 13]],
      "the conditional range expanded with the insertion",
    );
    const validation = controller.getCellDataValidation({ row: 12, col: 0 });
    assert.ok(validation, "the validation moved with the range");
    assert.deepEqual(
      validation.ranges.map((range) => [range.start.row, range.end.row]),
      [[10, 13]],
    );
    assert.deepEqual(controller.freezePanes, { col: 0, row: 5 });
    assert.equal(cell(controller, 12, 0), "beta");
    assert.equal(controller.images[0]?.anchor?.from?.row, 18, "the image anchor moved with the inserted row");
    assert.equal(controller.charts[0]?.anchor?.from?.row, 1, "anchors above the insertion stay put");
  } finally {
    app.unmount();
  }
});

test("structure edits survive export/reload and undo through the snapshot", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController({ file: sampleBuffer() });
  try {
    controller.setCellValue({ row: 10, col: 0 }, "alpha");
    controller.setCellValue({ row: 11, col: 0 }, "beta");
    controller.setCellValue({ row: 12, col: 0 }, "gamma");
    controller.setCellFormula({ row: 11, col: 1 }, "=A12");
    controller.selectCell({ row: 11, col: 0 });
    controller.setCellHyperlink({ row: 11, col: 0 }, { target: "#Dashboard!A12" });
    controller.setNamedRange("Revenue", "'Dashboard'!$A$12");
    assert.equal(controller.insertRows(11), true);
    assert.equal(cell(controller, 12, 0), "beta");
    assert.equal(controller.getCellHyperlink({ row: 12, col: 0 })?.target, "#Dashboard!A13");
    assert.equal(
      controller.getNamedRanges().find((entry) => entry.name === "Revenue")?.refersTo,
      "'Dashboard'!$A$13",
      "a defined name follows the inserted row",
    );

    const bytes = await exportedBytes(downloads, controller);
    const reloaded = await mountController({ file: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) });
    try {
      assert.equal(cell(reloaded.controller, 12, 0), "beta");
      assert.equal(formula(reloaded.controller, 12, 1), "=A13");
      assert.equal(reloaded.controller.getCellHyperlink({ row: 12, col: 0 })?.target, "#Dashboard!A13");
      assert.equal(reloaded.controller.images[0]?.anchor?.from?.row, 18, "the shifted image anchor survives export");
      assert.equal(
        reloaded.controller.getNamedRanges().find((entry) => entry.name === "Revenue")?.refersTo,
        "'Dashboard'!$A$13",
      );
    } finally {
      reloaded.app.unmount();
    }

    controller.undo();
    await waitFor(() => cell(controller, 11, 0) === "beta", "undo restores the inserted row");
    assert.equal(cell(controller, 12, 0), "gamma");
    assert.equal(controller.getCellHyperlink({ row: 11, col: 0 })?.target, "#Dashboard!A12");
  } finally {
    app.unmount();
  }
});

test("deleteRows pulls values up and gives removed references #REF!", async () => {
  const { app, controller } = await mountController({ file: sampleBuffer() });
  try {
    controller.setCellValue({ row: 20, col: 0 }, "one");
    controller.setCellValue({ row: 21, col: 0 }, "two");
    controller.setCellValue({ row: 22, col: 0 }, "three");
    controller.setCellFormula({ row: 20, col: 1 }, "=A22");
    controller.setCellFormula({ row: 22, col: 2 }, "=A22");

    assert.equal(controller.deleteRows(21, 1), true);

    assert.equal(cell(controller, 20, 0), "one");
    assert.equal(cell(controller, 21, 0), "three");
    assert.equal(formula(controller, 20, 1).endsWith("#REF!"), true, "a formula that referenced the deleted row becomes #REF!");
    assert.equal(formula(controller, 21, 2).endsWith("#REF!"), true);
  } finally {
    app.unmount();
  }
});

test("insertColumns and deleteColumns shift values and formulas", async () => {
  const { app, controller } = await mountController({ file: sampleBuffer() });
  try {
    controller.setCellValue({ row: 30, col: 0 }, "left");
    controller.setCellValue({ row: 30, col: 1 }, "middle");
    controller.setCellValue({ row: 30, col: 2 }, "right");
    controller.setCellFormula({ row: 31, col: 0 }, "=B31");
    controller.setCellValue({ row: 31, col: 1 }, "7");

    assert.equal(controller.insertColumns(1, 1), true);
    assert.equal(cell(controller, 30, 0), "left");
    assert.equal(cell(controller, 30, 1), "");
    assert.equal(cell(controller, 30, 2), "middle");
    assert.equal(formula(controller, 31, 0), "=C31");
    assert.equal(cell(controller, 31, 2), "7");

    assert.equal(controller.deleteColumns(1, 1), true);
    assert.equal(cell(controller, 30, 1), "middle");
    assert.equal(cell(controller, 30, 2), "right");
    assert.equal(formula(controller, 31, 0), "=B31");
  } finally {
    app.unmount();
  }
});

test("the table range follows row insertions and deletions", async () => {
  const { app, controller } = await mountController({ file: sampleBuffer("sales-table.xlsx") });
  try {
    assert.equal(controller.tables[0]?.reference, "A1:H181");
    assert.equal(controller.insertRows(5, 2), true);
    assert.equal(controller.tables[0]?.reference, "A1:H183");
    assert.equal(controller.deleteRows(5, 2), true);
    assert.equal(controller.tables[0]?.reference, "A1:H181");
  } finally {
    app.unmount();
  }
});

test("paste special writes values, formats and transposed blocks", async () => {
  const { app, controller } = await mountController({ file: sampleBuffer() });
  try {
    controller.setCellValue({ row: 40, col: 0 }, "1");
    controller.setCellValue({ row: 40, col: 1 }, "2");
    controller.setCellValue({ row: 41, col: 1 }, "2");
    controller.setCellFormula({ row: 41, col: 0 }, "=B41");
    controller.setRangeStyle(
      { start: { row: 40, col: 0 }, end: { row: 40, col: 1 } },
      { fill: { color: { colorType: "rgb", hex: "FF0000" }, fillType: "solid" }, font: { bold: true } },
    );
    controller.selectRange({ start: { row: 40, col: 0 }, end: { row: 41, col: 1 } });
    const payload = controller.getClipboardData();
    assert.ok(payload);

    controller.selectCell({ row: 45, col: 3 });
    assert.equal(controller.pasteSpecialStructuredData(payload.structured, { mode: "values" }), true);
    assert.equal(cell(controller, 45, 3), "1");
    assert.equal(cell(controller, 45, 4), "2");
    assert.equal(formula(controller, 46, 3), "", "values-only paste drops the formula");
    assert.equal(cell(controller, 46, 3), "2");

    controller.selectCell({ row: 45, col: 6 });
    assert.equal(controller.pasteSpecialStructuredData(payload.structured, { mode: "formats" }), true);
    assert.equal(cell(controller, 45, 6), "", "formats-only paste leaves the target value alone");
    assert.equal(controller.getActiveWorksheet().getCellStyleAt(45, 6)?.font?.bold, true);
    assert.equal(controller.getActiveWorksheet().getCellStyleAt(45, 6)?.fill?.fillType, "solid");

    controller.selectCell({ row: 45, col: 8 });
    assert.equal(controller.pasteSpecialStructuredData(payload.structured, { mode: "all", transpose: true }), true);
    assert.equal(cell(controller, 45, 8), "1");
    assert.equal(cell(controller, 46, 8), "2");
    assert.equal(formula(controller, 45, 9).startsWith("="), true, "transpose moves the formula cell");

    // Cut moves values instead of copying them.
    controller.selectRange({ start: { row: 40, col: 0 }, end: { row: 40, col: 0 } });
    assert.equal(controller.cutSelection(), true);
    const cutPayload = controller.getClipboardData().structured;
    controller.selectCell({ row: 50, col: 0 });
    assert.equal(controller.pasteSpecialStructuredData(cutPayload, { mode: "all" }), true);
    assert.equal(cell(controller, 50, 0), "1");
    assert.equal(cell(controller, 40, 0), "", "cut clears the source cell");
  } finally {
    app.unmount();
  }
});

test("loaded defined names follow structure edits and survive export", async (t) => {
  const downloads = installDownloadCapture(t);
  const source = applyWorkbookDefinedNames(new Uint8Array(sampleBuffer()), [
    { name: "LoadedRange", refersTo: "'Dashboard'!$A$12:$B$12" },
  ]);
  const { app, controller } = await mountController({
    file: source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength),
  });
  try {
    controller.setCellValue({ row: 11, col: 0 }, "marker");
    assert.equal(
      controller.getNamedRanges().find((entry) => entry.name === "LoadedRange")?.refersTo,
      "'Dashboard'!$A$12:$B$12",
      "the loaded name is visible before the structure edit",
    );
    assert.equal(controller.insertRows(11), true);
    assert.equal(
      controller.getNamedRanges().find((entry) => entry.name === "LoadedRange")?.refersTo,
      "'Dashboard'!$A$13:$B$13",
      "a name loaded from workbook.xml follows the inserted row",
    );
    assert.equal(controller.deleteRows(11), true);
    assert.equal(
      controller.getNamedRanges().find((entry) => entry.name === "LoadedRange")?.refersTo,
      "'Dashboard'!$A$12:$B$12",
      "the loaded name moves back on delete",
    );
    assert.equal(controller.insertRows(11), true);

    const bytes = await exportedBytes(downloads, controller);
    const reloaded = await mountController({ file: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) });
    try {
      assert.equal(
        reloaded.controller.getNamedRanges().find((entry) => entry.name === "LoadedRange")?.refersTo,
        "'Dashboard'!$A$13:$B$13",
        "the shifted loaded name survives export/reload",
      );
    } finally {
      reloaded.app.unmount();
    }
  } finally {
    app.unmount();
  }
});

test("formulas on other sheets follow edits to the referenced sheet", async () => {
  const { app, controller } = await mountController({ file: sampleBuffer("financial-model.xlsx") });
  try {
    controller.setActiveSheetIndex(0);
    controller.setCellValue({ row: 11, col: 0 }, "twelve");
    controller.setActiveSheetIndex(1);
    controller.setCellFormula({ row: 40, col: 0 }, "='Assumptions'!A12");
    controller.setCellFormula({ row: 41, col: 0 }, "=A2");
    controller.setCellFormula({ row: 42, col: 0 }, "=Assumptions!A12");
    controller.setActiveSheetIndex(0);
    controller.setCellFormula({ row: 42, col: 0 }, "='P&L'!A1");

    assert.equal(controller.insertRows(11), true);

    controller.setActiveSheetIndex(1);
    assert.equal(formula(controller, 40, 0), "='Assumptions'!A13", "the cross-sheet reference moves");
    assert.equal(formula(controller, 41, 0), "=A2", "unqualified references on the other sheet stay put");
    assert.equal(formula(controller, 42, 0), "=Assumptions!A13", "unquoted sheet references also move");

    controller.setActiveSheetIndex(0);
    assert.equal(formula(controller, 42, 0), "='P&L'!A1", "references to another sheet stay put");

    controller.deleteRows(11);
    controller.setActiveSheetIndex(1);
    assert.equal(formula(controller, 40, 0), "='Assumptions'!A12", "the cross-sheet reference moves back");
  } finally {
    app.unmount();
  }
});

function workbookWithUnmodeledRule(bytes) {
  const archive = unzipSync(bytes);
  const sheetPath = "xl/worksheets/sheet1.xml";
  const xml = strFromU8(archive[sheetPath]);
  archive[sheetPath] = strToU8(xml.replace(
    "</worksheet>",
    '<conditionalFormatting sqref="A11:A12"><cfRule type="expression" dxfId="0" priority="10"><formula>$A11&gt;0</formula></cfRule></conditionalFormatting></worksheet>',
  ));
  return zipSync(archive);
}

test("unmodeled conditional-format rules keep a remapped sqref after export", async (t) => {
  const downloads = installDownloadCapture(t);
  const source = workbookWithUnmodeledRule(new Uint8Array(sampleBuffer()));
  const { app, controller } = await mountController({
    file: source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength),
  });
  try {
    controller.setCellValue({ row: 11, col: 0 }, "marker");
    assert.equal(controller.insertRows(11), true);
    const bytes = await exportedBytes(downloads, controller);
    const xml = strFromU8(unzipSync(bytes)["xl/worksheets/sheet1.xml"]);
    assert.match(xml, /conditionalFormatting sqref="A11:A13"/);
    assert.match(xml, /type="expression"/, "the unmodeled rule survives export");
  } finally {
    app.unmount();
  }
});

test("number formats and defined names persist through export and reload", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController({ file: sampleBuffer() });
  try {
    controller.selectCell({ row: 55, col: 0 });
    controller.setCellValue({ row: 55, col: 0 }, "1234.5");
    controller.selectCell({ row: 55, col: 0 });
    assert.equal(
      controller.setSelectionNumberFormat({ formatType: "custom", formatString: "#,##0.00" }),
      true,
    );
    assert.equal(
      controller.getActiveWorksheet().getCellStyleAt(55, 0)?.numberFormat?.formatString,
      "#,##0.00",
    );

    assert.equal(controller.setNamedRange("Revenue", "'Dashboard'!$A$1:$B$2"), true);
    assert.ok(controller.getNamedRanges().some((entry) => entry.name === "Revenue"));
    assert.equal(controller.setNamedRange("Revenue", "'Dashboard'!$A$1:$C$3"), true);
    assert.equal(
      controller.getNamedRanges().find((entry) => entry.name === "Revenue")?.refersTo,
      "'Dashboard'!$A$1:$C$3",
      "updating a name replaces its reference",
    );

    let bytes = await exportedBytes(downloads, controller);
    let reloaded = await mountController({ file: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) });
    try {
      assert.equal(
        reloaded.controller.getActiveWorksheet().getCellStyleAt(55, 0)?.numberFormat?.formatString,
        "#,##0.00",
      );
      assert.equal(
        reloaded.controller.getNamedRanges().find((entry) => entry.name === "Revenue")?.refersTo,
        "'Dashboard'!$A$1:$C$3",
      );
      assert.equal(reloaded.controller.deleteNamedRange("Revenue"), true);
      bytes = await exportedBytes(downloads, reloaded.controller);
    } finally {
      reloaded.app.unmount();
    }

    const finalReload = await mountController({ file: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) });
    try {
      assert.equal(
        finalReload.controller.getNamedRanges().some((entry) => entry.name === "Revenue"),
        false,
        "deleting a name removes it from the exported workbook",
      );
    } finally {
      finalReload.app.unmount();
    }
  } finally {
    app.unmount();
  }
});

test("selection helpers select whole rows, columns and the used range", async () => {
  const { app, controller } = await mountController({ file: sampleBuffer() });
  try {
    controller.selectRow(4);
    assert.equal(controller.selection.start.row, 4);
    assert.equal(controller.selection.end.row, 4);
    assert.equal(controller.selection.start.col, 0);
    assert.equal(controller.selection.end.col, Math.max(controller.activeSheet.maxUsedCol, controller.activeSheet.visibleCols.at(-1) ?? 0));

    controller.selectColumn(2);
    assert.equal(controller.selection.start.col, 2);
    assert.equal(controller.selection.end.col, 2);
    assert.equal(controller.selection.start.row, 0);

    controller.selectAll();
    assert.equal(controller.selection.start.row, 0);
    assert.equal(controller.selection.start.col, 0);
    assert.equal(controller.selection.end.row, Math.max(controller.activeSheet.maxUsedRow, controller.activeSheet.visibleRows.at(-1) ?? 0));
  } finally {
    app.unmount();
  }
});

test("insertRows shifts chart and image anchors in the drawing parts", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController({ file: sampleBuffer() });
  try {
    assert.equal(controller.charts[0]?.anchor?.from?.row, 1);
    assert.equal(controller.images[0]?.anchor?.from?.row, 17);
    assert.equal(controller.insertRows(0), true);
    assert.equal(controller.charts[0]?.anchor?.from?.row, 2);
    assert.equal(controller.images[0]?.anchor?.from?.row, 18);

    const bytes = await exportedBytes(downloads, controller);
    const reloaded = await mountController({ file: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) });
    try {
      assert.equal(reloaded.controller.charts[0]?.anchor?.from?.row, 2);
      assert.equal(reloaded.controller.images[0]?.anchor?.from?.row, 18);
    } finally {
      reloaded.app.unmount();
    }
  } finally {
    app.unmount();
  }
});

test("mergeAcrossSelection merges each row independently", async () => {
  const { app, controller } = await mountController({ file: sampleBuffer() });
  try {
    controller.selectRange({ start: { row: 60, col: 0 }, end: { row: 61, col: 2 } });
    controller.mergeAcrossSelection();
    assert.deepEqual(
      controller.activeSheet.mergedRegions.filter((range) => range.start.row >= 60).map((range) => [
        range.start.row,
        range.start.col,
        range.end.row,
        range.end.col,
      ]),
      [[60, 0, 60, 2], [61, 0, 61, 2]],
    );
  } finally {
    app.unmount();
  }
});

test("the ribbon exposes structural, paste-special, number-format and name entry points", async () => {
  const calls = [];
  const controller = {
    activeCell: { row: 2, col: 1 },
    activeSheet: { maxUsedCol: 3, maxUsedRow: 5, name: "Sheet1", visibleCols: [0, 1, 2, 3], visibleRows: [0, 1, 2], rowStyleIds: {}, colStyleIds: {}, styleById: {} },
    getActiveWorksheet: () => null,
    canRedo: false,
    canUndo: false,
    canZoomIn: false,
    canZoomOut: false,
    error: null,
    freezePanes: null,
    isLoading: false,
    readOnly: false,
    revision: 0,
    selection: { start: { row: 2, col: 1 }, end: { row: 2, col: 1 } },
    tabs: [],
    zoomScale: 100,
    cutSelection: () => calls.push(["cut"]),
    copySelectionToClipboard: async () => { calls.push(["copy"]); return true; },
    pasteFromClipboard: async () => { calls.push(["paste"]); return true; },
    insertRows: (row) => calls.push(["insertRows", row]),
    deleteRows: (row) => calls.push(["deleteRows", row]),
    insertColumns: (col) => calls.push(["insertColumns", col]),
    deleteColumns: (col) => calls.push(["deleteColumns", col]),
  };
  const authoring = [];
  const mounted = await mount(XlsxRibbon, {
    controller,
    onAuthoring: (kind) => authoring.push(kind),
  });
  try {
    findByTestId(mounted.root, "xlsx-ribbon-cut").props.onClick();
    findByTestId(mounted.root, "xlsx-ribbon-copy").props.onClick();
    findByTestId(mounted.root, "xlsx-ribbon-paste").props.onClick();
    findByTestId(mounted.root, "xlsx-ribbon-insert-row-above").props.onClick();
    findByTestId(mounted.root, "xlsx-ribbon-insert-row-below").props.onClick();
    findByTestId(mounted.root, "xlsx-ribbon-delete-row").props.onClick();
    findByTestId(mounted.root, "xlsx-ribbon-insert-column-left").props.onClick();
    findByTestId(mounted.root, "xlsx-ribbon-insert-column-right").props.onClick();
    findByTestId(mounted.root, "xlsx-ribbon-delete-column").props.onClick();
    findByTestId(mounted.root, "xlsx-ribbon-paste-special").props.onClick();
    findByTestId(mounted.root, "xlsx-ribbon-number-format").props.onClick();
    findByTestId(mounted.root, "xlsx-ribbon-names").props.onClick();
    await vue.nextTick();
    assert.deepEqual(calls, [
      ["cut"],
      ["copy"],
      ["paste"],
      ["insertRows", 2],
      ["insertRows", 3],
      ["deleteRows", 2],
      ["insertColumns", 1],
      ["insertColumns", 2],
      ["deleteColumns", 1],
    ]);
    assert.deepEqual(authoring, ["pasteSpecial", "numberFormat", "names"]);
  } finally {
    mounted.app.unmount();
  }
});

function fakeController(overrides = {}) {
  const calls = [];
  const base = {
    activeCell: { row: 2, col: 1 },
    activeCellAddress: "B3",
    activeSheet: { maxUsedCol: 3, maxUsedRow: 5, name: "Sheet1", visibleCols: [0, 1, 2, 3], visibleRows: [0, 1, 2], rowStyleIds: {}, colStyleIds: {}, styleById: {} },
    getActiveWorksheet: () => null,
    revision: 0,
    selection: { start: { row: 2, col: 1 }, end: { row: 4, col: 3 } },
    selectedRangeAddress: "B3:D5",
    getNamedRanges: () => [],
    getCellComment: () => null,
    getCellHyperlink: () => null,
    getTableAtCell: () => null,
    getConditionalFormatRules: () => [],
    getCellDataValidation: () => null,
    getChartById: () => null,
    selectedChartId: null,
    ...overrides,
  };
  return { calls, controller: base };
}

