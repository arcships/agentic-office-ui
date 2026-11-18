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

function findByTestIdInBody(testId) {
  return walk(teleportTarget("body")).find((node) => node.props?.["data-testid"] === testId);
}

function setModel(node, value) {
  const update = node?.props?.["onUpdate:modelValue"];
  assert.equal(typeof update, "function", "control exposes a model update handler");
  update(value);
}

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

async function mountDialog(kind, overrides) {
  const { calls, controller } = fakeController(overrides);
  const Harness = vue.defineComponent({
    render: () => vue.h("div", [vue.h(XlsxAuthoringDialog, { controller, kind })]),
  });
  const mounted = await mount(Harness);
  return { ...mounted, calls, controller };
}

test("paste-special dialog applies the chosen mode and transpose flag", async () => {
  const calls = [];
  const { app } = await mountDialog("pasteSpecial", {
    pasteSpecialFromClipboard: async (options) => { calls.push(options); return true; },
  });
  try {
    assert.ok(findByTestIdInBody("xlsx-paste-special-section"));
    setModel(findByTestIdInBody("xlsx-paste-special-mode"), "values");
    setModel(findByTestIdInBody("xlsx-paste-special-transpose"), true);
    await vue.nextTick();
    findByTestIdInBody("xlsx-paste-special-apply").props.onClick();
    await vue.nextTick();
    assert.deepEqual(calls, [{ mode: "values", transpose: true }]);
  } finally {
    app.unmount();
  }
});

test("number-format dialog builds decimal, separator and custom patterns", async () => {
  const calls = [];
  const { app } = await mountDialog("numberFormat", {
    setSelectionNumberFormat: (format) => { calls.push(format); return true; },
  });
  try {
    assert.ok(findByTestIdInBody("xlsx-number-format-section"));
    setModel(findByTestIdInBody("xlsx-number-format-decimals"), 0);
    setModel(findByTestIdInBody("xlsx-number-format-thousands"), false);
    await vue.nextTick();
    assert.match(textContent(findByTestIdInBody("xlsx-number-format-preview")), /0/);
    findByTestIdInBody("xlsx-number-format-apply").props.onClick();
    assert.deepEqual(calls.at(-1), { formatString: "0", formatType: "custom" });

    setModel(findByTestIdInBody("xlsx-number-format-category"), "custom");
    await vue.nextTick();
    setModel(findByTestIdInBody("xlsx-number-format-pattern"), '0.0"%"');
    await vue.nextTick();
    findByTestIdInBody("xlsx-number-format-apply").props.onClick();
    assert.deepEqual(calls.at(-1), { formatString: '0.0"%"', formatType: "custom" });
  } finally {
    app.unmount();
  }
});

test("name manager lists, adds, updates and deletes defined names", async () => {
  let names = [{ name: "Existing", refersTo: "'Sheet1'!$A$1:$B$2" }];
  const calls = [];
  const { app } = await mountDialog("names", {
    getNamedRanges: () => names,
    setNamedRange: (name, refersTo) => { calls.push(["set", name, refersTo]); names = [{ name, refersTo }]; return true; },
    deleteNamedRange: (name) => { calls.push(["delete", name]); names = names.filter((entry) => entry.name !== name); return true; },
  });
  try {
    assert.ok(findByTestIdInBody("xlsx-name-item"), "existing names are listed");
    findByTestIdInBody("xlsx-name-use-selection").props.onClick();
    await vue.nextTick();
    assert.equal(
      findByTestIdInBody("xlsx-name-refers-to").value,
      "'Sheet1'!$B$3:$D$5",
      "the name manager can point a name at the current selection",
    );
    setModel(findByTestIdInBody("xlsx-name-input"), "Revenue");
    setModel(findByTestIdInBody("xlsx-name-refers-to"), "'Sheet1'!$C$1:$C$3");
    await vue.nextTick();
    findByTestIdInBody("xlsx-name-save").props.onClick();
    await vue.nextTick();
    assert.deepEqual(calls[0], ["set", "Revenue", "'Sheet1'!$C$1:$C$3"]);
    assert.ok(names.some((entry) => entry.name === "Revenue"));
    findByTestIdInBody("xlsx-name-delete").props.onClick();
    await vue.nextTick();
    assert.deepEqual(calls[1], ["delete", "Revenue"]);
  } finally {
    app.unmount();
  }
});

test("context menu exposes cut/copy, paste special, structural, merge and selection commands", () => {
  const source = readFileSync(
    new URL("../../packages/vue-xlsx/src/components/XlsxContextMenu.vue", import.meta.url),
    "utf8",
  );
  for (const testId of [
    "xlsx-context-cut",
    "xlsx-context-copy",
    "xlsx-context-paste",
    "xlsx-context-paste-special",
    "xlsx-context-paste-all",
    "xlsx-context-paste-values",
    "xlsx-context-paste-formats",
    "xlsx-context-paste-transpose",
    "xlsx-context-edit-cell",
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
  ]) {
    assert.match(source, new RegExp(`data-testid="${testId}"`), `context menu has ${testId}`);
  }
  assert.match(source, /Ctrl\+X/, "cut shows its shortcut");
  assert.match(source, /Ctrl\+V/, "paste shows its shortcut");
  assert.match(source, /data-submenu="pasteSpecial"/, "paste special opens a submenu");
  assert.match(source, /aria-checked="selectionIsRow"/, "row selection shows a checked state");
  assert.match(source, /event\.key === "ArrowRight"/, "right arrow opens the submenu");
  assert.match(source, /event\.key === "ArrowLeft"/, "left arrow closes the submenu");
  assert.match(source, /Type-ahead/, "menus support type-ahead");

  const headerMenu = readFileSync(
    new URL("../../packages/vue-xlsx/src/components/XlsxHeaderMenu.vue", import.meta.url),
    "utf8",
  );
  assert.match(headerMenu, /xlsx-column-menu-list/);
  assert.match(headerMenu, /xlsx-row-menu-list/);
  assert.match(headerMenu, /xlsx-column-menu-insert-left/);
  assert.match(headerMenu, /xlsx-column-menu-insert-right/);
  assert.match(headerMenu, /xlsx-column-menu-delete/);
  assert.match(headerMenu, /xlsx-row-menu-insert-above/);
  assert.match(headerMenu, /xlsx-row-menu-delete/);
  assert.match(headerMenu, /resolveOfficeMenuKeyCommand/);

  const grid = readFileSync(
    new URL("../../packages/vue-xlsx/src/components/XlsxGrid.vue", import.meta.url),
    "utf8",
  );
  assert.match(grid, /data-testid="xlsx-column-menu"/);
  assert.match(grid, /data-testid="xlsx-row-menu"/);
  assert.match(grid, /<XlsxHeaderMenu/);
  assert.match(grid, /controller\.editRequest/, "the grid consumes inline-edit requests");
});

test("context menu routes header and structural commands to the controller", async () => {
  const calls = [];
  const controller = {
    activeCell: { row: 1, col: 1 },
    activeSheet: {
      maxUsedCol: 3,
      maxUsedRow: 5,
      visibleCols: [0, 1, 2, 3],
      colWidths: [80, 80, 80, 80],
      rowHeights: [24, 24, 24, 24, 24, 24],
      defaultColWidthPx: 80,
      defaultRowHeightPx: 24,
      visibleRows: [0, 1, 2, 3, 4, 5],
    },
    freezePanes: null,
    isEditingLocked: false,
    readOnly: false,
    revision: 0,
    selection: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } },
    tables: [],
    zoomScale: 100,
    clearSelectedCells: () => calls.push(["clear"]),
    insertRows: (row) => calls.push(["insertRows", row]),
    deleteRows: (row) => calls.push(["deleteRows", row]),
    insertColumns: (col) => calls.push(["insertColumns", col]),
    deleteColumns: (col) => calls.push(["deleteColumns", col]),
    selectRow: (row) => calls.push(["selectRow", row]),
    selectColumn: (col) => calls.push(["selectColumn", col]),
    selectAll: () => calls.push(["selectAll"]),
    selectCell: (cell) => calls.push(["selectCell", cell]),
    cutSelection: () => { calls.push(["cut"]); return true; },
    pasteFromClipboard: async () => true,
    pasteSpecialFromClipboard: async (options) => { calls.push(["pasteSpecial", options]); return true; },
    editCell: () => { calls.push(["editCell"]); return true; },
    undo: () => {},
    redo: () => {},
  };
  const listeners = new Map();
  const target = {
    addEventListener(type, listener) { listeners.set(type, listener); },
    getBoundingClientRect: () => ({ left: 0, top: 0 }),
    removeEventListener(type) { listeners.delete(type); },
    scrollLeft: 0,
    scrollTop: 0,
  };
  const dispatchContextMenu = (clientX, clientY) => {
    listeners.get("contextmenu")?.({
      type: "contextmenu",
      clientX,
      clientY,
      preventDefault() {},
      target: undefined,
    });
  };
  const mounted = await mount(XlsxContextMenu, { controller, targetElement: target });
  try {
    // Column header right-click: 10px from the left, inside the 24px strip.
    dispatchContextMenu(90, 10);
    await vue.nextTick();
    const menu = findByTestIdInBody("xlsx-context-menu");
    assert.ok(menu, "the menu renders for a header hit");
    assert.equal(menu.props["data-menu-kind"], "column");
    findByTestIdInBody("xlsx-context-insert-column-right").props.onClick();
    assert.deepEqual(calls.at(-1), ["insertColumns", 1]);

    const rowListeners = new Map();
    const rowTarget = {
      addEventListener(type, listener) { rowListeners.set(type, listener); },
      getBoundingClientRect: () => ({ left: 0, top: 0 }),
      removeEventListener(type) { rowListeners.delete(type); },
      scrollLeft: 0,
      scrollTop: 0,
    };
    const rowMounted = await mount(XlsxContextMenu, { controller, targetElement: rowTarget });
    try {
      rowListeners.get("contextmenu")?.({ type: "contextmenu", clientX: 10, clientY: 40, preventDefault() {}, target: undefined });
      await vue.nextTick();
      const rowMenu = findByTestIdInBody("xlsx-context-menu");
      assert.equal(rowMenu.props["data-menu-kind"], "row");
      findByTestIdInBody("xlsx-context-delete-row").props.onClick();
      assert.deepEqual(calls.at(-1), ["deleteRows", 0]);
    } finally {
      rowMounted.app.unmount();
    }
  } finally {
    mounted.app.unmount();
  }
});
