import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

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
const core = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);
const wasmBytes = readFileSync(
  requireFromXlsxCore.resolve("@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm")
);
core.setWasmSource(
  wasmBytes.buffer.slice(
    wasmBytes.byteOffset,
    wasmBytes.byteOffset + wasmBytes.byteLength
  )
);
const wasm = await core.getSheetsWasmModule();

const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

function bytesFromFile(name) {
  return new Uint8Array(readFileSync(new URL(name, samples)));
}

function patchArchive(bytes, replacements) {
  const archive = unzipSync(bytes);
  for (const [path, replace] of Object.entries(replacements)) {
    archive[path] = strToU8(replace(strFromU8(archive[path])));
  }
  return zipSync(archive);
}

function patchWorksheetAutoFilter(bytes, snippet) {
  const replacements = {};
  for (const path of Object.keys(unzipSync(bytes))) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(path)) continue;
    replacements[path] = (xml) => xml.replace("</sheetData>", `</sheetData>${snippet}`);
  }
  return patchArchive(bytes, replacements);
}

const WORKSHEET_AUTO_FILTER =
  '<autoFilter ref="A1:D8"><filterColumn colId="0"><filters blank="1"><filter val="North"/><filter val="West"/></filters></filterColumn><filterColumn colId="2" hiddenButton="1"/></autoFilter>';

function worksheetFilterFixture() {
  return patchWorksheetAutoFilter(bytesFromFile("financial-model.xlsx"), WORKSHEET_AUTO_FILTER);
}

function tableFilterFixture(extraTableAttributes = "") {
  return patchArchive(bytesFromFile("sales-table.xlsx"), {
    "xl/tables/table1.xml": (xml) =>
      xml.replace(
        "<table ",
        `<table${extraTableAttributes ? ` ${extraTableAttributes}` : ""} `,
      ).replace(
        '<autoFilter ref="A1:H181"/>',
        '<autoFilter ref="A1:H181"><filterColumn colId="1"><filters><filter val="Rep 2"/></filters></filterColumn></autoFilter>',
      ),
  });
}

test("worksheet <autoFilter> parses its range and value filter columns", () => {
  const assets = core.parseWorkbookStructureAssets(worksheetFilterFixture());
  const ranges = assets.sheetStatesByWorkbookSheetIndex[0].autoFilterRanges;
  assert.equal(ranges.length, 1);
  const range = ranges[0];
  assert.equal(range.source, "worksheet");
  assert.equal(range.reference, "A1:D8");
  assert.deepEqual(range.start, { row: 0, col: 0 });
  assert.deepEqual(range.end, { row: 7, col: 3 });
  assert.equal(range.headerRowCount, 1);
  assert.equal(range.dataEndRow, 7);
  assert.equal(range.tableName, undefined);
  assert.deepEqual(range.columns, [
    {
      columnIndex: 0,
      includeBlank: true,
      showButton: true,
      values: ["North", "West"],
    },
    {
      columnIndex: 2,
      includeBlank: false,
      showButton: false,
      values: [],
    },
  ]);
});

test("table parts keep their own <autoFilter> in the table metadata and the composed ranges", () => {
  const assets = core.parseWorkbookStructureAssets(tableFilterFixture());
  const table = assets.tableMetadataByWorkbookSheetIndex[0][0];
  assert.equal(table.name, "SalesRecords");
  assert.equal(table.autoFilter.reference, "A1:H181");
  assert.equal(table.autoFilter.source, "table");
  assert.equal(table.autoFilter.tableName, "SalesRecords");
  assert.deepEqual(table.autoFilter.columns, [
    { columnIndex: 1, includeBlank: false, showButton: true, values: ["Rep 2"] },
  ]);
  const ranges = assets.sheetStatesByWorkbookSheetIndex[0].autoFilterRanges;
  assert.deepEqual(ranges.map((range) => [range.reference, range.source, range.tableName]), [
    ["A1:H181", "table", "SalesRecords"],
  ]);
});

test("a table filter excludes its totals row from the data range", () => {
  const assets = core.parseWorkbookStructureAssets(
    tableFilterFixture('totalsRowShown="1" totalsRowCount="1"'),
  );
  const range = assets.sheetStatesByWorkbookSheetIndex[0].autoFilterRanges[0];
  assert.equal(range.end.row, 180);
  assert.equal(range.dataEndRow, 179, "the totals row is not data");
});

test("composeAutoFilterRanges dedupes a worksheet filter already owned by a table", () => {
  const tableFilter = {
    columns: [],
    dataEndRow: 9,
    end: { row: 9, col: 3 },
    headerRowCount: 1,
    reference: "A1:D10",
    source: "table",
    start: { row: 0, col: 0 },
    tableName: "SalesRecords",
  };
  const worksheetFilter = { ...tableFilter, source: "worksheet", tableName: undefined };
  const composed = core.composeAutoFilterRanges(
    [worksheetFilter],
    [{ autoFilter: tableFilter, name: "SalesRecords" }],
  );
  assert.equal(composed.length, 1, "one reference yields one header menu");
  assert.equal(composed[0].source, "table");
  assert.equal(composed[0].tableName, "SalesRecords");

  const separate = core.composeAutoFilterRanges(
    [{ ...worksheetFilter, reference: "A12:D20", start: { row: 11, col: 0 }, end: { row: 19, col: 3 } }],
    [{ autoFilter: tableFilter, name: "SalesRecords" }],
  );
  assert.deepEqual(separate.map((range) => range.reference), ["A1:D10", "A12:D20"]);
});

test("multi-range autoFilter refs keep the filter columns on the first range", () => {
  const bytes = patchWorksheetAutoFilter(
    bytesFromFile("financial-model.xlsx"),
    '<autoFilter ref="A1:D5 A8:D12"><filterColumn colId="3"><filters><filter val="Q1"/></filters></filterColumn></autoFilter>',
  );
  const ranges = core.parseWorkbookStructureAssets(bytes).sheetStatesByWorkbookSheetIndex
    .flatMap((state) => state?.autoFilterRanges ?? [])
    .filter((range) => range.source === "worksheet");
  const firstSheetRanges = [ranges[0], ranges[1]];
  assert.deepEqual(firstSheetRanges.map((range) => range.reference), ["A1:D5", "A8:D12"]);
  assert.deepEqual(firstSheetRanges[0].columns.map((column) => [column.columnIndex, column.values]), [[3, ["Q1"]]]);
  assert.deepEqual(firstSheetRanges[1].columns, []);
});

test("applyWorkbookAutoFilterOverrides rewrites cleared and applied criteria", () => {
  const source = worksheetFilterFixture();
  const cleared = core.applyWorkbookAutoFilterOverrides(source, {
    0: { "A1:D8": [] },
  });
  const clearedRanges = core.parseWorkbookStructureAssets(cleared).sheetStatesByWorkbookSheetIndex[0].autoFilterRanges;
  assert.deepEqual(clearedRanges[0].columns, [], "clearing removes the stored filter columns");

  const applied = core.applyWorkbookAutoFilterOverrides(source, {
    0: {
      "A1:D8": [
        { columnIndex: 2, includeBlank: true, values: ["Rep 2", "Rep 3"] },
      ],
    },
  });
  const appliedRange = core.parseWorkbookStructureAssets(applied).sheetStatesByWorkbookSheetIndex[0].autoFilterRanges[0];
  assert.deepEqual(appliedRange.columns, [
    { columnIndex: 2, includeBlank: true, showButton: true, values: ["Rep 2", "Rep 3"] },
  ]);

  assert.equal(
    core.applyWorkbookAutoFilterOverrides(source, {}),
    source,
    "a workbook without overrides is not rezipped",
  );
});

test("applyWorkbookAutoFilterOverrides patches a table filter too", () => {
  const source = tableFilterFixture();
  const applied = core.applyWorkbookAutoFilterOverrides(source, {
    0: { "A1:H181": [{ columnIndex: 0, includeBlank: false, values: ["North"] }] },
  });
  const archive = unzipSync(applied);
  const tableXml = strFromU8(archive["xl/tables/table1.xml"]);
  assert.match(
    tableXml,
    /<autoFilter ref="A1:H181"><filterColumn colId="0"><filters><filter val="North"\/><\/filters><\/filterColumn><\/autoFilter>/,
  );
  const range = core.parseWorkbookStructureAssets(applied).sheetStatesByWorkbookSheetIndex[0].autoFilterRanges[0];
  assert.deepEqual(range.columns[0].values, ["North"]);
});

test("export -> reload keeps worksheet and table AutoFilter definitions", () => {
  const source = worksheetFilterFixture();
  const workbook = wasm.Workbook.fromBytes(source);
  const saved = workbook.saveXlsxBytes();
  workbook.free();
  const merged = core.mergeWorkbookSheetAuthoringAssets(saved, unzipSync(source));
  const reloaded = core.parseWorkbookStructureAssets(merged);
  const worksheetRanges = reloaded.sheetStatesByWorkbookSheetIndex
    .flatMap((state) => state?.autoFilterRanges ?? [])
    .filter((range) => range.source === "worksheet");
  assert.equal(worksheetRanges.length, 3);
  assert.equal(worksheetRanges[0].reference, "A1:D8");
  assert.deepEqual(worksheetRanges[0].columns[0].values, ["North", "West"]);
});

test("export -> reload keeps a table's AutoFilter through the authoring merge", () => {
  const source = tableFilterFixture();
  const workbook = wasm.Workbook.fromBytes(source);
  const saved = workbook.saveXlsxBytes();
  workbook.free();
  // The engine drops table parts on save; the viewer's merge puts them back.
  const merged = core.mergeWorkbookSheetAuthoringAssets(saved, unzipSync(source));
  const tableRange = core.parseWorkbookStructureAssets(merged).sheetStatesByWorkbookSheetIndex[0]
    .autoFilterRanges.find((range) => range.source === "table");
  assert.equal(tableRange?.reference, "A1:H181");
  assert.equal(tableRange?.tableName, "SalesRecords");
  assert.deepEqual(tableRange?.columns[0].values, ["Rep 2"]);
});
