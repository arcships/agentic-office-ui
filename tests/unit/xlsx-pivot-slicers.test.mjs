import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

// Issue #22 acceptance: pivot tables and slicers are detected from the package
// instead of being invisible, their metadata survives export/reload, and the
// worksheet/model exposes it to hosts.

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);
const xmldom = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= xmldom.DOMParser;
globalThis.Node ??= xmldom.Node;
globalThis.XMLSerializer ??= xmldom.XMLSerializer;

const core = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);

const wasmPath = requireFromXlsxCore.resolve(
  "@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm",
);
const wasmBytes = readFileSync(wasmPath);
core.setWasmSource(
  wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
);
const wasm = await core.getSheetsWasmModule();

function corpusBytes(name) {
  return new Uint8Array(readFileSync(new URL(`../../test-data/corpus/${name}`, import.meta.url)));
}

function sampleBytes(name) {
  return new Uint8Array(
    readFileSync(new URL(`../../apps/demo/public/samples/${name}`, import.meta.url)),
  );
}

function loadAssets(bytes) {
  return core.parseWorkbookImageAssets(bytes);
}

test("pivot table metadata is parsed from the worksheet parts", () => {
  const assets = loadAssets(corpusBytes("xlsx-pivot-slicers.xlsx"));
  const state = assets.sheetStatesByWorkbookSheetIndex[0];
  assert.equal(state.pivotTables.length, 1);
  const pivot = state.pivotTables[0];

  assert.equal(pivot.name, "PivotTable1");
  assert.equal(pivot.cacheId, 1);
  assert.deepEqual(pivot.range, {
    end: { col: 12, row: 13 },
    start: { col: 9, row: 9 },
  });
  assert.equal(pivot.layout, "outline");
  assert.equal(pivot.refreshOnLoad, true);
  assert.deepEqual(pivot.source, { sheet: "Data", range: "A1:C6", table: null });
  assert.deepEqual(pivot.rowFields, [0]);
  assert.deepEqual(pivot.columnFields, [1]);
  assert.deepEqual(pivot.pageFields, []);
  assert.equal(pivot.styleInfo.name, "PivotStyleLight16");

  assert.deepEqual(
    pivot.fields.map((field) => ({ axis: field.axis, name: field.name })),
    [
      { axis: "row", name: "Region" },
      { axis: "column", name: "Product" },
      { axis: "data", name: "Revenue" },
    ],
  );
  assert.deepEqual(
    pivot.fields[0].items.map((item) => item.caption),
    ["East", "West", "North", "South"],
  );
  assert.deepEqual(pivot.dataFields, [
    {
      aggregation: "sum",
      fieldIndex: 2,
      name: "Sum of Revenue",
      numberFormatId: 0,
      showAs: "normal",
    },
  ]);
});

test("slicers are detected with their cache, items and drawing anchor", () => {
  const assets = loadAssets(corpusBytes("xlsx-pivot-slicers.xlsx"));
  const state = assets.sheetStatesByWorkbookSheetIndex[0];
  assert.equal(state.slicers.length, 1);
  const [slicer] = state.slicers;

  assert.equal(slicer.name, "Region");
  assert.equal(slicer.caption, "Region");
  assert.equal(slicer.cacheName, "Slicer_Region");
  assert.equal(slicer.sourceName, "Region");
  assert.equal(slicer.style, "SlicerStyleLight1");
  assert.equal(slicer.columnCount, 1);
  assert.deepEqual(slicer.pivotTables, ["PivotTable1"]);
  assert.deepEqual(
    slicer.items.map((item) => ({ caption: item.caption, selected: item.selected })),
    [
      { caption: "East", selected: false },
      { caption: "West", selected: true },
      { caption: "North", selected: false },
      { caption: "South", selected: false },
    ],
  );
  assert.equal(slicer.anchor.kind, "two-cell");
  assert.deepEqual(slicer.anchor.from, { col: 13, colOffsetEmu: 0, row: 2, rowOffsetEmu: 0 });
  assert.deepEqual(slicer.anchor.to, { col: 16, colOffsetEmu: 0, row: 14, rowOffsetEmu: 0 });
});

test("defined names expose workbook and sheet scope without built-in print names", () => {
  const assets = loadAssets(corpusBytes("xlsx-pivot-slicers.xlsx"));
  const state = assets.sheetStatesByWorkbookSheetIndex[0];
  assert.deepEqual(state.definedNames, [
    { comment: undefined, hidden: false, name: "SalesData", range: "Data!$A$2:$C$6", scope: "workbook" },
    { comment: undefined, hidden: false, name: "RevenueTotal", range: "SUM(Data!$C$2:$C$6)", scope: "workbook" },
    {
      comment: undefined,
      hidden: false,
      name: "CurrentProduct",
      range: "Data!$B$5",
      scope: "sheet",
      sheetName: "Data",
    },
  ]);
});

test("pivot layouts, aggregations and show-as values are classified", () => {
  const base = corpusBytes("xlsx-pivot-slicers.xlsx");
  const patchPivot = (replace) => {
    const archive = unzipSync(base);
    archive["xl/pivotTables/pivotTable1.xml"] = strToU8(
      replace(strFromU8(archive["xl/pivotTables/pivotTable1.xml"])),
    );
    return zipSync(archive);
  };

  const compact = core.parseWorksheetPivotTables(
    unzipSync(patchPivot((xml) => xml.replace(/compact="0" outline="1" outlineData="1" gridDropZones="1"/, 'compact="1"'))),
    "xl/worksheets/sheet1.xml",
  );
  assert.equal(compact[0].layout, "compact");

  const tabular = core.parseWorksheetPivotTables(
    unzipSync(patchPivot((xml) => xml.replace(/outline="1" outlineData="1"/, 'outline="0" outlineData="0"'))),
    "xl/worksheets/sheet1.xml",
  );
  assert.equal(tabular[0].layout, "tabular");

  const averagePercent = core.parseWorksheetPivotTables(
    unzipSync(patchPivot((xml) => xml
      .replace('subtotal="sum"', 'subtotal="average"')
      .replace('showDataAs="normal"', 'showDataAs="percentOfCol"'))),
    "xl/worksheets/sheet1.xml",
  );
  assert.equal(averagePercent[0].dataFields[0].aggregation, "average");
  assert.equal(averagePercent[0].dataFields[0].showAs, "percentOfColumnTotal");
});

test("workbooks without pivots or slicers report empty metadata", () => {
  for (const name of ["financial-model.xlsx", "sales-table.xlsx"]) {
    const assets = core.parseWorkbookStructureAssets(sampleBytes(name));
    const state = assets.sheetStatesByWorkbookSheetIndex[0];
    assert.deepEqual(state.pivotTables, [], name);
    assert.deepEqual(state.slicers, [], name);
    assert.deepEqual(state.definedNames, [], name);
  }
});

test("pivot, slicer and defined-name parts survive export and reload", () => {
  const bytes = corpusBytes("xlsx-pivot-slicers.xlsx");
  const assets = loadAssets(bytes);
  const workbook = wasm.Workbook.fromBytes(bytes);
  const saved = core.mergeWorkbookSheetAuthoringAssets(
    workbook.saveXlsxBytes(),
    assets.archive,
  );
  const archive = unzipSync(saved);
  for (const part of [
    "xl/pivotTables/pivotTable1.xml",
    "xl/pivotCache/pivotCacheDefinition1.xml",
    "xl/pivotCache/pivotCacheRecords1.xml",
    "xl/slicers/slicer1.xml",
    "xl/slicerCaches/slicerCache1.xml",
    "xl/drawings/drawing1.xml",
  ]) {
    assert.ok(archive[part], `${part} is kept`);
  }
  assert.match(strFromU8(archive["xl/worksheets/sheet1.xml"]), /<pivotTableParts[^>]*>/);
  assert.match(strFromU8(archive["xl/worksheets/sheet1.xml"]), /<x14:slicerList>/);
  assert.match(strFromU8(archive["xl/workbook.xml"]), /<pivotCaches[^>]*>/);
  assert.match(strFromU8(archive["xl/workbook.xml"]), /<x14:slicerCaches>/);
  assert.match(
    strFromU8(archive["xl/_rels/workbook.xml.rels"]),
    /relationships\/pivotCacheDefinition/,
  );
  assert.match(strFromU8(archive["xl/workbook.xml"]), /name="SalesData"/);

  const reloaded = core.parseWorkbookStructureAssets(saved);
  const state = reloaded.sheetStatesByWorkbookSheetIndex[0];
  assert.equal(state.pivotTables.length, 1);
  assert.equal(state.pivotTables[0].name, "PivotTable1");
  assert.equal(state.pivotTables[0].dataFields[0].name, "Sum of Revenue");
  assert.equal(state.pivotTables[0].fields[0].items.length, 4);
  assert.equal(state.slicers.length, 1);
  assert.equal(state.slicers[0].items[1].caption, "West");
  assert.equal(state.slicers[0].items[1].selected, true);
  assert.deepEqual(state.definedNames.map((entry) => entry.name), [
    "SalesData",
    "RevenueTotal",
    "CurrentProduct",
  ]);
});

test("the workbook-level cache and slicer relationships are valid", () => {
  const bytes = corpusBytes("xlsx-pivot-slicers.xlsx");
  const assets = loadAssets(bytes);
  const saved = core.mergeWorkbookSheetAuthoringAssets(
    wasm.Workbook.fromBytes(bytes).saveXlsxBytes(),
    assets.archive,
  );
  const archive = unzipSync(saved);
  const parseRelationships = (path) => {
    const document = new DOMParser().parseFromString(strFromU8(archive[path]), "application/xml");
    return Array.from(document.getElementsByTagName("Relationship")).map((node) => ({
      id: node.getAttribute("Id"),
      target: node.getAttribute("Target"),
      type: node.getAttribute("Type"),
    }));
  };
  const workbookRelationships = parseRelationships("xl/_rels/workbook.xml.rels");
  const sheetRelationships = parseRelationships("xl/worksheets/_rels/sheet1.xml.rels");
  assert.ok(
    workbookRelationships.some((relationship) => relationship.type.endsWith("/pivotCacheDefinition")),
    "workbook keeps the pivot cache definition relationship",
  );
  assert.ok(
    workbookRelationships.some((relationship) => relationship.type.endsWith("/slicerCache")),
    "workbook keeps the slicer cache relationship",
  );
  assert.ok(
    sheetRelationships.some((relationship) => relationship.type.endsWith("/pivotTable")),
    "sheet keeps the pivot table relationship",
  );
  assert.ok(
    sheetRelationships.some((relationship) => relationship.type.endsWith("/slicer")),
    "sheet keeps the slicer relationship",
  );
  assert.ok(
    sheetRelationships.some((relationship) => relationship.type.endsWith("/drawing")),
    "sheet keeps its drawing relationship",
  );
});
