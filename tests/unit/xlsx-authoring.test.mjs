import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
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

const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

function sampleBytes(name) {
  return new Uint8Array(readFileSync(new URL(name, samples)));
}

const wasmPath = requireFromXlsxCore.resolve(
  "@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm",
);
const wasmBytes = readFileSync(wasmPath);
core.setWasmSource(
  wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
);
const wasm = await core.getSheetsWasmModule();

/** Loads a sample workbook and returns its mutable authoring assets. */
function loadAssets(name = "financial-model.xlsx") {
  const bytes = sampleBytes(name);
  return {
    bytes,
    assets: core.parseWorkbookImageAssets(bytes),
    workbook: wasm.Workbook.fromBytes(bytes),
  };
}

function sheetXml(archive, path = "xl/worksheets/sheet1.xml") {
  return strFromU8(archive[path]);
}

function sheetRels(archive, path = "xl/worksheets/_rels/sheet1.xml.rels") {
  return strFromU8(archive[path] ?? strToU8(""));
}

/** Saves through the engine, then merges authored objects exactly like export. */
function saveWithAuthoring({ workbook, assets }, mergeImages = false) {
  const saved = workbook.saveXlsxBytes();
  const authored = core.mergeWorkbookSheetAuthoringAssets(saved, assets.archive);
  return mergeImages
    ? core.mergeWorkbookImageAssets(authored, assets, assets.sheetOrigins)
    : authored;
}

test("freeze panes are written to the worksheet view and can be cleared", () => {
  const { assets } = loadAssets();
  assert.equal(core.setWorkbookFreezePanes(assets, 0, { row: 1, col: 2 }), true);

  const xml = sheetXml(assets.archive);
  assert.match(
    xml,
    /<pane[^>]*xSplit="2"[^>]*ySplit="1"[^>]*topLeftCell="C2"[^>]*state="frozen"/,
    "the frozen pane records the split counts and the first visible cell",
  );
  const sheetDocument = new DOMParser().parseFromString(xml, "application/xml");
  assert.deepEqual(
    core.parseWorksheetFreezePanesDocument(sheetDocument),
    { row: 1, col: 2 },
  );

  assert.equal(core.setWorkbookFreezePanes(assets, 0, null), true);
  assert.doesNotMatch(sheetXml(assets.archive), /<pane\b/);
  const cleared = new DOMParser().parseFromString(sheetXml(assets.archive), "application/xml");
  assert.equal(core.parseWorksheetFreezePanesDocument(cleared), null);
});

test("comments create, edit and delete the comments part with its VML drawing", () => {
  const { assets } = loadAssets();
  const cell = { row: 3, col: 2 };
  const created = core.setWorkbookCellComment(assets, 0, cell, {
    author: "Dana",
    text: "Check <this> & that",
  });
  assert.deepEqual(created, [
    { author: "Dana", cell, text: "Check <this> & that" },
  ]);

  const commentsPath = Object.keys(assets.archive).find((path) => path.startsWith("xl/comments/"));
  const vmlPath = Object.keys(assets.archive).find((path) => path.endsWith(".vml"));
  assert.ok(commentsPath, "a comments part is written");
  assert.ok(vmlPath, "a legacy VML drawing is written next to it");
  const commentsXml = strFromU8(assets.archive[commentsPath]);
  assert.match(commentsXml, /<author>Dana<\/author>/);
  assert.match(commentsXml, /Check &lt;this&gt; &amp; that/);
  assert.match(sheetXml(assets.archive), /<legacyDrawing r:id="[^"]+"\/>/);
  assert.match(sheetRels(assets.archive), /relationships\/comments/);
  assert.match(sheetRels(assets.archive), /relationships\/vmlDrawing/);

  const edited = core.setWorkbookCellComment(assets, 0, cell, { text: "Revised" });
  assert.deepEqual(edited, [{ author: undefined, cell, text: "Revised" }]);

  const cleared = core.setWorkbookCellComment(assets, 0, cell, null);
  assert.deepEqual(cleared, []);
  assert.doesNotMatch(sheetXml(assets.archive), /<legacyDrawing\b/);
  assert.equal(
    Object.keys(assets.archive).some((path) => path.startsWith("xl/comments/")),
    false,
    "deleting the last comment removes the comments part",
  );
});

test("hyperlinks keep external URLs as relationships and internal targets as locations", () => {
  const { assets } = loadAssets();
  const cell = { row: 4, col: 1 };
  const created = core.setWorkbookCellHyperlink(assets, 0, cell, {
    target: "https://example.com/report?a=1&b=2",
    tooltip: "Report",
  });
  assert.equal(created.length, 1);
  assert.equal(created[0].target, "https://example.com/report?a=1&b=2");
  assert.equal(created[0].tooltip, "Report");
  const rels = sheetRels(assets.archive);
  assert.match(rels, /Target="https:\/\/example.com\/report\?a=1&amp;b=2"/);
  assert.match(rels, /TargetMode="External"/);
  assert.match(sheetXml(assets.archive), /<hyperlink ref="B5" r:id="[^"]+" tooltip="Report"\/>/);

  const internal = core.setWorkbookCellHyperlink(assets, 0, { row: 5, col: 1 }, {
    target: "#P&L!A1",
  });
  assert.deepEqual(internal.map((link) => link.target).sort(), [
    "#P&L!A1",
    "https://example.com/report?a=1&b=2",
  ]);
  assert.match(sheetXml(assets.archive), /location="P&amp;L!A1"/);

  assert.deepEqual(core.setWorkbookCellHyperlink(assets, 0, cell, null), [
    { cell: { row: 5, col: 1 }, target: "#P&L!A1", tooltip: undefined },
  ]);
});

test("conditional formats round-trip color scales, data bars and icon sets", () => {
  const { assets } = loadAssets();
  const range = { start: { row: 2, col: 1 }, end: { row: 6, col: 1 } };

  const colorScale = core.createConditionalFormatRuleFromDraft(
    { colors: ["F8696B", "FFEB84", "63BE7B"], kind: "colorScale", thresholds: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }] },
    [range],
    1,
  );
  const added = core.addWorkbookConditionalFormatRule(assets, 0, colorScale);
  assert.equal(added.length, 1);
  assert.equal(added[0].kind, "colorScale");
  assert.deepEqual(
    added[0].cfvos.map((cfvo) => [cfvo.type, cfvo.value ?? null]),
    [["min", null], ["percentile", 50], ["max", null]],
  );

  const dataBar = core.createConditionalFormatRuleFromDraft(
    { colors: ["5B9BD5"], kind: "dataBar" },
    [range],
    2,
  );
  const iconSet = core.createConditionalFormatRuleFromDraft(
    { iconSet: "3TrafficLights1", kind: "iconSet" },
    [range],
    3,
  );
  const withAll = core.addWorkbookConditionalFormatRule(assets, 0, dataBar);
  const withIcons = core.addWorkbookConditionalFormatRule(assets, 0, iconSet);
  assert.equal(withAll.length, 2);
  assert.deepEqual(withIcons.map((rule) => rule.kind), ["colorScale", "dataBar", "iconSet"]);
  assert.deepEqual(withIcons.map((rule) => rule.priority), [1, 2, 3]);
  assert.deepEqual(withIcons[2].icons.map((icon) => icon.iconId), [0, 1, 2]);

  const updated = core.updateWorkbookConditionalFormatRule(assets, 0, 1, dataBar);
  assert.equal(updated[1].kind, "dataBar");

  const deleted = core.deleteWorkbookConditionalFormatRule(assets, 0, 0);
  assert.deepEqual(deleted.map((rule) => rule.kind), ["dataBar", "iconSet"]);
  assert.equal(core.deleteWorkbookConditionalFormatRule(assets, 0, 9), null);

  const xml = sheetXml(assets.archive);
  assert.match(xml, /<conditionalFormatting sqref="B3:B7"><cfRule type="dataBar" priority="1">/);
});

test("data validations are replaced per cell and can be removed", () => {
  const { assets } = loadAssets();
  const range = { start: { row: 1, col: 0 }, end: { row: 4, col: 0 } };
  const applied = core.setWorkbookDataValidation(assets, 0, [range], {
    allowBlank: true,
    listSource: '"Open,Closed"',
    ranges: [range],
    showDropdown: true,
    showErrorAlert: true,
    validationType: "list",
    errorMessage: "Pick a status",
  });
  assert.equal(applied.length, 1);
  assert.equal(applied[0].validationType, "list");
  assert.equal(applied[0].listSource, '"Open,Closed"');
  assert.equal(applied[0].errorMessage, "Pick a status");
  assert.doesNotMatch(
    sheetXml(assets.archive),
    /dataValidation[^>]*showDropDown="1"/,
    "showDropdown=true must not write the inverted OOXML flag",
  );

  const removed = core.setWorkbookDataValidation(assets, 0, [range], null);
  assert.deepEqual(removed, []);
  assert.doesNotMatch(sheetXml(assets.archive), /<dataValidations\b/);
});

test("tables create a ListObject part, relationship and tableParts entry", () => {
  const { assets } = loadAssets();
  const range = { start: { row: 0, col: 0 }, end: { row: 4, col: 3 } };
  const created = core.addWorkbookTable(assets, 0, {
    columnNames: ["Driver", "FY2026", "FY2027", "FY2028"],
    name: "BudgetTable",
    range,
  });
  assert.ok(created);
  assert.equal(created.table.name, "BudgetTable");
  assert.equal(created.table.reference, "A1:D5");
  assert.deepEqual(created.table.columns.map((column) => column.name), [
    "Driver",
    "FY2026",
    "FY2027",
    "FY2028",
  ]);
  assert.equal(created.table.styleInfo?.name, "TableStyleMedium2");

  const tablePath = Object.keys(assets.archive).find((path) => path.startsWith("xl/tables/"));
  assert.ok(tablePath);
  assert.match(sheetXml(assets.archive), /<tableParts count="1"><tablePart r:id="[^"]+"\/><\/tableParts>/);
  assert.match(sheetRels(assets.archive), /relationships\/table/);
  assert.match(
    strFromU8(assets.archive["[Content_Types].xml"]),
    /spreadsheetml\.table\+xml/,
  );

  assert.equal(core.addWorkbookTable(assets, 0, {
    columnNames: ["A", "B"],
    name: "BudgetTable",
    range: { start: { row: 6, col: 0 }, end: { row: 7, col: 1 } },
  }), null, "duplicate table names on the same sheet are rejected");
  assert.equal(core.addWorkbookTable(assets, 1, {
    columnNames: ["A", "B"],
    name: "budgettable",
    range: { start: { row: 6, col: 0 }, end: { row: 7, col: 1 } },
  }), null, "table names are unique across the whole workbook");

  const remaining = core.deleteWorkbookTable(assets, 0, "BudgetTable");
  assert.deepEqual(remaining, []);
  assert.doesNotMatch(sheetXml(assets.archive), /<tableParts\b/);
});

test("inserted charts write a chart part, drawing part and drawing attachment", () => {
  const { assets, workbook } = loadAssets();
  const range = { start: { row: 1, col: 0 }, end: { row: 4, col: 3 } };

  for (const [family, expectedType] of [
    ["bar", "ColumnClustered"],
    ["line", "Line"],
    ["pie", "Pie"],
  ]) {
    const inserted = core.insertWorkbookChart(assets, 0, {
      family,
      range,
      sheetName: "Dashboard",
      title: `${family} chart`,
    });
    assert.ok(inserted, `${family} chart is inserted`);
    assert.equal(inserted.chartType, expectedType);
    assert.ok(assets.archive[inserted.chartPath], "chart part exists");
    assert.ok(assets.archive[inserted.drawingPath], "drawing part exists");
    assert.ok(assets.archive[inserted.drawingRelsPath], "drawing relationships exist");
    assert.match(
      strFromU8(assets.archive[inserted.chartPath]),
      new RegExp(`<c:${family === "bar" ? "barChart" : family === "line" ? "lineChart" : "pieChart"}>`),
    );
  }

  const chartAssets = core.loadWorkbookChartAssets(
    workbook,
    assets,
    new Map([[0, 0], [1, 1], [2, 2]]),
    false,
  );
  const sheetCharts = chartAssets.chartsByWorkbookSheetIndex[0];
  assert.deepEqual(
    sheetCharts.map((chart) => chart.chartType),
    ["ColumnClustered", "Line", "Pie"],
  );
  assert.deepEqual(
    sheetCharts.map((chart) => chart.title),
    ["bar chart", "line chart", "pie chart"],
  );
  assert.equal(sheetCharts[0].series.length, 3, "three value columns become three series");
});

test("exported workbooks keep authored objects through a save/reload cycle", () => {
  const { assets, workbook } = loadAssets();
  const commentCell = { row: 3, col: 2 };
  core.setWorkbookFreezePanes(assets, 0, { row: 1, col: 0 });
  core.setWorkbookCellComment(assets, 0, commentCell, { author: "Dana", text: "Note" });
  core.setWorkbookCellHyperlink(assets, 0, { row: 4, col: 1 }, { target: "https://example.com" });
  core.addWorkbookConditionalFormatRule(
    assets,
    0,
    core.createConditionalFormatRuleFromDraft(
      { colors: ["5B9BD5"], kind: "dataBar" },
      [{ start: { row: 2, col: 1 }, end: { row: 4, col: 1 } }],
      1,
    ),
  );
  core.setWorkbookDataValidation(
    assets,
    0,
    [{ start: { row: 1, col: 0 }, end: { row: 4, col: 0 } }],
    {
      listSource: '"A,B"',
      ranges: [{ start: { row: 1, col: 0 }, end: { row: 4, col: 0 } }],
      validationType: "list",
    },
  );
  core.addWorkbookTable(assets, 0, {
    columnNames: ["Driver", "FY2026", "FY2027", "FY2028"],
    name: "BudgetTable",
    range: { start: { row: 0, col: 0 }, end: { row: 4, col: 3 } },
  });
  core.insertWorkbookChart(assets, 0, {
    family: "line",
    range: { start: { row: 1, col: 0 }, end: { row: 4, col: 3 } },
    sheetName: "Dashboard",
    title: "Budget",
  });

  const exported = saveWithAuthoring({ assets, workbook }, true);
  const archive = unzipSync(exported);
  const sheet = sheetXml(archive);
  const order = [
    "<conditionalFormatting",
    "<dataValidations",
    "<hyperlinks",
    "<drawing",
    "<legacyDrawing",
    "<tableParts",
  ].map((marker) => sheet.indexOf(marker));
  assert.ok(order.every((index) => index >= 0), "every authored element is present");
  assert.deepEqual(
    [...order].sort((left, right) => left - right),
    order,
    "authored elements keep the OOXML worksheet child order",
  );

  const structure = core.parseWorkbookStructureAssets(exported);
  const state = structure.sheetStatesByWorkbookSheetIndex[0];
  assert.deepEqual(state.comments, [{ author: "Dana", cell: commentCell, text: "Note" }]);
  assert.deepEqual(state.hyperlinks, [
    { cell: { row: 4, col: 1 }, target: "https://example.com", tooltip: undefined },
  ]);
  assert.deepEqual(state.conditionalFormatRules.map((rule) => rule.kind), ["dataBar"]);
  assert.deepEqual(state.dataValidations.map((rule) => rule.validationType), ["list"]);
  assert.deepEqual(
    structure.tableMetadataByWorkbookSheetIndex[0].map((table) => table.name),
    ["BudgetTable"],
  );

  const reloaded = wasm.Workbook.fromBytes(exported);
  try {
    const worksheet = reloaded.getSheet(0);
    try {
      assert.deepEqual(worksheet.freezePanes, { row: 1, col: 0 });
      assert.deepEqual(worksheet.tables.map((table) => table.name), ["BudgetTable"]);
    } finally {
      worksheet.free();
    }
    const chartAssets = core.loadWorkbookChartAssets(
      reloaded,
      core.parseWorkbookImageAssets(exported),
      new Map([[0, 0]]),
      false,
    );
    assert.deepEqual(
      chartAssets.chartsByWorkbookSheetIndex[0].map((chart) => chart.chartType),
      ["Line"],
    );
  } finally {
    reloaded.free();
  }
});

test("authoring merge preserves conditional-format rules the editor does not model", () => {
  const bytes = sampleBytes("financial-model.xlsx");
  const archive = unzipSync(bytes);
  archive["xl/worksheets/sheet1.xml"] = strToU8(
    strFromU8(archive["xl/worksheets/sheet1.xml"]).replace(
      "<sheetData>",
      '<conditionalFormatting sqref="B2:B5"><cfRule type="cellIs" operator="greaterThan" priority="9"><formula>10</formula></cfRule></conditionalFormatting><sheetData>',
    ),
  );
  const patched = zipSync(archive);
  const assets = core.parseWorkbookImageAssets(patched);
  const workbook = wasm.Workbook.fromBytes(patched);
  core.addWorkbookConditionalFormatRule(
    assets,
    0,
    core.createConditionalFormatRuleFromDraft(
      { colors: ["5B9BD5"], kind: "dataBar" },
      [{ start: { row: 1, col: 1 }, end: { row: 4, col: 1 } }],
      1,
    ),
  );
  const exported = core.mergeWorkbookSheetAuthoringAssets(workbook.saveXlsxBytes(), assets.archive);
  const xml = sheetXml(unzipSync(exported));
  assert.match(xml, /cfRule type="cellIs"/, "unknown rules are preserved");
  assert.match(xml, /cfRule type="dataBar"/);
  workbook.free();
});

test("authoring merge is a no-op for workbooks without authored objects", () => {
  const { assets, workbook } = loadAssets("financial-model.xlsx");
  const saved = workbook.saveXlsxBytes();
  const merged = core.mergeWorkbookSheetAuthoringAssets(saved, assets.archive);
  assert.deepEqual(
    [...merged],
    [...saved],
    "unmodified workbooks are copied byte for byte",
  );
});
