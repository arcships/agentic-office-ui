import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

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

const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);

const insertRow = (index, count = 1) => ({ axis: "row", count, index, kind: "insert" });
const deleteRow = (index, count = 1) => ({ axis: "row", count, index, kind: "delete" });
const insertColumn = (index, count = 1) => ({ axis: "column", count, index, kind: "insert" });
const deleteColumn = (index, count = 1) => ({ axis: "column", count, index, kind: "delete" });

function sampleBytes(name = "charts-images.xlsx") {
  return new Uint8Array(
    readFileSync(new URL(`../../apps/demo/public/samples/${name}`, import.meta.url)),
  );
}

test("axis indices shift on insert and disappear on delete", () => {
  assert.equal(core.remapAxisIndexForStructure(2, insertRow(3)), 2);
  assert.equal(core.remapAxisIndexForStructure(3, insertRow(3)), 4);
  assert.equal(core.remapAxisIndexForStructure(4, insertRow(3, 2)), 6);
  assert.equal(core.remapAxisIndexForStructure(2, deleteRow(3, 2)), 2);
  assert.equal(core.remapAxisIndexForStructure(3, deleteRow(3, 2)), null);
  assert.equal(core.remapAxisIndexForStructure(4, deleteRow(3, 2)), null);
  assert.equal(core.remapAxisIndexForStructure(5, deleteRow(3, 2)), 3);
  assert.equal(core.remapAxisIndexForStructure(1, insertColumn(1)), 2);
});

test("ranges expand on insert, shrink on delete and disappear when fully removed", () => {
  const range = { start: { row: 2, col: 1 }, end: { row: 5, col: 3 } };
  assert.deepEqual(
    core.remapRangeForStructure(range, insertRow(3)),
    { start: { row: 2, col: 1 }, end: { row: 6, col: 3 } },
    "a range crossing the insertion point grows",
  );
  assert.deepEqual(
    core.remapRangeForStructure(range, insertRow(2)),
    { start: { row: 3, col: 1 }, end: { row: 6, col: 3 } },
    "a range starting at the insertion point moves down",
  );
  assert.deepEqual(
    core.remapRangeForStructure(range, deleteRow(3, 1)),
    { start: { row: 2, col: 1 }, end: { row: 4, col: 3 } },
  );
  assert.deepEqual(
    core.remapRangeForStructure(range, deleteRow(2, 2)),
    { start: { row: 2, col: 1 }, end: { row: 3, col: 3 } },
  );
  assert.equal(core.remapRangeForStructure(range, deleteRow(2, 4)), null);
  assert.deepEqual(
    core.remapRangeForStructure(range, deleteColumn(1, 1)),
    { start: { row: 2, col: 1 }, end: { row: 5, col: 2 } },
  );
  assert.deepEqual(
    core.remapCellForStructure({ row: 4, col: 2 }, deleteRow(4, 1)),
    null,
  );
});

test("A1 references follow Excel semantics for insertions, deletions, anchors and sheets", () => {
  assert.equal(core.remapA1References("=SUM(A1:A10)", insertRow(4)), "=SUM(A1:A11)");
  assert.equal(core.remapA1References("=SUM(A1:A10)", deleteRow(1, 1)), "=SUM(A1:A9)");
  assert.equal(core.remapA1References("=A1+$B$2", insertRow(0, 2)), "=A3+$B$4");
  assert.equal(core.remapA1References("=A5", deleteRow(4, 1)), "=#REF!");
  assert.equal(core.remapA1References("=A5:B8", deleteRow(4, 1)), "=A5:B7");
  assert.equal(core.remapA1References("=A5:A6", deleteRow(4, 2)), "=#REF!");
  assert.equal(
    core.remapA1References("=A2+'Other Sheet'!B2", insertRow(0), { sheetName: "Data" }),
    "=A3+'Other Sheet'!B2",
    "references to other sheets are untouched",
  );
  assert.equal(
    core.remapA1References("='Data'!B2", insertRow(0), { sheetName: "Data" }),
    "='Data'!B3",
    "quoted references to the edited sheet move",
  );
  assert.equal(
    core.remapA1References("=A5+'Data'!A5", insertRow(0), { sheetName: "Data", skipUnqualified: true }),
    "=A5+'Data'!A6",
    "another sheet's unqualified references stay put while qualified ones move",
  );
  assert.equal(
    core.remapA1References('=IF(A1="B2",LOG10(A1),A1)', insertRow(0)),
    '=IF(A2="B2",LOG10(A2),A2)',
    "strings and function names that look like addresses are untouched",
  );
  assert.equal(
    core.remapDefinedNameForStructure("'Data'!$A$1:$C$3", insertRow(1), "Data"),
    "'Data'!$A$1:$C$4",
  );
});

test("freeze panes and drawing markers follow the structure edit", () => {
  assert.deepEqual(core.remapFreezePanesForStructure({ col: 1, row: 2 }, insertRow(1)), { col: 1, row: 3 });
  assert.deepEqual(core.remapFreezePanesForStructure({ col: 1, row: 2 }, deleteRow(1)), { col: 1, row: 1 });
  assert.deepEqual(core.remapFreezePanesForStructure({ col: 1, row: 2 }, deleteRow(3)), { col: 1, row: 2 });
  assert.deepEqual(core.remapFreezePanesForStructure({ col: 1, row: 2 }, deleteRow(0, 5)), { col: 1, row: 0 });

  const anchor = {
    from: { col: 1, colOffsetEmu: 5, row: 2, rowOffsetEmu: 7 },
    kind: "one-cell",
    sizeEmu: { cx: 100, cy: 200 },
  };
  assert.deepEqual(core.remapImageAnchorForStructure(anchor, insertRow(0)), {
    from: { col: 1, colOffsetEmu: 5, row: 3, rowOffsetEmu: 7 },
    kind: "one-cell",
    sizeEmu: { cx: 100, cy: 200 },
  });
  assert.deepEqual(core.remapImageAnchorForStructure(anchor, deleteRow(2)), {
    from: { col: 1, colOffsetEmu: 5, row: 2, rowOffsetEmu: 7 },
    kind: "one-cell",
    sizeEmu: { cx: 100, cy: 200 },
  });
});

test("conditional formats, validations, sparklines and tables remap their references", () => {
  const rule = {
    cfvos: [{ type: "min" }, { type: "formula", value: "A5" }, { type: "max" }],
    colors: [{ rgb: "#f8696b" }, { rgb: "#ffeb84" }, { rgb: "#63be7b" }],
    kind: "colorScale",
    priority: 1,
    ranges: [{ start: { row: 1, col: 1 }, end: { row: 4, col: 1 } }],
  };
  const nextRule = core.remapConditionalFormatRuleForStructure(rule, insertRow(2));
  assert.deepEqual(nextRule.ranges, [{ start: { row: 1, col: 1 }, end: { row: 5, col: 1 } }]);
  assert.equal(nextRule.cfvos[1].value, "A6");
  assert.equal(core.remapConditionalFormatRuleForStructure(rule, deleteRow(1, 4)), null);

  const validation = {
    listSource: "A1:A5",
    ranges: [{ start: { row: 3, col: 0 }, end: { row: 3, col: 0 } }],
    validationType: "list",
    value1: "A5",
  };
  const nextValidation = core.remapDataValidationForStructure(validation, insertRow(1));
  assert.equal(nextValidation.listSource, "A1:A6");
  assert.equal(nextValidation.value1, "A6");
  assert.deepEqual(nextValidation.ranges, [{ start: { row: 4, col: 0 }, end: { row: 4, col: 0 } }]);
  assert.equal(core.remapDataValidationForStructure(validation, deleteRow(3)), null);

  const sparkline = {
    range: { start: { row: 1, col: 0 }, end: { row: 4, col: 0 } },
    target: { row: 5, col: 1 },
    type: "line",
  };
  assert.deepEqual(core.remapSparklineForStructure(sparkline, insertRow(2)), {
    range: { start: { row: 1, col: 0 }, end: { row: 5, col: 0 } },
    target: { row: 6, col: 1 },
    type: "line",
  });
  assert.equal(core.remapSparklineForStructure(sparkline, deleteRow(5)), null);

  assert.deepEqual(
    core.remapTableForStructure({ start: { row: 0, col: 0 }, end: { row: 9, col: 2 } }, insertRow(4)),
    { end: { row: 10, col: 2 }, start: { row: 0, col: 0 } },
  );
});

test("shiftWorkbookSheetStructure moves links, notes, ranges and tables in the archive", () => {
  const assets = core.parseWorkbookImageAssets(sampleBytes());
  assert.equal(core.setWorkbookCellHyperlink(
    assets,
    0,
    { row: 11, col: 0 },
    { target: "#Dashboard!A12" },
  )?.length, 2);
  assert.equal(core.setWorkbookCellComment(
    assets,
    0,
    { row: 11, col: 0 },
    { author: "Dana", text: "row note" },
  )?.length, 2);
  assert.equal(core.addWorkbookConditionalFormatRule(assets, 0, {
    cfvos: [{ type: "min" }, { type: "formula", value: "A12" }, { type: "max" }],
    colors: [{ rgb: "#f8696b" }, { rgb: "#ffeb84" }, { rgb: "#63be7b" }],
    kind: "colorScale",
    priority: 4,
    ranges: [{ start: { row: 10, col: 0 }, end: { row: 12, col: 0 } }],
  })?.length, 4);

  assert.equal(
    core.shiftWorkbookSheetStructure(assets, 0, insertRow(11), { sheetName: "Dashboard" }),
    true,
  );
  const state = core.readWorkbookSheetAuthoringState(assets, 0);
  assert.deepEqual(
    state.hyperlinks.map((hyperlink) => [hyperlink.cell.row, hyperlink.target]),
    [[1, "#Dashboard!A7"], [12, "#Dashboard!A13"]],
    "the link cell and its internal target both move",
  );
  assert.deepEqual(
    state.comments.map((comment) => [comment.cell.row, comment.text]),
    [[1, "该季度收入已通过财务复核。"], [12, "row note"]],
  );
  const rule = state.conditionalFormatRules.at(-1);
  assert.deepEqual(rule.ranges, [{ start: { row: 10, col: 0 }, end: { row: 13, col: 0 } }]);
  assert.equal(rule.cfvos[1].value, "A13");

  const tableAssets = core.parseWorkbookImageAssets(sampleBytes("sales-table.xlsx"));
  core.shiftWorkbookSheetStructure(tableAssets, 0, insertRow(5, 2), { sheetName: "Sales Data" });
  const tableState = core.readWorkbookSheetAuthoringState(tableAssets, 0);
  assert.equal(tableState.tables[0].reference, "A1:H183");
  core.shiftWorkbookSheetStructure(tableAssets, 0, deleteRow(5, 2), { sheetName: "Sales Data" });
  assert.equal(core.readWorkbookSheetAuthoringState(tableAssets, 0).tables[0].reference, "A1:H181");
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

function sheetXml(assets) {
  return strFromU8(assets.archive["xl/worksheets/sheet1.xml"]);
}

test("unmodeled conditional-format rules shift their sqref and formula", () => {
  const assets = core.parseWorkbookImageAssets(workbookWithUnmodeledRule(sampleBytes()));
  assert.match(sheetXml(assets), /sqref="A11:A12"/);
  assert.equal(
    core.shiftWorkbookSheetStructure(assets, 0, insertRow(11), { sheetName: "Dashboard" }),
    true,
  );
  assert.match(sheetXml(assets), /conditionalFormatting sqref="A11:A13"/);
  assert.match(sheetXml(assets), /<formula>\$A11&gt;0<\/formula>/);

  assert.equal(
    core.shiftWorkbookSheetStructure(assets, 0, deleteRow(10, 3), { sheetName: "Dashboard" }),
    true,
  );
  assert.doesNotMatch(sheetXml(assets), /type="expression"/, "a rule whose range is deleted disappears");
});

test("defined names round-trip through saved workbook bytes", () => {
  const names = [
    { name: "Revenue", refersTo: "'Dashboard'!$A$1:$B$2" },
    { name: "LocalRate", refersTo: "'Dashboard'!$C$1", sheetIndex: 0 },
  ];
  const bytes = core.applyWorkbookDefinedNames(sampleBytes(), names);
  assert.deepEqual(core.readWorkbookDefinedNames(bytes), [
    { comment: undefined, hidden: false, name: "Revenue", refersTo: "'Dashboard'!$A$1:$B$2", sheetIndex: undefined },
    { comment: undefined, hidden: false, name: "LocalRate", refersTo: "'Dashboard'!$C$1", sheetIndex: 0 },
  ]);

  const normalized = core.normalizeWorkbookDefinedNames([
    { name: "Revenue", refersTo: "'Dashboard'!$A$1", scope: "workbook" },
    { name: "Local", refersTo: "'Dashboard'!$B$1", sheetIndex: 0, comment: "note", hidden: true },
    { name: "", refersTo: "A1" },
  ]);
  assert.deepEqual(normalized, [
    { comment: undefined, hidden: false, name: "Revenue", refersTo: "'Dashboard'!$A$1", sheetIndex: undefined },
    { comment: "note", hidden: true, name: "Local", refersTo: "'Dashboard'!$B$1", sheetIndex: 0 },
  ]);
});
