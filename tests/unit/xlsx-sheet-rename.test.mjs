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

function configureWasm() {
  const wasmBytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  core.setWasmSource(
    wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
  );
  return core.initWasm();
}

const wasm = await configureWasm();

function archiveText(archive, path) {
  return strFromU8(archive[path]);
}

// ── Name rules ───────────────────────────────────────────────────────────────

test("worksheet names follow Excel's trimming, length, character and apostrophe rules", () => {
  assert.equal(core.validateWorksheetName("  Data  "), null, "surrounding whitespace is trimmed");
  assert.equal(core.validateWorksheetName(""), "empty");
  assert.equal(core.validateWorksheetName("   "), "empty");
  assert.equal(core.validateWorksheetName("a".repeat(31)), null);
  assert.equal(core.validateWorksheetName("a".repeat(32)), "too-long");
  assert.equal(core.validateWorksheetName("😀".repeat(31)), null, "length counts code points");
  assert.equal(core.validateWorksheetName("😀".repeat(32)), "too-long");
  assert.equal(core.validateWorksheetName("Q1 [forecast]"), "invalid-characters");
  for (const character of ["\\", "/", ":", "*", "?", "[", "]"]) {
    assert.equal(
      core.validateWorksheetName(`a${character}b`),
      "invalid-characters",
      `${JSON.stringify(character)} is rejected`,
    );
  }
  assert.equal(core.validateWorksheetName("'Data"), "apostrophe");
  assert.equal(core.validateWorksheetName("Data'"), "apostrophe");
  assert.equal(core.validateWorksheetName("Da'ta"), null, "an inner apostrophe is legal");
});

test("worksheet names are unique case-insensitively across the supplied names", () => {
  assert.equal(core.validateWorksheetName("Data", ["Sheet1"]), null);
  assert.equal(core.validateWorksheetName("data", ["Data"]), "duplicate");
  assert.equal(core.validateWorksheetName(" DATA ", ["data"]), "duplicate");
  assert.equal(core.validateWorksheetName("Data", ["", "Data "]), "duplicate");
  assert.equal(core.validateWorksheetName("Data", []), null);
  assert.equal(core.XLSX_MAX_SHEET_NAME_LENGTH, 31);
});

// ── Formula rewriting ────────────────────────────────────────────────────────

test("formula sheet prefixes are renamed and re-quoted the way Excel writes them", () => {
  assert.equal(
    core.renameSheetInFormula("Assumptions!B2*Assumptions!B3", "Assumptions", "Inputs"),
    "Inputs!B2*Inputs!B3",
  );
  assert.equal(
    core.renameSheetInFormula("SUM('My Sheet'!A1:A3)", "My Sheet", "Data"),
    "SUM(Data!A1:A3)",
    "a simple new name drops the quotes",
  );
  assert.equal(
    core.renameSheetInFormula("Sheet1!A1", "Sheet1", "My Inputs"),
    "'My Inputs'!A1",
    "a new name with spaces gains quotes",
  );
  assert.equal(
    core.renameSheetInFormula("Sheet1!A1", "Sheet1", "2024"),
    "'2024'!A1",
    "a name starting with a digit is quoted",
  );
  assert.equal(
    core.renameSheetInFormula("Sheet1!A1", "Sheet1", "A1"),
    "'A1'!A1",
    "a name shaped like a cell reference is quoted",
  );
  assert.equal(
    core.renameSheetInFormula("Sheet1!A1", "Sheet1", "O'Brien"),
    "'O''Brien'!A1",
    "apostrophes double inside a quoted prefix",
  );
  assert.equal(
    core.renameSheetInFormula("Sheet1!A1", "sheet1", "SHEET-2"),
    "'SHEET-2'!A1",
    "the match ignores case",
  );
});

test("formula rewriting leaves literals, errors, tables and other sheets alone", () => {
  assert.equal(
    core.renameSheetInFormula('="Sheet1!A1"&Sheet1!B2', "Sheet1", "Data"),
    '="Sheet1!A1"&Data!B2',
    "string literals keep the old name",
  );
  assert.equal(core.renameSheetInFormula("#REF!+A1", "REF", "Data"), "#REF!+A1");
  assert.equal(core.renameSheetInFormula("Table1[Column]", "Table1", "Data"), "Table1[Column]");
  assert.equal(
    core.renameSheetInFormula("Other!A1+Sheet1!A1", "Sheet1", "Data"),
    "Other!A1+Data!A1",
  );
  assert.equal(
    core.renameSheetInFormula("[Book1.xlsx]Sheet1!A1+Sheet1!A1", "Sheet1", "Data"),
    "[Book1.xlsx]Sheet1!A1+Data!A1",
    "an external workbook keeps its own sheet name",
  );
  assert.equal(
    core.renameSheetInFormula("MySheet1!A1", "Sheet1", "Data"),
    "MySheet1!A1",
    "a longer sheet name is not matched by suffix",
  );
  assert.equal(core.renameSheetInFormula("SUM(A1:A3)", "Sheet1", "Data"), "SUM(A1:A3)");
});

test("3D references rewrite either endpoint", () => {
  assert.equal(
    core.renameSheetInFormula("Sheet1:Sheet3!A1", "Sheet1", "Data"),
    "Data:Sheet3!A1",
  );
  assert.equal(
    core.renameSheetInFormula("Sheet1:Sheet3!A1", "Sheet3", "Data 3"),
    "Sheet1:'Data 3'!A1",
  );
  assert.equal(
    core.renameSheetInFormula("'Sheet 1:Sheet 3'!A1", "Sheet 1", "Data"),
    "'Data:Sheet 3'!A1",
  );
});

// ── Saved-byte rewriting ─────────────────────────────────────────────────────

test("renaming a worksheet rewrites the workbook part and every reference to it", () => {
  const workbook = wasm.Workbook.fromBytes(sampleBytes("financial-model.xlsx"));
  const saved = workbook.saveXlsxBytes();
  workbook.free();

  const renamed = core.applyWorkbookSheetRename(saved, 0, "Inputs & Rates");
  assert.ok(renamed instanceof Uint8Array, "the archive is rewritten");

  const archive = unzipSync(renamed);
  assert.match(archiveText(archive, "xl/workbook.xml"), /<sheet name="Inputs &amp; Rates" sheetId="1"/);
  assert.match(archiveText(archive, "xl/workbook.xml"), /<sheet name="P&amp;L"/);
  const profitAndLoss = archiveText(archive, "xl/worksheets/sheet2.xml");
  assert.doesNotMatch(profitAndLoss, /Assumptions/);
  assert.match(profitAndLoss, /<f>'Inputs &amp; Rates'!B2\*'Inputs &amp; Rates'!B3<\/f>/);

  const reloaded = wasm.Workbook.fromBytes(renamed);
  try {
    assert.deepEqual(reloaded.sheetNames, ["Inputs & Rates", "P&L", "Notes"]);
    assert.equal(core.safeCalculate(reloaded).calculated, true);
    const sheet = reloaded.getSheet(1);
    try {
      assert.equal(sheet.getFormattedValueAt(1, 1), "$50,400", "cross-sheet formulas still calculate");
    } finally {
      sheet.free();
    }
  } finally {
    reloaded.free();
  }
});

test("a rename keeps worksheet order and leaves chartsheets untouched", () => {
  const workbook = wasm.Workbook.fromBytes(sampleBytes("charts-images.xlsx"));
  const saved = workbook.saveXlsxBytes();
  const before = workbook.sheetOrder;
  workbook.free();

  const renamed = core.applyWorkbookSheetRename(saved, 0, "Overview");
  assert.ok(renamed);
  const workbookXml = archiveText(unzipSync(renamed), "xl/workbook.xml");
  assert.match(workbookXml, /<sheet name="Overview" sheetId="1" r:id="rId1"\/><sheet name="Revenue Chart" sheetId="2"/);

  const reloaded = wasm.Workbook.fromBytes(renamed);
  try {
    assert.deepEqual(reloaded.sheetNames, ["Overview"]);
    assert.deepEqual(reloaded.sheetOrder, before, "slot order and chartsheet placement are unchanged");
    assert.equal(reloaded.chartsheets[0]?.name, "Revenue Chart");
  } finally {
    reloaded.free();
  }
});

test("a sheet name with XML-escaped characters still rewrites its references", () => {
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="P&amp;L" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets>
</workbook>`;
  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`;
  const dataSheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1"><f>'P&amp;L'!B2+1</f></c></row></sheetData>
</worksheet>`;
  const bytes = zipSync({
    "xl/workbook.xml": strToU8(workbookXml),
    "xl/_rels/workbook.xml.rels": strToU8(relsXml),
    "xl/worksheets/sheet1.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
    "xl/worksheets/sheet2.xml": strToU8(dataSheetXml),
  });

  const renamed = core.applyWorkbookSheetRename(bytes, 0, "Profit & Loss");
  assert.ok(renamed);
  const archive = unzipSync(renamed);
  assert.match(archiveText(archive, "xl/workbook.xml"), /<sheet name="Profit &amp; Loss" sheetId="1"/);
  assert.match(archiveText(archive, "xl/worksheets/sheet2.xml"), /<f>'Profit &amp; Loss'!B2\+1<\/f>/);

  // A quoted sheet prefix doubles apostrophes, which the part prefilter has to
  // account for as well.
  const apostropheBytes = zipSync({
    "xl/workbook.xml": strToU8(workbookXml.replace("P&amp;L", "O'Brien")),
    "xl/_rels/workbook.xml.rels": strToU8(relsXml),
    "xl/worksheets/sheet1.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`),
    "xl/worksheets/sheet2.xml": strToU8(dataSheetXml.replace("'P&amp;L'", "'O''Brien'")),
  });
  const apostropheRenamed = core.applyWorkbookSheetRename(apostropheBytes, 0, "Owner");
  assert.ok(apostropheRenamed);
  assert.match(
    archiveText(unzipSync(apostropheRenamed), "xl/worksheets/sheet2.xml"),
    /<f>Owner!B2\+1<\/f>/,
  );
});

test("defined names that point at the renamed sheet follow the rename", () => {
  const workbook = wasm.Workbook.fromBytes(sampleBytes("financial-model.xlsx"));
  const saved = core.applyWorkbookDefinedNames(workbook.saveXlsxBytes(), [
    { name: "Rates", refersTo: "Assumptions!$B$2:$D$2" },
    { name: "Notes", refersTo: "Notes!$A$1:$A$3" },
  ]);
  workbook.free();

  const renamed = core.applyWorkbookSheetRename(saved, 0, "Inputs");
  assert.ok(renamed);
  const names = core.readWorkbookDefinedNames(renamed);
  assert.deepEqual(
    names.map((entry) => [entry.name, entry.refersTo]),
    [["Rates", "Inputs!$B$2:$D$2"], ["Notes", "Notes!$A$1:$A$3"]],
  );
});

test("worksheet XML formulas, validations, table formulas and links are rewritten", () => {
  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1"><f>Sheet1!B1+1</f></c></row></sheetData>
  <hyperlinks><hyperlink ref="A1" location="#Sheet1!B2" display="Sheet1!B2"/><hyperlink ref="B1" location="Sheet1!C3"/></hyperlinks>
  <dataValidations count="1"><dataValidation type="list"><formula1>Sheet1!$A$1:$A$5</formula1></dataValidation></dataValidations>
  <conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>Sheet1!A1&gt;0</formula></cfRule></conditionalFormatting>
  <tableParts count="1"><tablePart r:id="rId1"/></tableParts>
</worksheet>`;
  const tableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B2">
  <tableColumns count="2">
    <tableColumn id="1" name="Sheet1"><calculatedColumnFormula>Sheet1!A1*2</calculatedColumnFormula></tableColumn>
    <tableColumn id="2" name="B"/>
  </tableColumns>
  <totalsRowFormula>SUBTOTAL(109,Sheet1!B2:B4)</totalsRowFormula>
</table>`;
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets>
</workbook>`;
  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`;
  const bytes = zipSync({
    "xl/workbook.xml": strToU8(workbookXml),
    "xl/_rels/workbook.xml.rels": strToU8(relsXml),
    "xl/worksheets/sheet1.xml": strToU8(sheetXml),
    "xl/worksheets/sheet2.xml": strToU8(sheetXml),
    "xl/tables/table1.xml": strToU8(tableXml),
  });

  const renamed = core.applyWorkbookSheetRename(bytes, 0, "Data 1");
  assert.ok(renamed);
  const archive = unzipSync(renamed);
  const sheet = archiveText(archive, "xl/worksheets/sheet2.xml");
  assert.match(sheet, /<f>'Data 1'!B1\+1<\/f>/);
  assert.match(sheet, /location="#'Data 1'!B2"/);
  assert.match(sheet, /location="'Data 1'!C3"/);
  assert.match(sheet, /display="Sheet1!B2"/, "display text is not a reference");
  assert.match(sheet, /<formula1>'Data 1'!\$A\$1:\$A\$5<\/formula1>/);
  assert.match(sheet, /<formula>'Data 1'!A1&gt;0<\/formula>/);
  const table = archiveText(archive, "xl/tables/table1.xml");
  assert.match(table, /name="Sheet1"/, "the table column name is untouched");
  assert.match(table, /'Data 1'!A1\*2/, "a calculated column formula is rewritten");
  assert.match(table, /SUBTOTAL\(109,'Data 1'!B2:B4\)/);
});

test("pivot cache worksheet sources and chart series references follow the rename", () => {
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Source" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
  const pivotXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="4">
  <cacheSource type="worksheet"><worksheetSource sheet="Source" ref="A1:C5"/></cacheSource>
  <cacheFields count="1"><cacheField name="Kind" formula="Source!A1"/></cacheFields>
</pivotCacheDefinition>`;
  const chartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Source!$B$2:$B$5</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart>
</c:chartSpace>`;
  const bytes = zipSync({
    "xl/workbook.xml": strToU8(workbookXml),
    "xl/_rels/workbook.xml.rels": strToU8(relsXml),
    "xl/pivotCache/pivotCacheDefinition1.xml": strToU8(pivotXml),
    "xl/charts/chart1.xml": strToU8(chartXml),
  });

  const renamed = core.applyWorkbookSheetRename(bytes, 0, "Raw Data");
  assert.ok(renamed);
  const archive = unzipSync(renamed);
  const pivot = archiveText(archive, "xl/pivotCache/pivotCacheDefinition1.xml");
  assert.match(pivot, /<worksheetSource sheet="Raw Data"/);
  assert.match(pivot, /formula="'Raw Data'!A1"/);
  assert.match(
    archiveText(archive, "xl/charts/chart1.xml"),
    /<c:f>'Raw Data'!\$B\$2:\$B\$5<\/c:f>/,
  );
});

test("renaming is a no-op for unknown sheets and is rejected for invalid input", () => {
  const workbook = wasm.Workbook.fromBytes(sampleBytes("financial-model.xlsx"));
  const saved = workbook.saveXlsxBytes();
  workbook.free();

  assert.equal(core.applyWorkbookSheetRename(saved, 9, "Data"), null);
  assert.equal(core.applyWorkbookSheetRename(saved, -1, "Data"), null);
  assert.equal(core.applyWorkbookSheetRename(saved, 0, ""), null);
  assert.equal(core.applyWorkbookSheetRename(new Uint8Array([1, 2, 3]), 0, "Data"), null);
});
