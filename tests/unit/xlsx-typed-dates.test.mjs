import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const ts = requireFromXlsxCore("typescript");
const { strToU8, zipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);

const typedDatesModuleUrl = (() => {
  const source = readFileSync(
    new URL("../../packages/xlsx-core/src/typed-dates.ts", import.meta.url),
    "utf8",
  );
  const fflateUrl = pathToFileURL(requireFromXlsxCore.resolve("fflate")).href;
  const emitted = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText.replace(/from\s+["']fflate["']/g, `from ${JSON.stringify(fflateUrl)}`);
  return `data:text/javascript;base64,${Buffer.from(emitted).toString("base64")}`;
})();

const {
  normalizeTypedDateCells,
  parseIsoDateCell,
} = await import(typedDatesModuleUrl);

const wasmEntry = pathToFileURL(
  requireFromXlsxCore.resolve("@dukelib/sheets-wasm"),
).href;
const wasmPath = requireFromXlsxCore.resolve(
  "@dukelib/sheets-wasm/duke_sheets_wasm_bg.wasm",
);
const { default: initWasm, Workbook } = await import(wasmEntry);
await initWasm({ module_or_path: readFileSync(wasmPath) });

function dateWorkbook(date1904) {
  const parts = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    </Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`,
    "xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <workbookPr date1904="${date1904 ? 1 : 0}"/>
      <sheets><sheet name="Dates" sheetId="1" r:id="rId1"/></sheets>
    </workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    </Relationships>`,
    "xl/styles.xml": `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <numFmts count="3">
        <numFmt numFmtId="164" formatCode="dd mmm yyyy"/>
        <numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss.000"/>
        <numFmt numFmtId="166" formatCode="hh:mm:ss.000"/>
      </numFmts>
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="4">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
      </cellXfs>
    </styleSheet>`,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
      <row r="1"><c r="A1" t="d" s="1"><v>2026-11-01T00:00:00.000Z</v></c></row>
      <row r="2"><c r="A2" t="d" s="1"><v>2026-11-11</v></c></row>
      <row r="3"><c r="A3" t="d" s="2"><v>2026-11-01T12:34:56.125</v></c></row>
      <row r="4"><c r="A4" t="d" s="3"><v>12:34:56.125</v></c></row>
      <row r="5"><c r="A5" t="inlineStr" s="1"><is><t>2026-11-01T00:00:00.000Z</t></is></c></row>
      <row r="6"><c r="A6" t="d" s="1"><v>2026-02-30T00:00:00Z</v></c></row>
      <row r="7"><c r="A7" t="d" s="1"><f>DATE(2026,11,1)</f><v>2026-11-01T00:00:00Z</v></c></row>
      <row r="8"><c r="A8" t="n" s="1"><v>${date1904 ? 44865 : 46327}</v></c></row>
      <row r="9"><c r="A9" t="d"><v>1900-01-01</v></c></row>
      <row r="10"><c r="A10" t="d"><v>1900-02-28</v></c></row>
      <row r="11"><c r="A11" t="d"><v>1900-03-01</v></c></row>
      <row r="12"><c r="A12" t="d"><v>1904-01-01</v></c></row>
      <row r="13"><c r="A13" t="d"><v>1900-02-29</v></c></row>
      <row r="14"><c r="A14" t="d"><v>23:59:60</v></c></row>
      <row r="15"><c r="A15" t="d"><v>2026-11-01T00:00:00+02:00</v></c></row>
    </sheetData></worksheet>`,
  };
  return zipSync(
    Object.fromEntries(Object.entries(parts).map(([name, xml]) => [name, strToU8(xml)])),
  );
}

test("ISO date cells convert to serials in both Excel date systems", () => {
  assert.equal(parseIsoDateCell("2026-11-01T00:00:00.000Z"), 46327);
  assert.equal(parseIsoDateCell("2026-11-11"), 46337);
  assert.equal(parseIsoDateCell("2026-11-01T00:00:00.000Z", true), 44865);
  assert.equal(parseIsoDateCell("1904-01-01", true), 0);
  assert.equal(parseIsoDateCell("1900-01-01"), 1);
  assert.equal(parseIsoDateCell("1900-02-28"), 59);
  assert.equal(parseIsoDateCell("1900-03-01"), 61);
  assert.equal(parseIsoDateCell("1900-02-29"), null);
  assert.equal(parseIsoDateCell("23:59:60"), null);
  assert.equal(parseIsoDateCell("2026-11-01T00:00:00+02:00"), null);
  assert.ok(
    Math.abs(parseIsoDateCell("12:34:56.125") - 45296.125 / 86_400) < 1e-12,
  );
});

for (const date1904 of [false, true]) {
  test(`typed ISO date workbook imports as numeric cells (${date1904 ? "1904" : "1900"} system)`, () => {
    const source = dateWorkbook(date1904);
    const normalized = normalizeTypedDateCells(source);
    assert.notEqual(normalized, source, "typed date cells should be rewritten");

    const workbook = Workbook.fromBytes(normalized);
    try {
      const sheet = workbook.getSheet(0);
      try {
        const serial = date1904 ? 44865 : 46327;
        const expected = [serial, serial + 10, serial + 45296.125 / 86_400, 45296.125 / 86_400];
        for (const [row, value] of expected.entries()) {
          const cell = sheet.getCellAt(row, 0);
          try {
            assert.equal(cell.is_number, true, `A${row + 1} should be numeric`);
            assert.ok(Math.abs(cell.asNumber() - value) < 1e-9);
          } finally {
            cell.free();
          }
        }

        assert.equal(sheet.getFormattedValue("A1"), "01 Nov 2026");
        assert.equal(sheet.getFormattedValue("A2"), "11 Nov 2026");
        assert.equal(sheet.getFormattedValue("A3"), "2026-11-01 12:34:56.125");
        assert.equal(sheet.getFormattedValue("A4"), "12:34:56.125");
        assert.equal(sheet.getFormattedValue("A6"), "2026-02-30T00:00:00Z");
        assert.equal(sheet.getFormattedValue("A7"), "01 Nov 2026");
        assert.equal(sheet.getFormulaAt(6, 0), "=DATE(2026,11,1)");
        assert.equal(sheet.getFormattedValue("A8"), "01 Nov 2026");
        assert.equal(sheet.getFormattedValue("A12"), date1904 ? "0" : "1462");
        assert.equal(sheet.getFormattedValue("A13"), "1900-02-29");
        assert.equal(sheet.getFormattedValue("A14"), "23:59:60");
        assert.equal(sheet.getFormattedValue("A15"), "2026-11-01T00:00:00+02:00");
        if (!date1904) {
          assert.equal(sheet.getFormattedValue("A9"), "1");
          assert.equal(sheet.getFormattedValue("A10"), "59");
          assert.equal(sheet.getFormattedValue("A11"), "61");
        }
      } finally {
        sheet.free();
      }
    } finally {
      workbook.free();
    }
  });
}

test("typed date normalization leaves non-archive inputs untouched", () => {
  const plain = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3]);
  assert.equal(normalizeTypedDateCells(plain), plain);
  assert.equal(normalizeTypedDateCells(new Uint8Array([1, 2, 3])).byteLength, 3);
});
