import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

// Issue #22 acceptance: a workbook that names a built-in table style must paint
// the same banded rows, header fill and total row Excel does. The resolver is
// shared by the canvas grid, the DOM grid and the print preview; this test
// covers the style algebra and the workbook export/reload path.

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
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

const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

function sampleBytes(name = "sales-table.xlsx") {
  return new Uint8Array(readFileSync(new URL(name, samples)));
}

function makeTable(overrides = {}) {
  return {
    columns: [
      { id: 1, index: 0, name: "Region" },
      { id: 2, index: 1, name: "Units" },
      { id: 3, index: 2, name: "Total" },
    ],
    displayName: "Sales",
    end: { col: 2, row: 5 },
    headerRowCount: 1,
    name: "Sales",
    reference: "A1:C6",
    start: { col: 0, row: 0 },
    styleInfo: {
      name: "TableStyleMedium2",
      showFirstColumn: false,
      showLastColumn: false,
      showRowStripes: true,
    },
    totalsRowCount: 0,
    totalsRowShown: false,
    ...overrides,
  };
}

function resolveCellStyle(table, cell, extra = {}) {
  return core.withTableCellStyle(undefined, [table], cell, {
    tableStyleByName: {},
    namedCellStyleByName: {},
    ...extra,
  });
}

test("all 60 built-in table styles resolve, unknown names do not", () => {
  for (const [family, count] of [["Light", 21], ["Medium", 28], ["Dark", 11]]) {
    for (let index = 1; index <= count; index += 1) {
      const style = core.resolveBuiltinTableStyle(`TableStyle${family}${index}`);
      assert.ok(style, `TableStyle${family}${index}`);
      assert.equal(typeof style, "object");
    }
  }
  assert.equal(core.resolveBuiltinTableStyle("TableStyleMedium999"), null);
  assert.equal(core.resolveBuiltinTableStyle(undefined), null);
  assert.equal(core.resolveBuiltinTableStyle("MyCustomStyle"), null);
});

test("header rows, banded body rows and totals rows resolve like Excel", () => {
  const table = makeTable({
    end: { col: 2, row: 5 },
    totalsRowCount: 1,
    totalsRowShown: true,
  });

  const header = resolveCellStyle(table, { col: 0, row: 0 });
  assert.deepEqual(header.fill, { color: { theme: 4 }, fillType: "solid" });
  assert.equal(header.font.bold, true);
  assert.deepEqual(header.font.color, { theme: 0 });

  // The first body row is the first stripe, the second row the second stripe.
  const firstStripe = resolveCellStyle(table, { col: 0, row: 1 });
  assert.deepEqual(firstStripe.fill, { color: { theme: 4, tint: 0.8 }, fillType: "solid" });
  const secondStripe = resolveCellStyle(table, { col: 0, row: 2 });
  assert.equal(secondStripe.fill, undefined, "TableStyleMedium2's second stripe is unfilled");

  // The totals row is emphasised and gets the double top border.
  const totals = resolveCellStyle(table, { col: 0, row: 5 });
  assert.equal(totals.font.bold, true);
  assert.deepEqual(totals.border.top, { color: { theme: 4 }, style: "double" });
});

test("column stripes, first/last column emphasis and style flags are honoured", () => {
  const table = makeTable({
    styleInfo: {
      name: "TableStyleMedium2",
      showColumnStripes: true,
      showFirstColumn: true,
      showLastColumn: true,
      showRowStripes: false,
    },
  });
  const firstBodyCell = resolveCellStyle(table, { col: 0, row: 1 });
  // First column emphasis wins over the column stripe at the same cell.
  assert.equal(firstBodyCell.font.bold, true);
  const middleColumn = resolveCellStyle(table, { col: 2, row: 2 });
  assert.deepEqual(middleColumn.fill, {
    color: { theme: 4, tint: 0.8 },
    fillType: "solid",
  });
  const lastColumn = resolveCellStyle(table, { col: 2, row: 1 });
  assert.equal(lastColumn.font.bold, true);

  const plain = resolveCellStyle(
    makeTable({
      styleInfo: {
        name: "TableStyleMedium2",
        showColumnStripes: false,
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: false,
      },
    }),
    { col: 1, row: 2 },
  );
  assert.equal(plain.fill, undefined);
  assert.equal(plain.font.bold, undefined);
});

test("a workbook-defined table style overrides the shipped preset", () => {
  const table = makeTable();
  const custom = {
    headerRow: {
      fill: { color: { rgb: "#112233" }, fillType: "solid" },
      font: { bold: true, color: { rgb: "#ffffff" } },
    },
    firstRowStripe: { fill: { color: { rgb: "#eeeeee" }, fillType: "solid" } },
  };
  const header = resolveCellStyle(table, { col: 0, row: 0 }, {
    tableStyleByName: { TableStyleMedium2: custom },
  });
  assert.deepEqual(header.fill, { color: { rgb: "#112233" }, fillType: "solid" });
  const stripe = resolveCellStyle(table, { col: 0, row: 1 }, {
    tableStyleByName: { TableStyleMedium2: custom },
  });
  assert.deepEqual(stripe.fill, { color: { rgb: "#eeeeee" }, fillType: "solid" });
});

test("explicit cell formatting wins over the table style", () => {
  const table = makeTable();
  const cellStyle = {
    fill: { color: { rgb: "#ff0000" }, fillType: "solid" },
    font: { bold: false, color: { rgb: "#00ff00" }, italic: true, name: "Georgia", size: 14 },
    border: {
      bottom: { color: { rgb: "#0000ff" }, style: "thick" },
      top: { style: "none" },
    },
  };
  const header = core.withTableCellStyle(cellStyle, [table], { col: 0, row: 0 }, {
    declaredTableStyleOverrides: {
      border: cellStyle.border,
      fill: cellStyle.fill,
      font: cellStyle.font,
    },
  });
  assert.deepEqual(header.fill, { color: { rgb: "#ff0000" }, fillType: "solid" });
  assert.equal(header.font.italic, true);
  assert.equal(header.font.bold, false);
  assert.equal(header.font.name, "Georgia");
  assert.deepEqual(header.font.color, { rgb: "#00ff00" });
  assert.deepEqual(header.border.bottom, { color: { rgb: "#0000ff" }, style: "thick" });
});

test("default cell fonts do not erase the table header emphasis", () => {
  const table = makeTable();
  const defaultStyle = {
    font: {
      bold: false,
      color: { theme: 1 },
      italic: false,
      name: "Calibri",
      size: 11,
      strikethrough: false,
      underline: "none",
      verticalAlign: "baseline",
    },
    numberFormat: { formatType: "general" },
  };
  const header = core.withTableCellStyle(defaultStyle, [table], { col: 0, row: 0 });
  assert.equal(header.font.bold, true, "the table header stays bold");
  assert.deepEqual(header.font.color, { theme: 0 }, "the table header stays white");
  assert.deepEqual(header.numberFormat, { formatType: "general" });
});

test("sheet tables are found by cell and table-free sheets short-circuit", () => {
  const table = makeTable();
  assert.equal(core.findTableAtCell([table], { col: 1, row: 3 })?.name, "Sales");
  assert.equal(core.findTableAtCell([table], { col: 3, row: 3 }), null);
  assert.equal(core.findTableAtCell([], { col: 0, row: 0 }), null);
  const base = { font: { bold: true } };
  assert.equal(core.withTableCellStyle(base, [], { col: 0, row: 0 }), base);
  assert.equal(core.withTableCellStyle(base, [table], { col: 9, row: 9 }), base);
});

test("the sample workbook carries its table style through export and reload", () => {
  const bytes = sampleBytes("sales-table.xlsx");
  const assets = core.parseWorkbookImageAssets(bytes);
  const table = assets.tableMetadataByWorkbookSheetIndex[0][0];
  assert.equal(table.styleInfo.name, "TableStyleMedium2");
  assert.equal(table.styleInfo.showRowStripes, true);

  const workbook = wasm.Workbook.fromBytes(bytes);
  const saved = core.mergeWorkbookSheetAuthoringAssets(
    workbook.saveXlsxBytes(),
    assets.archive,
  );
  const reloaded = core.parseWorkbookStructureAssets(saved);
  const reloadedTable = reloaded.tableMetadataByWorkbookSheetIndex[0][0];
  assert.equal(reloadedTable.styleInfo.name, "TableStyleMedium2");
  assert.equal(reloadedTable.styleInfo.showRowStripes, true);
  assert.equal(reloadedTable.reference, table.reference);
  assert.equal(reloaded.sheetStatesByWorkbookSheetIndex[0].definedNames?.length, 0);

  const style = core.resolveBuiltinTableStyle(reloadedTable.styleInfo.name);
  assert.equal(
    core.resolveWorkbookFillStyle(style.headerRow.fill, reloaded.themePalette).backgroundColor,
    "#4f81bd",
  );
});
