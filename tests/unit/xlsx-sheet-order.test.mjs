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

function workbookDocument(bytes) {
  return new XmlDomParser().parseFromString(
    archiveText(unzipSync(bytes), "xl/workbook.xml"),
    "application/xml",
  );
}

function sheetNamesInWorkbookXml(bytes) {
  return Array.from(workbookDocument(bytes).getElementsByTagName("*"))
    .filter((node) => node.localName === "sheet")
    .map((node) => node.getAttribute("name"));
}

// ── Saved-byte reorder ───────────────────────────────────────────────────────

test("applyWorkbookSheetOrder rewrites the sheet slot order and keeps the cells with their sheets", () => {
  const workbook = wasm.Workbook.fromBytes(sampleBytes("financial-model.xlsx"));
  const saved = workbook.saveXlsxBytes();
  workbook.free();

  const reordered = core.applyWorkbookSheetOrder(saved, [2, 0, 1]);
  assert.ok(reordered instanceof Uint8Array, "the archive is rewritten");
  assert.deepEqual(sheetNamesInWorkbookXml(reordered), ["Notes", "Assumptions", "P&L"]);

  const reloaded = wasm.Workbook.fromBytes(reordered);
  try {
    assert.deepEqual(reloaded.sheetNames, ["Notes", "Assumptions", "P&L"]);
    assert.deepEqual(reloaded.sheetOrder, [
      { index: 0, slotType: "worksheet" },
      { index: 1, slotType: "worksheet" },
      { index: 2, slotType: "worksheet" },
    ]);
    const notes = reloaded.getSheet(0);
    try {
      assert.equal(notes.name, "Notes");
      assert.match(notes.getFormattedValueAt(0, 0), /Financial model fixture/);
    } finally {
      notes.free();
    }
    const assumptions = reloaded.getSheet(1);
    try {
      assert.equal(assumptions.name, "Assumptions");
      assert.equal(assumptions.getFormattedValueAt(0, 0), "Driver");
    } finally {
      assumptions.free();
    }
    assert.equal(core.safeCalculate(reloaded).calculated, true);
    const profitAndLoss = reloaded.getSheet(2);
    try {
      assert.equal(profitAndLoss.name, "P&L");
      assert.equal(
        profitAndLoss.getFormattedValueAt(1, 1),
        "$50,400",
        "cross-sheet formulas still calculate after the move",
      );
    } finally {
      profitAndLoss.free();
    }
  } finally {
    reloaded.free();
  }
});

test("a reorder keeps the worksheet parts, formulas, images and charts untouched", () => {
  const workbook = wasm.Workbook.fromBytes(sampleBytes("charts-images.xlsx"));
  const saved = workbook.saveXlsxBytes();
  const before = unzipSync(saved);
  workbook.free();

  const reordered = core.applyWorkbookSheetOrder(saved, [1, 0]);
  assert.ok(reordered);
  const after = unzipSync(reordered);
  assert.equal(
    archiveText(after, "xl/worksheets/sheet1.xml"),
    archiveText(before, "xl/worksheets/sheet1.xml"),
    "the worksheet part itself does not change",
  );
  for (const path of Object.keys(before)) {
    if (path === "xl/workbook.xml") continue;
    assert.deepEqual(after[path], before[path], `${path} is carried over byte-for-byte`);
  }

  const reloaded = wasm.Workbook.fromBytes(reordered);
  try {
    assert.deepEqual(reloaded.sheetNames, ["Dashboard"]);
    assert.deepEqual(reloaded.sheetOrder, [
      { index: 0, slotType: "chartsheet" },
      { index: 0, slotType: "worksheet" },
    ]);
    assert.equal(reloaded.chartsheets[0]?.name, "Revenue Chart");
  } finally {
    reloaded.free();
  }
});

test("an identity order returns the input bytes and invalid permutations return null", () => {
  const workbook = wasm.Workbook.fromBytes(sampleBytes("financial-model.xlsx"));
  const saved = workbook.saveXlsxBytes();
  workbook.free();

  assert.equal(core.applyWorkbookSheetOrder(saved, [0, 1, 2]), saved, "identity is a no-op");
  for (const order of [
    [],
    [0, 1],
    [0, 1, 1],
    [0, 1, 3],
    [0, 1, -1],
    [0, 1, 2, 3],
    [0, 1, 1.5],
    ["0", 1, 2],
  ]) {
    assert.equal(
      core.applyWorkbookSheetOrder(saved, order),
      null,
      `${JSON.stringify(order)} is rejected`,
    );
  }
  assert.equal(core.applyWorkbookSheetOrder(new Uint8Array([1, 2, 3]), [0]), null);
  assert.equal(core.applyWorkbookSheetOrder(saved, null), null);
});

test("a reorder remaps sheet-scoped names and workbook-view tab indices", () => {
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <bookViews><workbookView activeTab="2" firstSheet="1"/></bookViews>
  <sheets>
    <sheet name="One" sheetId="1" r:id="rId1"/>
    <sheet name="Two" sheetId="2" r:id="rId2"/>
    <sheet name="Three" sheetId="3" r:id="rId3"/>
  </sheets>
  <definedNames>
    <definedName name="LocalThree" localSheetId="2">Three!$A$1</definedName>
    <definedName name="Global">One!$A$1</definedName>
  </definedNames>
  <extLst>
    <ext uri="{B58B0392-4F1F-4190-BB64-865C28DF51EA}">
      <x14:definedNames>
        <x14:definedName name="LocalTwo" localSheetId="1">Two!$A$1</x14:definedName>
      </x14:definedNames>
    </ext>
  </extLst>
</workbook>`;
  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>`;
  const worksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`;
  const bytes = zipSync({
    "xl/workbook.xml": strToU8(workbookXml),
    "xl/_rels/workbook.xml.rels": strToU8(relsXml),
    "xl/worksheets/sheet1.xml": strToU8(worksheet),
    "xl/worksheets/sheet2.xml": strToU8(worksheet),
    "xl/worksheets/sheet3.xml": strToU8(worksheet),
  });

  const reordered = core.applyWorkbookSheetOrder(bytes, [2, 0, 1]);
  assert.ok(reordered);
  const document = new XmlDomParser().parseFromString(
    archiveText(unzipSync(reordered), "xl/workbook.xml"),
    "application/xml",
  );
  const elements = Array.from(document.getElementsByTagName("*"));
  const byLocalName = (name) => elements.filter((node) => node.localName === name);

  assert.deepEqual(
    byLocalName("sheet").map((node) => node.getAttribute("name")),
    ["Three", "One", "Two"],
  );
  const workbookView = byLocalName("workbookView")[0];
  assert.equal(workbookView.getAttribute("activeTab"), "0", "old slot 2 is now slot 0");
  assert.equal(workbookView.getAttribute("firstSheet"), "2", "old slot 1 is now slot 2");
  const definedNames = byLocalName("definedName");
  assert.equal(
    definedNames.find((node) => node.getAttribute("name") === "LocalThree")?.getAttribute("localSheetId"),
    "0",
  );
  assert.equal(
    definedNames.find((node) => node.getAttribute("name") === "LocalTwo")?.getAttribute("localSheetId"),
    "2",
  );
  assert.equal(
    definedNames.find((node) => node.getAttribute("name") === "Global")?.getAttribute("localSheetId"),
    null,
    "a workbook-scoped name keeps no local sheet id",
  );
});

test("a chartsheet-first workbook keeps worksheet assets indexed by worksheet", () => {
  const source = sampleBytes("charts-images.xlsx");
  const workbook = wasm.Workbook.fromBytes(source);
  const saved = workbook.saveXlsxBytes();
  workbook.free();

  const assets = core.parseWorkbookImageAssets(source);
  const merged = core.mergeWorkbookImageAssets(
    core.mergeWorkbookSheetAuthoringAssets(saved, assets.archive),
    assets,
    assets.sheetOrigins,
  );
  const reordered = core.applyWorkbookSheetOrder(merged, [1, 0]);
  assert.ok(reordered);

  const parsed = core.parseWorkbookImageAssets(reordered);
  assert.deepEqual(
    parsed.sheetPathsByWorkbookSheetIndex,
    ["xl/worksheets/sheet1.xml"],
    "chartsheets are not worksheet slots",
  );
  assert.deepEqual(parsed.imagesByWorkbookSheetIndex.map((images) => images.length), [1]);
  assert.deepEqual(parsed.shapesByWorkbookSheetIndex.map((shapes) => shapes.length), [1]);
  assert.equal(parsed.sheetOrigins[0]?.attachments.length, 1, "the drawing stays on the worksheet");
  assert.equal(
    core.parseWorkbookStructureAssets(reordered).sheetStatesByWorkbookSheetIndex.length,
    1,
    "sheet state is indexed by worksheet, not by tab slot",
  );
});

test("moving a sheet and moving it back restores the original workbook order", () => {
  const workbook = wasm.Workbook.fromBytes(sampleBytes("financial-model.xlsx"));
  const saved = workbook.saveXlsxBytes();
  workbook.free();

  const moved = core.applyWorkbookSheetOrder(saved, [1, 2, 0]);
  assert.ok(moved);
  assert.deepEqual(sheetNamesInWorkbookXml(moved), ["P&L", "Notes", "Assumptions"]);

  const back = core.applyWorkbookSheetOrder(moved, [2, 0, 1]);
  assert.ok(back);
  assert.deepEqual(sheetNamesInWorkbookXml(back), ["Assumptions", "P&L", "Notes"]);

  const reloaded = wasm.Workbook.fromBytes(back);
  try {
    assert.deepEqual(reloaded.sheetNames, ["Assumptions", "P&L", "Notes"]);
  } finally {
    reloaded.free();
  }
});
