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

const wasmPath = requireFromXlsxCore.resolve(
  "@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm",
);
const wasmBytes = readFileSync(wasmPath);
core.setWasmSource(
  wasmBytes.buffer.slice(
    wasmBytes.byteOffset,
    wasmBytes.byteOffset + wasmBytes.byteLength,
  ),
);

const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

/**
 * A bold word, a red suffix and a superscript marker — the three-run cell from
 * issue #138 (`<r><rPr><b/></rPr>…`, `<rPr><color rgb=…/>`, `<rPr><vertAlign
 * val="superscript"/>`).
 */
const RICH_TEXT_RUNS_XML = [
  "<r><rPr><b/></rPr><t>Quarter</t></r>",
  '<r><rPr><color rgb="FFFF0000"/></rPr><t>ly</t></r>',
  '<r><rPr><vertAlign val="superscript"/></rPr><t>2</t></r>',
].join("");

const RICH_TEXT_PLAIN = "Quarterly2";

function sampleBytes(name = "charts-images.xlsx") {
  return new Uint8Array(readFileSync(new URL(name, samples)));
}

/**
 * `charts-images.xlsx` with A1 as an inline rich string and B1 as a shared
 * string that carries the same three runs.
 */
function richTextWorkbookBytes() {
  const archive = unzipSync(sampleBytes());
  const sheet = strFromU8(archive["xl/worksheets/sheet1.xml"]);
  const patchedSheet = sheet
    .replace(
      '<c r="A1" s="1" t="inlineStr"><is><t>Quarter</t></is></c>',
      `<c r="A1" s="1" t="inlineStr"><is>${RICH_TEXT_RUNS_XML}</is></c>`,
    )
    .replace(
      '<c r="B1" s="1" t="inlineStr"><is><t>Revenue</t></is></c>',
      '<c r="B1" s="1" t="s"><v>0</v></c>',
    );
  assert.match(patchedSheet, /Quarter/, "the inline fixture cell must be patched");
  assert.match(patchedSheet, /t="s"><v>0<\/v>/, "the shared-string fixture cell must be patched");
  archive["xl/worksheets/sheet1.xml"] = strToU8(patchedSheet);
  archive["xl/sharedStrings.xml"] = strToU8(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">'
      + `<si>${RICH_TEXT_RUNS_XML}</si></sst>`,
  );
  const relationships = strFromU8(archive["xl/_rels/workbook.xml.rels"]).replace(
    "</Relationships>",
    '<Relationship Id="rIdSharedStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
  );
  archive["xl/_rels/workbook.xml.rels"] = strToU8(relationships);
  const contentTypes = strFromU8(archive["[Content_Types].xml"]).replace(
    "</Types>",
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
  );
  archive["[Content_Types].xml"] = strToU8(contentTypes);
  return zipSync(archive);
}

function sheetState(bytes, index = 0) {
  const assets = core.parseWorkbookStructureAssets(bytes);
  return assets.sheetStatesByWorkbookSheetIndex[index];
}

function runTexts(runs) {
  return runs.map((run) => run.text);
}

test("XLSX rich-text cells keep every inline run from shared strings and inline strings", () => {
  const bytes = richTextWorkbookBytes();
  const state = sheetState(bytes);

  const inlineRuns = state.richTextCells.A1;
  assert.ok(inlineRuns, "the inline rich cell keeps its runs");
  assert.deepEqual(runTexts(inlineRuns), ["Quarter", "ly", "2"]);
  assert.equal(inlineRuns[0].font.bold, true);
  assert.deepEqual(inlineRuns[1].font.color, { rgb: "#ff0000" });
  assert.equal(inlineRuns[2].font.verticalAlign, "superscript");
  assert.equal(
    core.richTextRunsToPlainText(inlineRuns),
    RICH_TEXT_PLAIN,
    "the runs concatenate to the display value",
  );

  const sharedRuns = state.richTextCells.B1;
  assert.ok(sharedRuns, "a `t=\"s\"` cell takes its runs from sharedStrings.xml");
  assert.deepEqual(
    sharedRuns,
    inlineRuns,
    "the shared-string path produces the same runs as the inline path",
  );

  assert.equal(
    state.richTextCells.C1,
    undefined,
    "cells without rich runs keep no entry and render their plain value",
  );
});

test("XLSX viewer assets carry rich-text runs for the renderer", () => {
  const assets = core.parseWorkbookImageAssets(richTextWorkbookBytes());
  const state = assets.sheetStatesByWorkbookSheetIndex[0];
  assert.deepEqual(runTexts(state.richTextCells.A1), ["Quarter", "ly", "2"]);
  assert.deepEqual(runTexts(state.richTextCells.B1), ["Quarter", "ly", "2"]);
});

test("XLSX rich-text cells keep the concatenated value for formulas and search", async () => {
  const bytes = richTextWorkbookBytes();
  const module = await core.getSheetsWasmModule();
  const workbook = module.Workbook.fromBytes(bytes);
  const worksheet = workbook.getSheet(0);
  try {
    assert.equal(
      worksheet.getFormattedValueAt(0, 0),
      RICH_TEXT_PLAIN,
      "the engine's display value stays the concatenation of the runs",
    );
    assert.equal(worksheet.getFormattedValueAt(0, 1), RICH_TEXT_PLAIN);
  } finally {
    worksheet.free();
    workbook.free();
  }
});

test("XLSX rich-text runs survive an export and reload", async () => {
  const module = await core.getSheetsWasmModule();
  const workbook = module.Workbook.fromBytes(richTextWorkbookBytes());
  let exported;
  try {
    exported = workbook.saveXlsxBytes();
  } finally {
    workbook.free();
  }

  const archive = unzipSync(exported);
  const savedSheet = strFromU8(archive["xl/worksheets/sheet1.xml"]);
  assert.match(
    savedSheet,
    /<is><r><rPr><b\/><\/rPr><t>Quarter<\/t>/,
    "the saved workbook keeps the bold run",
  );
  assert.match(savedSheet, /<color rgb="FFFF0000"\/>/, "the saved workbook keeps the red run");
  assert.match(savedSheet, /<vertAlign val="superscript"\/>/, "the saved workbook keeps the superscript run");

  const reloaded = sheetState(exported);
  assert.deepEqual(runTexts(reloaded.richTextCells.A1), ["Quarter", "ly", "2"]);
  assert.equal(reloaded.richTextCells.A1[0].font.bold, true);
  assert.equal(reloaded.richTextCells.A1[2].font.verticalAlign, "superscript");
  assert.deepEqual(
    runTexts(reloaded.richTextCells.B1),
    ["Quarter", "ly", "2"],
    "the saved shared-string cell reloads with its runs",
  );
});

test("XLSX rich-text runs are dropped once the cell text no longer matches", () => {
  const runs = [
    { text: "Quarter", font: { bold: true } },
    { text: "ly" },
  ];
  assert.deepEqual(core.richTextRunsToPlainText(runs), "Quarterly");
  assert.equal(
    core.resolveRichTextRuns(runs, "Quarterly"),
    runs,
    "matching runs are returned as-is",
  );
  assert.equal(
    core.resolveRichTextRuns(runs, "Edited"),
    undefined,
    "an edited cell drops the stale runs instead of painting them over the new value",
  );
  assert.equal(
    core.resolveRichTextRuns([{ text: "only" }], "only"),
    undefined,
    "a single run is not rich text",
  );
  assert.equal(core.resolveRichTextRuns(undefined, "plain"), undefined);
  assert.equal(core.resolveRichTextRuns([], ""), undefined);
});

test("XLSX plain shared strings stay plain", () => {
  const archive = unzipSync(sampleBytes("charts-images.xlsx"));
  archive["xl/sharedStrings.xml"] = strToU8(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">'
      + "<si><t>Quarter</t></si></sst>",
  );
  const sheet = strFromU8(archive["xl/worksheets/sheet1.xml"]).replace(
    '<c r="A1" s="1" t="inlineStr"><is><t>Quarter</t></is></c>',
    '<c r="A1" s="1" t="s"><v>0</v></c>',
  );
  archive["xl/worksheets/sheet1.xml"] = strToU8(sheet);
  const state = sheetState(zipSync(archive));
  assert.deepEqual(
    state.richTextCells,
    {},
    "a shared string without runs allocates no rich-text entry",
  );
});
