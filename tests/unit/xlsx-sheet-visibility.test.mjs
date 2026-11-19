import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
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
globalThis.Document ??= class Document {};
globalThis.ShadowRoot ??= class ShadowRoot {};

const core = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);
const sheetsWasm = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@dukelib/sheets-wasm")).href
);
const wasmBytes = readFileSync(
  new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
);
await sheetsWasm.default({
  module_or_path: wasmBytes.buffer.slice(
    wasmBytes.byteOffset,
    wasmBytes.byteOffset + wasmBytes.byteLength,
  ),
});

function workbookBytes(name) {
  return new Uint8Array(readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  ));
}

function saveWorkbook(bytes) {
  const workbook = sheetsWasm.Workbook.fromBytes(bytes);
  try {
    return workbook.saveXlsxBytes();
  } finally {
    workbook.free();
  }
}

function visibilityByName(bytes) {
  const workbook = sheetsWasm.Workbook.fromBytes(bytes);
  try {
    return Object.fromEntries(workbook.sheetNames.map((name, index) => {
      const worksheet = workbook.getSheet(index);
      try {
        return [name, worksheet.visibility];
      } finally {
        worksheet.free();
      }
    }));
  } finally {
    workbook.free();
  }
}

test("applyWorkbookSheetVisibility writes hidden, veryHidden and visible states", () => {
  const saved = saveWorkbook(workbookBytes("financial-model.xlsx"));
  const hidden = core.applyWorkbookSheetVisibility(saved, "P&L", "hidden");
  assert.ok(hidden, "hiding a worksheet rewrites the saved package");
  assert.deepEqual(visibilityByName(hidden), {
    Assumptions: "visible",
    "P&L": "hidden",
    Notes: "visible",
  });

  const savedHidden = saveWorkbook(hidden);
  assert.deepEqual(
    visibilityByName(savedHidden),
    { Assumptions: "visible", "P&L": "hidden", Notes: "visible" },
    "the engine's own save keeps the rewritten visibility",
  );

  const veryHidden = core.applyWorkbookSheetVisibility(saved, "P&L", "veryHidden");
  assert.ok(veryHidden);
  assert.deepEqual(visibilityByName(veryHidden), {
    Assumptions: "visible",
    "P&L": "veryHidden",
    Notes: "visible",
  });
  assert.deepEqual(
    visibilityByName(saveWorkbook(veryHidden)),
    { Assumptions: "visible", "P&L": "veryHidden", Notes: "visible" },
    "the engine's own save keeps veryHidden too",
  );

  const visible = core.applyWorkbookSheetVisibility(veryHidden, "P&L", "visible");
  assert.ok(visible);
  assert.deepEqual(visibilityByName(visible), {
    Assumptions: "visible",
    "P&L": "visible",
    Notes: "visible",
  });
});

test("applyWorkbookSheetVisibility rejects unknown sheets and invalid values", () => {
  const saved = saveWorkbook(workbookBytes("financial-model.xlsx"));
  assert.equal(core.applyWorkbookSheetVisibility(saved, "Missing", "hidden"), null);
  assert.equal(core.applyWorkbookSheetVisibility(saved, "", "hidden"), null);
  assert.equal(core.applyWorkbookSheetVisibility(saved, "P&L", "gone"), null);
  assert.equal(core.applyWorkbookSheetVisibility(new Uint8Array([1, 2, 3]), "P&L", "hidden"), null);
});

test("applyWorkbookSheetVisibility returns the input for an unchanged visibility", () => {
  const saved = saveWorkbook(workbookBytes("financial-model.xlsx"));
  assert.equal(core.applyWorkbookSheetVisibility(saved, "P&L", "visible"), saved);
  const hidden = core.applyWorkbookSheetVisibility(saved, "P&L", "hidden");
  assert.equal(core.applyWorkbookSheetVisibility(hidden, "P&L", "hidden"), hidden);
});

test("isWorksheetTabListed keeps veryHidden sheets out of every tab list", () => {
  assert.equal(core.isWorksheetTabListed("visible"), true);
  assert.equal(core.isWorksheetTabListed("hidden"), false);
  assert.equal(core.isWorksheetTabListed("hidden", true), true);
  assert.equal(core.isWorksheetTabListed("veryHidden"), false);
  assert.equal(core.isWorksheetTabListed("veryHidden", true), false);
});
