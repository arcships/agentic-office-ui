import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

import {
  importFromDemo,
  mount,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

const { getSheetsWasmModule, setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

const requireFromVueXlsx = createRequire(
  new URL("../../packages/vue-xlsx/package.json", import.meta.url),
);
const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromVueXlsx.resolve("fflate")).href
);
// Structure/style parsing needs a DOM in Node; the browser provides it.
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;

function configureWasm() {
  const bytes = readFileSync(
    new URL(
      "../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm",
      import.meta.url,
    ),
  );
  setWasmSource(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  );
}

function toArrayBuffer(bytes) {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function mountController(file, options = {}) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file,
        fileName: "number-format-undo.xlsx",
        useWorker: false,
        ...options,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "editable XLSX controller",
  );
  assert.ifError(controller.error);
  return { ...mounted, controller };
}

async function mountCsvController() {
  const wasm = await getSheetsWasmModule();
  const source = wasm.Workbook.fromCsvString("a,b,c\n");
  const bytes = source.saveXlsxBytes();
  source.free();
  return mountController(toArrayBuffer(bytes), { skipXmlParsing: true });
}

function readCell(controller, cell) {
  const worksheet = controller.getActiveWorksheet();
  try {
    const value = worksheet.getCellAt(cell.row, cell.col);
    try {
      return {
        isText: value.is_text,
        numberFormat: worksheet.getCellStyleAt(cell.row, cell.col)?.numberFormat ?? null,
        style: worksheet.getCellStyleAt(cell.row, cell.col),
        value: value.toJs(),
      };
    } finally {
      value.free();
    }
  } finally {
    worksheet.free();
  }
}

async function selectRange(controller, range) {
  controller.selectRange(range);
  await vue.nextTick();
}

async function applyRangeStyle(controller, range, style) {
  await selectRange(controller, range);
  controller.setRangeStyle(range, style);
  await vue.nextTick();
}

/**
 * Workbook whose column A carries a `@` (text) column style and no cell style,
 * so the only way a cell can be text is the inherited column format.
 */
async function createTextColumnWorkbook() {
  const wasm = await getSheetsWasmModule();
  const source = wasm.Workbook.fromCsvString("value\n123\n");
  const sheet = source.getSheet(0);
  sheet.setCellStyleAt(0, 0, { numberFormat: { formatType: "builtin", id: 49 } });
  sheet.free();
  const archive = unzipSync(source.saveXlsxBytes());
  source.free();

  const styles = strFromU8(archive["xl/styles.xml"] ?? new Uint8Array());
  const cellXfs = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? "";
  const xfTags = [...cellXfs.matchAll(/<xf\b[^>]*>/g)].map((match) => match[0]);
  const styleIndex = xfTags.findIndex((tag) => tag.includes('numFmtId="49"'));
  assert.ok(styleIndex > 0, "fixture needs a text cell style");

  const sheetXml = strFromU8(archive["xl/worksheets/sheet1.xml"] ?? new Uint8Array());
  const patched = sheetXml.replace(
    "<sheetData>",
    `<cols><col min="1" max="1" width="18" customWidth="1" style="${styleIndex}"/></cols><sheetData>`,
  );
  assert.notEqual(patched, sheetXml, "fixture worksheet needs a sheetData block");
  archive["xl/worksheets/sheet1.xml"] = strToU8(patched);

  return toArrayBuffer(zipSync(archive));
}

configureWasm();

test("undo restores a range where only some cells had a style", async (t) => {
  const { app, controller } = await mountCsvController();
  t.after(() => app.unmount());

  const styled = { col: 0, row: 2 };
  const unstyled = { col: 1, row: 2 };
  controller.setCellStyle(styled, { font: { bold: true }, numberFormat: { formatType: "custom", formatString: "#,##0" } });
  controller.setCellValue(styled, "1234");
  controller.setCellValue(unstyled, "5678");
  await vue.nextTick();
  assert.equal(readCell(controller, styled).style?.font?.bold, true);
  assert.equal(readCell(controller, unstyled).style, null);

  const range = { start: styled, end: unstyled };
  await applyRangeStyle(controller, range, { numberFormat: { formatType: "custom", formatString: "$#,##0.00" } });
  assert.equal(readCell(controller, styled).numberFormat?.formatString, "$#,##0.00");
  assert.equal(readCell(controller, unstyled).numberFormat?.formatString, "$#,##0.00");

  controller.undo();
  await vue.nextTick();

  const restored = readCell(controller, styled);
  assert.equal(restored.style?.font?.bold, true, "the styled cell keeps its font");
  assert.equal(restored.numberFormat?.formatString, "#,##0", "the styled cell gets its old format back");
  assert.equal(restored.value, 1234);
  assert.equal(readCell(controller, unstyled).style, null, "the unstyled cell returns to no style");
  assert.equal(readCell(controller, unstyled).value, 5678);
});

test("redo re-applies a cleared number format and undo clears it again", async (t) => {
  const { app, controller } = await mountCsvController();
  t.after(() => app.unmount());

  const cell = { col: 0, row: 3 };
  const range = { start: cell, end: cell };
  await applyRangeStyle(controller, range, { numberFormat: { formatType: "custom", formatString: "0.00%" } });
  assert.equal(readCell(controller, cell).numberFormat?.formatString, "0.00%");

  controller.undo();
  await vue.nextTick();
  assert.equal(readCell(controller, cell).numberFormat, null);

  assert.equal(controller.canRedo, true);
  controller.redo();
  await vue.nextTick();
  assert.equal(readCell(controller, cell).numberFormat?.formatString, "0.00%");

  controller.undo();
  await vue.nextTick();
  assert.equal(readCell(controller, cell).numberFormat, null);
});

test("undo removes the format that typing added to a General cell", async (t) => {
  const { app, controller } = await mountCsvController();
  t.after(() => app.unmount());

  const cell = { col: 0, row: 4 };
  controller.setCellValue(cell, "50%");
  await vue.nextTick();
  const typed = readCell(controller, cell);
  assert.equal(typed.value, 0.5);
  assert.equal(typed.numberFormat?.formatString, "0.00%");

  controller.undo();
  await vue.nextTick();
  const restored = readCell(controller, cell);
  assert.equal(restored.value, "", "undo restores the empty cell");
  assert.equal(restored.style, null, "undo clears the auto-applied format");

  controller.redo();
  await vue.nextTick();
  assert.equal(readCell(controller, cell).value, 0.5);
  assert.equal(readCell(controller, cell).numberFormat?.formatString, "0.00%");
});

test("undo removes a number format that paste carried", async (t) => {
  const { app, controller } = await mountCsvController();
  t.after(() => app.unmount());

  const source = { col: 0, row: 5 };
  const target = { col: 2, row: 5 };
  controller.setCellValue(source, "$1,234.56");
  await selectRange(controller, { start: source, end: source });
  const clipboard = controller.getClipboardData();
  controller.selectCell(target);
  assert.equal(controller.pasteStructuredClipboardData(clipboard.structured), true);
  await vue.nextTick();
  assert.equal(readCell(controller, target).numberFormat?.formatString, "$#,##0.00");

  controller.undo();
  await vue.nextTick();
  const restored = readCell(controller, target);
  assert.equal(restored.style, null, "paste undo clears the pasted style");
  assert.equal(restored.value, "");
});

test("undo restores the inherited column number format", async (t) => {
  const file = await createTextColumnWorkbook();
  const { app, controller } = await mountController(file, { skipXmlParsing: false });
  t.after(() => app.unmount());

  const cell = { col: 0, row: 3 };
  assert.equal(readCell(controller, cell).style, null, "the fixture cell has no explicit style");

  controller.setCellValue(cell, "123");
  await vue.nextTick();
  assert.equal(readCell(controller, cell).isText, true, "the column's @ format keeps digits as text");

  const range = { start: cell, end: cell };
  await applyRangeStyle(controller, range, {
    numberFormat: { formatType: "custom", formatString: "$#,##0.00" },
  });
  assert.equal(readCell(controller, cell).numberFormat?.formatString, "$#,##0.00");

  controller.undo();
  await vue.nextTick();
  assert.equal(readCell(controller, cell).style, null, "undo returns the cell to the inherited column format");

  controller.setCellValue(cell, "456");
  await vue.nextTick();
  const typed = readCell(controller, cell);
  assert.equal(typed.isText, true, "the inherited column format applies again after undo");
  assert.equal(typed.value, "456");
});
