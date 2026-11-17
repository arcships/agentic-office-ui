import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const { getSheetsWasmModule, setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { XlsxRibbon, useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

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

async function mountWorkbookController() {
  const wasm = await getSheetsWasmModule();
  const source = wasm.Workbook.fromCsvString("a,b,c\n");
  const bytes = source.saveXlsxBytes();
  source.free();
  const file = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);

  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file,
        fileName: "number-locale.xlsx",
        skipXmlParsing: true,
        useWorker: false,
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

function readCell(controller, cell) {
  const worksheet = controller.getActiveWorksheet();
  try {
    const value = worksheet.getCellAt(cell.row, cell.col);
    try {
      return {
        numberFormat: worksheet.getCellStyleAt(cell.row, cell.col)?.numberFormat ?? null,
        value: value.toJs(),
      };
    } finally {
      value.free();
    }
  } finally {
    worksheet.free();
  }
}

async function applyNumberFormat(controller, cell, formatString) {
  controller.selectRange({ start: cell, end: cell });
  controller.setRangeStyle(
    { start: cell, end: cell },
    { numberFormat: { formatType: "custom", formatString } },
  );
  await vue.nextTick();
}

function numberFormatSelect(root) {
  return walk(root).find(
    (node) => node.type === "select" && node.props?.["aria-label"] === "Number format",
  );
}

configureWasm();

test("locale decimal and group separators are honoured when typing", async (t) => {
  const { app, controller } = await mountWorkbookController();
  t.after(() => app.unmount());

  const cases = [
    // German cell: dot is the group separator, comma the decimal mark.
    { code: "[$-407]#,##0.00", input: "1.234,56", row: 1, value: 1234.56 },
    { code: "[$-407]#,##0.00", input: "1.234", row: 2, value: 1234 },
    { code: "[$-407]#,##0.00", input: "1,5", row: 3, value: 1.5 },
    // US cell keeps the opposite reading.
    { code: "[$-409]#,##0.00", input: "1,234.56", row: 4, value: 1234.56 },
    { code: "[$-409]#,##0.00", input: "1.234", row: 5, value: 1.234 },
    // Both separators are unambiguous even without a locale.
    { code: "General", input: "1.234,56", row: 6, value: 1234.56 },
    // Locale separators also apply around percent and currency markings.
    { code: "[$-407]0.00%", input: "1.234,5%", row: 7, value: 12.345 },
    { code: "[$€-407]#,##0.00", input: "1.234,56 €", row: 8, value: 1234.56 },
  ];

  for (const entry of cases) {
    const cell = { col: 0, row: entry.row };
    await applyNumberFormat(controller, cell, entry.code);
    controller.setCellValue(cell, entry.input);
    await vue.nextTick();

    const read = readCell(controller, cell);
    assert.equal(read.value, entry.value, `${entry.code} + ${entry.input}`);
  }
});

test("locale-aware display renders numbers with the format's separators", async (t) => {
  const { app, controller } = await mountWorkbookController();
  t.after(() => app.unmount());

  const worksheet = controller.getActiveWorksheet();
  try {
    worksheet.setCell("A1", 1234.56);
    worksheet.setCell("A2", 0.5);
    worksheet.setCell("A3", -1234.5);
    worksheet.setCell("A4", 1234.56);
    worksheet.setCell("A5", 1234.56);
    worksheet.setCell("A6", 45306);
    worksheet.setCell("A7", "1,234");
    worksheet.setCell("A8", 1234.56);
  } finally {
    worksheet.free();
  }

  const formats = [
    { cell: { col: 0, row: 0 }, code: "[$-407]#,##0.00", display: "1.234,56" },
    { cell: { col: 0, row: 1 }, code: "[$-407]0.00%", display: "50,00%" },
    { cell: { col: 0, row: 2 }, code: "[$-407]#,##0.00", display: "-1.234,50" },
    { cell: { col: 0, row: 3 }, code: "[$-409]#,##0.00", display: "1,234.56" },
    { cell: { col: 0, row: 4 }, code: "#,##0.00", display: "1,234.56" },
    { cell: { col: 0, row: 5 }, code: "[$-407]dd.mm.yyyy", display: "15.01.2024" },
    { cell: { col: 0, row: 6 }, code: "[$-407]@", display: "1,234" },
    { cell: { col: 0, row: 7 }, code: "[$€-407]#,##0.00", display: "€1.234,56" },
  ];

  for (const entry of formats) {
    await applyNumberFormat(controller, entry.cell, entry.code);
    assert.equal(
      controller.getCellDisplayValue(entry.cell),
      entry.display,
      `${entry.code} display`,
    );
  }

  // A German cell keeps its text payload untouched.
  assert.equal(readCell(controller, { col: 0, row: 6 }).value, "1,234");
});

test("locale date order and month names are parsed for typed dates", async (t) => {
  const { app, controller } = await mountWorkbookController();
  t.after(() => app.unmount());

  const cases = [
    { code: "[$-407]dd.mm.yyyy", display: "31.12.2024", input: "31.12.2024", row: 1 },
    { code: "[$-407]dd.mm.yyyy", display: "15.01.2024", input: "15. Januar 2024", row: 2 },
    { code: "[$-407]dd.mm.yyyy", display: "15.01.2024", input: "15-Jan-2024", row: 3 },
    { code: "[$-40C]dd/mm/yyyy", display: "31/12/2024", input: "31/12/2024", row: 4 },
    { code: "[$-411]yyyy\"年\"m\"月\"d\"日\"", display: "2024年1月15日", input: "2024年1月15日", row: 5 },
    { code: "General", display: "1/15/2024", input: "Jan 15, 2024", row: 6 },
    { code: "General", display: "1/15/2024", input: "15 Jan 2024", row: 7 },
  ];

  for (const entry of cases) {
    const cell = { col: 0, row: entry.row };
    await applyNumberFormat(controller, cell, entry.code);
    controller.setCellValue(cell, entry.input);
    await vue.nextTick();

    const read = readCell(controller, cell);
    assert.equal(typeof read.value, "number", `${entry.input} must become a date serial`);
    assert.equal(controller.getCellDisplayValue(cell), entry.display, `${entry.input} display`);
  }
});

test("a General cell with a locale section picks that locale's date order", async (t) => {
  const { app, controller } = await mountWorkbookController();
  t.after(() => app.unmount());

  const cell = { col: 0, row: 1 };
  await applyNumberFormat(controller, cell, "[$-407]General");
  controller.setCellValue(cell, "1/2/2024");
  await vue.nextTick();

  // 1 February 2024 in day-first German order, with a German date format.
  const read = readCell(controller, cell);
  assert.equal(read.value, 45323);
  assert.equal(read.numberFormat?.formatString, "dd.mm.yyyy");
  assert.equal(controller.getCellDisplayValue(cell), "01.02.2024");
});

test("the ribbon classifies locale sections without misreading them", async (t) => {
  const state = vue.reactive({
    numberFormat: { formatString: "[$-407]#,##0.00", formatType: "custom" },
    revision: 0,
  });
  const worksheet = {
    free() {},
    getCellStyleAt: () => ({ numberFormat: state.numberFormat }),
  };
  const controller = {
    get activeCell() { return { col: 0, row: 0 }; },
    get activeSheet() { return { themePalette: null }; },
    get revision() { return state.revision; },
    get isWorkerBacked() { return false; },
    get readOnly() { return false; },
    get selection() { return null; },
    get tabs() { return []; },
    getActiveWorksheet: () => worksheet,
  };

  const { app, root } = await mount(XlsxRibbon, { controller });
  t.after(() => app.unmount());
  const select = numberFormatSelect(root);
  assert.ok(select);

  const cases = [
    [{ formatString: "[$-407]#,##0.00", formatType: "custom" }, "number"],
    [{ formatString: "[$-407]dd.mm.yyyy", formatType: "custom" }, "date"],
    [{ formatString: "[$-411]yyyy\"年\"m\"月\"d\"日\"", formatType: "custom" }, "date"],
    [{ formatString: "[$-409]h:mm AM/PM", formatType: "custom" }, "date"],
    [{ formatType: "builtin", id: 14 }, "date"],
    [{ formatType: "builtin", id: 49 }, "text"],
  ];

  for (const [numberFormat, expected] of cases) {
    state.numberFormat = numberFormat;
    state.revision += 1;
    await vue.nextTick();
    assert.equal(select.value, expected, JSON.stringify(numberFormat));
  }
});
