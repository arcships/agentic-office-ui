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
  const source = wasm.Workbook.fromCsvString("alpha,beta,gamma\n");
  const bytes = source.saveXlsxBytes();
  source.free();
  const file = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);

  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file,
        fileName: "number-formats.xlsx",
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
        isNumber: value.is_number,
        isText: value.is_text,
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

function numberFormatSelect(root) {
  return walk(root).find(
    (node) => node.type === "select" && node.props?.["aria-label"] === "Number format",
  );
}

configureWasm();

test("typed dates, percentages, currency and grouped numbers become numbers with Excel's formats", async (t) => {
  const { app, controller } = await mountWorkbookController();
  t.after(() => app.unmount());

  const cases = [
    { cell: { col: 0, row: 1 }, display: "1/15/2024", format: "m/d/yyyy", text: "2024-01-15", value: 45306 },
    { cell: { col: 0, row: 2 }, display: "50.00%", format: "0.00%", text: "50%", value: 0.5 },
    { cell: { col: 0, row: 3 }, display: "$1,234.56", format: "$#,##0.00", text: "$1,234.56", value: 1234.56 },
    { cell: { col: 0, row: 4 }, display: "1,234", format: "#,##0", text: "1,234", value: 1234 },
    { cell: { col: 0, row: 5 }, display: "-500", format: undefined, text: "(500)", value: -500 },
  ];

  for (const entry of cases) {
    assert.equal(
      readCell(controller, entry.cell).numberFormat,
      null,
      `${entry.text} starts in a General cell`,
    );
    controller.setCellValue(entry.cell, entry.text);
    await vue.nextTick();

    const cell = readCell(controller, entry.cell);
    assert.equal(cell.isNumber, true, `${entry.text} must be stored as a number`);
    assert.equal(cell.value, entry.value, `${entry.text} value`);
    if (entry.format) {
      assert.equal(cell.numberFormat?.formatString, entry.format, `${entry.text} format`);
    } else {
      assert.equal(cell.numberFormat, null, `${entry.text} keeps General`);
    }
    assert.equal(controller.getCellDisplayValue(entry.cell), entry.display, `${entry.text} display`);
  }
});

test("typed text that is not a number keeps its type", async (t) => {
  const { app, controller } = await mountWorkbookController();
  t.after(() => app.unmount());

  const cell = { col: 1, row: 2 };
  controller.setCellValue(cell, "2024-13-45");
  controller.setCellValue({ col: 2, row: 2 }, "hello");
  await vue.nextTick();

  assert.equal(readCell(controller, cell).isText, true);
  assert.equal(readCell(controller, cell).value, "2024-13-45");
  assert.equal(readCell(controller, { col: 2, row: 2 }).isText, true);
});

test("typing a number into a text-formatted cell keeps text", async (t) => {
  const { app, controller } = await mountWorkbookController();
  t.after(() => app.unmount());

  const cell = { col: 1, row: 3 };
  controller.selectRange({ start: cell, end: cell });
  controller.setRangeStyle(
    { start: cell, end: cell },
    { numberFormat: { formatType: "builtin", id: 49 } },
  );
  assert.equal(readCell(controller, cell).numberFormat?.formatString, "@");

  controller.setCellValue(cell, "123");
  await vue.nextTick();

  const entered = readCell(controller, cell);
  assert.equal(entered.isText, true, "a @ cell must not coerce digits into a number");
  assert.equal(entered.value, "123");
  assert.equal(controller.getCellDisplayValue(cell), "123");
});

test("internal copy/paste keeps the number format and the value type", async (t) => {
  const { app, controller } = await mountWorkbookController();
  t.after(() => app.unmount());

  const source = { col: 0, row: 1 };
  controller.setCellValue(source, "$1,234.56");
  controller.selectRange({ start: source, end: source });
  await vue.nextTick();
  const clipboard = controller.getClipboardData();
  assert.equal(clipboard.text, "$1,234.56", "plain-text flavour keeps the display string");

  const generalTarget = { col: 2, row: 5 };
  assert.equal(readCell(controller, generalTarget).numberFormat, null);
  controller.selectCell(generalTarget);
  assert.equal(controller.pasteStructuredClipboardData(clipboard.structured), true);
  await vue.nextTick();

  const pasted = readCell(controller, generalTarget);
  assert.equal(pasted.isNumber, true, "paste must keep the typed value");
  assert.equal(pasted.value, 1234.56);
  assert.equal(pasted.numberFormat?.formatString, "$#,##0.00");
  assert.equal(controller.getCellDisplayValue(generalTarget), "$1,234.56");

  const textSource = { col: 0, row: 6 };
  controller.setCellValue(textSource, "'007");
  controller.selectRange({ start: textSource, end: textSource });
  const textClipboard = controller.getClipboardData();
  const textTarget = { col: 2, row: 6 };
  controller.selectCell(textTarget);
  assert.equal(controller.pasteStructuredClipboardData(textClipboard.structured), true);
  await vue.nextTick();

  const pastedText = readCell(controller, textTarget);
  assert.equal(pastedText.isText, true, "a text cell must stay text after paste");
  assert.equal(pastedText.value, "007");
});

test("undoing an applied number format restores the cell's previous format", async (t) => {
  const { app, controller } = await mountWorkbookController();
  t.after(() => app.unmount());

  const cell = { col: 2, row: 3 };
  controller.setCellValue(cell, "77");
  await vue.nextTick();
  assert.equal(readCell(controller, cell).numberFormat, null, "cell starts without a style");

  controller.selectRange({ start: cell, end: cell });
  controller.setRangeStyle(
    { start: cell, end: cell },
    { numberFormat: { formatType: "custom", formatString: "$#,##0.00" } },
  );
  await vue.nextTick();
  assert.equal(readCell(controller, cell).numberFormat?.formatString, "$#,##0.00");
  assert.equal(controller.getCellDisplayValue(cell), "$77.00");

  assert.equal(controller.canUndo, true);
  controller.undo();
  await vue.nextTick();

  const restored = readCell(controller, cell);
  assert.equal(restored.numberFormat, null, "undo must clear the applied format");
  assert.equal(restored.value, 77, "undo must keep the cell value");
  assert.equal(controller.getCellDisplayValue(cell), "77");
});

test("the ribbon reports built-in ids and section formats", async (t) => {
  const state = vue.reactive({
    numberFormat: { formatType: "builtin", id: 14 },
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
  assert.ok(select, "ribbon renders the number format select");

  const cases = [
    [{ formatType: "builtin", id: 14 }, "date"],
    [{ formatType: "builtin", id: 49 }, "text"],
    [{ formatType: "custom", formatString: "#,##0.00_);[Red](#,##0.00)" }, "number"],
    [{ formatType: "custom", formatString: "$#,##0.00" }, "currency"],
    [{ formatType: "custom", formatString: "0.00%" }, "percent"],
    [{ formatType: "custom", formatString: "h:mm:ss" }, "date"],
    [{ formatType: "general", formatString: "General" }, "general"],
  ];

  for (const [numberFormat, expected] of cases) {
    state.numberFormat = numberFormat;
    state.revision += 1;
    await vue.nextTick();
    assert.equal(select.value, expected, JSON.stringify(numberFormat));
  }
});

test("formula, Lotus-prefix and scientific input follow Excel's rules", async (t) => {
  const { app, controller } = await mountWorkbookController();
  t.after(() => app.unmount());

  const read = (cell) => {
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
  };

  controller.setCellValue({ col: 0, row: 1 }, "1");
  controller.setCellValue({ col: 0, row: 2 }, "2");
  controller.setCellFormula({ col: 1, row: 1 }, "=SUM(a2 : a3)");
  await vue.nextTick();
  assert.equal(
    controller.getCellFormula({ col: 1, row: 1 }),
    "=SUM(A2:A3)",
    "references are canonicalized on commit",
  );
  assert.equal(controller.getCellDisplayValue({ col: 1, row: 1 }), "3");

  // A leading "=" and the Lotus "+" prefix create formulas.
  controller.setCellValue({ col: 2, row: 1 }, "=A2*3");
  controller.setCellValue({ col: 3, row: 1 }, "+A2*4");
  await vue.nextTick();
  assert.equal(controller.getCellDisplayValue({ col: 2, row: 1 }), "3");
  assert.equal(controller.getCellFormula({ col: 3, row: 1 }), "=+A2*4");
  assert.equal(controller.getCellDisplayValue({ col: 3, row: 1 }), "4");

  // A signed number stays a number.
  controller.setCellValue({ col: 4, row: 1 }, "-3");
  await vue.nextTick();
  assert.equal(controller.getCellFormula({ col: 4, row: 1 }), "");
  assert.equal(read({ col: 4, row: 1 }).value, -3);

  // Exponent notation gets a scientific format, DATE() a date format.
  controller.setCellValue({ col: 5, row: 1 }, "1e5");
  controller.setCellValue({ col: 6, row: 1 }, "=DATE(2024,1,15)");
  await vue.nextTick();

  const scientific = read({ col: 5, row: 1 });
  assert.equal(scientific.value, 100000);
  assert.equal(scientific.numberFormat?.formatString, "0.00E+00");
  assert.equal(controller.getCellDisplayValue({ col: 5, row: 1 }), "1.00E+05");

  const date = read({ col: 6, row: 1 });
  assert.equal(date.value, 45306);
  assert.equal(date.numberFormat?.formatString, "m/d/yyyy");
  assert.equal(controller.getCellDisplayValue({ col: 6, row: 1 }), "1/15/2024");
});
