import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const { getSheetsWasmModule, setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { XlsxRibbon, XlsxViewer, useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

// The composite viewer mounts the real grid. Only the browser globals that its
// render/paint paths touch are stubbed; the component renderer supplies nodes.
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
globalThis.ResizeObserver ??= class ResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
};
globalThis.requestAnimationFrame ??= (callback) => setTimeout(() => callback(Date.now()), 0);
globalThis.cancelAnimationFrame ??= (handle) => clearTimeout(handle);
if (typeof Object.prototype.getContext !== "function") {
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value: () => null,
  });
}
globalThis.window ??= {
  addEventListener() {},
  cancelAnimationFrame: globalThis.cancelAnimationFrame,
  clearInterval: globalThis.clearInterval,
  clearTimeout: globalThis.clearTimeout,
  devicePixelRatio: 1,
  getComputedStyle: () => ({ getPropertyValue: () => "" }),
  removeEventListener() {},
  requestAnimationFrame: globalThis.requestAnimationFrame,
  setInterval: globalThis.setInterval,
  setTimeout: globalThis.setTimeout,
};

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

async function mountWorkbookController(options = {}) {
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

/** Mounts the real composite viewer and returns a writable locale prop. */
async function mountViewerController(options = {}, viewerProps = {}) {
  const wasm = await getSheetsWasmModule();
  const source = wasm.Workbook.fromCsvString("a,b,c\n");
  const bytes = source.saveXlsxBytes();
  source.free();
  const file = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);

  let controller;
  const locale = vue.ref(viewerProps.locale ?? "en");
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file,
        fileName: "number-locale.xlsx",
        skipXmlParsing: true,
        useWorker: false,
        ...options,
      });
      return () => vue.h(XlsxViewer, {
        controller,
        documentLocale: viewerProps.documentLocale,
        locale: locale.value,
        showDefaultToolbar: false,
        showFormulaBar: false,
        showImages: false,
        showRibbon: false,
        showSheetTabs: false,
      });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "XLSX viewer controller",
  );
  assert.ifError(controller.error);
  return { ...mounted, controller, locale };
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

test("the controller locale formats and parses values that name no locale", async (t) => {
  const { app, controller } = await mountWorkbookController({ locale: "de-DE" });
  t.after(() => app.unmount());
  assert.equal(controller.locale, "de-DE");

  const worksheet = controller.getActiveWorksheet();
  try {
    worksheet.setCell("A1", 1234.56);
    worksheet.setCell("A2", 0.5);
  } finally {
    worksheet.free();
  }

  // Formats without a `[$-…]` section follow the host locale.
  const money = { col: 0, row: 0 };
  await applyNumberFormat(controller, money, "#,##0.00");
  assert.equal(controller.getCellDisplayValue(money), "1.234,56");

  const percent = { col: 0, row: 1 };
  await applyNumberFormat(controller, percent, "0.00%");
  assert.equal(controller.getCellDisplayValue(percent), "50,00%");

  // Typing follows the same locale: 1.234 is grouped, 31.12.2024 is a date.
  const grouped = { col: 0, row: 2 };
  controller.setCellValue(grouped, "1.234");
  await vue.nextTick();
  assert.equal(readCell(controller, grouped).value, 1234);

  const date = { col: 0, row: 3 };
  controller.setCellValue(date, "31.12.2024");
  await vue.nextTick();
  assert.equal(typeof readCell(controller, date).value, "number");
  assert.equal(readCell(controller, date).numberFormat?.formatString, "dd.mm.yyyy");
  assert.equal(controller.getCellDisplayValue(date), "31.12.2024");
});

test("setDocumentLocale switches formats and rejects unknown locales", async (t) => {
  const { app, controller } = await mountWorkbookController({ locale: "de-DE" });
  t.after(() => app.unmount());

  const worksheet = controller.getActiveWorksheet();
  try {
    worksheet.setCell("A1", 1234.56);
  } finally {
    worksheet.free();
  }
  const cell = { col: 0, row: 0 };
  await applyNumberFormat(controller, cell, "#,##0.00");
  assert.equal(controller.getCellDisplayValue(cell), "1.234,56");

  controller.setDocumentLocale("en-US");
  await vue.nextTick();
  assert.equal(controller.locale, "en-US");
  assert.equal(controller.getCellDisplayValue(cell), "1,234.56");

  controller.setDocumentLocale("not a locale");
  await vue.nextTick();
  assert.equal(controller.locale, "en-US");

  controller.setDocumentLocale("de-DE");
  await vue.nextTick();
  assert.equal(controller.getCellDisplayValue(cell), "1.234,56");
});

test("XlsxViewer forwards documentLocale to its controller", async (t) => {
  const calls = [];
  const controller = {
    activeSheet: null,
    canRedo: false,
    canUndo: false,
    error: null,
    isEditingLocked: false,
    isLoading: false,
    readOnly: false,
    redo() {},
    setDocumentLocale(locale) { calls.push(locale); },
    tabs: [],
    undo() {},
  };

  const mounted = await mount(XlsxViewer, {
    controller,
    documentLocale: "de-DE",
    showDefaultToolbar: false,
    showFormulaBar: false,
    showImages: false,
    showRibbon: false,
  });
  t.after(() => mounted.app.unmount());

  assert.deepEqual(calls, ["de-DE"]);
  assert.deepEqual(mounted.warnings, []);
});

test("XlsxViewer locale drives number display and typed input", async (t) => {
  const { app, controller, locale } = await mountViewerController({}, { locale: "de-DE" });
  t.after(() => app.unmount());
  await vue.nextTick();
  assert.equal(controller.locale, "de-DE");

  // Typing a German-grouped number in a General cell picks an Excel number
  // format and renders it with the viewer's separators.
  const numberCell = { col: 0, row: 1 };
  controller.setCellValue(numberCell, "1.234,5");
  await vue.nextTick();
  assert.equal(readCell(controller, numberCell).value, 1234.5);
  assert.equal(readCell(controller, numberCell).numberFormat?.formatString, "#,##0.00");
  assert.equal(controller.getCellDisplayValue(numberCell), "1.234,50");

  // An explicit `@` format is still text: the typed string is not coerced.
  const textCell = { col: 1, row: 1 };
  await applyNumberFormat(controller, textCell, "@");
  controller.setCellValue(textCell, "1.234,5");
  await vue.nextTick();
  assert.equal(readCell(controller, textCell).value, "1.234,5");
  assert.equal(controller.getCellDisplayValue(textCell), "1.234,5");

  // Dates follow the same locale.
  const dateCell = { col: 2, row: 1 };
  controller.setCellValue(dateCell, "31.12.2024");
  await vue.nextTick();
  assert.equal(typeof readCell(controller, dateCell).value, "number");
  assert.equal(controller.getCellDisplayValue(dateCell), "31.12.2024");

  // Switching the prop repaints existing formats and changes parsing.
  locale.value = "en-US";
  await vue.nextTick();
  assert.equal(controller.locale, "en-US");
  assert.equal(controller.getCellDisplayValue(numberCell), "1,234.50");

  const usCell = { col: 0, row: 2 };
  controller.setCellValue(usCell, "1,234.5");
  await vue.nextTick();
  assert.equal(readCell(controller, usCell).value, 1234.5);
  assert.equal(controller.getCellDisplayValue(usCell), "1,234.50");

  // The issue's en-US wording keeps one decimal when the cell asks for one.
  const oneDecimalCell = { col: 1, row: 2 };
  await applyNumberFormat(controller, oneDecimalCell, "#,##0.0");
  controller.setCellValue(oneDecimalCell, "1,234.5");
  await vue.nextTick();
  assert.equal(readCell(controller, oneDecimalCell).value, 1234.5);
  assert.equal(controller.getCellDisplayValue(oneDecimalCell), "1,234.5");
});

test("a viewer without a locale keeps the controller's locale", async (t) => {
  const { app, controller } = await mountViewerController({ locale: "de-DE" });
  t.after(() => app.unmount());
  await vue.nextTick();

  assert.equal(controller.locale, "de-DE");
});

test("documentLocale overrides the viewer locale for numbers", async (t) => {
  const { app, controller } = await mountViewerController(
    {},
    { documentLocale: "en-US", locale: "de-DE" },
  );
  t.after(() => app.unmount());
  await vue.nextTick();
  assert.equal(controller.locale, "en-US");

  const cell = { col: 0, row: 1 };
  await applyNumberFormat(controller, cell, "#,##0.00");
  const worksheet = controller.getActiveWorksheet();
  try {
    worksheet.setCell("A2", 1234.5);
  } finally {
    worksheet.free();
  }
  assert.equal(controller.getCellDisplayValue(cell), "1,234.50");
});
