import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

const { getSheetsWasmModule, setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

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

configureWasm();
const wasm = await getSheetsWasmModule();

function formulaWorkbookBuffer() {
  const sourceWorkbook = wasm.Workbook.fromCsvString("1,2\n");
  const sourceSheet = sourceWorkbook.getSheet(0);
  sourceSheet.setFormula("C1", "=A1+B1");
  sourceSheet.free();
  const bytes = sourceWorkbook.saveXlsxBytes();
  sourceWorkbook.free();
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function mountController(buffer) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: buffer,
        fileName: "calculation-safety.xlsx",
        skipXmlParsing: true,
        useWorker: false,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "editable XLSX calculation controller",
  );
  assert.ifError(controller.error);
  return { ...mounted, controller };
}

test("XLSX recalc trap keeps cached values readable and exposes recovery status", async () => {
  const mounted = await mountController(formulaWorkbookBuffer());
  const { controller } = mounted;
  try {
    assert.equal(controller.calculationStatus, null);
    assert.equal(controller.getCellDisplayValue({ row: 0, col: 2 }), "3");

    const originalWorkbook = controller.workbook;
    originalWorkbook.calculate = () => {
      throw new Error("wasm unreachable");
    };

    controller.setCellValue({ row: 0, col: 0 }, "10");
    await vue.nextTick();

    assert.notEqual(controller.workbook, originalWorkbook);
    assert.equal(controller.calculationStatus?.reason, "calculate-trapped");
    assert.match(controller.calculationStatus.message, /cached values/);
    assert.equal(controller.getCellDisplayValue({ row: 0, col: 0 }), "10");
    assert.equal(controller.getCellDisplayValue({ row: 0, col: 2 }), "3");
    assert.equal(controller.canUndo, true);

    controller.recalculate();
    await vue.nextTick();

    assert.equal(controller.calculationStatus, null);
    assert.equal(controller.getCellDisplayValue({ row: 0, col: 2 }), "12");

    controller.setCellValue({ row: 0, col: 1 }, "5");
    await vue.nextTick();

    assert.equal(controller.getCellDisplayValue({ row: 0, col: 2 }), "15");
  } finally {
    mounted.app.unmount();
  }
});

test("XLSX load surfaces unresolved formula references without trapping", async () => {
  const sourceWorkbook = wasm.Workbook.fromCsvString("1,2\n");
  const sourceSheet = sourceWorkbook.getSheet(0);
  sourceSheet.setFormula("C1", "=MissingSheet!A1");
  sourceSheet.free();
  const bytes = sourceWorkbook.saveXlsxBytes();
  sourceWorkbook.free();
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);

  const mounted = await mountController(buffer);
  const { controller } = mounted;
  try {
    assert.equal(controller.calculationStatus?.reason, "unresolved-sheet-refs");
    assert.match(controller.calculationStatus.message, /cached values/);
    assert.equal(controller.getCellDisplayValue({ row: 0, col: 0 }), "1");
  } finally {
    mounted.app.unmount();
  }
});
