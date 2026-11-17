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

function workbookBuffer() {
  const bytes = readFileSync(
    new URL("../../apps/demo/public/samples/sales-table.xlsx", import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

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

async function mountController(historyOptions) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(),
        fileName: "history-budget.xlsx",
        skipXmlParsing: true,
        useWorker: false,
        ...historyOptions,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "editable XLSX history controller",
  );
  assert.ifError(controller.error);
  return { ...mounted, controller };
}

async function setValues(controller, values) {
  for (const value of values) {
    controller.setCellValue({ col: 0, row: 0 }, value);
    await vue.nextTick();
  }
}

async function countUndos(controller) {
  let count = 0;
  while (controller.canUndo && count < 50) {
    controller.undo();
    count += 1;
    await vue.nextTick();
  }
  return count;
}

configureWasm();

test("XLSX controller opens local UTF-8 CSV as a single editable sheet", async () => {
  const bytes = new TextEncoder().encode("name,amount,note\nAlice,12,\"hello, world\"\nBob,34,中文\n");
  const csvBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  let controller;
  const diagnostics = [];
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: csvBuffer,
        fileName: "sales.csv",
        onDiagnostic: (event) => diagnostics.push(event),
        useWorker: false,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  try {
    await waitFor(
      () => controller.sourceState === "ready" || controller.error,
      "CSV-backed XLSX controller",
    );
    assert.ifError(controller.error);
    assert.equal(controller.displayFileName, "sales.csv");
    assert.equal(controller.sheets.length, 1);
    assert.equal(controller.getCellDisplayValue({ row: 0, col: 0 }), "name");
    assert.equal(controller.getCellDisplayValue({ row: 1, col: 1 }), "12");
    assert.equal(controller.getCellDisplayValue({ row: 1, col: 2 }), "hello, world");
    assert.equal(controller.getCellDisplayValue({ row: 2, col: 2 }), "中文");
    assert.deepEqual(diagnostics.map((event) => event.type), ["load-start", "load-success"]);
    assert.equal(diagnostics[1].bytes, csvBuffer.byteLength);
  } finally {
    mounted.app.unmount();
  }
});

test("XLSX controller opens XLSB through the basic workbook asset path", async () => {
  const wasm = await getSheetsWasmModule();
  const sourceWorkbook = wasm.Workbook.fromCsvString("name,amount\nAlice,12\n");
  const bytes = sourceWorkbook.saveXlsbBytes();
  sourceWorkbook.free();
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: buffer,
        fileName: "sales.xlsb",
        useWorker: false,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  try {
    await waitFor(
      () => controller.sourceState === "ready" || controller.error,
      "XLSB-backed XLSX controller",
    );
    assert.ifError(controller.error);
    assert.equal(controller.displayFileName, "sales.xlsb");
    assert.equal(controller.sheets.length, 1);
    assert.equal(controller.getCellDisplayValue({ row: 0, col: 0 }), "name");
    assert.equal(controller.getCellDisplayValue({ row: 1, col: 0 }), "Alice");
  } finally {
    mounted.app.unmount();
  }
});

test("deferred CSV loading waits for explicit resume before conversion", async () => {
  const bytes = new TextEncoder().encode("name,amount\nAlice,12\n");
  const csvBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        deferLoadingAboveBytes: 1,
        file: csvBuffer,
        fileName: "deferred.csv",
        useWorker: false,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  try {
    await waitFor(() => controller.isLoadDeferred, "deferred CSV state");
    assert.equal(controller.sourceState, "loading");
    assert.equal(controller.isLoading, false);
    assert.equal(controller.activeSheet, null);
    controller.continueDeferredLoad();
    await waitFor(
      () => controller.sourceState === "ready" || controller.error,
      "resumed CSV controller",
    );
    assert.ifError(controller.error);
    assert.equal(controller.getCellDisplayValue({ row: 1, col: 1 }), "12");
  } finally {
    mounted.app.unmount();
  }
});

test("XLSX undo byte budgets belong to each controller instance", async () => {
  const low = await mountController({
    historyMaxBytes: 64 * 1024,
    historyMaxEntries: 20,
  });
  const high = await mountController({
    historyMaxBytes: 2 * 1024 * 1024,
    historyMaxEntries: 20,
  });
  const values = Array.from(
    { length: 6 },
    (_, index) => `${index}:`.padEnd(16 * 1024, String.fromCharCode(65 + index)),
  );

  try {
    await setValues(low.controller, values);
    await setValues(high.controller, values);
    const lowUndoCount = await countUndos(low.controller);
    const highUndoCount = await countUndos(high.controller);

    assert.ok(lowUndoCount >= 1, "the latest low-budget edit must remain undoable");
    assert.ok(lowUndoCount < highUndoCount);
    assert.equal(highUndoCount, values.length);
  } finally {
    low.app.unmount();
    high.app.unmount();
  }
});

test("XLSX history keeps an oversized latest edit, caps count, and clears redo", async () => {
  const oversized = await mountController({
    historyMaxBytes: 512,
    historyMaxEntries: 10,
  });
  const countLimited = await mountController({
    historyMaxBytes: 1024 * 1024,
    historyMaxEntries: 2,
  });
  const cell = { col: 0, row: 0 };
  const initialValue = oversized.controller.getCellDisplayValue(cell);

  try {
    oversized.controller.setCellValue(cell, "超".repeat(8 * 1024));
    await vue.nextTick();
    assert.equal(oversized.controller.canUndo, true);
    oversized.controller.undo();
    await vue.nextTick();
    assert.equal(oversized.controller.getCellDisplayValue(cell), initialValue);
    assert.equal(oversized.controller.canRedo, true);

    oversized.controller.setCellValue(cell, "new branch");
    await vue.nextTick();
    assert.equal(oversized.controller.canRedo, false);
    assert.equal(oversized.controller.canUndo, true);

    await setValues(countLimited.controller, ["one", "two", "three", "four"]);
    assert.equal(await countUndos(countLimited.controller), 2);
  } finally {
    oversized.app.unmount();
    countLimited.app.unmount();
  }
});

test("XLSX undo after a paste restores the pre-paste selection", async () => {
  const mounted = await mountController();
  const target = { col: 2, row: 2 };
  const pastedRange = {
    start: target,
    end: { col: 3, row: 3 },
  };
  const pastedCells = [
    { col: 2, row: 2 },
    { col: 3, row: 2 },
    { col: 2, row: 3 },
    { col: 3, row: 3 },
  ];
  const originalValues = pastedCells.map((cell) =>
    mounted.controller.getCellDisplayValue(cell),
  );

  try {
    mounted.controller.selectCell(target);
    assert.equal(mounted.controller.pasteText("a\tb\nc\td"), true);
    await vue.nextTick();

    assert.deepEqual(mounted.controller.selection, pastedRange);
    assert.equal(mounted.controller.activeCellAddress, "C3");
    assert.equal(mounted.controller.getCellDisplayValue(target), "a");
    assert.equal(mounted.controller.getCellDisplayValue({ col: 3, row: 3 }), "d");

    mounted.controller.undo();
    await vue.nextTick();

    assert.deepEqual(mounted.controller.selection, { start: target, end: target });
    assert.equal(mounted.controller.activeCellAddress, "C3");
    assert.deepEqual(
      pastedCells.map((cell) => mounted.controller.getCellDisplayValue(cell)),
      originalValues,
    );

    mounted.controller.redo();
    await vue.nextTick();

    assert.deepEqual(mounted.controller.selection, pastedRange);
    assert.equal(mounted.controller.activeCellAddress, "C3");
    assert.equal(mounted.controller.getCellDisplayValue(target), "a");
  } finally {
    mounted.app.unmount();
  }
});

test("XLSX undo after a fill restores the pre-fill selection", async () => {
  const mounted = await mountController();
  const sourceRange = {
    start: { col: 0, row: 0 },
    end: { col: 1, row: 0 },
  };
  const filledRange = {
    start: sourceRange.start,
    end: { col: 1, row: 2 },
  };
  const sourceValues = [
    mounted.controller.getCellDisplayValue({ col: 0, row: 0 }),
    mounted.controller.getCellDisplayValue({ col: 1, row: 0 }),
  ];
  const filledCells = [
    { col: 0, row: 1 },
    { col: 1, row: 1 },
    { col: 0, row: 2 },
    { col: 1, row: 2 },
  ];
  const originalValues = filledCells.map((cell) =>
    mounted.controller.getCellDisplayValue(cell),
  );

  try {
    mounted.controller.selectRange(sourceRange);
    mounted.controller.fillSelection(filledRange);
    await vue.nextTick();

    assert.deepEqual(mounted.controller.selection, filledRange);
    assert.equal(mounted.controller.activeCellAddress, "B3");
    assert.equal(mounted.controller.getCellDisplayValue({ col: 0, row: 1 }), sourceValues[0]);
    assert.equal(mounted.controller.getCellDisplayValue({ col: 1, row: 2 }), sourceValues[1]);

    mounted.controller.undo();
    await vue.nextTick();

    assert.deepEqual(mounted.controller.selection, sourceRange);
    assert.equal(mounted.controller.activeCellAddress, "B1");
    assert.deepEqual(
      filledCells.map((cell) => mounted.controller.getCellDisplayValue(cell)),
      originalValues,
    );

    mounted.controller.redo();
    await vue.nextTick();

    assert.deepEqual(mounted.controller.selection, filledRange);
    assert.equal(mounted.controller.activeCellAddress, "B3");
  } finally {
    mounted.app.unmount();
  }
});
