import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
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
