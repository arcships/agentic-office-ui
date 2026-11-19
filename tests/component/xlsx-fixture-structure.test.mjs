import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const xmldom = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= xmldom.DOMParser;
globalThis.Node ??= xmldom.Node;
globalThis.XMLSerializer ??= xmldom.XMLSerializer;

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { useXlsxViewerController, XlsxGrid } = await importFromDemo("@arcships/vue-xlsx");

const wasm = readFileSync(
  new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
);
setWasmSource(wasm.buffer.slice(wasm.byteOffset, wasm.byteOffset + wasm.byteLength));

function sampleBuffer(name) {
  const bytes = readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function mountWorkbook(t, name, { readOnly = true, timeoutMs = 2_000 } = {}) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: sampleBuffer(name),
        fileName: name,
        readOnly,
        useWorker: false,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  t.after(() => mounted.app.unmount());
  await waitFor(
    () => controller.sourceState === "ready" || controller.sourceError,
    `${name} ready`,
    timeoutMs,
  );
  assert.ifError(controller.sourceError);
  assert.equal(controller.sourceState, "ready");
  return { ...mounted, controller };
}

function installGridEnvironment() {
  const previousWindow = globalThis.window;
  const previousResizeObserver = globalThis.ResizeObserver;
  const previousDocument = globalThis.Document;
  const previousShadowRoot = globalThis.ShadowRoot;
  globalThis.window = {
    devicePixelRatio: 1,
    requestAnimationFrame: () => 1,
    cancelAnimationFrame() {},
    addEventListener() {},
    removeEventListener() {},
  };
  globalThis.ResizeObserver = class ResizeObserver {
    observe() {}
    disconnect() {}
  };
  globalThis.Document = class Document {};
  globalThis.ShadowRoot = class ShadowRoot {};
  return () => {
    if (previousWindow === undefined) delete globalThis.window;
    else globalThis.window = previousWindow;
    if (previousResizeObserver === undefined) delete globalThis.ResizeObserver;
    else globalThis.ResizeObserver = previousResizeObserver;
    if (previousDocument === undefined) delete globalThis.Document;
    else globalThis.Document = previousDocument;
    if (previousShadowRoot === undefined) delete globalThis.ShadowRoot;
    else globalThis.ShadowRoot = previousShadowRoot;
  };
}

test("charts-images.xlsx exposes its merge, multi-axis freeze and list validation", async (t) => {
  const { controller } = await mountWorkbook(t, "charts-images.xlsx");
  const sheet = controller.activeSheet;
  assert.equal(sheet?.name, "Dashboard");

  assert.deepEqual(
    controller.freezePanes,
    { row: 2, col: 2 },
    "C3 freeze panes must pin two rows and two columns",
  );
  assert.deepEqual(
    sheet.mergedRegions.map((range) => [
      range.start.row,
      range.start.col,
      range.end.row,
      range.end.col,
    ]),
    [[6, 0, 7, 2]],
    "the fixture must keep its A7:C8 merged region",
  );

  assert.equal(sheet.dataValidations.length, 1);
  const validation = sheet.dataValidations[0];
  assert.equal(validation.validationType, "list");
  assert.match(validation.listSource ?? "", /1,2,3/);
  assert.deepEqual(
    validation.ranges.map((range) => [range.start.row, range.start.col, range.end.row, range.end.col]),
    [[1, 3, 4, 3]],
    "the list validation covers D2:D5",
  );

  assert.equal(sheet.conditionalFormatRules.length, 3, "color scale, data bar and icon set");
  assert.equal(sheet.sparklines.length, 1);
  assert.equal(controller.images.length, 1, "the dashboard image must be parsed");
});

test("sales-table.xlsx keeps multi-axis freeze, a validated status column and two-letter widths", async (t) => {
  const { controller } = await mountWorkbook(t, "sales-table-validated.xlsx");
  const sheet = controller.activeSheet;
  assert.equal(sheet?.name, "Sales Data");
  assert.deepEqual(controller.freezePanes, { row: 2, col: 2 });

  assert.equal(sheet.dataValidations.length, 1);
  const validation = sheet.dataValidations[0];
  assert.equal(validation.validationType, "list");
  assert.match(validation.listSource ?? "", /Active/);
  assert.deepEqual(
    validation.ranges.map((range) => [range.start.row, range.start.col, range.end.row, range.end.col]),
    [[1, 8, 180, 8]],
    "the status list validation covers I2:I181",
  );

  // Column letters beyond Z are workbook indices 26 and 27. Their explicit
  // widths must survive parsing instead of falling back to the default.
  assert.ok(sheet.colWidthOverridesPx[26] > sheet.colWidthOverridesPx[27]);
  assert.ok(sheet.colWidthOverridesPx[27] > 0);
  assert.equal(controller.getCellDisplayValue({ row: 0, col: 26 }), "Region rollup");
  assert.equal(controller.getCellDisplayValue({ row: 0, col: 27 }), "Owner");
  assert.equal(controller.getCellDisplayValue({ row: 1, col: 8 }), "Closed");
});

test("large-grid.xlsx still spans the full fixture range", async (t) => {
  const { controller } = await mountWorkbook(t, "large-grid.xlsx", { timeoutMs: 15_000 });
  const sheet = controller.activeSheet;
  assert.equal(sheet?.name, "Large Grid");
  assert.equal(sheet.maxUsedRow, 550);
  assert.equal(sheet.maxUsedCol, 59);
  assert.equal(controller.getCellDisplayValue({ row: 550, col: 59 }), "R550C60");
});

test("XlsxGrid publishes the fixture merge, freeze and validation counts to the DOM", async (t) => {
  const { controller } = await mountWorkbook(t, "charts-images.xlsx");
  const restoreEnvironment = installGridEnvironment();
  const grid = await mount(XlsxGrid, { controller });
  t.after(() => {
    grid.app.unmount();
    restoreEnvironment();
  });
  const root = findByTestId(grid.root, "xlsx-grid");
  assert.ok(root);
  assert.equal(root.props["data-merged-region-count"], 1);
  assert.equal(root.props["data-freeze-row"], 2);
  assert.equal(root.props["data-freeze-column"], 2);
  assert.equal(root.props["data-data-validation-count"], 1);
});

test("the sales fixture's list validation drives a real dropdown", async (t) => {
  const { controller } = await mountWorkbook(t, "sales-table-validated.xlsx", { readOnly: false });
  const restoreEnvironment = installGridEnvironment();
  const grid = await mount(XlsxGrid, { controller });
  t.after(() => {
    grid.app.unmount();
    restoreEnvironment();
  });

  controller.selectCell({ row: 1, col: 8 });
  await vue.nextTick();
  const trigger = findByTestId(grid.root, "xlsx-validation-dropdown-trigger");
  assert.ok(trigger, "I2 has a list validation and must show its dropdown trigger");
  trigger.props.onPointerdown({
    stopPropagation() {},
    preventDefault() {},
  });
  await vue.nextTick();

  const options = walk(grid.root).filter(
    (node) => node.props?.["data-testid"] === "xlsx-validation-option",
  );
  assert.deepEqual(
    options.map((node) => textContent(node).trim()),
    ["Active", "On Hold", "Closed"],
  );
  options[1].props.onClick({ stopPropagation() {} });
  await vue.nextTick();
  assert.equal(controller.getCellDisplayValue({ row: 1, col: 8 }), "On Hold");
});
