import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

import {
  findByTestId,
  importFromDemo,
  mount,
  teleportTarget,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

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
// `vModelText` feature-detects focus restoration through these globals; the
// custom test renderer hands it plain host nodes, which are neither.
globalThis.Document ??= class Document {};
globalThis.ShadowRoot ??= class ShadowRoot {};
const { strFromU8, unzipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const {
  useXlsxViewerController,
  XlsxAuthoringDialog,
  XlsxRibbon,
  XlsxViewer,
} = await importFromDemo("@arcships/vue-xlsx");

function sampleBuffer(name) {
  const bytes = readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

let wasmConfigured = false;
function configureWasm() {
  if (wasmConfigured) return;
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  );
  wasmConfigured = true;
}
configureWasm();

async function mountController(options) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({ useWorker: false, ...options });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, controller };
}

async function waitForReady(controller) {
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "XLSX controller ready state",
  );
  assert.ifError(controller.error);
  assert.equal(controller.sourceState, "ready");
}

function installDownloadCapture(t) {
  const downloads = [];
  const blobs = new Map();
  const originalDocument = globalThis.document;
  const originalCreateObjectUrl = URL.createObjectURL;
  const originalRevokeObjectUrl = URL.revokeObjectURL;
  let nextObjectUrl = 0;

  globalThis.document = {
    body: { append() {} },
    createElement() {
      return {
        href: "",
        download: "",
        click() {
          downloads.push({ blob: blobs.get(this.href), fileName: this.download });
        },
        remove() {},
      };
    },
  };
  URL.createObjectURL = (blob) => {
    const objectUrl = `blob:test-${nextObjectUrl += 1}`;
    blobs.set(objectUrl, blob);
    return objectUrl;
  };
  URL.revokeObjectURL = () => {};

  t.after(() => {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    if (originalCreateObjectUrl === undefined) delete URL.createObjectURL;
    else URL.createObjectURL = originalCreateObjectUrl;
    if (originalRevokeObjectUrl === undefined) delete URL.revokeObjectURL;
    else URL.revokeObjectURL = originalRevokeObjectUrl;
  });

  return downloads;
}

/** Drives a `v-model` control through its compiled update handler. */
function sheetPart(archive, name = "xl/worksheets/sheet1.xml") {
  return strFromU8(archive[name]);
}

const actionSelection = { start: { row: 1, col: 0 }, end: { row: 4, col: 3 } };

test("controller authoring actions mutate the model, are undoable and survive export", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
  });
  t.after(() => app.unmount());
  await waitForReady(controller);

  // Freeze panes at the active cell.
  controller.selectCell({ row: 1, col: 0 });
  assert.equal(controller.setFreezePanes({ row: 1, col: 0 }), true);
  assert.deepEqual(controller.freezePanes, { row: 1, col: 0 });
  assert.deepEqual(controller.activeSheet?.freezePanes, { row: 1, col: 0 });
  assert.equal(controller.setFreezePanes({ row: 1, col: 0 }), false, "unchanged panes are a no-op");

  // Comment on a cell.
  controller.selectCell({ row: 2, col: 1 });
  assert.equal(
    controller.setCellComment({ row: 2, col: 1 }, { author: "Dana", text: "Check margin" }),
    true,
  );
  assert.deepEqual(controller.getCellComment({ row: 2, col: 1 }), {
    author: "Dana",
    cell: { row: 2, col: 1 },
    text: "Check margin",
  });
  assert.equal(controller.activeSheet?.comments.length, 1);

  // Hyperlink on a cell.
  assert.equal(
    controller.setCellHyperlink({ row: 3, col: 2 }, { target: "https://example.com/report" }),
    true,
  );
  assert.equal(controller.getCellHyperlink({ row: 3, col: 2 })?.target, "https://example.com/report");
  assert.equal(controller.activeSheet?.hyperlinks.length, 1);

  // Conditional format over the selection.
  controller.selectRange(actionSelection);
  assert.equal(
    controller.addConditionalFormatRule({
      kind: "colorScale",
      ranges: [actionSelection],
      colors: ["F8696B", "FFEB84", "63BE7B"],
    }),
    true,
  );
  assert.equal(controller.getConditionalFormatRules().length, 1);

  // ListObject table from the selection.
  controller.selectRange({ start: { row: 0, col: 0 }, end: { row: 4, col: 3 } });
  const table = controller.createTable({ hasHeaders: true, name: "AuthoredTable" });
  assert.ok(table);
  assert.equal(table.name, "AuthoredTable");
  assert.equal(controller.getTableAtCell({ row: 2, col: 2 })?.name, "AuthoredTable");
  assert.ok(controller.tables.some((entry) => entry.name === "AuthoredTable"));

  // Data validation on the selection.
  controller.selectRange({ start: { row: 1, col: 0 }, end: { row: 4, col: 0 } });
  assert.equal(
    controller.setSelectionDataValidation({ listSource: "Open,Closed", validationType: "list" }),
    true,
  );
  assert.equal(controller.activeSheet?.dataValidations.length, 1);

  // A later cell edit rebuilds the sheet model from the parsed state; authored
  // objects must not disappear from the view.
  controller.setCellValue({ row: 0, col: 0 }, "Edited");
  await vue.nextTick();
  assert.deepEqual(controller.activeSheet?.freezePanes, { row: 1, col: 0 });
  assert.equal(controller.activeSheet?.comments.length, 1);
  assert.equal(controller.activeSheet?.conditionalFormatRules.length, 1);
  assert.equal(controller.activeSheet?.hyperlinks.length, 1);
  assert.ok(controller.tables.some((entry) => entry.name === "AuthoredTable"));

  // Chart from the selected range.
  const chart = controller.insertChart({
    family: "line",
    range: actionSelection,
    title: "Authored chart",
  });
  assert.ok(chart, "the chart is inserted and returned");
  assert.equal(chart.chartType, "Line");
  assert.equal(chart.title, "Authored chart");
  assert.ok(
    controller.charts.some((entry) => entry.id === chart.id),
    "the inserted chart is part of the controller chart list",
  );

  // Every action is undoable: undo drops the chart, redo restores it.
  assert.equal(controller.canUndo, true);
  controller.undo();
  await waitFor(
    () => !controller.charts.some((entry) => entry.id === chart.id),
    "undo removes the inserted chart",
  );
  await waitFor(() => controller.canRedo, "undo records a redo entry");
  assert.equal(controller.activeSheet?.comments.length, 1, "earlier authoring survives undo");
  controller.redo();
  await waitFor(
    () => controller.charts.some((entry) => entry.id === chart.id),
    "redo restores the inserted chart",
  );

  // Editing a chart writes the chart part, not just the in-memory model.
  controller.updateChart(chart.id, {
    dataLabels: { showValue: true },
    title: "Renamed chart",
  });
  await vue.nextTick();
  assert.equal(controller.getChartById(chart.id)?.title, "Renamed chart");

  // Export keeps every authored object inside the package.
  controller.exportXlsx();
  await vue.nextTick();
  assert.equal(downloads.length, 1);
  const exported = new Uint8Array(await downloads[0].blob.arrayBuffer());
  const archive = unzipSync(exported);
  const xml = sheetPart(archive);
  assert.match(xml, /<pane[^>]*ySplit="1"/);
  assert.match(xml, /<legacyDrawing\b/);
  assert.match(xml, /<hyperlinks/g);
  assert.match(xml, /<conditionalFormatting /);
  assert.match(xml, /<color rgb="FFF8696B"\/>/);
  assert.match(xml, /<color rgb="FFFFEB84"\/>/);
  assert.match(xml, /<color rgb="FF63BE7B"\/>/);
  assert.doesNotMatch(xml, /<color[^>]*rgb="#/, "exported colours are ARGB, not CSS hex");
  assert.match(xml, /<dataValidations /);
  assert.match(xml, /<tableParts /);
  assert.match(xml, /<drawing /);
  assert.ok(Object.keys(archive).some((path) => path.startsWith("xl/comments/")));
  assert.ok(Object.keys(archive).some((path) => path.startsWith("xl/tables/")));
  assert.ok(Object.keys(archive).some((path) => path.startsWith("xl/charts/chartAuthoring")));
  assert.match(
    strFromU8(archive["[Content_Types].xml"]),
    /spreadsheetml\.comments\+xml/,
  );

  // Reload the exported bytes with a fresh controller.
  const reloaded = await mountController({
    file: exported.buffer.slice(exported.byteOffset, exported.byteOffset + exported.byteLength),
    fileName: "reloaded.xlsx",
  });
  t.after(() => reloaded.app.unmount());
  await waitForReady(reloaded.controller);
  const sheet = reloaded.controller.activeSheet;
  assert.deepEqual(sheet?.freezePanes, { row: 1, col: 0 });
  assert.deepEqual(
    sheet?.comments.map((comment) => [comment.cell, comment.text]),
    [[{ col: 1, row: 2 }, "Check margin"]],
  );
  assert.deepEqual(
    sheet?.hyperlinks.map((link) => [link.cell, link.target]),
    [[{ col: 2, row: 3 }, "https://example.com/report"]],
  );
  assert.deepEqual(sheet?.conditionalFormatRules.map((rule) => rule.kind), ["colorScale"]);
  assert.deepEqual(sheet?.dataValidations.map((rule) => rule.validationType), ["list"]);
  assert.ok(
    reloaded.controller.tables.some((entry) => entry.name === "AuthoredTable"),
    "the table survives a reload",
  );
  assert.deepEqual(
    reloaded.controller.charts.map((entry) => entry.chartType),
    ["Line"],
  );
  assert.equal(reloaded.controller.charts[0].title, "Renamed chart");
  assert.equal(reloaded.controller.charts[0].dataLabels?.showValue, true);
});

test("authoring is a no-op while the sheet is read-only", async (t) => {
  const { app, controller } = await mountController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
  });
  t.after(() => app.unmount());
  await waitForReady(controller);

  const options = controller.readOnly;
  assert.equal(options, false);
  controller.selectCell({ row: 1, col: 1 });
  const mountingReadOnly = await mountController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
    readOnly: true,
  });
  t.after(() => mountingReadOnly.app.unmount());
  await waitForReady(mountingReadOnly.controller);
  mountingReadOnly.controller.selectCell({ row: 1, col: 1 });
  assert.equal(
    mountingReadOnly.controller.setCellComment({ row: 1, col: 1 }, { text: "Read-only" }),
    false,
  );
  assert.equal(mountingReadOnly.controller.setFreezePanes({ row: 1, col: 0 }), false);
  assert.equal(
    mountingReadOnly.controller.addConditionalFormatRule({ kind: "dataBar" }),
    false,
  );
  assert.equal(mountingReadOnly.controller.insertChart({ family: "bar" }), null);
  assert.equal(mountingReadOnly.controller.createTable({ name: "Nope" }), null);
});

test("authoring objects undo and redo through the snapshot history", async (t) => {
  const { app, controller } = await mountController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
  });
  t.after(() => app.unmount());
  await waitForReady(controller);

  const cell = { row: 2, col: 1 };
  controller.selectCell(cell);
  assert.equal(controller.setCellComment(cell, { author: "Dana", text: "Undo me" }), true);
  assert.equal(controller.setCellHyperlink(cell, { target: "https://example.com/undo" }), true);
  assert.equal(controller.activeSheet?.comments.length, 1);
  assert.equal(controller.activeSheet?.hyperlinks.length, 1);

  controller.undo();
  await waitFor(
    () => (controller.activeSheet?.hyperlinks.length ?? 0) === 0,
    "undo removes the hyperlink",
  );
  await waitFor(() => controller.canRedo, "undo records a redo entry");
  assert.equal(controller.activeSheet?.comments.length, 1, "the earlier comment stays");

  controller.redo();
  await waitFor(
    () => (controller.activeSheet?.hyperlinks.length ?? 0) === 1,
    "redo restores the hyperlink",
  );
  assert.equal(controller.activeSheet?.hyperlinks[0]?.target, "https://example.com/undo");
  await waitFor(() => controller.canUndo, "redo settles before the next undo");

  controller.undo();
  await waitFor(
    () => (controller.activeSheet?.hyperlinks.length ?? 0) === 0,
    "second undo removes the hyperlink again",
  );
  await waitFor(() => controller.canRedo, "second undo records a redo entry");
  controller.undo();
  await waitFor(
    () => (controller.activeSheet?.comments.length ?? 0) === 0,
    "undo also removes the earlier comment",
  );
});

test("insertChart falls back to the used range when only one cell is selected", async (t) => {
  const { app, controller } = await mountController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
  });
  t.after(() => app.unmount());
  await waitForReady(controller);

  controller.selectCell({ row: 0, col: 0 });
  const chart = controller.insertChart({ family: "bar" });
  assert.ok(chart, "a single-cell selection still inserts a chart");
  assert.equal(chart.chartType, "ColumnClustered");
  assert.ok(chart.series.length > 0);
});

test("XlsxRibbon exposes authoring entry points and toggles freeze panes", async () => {
  const calls = [];
  const controller = {
    activeCell: { row: 2, col: 1 },
    activeSheet: null,
    canRedo: false,
    canUndo: false,
    canZoomIn: false,
    canZoomOut: false,
    error: null,
    freezePanes: null,
    isLoading: false,
    readOnly: false,
    revision: 0,
    selection: { start: { row: 2, col: 1 }, end: { row: 3, col: 2 } },
    selectedRangeAddress: "B3:C4",
    tabs: [],
    toggleFreezePanesAtActiveCell: () => calls.push("freeze"),
    zoomScale: 100,
  };
  const authoring = [];
  const mounted = await mount(XlsxRibbon, {
    controller,
    onAuthoring: (kind) => authoring.push(kind),
  });
  try {
    for (const [testId, kind] of [
      ["xlsx-ribbon-comment", "comment"],
      ["xlsx-ribbon-hyperlink", "hyperlink"],
      ["xlsx-ribbon-table", "table"],
      ["xlsx-ribbon-conditional", "conditional"],
      ["xlsx-ribbon-chart", "chart"],
      ["xlsx-ribbon-validation", "validation"],
    ]) {
      const button = findByTestId(mounted.root, testId);
      assert.ok(button, `${testId} is rendered`);
      button.props.onClick();
      assert.equal(authoring.at(-1), kind);
    }
    const freeze = findByTestId(mounted.root, "xlsx-ribbon-freeze");
    assert.ok(freeze);
    freeze.props.onClick();
    assert.deepEqual(calls, ["freeze"]);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

