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
globalThis.Document ??= class Document {};
globalThis.ShadowRoot ??= class ShadowRoot {};
const { strFromU8, unzipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { useXlsxViewerController, XlsxChartOverlay, XlsxContextMenu } = await importFromDemo(
  "@arcships/vue-xlsx",
);

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
  setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
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

function chartByPath(controller, chartPath) {
  return controller.charts.find((chart) => chart.chartPath === chartPath) ?? null;
}

const chartRange = { start: { row: 0, col: 0 }, end: { row: 4, col: 2 } };

test("controller.changeChartType rewrites the part, is undoable and survives export", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController({
    file: sampleBuffer("charts-images.xlsx"),
    fileName: "charts-images.xlsx",
  });
  t.after(() => app.unmount());
  await waitForReady(controller);

  const inserted = controller.insertChart({
    family: "bar",
    range: chartRange,
    title: "Quarterly revenue",
  });
  assert.ok(inserted, "the chart is inserted");
  const chartPath = inserted.chartPath;
  assert.deepEqual(
    inserted.series.map((series) => series.name),
    ["Revenue", "Exceptions"],
  );

  const originalSeries = inserted.series.map((series) => ({
    color: series.color,
    name: series.name,
    values: series.values,
  }));

  assert.equal(controller.changeChartType(inserted.id, "Pie3D"), true);
  await vue.nextTick();
  const pie = chartByPath(controller, chartPath);
  assert.ok(pie, "the rewritten chart is part of the controller chart list");
  assert.equal(pie.chartType, "Pie3D");
  assert.equal(pie.is3d, true);
  assert.equal(controller.selectedChartId, pie.id, "the rewritten chart stays selected");
  assert.equal(controller.canUndo, true);

  assert.equal(controller.changeChartType(pie.id, "Treemap"), true);
  await vue.nextTick();
  const treemap = chartByPath(controller, chartPath);
  assert.equal(treemap.chartType, "Treemap");
  assert.equal(treemap.chartExLayout, "treemap");

  assert.equal(controller.changeChartType(treemap.id, "ColumnClustered"), true);
  await vue.nextTick();
  const clustered = chartByPath(controller, chartPath);
  assert.equal(clustered.chartType, "ColumnClustered");
  assert.deepEqual(
    clustered.series.map((series) => ({
      color: series.color,
      name: series.name,
      values: series.values,
    })),
    originalSeries,
    "all four switches keep series names, colours and values",
  );

  controller.undo();
  await waitFor(
    () => chartByPath(controller, chartPath)?.chartType === "Treemap",
    "undo restores the treemap step",
  );
  await waitFor(() => controller.canRedo, "undo records a redo entry");
  controller.undo();
  await waitFor(
    () => chartByPath(controller, chartPath)?.chartType === "Pie3D",
    "undo restores the 3-D pie step",
  );
  await waitFor(() => controller.canRedo, "the second undo records a redo entry");
  controller.redo();
  await waitFor(
    () => chartByPath(controller, chartPath)?.chartType === "Treemap",
    "redo applies the treemap again",
    5_000,
  );

  assert.equal(controller.changeChartType("missing-chart", "Pie"), false);
  assert.equal(controller.changeChartType(treemap.id, "NotAChartType"), false);

  controller.exportXlsx();
  await vue.nextTick();
  assert.equal(downloads.length, 1);
  const exported = new Uint8Array(await downloads[0].blob.arrayBuffer());
  const archive = unzipSync(exported);
  const contentTypes = strFromU8(archive["[Content_Types].xml"]);
  assert.match(contentTypes, /application\/vnd\.ms-office\.chartex\+xml/);
  assert.match(contentTypes, /application\/vnd\.ms-office\.chartstyle\+xml/);
  assert.match(strFromU8(archive[chartPath]), /layoutId="treemap"/);
});

test("controller.changeChartType refuses read-only workbooks", async (t) => {
  const { app, controller } = await mountController({
    file: sampleBuffer("charts-images.xlsx"),
    fileName: "charts-images.xlsx",
    readOnly: true,
  });
  t.after(() => app.unmount());
  await waitForReady(controller);
  const chart = controller.charts[0];
  assert.ok(chart, "the sample exposes a chart");
  assert.equal(controller.changeChartType(chart.id, "Pie3D"), false);
  assert.equal(chartByPath(controller, chart.chartPath)?.chartType, chart.chartType);
});

const menuChart = {
  anchor: {
    kind: "one-cell",
    from: { col: 0, colOffsetEmu: 0, row: 0, rowOffsetEmu: 0 },
    sizeEmu: { cx: 914400, cy: 914400 },
  },
  chartPath: "xl/charts/chart1.xml",
  chartType: "ColumnClustered",
  editable: true,
  id: "chart-1",
  series: [
    { categories: ["A"], dataPoints: [], id: "series-1", name: "Series 1", values: [1] },
    { categories: ["A"], dataPoints: [], id: "series-2", name: "Series 2", values: [2] },
  ],
  workbookSheetIndex: 0,
  zIndex: 0,
};

test("the chart overlay forwards right clicks to the chart menu", async () => {
  const selected = [];
  const calls = [];
  const controller = {
    activeSheet: {
      colWidths: [],
      defaultColWidthPx: 80,
      rowHeights: [],
      defaultRowHeightPx: 24,
      visibleCols: [],
      visibleRows: [],
    },
    charts: [menuChart],
    isChartsLoading: false,
    moveChartBy: (...args) => calls.push(["move", ...args]),
    resizeChartBy: (...args) => calls.push(["resize", ...args]),
    selectedChartElement: null,
    selectedChartId: null,
    selectChart: (id) => selected.push(id),
    selectChartElement: () => {},
    zoomScale: 100,
  };
  const payloads = [];
  const mounted = await mount(XlsxChartOverlay, {
    controller,
    isDark: false,
    onChartContextMenu: (payload) => payloads.push(payload),
    scrollLeft: 0,
    scrollTop: 0,
    viewportHeight: 600,
    viewportWidth: 800,
  });
  try {
    const item = findByTestId(mounted.root, "xlsx-chart-item");
    assert.ok(item, "the chart item is rendered");
    item.props.onContextmenu({
      clientX: 120,
      clientY: 140,
      preventDefault() {},
      stopPropagation() {},
    });
    assert.deepEqual(selected, [menuChart.id], "right click selects the chart");
    assert.deepEqual(payloads, [{
      chartId: menuChart.id,
      clientX: 120,
      clientY: 140,
    }]);
  } finally {
    mounted.app.unmount();
  }
});

test("the viewers wire the chart menu and the controller exposes changeChartType", () => {
  const viewer = readFileSync(
    new URL("../../packages/vue-xlsx/src/components/XlsxViewer.vue", import.meta.url),
    "utf8",
  );
  const surface = readFileSync(
    new URL("../../packages/vue-xlsx/src/components/XlsxSheetSurface.vue", import.meta.url),
    "utf8",
  );
  const overlay = readFileSync(
    new URL("../../packages/vue-xlsx/src/components/XlsxChartOverlay.vue", import.meta.url),
    "utf8",
  );
  const controller = readFileSync(
    new URL("../../packages/vue-xlsx/src/composables/chart-controller.ts", import.meta.url),
    "utf8",
  );
  for (const source of [viewer, surface]) {
    assert.match(source, /@chart-context-menu="onChartContextMenu"/);
    assert.match(source, /contextMenuRef\.value\?\.openForChart\(payload\.chartId, payload\.clientX, payload\.clientY\)/);
  }
  assert.match(overlay, /@contextmenu\.prevent\.stop="onChartContextMenu\(item, \$event\)"/);
  assert.match(overlay, /emit\("chartContextMenu", \{/);
  assert.match(controller, /changeWorkbookChartType\(/);
  assert.match(controller, /changeChartType,/);
});
