import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  walk,
  waitFor,
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

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const {
  CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE,
  XLSX_RANGE_HIGHLIGHT_DEFAULT_COLOR,
  XlsxRangeHighlightLayer,
  XlsxSheetSurface,
  useXlsxViewerController,
} = await importFromDemo("@arcships/vue-xlsx");

let wasmConfigured = false;
function configureWasm() {
  if (wasmConfigured) return;
  wasmConfigured = true;
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

function workbookBuffer(name = "charts-images.xlsx") {
  const bytes = readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function sheet(overrides = {}) {
  return {
    colCount: 4,
    colWidths: [80, 80, 80, 80],
    colWidthOverridesPx: {},
    defaultColWidthPx: 80,
    defaultRowHeightPx: 24,
    freezePanes: null,
    rowCount: 6,
    rowHeights: [24, 24, 24, 24, 24, 24],
    rowHeightOverridesPx: {},
    visibleCols: [0, 1, 2, 3],
    visibleRows: [0, 1, 2, 3, 4, 5],
    ...overrides,
  };
}

function highlightNodes(root) {
  return walk(root).filter((node) => node.props?.["data-testid"] === "xlsx-range-highlight");
}

async function mountController() {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(),
        fileName: "chart-highlight.xlsx",
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

test("range highlight layer paints, scrolls, zooms and clips inside freeze panes", async () => {
  const controller = vue.reactive({
    activeSheet: sheet(),
    activeSheetIndex: 0,
    highlightedRanges: [],
    zoomScale: 100,
  });
  const viewport = vue.reactive({
    scrollLeft: 0,
    scrollTop: 0,
    viewportHeight: 400,
    viewportWidth: 500,
  });
  const LayerHost = vue.defineComponent({
    setup() {
      return () => vue.h(XlsxRangeHighlightLayer, { controller, ...viewport });
    },
  });
  const mounted = await mount(LayerHost);
  try {
    const layer = findByTestId(mounted.root, "xlsx-range-highlight-layer");
    assert.ok(layer, "the layer renders above the grid body");
    assert.equal(layer.props["aria-hidden"], "true");

    controller.highlightedRanges = [{
      color: "#2563eb",
      id: "search-hit",
      range: { end: { col: 2, row: 2 }, start: { col: 1, row: 1 } },
      role: "custom",
      sheetIndex: 0,
    }];
    await vue.nextTick();
    let nodes = highlightNodes(mounted.root);
    assert.equal(nodes.length, 1);
    assert.equal(layer.props["data-highlight-count"], 1);
    assert.equal(nodes[0].props["data-role"], "custom");
    assert.equal(nodes[0].props["data-range"], "B2:C3");
    assert.equal(nodes[0].props["data-highlight-id"], "search-hit");
    assert.equal(nodes[0].props.style.left, "80px");
    assert.equal(nodes[0].props.style.top, "24px");
    assert.equal(nodes[0].props.style.width, "160px");
    assert.equal(nodes[0].props.style.height, "48px");
    assert.equal(nodes[0].props.style.border, "2px solid #2563eb");
    assert.equal(nodes[0].props.style.backgroundColor, "rgba(37, 99, 235, 0.18)");

    viewport.scrollLeft = 40;
    viewport.scrollTop = 20;
    await vue.nextTick();
    nodes = highlightNodes(mounted.root);
    assert.equal(nodes[0].props.style.left, "40px");
    assert.equal(nodes[0].props.style.top, "4px");
    assert.equal(nodes[0].props.style.width, "160px");
    assert.equal(nodes[0].props.style.height, "48px");

    viewport.scrollLeft = 0;
    viewport.scrollTop = 0;
    controller.zoomScale = 200;
    await vue.nextTick();
    nodes = highlightNodes(mounted.root);
    assert.equal(nodes[0].props.style.left, "160px");
    assert.equal(nodes[0].props.style.top, "48px");
    assert.equal(nodes[0].props.style.width, "320px");
    assert.equal(nodes[0].props.style.height, "96px");

    controller.zoomScale = 100;
    controller.activeSheet = sheet({ freezePanes: { col: 2, row: 2 } });
    await vue.nextTick();
    nodes = highlightNodes(mounted.root);
    assert.equal(nodes.length, 4, "a range crossing the freeze lines is split per pane");
    assert.deepEqual(
      nodes.map((node) => node.parent.props["data-grid-pane"]).sort(),
      ["frozen-cols", "frozen-corner", "frozen-rows", "main"],
    );

    // Highlights on another worksheet never paint on the active sheet.
    controller.activeSheetIndex = 0;
    controller.highlightedRanges = [{
      color: "#f59e0b",
      range: { end: { col: 0, row: 0 }, start: { col: 0, row: 0 } },
      sheetIndex: 1,
    }];
    await vue.nextTick();
    assert.equal(highlightNodes(mounted.root).length, 0);
    assert.equal(layer.props["data-highlight-count"], 0);
  } finally {
    mounted.app.unmount();
  }
});

test("sheet surface mounts the highlight layer above the grid canvas", async (t) => {
  configureWasm();
  const { app, controller } = await mountController();
  t.after(() => app.unmount());

  controller.highlightRanges(
    [{ end: { col: 2, row: 3 }, start: { col: 0, row: 1 } }],
    { color: "#22c55e" },
  );
  const surface = await mount(XlsxSheetSurface, { controller });
  try {
    const grid = findByTestId(surface.root, "xlsx-grid");
    if (grid) {
      grid.clientWidth = 1600;
      grid.clientHeight = 600;
      grid.getBoundingClientRect = () => ({
        bottom: 600, height: 600, left: 0, right: 1600, top: 0, width: 1600,
      });
    }
    controller.clearHighlightedRanges();
    controller.highlightRanges(
      [{ end: { col: 2, row: 3 }, start: { col: 0, row: 1 } }],
      { color: "#22c55e" },
    );
    await vue.nextTick();
    const layer = findByTestId(surface.root, "xlsx-range-highlight-layer");
    assert.ok(layer, "XlsxSheetSurface renders the range highlight layer");
    assert.equal(highlightNodes(surface.root).length, 4, "freeze panes split the range into four fragments");
    const order = walk(surface.root);
    assert.ok(
      order.indexOf(findByTestId(surface.root, "xlsx-grid")) < order.indexOf(layer),
      "the highlight layer paints after the grid canvas",
    );

    controller.clearHighlightedRanges();
    await vue.nextTick();
    assert.equal(highlightNodes(surface.root).length, 0);
  } finally {
    surface.app.unmount();
  }
});

test("controller derives chart source ranges, exposes the public API and clears on deselect", async (t) => {
  configureWasm();
  const { app, controller } = await mountController();
  t.after(() => app.unmount());

  const inserted = controller.insertChart({
    family: "bar",
    range: { end: { col: 2, row: 4 }, start: { col: 0, row: 0 } },
    title: "Highlighted sources",
  });
  assert.ok(inserted, "a chart is inserted from the fixture range");
  await waitFor(() => controller.charts.length > 0, "the chart to load");
  controller.selectChart(inserted.id);
  await waitFor(() => controller.highlightedRanges.length > 0, "chart source highlights");

  const chartLevel = controller.highlightedRanges;
  assert.ok(
    chartLevel.some((highlight) => highlight.role === "category"),
    "chart selection highlights category ranges",
  );
  assert.ok(
    chartLevel.some((highlight) => highlight.role === "value"),
    "chart selection highlights value ranges",
  );
  assert.equal(
    chartLevel.find((highlight) => highlight.role === "category").color,
    CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE.category,
  );
  assert.equal(
    chartLevel.find((highlight) => highlight.role === "value").color,
    CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE.value,
  );
  assert.notEqual(
    CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE.category,
    CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE.value,
  );

  const series = inserted.series[0];
  controller.selectChartElement({
    chartId: inserted.id,
    kind: "series",
    seriesId: series.id,
    seriesIndex: 0,
  });
  await vue.nextTick();
  const seriesHighlights = controller.highlightedRanges;
  assert.deepEqual(seriesHighlights.map((highlight) => highlight.role), ["category", "value"]);
  assert.equal(seriesHighlights[0].seriesName, series.name);
  assert.equal(seriesHighlights[0].sheetIndex, controller.activeSheetIndex);
  assert.equal(seriesHighlights[0].color, CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE.category);
  assert.equal(seriesHighlights[1].color, CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE.value);

  const layerMount = await mount(XlsxRangeHighlightLayer, {
    controller,
    viewportHeight: 500,
    viewportWidth: 800,
  });
  try {
    await vue.nextTick();
    const nodes = highlightNodes(layerMount.root);
    assert.ok(nodes.length >= 2, "the viewer layer paints the series category and value ranges");
    assert.ok(nodes.some((node) => node.props["data-role"] === "category"));
    assert.ok(nodes.some((node) => node.props["data-role"] === "value"));

    controller.setZoomScale(150);
    await vue.nextTick();
    assert.ok(highlightNodes(layerMount.root).length >= 2, "highlights survive zoom");

    controller.clearHighlightedRanges();
    await vue.nextTick();
    assert.deepEqual(controller.highlightedRanges, [], "the public API can clear every visible range");
    assert.equal(highlightNodes(layerMount.root).length, 0);

    // A new selection restores chart source highlights after an explicit clear.
    controller.selectChartElement({
      chartId: inserted.id,
      kind: "series",
      seriesId: series.id,
      seriesIndex: 0,
    });
    await vue.nextTick();
    assert.equal(controller.highlightedRanges.length, 2);

    controller.highlightRanges(
      [{ end: { col: 1, row: 5 }, start: { col: 0, row: 5 } }],
      { id: "host-range" },
    );
    await vue.nextTick();
    assert.equal(
      controller.highlightedRanges.length,
      3,
      "host ranges and chart source ranges are merged",
    );

    controller.clearSelectedChart();
    await vue.nextTick();
    assert.deepEqual(
      controller.highlightedRanges.map((highlight) => highlight.id),
      ["host-range"],
      "deselecting the chart clears only its ranges",
    );
    assert.equal(highlightNodes(layerMount.root).length, 1);
    controller.clearHighlightedRanges();
    await vue.nextTick();
    assert.equal(highlightNodes(layerMount.root).length, 0);
  } finally {
    layerMount.app.unmount();
  }

  // Public arbitrary-range API, independent of the chart selection.
  controller.highlightRanges(
    [{ end: { col: 1, row: 5 }, start: { col: 0, row: 4 } }],
    { color: "#123456", id: "search", label: "Search hit", role: "custom" },
  );
  assert.equal(controller.highlightedRanges.length, 1);
  assert.deepEqual(controller.highlightedRanges[0], {
    color: "#123456",
    id: "search",
    label: "Search hit",
    range: { end: { col: 1, row: 5 }, start: { col: 0, row: 4 } },
    role: "custom",
    sheetIndex: controller.activeSheetIndex,
  });
  controller.clearHighlightedRanges();
  assert.deepEqual(controller.highlightedRanges, []);

  controller.highlightRanges([{ end: { col: 0, row: 0 }, start: { col: 0, row: 0 } }]);
  assert.equal(controller.highlightedRanges[0].color, XLSX_RANGE_HIGHLIGHT_DEFAULT_COLOR);
  assert.equal(controller.highlightedRanges[0].role, "custom");
  controller.clearHighlightedRanges();

  assert.throws(
    () => controller.highlightRanges([{ end: { col: 0, row: 0 }, start: { col: -1, row: 0 } }]),
    TypeError,
  );
  assert.deepEqual(controller.highlightedRanges, []);
});
