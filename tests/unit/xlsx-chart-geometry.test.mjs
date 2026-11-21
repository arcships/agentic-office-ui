import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

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

const core = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);
const chartModule = await import(
  new URL("../../packages/vue-xlsx/dist/chart.js", import.meta.url).href
);
const interaction = await import(
  new URL("../../packages/office-interaction/dist/index.js", import.meta.url).href
);

const wasmPath = requireFromXlsxCore.resolve(
  "@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm",
);
const wasmBytes = readFileSync(wasmPath);
core.setWasmSource(
  wasmBytes.buffer.slice(
    wasmBytes.byteOffset,
    wasmBytes.byteOffset + wasmBytes.byteLength,
  ),
);
const wasm = await core.getSheetsWasmModule();
const sample = new URL(
  "../../apps/demo/public/samples/xlsx-chart-families.xlsx",
  import.meta.url,
);

function loadCharts() {
  const bytes = new Uint8Array(readFileSync(sample));
  const workbook = wasm.Workbook.fromBytes(bytes);
  try {
    const imageAssets = core.parseWorkbookChartStyleAssets(bytes);
    const assets = core.loadWorkbookChartAssets(
      workbook,
      imageAssets,
      new Map(Array.from({ length: workbook.sheetCount }, (_, index) => [index, index])),
    );
    return assets.chartsByWorkbookSheetIndex.flat();
  } finally {
    workbook.free();
  }
}

function chartByTitle(charts, title) {
  const chart = charts.find((entry) => entry.title === title);
  assert.ok(chart, `${title} must be parsed from the fixture`);
  return chart;
}

function close(actual, expected, message, epsilon = 1e-9) {
  assert.ok(
    Math.abs(actual - expected) <= epsilon,
    `${message}: expected ${expected}, got ${actual}`,
  );
}

function closeRect(actual, expected, message) {
  assert.ok(actual, `${message}: missing rect`);
  close(actual.x, expected.x, `${message}.x`);
  close(actual.y, expected.y, `${message}.y`);
  close(actual.width, expected.width, `${message}.width`);
  close(actual.height, expected.height, `${message}.height`);
}

const frame = { width: 567, height: 283 };
const charts = loadCharts();

test("the combo fixture exposes exact plot, axis, bar and line geometry", () => {
  const chart = chartByTitle(charts, "Combo Two Axes");
  const snapshot = chartModule.computeXlsxChartGeometry(chart, { frame });

  assert.equal(snapshot.provider, chartModule.XLSX_CHART_GEOMETRY_PROVIDER_ID);
  assert.equal(snapshot.format, "xlsx");
  assert.equal(snapshot.space, "chart-local-px");
  assert.deepEqual(snapshot.frame, frame);
  assert.equal(snapshot.exact, true);
  assert.deepEqual(
    snapshot.limitations.map((entry) => entry.code),
    ["xlsx.text-metrics-estimated", "xlsx.manual-layout-not-parsed", "xlsx.combo-negative-zero-clamped"],
  );
  assert.deepEqual(interaction.validateChartGeometrySnapshot(snapshot), []);

  closeRect(snapshot.layout.chartArea, { x: 0, y: 0, width: 567, height: 283 }, "chartArea");
  closeRect(snapshot.layout.plotArea, { x: 42, y: 24, width: 398.6, height: 231 }, "plotArea");
  closeRect(snapshot.layout.title, { x: 0, y: 0, width: 567, height: 24 }, "title band");
  assert.equal(snapshot.layout.legend?.position, "right");
  assert.equal(snapshot.layout.legend?.items.length, 2);
  assert.deepEqual(
    snapshot.layout.legend?.items.map((item) => [item.seriesIndex, item.label]),
    [[0, "Revenue"], [1, "Cost"]],
  );
  closeRect(snapshot.layout.legend.items[0].rect, {
    x: 491.8,
    y: 23,
    width: 59.2,
    height: 18,
  }, "legend 0");
  closeRect(snapshot.layout.legend.items[1].rect, {
    x: 491.8,
    y: 41,
    width: 42.4,
    height: 18,
  }, "legend 1");

  assert.deepEqual(
    snapshot.axes.map((axis) => [axis.axisId, axis.kind]),
    [[10, "category"], [100, "value"], [200, "value"]],
  );
  assert.deepEqual(
    snapshot.axes[0].ticks.map((tick) => tick.label),
    ["Q1", "Q2", "Q3", "Q4"],
  );
  [91.825, 191.475, 291.125, 390.775].forEach((expected, index) => {
    close(snapshot.axes[0].ticks[index].labelAnchor.x, expected, `category ${index} x`);
    close(snapshot.axes[0].ticks[index].labelAnchor.y, 268, `category ${index} y`);
  });
  assert.deepEqual(
    snapshot.axes[1].ticks.map((tick) => tick.value),
    [-10, 0, 10, 20, 30],
  );
  assert.deepEqual(
    snapshot.axes[2].ticks.map((tick) => tick.value),
    [-5, 0, 5, 10, 15],
  );
  close(snapshot.axes[1].zeroLine.start.x, 42, "primary zero line start x");
  close(snapshot.axes[1].zeroLine.start.y, 197.25, "primary zero line y");
  close(snapshot.axes[2].ticks[4].labelAnchor.x, 446.6, "secondary tick x");
  close(snapshot.axes[2].ticks[4].labelAnchor.y, 27, "secondary tick y");

  assert.equal(snapshot.series.length, 2);
  const [bars, line] = snapshot.series;
  assert.equal(bars.shape, "bar");
  assert.equal(bars.valueAxisIndex, 1);
  assert.deepEqual(bars.points.map((point) => point.value), [10, 20, -5, 30]);
  closeRect(bars.points[0].rect, {
    x: 75.0838,
    y: 139.5,
    width: 33.4824,
    height: 57.75,
  }, "bar 0");
  closeRect(bars.points[3].rect, {
    x: 374.0338,
    y: 24,
    width: 33.4824,
    height: 173.25,
  }, "bar 3");
  closeRect(bars.points[2].rect, {
    x: 274.3838,
    y: 197.25,
    width: 33.4824,
    height: 1,
  }, "combo negative bar is painted clamped to zero");
  assert.equal(bars.points[0].confidence, "exact");
  assert.equal(bars.points[0].rectConfidence, "exact");

  assert.equal(line.shape, "line");
  assert.equal(line.valueAxisIndex, 2);
  assert.deepEqual(line.points.map((point) => point.value), [4, -3, 8, 12]);
  [[91.825, 151.05], [191.475, 231.9], [291.125, 104.85], [390.775, 58.65]]
    .forEach(([expectedX, expectedY], index) => {
      close(line.points[index].position.x, expectedX, `line point ${index} x`);
      close(line.points[index].position.y, expectedY, `line point ${index} y`);
      assert.equal(line.points[index].confidence, "exact");
      assert.equal(line.points[index].rectConfidence, "approximate");
    });
});

test("the provider is bound to one chart id and can return layout only", () => {
  const chart = chartByTitle(charts, "Column");
  const provider = chartModule.createXlsxChartGeometryProvider();
  assert.equal(provider.id, chartModule.XLSX_CHART_GEOMETRY_PROVIDER_ID);
  assert.equal(provider.format, "xlsx");
  assert.equal(
    provider.getChartGeometry(chart, { chartId: "chart-other", frame }),
    undefined,
  );
  const snapshot = provider.getChartGeometry(chart, { chartId: chart.id, frame });
  assert.ok(snapshot);
  assert.equal(snapshot.chartId, chart.id);
  assert.equal(snapshot.series.length, 1);
  const layoutOnly = provider.getChartGeometry(chart, {
    chartId: chart.id,
    detail: "layout",
    frame,
  });
  assert.deepEqual(layoutOnly.axes, []);
  assert.deepEqual(layoutOnly.series, []);
  assert.equal(layoutOnly.exact, false);
});

test("hit testing resolves marks, legends, ticks and the plot area", () => {
  const chart = chartByTitle(charts, "Combo Two Axes");
  const snapshot = chartModule.computeXlsxChartGeometry(chart, { frame });

  assert.deepEqual(
    interaction.hitTestChartGeometry(snapshot, { x: 91.825, y: 168.375 }),
    {
      kind: "series",
      chartId: chart.id,
      seriesIndex: 0,
      pointIndex: 0,
      confidence: "exact",
    },
  );
  assert.deepEqual(
    interaction.hitTestChartGeometry(snapshot, { x: 503.8, y: 30 }),
    {
      kind: "legendEntry",
      chartId: chart.id,
      legendIndex: 0,
      seriesIndex: 0,
      confidence: "approximate",
    },
  );
  assert.deepEqual(
    interaction.hitTestChartGeometry(snapshot, { x: 30, y: 255 }),
    {
      kind: "axisTick",
      chartId: chart.id,
      axisIndex: 1,
      tickIndex: 0,
      confidence: "approximate",
    },
  );
  assert.deepEqual(
    interaction.hitTestChartGeometry(snapshot, { x: 300, y: 220 }),
    { kind: "plotArea", chartId: chart.id, confidence: "exact" },
  );
  assert.equal(interaction.hitTestChartGeometry(snapshot, { x: 700, y: 400 }), null);
});

test("unsupported chart families expose layout but no series geometry", () => {
  const chart = chartByTitle(charts, "Pie");
  const snapshot = chartModule.computeXlsxChartGeometry(chart, { frame });
  assert.equal(snapshot.exact, false);
  assert.deepEqual(snapshot.series, []);
  assert.deepEqual(snapshot.axes, []);
  assert.ok(snapshot.limitations.some((entry) => entry.code === "xlsx.pie-provider"));
  assert.ok(snapshot.layout.legend.items.length > 0);
  assert.deepEqual(
    snapshot.layout.legend.items.map((item) => item.pointIndex),
    [0, 1, 3],
    "pie legend entries keep the source point index when non-positive points are skipped",
  );
  assert.deepEqual(interaction.validateChartGeometrySnapshot(snapshot), []);

  const fallback = chartByTitle(charts, "3D Fallback");
  const fallbackSnapshot = chartModule.computeXlsxChartGeometry(fallback, { frame });
  assert.equal(fallbackSnapshot.exact, false);
  assert.ok(fallbackSnapshot.limitations.some((entry) => entry.code === "xlsx.3d-projection-provider"));
});

test("bar charts keep negative bars below the zero baseline", () => {
  const chart = chartByTitle(charts, "Column");
  const snapshot = chartModule.computeXlsxChartGeometry(chart, { frame });
  const negative = snapshot.series[0].points.find((point) => point.value === -5);
  assert.ok(negative);
  close(negative.rect.y, 197.25, "negative bar top");
  close(negative.rect.height, 28.875, "negative bar height");
});

test("the provider rejects invalid frames instead of fabricating geometry", () => {
  const chart = chartByTitle(charts, "Column");
  assert.throws(
    () => chartModule.computeXlsxChartGeometry(chart, { frame: { width: 0, height: 283 } }),
    RangeError,
  );
  assert.throws(
    () => chartModule.computeXlsxChartGeometry(chart, { frame: { width: 567, height: Number.NaN } }),
    RangeError,
  );
});
