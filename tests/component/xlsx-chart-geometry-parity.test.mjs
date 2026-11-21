import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

import { MemoChartSvg, computeXlsxChartGeometry } from "../../packages/vue-xlsx/dist/chart.js";
import { mount, textContent, walk } from "./vue-test-renderer.mjs";

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

const palette = {
  border: "#cbd5e1",
  mutedText: "#64748b",
  surface: "#ffffff",
  text: "#0f172a",
};
const frame = { width: 567, height: 283 };

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

function translateOf(node) {
  const match = /^translate\(\s*(-?[\d.]+)[,\s]+(-?[\d.]+)\s*\)$/u.exec(
    String(node.props?.transform ?? ""),
  );
  if (!match) {
    return null;
  }
  return { x: Number(match[1]), y: Number(match[2]) };
}

test("provider geometry equals the combo chart's painted SVG marks and axis anchors", async () => {
  const chart = chartByTitle(loadCharts(), "Combo Two Axes");
  const snapshot = computeXlsxChartGeometry(chart, { frame });
  const { root } = await mount(MemoChartSvg, {
    chart,
    palette,
    rect: { ...frame, left: 0, top: 0 },
  });
  const nodes = walk(root);

  const paintedBars = nodes
    .filter((node) => node.type === "rect")
    .map((node) => ({
      pointIndex: node.props?.["data-xlsx-chart-point-index"],
      seriesIndex: node.props?.["data-xlsx-chart-series-index"],
      x: Number(node.props?.x),
      y: Number(node.props?.y),
      width: Number(node.props?.width),
      height: Number(node.props?.height),
    }))
    .filter((node) => node.seriesIndex === "0" && node.pointIndex != null);

  assert.equal(paintedBars.length, 4);
  const providerBars = snapshot.series[0].points;
  paintedBars.forEach((painted) => {
    const point = providerBars.find((entry) => entry.categoryIndex === Number(painted.pointIndex));
    assert.ok(point, `provider bar for category ${painted.pointIndex}`);
    close(point.rect.x, painted.x, `bar ${painted.pointIndex} x`);
    close(point.rect.y, painted.y, `bar ${painted.pointIndex} y`);
    close(point.rect.width, painted.width, `bar ${painted.pointIndex} width`);
    close(point.rect.height, painted.height, `bar ${painted.pointIndex} height`);
  });

  const paintedMarkers = nodes
    .filter((node) => node.type === "path" && node.props?.transform)
    .map((node) => translateOf(node))
    .filter(Boolean);
  assert.equal(paintedMarkers.length, 4);
  const providerLine = snapshot.series[1].points;
  assert.equal(providerLine.length, 4);
  paintedMarkers.forEach((marker, index) => {
    close(providerLine[index].position.x, marker.x, `line point ${index} x`);
    close(providerLine[index].position.y, marker.y, `line point ${index} y`);
  });

  const texts = nodes
    .filter((node) => node.type === "text")
    .map((node) => ({
      anchor: node.props?.["text-anchor"],
      text: textContent(node),
      x: Number(node.props?.x),
      y: Number(node.props?.y),
    }));
  const categoryTicks = snapshot.axes[0].ticks;
  categoryTicks.forEach((tick, index) => {
    const painted = texts.find((entry) => entry.text === tick.label && entry.y === 268);
    assert.ok(painted, `painted category label ${tick.label}`);
    close(painted.x, tick.labelAnchor.x, `category label ${index} x`);
    close(painted.y, tick.labelAnchor.y, `category label ${index} y`);
  });
  const paintedPrimaryTicks = texts.filter((entry) => entry.anchor === "end" && entry.x === 36);
  assert.deepEqual(
    paintedPrimaryTicks.map((entry) => entry.text),
    snapshot.axes[1].ticks.map((tick) => tick.label),
  );
  snapshot.axes[1].ticks.forEach((tick, index) => {
    close(paintedPrimaryTicks[index].y, tick.labelAnchor.y, `primary tick ${index} y`);
  });
  const paintedSecondaryTicks = texts.filter((entry) => entry.anchor === "start" && entry.x === 446.6);
  assert.deepEqual(
    paintedSecondaryTicks.map((entry) => entry.text),
    snapshot.axes[2].ticks.map((tick) => tick.label),
  );
  snapshot.axes[2].ticks.forEach((tick, index) => {
    close(paintedSecondaryTicks[index].y, tick.labelAnchor.y, `secondary tick ${index} y`);
  });

  const title = texts.find((entry) => entry.text === chart.title);
  assert.ok(title);
  close(title.x, snapshot.layout.titleLabel.anchor.x, "title x");
  close(title.y, snapshot.layout.titleLabel.anchor.y, "title y");
});

test("horizontal bar geometry follows the reversed category axis and negative bars", async () => {
  const chart = chartByTitle(loadCharts(), "Stacked Bar");
  const snapshot = computeXlsxChartGeometry(chart, { frame });
  const { root } = await mount(MemoChartSvg, {
    chart,
    palette,
    rect: { ...frame, left: 0, top: 0 },
  });
  const nodes = walk(root);

  assert.deepEqual(
    snapshot.axes[0].ticks.map((tick) => tick.label),
    ["Q4", "Q3", "Q2", "Q1"],
  );
  const paintedBars = nodes
    .filter((node) => node.type === "rect")
    .map((node) => ({
      pointIndex: node.props?.["data-xlsx-chart-point-index"],
      seriesIndex: node.props?.["data-xlsx-chart-series-index"],
      x: Number(node.props?.x),
      y: Number(node.props?.y),
      width: Number(node.props?.width),
      height: Number(node.props?.height),
    }))
    .filter((node) => node.seriesIndex === "0" && node.pointIndex != null);
  assert.equal(paintedBars.length, 4);
  paintedBars.forEach((painted) => {
    const point = snapshot.series[0].points.find(
      (entry) => entry.categoryIndex === Number(painted.pointIndex),
    );
    assert.ok(point, `provider bar for painted category ${painted.pointIndex}`);
    close(point.rect.x, painted.x, `horizontal bar ${painted.pointIndex} x`);
    close(point.rect.y, painted.y, `horizontal bar ${painted.pointIndex} y`);
    close(point.rect.width, painted.width, `horizontal bar ${painted.pointIndex} width`);
    close(point.rect.height, painted.height, `horizontal bar ${painted.pointIndex} height`);
  });
  const paintedLabels = nodes
    .filter((node) => node.type === "text")
    .map((node) => textContent(node))
    .filter((text) => ["Q1", "Q2", "Q3", "Q4"].includes(text));
  assert.deepEqual(paintedLabels, ["Q4", "Q3", "Q2", "Q1"]);
});

test("line chart point anchors and category ticks match the painted SVG", async () => {
  const chart = chartByTitle(loadCharts(), "Smooth Line");
  const snapshot = computeXlsxChartGeometry(chart, { frame });
  const { root } = await mount(MemoChartSvg, {
    chart,
    palette,
    rect: { ...frame, left: 0, top: 0 },
  });
  const nodes = walk(root);
  const markers = nodes
    .filter((node) => node.type === "path" && node.props?.transform)
    .map((node) => translateOf(node))
    .filter(Boolean);
  assert.equal(markers.length, 4);
  snapshot.series[0].points.forEach((point, index) => {
    close(point.position.x, markers[index].x, `line ${index} x`);
    close(point.position.y, markers[index].y, `line ${index} y`);
  });
  const labels = nodes
    .filter((node) => node.type === "text")
    .map((node) => ({ text: textContent(node), x: Number(node.props?.x), y: Number(node.props?.y) }));
  snapshot.axes[0].ticks.forEach((tick) => {
    const painted = labels.find((entry) => entry.text === tick.label && entry.y === 268);
    assert.ok(painted, `painted category label ${tick.label}`);
    close(painted.x, tick.labelAnchor.x, `category ${tick.label} x`);
  });
  const valueLabels = labels.filter((entry) => entry.anchor === undefined && /^-?\d+$/u.test(entry.text));
  assert.ok(valueLabels.length > 0);
});
