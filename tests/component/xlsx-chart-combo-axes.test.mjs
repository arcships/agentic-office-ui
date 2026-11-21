import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

import { MemoChartSvg } from "../../packages/vue-xlsx/dist/chart.js";
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
const rect = { height: 283, width: 567, x: 0, y: 0 };

function loadComboCharts() {
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

async function renderChart(chart) {
  const { root } = await mount(MemoChartSvg, { chart, palette, rect });
  return walk(root)
    .filter((node) => node.type === "text")
    .map((node) => ({
      anchor: node.props?.["text-anchor"],
      text: textContent(node),
      transform: node.props?.transform,
      x: Number(node.props?.x),
    }));
}

function isNumericTick(entry) {
  return /^-?\d/.test(entry.text);
}

const charts = loadComboCharts();
const oneAxis = charts.find((chart) => chart.title === "Combo One Axis");
const twoAxes = charts.find((chart) => chart.title === "Combo Two Axes");
assert.ok(oneAxis, "Combo One Axis must be parsed from the fixture");
assert.ok(twoAxes, "Combo Two Axes must be parsed from the fixture");

test("the two-axis combo renders both value axes with their own scales and title", async () => {
  const texts = await renderChart(twoAxes);
  const leftTicks = texts.filter((entry) => isNumericTick(entry) && entry.x < 100);
  const rightTicks = texts.filter((entry) => isNumericTick(entry) && entry.x > rect.width * 0.6);

  assert.deepEqual(
    leftTicks.map((entry) => entry.text),
    ["-10", "0", "10", "20", "30"],
    "the primary value axis keeps the column scale",
  );
  assert.ok(
    rightTicks.some((entry) => entry.text === "-5"),
    `the secondary value axis renders its own lower bound, got ${JSON.stringify(rightTicks.map((entry) => entry.text))}`,
  );
  assert.ok(
    rightTicks.some((entry) => entry.text === "15"),
    `the secondary value axis renders its own upper bound, got ${JSON.stringify(rightTicks.map((entry) => entry.text))}`,
  );
  assert.ok(
    rightTicks.every((entry) => entry.text !== "30"),
    "the secondary axis must not reuse the primary scale",
  );

  const secondaryTitle = texts.find((entry) => entry.text === "Secondary");
  assert.ok(secondaryTitle, "the secondary axis title renders");
  assert.match(
    secondaryTitle.transform ?? "",
    /^rotate\(90\b/,
    "the secondary axis title is rotated along the right-hand axis",
  );

  const labels = texts.map((entry) => entry.text);
  assert.ok(labels.includes("Revenue"), "the primary legend still renders");
  assert.ok(labels.includes("Cost"), "the secondary legend still renders");
});

test("the one-axis combo keeps rendering a single value axis", async () => {
  const texts = await renderChart(oneAxis);
  const rightTicks = texts.filter((entry) => isNumericTick(entry) && entry.x > rect.width * 0.6);
  assert.deepEqual(rightTicks.map((entry) => entry.text), [], "no right-hand value axis");
  assert.ok(
    !texts.some((entry) => entry.text === "Secondary"),
    "the one-axis combo has no secondary axis title",
  );

  const leftTicks = texts.filter((entry) => isNumericTick(entry) && entry.x < 100);
  assert.deepEqual(
    leftTicks.map((entry) => entry.text),
    ["-10", "0", "10", "20", "30"],
    "the shared axis covers the column scale",
  );
});
