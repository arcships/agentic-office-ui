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

const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);
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

function fixtureBytes() {
  return new Uint8Array(readFileSync(sample));
}

function loadCharts(bytes) {
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

test("the two-axis combo keeps both value axes and their series groups", () => {
  const chart = chartByTitle(loadCharts(fixtureBytes()), "Combo Two Axes");

  assert.deepEqual(
    chart.axes.map((axis) => axis.id),
    [10, 100, 200],
    "the category axis and both value axes are retained in part order",
  );
  const categoryAxis = chart.axes.find((axis) => axis.id === 10);
  const primaryValueAxis = chart.axes.find((axis) => axis.id === 100);
  const secondaryValueAxis = chart.axes.find((axis) => axis.id === 200);
  assert.equal(categoryAxis?.axisType, "category");
  assert.equal(primaryValueAxis?.axisType, "value");
  assert.equal(secondaryValueAxis?.axisType, "value");
  assert.equal(secondaryValueAxis?.title, "Secondary");
  assert.equal(secondaryValueAxis?.position, "l");
  assert.equal(secondaryValueAxis?.crossId, 10, "the secondary axis crosses the category axis");
  assert.equal(secondaryValueAxis?.crosses, "max");
  assert.equal(primaryValueAxis?.title, undefined, "a missing axis title stays absent");
  assert.equal(chart.valueAxis?.id, 100, "the primary value axis remains the chart value axis");

  assert.deepEqual(
    chart.typeGroups?.map((group) => ({
      axisIds: group.axisIds,
      chartType: group.chartType,
    })),
    [
      { axisIds: [10, 100], chartType: "ColumnClustered" },
      { axisIds: [10, 200], chartType: "Line" },
    ],
    "each series group keeps the axis ids declared by its chart type node",
  );
});

test("the one-axis combo keeps a single value axis", () => {
  const chart = chartByTitle(loadCharts(fixtureBytes()), "Combo One Axis");

  assert.deepEqual(chart.axes.map((axis) => axis.id), [10, 100]);
  const valueAxes = chart.axes.filter((axis) => axis.axisType === "value");
  assert.equal(valueAxes.length, 1, "a shared combo axis is not duplicated");
  assert.equal(valueAxes[0]?.title, undefined);
  assert.deepEqual(
    chart.typeGroups?.map((group) => group.axisIds),
    [[10, 100], [10, 100]],
    "both combo groups declare the same value axis",
  );
});

test("axis parsing keeps the secondary axis scale and number format", () => {
  const bytes = fixtureBytes();
  const archive = unzipSync(bytes);
  const chartPath = "xl/charts/chart11.xml";
  const chartXml = strFromU8(archive[chartPath]);
  const scaledChartXml = chartXml.replace(
    "<valAx><axId val=\"200\"/><scaling><orientation val=\"minMax\"/></scaling>",
    "<valAx><axId val=\"200\"/><scaling><orientation val=\"minMax\"/><max val=\"50\"/><min val=\"-5\"/></scaling>",
  );
  assert.notEqual(scaledChartXml, chartXml, "the fixture must declare a second value axis");
  const formattedChartXml = scaledChartXml.replace(
    "<axPos val=\"l\"/><majorGridlines/><title>",
    "<axPos val=\"l\"/><numFmt formatCode=\"0.0&quot;x&quot;\" sourceLinked=\"0\"/><majorGridlines/><title>",
  );
  assert.notEqual(
    formattedChartXml,
    scaledChartXml,
    "the secondary value axis must carry the title marker",
  );
  archive[chartPath] = strToU8(formattedChartXml);

  const chart = chartByTitle(loadCharts(zipSync(archive)), "Combo Two Axes");
  const secondaryValueAxis = chart.axes.find((axis) => axis.id === 200);
  assert.equal(secondaryValueAxis?.title, "Secondary");
  assert.equal(secondaryValueAxis?.min, -5);
  assert.equal(secondaryValueAxis?.max, 50);
  assert.deepEqual(secondaryValueAxis?.numberFormat, {
    formatCode: '0.0"x"',
    sourceLinked: false,
  });
  assert.equal(
    chart.axes.find((axis) => axis.id === 100)?.min,
    undefined,
    "scaling on the secondary axis does not leak onto the primary axis",
  );
});
