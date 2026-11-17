import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  textContent,
  walk,
} from "./vue-test-renderer.mjs";

const { MemoChartSvg } = await importFromDemo("@arcships/vue-xlsx/chart");

const palette = {
  border: "#cbd5e1",
  mutedText: "#64748b",
  surface: "#ffffff",
  text: "#0f172a",
};

function baseChart(overrides = {}) {
  return {
    anchor: {},
    axes: [],
    categoryAxis: {
      labelPosition: "nextTo",
      majorGridlines: false,
    },
    chartType: "ColumnClustered",
    id: "chart-axis-test",
    legend: { position: "r" },
    series: [
      {
        categories: ["A", "B"],
        name: "Series 1",
        values: [0.25, 0.5],
      },
    ],
    sheetIndex: 0,
    title: "Chart",
    valueAxis: {
      labelPosition: "nextTo",
      majorGridlines: false,
    },
    workbookSheetIndex: 0,
    zIndex: 0,
    ...overrides,
  };
}

async function renderChart(chart, { height = 300, width = 400 } = {}) {
  const mounted = await mount(MemoChartSvg, {
    chart,
    palette,
    rect: { height, width },
  });
  const nodes = walk(mounted.root);
  const textNodes = nodes.filter((node) => node.type === "text");
  return {
    ...mounted,
    nodes,
    textNodes,
    texts: textNodes.map((node) => textContent(node)),
    textOf(node) {
      return textContent(node);
    },
  };
}

function legendEntries(nodes) {
  return nodes.filter((node) => node.props?.["data-xlsx-chart-element-kind"] === "legendEntry");
}

test("value axis tick labels use the axis percentage number format", async () => {
  const chart = baseChart({
    valueAxis: {
      labelPosition: "nextTo",
      majorGridlines: true,
      majorUnit: 0.25,
      max: 0.5,
      min: 0,
      numberFormat: { formatCode: "0%" },
    },
  });
  const { texts } = await renderChart(chart);
  assert.ok(texts.includes("25%"), `expected 25% tick, got ${JSON.stringify(texts)}`);
  assert.ok(texts.includes("50%"), `expected 50% tick, got ${JSON.stringify(texts)}`);
  assert.ok(!texts.includes("0.25"), "percentage tick must not fall back to a raw fraction");
});

test("value axis tick labels keep the configured currency symbol and decimals", async () => {
  const chart = baseChart({
    valueAxis: {
      labelPosition: "nextTo",
      majorGridlines: false,
      majorUnit: 1000,
      max: 2000,
      min: 0,
      numberFormat: { formatCode: "$#,##0.00" },
    },
    series: [
      {
        categories: ["A", "B"],
        name: "Series 1",
        values: [500, 1500],
      },
    ],
  });
  const { texts } = await renderChart(chart);
  assert.ok(
    texts.includes("$0.00"),
    `expected a currency tick, got ${JSON.stringify(texts)}`,
  );
  assert.ok(
    texts.some((text) => /^\$1[.,\s]?000\.00$/.test(text)),
    `expected a grouped currency tick, got ${JSON.stringify(texts)}`,
  );
  assert.ok(!texts.includes("1000"), "currency ticks must not drop the symbol");
});

test("value axis tick labels render date serials using the axis date format", async () => {
  const chart = baseChart({
    valueAxis: {
      labelPosition: "nextTo",
      majorGridlines: true,
      majorUnit: 15,
      max: 45031,
      min: 45000,
      numberFormat: { formatCode: "yyyy-mm-dd" },
    },
  });
  const { texts } = await renderChart(chart);
  assert.ok(
    texts.some((text) => /2023/.test(text)),
    `expected a formatted date tick, got ${JSON.stringify(texts)}`,
  );
  assert.ok(
    !texts.some((text) => /^450\d\d$/.test(text)),
    `date axis must not render raw serials, got ${JSON.stringify(texts)}`,
  );
});

test("majorGridlines controls value and category gridlines", async () => {
  const withoutGridlines = await renderChart(baseChart());
  assert.equal(
    withoutGridlines.nodes.filter((node) => node.type === "line" && node.props?.["stroke-width"] === 1).length,
    0,
  );

  const withValueGridlines = await renderChart(baseChart({
    valueAxis: {
      labelPosition: "nextTo",
      majorGridlines: true,
      majorUnit: 0.25,
      max: 0.5,
      min: 0,
    },
  }));
  assert.equal(
    withValueGridlines.nodes.filter((node) => node.type === "line" && node.props?.["stroke-width"] === 1).length,
    3,
  );

  const withCategoryGridlines = await renderChart(baseChart({
    categoryAxis: {
      labelPosition: "nextTo",
      majorGridlines: true,
    },
  }));
  assert.equal(
    withCategoryGridlines.nodes.filter((node) => node.type === "line" && node.props?.["stroke-width"] === 1).length,
    2,
  );
});

test("labelPosition none hides axis labels without hiding the axis itself", async () => {
  const valueHidden = await renderChart(baseChart({
    valueAxis: {
      labelPosition: "none",
      majorGridlines: true,
      majorUnit: 0.25,
      max: 0.5,
      min: 0,
      numberFormat: { formatCode: "0%" },
    },
  }));
  assert.ok(!valueHidden.texts.includes("25%"));
  assert.ok(valueHidden.texts.includes("A"));

  const categoryHidden = await renderChart(baseChart({
    categoryAxis: {
      labelPosition: "none",
      majorGridlines: false,
    },
  }));
  assert.ok(!categoryHidden.texts.includes("A"));
  assert.ok(categoryHidden.texts.includes("0.5"));

  const labelsHigh = await renderChart(baseChart({
    categoryAxis: {
      labelPosition: "high",
      majorGridlines: false,
    },
    valueAxis: {
      labelPosition: "high",
      majorGridlines: false,
      majorUnit: 0.25,
      max: 0.5,
      min: 0,
      numberFormat: { formatCode: "0%" },
    },
  }));
  assert.ok(
    labelsHigh.textNodes.some((node) => (
      textContent(node) === "50%" && node.props?.["text-anchor"] === "start"
    )),
    "value labels with labelPosition high must move to the far side of the plot",
  );
  assert.ok(
    labelsHigh.textNodes.some((node) => textContent(node) === "A" && Number(node.props?.y) < 24),
    "category labels with labelPosition high must move above the plot",
  );
});

test("tickLabelSkip renders every Nth category label", async () => {
  const chart = baseChart({
    categoryAxis: {
      labelPosition: "nextTo",
      majorGridlines: false,
      tickLabelSkip: 2,
    },
    series: [
      {
        categories: ["A", "B", "C", "D", "E", "F"],
        name: "Series 1",
        values: [1, 2, 3, 4, 5, 6],
      },
    ],
  });
  const { texts } = await renderChart(chart);
  for (const label of ["A", "C", "E"]) {
    assert.ok(texts.includes(label), `expected skipped label ${label}, got ${JSON.stringify(texts)}`);
  }
  for (const label of ["B", "D", "F"]) {
    assert.ok(!texts.includes(label), `label ${label} should have been skipped`);
  }
});

test("dense long category labels are rotated or skipped instead of overlapping", async () => {
  const categories = Array.from({ length: 12 }, (_, index) => `Category ${String(index + 1).padStart(2, "0")}`);
  const chart = baseChart({
    categoryAxis: {
      labelPosition: "nextTo",
      majorGridlines: false,
    },
    series: [
      {
        categories,
        name: "Series 1",
        values: categories.map((_, index) => index + 1),
      },
    ],
  });
  const { textNodes } = await renderChart(chart, { width: 400 });
  const categoryTexts = textNodes.filter((node) => textContent(node).startsWith("Category"));
  assert.ok(categoryTexts.length > 0, "expected at least one category label");
  assert.ok(categoryTexts.length < categories.length, "dense category labels must be skipped");
  for (const node of categoryTexts) {
    assert.match(
      String(node.props?.transform ?? ""),
      /rotate\(-45 /,
      `category label ${textContent(node)} must be rotated`,
    );
  }
});

test("long legend labels wrap inside the reserved legend width", async () => {
  const chart = baseChart({
    legend: { position: "r" },
    series: [
      {
        categories: ["A", "B"],
        name: "A very long series name that must wrap across several legend lines",
        values: [0.25, 0.5],
      },
    ],
  });
  const { nodes, texts } = await renderChart(chart, { height: 260, width: 320 });
  const entries = legendEntries(nodes);
  assert.equal(entries.length, 1);
  const lineTexts = walk(entries[0]).filter((node) => node.type === "text");
  assert.ok(lineTexts.length > 1, `expected wrapped legend text, got ${JSON.stringify(texts)}`);
  const legendText = lineTexts.map((node) => textContent(node)).join(" ");
  assert.match(legendText, /long series name/);
});

test("horizontal legends wrap onto multiple rows", async () => {
  const series = Array.from({ length: 6 }, (_, index) => ({
    categories: ["A", "B"],
    name: `Series name ${index + 1}`,
    values: [index + 1, index + 2],
  }));
  const chart = baseChart({
    legend: { position: "b" },
    series,
  });
  const { nodes } = await renderChart(chart, { height: 320, width: 360 });
  const entries = legendEntries(nodes);
  assert.equal(entries.length, series.length);
  const rowTranslations = new Set(entries.map((entry) => String(entry.props?.transform ?? "").replace(/^translate\([^,]+, /, "")));
  assert.ok(rowTranslations.size > 1, "legend entries should wrap onto multiple rows");
});

test("value axis reserve grows for wide tick labels instead of clipping them", async () => {
  const chart = baseChart({
    legend: null,
    series: [
      {
        categories: ["A", "B"],
        name: "Series 1",
        values: [12345678, 23456789],
      },
    ],
  });
  const { textNodes } = await renderChart(chart, { width: 400 });
  const valueLabels = textNodes
    .filter((node) => node.props?.["text-anchor"] === "end")
    .map((node) => Number(node.props?.x))
    .filter((x) => Number.isFinite(x));
  assert.ok(valueLabels.length >= 3, "expected value axis tick labels");
  assert.ok(
    Math.min(...valueLabels) > 40,
    `wide value labels must reserve more than the old 42px plot inset, got x=${Math.min(...valueLabels)}`,
  );
});

test("pie legend and geometry agree when values are zero or negative", async () => {
  const chart = baseChart({
    chartType: "Pie",
    legend: { position: "r" },
    series: [
      {
        categories: ["A", "B", "C"],
        name: "Series 1",
        values: [10, 0, 5],
      },
    ],
  });
  const { nodes, texts } = await renderChart(chart);
  const slicePaths = nodes.filter((node) => (
    node.type === "path" && node.props?.["data-xlsx-chart-point-index"] != null
  ));
  assert.equal(slicePaths.length, 2);
  const legendLabels = texts.filter((text) => ["A", "B", "C"].includes(text));
  assert.deepEqual(legendLabels, ["A", "C"]);

  const negativeChart = baseChart({
    chartType: "Pie",
    legend: { position: "r" },
    series: [
      {
        categories: ["A", "B", "C"],
        name: "Series 1",
        values: [10, -5, 5],
      },
    ],
  });
  const negative = await renderChart(negativeChart);
  assert.equal(
    negative.nodes.filter((node) => node.type === "path" && node.props?.["data-xlsx-chart-point-index"] != null).length,
    2,
  );
  assert.deepEqual(
    negative.texts.filter((text) => ["A", "B", "C"].includes(text)),
    ["A", "C"],
  );
});

test("bar-of-pie starts at the same angle as a regular pie", async () => {
  const chart = baseChart({
    chartType: "BarOfPie",
    firstSliceAngle: 0,
    legend: null,
    series: [
      {
        categories: ["A", "B", "C", "D"],
        name: "Series 1",
        values: [10, 20, 30, 40],
      },
    ],
  });
  const { nodes } = await renderChart(chart);
  const firstSlice = nodes.find((node) => (
    node.type === "path" && node.props?.["data-xlsx-chart-point-index"] != null
  ));
  assert.ok(firstSlice, "expected a bar-of-pie slice path");
  const match = /^\s*M\s*(-?[\d.]+)[ ,]+(-?[\d.]+)/.exec(String(firstSlice.props?.d ?? ""));
  assert.ok(match, `could not parse slice path: ${String(firstSlice.props?.d)}`);
  const startX = Number(match[1]);
  const startY = Number(match[2]);
  assert.ok(Math.abs(startX) < 0.001, `first slice should start at 12 o'clock, got x=${startX}`);
  assert.ok(startY < 0, `first slice should start above the center, got y=${startY}`);
});
