import assert from "node:assert/strict";
import { test } from "node:test";

import { MemoChartSvg } from "../../packages/vue-xlsx/dist/chart.js";
import { mount, textContent, walk } from "./vue-test-renderer.mjs";

const palette = {
  border: "#cbd5e1",
  mutedText: "#64748b",
  surface: "#ffffff",
  text: "#111827",
};
const rect = { height: 320, width: 480, x: 0, y: 0 };

function makeChart(overrides = {}) {
  return {
    anchor: {
      kind: "one-cell",
      from: { col: 0, colOffsetEmu: 0, row: 0, rowOffsetEmu: 0 },
      sizeEmu: { cx: 914400, cy: 914400 },
    },
    axes: [],
    chartType: "Line",
    id: "chart-1",
    legend: null,
    series: [
      {
        categories: ["A", "B"],
        dataPoints: [],
        id: "series-1",
        name: "Series 1",
        values: [10, 20],
      },
    ],
    sheetIndex: 0,
    workbookSheetIndex: 0,
    zIndex: 0,
    ...overrides,
  };
}

async function renderChart(chart) {
  const { root } = await mount(MemoChartSvg, { chart, palette, rect });
  return root;
}

function textNodes(root) {
  return walk(root)
    .filter((node) => node.type === "text")
    .map((node) => ({
      anchor: node.props?.["text-anchor"],
      text: textContent(node),
      x: Number(node.props?.x),
      y: Number(node.props?.y),
    }));
}

function markerNodes(root, seriesIndex = 0) {
  return walk(root).filter((node) => (
    node.type === "path"
    && node.props?.["data-xlsx-chart-point-index"] != null
    && node.props?.["data-xlsx-chart-series-index"] === String(seriesIndex)
  ));
}

function dataLabelNodes(root) {
  return walk(root)
    .filter((node) => node.type === "text" && node.props?.["pointer-events"] === "none")
    .map((node) => ({
      anchor: node.props?.["text-anchor"],
      text: textContent(node),
      x: Number(node.props?.x),
      y: Number(node.props?.y),
    }));
}

test("cartesian bar, line, and area charts render value data labels", async () => {
  for (const chartType of ["ColumnClustered", "Line", "Area"]) {
    const root = await renderChart(makeChart({
      chartType,
      dataLabels: { showValue: true },
      series: [{
        categories: ["A", "B"],
        dataPoints: [],
        id: "series-1",
        name: "Series 1",
        values: [7, 13],
      }],
    }));
    const labels = textNodes(root)
      .filter((node) => node.anchor === "middle")
      .map((node) => node.text);
    assert.ok(labels.includes("7"), `${chartType} renders the first value label`);
    assert.ok(labels.includes("13"), `${chartType} renders the second value label`);
  }
});

test("bar data labels honour number format and series name flags", async () => {
  const root = await renderChart(makeChart({
    chartType: "ColumnClustered",
    dataLabels: {
      numberFormat: { formatCode: "0.0" },
      showSeriesName: true,
      showValue: true,
    },
    series: [{
      categories: ["A"],
      dataPoints: [],
      id: "series-1",
      name: "Revenue",
      values: [42],
    }],
  }));
  const labels = textNodes(root).map((node) => node.text);
  assert.ok(labels.includes("Revenue, 42.0"));
});

test("100% stacked charts divide by the sum of absolute values", async () => {
  for (const chartType of ["LinePercentStacked", "ColumnPercentStacked"]) {
    const root = await renderChart(makeChart({
      chartType,
      dataLabels: { showValue: true },
      series: [
        {
          categories: ["Mixed"],
          dataPoints: [],
          id: "positive",
          name: "Positive",
          values: [5],
        },
        {
          categories: ["Mixed"],
          dataPoints: [],
          id: "negative",
          name: "Negative",
          values: [-5],
        },
      ],
    }));
    const labels = textNodes(root).map((node) => node.text);
    assert.ok(labels.includes("50%"), `${chartType} renders +50%, not +500%`);
    assert.ok(labels.includes("-50%"), `${chartType} renders -50%, not -500%`);
    assert.ok(!labels.includes("500%"), `${chartType} does not explode the percentage`);
  }
});

test("log axes use logarithmic tick spacing", async () => {
  const root = await renderChart(makeChart({
    chartType: "Line",
    series: [{
      categories: ["A", "B", "C"],
      dataPoints: [],
      id: "series-1",
      name: "Series 1",
      values: [1, 10, 100],
    }],
    valueAxis: { logBase: 10, max: 100, min: 1 },
  }));
  const valueTicks = textNodes(root)
    .filter((node) => node.anchor === "end")
    .map((node) => ({ text: node.text, y: node.y }));
  const tickY = new Map(valueTicks.map((tick) => [tick.text, tick.y]));
  assert.ok(tickY.has("1") && tickY.has("10") && tickY.has("100"), "log ticks include 1, 10, 100");
  const gapOneToTen = Math.abs((tickY.get("1") ?? 0) - (tickY.get("10") ?? 0));
  const gapTenToHundred = Math.abs((tickY.get("10") ?? 0) - (tickY.get("100") ?? 0));
  assert.ok(Math.abs(gapOneToTen - gapTenToHundred) < 0.01, "each decade occupies the same number of pixels");
});

test("smooth series draw curves and marker symbols render their own shape", async () => {
  const root = await renderChart(makeChart({
    chartType: "Line",
    series: [{
      categories: ["A", "B", "C"],
      dataPoints: [],
      id: "series-1",
      markerSymbol: "triangle",
      name: "Series 1",
      smooth: true,
      values: [10, 20, 15],
    }],
  }));
  const linePath = walk(root).find((node) => (
    node.type === "path"
    && typeof node.props?.d === "string"
    && node.props.d.length > 0
    && node.props["data-xlsx-chart-point-index"] == null
    && node.props.transform == null
  ));
  assert.ok(linePath, "line path exists");
  assert.ok(linePath.props.d.includes("C"), "smooth line uses a cubic curve");
  const markers = markerNodes(root);
  assert.equal(markers.length, 3);
  for (const marker of markers) {
    assert.ok(marker.props.d.includes("L"), "triangle marker path uses straight segments");
    assert.ok(!marker.props.d.includes("A"), "triangle marker is not a circle");
  }
});

test("3D line, area, and column charts render value data labels", async () => {
  for (const chartType of ["Line", "Area", "ColumnClustered"]) {
    const root = await renderChart(makeChart({
      chartType,
      dataLabels: { showValue: true },
      is3d: true,
      series: [{
        categories: ["A", "B"],
        dataPoints: [],
        id: "series-1",
        name: "Series 1",
        values: [7, 13],
      }],
    }));
    const labels = textNodes(root).map((node) => node.text);
    assert.ok(labels.includes("7"), `${chartType} 3D renders the first value label`);
    assert.ok(labels.includes("13"), `${chartType} 3D renders the second value label`);
  }
});

test("combo charts sharing a value axis do not render a second axis", async () => {
  const makeCombo = (shared) => {
    const categoryAxis = { axisType: "category", id: 1, position: "b" };
    const columnAxis = { axisType: "value", id: 2, position: "l" };
    const lineAxis = shared
      ? { axisType: "value", id: 2, position: "l" }
      : { axisType: "value", id: 3, position: "r" };
    const columnSeries = {
      categories: ["A", "B"],
      dataPoints: [],
      id: "column",
      name: "Column",
      values: [10, 20],
    };
    const lineSeries = {
      categories: ["A", "B"],
      dataPoints: [],
      id: "line",
      name: "Line",
      values: [100, 200],
    };
    return makeChart({
      axes: [categoryAxis, columnAxis, ...(shared ? [] : [lineAxis])],
      categoryAxis,
      chartType: "Combo",
      series: [columnSeries, lineSeries],
      typeGroups: [
        { axisIds: [1, 2], chartType: "ColumnClustered", series: [columnSeries] },
        { axisIds: [1, shared ? 2 : 3], chartType: "Line", series: [lineSeries] },
      ],
      valueAxis: columnAxis,
    });
  };

  const sharedRoot = await renderChart(makeCombo(true));
  const splitRoot = await renderChart(makeCombo(false));
  const secondaryAxisLabels = (root) => textNodes(root).filter((node) => (
    node.anchor === "start" && node.x > rect.width / 2
  ));
  assert.equal(secondaryAxisLabels(sharedRoot).length, 0, "shared axis has no right-hand labels");
  assert.ok(secondaryAxisLabels(splitRoot).length > 0, "independent axes still render the right-hand axis");
  const sharedTickValues = textNodes(sharedRoot)
    .filter((node) => node.anchor === "end")
    .map((node) => node.text);
  assert.ok(sharedTickValues.includes("200"), "shared domain covers both column and line values");
});

function makeComboChart({ columnLabels, lineLabels, secondaryType = "Line", columnSeries, lineSeries, chartLabels } = {}) {
  const categoryAxis = { axisType: "category", id: 1, position: "b" };
  const columnAxis = { axisType: "value", id: 2, position: "l" };
  const lineAxis = { axisType: "value", id: 3, position: "r" };
  const columns = columnSeries ?? [{
    categories: ["A", "B"],
    dataPoints: [],
    id: "column",
    name: "Column",
    values: [10, 20],
  }];
  const line = lineSeries ?? [{
    categories: ["A", "B"],
    dataPoints: [],
    id: "line",
    name: `${secondaryType} series`,
    values: [100, 200],
  }];
  return makeChart({
    axes: [categoryAxis, columnAxis, lineAxis],
    categoryAxis,
    chartType: "Combo",
    dataLabels: chartLabels,
    series: [...columns, ...line],
    typeGroups: [
      {
        axisIds: [1, 2],
        chartType: "ColumnClustered",
        dataLabels: columnLabels,
        series: columns,
      },
      {
        axisIds: [1, 3],
        chartType: secondaryType,
        dataLabels: lineLabels,
        series: line,
      },
    ],
    valueAxis: columnAxis,
  });
}

test("combo charts use group-level labels for column, line, and area groups", async () => {
  for (const secondaryType of ["Line", "Area"]) {
    const root = await renderChart(makeComboChart({
      columnLabels: { numberFormat: { formatCode: "0.0" }, position: "outEnd", showValue: true },
      lineLabels: { position: "t", showCategoryName: true, showSeriesName: true, showValue: true },
      secondaryType,
    }));
    const labels = dataLabelNodes(root).map((node) => node.text);
    assert.ok(labels.includes("10.0"), `${secondaryType} combo column label uses column number format`);
    assert.ok(labels.includes("20.0"), `${secondaryType} combo column label uses column number format`);
    assert.ok(
      labels.includes(`${secondaryType} series, A, 100`),
      `${secondaryType} combo secondary label uses its group flags`,
    );
    assert.ok(
      labels.includes(`${secondaryType} series, B, 200`),
      `${secondaryType} combo secondary label uses its group flags`,
    );
  }
});

test("combo charts honour chart-level labels, position, and showPercent", async () => {
  const fallbackRoot = await renderChart(makeComboChart({
    chartLabels: { showValue: true },
  }));
  const fallbackLabels = dataLabelNodes(fallbackRoot).map((node) => node.text);
  assert.ok(fallbackLabels.includes("10"), "column labels fall back to chart.dataLabels");
  assert.ok(fallbackLabels.includes("100"), "secondary labels fall back to chart.dataLabels");

  const outEndRoot = await renderChart(makeComboChart({
    columnLabels: { numberFormat: { formatCode: "0.0" }, position: "outEnd", showValue: true },
    lineLabels: { showValue: false },
  }));
  const centerRoot = await renderChart(makeComboChart({
    columnLabels: { numberFormat: { formatCode: "0.0" }, position: "ctr", showValue: true },
    lineLabels: { showValue: false },
  }));
  const firstOutEnd = dataLabelNodes(outEndRoot).find((node) => node.text === "10.0");
  const firstCenter = dataLabelNodes(centerRoot).find((node) => node.text === "10.0");
  assert.ok(firstOutEnd && firstCenter, "both combo label positions render");
  assert.notEqual(firstOutEnd.y, firstCenter.y, "combo bar labels honour the configured position");

  const percentRoot = await renderChart(makeComboChart({
    columnLabels: { showPercent: true },
    columnSeries: [
      { categories: ["A", "B"], dataPoints: [], id: "c1", name: "C1", values: [30, 40] },
      { categories: ["A", "B"], dataPoints: [], id: "c2", name: "C2", values: [70, 60] },
    ],
    lineLabels: { showValue: false },
  }));
  const percentLabels = dataLabelNodes(percentRoot).map((node) => node.text);
  assert.ok(percentLabels.includes("30%"), "combo showPercent uses the category total");
  assert.ok(percentLabels.includes("70%"), "combo showPercent uses the category total");
});

test("date axes space points by elapsed time instead of category order", async () => {
  const root = await renderChart(makeChart({
    categoryAxis: {
      axisType: "date",
      numberFormat: { formatCode: "m/d/yyyy" },
    },
    chartType: "Line",
    series: [{
      categories: [44927, 44928, 44986], // 2023-01-01, 2023-01-02, 2023-03-01
      dataPoints: [],
      id: "series-1",
      name: "Series 1",
      values: [10, 20, 30],
    }],
  }));
  const markers = markerNodes(root);
  assert.equal(markers.length, 3);
  const xs = markers.map((marker) => {
    const match = /translate\(([-\d.]+),\s*([-\d.]+)\)/.exec(String(marker.props.transform));
    return Number(match?.[1] ?? Number.NaN);
  });
  const firstGap = xs[1] - xs[0];
  const fullSpan = xs[2] - xs[0];
  assert.ok(Math.abs(firstGap / fullSpan - 1 / 59) < 0.005, "one day is 1/59 of the Jan 1 to Mar 1 span");
});
