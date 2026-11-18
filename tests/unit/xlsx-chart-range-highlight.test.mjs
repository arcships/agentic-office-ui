import assert from "node:assert/strict";
import { test } from "node:test";

const {
  CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE,
  XLSX_RANGE_HIGHLIGHT_DEFAULT_COLOR,
  chartRangeCellAt,
  collectChartRangeHighlights,
  createRangeHighlightLayout,
  normalizeHighlightRange,
  parseChartSourceA1Range,
  parseChartSourceRange,
  rangeHighlightStyle,
} = await import(new URL("../../packages/vue-xlsx/dist/index.js", import.meta.url).href);

const SHEETS = [
  { index: 0, name: "Sheet1", workbookSheetIndex: 0 },
  { index: 1, name: "Sales Data", workbookSheetIndex: 1 },
];

function sheet(overrides = {}) {
  return {
    colCount: 4,
    colWidths: [80, 80, 80, 80],
    colWidthOverridesPx: {},
    defaultColWidthPx: 80,
    defaultRowHeightPx: 24,
    freezePanes: null,
    rowCount: 5,
    rowHeights: [24, 24, 24, 24, 24],
    rowHeightOverridesPx: {},
    visibleCols: [0, 1, 2, 3],
    visibleRows: [0, 1, 2, 3, 4],
    ...overrides,
  };
}

function series(overrides = {}) {
  return {
    categories: ["A", "B"],
    categoriesRef: { formula: "Sheet1!$A$2:$A$3" },
    dataPoints: [],
    id: "series-1",
    name: "Revenue",
    values: [10, 20],
    valuesRef: { formula: "Sheet1!$B$2:$B$3" },
    ...overrides,
  };
}

function chart(overrides = {}) {
  return {
    chartType: "ColumnClustered",
    id: "chart-1",
    series: [series()],
    sheetIndex: 0,
    workbookSheetIndex: 0,
    ...overrides,
  };
}

test("chart source formulas parse cells, ranges, whole rows/columns and quoted sheets", () => {
  assert.deepEqual(
    parseChartSourceRange("Sheet1!$B$2:$B$5", { sheets: SHEETS }),
    {
      range: { end: { col: 1, row: 4 }, start: { col: 1, row: 1 } },
      sheetIndex: 0,
    },
  );
  assert.deepEqual(
    parseChartSourceRange("'Sales Data'!$C$1", { sheets: SHEETS }),
    {
      range: { end: { col: 2, row: 0 }, start: { col: 2, row: 0 } },
      sheetIndex: 1,
    },
  );
  // Sheet names are matched case-insensitively the way Excel resolves them.
  assert.equal(parseChartSourceRange("sheet1!A1", { sheets: SHEETS })?.sheetIndex, 0);
  // A formula without a qualifier belongs to the chart's own sheet.
  assert.deepEqual(
    parseChartSourceRange("=A1:B2", { fallbackSheetIndex: 1 }),
    {
      range: { end: { col: 1, row: 1 }, start: { col: 0, row: 0 } },
      sheetIndex: 1,
    },
  );
  // Whole column / whole row references cover Excel's sheet bounds.
  assert.deepEqual(parseChartSourceA1Range("$A:$A"), {
    end: { col: 0, row: 1_048_575 },
    start: { col: 0, row: 0 },
  });
  assert.deepEqual(parseChartSourceA1Range("3:3"), {
    end: { col: 16_383, row: 2 },
    start: { col: 0, row: 2 },
  });
  // External workbooks, unions and unknown sheets are not highlighted.
  assert.equal(parseChartSourceRange("[Book.xlsx]Sheet1!A1", { sheets: SHEETS }), null);
  assert.equal(parseChartSourceRange("Sheet1!A1:B2,C1:C2", { sheets: SHEETS }), null);
  assert.equal(parseChartSourceRange("Missing!A1", { sheets: SHEETS }), null);
  // Defined names resolve through the supplied resolver, but self-references stop.
  assert.equal(
    parseChartSourceRange("Total", {
      fallbackSheetIndex: 0,
      resolveNamedRange: () => "Sheet1!$D$1",
      sheets: SHEETS,
    })?.range.start.col,
    3,
  );
  assert.equal(
    parseChartSourceRange("Total", {
      fallbackSheetIndex: 0,
      resolveNamedRange: () => "Total",
      sheets: SHEETS,
    }),
    null,
  );
});

test("chart selection derives category and value ranges with distinct role colours", () => {
  const highlights = collectChartRangeHighlights(chart(), {
    selection: { chartId: "chart-1", kind: "series", seriesId: "series-1", seriesIndex: 0 },
    sheets: SHEETS,
  });
  assert.deepEqual(highlights.map(({ range, role, seriesName, sheetIndex }) => ({
    range,
    role,
    seriesName,
    sheetIndex,
  })), [
    {
      range: { end: { col: 0, row: 2 }, start: { col: 0, row: 1 } },
      role: "category",
      seriesName: "Revenue",
      sheetIndex: 0,
    },
    {
      range: { end: { col: 1, row: 2 }, start: { col: 1, row: 1 } },
      role: "value",
      seriesName: "Revenue",
      sheetIndex: 0,
    },
  ]);
  assert.equal(highlights[0].color, CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE.category);
  assert.equal(highlights[1].color, CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE.value);
  assert.notEqual(highlights[0].color, highlights[1].color);
});

test("chart-level selection covers every series while a point narrowing singles out one cell", () => {
  const multi = chart({
    series: [
      series(),
      series({
        categoriesRef: { formula: "Sheet1!$A$2:$A$3" },
        id: "series-2",
        name: "Cost",
        valuesRef: { formula: "'Sales Data'!$D$2:$D$3" },
      }),
    ],
  });
  const all = collectChartRangeHighlights(multi, {
    selection: { chartId: "chart-1", kind: "chart" },
    sheets: SHEETS,
  });
  assert.deepEqual(all.map((highlight) => [highlight.role, highlight.sheetIndex]), [
    ["category", 0],
    ["value", 0],
    ["value", 1],
  ], "shared category ranges are emitted once; cross-sheet values point at their sheet");
  assert.equal(all[2].seriesName, "Cost");

  const point = collectChartRangeHighlights(multi, {
    selection: {
      chartId: "chart-1",
      kind: "point",
      pointIndex: 1,
      seriesId: "series-2",
      seriesIndex: 1,
    },
    sheets: SHEETS,
  });
  assert.deepEqual(point.map(({ range, role, sheetIndex }) => ({ range, role, sheetIndex })), [
    {
      range: { end: { col: 0, row: 2 }, start: { col: 0, row: 2 } },
      role: "category",
      sheetIndex: 0,
    },
    {
      range: { end: { col: 3, row: 2 }, start: { col: 3, row: 2 } },
      role: "value",
      sheetIndex: 1,
    },
  ]);
});

test("bubble size references and stale selections are handled explicitly", () => {
  const bubble = chart({
    chartType: "Bubble",
    series: [series({ bubbleSizeRef: { formula: "Sheet1!$C$2:$C$3" } })],
  });
  const highlights = collectChartRangeHighlights(bubble, {
    selection: { chartId: "chart-1", kind: "legendEntry", seriesId: "series-1", seriesIndex: 0 },
    sheets: SHEETS,
  });
  assert.equal(highlights.length, 3);
  assert.equal(highlights[2].role, "bubbleSize");
  assert.equal(highlights[2].color, CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE.bubbleSize);

  assert.deepEqual(collectChartRangeHighlights(bubble, {
    selection: { chartId: "another-chart", kind: "chart" },
    sheets: SHEETS,
  }), []);
  assert.deepEqual(collectChartRangeHighlights(bubble, { sheets: SHEETS }), []);
});

test("point indices map along the axis a chart source actually uses", () => {
  assert.deepEqual(
    chartRangeCellAt({ end: { col: 0, row: 4 }, start: { col: 0, row: 1 } }, 2),
    { col: 0, row: 3 },
  );
  assert.deepEqual(
    chartRangeCellAt({ end: { col: 4, row: 1 }, start: { col: 1, row: 1 } }, 2),
    { col: 3, row: 1 },
  );
  assert.deepEqual(
    chartRangeCellAt({ end: { col: 3, row: 4 }, start: { col: 1, row: 1 } }, 1),
    { col: 1, row: 2 },
  );
  assert.equal(chartRangeCellAt({ end: { col: 0, row: 1 }, start: { col: 0, row: 1 } }, 2), null);
});

test("highlight ranges normalize and reject impossible coordinates", () => {
  assert.deepEqual(
    normalizeHighlightRange({ end: { col: 1, row: 1 }, start: { col: 3, row: 4 } }),
    { end: { col: 3, row: 4 }, start: { col: 1, row: 1 } },
  );
  assert.throws(
    () => normalizeHighlightRange({ end: { col: 0, row: 0 }, start: { col: -1, row: 0 } }),
    TypeError,
  );
  assert.throws(
    () => normalizeHighlightRange({ end: { col: 0, row: 0 }, start: { col: 0, row: 0.5 } }),
    TypeError,
  );
});

test("highlight layout maps ranges to pane-local pixels at any zoom", () => {
  const range = { end: { col: 2, row: 2 }, start: { col: 1, row: 1 } };
  assert.deepEqual(createRangeHighlightLayout(sheet(), 100).resolve(range), [
    { pane: "main", rect: { height: 48, left: 80, top: 24, width: 160 } },
  ]);
  assert.deepEqual(createRangeHighlightLayout(sheet(), 200).resolve(range), [
    { pane: "main", rect: { height: 96, left: 160, top: 48, width: 320 } },
  ]);
  // Scrolling only moves the non-frozen pane; a viewport trims off-screen rows.
  assert.deepEqual(
    createRangeHighlightLayout(sheet(), 100, 40, 20)
      .resolve({ end: { col: 3, row: 4 }, start: { col: 0, row: 0 } }, { height: 72, width: 160 }),
    [{ pane: "main", rect: { height: 96, left: -40, top: -20, width: 240 } }],
    "partially visible rows/columns stay in the rect so the pane clips the edge",
  );
});

test("highlight layout splits ranges at frozen rows and columns", () => {
  const layout = createRangeHighlightLayout(
    sheet({ freezePanes: { col: 2, row: 2 } }),
    100,
    10,
    30,
  );
  assert.equal(layout.frozenWidth, 160);
  assert.equal(layout.frozenHeight, 48);
  assert.deepEqual(
    layout.resolve({ end: { col: 3, row: 3 }, start: { col: 1, row: 1 } }, { height: 200, width: 300 }),
    [
      { pane: "frozen-corner", rect: { height: 24, left: 80, top: 24, width: 80 } },
      { pane: "frozen-rows", rect: { height: 24, left: -10, top: 24, width: 160 } },
      { pane: "frozen-cols", rect: { height: 24, left: 80, top: -6, width: 80 } },
      { pane: "main", rect: { height: 24, left: -10, top: -6, width: 160 } },
    ],
  );
});

test("highlight layout skips hidden axes and keeps zoomed row/column overrides", () => {
  const hidden = sheet({
    colWidthOverridesPx: { 1: 120 },
    rowHeightOverridesPx: { 2: 40 },
    visibleCols: [0, 2, 3],
    visibleRows: [0, 1, 3, 4],
  });
  assert.deepEqual(
    createRangeHighlightLayout(hidden, 100).resolve({
      end: { col: 3, row: 4 },
      start: { col: 1, row: 1 },
    }),
    [{ pane: "main", rect: { height: 72, left: 80, top: 24, width: 160 } }],
    "hidden row 2 / column 1 are skipped and the display-ordered sizes are used",
  );
});

test("range highlight styles use a solid role colour and translucent fill", () => {
  assert.equal(XLSX_RANGE_HIGHLIGHT_DEFAULT_COLOR, "#f59e0b");
  assert.deepEqual(rangeHighlightStyle({ height: 24, left: 10, top: 20, width: 30 }, "#2563eb"), {
    backgroundColor: "rgba(37, 99, 235, 0.18)",
    border: "2px solid #2563eb",
    boxSizing: "border-box",
    height: "24px",
    left: "10px",
    pointerEvents: "none",
    position: "absolute",
    top: "20px",
    width: "30px",
  });
});
