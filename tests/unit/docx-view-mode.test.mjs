import assert from "node:assert/strict";
import test from "node:test";

const docx = await import(
  new URL("../../packages/vue-docx/dist/index.js", import.meta.url).href,
);

const {
  DOCX_VIEW_MODES,
  docxPageAtPoint,
  docxPageRows,
  docxRowIndexAtOffset,
  docxRowIndexForPage,
  docxRowsInViewport,
  docxSnapTopForRow,
  docxSpreadStepPage,
  isDocxSpreadMode,
  normalizeDocxFitMode,
  normalizeDocxViewMode,
  resolveDocxViewLayout,
} = docx;

const PAGE = { widthPx: 200, heightPx: 300 };

function layoutFor(pageCount, mode, overrides = {}) {
  return resolveDocxViewLayout({
    pageCount,
    mode,
    pageSize: () => PAGE,
    zoom: 1,
    gapPx: 28,
    topPaddingPx: 24,
    spreadGapPx: 24,
    ...overrides,
  });
}

test("DOCX view modes normalise to continuous, single or spread", () => {
  assert.deepEqual(DOCX_VIEW_MODES, ["continuous", "single", "spread"]);
  assert.equal(normalizeDocxViewMode(undefined), "continuous");
  assert.equal(normalizeDocxViewMode("bogus"), "continuous");
  assert.equal(normalizeDocxViewMode("single"), "single");
  assert.equal(normalizeDocxViewMode("spread"), "spread");
  assert.equal(isDocxSpreadMode("spread"), true);
  assert.equal(isDocxSpreadMode("single"), false);
  assert.equal(normalizeDocxFitMode("page"), "page");
  assert.equal(normalizeDocxFitMode("width"), "width");
  assert.equal(normalizeDocxFitMode("none"), null);
  assert.equal(normalizeDocxFitMode(undefined), null);
});

test("DOCX pages pair up into spread rows", () => {
  assert.deepEqual(docxPageRows(0, "spread"), []);
  assert.deepEqual(docxPageRows(1, "spread"), [[0]]);
  assert.deepEqual(docxPageRows(4, "spread"), [[0, 1], [2, 3]]);
  assert.deepEqual(docxPageRows(5, "spread"), [[0, 1], [2, 3], [4]]);
  assert.deepEqual(docxPageRows(3, "continuous"), [[0], [1], [2]]);
  assert.deepEqual(docxPageRows(3, "single"), [[0], [1], [2]]);
});

test("continuous layout keeps the single-column page stack", () => {
  const layout = layoutFor(3, "continuous");
  assert.equal(layout.mode, "continuous");
  assert.deepEqual(layout.pageOffsets.map((offset) => Math.round(offset)), [24, 352, 680]);
  assert.deepEqual(layout.pageLefts, [0, 0, 0]);
  assert.equal(Math.round(layout.totalHeightPx), 1008);
  assert.equal(Math.round(layout.maxRowWidthPx), 200);
  assert.deepEqual(layout.pageRows, [0, 1, 2]);
  assert.equal(docxRowIndexForPage(layout, 2), 2);
});

test("single layout reuses the same scroll space as continuous", () => {
  const continuous = layoutFor(3, "continuous");
  const single = layoutFor(3, "single");
  assert.deepEqual(single.pageOffsets, continuous.pageOffsets);
  assert.deepEqual(single.pageHeights, continuous.pageHeights);
  assert.equal(single.totalHeightPx, continuous.totalHeightPx);
});

test("spread layout places both pages of a row side by side with one row top", () => {
  const layout = layoutFor(5, "spread");
  assert.equal(layout.rows.length, 3);
  const [firstRow, secondRow, lastRow] = layout.rows;
  assert.deepEqual(firstRow.pageIndexes, [0, 1]);
  assert.equal(firstRow.widthPx, 200 + 24 + 200, "one spread is two pages plus the column gap");
  assert.equal(firstRow.pageLeftsPx[0], 0);
  assert.equal(firstRow.pageLeftsPx[1], 224);
  assert.equal(layout.pageOffsets[0], layout.pageOffsets[1], "a spread shares one row top");
  assert.deepEqual(secondRow.pageIndexes, [2, 3]);
  assert.equal(secondRow.topPx, firstRow.topPx + 300 + 28, "rows keep the vertical page gap");
  assert.deepEqual(lastRow.pageIndexes, [4]);
  assert.equal(lastRow.widthPx, firstRow.widthPx, "a trailing lone page keeps the spread width");
  assert.equal(lastRow.pageLeftsPx[0], 112, "the lone page is centered on the spread row");
  assert.equal(layout.totalHeightPx, 24 + 3 * (300 + 28), "the scroll length keeps the trailing page gap");
  assert.equal(layout.maxRowWidthPx, 424);
  assert.deepEqual(layout.pageRows, [0, 0, 1, 1, 2]);
});

test("spread geometry follows different page sizes and zoom", () => {
  const sizes = [
    { widthPx: 200, heightPx: 300 },
    { widthPx: 400, heightPx: 200 },
    { widthPx: 200, heightPx: 300 },
  ];
  const layout = resolveDocxViewLayout({
    pageCount: 3,
    mode: "spread",
    pageSize: (pageIndex) => sizes[pageIndex],
    zoom: 2,
    gapPx: 28,
    topPaddingPx: 24,
    spreadGapPx: 24,
  });
  assert.equal(layout.rows[0].heightPx, 600, "a landscape page does not shrink its portrait neighbour");
  assert.equal(layout.pageHeights[0], 600);
  assert.equal(layout.pageHeights[1], 400);
  assert.equal(layout.pageOffsets[0], layout.pageOffsets[1]);
  assert.equal(layout.rows[1].topPx, 24 + 600 + 28);
  assert.deepEqual(layout.pageWidths.map((width) => Math.round(width)), [400, 800, 400]);
});

test("print layout stacks every page even in spread mode", () => {
  const layout = layoutFor(4, "spread", { print: true });
  assert.equal(layout.mode, "continuous");
  assert.deepEqual(layout.pageRows, [0, 1, 2, 3]);
  assert.equal(layout.rows.length, 4);
});

test("page navigation resolves rows and snapping targets", () => {
  const layout = layoutFor(5, "spread");
  // The viewport centre decides the row: 24 + 300/2 is on the first spread.
  assert.equal(docxRowIndexAtOffset(layout.rows, 24, 300), 0);
  assert.equal(docxRowIndexAtOffset(layout.rows, 24 + 328, 300), 1);
  assert.equal(docxSnapTopForRow(layout.rows, 24 + 328 - 40, 300), 24 + 328);
  assert.equal(docxSnapTopForRow([], 0, 300), undefined);
  assert.equal(docxSpreadStepPage(5, 0, 1), 2);
  assert.equal(docxSpreadStepPage(5, 1, 1), 2, "stepping from a right page moves by one spread");
  assert.equal(docxSpreadStepPage(5, 2, -1), 0);
  assert.equal(docxSpreadStepPage(5, 4, 1), 4, "the last lone page is the floor");
});

test("rows in the viewport include both pages of a spread", () => {
  const layout = layoutFor(5, "spread");
  const rows = docxRowsInViewport(layout.rows, { startPx: 30, endPx: 640 });
  assert.deepEqual(rows.map((row) => row.rowIndex), [0, 1]);
  const pages = rows.flatMap((row) => [...row.pageIndexes]);
  assert.deepEqual(pages, [0, 1, 2, 3]);
});

test("hit testing resolves the left/right page of a spread by x", () => {
  const layout = layoutFor(5, "spread");
  const contentWidth = 424;
  // Row 0 spans x 0..424 with page 0 at 0..200 and page 1 at 224..424.
  assert.equal(docxPageAtPoint(layout.rows, { x: 100, y: 100, contentWidth }), 0);
  assert.equal(docxPageAtPoint(layout.rows, { x: 300, y: 100, contentWidth }), 1);
  // A pointer in the spread gap snaps to the nearest page centre.
  assert.equal(docxPageAtPoint(layout.rows, { x: 212, y: 100, contentWidth }), 0);
  assert.equal(docxPageAtPoint(layout.rows, { x: 224, y: 100, contentWidth }), 1);
  // Content wider than the row centers it: the same x now lands left of both.
  const wide = docxPageAtPoint(layout.rows, { x: 400, y: 100, contentWidth: 824 });
  assert.equal(wide, 0);
  // The y axis still selects the row; page 5 is alone on the last row.
  assert.equal(docxPageAtPoint(layout.rows, { x: 200, y: 700, contentWidth }), 4);
  assert.equal(docxPageAtPoint([], { x: 0, y: 0, contentWidth: 0 }), undefined);
});
