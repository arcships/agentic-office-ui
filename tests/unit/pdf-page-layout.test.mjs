import assert from "node:assert/strict";
import test from "node:test";

const pdf = await import(
  new URL("../../packages/vue-pdf/dist/index.js", import.meta.url).href,
);

const {
  PDF_PAGE_LAYOUTS,
  isPdfSpreadLayout,
  normalizePdfPageLayout,
  pdfPageAtPoint,
  pdfPageIndexAtOffset,
  pdfPageLayoutRows,
  pdfRowIndexAtOffset,
  pdfRowsInViewport,
  pdfSnapTopForRow,
  pdfSpreadPagePair,
  pdfSpreadStepPage,
  resolvePdfPageLayoutMetrics,
} = pdf;

const PAGE = { width: 600, height: 800 };

function metricsFor(pageCount, layout, overrides = {}) {
  return resolvePdfPageLayoutMetrics({
    pageCount,
    layout,
    pageSize: () => PAGE,
    rowGap: 8,
    columnGap: 8,
    paddingTop: 16,
    ...overrides,
  });
}

test("PDF page layouts normalise to continuous, single or spread", () => {
  assert.deepEqual(PDF_PAGE_LAYOUTS, ["continuous", "single", "spread"]);
  assert.equal(normalizePdfPageLayout(undefined), "continuous");
  assert.equal(normalizePdfPageLayout("nope"), "continuous");
  assert.equal(normalizePdfPageLayout("single"), "single");
  assert.equal(normalizePdfPageLayout("spread"), "spread");
  assert.equal(isPdfSpreadLayout("spread"), true);
  assert.equal(isPdfSpreadLayout("continuous"), false);
});

test("PDF pages pair into facing spreads", () => {
  assert.deepEqual(pdfPageLayoutRows(0, "spread"), []);
  assert.deepEqual(pdfPageLayoutRows(5, "spread"), [[0, 1], [2, 3], [4]]);
  assert.deepEqual(pdfPageLayoutRows(3, "continuous"), [[0], [1], [2]]);
  assert.deepEqual(pdfSpreadPagePair(5, 0), [0, 1]);
  assert.deepEqual(pdfSpreadPagePair(5, 1), [0, 1]);
  assert.deepEqual(pdfSpreadPagePair(5, 2), [2, 3]);
  assert.deepEqual(pdfSpreadPagePair(5, 4), [4], "an odd last page has no facing page");
  assert.deepEqual(pdfSpreadPagePair(0, 0), []);
  assert.equal(pdfSpreadStepPage(5, 1, 1), 2);
  assert.equal(pdfSpreadStepPage(5, 2, -1), 0);
  assert.equal(pdfSpreadStepPage(5, 4, 1), 4);
  assert.equal(pdfSpreadStepPage(0, 3, 1), 0);
});

test("continuous metrics reproduce the stacked page offsets the surface already used", () => {
  const metrics = metricsFor(3, "continuous");
  assert.deepEqual(metrics.pageOffsets, [16, 824, 1632]);
  assert.deepEqual(metrics.pageHeights, [800, 800, 800]);
  assert.deepEqual(metrics.pageGridColumns, [1, 1, 1]);
  assert.deepEqual(metrics.pageGridRows, [1, 2, 3]);
  assert.deepEqual(metrics.columnWidths, [600]);
  assert.equal(metrics.totalHeightPx, 16 + 3 * 800 + 3 * 8, "the scroll length keeps the trailing row gap");
  assert.equal(metrics.maxRowWidthPx, 600);
});

test("single metrics keep one page per row so the surface can snap", () => {
  const continuous = metricsFor(3, "continuous");
  const single = metricsFor(3, "single");
  assert.deepEqual(single.pageOffsets, continuous.pageOffsets);
  assert.deepEqual(single.rows.map((row) => row.pageIndexes), [[0], [1], [2]]);
  assert.equal(pdfRowIndexAtOffset(single.rows, 0, 800), 0);
  assert.equal(pdfRowIndexAtOffset(single.rows, 824, 800), 1);
  assert.equal(pdfSnapTopForRow(single.rows, 824 + 120, 800), 824);
  assert.equal(pdfSnapTopForRow([], 0, 800), undefined);
});

test("spread metrics put two facing pages on one row with correct page numbers", () => {
  const metrics = metricsFor(3, "spread");
  assert.equal(metrics.columnWidths.length, 2);
  assert.deepEqual(metrics.columnWidths, [600, 600]);
  assert.deepEqual(metrics.pageGridColumns, [1, 2, 0]);
  assert.deepEqual(metrics.pageGridRows, [1, 1, 2]);
  assert.equal(metrics.pageOffsets[0], 16);
  assert.equal(metrics.pageOffsets[1], 16, "both pages of a spread share the row top");
  assert.equal(metrics.rows[0].widthPx, 600 + 8 + 600);
  assert.equal(metrics.pageLefts[0], 0);
  assert.equal(metrics.pageLefts[1], 608);
  assert.equal(metrics.rows[1].pageIndexes.length, 1, "the odd last page is alone");
  assert.equal(metrics.rows[1].pageLeftsPx[0], 304, "the lone page is centered on the spread");
  assert.equal(pdfPageIndexAtOffset(metrics.rows, 16, 800), 0);
  assert.equal(pdfPageIndexAtOffset(metrics.rows, 824 + 16, 800), 2, "a spread reports its left page");
});

test("spread columns follow CSS grid max-content sizing", () => {
  const sizes = [
    { width: 600, height: 800 },
    { width: 1000, height: 400 },
    { width: 500, height: 700 },
    { width: 900, height: 300 },
  ];
  const metrics = resolvePdfPageLayoutMetrics({
    pageCount: 4,
    layout: "spread",
    pageSize: (pageIndex) => sizes[pageIndex],
    rowGap: 8,
    columnGap: 8,
    paddingTop: 0,
  });
  assert.deepEqual(metrics.columnWidths, [600, 1000]);
  assert.equal(metrics.rows[0].widthPx, 1608);
  assert.equal(metrics.pageLefts[0], 0);
  assert.equal(metrics.pageLefts[1], 608);
  // Row 2 pages are centered inside their column.
  assert.equal(metrics.pageLefts[2], 50);
  assert.equal(metrics.pageLefts[3], 600 + 8 + 50);
  assert.equal(metrics.rows[0].heightPx, 800);
  assert.equal(metrics.rows[1].topPx, 808);
});

test("PDF print layout keeps one page per row even in spread mode", () => {
  const metrics = metricsFor(4, "spread", { print: true });
  assert.equal(metrics.layout, "continuous");
  assert.deepEqual(metrics.pageGridColumns, [1, 1, 1, 1]);
  assert.equal(metrics.rows.length, 4);
});

test("hit testing resolves the facing page by x", () => {
  const metrics = metricsFor(3, "spread");
  const contentWidth = metrics.maxRowWidthPx;
  assert.equal(pdfPageAtPoint(metrics.rows, { x: 100, y: 100, contentWidth }), 0);
  assert.equal(pdfPageAtPoint(metrics.rows, { x: 900, y: 100, contentWidth }), 1);
  // The gap between the pages snaps to the nearer page centre.
  assert.equal(pdfPageAtPoint(metrics.rows, { x: 604, y: 100, contentWidth }), 0);
  assert.equal(pdfPageAtPoint(metrics.rows, { x: 607, y: 100, contentWidth }), 1);
  // A wider scroll content centers the whole spread.
  assert.equal(pdfPageAtPoint(metrics.rows, { x: 0, y: 100, contentWidth: 2000 }), 0);
  // The last row holds the lone page.
  assert.equal(pdfPageAtPoint(metrics.rows, { x: 400, y: 900, contentWidth }), 2);
  assert.equal(pdfPageAtPoint([], { x: 0, y: 0, contentWidth: 0 }), undefined);
});

test("rows in the viewport include both pages of a spread row", () => {
  const metrics = metricsFor(5, "spread");
  const rows = pdfRowsInViewport(metrics.rows, { startPx: 16, endPx: 900 });
  assert.deepEqual(rows.map((row) => row.rowIndex), [0, 1]);
  assert.deepEqual(rows.flatMap((row) => [...row.pageIndexes]), [0, 1, 2, 3]);
});
