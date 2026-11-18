import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";

// The module is browser-independent, so exercise the TypeScript source directly
// instead of relying on a built bundle.
const requireFromPdf = createRequire(
  new URL("../../packages/vue-pdf/package.json", import.meta.url),
);
const typescript = requireFromPdf("typescript");
const viewportSource = readFileSync(
  new URL("../../packages/vue-pdf/src/pdf/pdf-page-viewport.ts", import.meta.url),
  "utf8",
);
const viewport = await import(`data:text/javascript;base64,${Buffer.from(
  typescript.transpileModule(viewportSource, {
    compilerOptions: {
      module: typescript.ModuleKind.ES2022,
      target: typescript.ScriptTarget.ES2022,
    },
  }).outputText,
).toString("base64")}`);

const {
  computePageOffsets,
  pageIndexAtPosition,
  pageRangeContains,
  pageRangeEquals,
  planPageRenders,
  visiblePageRange,
} = viewport;

test("computePageOffsets stacks page heights with the stack gap and padding", () => {
  assert.deepEqual(computePageOffsets([]), []);
  assert.deepEqual(
    computePageOffsets([800, 900, 1000], { gap: 8, paddingTop: 16 }),
    [16, 824, 1732],
  );
  assert.deepEqual(computePageOffsets([100, 100]), [0, 100]);
  assert.deepEqual(
    computePageOffsets([Number.NaN, -40, 60], { gap: 8, paddingTop: 16 }),
    [16, 24, 32],
    "broken measurements collapse instead of producing NaN offsets",
  );
});

test("pageIndexAtPosition resolves the page under a scroll position", () => {
  const heights = [800, 800, 800];
  const offsets = computePageOffsets(heights, { gap: 8, paddingTop: 16 });
  assert.equal(pageIndexAtPosition(offsets, heights, 0), 0, "above the first page");
  assert.equal(pageIndexAtPosition(offsets, heights, 16), 0);
  assert.equal(pageIndexAtPosition(offsets, heights, 815), 0);
  assert.equal(pageIndexAtPosition(offsets, heights, 820), 0, "the gap belongs to the previous page");
  assert.equal(pageIndexAtPosition(offsets, heights, 824), 1);
  assert.equal(pageIndexAtPosition(offsets, heights, 5000), 2, "past the last page");
  assert.equal(pageIndexAtPosition([], [], 500), 0, "an empty document resolves to page 0");
  assert.equal(pageIndexAtPosition(offsets, heights, Number.NaN), 0);
});

test("visiblePageRange keeps the window bounded to the viewport plus overscan", () => {
  assert.deepEqual(
    visiblePageRange([], [], { scrollTop: 0, viewportHeight: 600 }),
    { first: 0, last: -1 },
    "an empty document renders nothing",
  );

  const heights = Array.from({ length: 40 }, (_, pageIndex) => 800 + pageIndex);
  const offsets = computePageOffsets(heights, { gap: 8, paddingTop: 16 });

  assert.deepEqual(
    visiblePageRange(offsets, heights, { scrollTop: 0, viewportHeight: 600, overscanPages: 1 }),
    { first: 0, last: 1 },
    "the first viewport renders two pages plus the clamped top overscan",
  );
  assert.deepEqual(
    visiblePageRange(offsets, heights, { scrollTop: 0, viewportHeight: 0, overscanPages: 1 }),
    { first: 0, last: 2 },
    "an unmeasured viewport rasterizes a bounded window instead of every page",
  );
  assert.deepEqual(
    visiblePageRange(offsets, heights, { scrollTop: 0, viewportHeight: 0, overscanPages: 2 }),
    { first: 0, last: 4 },
  );
  assert.deepEqual(
    visiblePageRange(offsets, heights, {
      scrollTop: offsets[20],
      viewportHeight: 600,
      overscanPages: 1,
    }),
    { first: 19, last: 21 },
    "a mid-document viewport keeps one page of overscan on each side",
  );
  assert.deepEqual(
    visiblePageRange(offsets, heights, { scrollTop: -500, viewportHeight: 600, overscanPages: 1 }),
    { first: 0, last: 1 },
    "negative scroll offsets clamp to the top",
  );
  assert.deepEqual(
    visiblePageRange(offsets, heights, { scrollTop: 10_000_000, viewportHeight: 600, overscanPages: 1 }),
    { first: 39, last: 39 },
    "a scroll offset past the content clamps to the last page",
  );

  const shortDocument = Array.from({ length: 3 }, () => 400);
  const shortOffsets = computePageOffsets(shortDocument, { gap: 8, paddingTop: 16 });
  assert.deepEqual(
    visiblePageRange(shortOffsets, shortDocument, { scrollTop: 0, viewportHeight: 5000, overscanPages: 1 }),
    { first: 0, last: 2 },
    "a viewport taller than the document renders every page it shows",
  );
  assert.deepEqual(
    visiblePageRange(shortOffsets, shortDocument, { scrollTop: 0, viewportHeight: 5000, overscanPages: 0 }),
    { first: 0, last: 2 },
  );
});

test("pageRange helpers compare inclusive ranges", () => {
  assert.equal(pageRangeContains({ first: 3, last: 5 }, 3), true);
  assert.equal(pageRangeContains({ first: 3, last: 5 }, 5), true);
  assert.equal(pageRangeContains({ first: 3, last: 5 }, 6), false);
  assert.equal(pageRangeContains({ first: 0, last: -1 }, 0), false, "the empty range contains nothing");
  assert.equal(pageRangeEquals({ first: 1, last: 2 }, { first: 1, last: 2 }), true);
  assert.equal(pageRangeEquals({ first: 1, last: 2 }, { first: 1, last: 3 }), false);
});

function pendingSlots(pageCount) {
  return Array.from({ length: pageCount }, () => ({ state: "pending", hasBitmap: false }));
}

test("planPageRenders rasterizes only pages entering the viewport window", () => {
  const slots = pendingSlots(40);
  const firstPlan = planPageRenders(slots, { first: 0, last: 2 });
  assert.deepEqual(firstPlan.render, [0, 1, 2], "the initial window is rasterized");
  assert.deepEqual(firstPlan.release, []);
  assert.ok(firstPlan.render.length < slots.length, "a load never rasterizes the whole document");

  for (const pageIndex of firstPlan.render) {
    slots[pageIndex] = { state: "ready", hasBitmap: true };
  }
  assert.deepEqual(
    planPageRenders(slots, { first: 0, last: 2 }),
    { render: [], release: [] },
    "an unchanged window is not rasterized again",
  );

  const scrolledPlan = planPageRenders(slots, { first: 20, last: 22 });
  assert.deepEqual(scrolledPlan.render, [20, 21, 22], "only the new window is rasterized");
  assert.deepEqual(scrolledPlan.release, [0, 1, 2], "bitmaps that left the window are released");
});

test("planPageRenders releases every held bitmap outside the window", () => {
  const slots = pendingSlots(6).map(() => ({ state: "ready", hasBitmap: true }));
  assert.deepEqual(planPageRenders(slots, { first: 0, last: -1 }), {
    render: [],
    release: [0, 1, 2, 3, 4, 5],
  });
  assert.deepEqual(
    planPageRenders(slots, { first: 2, last: 3 }),
    { render: [], release: [0, 1, 4, 5] },
    "a document-wide release keeps the viewport window intact",
  );
});

test("planPageRenders retries failed pages and re-renders stale loading slots", () => {
  const slots = pendingSlots(5);
  slots[0] = { state: "error", hasBitmap: false };
  slots[1] = { state: "loading", hasBitmap: false };
  slots[2] = { state: "ready", hasBitmap: true };
  slots[3] = { state: "error", hasBitmap: false };
  assert.deepEqual(planPageRenders(slots, { first: 0, last: 2 }).render, [0, 1]);
  assert.deepEqual(
    planPageRenders(slots, { first: 2, last: 4 }).release,
    [0, 1],
    "failed and abandoned slots leave the window as pending",
  );
  const missing = pendingSlots(3);
  missing[1] = undefined;
  assert.deepEqual(
    planPageRenders(missing, { first: 0, last: 2 }).render,
    [0, 1, 2],
    "slots without a record are treated as renderable",
  );
});

test("planPageRenders re-rasterizes stale bitmaps without releasing them", () => {
  const slots = pendingSlots(5);
  slots[0] = { state: "ready", hasBitmap: true };
  slots[1] = { state: "ready", hasBitmap: true, stale: true };
  slots[2] = { state: "ready", hasBitmap: true };
  assert.deepEqual(
    planPageRenders(slots, { first: 0, last: 2 }),
    { render: [1], release: [] },
    "a bitmap held across a zoom/rotation/DPR change is replaced in place",
  );
  assert.deepEqual(
    planPageRenders(slots, { first: 2, last: 3 }).release,
    [0, 1],
    "stale bitmaps still leave the window like any other bitmap",
  );
  slots[1] = { state: "ready", hasBitmap: true, stale: false };
  assert.deepEqual(
    planPageRenders(slots, { first: 0, last: 2 }),
    { render: [], release: [] },
    "an up-to-date bitmap is not rasterized again",
  );
});

test("a full scroll through a document never plans a whole-document re-rasterization", () => {
  const pageCount = 60;
  const heights = Array.from({ length: pageCount }, () => 800);
  const offsets = computePageOffsets(heights, { gap: 8, paddingTop: 16 });
  const slots = pendingSlots(pageCount);
  let renderPasses = 0;
  for (let scrollTop = 0; scrollTop <= offsets[pageCount - 1]; scrollTop += 600) {
    const range = visiblePageRange(offsets, heights, { scrollTop, viewportHeight: 600, overscanPages: 1 });
    const plan = planPageRenders(slots, range);
    renderPasses += plan.render.length;
    assert.ok(
      plan.render.length <= 4,
      `a ${range.first}-${range.last} window rasterized ${plan.render.length} pages`,
    );
    for (const pageIndex of plan.render) slots[pageIndex] = { state: "ready", hasBitmap: true };
    for (const pageIndex of plan.release) slots[pageIndex] = { state: "pending", hasBitmap: false };
    const held = slots.filter((slot) => slot.hasBitmap).length;
    assert.ok(held <= 4, `the bitmap budget tracked the viewport instead of the document (${held})`);
  }
  assert.deepEqual(
    planPageRenders(slots, { first: 0, last: -1 }).release,
    [pageCount - 3, pageCount - 2, pageCount - 1],
    "the final window is released",
  );
  assert.ok(renderPasses <= pageCount + 4, "scrolling rasterizes each page about once");
});
