// Measure-driven pagination iteration (issue #99).
//
// Covers the loop itself (bounded iterations, convergence, oscillation guard,
// advisory documentPageCount target), reference-stable page reuse, scroll
// anchoring, and the legal/invoice acceptance fixtures.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

// pretext needs a text-measurement canvas; Node has none, so provide a
// deterministic shim before the package (and pretext) loads.
const originalOffscreenCanvas = globalThis.OffscreenCanvas;
globalThis.OffscreenCanvas = class {
  constructor(width, height) {
    this.width = width;
    this.height = height;
  }
  getContext() {
    return {
      font: "12px sans-serif",
      measureText(text) {
        const fontSize = Number.parseFloat(
          /(\d+(?:\.\d+)?)px/.exec(this.font)?.[1] ?? "12"
        );
        return { width: text.length * fontSize * 0.5 };
      },
    };
  }
};
process.on("exit", () => {
  if (originalOffscreenCanvas === undefined) {
    delete globalThis.OffscreenCanvas;
  }
});

const core = await import(
  new URL("../../packages/docx-core/dist/index.js", import.meta.url).href,
);
const pageCountBaselines = JSON.parse(
  readFileSync(
    new URL("../fixtures/docx/page-count-baselines.json", import.meta.url),
    "utf8",
  ),
);

function paragraph(text) {
  return {
    type: "paragraph",
    children: [{ type: "text", text }],
  };
}

function table(cellsByRow) {
  return {
    type: "table",
    rows: cellsByRow.map((cells, rowIndex) => ({
      id: `row-${rowIndex}`,
      cells: cells.map((text, cellIndex) => ({
        id: `row-${rowIndex}-cell-${cellIndex}`,
        nodes: [paragraph(text)],
      })),
    })),
  };
}

function modelWith(nodes, metadata = {}) {
  return {
    nodes,
    metadata: {
      sourceParts: 0,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
      ...metadata,
    },
  };
}

async function loadFixtureModel(fixture) {
  const file = readFileSync(
    new URL(`../../apps/demo/public/samples/${fixture}`, import.meta.url),
  );
  const bytes = file.buffer.slice(
    file.byteOffset,
    file.byteOffset + file.byteLength,
  );
  const { model } = await core.wasmBuildDocModelFromBytes(bytes);
  return model;
}

function pageSegmentationKey(page) {
  return page.blocks
    .map((block) =>
      [
        block.kind,
        block.source?.nodeIndex ?? "",
        block.source?.columnIndex ?? 0,
        block.paragraphLineRange?.startLineIndex ?? "",
        block.paragraphLineRange?.endLineIndex ?? "",
        block.tableRowRange?.startRowIndex ?? "",
        block.tableRowRange?.endRowIndex ?? "",
      ].join(":"),
    )
    .join(",");
}

function pageContentHeightPx(page) {
  const bottom = page.blocks.reduce(
    (value, block) => Math.max(value, block.framePx.y + block.framePx.height),
    page.contentBoxPx.y,
  );
  return Math.max(1, Math.round(bottom - page.contentBoxPx.y));
}

function pageContentBottomPx(page) {
  return page.blocks.reduce(
    (value, block) => Math.max(value, block.framePx.y + block.framePx.height),
    page.contentBoxPx.y,
  );
}

function measureSnapshot(model) {
  const layoutDefaults = core.resolveDocumentLayoutDefaults(model.metadata);
  return (snapshot) => {
    const paragraphHeightsPxByNodeIndex = {};
    for (const page of snapshot.pages) {
      for (const block of page.blocks) {
        const nodeIndex = block.source?.nodeIndex;
        if (nodeIndex === undefined || block.kind !== "paragraph") continue;
        const node = model.nodes[nodeIndex];
        // Frame heights suppress space-before at the top of a page. A DOM
        // measurement (offsetHeight + margins) does not, so add it back to
        // model what the viewer feeds into the segmenter.
        const atPageTop = Math.abs(block.framePx.y - page.contentBoxPx.y) < 1;
        const beforeSpacingPx =
          atPageTop && node?.type === "paragraph"
            ? core.paragraphBeforeSpacingPx(node, layoutDefaults)
            : 0;
        paragraphHeightsPxByNodeIndex[nodeIndex] = Math.round(
          block.framePx.height + beforeSpacingPx,
        );
      }
    }
    return {
      pageContentHeightsPx: snapshot.pages.map(pageContentHeightPx),
      feedback: { paragraphHeightsPxByNodeIndex },
    };
  };
}

function runFixtureLoop(model, options, targetPageCount) {
  return core.iteratePaginationWithMeasurements({
    buildPages: (feedback) =>
      core.buildLayoutSnapshot(model, {
        ...options,
        measurementFeedback: feedback,
      }),
    measurePages: measureSnapshot(model),
    pageIdentityKeys: (snapshot) => snapshot.pages.map(pageSegmentationKey),
    maxIterations: 3,
    tolerancePx: 1,
    targetPageCount,
  });
}

test("measurement gate converges within the iteration bound", () => {
  const loop = new core.PaginationMeasurementLoop({
    maxIterations: 3,
    tolerancePx: 1,
    targetPageCount: 3,
  });
  const feedback = { paragraphHeightsPxByNodeIndex: { 0: 120 } };
  const observation = {
    pageIdentityKeys: ["page-a", "page-b", "page-c"],
    pageContentHeightsPx: [900, 800, 700],
    feedback,
  };

  const first = loop.observe(observation);
  assert.equal(first.action, "repaginate");
  assert.equal(first.iteration, 1);
  assert.equal(first.targetReached, true, "three rendered pages match target 3");

  const second = loop.observe(observation);
  assert.equal(second.action, "converged");
  assert.equal(second.iteration, 1);
  assert.equal(loop.state, "converged");

  // Re-measuring the same render after convergence must not start a new pass.
  const third = loop.observe(observation);
  assert.equal(third.action, "converged");
  assert.equal(loop.iteration, 1);
});

test("measurement gate stops on oscillation instead of looping", () => {
  const loop = new core.PaginationMeasurementLoop({ maxIterations: 6 });
  const first = loop.observe({
    pageIdentityKeys: ["a", "b"],
    pageContentHeightsPx: [100, 100],
    feedback: { paragraphHeightsPxByNodeIndex: { 0: 100 } },
  });
  assert.equal(first.action, "repaginate");

  const second = loop.observe({
    pageIdentityKeys: ["a", "b"],
    pageContentHeightsPx: [200, 200],
    feedback: { paragraphHeightsPxByNodeIndex: { 0: 100 } },
  });
  assert.equal(second.action, "repaginate");

  const third = loop.observe({
    pageIdentityKeys: ["a", "b"],
    pageContentHeightsPx: [100, 100],
    feedback: { paragraphHeightsPxByNodeIndex: { 0: 100 } },
  });
  assert.equal(third.action, "oscillating");
  assert.ok(loop.iteration < 6);
});

test("measurement gate exhausts after the configured number of repaginations", () => {
  const loop = new core.PaginationMeasurementLoop({ maxIterations: 2 });
  let height = 100;
  let decision;
  for (let pass = 0; pass < 6; pass += 1) {
    height += 10;
    decision = loop.observe({
      pageIdentityKeys: [`page-${pass}`],
      pageContentHeightsPx: [height],
      feedback: { paragraphHeightsPxByNodeIndex: { [pass]: 50 } },
    });
    if (decision.action !== "repaginate") break;
  }
  assert.equal(decision.action, "exhausted");
  assert.equal(loop.iteration, 2);
});

test("documentPageCount is a soft target: legal-contract stays nine pages", async () => {
  assert.deepEqual(core.resolvePaginationPageCountTarget(undefined), {
    enforced: false,
  });
  assert.deepEqual(core.resolvePaginationPageCountTarget(0), {
    enforced: false,
  });
  assert.deepEqual(core.resolvePaginationPageCountTarget(9), {
    targetPageCount: 9,
    enforced: false,
  });

  const model = await loadFixtureModel("legal-contract.docx");
  assert.equal(model.metadata.documentPageCount, 1, "fixture caches a stale <Pages>1</Pages>");
  const layout = core.resolveDocumentLayout(model);
  const options = {
    pageWidth: layout.pageWidthPx,
    pageHeight: layout.pageHeightPx,
    margin: layout.marginsPx.top,
  };

  const overridden = core.buildLayoutSnapshot(model, {
    ...options,
    pageCountTarget: 2,
  });
  assert.equal(overridden.pageCountTarget.targetPageCount, 2);
  assert.equal(overridden.pageCountTarget.enforced, false);
  assert.equal(
    overridden.pages.length,
    9,
    "the advisory override must not reshape the document",
  );

  const result = runFixtureLoop(model, options, model.metadata.documentPageCount);
  assert.equal(result.pageCount, 9, "the rendered page count wins over the cached target");
  assert.equal(result.targetPageCount, 1);
  assert.equal(result.targetReached, false);
  assert.equal(result.converged, true);
  assert.ok(result.iterations <= 3);
  assert.deepEqual(result.pages.pageCountTarget, {
    targetPageCount: 1,
    enforced: false,
  });
});

test("measurement feedback includes measured table row heights", () => {
  const rows = Array.from({ length: 10 }, (_, index) => [`Row ${index}`]);
  const model = modelWith([table(rows)]);
  const options = { pageWidth: 816, pageHeight: 1056, margin: 72 };
  const estimated = core.paginateDocumentLayout(model, {
    pageWidthPx: options.pageWidth,
    pageHeightPx: options.pageHeight,
    marginPx: options.margin,
  });
  const measured = core.paginateDocumentLayout(model, {
    pageWidthPx: options.pageWidth,
    pageHeightPx: options.pageHeight,
    marginPx: options.margin,
    measurementFeedback: {
      tableRowHeightsPxByNodeIndex: {
        0: rows.map(() => 200),
      },
    },
  });

  assert.equal(estimated.pages.length, 1);
  assert.ok(
    measured.pages.length > estimated.pages.length,
    "measured rows taller than the estimate must split the table over more pages",
  );
  assert.ok(
    measured.pages.some((page) =>
      page.blocks.some(
        (block) => block.kind === "table" && block.tableRowRange !== undefined,
      ),
    ),
    "measured rows must split the table across pages",
  );
});

test("partial table row measurements are ignored instead of zero-filling rows", () => {
  const rows = Array.from({ length: 6 }, (_, index) => [`Row ${index}`]);
  const model = modelWith([table(rows)]);
  const options = { pageWidth: 816, pageHeight: 1056, margin: 72 };
  const baseline = core.paginateDocumentLayout(model, {
    pageWidthPx: options.pageWidth,
    pageHeightPx: options.pageHeight,
    marginPx: options.margin,
  });
  const partial = core.paginateDocumentLayout(model, {
    pageWidthPx: options.pageWidth,
    pageHeightPx: options.pageHeight,
    marginPx: options.margin,
    measurementFeedback: {
      // Only the rows of the first page slice were in the DOM.
      tableRowHeightsPxByNodeIndex: { 0: [200, 200] },
    },
  });

  assert.deepEqual(
    partial.pages.map((page) => page.segments),
    baseline.pages.map((page) => page.segments),
    "an incomplete row-height array must not reshape the table",
  );
});

test("reused pages keep their previous arrays when segmentation is unchanged", () => {
  const previousSegments = [[0], [1, 2]];
  const nextSegments = [[0], [1, 3]];
  const previousKeys = ["0", "1-2"];
  const nextKeys = ["0", "1-3"];
  const result = core.reuseUnchangedPaginationItems({
    previous: previousSegments,
    next: nextSegments,
    isUnchanged: (_previous, _next, pageIndex) =>
      previousKeys[pageIndex] === nextKeys[pageIndex],
  });

  assert.equal(result.items[0], previousSegments[0], "unchanged page keeps its array");
  assert.equal(result.items[1], nextSegments[1], "changed page gets the new array");
  assert.deepEqual(result.changedIndexes, [1]);
  assert.deepEqual(result.reusedIndexes, [0]);
});

test("scroll anchoring keeps the same page offset when pages above change", () => {
  const before = [24, 1100, 2176];
  const anchor = core.resolvePaginationScrollAnchor(before, 1200);
  assert.deepEqual(anchor, { pageIndex: 1, offsetPx: 100 });
  assert.equal(core.resolvePaginationScrollTop(before, anchor), 1200);

  // A measurement pass inserted a shorter page before the anchored page.
  const after = [24, 500, 1100, 2176];
  assert.equal(core.resolvePaginationScrollTop(after, anchor), 600);
  assert.equal(
    core.resolvePaginationScrollTop([24], anchor),
    undefined,
    "removing the anchored page leaves scroll position untouched",
  );
  assert.equal(core.resolvePaginationScrollAnchor([], 400), undefined);
});

for (const fixture of ["legal-contract.docx", "invoice-table.docx"]) {
  test(`${fixture}: every page matches the paper and the stored baseline count`, async () => {
    const baseline = pageCountBaselines[fixture];
    const model = await loadFixtureModel(fixture);
    assert.equal(model.metadata.documentPageCount, baseline.embeddedDocumentPageCount);
    const layout = core.resolveDocumentLayout(model);
    const options = {
      pageWidth: layout.pageWidthPx,
      pageHeight: layout.pageHeightPx,
      margin: layout.marginsPx.top,
    };

    const estimate = core.buildLayoutSnapshot(model, options);
    assert.equal(estimate.pageCountTarget.enforced, false);
    assert.equal(
      estimate.pageCountTarget.targetPageCount,
      model.metadata.documentPageCount,
    );

    const first = runFixtureLoop(model, options, model.metadata.documentPageCount);
    const second = runFixtureLoop(model, options, model.metadata.documentPageCount);

    assert.equal(first.converged, true);
    assert.equal(first.oscillationDetected, false);
    assert.ok(first.iterations <= 3);
    assert.equal(first.pageCount, baseline.documentPageCount);
    assert.equal(second.pageCount, baseline.documentPageCount);
    assert.deepEqual(
      first.pages.pages.map(pageSegmentationKey),
      second.pages.pages.map(pageSegmentationKey),
      "repeated renders must produce identical segmentation",
    );

    for (const snapshot of [estimate, first.pages]) {
      assert.equal(snapshot.pages.length, baseline.documentPageCount);
      for (const [pageIndex, page] of snapshot.pages.entries()) {
        assert.ok(
          Math.abs(page.pageSizePx.height - baseline.paperHeightPx) <= 1,
          `${fixture} page ${pageIndex} height ${page.pageSizePx.height} must match the paper`,
        );
        assert.ok(
          pageContentBottomPx(page) <=
            page.contentBoxPx.y + page.contentBoxPx.height + 1,
          `${fixture} page ${pageIndex} content must fit inside the paper content box`,
        );
      }
    }
  });
}
