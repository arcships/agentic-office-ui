import assert from "node:assert/strict";
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

function paragraph(text, options = {}) {
  return {
    type: "paragraph",
    children: [{ type: "text", text }],
    ...(options.style ? { style: options.style } : {}),
    ...(options.sourceXml ? { sourceXml: options.sourceXml } : {}),
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

function columnSectionXml(num = 2, spaceTwips = 720) {
  return (
    `<w:sectPr><w:type w:val="nextPage"/>` +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>` +
    `<w:cols w:num="${num}" w:space="${spaceTwips}"/>` +
    `</w:sectPr>`
  );
}

function columnModel(nodes) {
  return modelWith(nodes, {
    sections: [
      {
        startNodeIndex: 0,
        headerSections: [],
        footerSections: [],
        sectionPropertiesXml: columnSectionXml(),
      },
    ],
  });
}

const PAGE_OPTIONS = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: 96 };

test("real wrap-line provider maps a mid-paragraph lastRenderedPageBreak to the pretext line", () => {
  const text = `A\nB\n${"C".repeat(100)}\nD`;
  const xml =
    `<w:p><w:r><w:t>A</w:t><w:br/><w:t>B</w:t><w:br/></w:r>` +
    `<w:r><w:lastRenderedPageBreak/><w:t>${"C".repeat(100)}</w:t><w:br/></w:r>` +
    `<w:r><w:t>D</w:t></w:r></w:p>`;
  const model = modelWith([paragraph(text, { sourceXml: xml })]);

  // The marker sits before the long line, so the real next-page line is line
  // 2. A character-proportional estimate maps the same offset to line 1.
  const realLineIndex = core.paragraphLastRenderedPageBreakLineIndex(
    model.nodes[0],
    4,
    10000,
  );
  assert.equal(realLineIndex, 2);

  const callbacks = {
    estimateDocNodeHeightPx: () => 80,
    paragraphHasVisibleText: () => true,
    paragraphIsStructuralSectionBreakSpacer: () => false,
    estimateParagraphHeightPx: () => 80,
    estimateParagraphLineHeightPx: () => 20,
    paragraphLineCountWithinWidth: () => 4,
    paragraphWidowControlEnabled: () => false,
    paragraphCanSplitAcrossPages: () => true,
    estimateTableRowHeightsPx: () => [],
  };

  const withProvider = core.buildDocumentPageNodeSegments(
    model,
    400,
    10000,
    callbacks,
    undefined,
    [],
    {
      allowParagraphLineSplitting: false,
      preferLastRenderedParagraphStartBreaks: true,
      paragraphLastRenderedPageBreakLineIndex:
        core.paragraphLastRenderedPageBreakLineIndex,
    },
  );
  assert.deepEqual(
    withProvider.map((page) =>
      page.map((segment) => [
        segment.paragraphLineRange.startLineIndex,
        segment.paragraphLineRange.endLineIndex,
      ]),
    ),
    [[[0, 2]], [[2, 4]]],
  );

  // Without the provider the proportional fallback lands on line 1.
  const withFallback = core.buildDocumentPageNodeSegments(
    model,
    400,
    10000,
    callbacks,
    undefined,
    [],
    {
      allowParagraphLineSplitting: false,
      preferLastRenderedParagraphStartBreaks: true,
    },
  );
  assert.deepEqual(
    withFallback.map((page) =>
      page.map((segment) => [
        segment.paragraphLineRange.startLineIndex,
        segment.paragraphLineRange.endLineIndex,
      ]),
    ),
    [[[0, 1]], [[1, 4]]],
  );
});

test("two-column sections expose real column geometry and metrics", () => {
  const geometries = core.resolvePaginationSectionGeometry(
    columnModel([paragraph("one")]),
    PAGE_OPTIONS,
  );
  const geometry = geometries[0];
  assert.equal(geometry.columnCount, 2);
  assert.equal(geometry.columnGapPx, 48);
  assert.equal(geometry.columnWidthsPx.length, 2);
  // Body width 816 - 96 - 96 = 624; gap 48 -> 288 per column.
  assert.deepEqual(geometry.columnWidthsPx, [288, 288]);

  const metrics = core.paginateDocumentLayout(
    columnModel([paragraph("one")]),
    { ...PAGE_OPTIONS, measurementFeedback: { paragraphHeightsPxByNodeIndex: { 0: 100 } } },
  ).sectionMetrics[0];
  assert.equal(metrics.pageContentWidthPx, 288);
  assert.equal(metrics.pageContentHeightMultiplier, 2);
  assert.equal(metrics.pageContentHeightPx, (1056 - 192) * 2);
});

test("column distribution balances the last page and keeps flow order", () => {
  const model = columnModel([
    paragraph("one"),
    paragraph("two"),
    paragraph("three"),
    paragraph("four"),
  ]);
  const layout = core.paginateDocumentLayout(model, {
    ...PAGE_OPTIONS,
    measurementFeedback: {
      paragraphHeightsPxByNodeIndex: { 0: 100, 1: 100, 2: 100, 3: 100 },
    },
  });

  assert.equal(layout.pages.length, 1);
  const page = layout.pages[0];
  assert.equal(page.columns.length, 2);
  assert.deepEqual(
    page.columns.map((column) => column.map((segment) => segment.nodeIndex)),
    [
      [0, 1],
      [2, 3],
    ],
  );
  assert.deepEqual(
    page.blocks.map((block) => block.columnIndex),
    [0, 0, 1, 1],
  );

  const snapshot = core.buildLayoutSnapshot(model, {
    pageWidth: 816,
    pageHeight: 1056,
    margin: 96,
    measurementFeedback: {
      paragraphHeightsPxByNodeIndex: { 0: 100, 1: 100, 2: 100, 3: 100 },
    },
  });
  assert.equal(snapshot.pages[0].columnCount, 2);
  assert.equal(snapshot.pages[0].columnGapPx, 48);
  assert.deepEqual(snapshot.pages[0].columnWidthsPx, [288, 288]);
  assert.deepEqual(
    snapshot.pages[0].blocks.map((block) => block.source?.columnIndex),
    [0, 0, 1, 1],
  );
});

test("a paragraph taller than a column is split across columns with contiguous line ranges", () => {
  const text = "word ".repeat(400);
  const model = columnModel([paragraph(text)]);
  const layout = core.paginateDocumentLayout(model, {
    ...PAGE_OPTIONS,
    measurementFeedback: {
      // Taller than one column (864px), shorter than both columns combined.
      paragraphHeightsPxByNodeIndex: { 0: 1000 },
    },
    allowParagraphLineSplitting: true,
  });

  assert.equal(layout.pages.length, 1);
  const page = layout.pages[0];
  assert.equal(page.columns.length, 2);
  const first = page.columns[0].at(-1);
  const second = page.columns[1][0];
  assert.equal(first.nodeIndex, 0);
  assert.equal(second.nodeIndex, 0);
  assert.ok(first.paragraphLineRange);
  assert.ok(second.paragraphLineRange);
  assert.equal(
    first.paragraphLineRange.endLineIndex,
    second.paragraphLineRange.startLineIndex,
  );
  assert.ok(first.paragraphLineRange.endLineIndex < first.paragraphLineRange.totalLineCount);
});

test("keep-with-next groups move together to the next column", () => {
  const model = columnModel([
    paragraph("filler"),
    paragraph("keep heading", { style: { keepNext: true } }),
    paragraph("keep body"),
    paragraph("tail"),
  ]);
  const layout = core.paginateDocumentLayout(model, {
    ...PAGE_OPTIONS,
    measurementFeedback: {
      paragraphHeightsPxByNodeIndex: { 0: 400, 1: 300, 2: 300, 3: 1000 },
    },
  });

  assert.ok(layout.pages.length >= 2);
  const firstPage = layout.pages[0];
  const columnNodeIndexes = firstPage.columns.map((column) =>
    column.map((segment) => segment.nodeIndex),
  );
  assert.deepEqual(columnNodeIndexes[0], [0]);
  assert.deepEqual(columnNodeIndexes[1], [1, 2]);
});

test("measure→repaginate loop converges with bounded iterations", () => {
  const blockCount = 10;
  const actualBlockHeights = Array.from({ length: blockCount }, () => 100);
  const estimatedBlockHeights = Array.from({ length: blockCount }, () => 20);
  const pageBudgetPx = 300;
  let builds = 0;

  const buildPages = (feedback) => {
    builds += 1;
    const heights = feedback?.blockHeights ?? estimatedBlockHeights;
    const pages = [];
    let current = [];
    let consumed = 0;
    const flush = () => {
      if (current.length > 0) pages.push(current);
      current = [];
      consumed = 0;
    };
    for (let index = 0; index < blockCount; index += 1) {
      const height = heights[index] ?? estimatedBlockHeights[index];
      if (consumed + height > pageBudgetPx && current.length > 0) flush();
      current.push(index);
      consumed += height;
    }
    flush();
    return pages;
  };

  const result = core.iteratePaginationWithMeasurements({
    buildPages,
    measurePages: (pages) => {
      return {
        pageContentHeightsPx: pages.map(
          (page) =>
            page.reduce(
              (sum, blockIndex) => sum + actualBlockHeights[blockIndex],
              0
            )
        ),
        feedback: { blockHeights: actualBlockHeights },
      };
    },
    pageIdentityKeys: (pages) => pages.map((page) => page.join(",")),
    maxIterations: 4,
    tolerancePx: 1,
  });

  assert.equal(result.converged, true);
  assert.ok(result.iterations <= 4);
  assert.deepEqual(result.pages.map((page) => page.length), [3, 3, 3, 1]);
  assert.ok(builds <= 4);
});

test("measure→repaginate loop stops on oscillation", () => {
  let toggle = 0;
  const result = core.iteratePaginationWithMeasurements({
    buildPages: (feedback) => (feedback?.pages ?? [0, 1]),
    measurePages: () => {
      toggle += 1;
      return {
        pageContentHeightsPx: toggle % 2 === 0 ? [100, 100] : [200, 200],
        feedback: { pages: [0, 1] },
      };
    },
    pageIdentityKeys: () => ["a", "b"],
    maxIterations: 6,
  });

  assert.equal(result.oscillationDetected, true);
  assert.ok(result.iterations < 6);
});

test("measured paragraph heights feed back into the runtime pagination", () => {
  const nodes = Array.from({ length: 12 }, (_, index) =>
    paragraph(`paragraph ${index} ${"text ".repeat(40)}`),
  );
  const model = modelWith(nodes);
  const estimated = core.paginateDocumentLayout(model, {
    ...PAGE_OPTIONS,
    marginPx: 72,
  });
  const withMeasurements = core.paginateDocumentLayout(model, {
    ...PAGE_OPTIONS,
    marginPx: 72,
    measurementFeedback: {
      paragraphHeightsPxByNodeIndex: Object.fromEntries(
        nodes.map((_, index) => [index, 180]),
      ),
    },
  });

  assert.ok(withMeasurements.pages.length > estimated.pages.length);
});

test("stale measured page budgets are ignored after pagination changes", () => {
  const model = modelWith(Array.from({ length: 4 }, (_, index) => paragraph(`p${index}`)));
  const baseline = core.paginateDocumentLayout(model, PAGE_OPTIONS);
  const stale = core.paginateDocumentLayout(model, {
    ...PAGE_OPTIONS,
    measurementFeedback: {
      pageContentHeightsPxByPageIndex: [2000, 2000, 2000],
      pageIdentityKeys: ["stale-a", "stale-b", "stale-c"],
    },
  });

  assert.deepEqual(
    stale.pages.map((page) => page.segments.map((segment) => segment.nodeIndex)),
    baseline.pages.map((page) => page.segments.map((segment) => segment.nodeIndex)),
  );
});

test("measure→repaginate loop fits rendered paragraph heights inside the paper", () => {
  const model = modelWith(
    Array.from({ length: 20 }, (_, index) =>
      paragraph(`paragraph ${index} ${"content ".repeat(30)}`),
    ),
  );
  const options = { pageWidth: 816, pageHeight: 1056, margin: 72 };
  const pageContentHeightPx = options.pageHeight - options.margin * 2;

  const result = core.iteratePaginationWithMeasurements({
    buildPages: (feedback) =>
      core.buildLayoutSnapshot(model, {
        ...options,
        measurementFeedback: feedback,
      }),
    measurePages: (snapshot) => {
      const paragraphHeightsPxByNodeIndex = {};
      snapshot.pages.forEach((page) => {
        page.blocks.forEach((block) => {
          const nodeIndex = block.source?.nodeIndex;
          if (block.kind !== "paragraph" || nodeIndex === undefined) return;
          // Simulate a renderer whose real paragraph is 20% taller than the
          // estimator's guess.
          paragraphHeightsPxByNodeIndex[nodeIndex] = Math.round(
            block.framePx.height * 1.2,
          );
        });
      });
      return {
        pageContentHeightsPx: snapshot.pages.map((page) =>
          page.blocks.reduce(
            (bottom, block) =>
              Math.max(bottom, block.framePx.y + block.framePx.height),
            page.contentBoxPx.y,
          ) - page.contentBoxPx.y,
        ),
        feedback: { paragraphHeightsPxByNodeIndex },
      };
    },
    pageIdentityKeys: (snapshot) =>
      snapshot.pages.map((page) =>
        page.blocks
          .map(
            (block) =>
              `${block.source?.nodeIndex}:${block.source?.columnIndex ?? 0}`,
          )
          .join(","),
      ),
    maxIterations: 3,
    tolerancePx: 1,
  });

  assert.ok(result.iterations <= 3);
  assert.ok(
    result.pages.pages.every(
      (page) =>
        page.blocks.reduce(
          (bottom, block) =>
            Math.max(bottom, block.framePx.y + block.framePx.height),
          page.contentBoxPx.y,
        ) -
          page.contentBoxPx.y <=
        pageContentHeightPx + 2,
    ),
    "final pages must not overflow the paper content box",
  );
});
