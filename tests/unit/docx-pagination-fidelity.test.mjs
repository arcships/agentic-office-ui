import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);
const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

function paragraph(text, options = {}) {
  return {
    type: "paragraph",
    children: [{ type: "text", text }],
    ...(options.style ? { style: options.style } : {}),
    ...(options.sourceXml ? { sourceXml: options.sourceXml } : {}),
  };
}

function docModel(nodes, metadata = {}) {
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

function segmentationCallbacks({
  paragraphHeight = 20,
  lineHeight = 20,
  lineCount = 1,
  rowHeights,
} = {}) {
  return {
    estimateDocNodeHeightPx: () => paragraphHeight,
    paragraphHasVisibleText: (value) =>
      value.children.some(
        (child) => child.type === "text" && child.text.trim().length > 0,
      ),
    paragraphIsStructuralSectionBreakSpacer: () => false,
    estimateParagraphHeightPx: () => paragraphHeight,
    estimateParagraphLineHeightPx: () => lineHeight,
    paragraphLineCountWithinWidth: () => lineCount,
    paragraphWidowControlEnabled: () => false,
    paragraphCanSplitAcrossPages: () => true,
    estimateTableRowHeightsPx: () => rowHeights ?? [],
  };
}

function tableCell(text) {
  return {
    type: "table-cell",
    nodes: [paragraph(text)],
  };
}

function tableRow(index, style) {
  return {
    type: "table-row",
    cells: [tableCell(`row ${index}`)],
    ...(style ? { style } : {}),
  };
}

test("OOXML paragraph spacing is additive instead of being collapsed", () => {
  const model = docModel([
    paragraph("one", {
      style: { spacing: { beforeTwips: 150, afterTwips: 300 } },
    }),
    paragraph("two", {
      style: { spacing: { beforeTwips: 150, afterTwips: 300 } },
    }),
    paragraph("three", {
      style: { spacing: { beforeTwips: 150, afterTwips: 300 } },
    }),
  ]);
  // 50 px each = 10 before + 20 content + 20 after.
  const pages = core.buildDocumentPageNodeSegments(
    model,
    120,
    500,
    segmentationCallbacks({ paragraphHeight: 50 }),
    undefined,
    [],
    { allowParagraphLineSplitting: false },
  );

  // Old max()-style collapsing made each paragraph 40 px, so all three fit on
  // one 120 px page. Word keeps both spacing values: only two fit.
  assert.deepEqual(
    pages.map((page) => page.map((segment) => segment.nodeIndex)),
    [[0, 1], [2]],
  );
  assert.equal(core.resolveParagraphConsumedHeightPx(50, 10, false), 50);
  assert.equal(core.resolveParagraphConsumedHeightPx(50, 10, true), 40);
  assert.equal(core.resolveParagraphConsumedHeightPx(20, 10, true), 10);
});

test("paragraph space-before is suppressed at the top of a page", () => {
  const height = 100;
  assert.equal(core.resolveParagraphConsumedHeightPx(height, 10, true), 90);
  assert.equal(core.resolveParagraphConsumedHeightPx(height, 10, false), 100);
});

test("lastRenderedPageBreak has one shared interpretation", () => {
  const internalXml =
    "<w:p><w:r><w:t>aaaaaaaa</w:t></w:r>" +
    "<w:r><w:lastRenderedPageBreak/><w:t>bbbbbbbb</w:t></w:r></w:p>";
  const leadingXml =
    "<w:p><w:r><w:lastRenderedPageBreak/><w:t>bbbbbbbb</w:t></w:r></w:p>";
  const internal = paragraph("a".repeat(8) + "b".repeat(8), {
    sourceXml: internalXml,
  });
  const leading = paragraph("b".repeat(8), { sourceXml: leadingXml });

  assert.equal(
    core.paragraphLastRenderedPageBreakPosition(internal),
    "internal",
  );
  assert.equal(core.paragraphHasLastRenderedPageBreak(internal), true);
  assert.equal(core.paragraphStartsWithLastRenderedPageBreak(internal), false);
  assert.equal(
    core.paragraphLastRenderedPageBreakPosition(leading),
    "leading",
  );
  assert.equal(core.paragraphStartsWithLastRenderedPageBreak(leading), true);

  const model = docModel([internal, leading]);
  assert.deepEqual(core.collectDocxLastRenderedPageBreakStartNodeIndexes(model), [
    1,
  ]);
  assert.deepEqual(
    core
      .buildDocumentPageNodeSegmentsFromLastRenderedPageBreakHints(
        docModel([internal]),
      )
      .map((page) => page.map((segment) => segment.nodeIndex)),
    [[0]],
  );
  assert.deepEqual(
    core
      .buildDocumentPageNodeSegmentsFromLastRenderedPageBreakHints(model)
      .map((page) => page.map((segment) => segment.nodeIndex)),
    [[0], [1]],
  );
});

test("a mid-paragraph lastRenderedPageBreak splits lines instead of the node", () => {
  const xml =
    "<w:p><w:r><w:t>aaaaaaaa</w:t></w:r>" +
    "<w:r><w:lastRenderedPageBreak/><w:t>bbbbbbbb</w:t></w:r></w:p>";
  const model = docModel([paragraph("a".repeat(8) + "b".repeat(8), { sourceXml: xml })]);

  const pages = core.buildDocumentPageNodeSegments(
    model,
    120,
    500,
    segmentationCallbacks({ paragraphHeight: 120, lineCount: 6 }),
    undefined,
    [],
    {
      allowParagraphLineSplitting: true,
      preferLastRenderedParagraphStartBreaks: true,
    },
  );

  assert.equal(pages.length, 2);
  assert.deepEqual(
    pages.map((page) => page[0].paragraphLineRange.startLineIndex),
    [0, 3],
  );
  assert.deepEqual(
    pages.map((page) => page[0].paragraphLineRange.endLineIndex),
    [3, 6],
  );
  // The paragraph itself stays on the page it started on; it is not moved
  // wholesale to the next page.
  assert.deepEqual(
    pages.map((page) => page.map((segment) => segment.nodeIndex)),
    [[0], [0]],
  );
});

test("explicit page breaks inside a paragraph are offset-level breaks", () => {
  const internalXml =
    "<w:p><w:r><w:t>aaa</w:t><w:br w:type=\"page\"/><w:t>bbb</w:t></w:r></w:p>";
  const trailingXml =
    "<w:p><w:r><w:t>aaa</w:t><w:br w:type=\"page\"/></w:r></w:p>";
  const internal = paragraph("aaabbb", { sourceXml: internalXml });
  const trailing = paragraph("aaa", { sourceXml: trailingXml });

  assert.deepEqual(core.collectParagraphExplicitPageBreakOffsets(internal), [3]);
  assert.equal(core.paragraphHasTrailingExplicitPageBreak(internal), false);
  assert.deepEqual(core.collectParagraphExplicitPageBreakOffsets(trailing), [3]);
  assert.equal(core.paragraphHasTrailingExplicitPageBreak(trailing), true);

  assert.deepEqual(
    [...core.collectTopLevelExplicitPageBreakStartNodeIndexes([
      internal,
      paragraph("next"),
    ])],
    [],
  );
  assert.deepEqual(
    [...core.collectTopLevelExplicitPageBreakStartNodeIndexes([
      trailing,
      paragraph("next"),
    ])],
    [1],
  );

  const pages = core.buildDocumentPageNodeSegments(
    docModel([internal]),
    120,
    500,
    segmentationCallbacks({ paragraphHeight: 120, lineCount: 6 }),
    undefined,
    [],
    { allowParagraphLineSplitting: false },
  );
  assert.deepEqual(
    pages.map((page) =>
      page.map((segment) => [
        segment.paragraphLineRange.startLineIndex,
        segment.paragraphLineRange.endLineIndex,
      ]),
    ),
    [[[0, 3]], [[3, 6]]],
  );

  // Even a paragraph the estimator calls a single line keeps the inline
  // break instead of silently dropping it.
  const singleLinePages = core.buildDocumentPageNodeSegments(
    docModel([internal]),
    120,
    500,
    segmentationCallbacks({ paragraphHeight: 40, lineCount: 1 }),
    undefined,
    [],
    { allowParagraphLineSplitting: false },
  );
  assert.deepEqual(
    singleLinePages.map((page) =>
      page.map((segment) => [
        segment.paragraphLineRange.startLineIndex,
        segment.paragraphLineRange.endLineIndex,
        segment.paragraphLineRange.totalLineCount,
      ]),
    ),
    [
      [[0, 1, 2]],
      [[1, 2, 2]],
    ],
  );
});

test("page-count reconciliation rejects over-tall virtual pages", () => {
  // #99 dropped the former 1.02–1.3 default candidates: a smaller page count
  // is never reached by growing the virtual page past the paper.
  const defaultUpward = core.reconcilePageCountCandidateToTargetCountByScalingHeight({
    initialPages: [1, 2, 3, 4, 5],
    targetPageCount: 4,
    buildPagesAtScale: (scale) =>
      Array.from({ length: Math.max(1, Math.ceil(5 / scale)) }, (_, index) => index),
  });
  assert.equal(defaultUpward.pageCount, 5);
  assert.equal(defaultUpward.scale, 1);

  const rejected = [];
  const options = {
    initialPages: [1, 2, 3, 4, 5],
    targetPageCount: 4,
    buildPagesAtScale: (scale) =>
      Array.from({ length: Math.max(1, Math.ceil(5 / scale)) }, (_, index) => index),
    scales: [1.3],
  };

  const guarded = core.reconcilePageCountCandidateToTargetCountByScalingHeight({
    ...options,
    onRejectedOverflowCandidate: (candidate) => rejected.push(candidate),
  });
  assert.equal(guarded.pageCount, 5);
  assert.equal(guarded.scale, 1);
  assert.deepEqual(
    rejected.map((candidate) => [candidate.pageCount, candidate.scale]),
    [[4, 1.3]],
  );

  const explicitlyClipped =
    core.reconcilePageCountCandidateToTargetCountByScalingHeight({
      ...options,
      allowPageContentOverflow: true,
    });
  assert.equal(explicitlyClipped.pageCount, 4);
  assert.ok(explicitlyClipped.scale > 1);

  const compressed = core.reconcilePageCountCandidateToTargetCountByScalingHeight({
    initialPages: [1, 2],
    targetPageCount: 3,
    buildPagesAtScale: (scale) =>
      Array.from({ length: Math.max(1, Math.ceil(2 / scale)) }, (_, index) => index),
    scales: [0.7],
  });
  assert.equal(compressed.pageCount, 3);
  assert.ok(compressed.scale <= 1);
});

test("tblHeader rows repeat on continuation pages", () => {
  const table = {
    type: "table",
    rows: [
      tableRow(0, { isHeader: true }),
      tableRow(1, { isHeader: true }),
      ...Array.from({ length: 6 }, (_, index) => tableRow(index + 2)),
    ],
  };
  const pages = core.buildDocumentPageNodeSegments(
    docModel([table]),
    100,
    500,
    segmentationCallbacks({ rowHeights: Array(8).fill(30) }),
    undefined,
    [],
    {},
  );

  const tableSegments = pages.map((page) => page[0].tableRowRange);
  assert.deepEqual(
    tableSegments.map((range) => [
      range?.startRowIndex,
      range?.endRowIndex,
      range?.repeatHeaderRowCount,
    ]),
    [
      [0, 4, undefined],
      [4, 6, 2],
      [6, 8, 2],
    ],
  );

  // Continuation pages reserve the two 30 px header rows, so only two body
  // rows fit in the 120 px content box (60 header + 2 × 30 body).
  assert.equal(pages.length, 3);
});

test("non-cantSplit table rows are sliced across pages", () => {
  const splittable = {
    type: "table",
    rows: [tableRow(0, { cantSplit: false })],
  };
  const pages = core.buildDocumentPageNodeSegments(
    docModel([splittable]),
    100,
    500,
    segmentationCallbacks({ rowHeights: [250] }),
    undefined,
    [],
    {},
  );

  assert.equal(pages.length, 3);
  assert.deepEqual(
    pages.map((page) => {
      const slice = page[0].tableRowSlice;
      return [
        slice.rowIndex,
        slice.startOffsetPx,
        slice.sliceHeightPx,
        slice.totalRowHeightPx,
      ];
    }),
    [
      [0, 0, 120, 250],
      [0, 120, 120, 250],
      [0, 240, 10, 250],
    ],
  );

  const cantSplit = {
    type: "table",
    rows: [tableRow(0, { cantSplit: true })],
  };
  const forcedPages = core.buildDocumentPageNodeSegments(
    docModel([cantSplit]),
    100,
    500,
    segmentationCallbacks({ rowHeights: [250] }),
    undefined,
    [],
    {},
  );
  assert.equal(forcedPages.length, 1);
  assert.equal(forcedPages[0][0].tableRowSlice, undefined);
  assert.equal(forcedPages[0][0].tableRowRange, undefined);
});

test("evenPage and oddPage section breaks insert parity blank pages", () => {
  const buildModel = (sectionType) =>
    docModel([paragraph("a"), paragraph("b"), paragraph("c")], {
      sections: [
        {
          startNodeIndex: 0,
          sectionPropertiesXml: "<w:sectPr/>",
          headerSections: [],
          footerSections: [],
        },
        {
          startNodeIndex: 2,
          sectionPropertiesXml: `<w:sectPr><w:type w:val="${sectionType}"/></w:sectPr>`,
          headerSections: [],
          footerSections: [],
        },
      ],
    });
  const callbacks = segmentationCallbacks({ paragraphHeight: 80 });

  const evenPages = core.buildDocumentPageNodeSegments(
    buildModel("evenPage"),
    100,
    500,
    callbacks,
    undefined,
    [],
    { allowParagraphLineSplitting: false },
  );
  assert.deepEqual(
    evenPages.map((page) => page.map((segment) => segment.nodeIndex)),
    [[0], [1], [], [2]],
  );

  const oddPages = core.buildDocumentPageNodeSegments(
    buildModel("oddPage"),
    100,
    500,
    callbacks,
    undefined,
    [],
    { allowParagraphLineSplitting: false },
  );
  assert.deepEqual(
    oddPages.map((page) => page.map((segment) => segment.nodeIndex)),
    [[0], [1], [2]],
  );

  const layoutPages = core.layoutDocument(buildModel("oddPage"), {
    pageHeight: 300,
    pageWidth: 400,
    margin: 50,
  });
  assert.deepEqual(
    layoutPages.map((page) => page.blocks.length),
    [2, 0, 1],
  );
});

test("per-section page geometry is used by layout and snapshot pages", () => {
  const model = docModel([paragraph("portrait"), paragraph("landscape")], {
    sections: [
      {
        startNodeIndex: 0,
        sectionPropertiesXml:
          '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
          '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>',
        headerSections: [],
        footerSections: [],
      },
      {
        startNodeIndex: 1,
        sectionPropertiesXml:
          '<w:sectPr><w:type w:val="nextPage"/>' +
          '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>' +
          '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="720"/></w:sectPr>',
        headerSections: [],
        footerSections: [],
      },
    ],
  });

  const pages = core.layoutDocument(model);
  assert.equal(pages.length, 2);
  assert.deepEqual(pages[0].layout, {
    pageWidth: 816,
    pageHeight: 1056,
    marginTop: 96,
    marginRight: 96,
    marginBottom: 96,
    marginLeft: 96,
  });
  assert.deepEqual(pages[1].layout, {
    pageWidth: 1056,
    pageHeight: 816,
    marginTop: 96,
    marginRight: 96,
    marginBottom: 96,
    marginLeft: 48,
  });
  assert.equal(pages[1].blocks[0].x, 48);

  const snapshot = core.buildLayoutSnapshot(model);
  assert.deepEqual(
    snapshot.pages.map((page) => page.pageSizePx),
    [
      { width: 816, height: 1056 },
      { width: 1056, height: 816 },
    ],
  );
  assert.deepEqual(
    snapshot.pages.map((page) => page.marginsPx.left),
    [96, 48],
  );
  assert.deepEqual(
    snapshot.resolvedDocument.sections.map((section) => section.layout.pageSizePx),
    [
      { width: 816, height: 1056 },
      { width: 1056, height: 816 },
    ],
  );
  assert.equal(snapshot.pages[1].blocks[0].framePx.x, 48);
});

test("font-load invalidation retires measurement caches and lets heights grow", () => {
  const generationBefore = core.textMeasurementGeneration();
  const keyBefore = core.heightEstimateCacheKeyPx(500);
  let notified = 0;
  const unsubscribe = core.onTextMeasurementInvalidated(() => {
    notified += 1;
  });
  core.invalidateTextMeasurements();
  unsubscribe();
  core.invalidateTextMeasurements();

  assert.ok(core.textMeasurementGeneration() > generationBefore);
  assert.equal(notified, 1);
  assert.notEqual(core.heightEstimateCacheKeyPx(500), keyBefore);

  // Without a generation change the stabilizer keeps its conservative "only
  // shrink" latch; a font change allows the measurement to grow again.
  assert.deepEqual(
    core.stabilizeMeasuredPageContentHeights([100], [140], {
      currentGeneration: 1,
      nextGeneration: 1,
    }),
    [100],
  );
  assert.deepEqual(
    core.stabilizeMeasuredPageContentHeights([100], [140], {
      currentGeneration: 1,
      nextGeneration: 2,
    }),
    [140],
  );
});

test("wasm parses w:tblHeader and w:cantSplit into table row styles", async () => {
  const bytes = readFileSync(new URL("invoice-table.docx", samples));
  const archive = unzipSync(new Uint8Array(bytes));
  const documentXml = strFromU8(archive["word/document.xml"]).replace(
    "<w:tr>",
    "<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>",
  );
  archive["word/document.xml"] = strToU8(documentXml);
  const zipped = zipSync(archive);
  const arrayBuffer = zipped.buffer.slice(
    zipped.byteOffset,
    zipped.byteOffset + zipped.byteLength,
  );

  const { model } = await core.wasmBuildDocModelFromBytes(arrayBuffer);
  const table = model.nodes.find((node) => node.type === "table");
  assert.ok(table);
  assert.deepEqual(table.rows[0].style, {
    cantSplit: true,
    isHeader: true,
  });
  assert.equal(table.rows[1].style?.isHeader, undefined);
});

test("document.fonts loading events invalidate text measurements", () => {
  const listeners = new Map();
  globalThis.document = {
    fonts: {
      addEventListener(type, listener) {
        listeners.set(type, listener);
      },
      ready: Promise.resolve({}),
    },
  };
  try {
    assert.equal(core.ensureDocumentFontInvalidation(), true);
    assert.equal(typeof listeners.get("loadingdone"), "function");
    const before = core.textMeasurementGeneration();
    listeners.get("loadingdone")();
    assert.ok(core.textMeasurementGeneration() > before);
  } finally {
    delete globalThis.document;
  }
});
