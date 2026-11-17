// Pagination runtime-path regression tests (issue #25).
//
// The viewer used to paginate with `layoutDocument()` — a one-line-per-paragraph
// estimator that ignores explicit page breaks, section breaks, keep-next/widow
// control, wrapped line counts and table row splitting. The real engine
// (`buildDocumentPageNodeSegments`) existed but had no runtime caller, so
// `buildLayoutSnapshot()` was the only place the viewer could get page
// boundaries from, and it used the estimator.
//
// These tests pin the runtime path: they assert page counts, page-start node
// indexes and split ranges through `buildLayoutSnapshot()`, and they also assert
// that the estimator gives different (wrong) results for the same fixtures. If
// `buildLayoutSnapshot()` is ever reverted to the estimator, these tests fail.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

// Import the built package entry directly (the `exports["."]` target), like
// tests/unit/docx-core.test.mjs does for vue-docx. A relative import keeps this
// test bound to the checkout it lives in.
const core = await import(
  new URL("../../packages/docx-core/dist/index.js", import.meta.url).href,
);

const PAGE_OPTIONS = { pageWidth: 816, pageHeight: 1056, margin: 72 };
const PAGE_CONTENT_HEIGHT_PX =
  PAGE_OPTIONS.pageHeight - PAGE_OPTIONS.margin * 2;

function paragraph(text, options = {}) {
  return {
    type: "paragraph",
    children: [{ type: "text", text }],
    style: options.style,
    sourceXml: options.sourceXml,
  };
}

function table(cellsByRow, options = {}) {
  return {
    type: "table",
    style: options.style,
    rows: cellsByRow.map((cells, rowIndex) => ({
      id: `row-${rowIndex}`,
      style: options.rowStyle,
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
      sourceParts: 1,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
      ...metadata,
    },
  };
}

function pageStartNodeIndexes(snapshot) {
  return snapshot.pages.map(
    (page) => page.blocks[0]?.source?.nodeIndex ?? null,
  );
}

function legacyPageStartNodeIndexes(pages) {
  return pages.map((page) => {
    const match = /^paragraph-(\d+)$/.exec(page.blocks[0]?.id ?? "");
    if (match) return Number.parseInt(match[1], 10);
    const tableMatch = /^table-(\d+)$/.exec(page.blocks[0]?.id ?? "");
    return tableMatch ? Number.parseInt(tableMatch[1], 10) : null;
  });
}

function snapshotAndLegacy(model, options = PAGE_OPTIONS) {
  return {
    snapshot: core.buildLayoutSnapshot(model, options),
    legacy: core.layoutDocument(model, options),
  };
}

test("explicit <w:br w:type=page> starts a new page in the runtime snapshot", () => {
  const model = modelWith([
    paragraph("Page one body", {
      sourceXml:
        '<w:p><w:r><w:t>Page one body</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>',
    }),
    paragraph("Page two body"),
    paragraph("Page two tail"),
  ]);

  const { snapshot, legacy } = snapshotAndLegacy(model);

  assert.equal(snapshot.pages.length, 2);
  assert.deepEqual(pageStartNodeIndexes(snapshot), [0, 1]);
  assert.equal(
    snapshot.pages[1].blocks.length,
    2,
    "the paragraph after the break and its successor share the second page",
  );

  // The estimator ignores the explicit break: one page, everything stacked.
  assert.equal(legacy.length, 1);
  assert.notEqual(
    snapshot.pages.length,
    legacy.length,
    "runtime pagination must not be the estimator result",
  );
});

test("<w:pageBreakBefore> starts a new page in the runtime snapshot", () => {
  const model = modelWith([
    paragraph("Intro"),
    paragraph("Chapter start", { style: { pageBreakBefore: true } }),
    paragraph("Chapter body"),
  ]);

  const { snapshot, legacy } = snapshotAndLegacy(model);

  assert.equal(snapshot.pages.length, 2);
  assert.deepEqual(pageStartNodeIndexes(snapshot), [0, 1]);
  assert.equal(legacy.length, 1);
});

test("a next-page section break starts a new page and uses its own geometry", () => {
  const model = modelWith(
    [paragraph("Section one"), paragraph("Section two heading"), paragraph("Section two body")],
    {
      sections: [
        {
          startNodeIndex: 0,
          headerSections: [],
          footerSections: [],
          sectionPropertiesXml:
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1440" w:bottom="1080" w:left="1440"/></w:sectPr>',
        },
        {
          startNodeIndex: 1,
          headerSections: [],
          footerSections: [],
          sectionPropertiesXml:
            '<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1584" w:right="1584" w:bottom="1584" w:left="1584"/></w:sectPr>',
        },
      ],
    },
  );

  const { snapshot, legacy } = snapshotAndLegacy(model);

  assert.equal(snapshot.pages.length, 2);
  assert.deepEqual(pageStartNodeIndexes(snapshot), [0, 1]);

  // Section 1 keeps the caller geometry (Letter, 72px margins).
  assert.deepEqual(snapshot.pages[0].pageSizePx, { width: 816, height: 1056 });
  assert.equal(snapshot.pages[0].marginsPx.top, 72);

  // Section 2 uses its own <w:sectPr> geometry: A4-ish, 1584 twips = 106px.
  assert.deepEqual(snapshot.pages[1].pageSizePx, { width: 794, height: 1123 });
  assert.equal(snapshot.pages[1].marginsPx.top, 106);
  assert.equal(
    snapshot.pages[1].marginsPx.left,
    106,
    "per-section margins must reach the page that starts the section",
  );
  // The estimator now honours section breaks and per-section geometry too
  // (issue #24 item 6), so it also produces two pages here.
  assert.equal(legacy.length, 2);
});

test("wrapped paragraph lines count toward page height, not one line each", () => {
  const longText = "The quick brown fox jumps over the lazy dog. ".repeat(40);
  const model = modelWith(
    Array.from({ length: 6 }, (_, index) => paragraph(`${index} ${longText}`)),
  );

  const { snapshot, legacy } = snapshotAndLegacy(model);

  assert.equal(
    snapshot.pages.length,
    2,
    "three wrapped paragraphs fit per page; the estimator would squeeze all six onto one",
  );
  assert.deepEqual(pageStartNodeIndexes(snapshot), [0, 3]);
  assert.equal(legacy.length, 1);
});

test("a multi-page table is split into disjoint row ranges inside the page content box", () => {
  const model = modelWith([
    paragraph("Table intro"),
    table(
      Array.from({ length: 24 }, (_, rowIndex) => [`Row ${rowIndex}`]),
      // 1440 twips / 15 = 96px per row, exact.
      { rowStyle: { heightTwips: 1440, heightRule: "exact" } },
    ),
    paragraph("After the table"),
  ]);

  const { snapshot, legacy } = snapshotAndLegacy(model);

  const tableSegments = snapshot.pages.flatMap((page) =>
    page.blocks
      .filter((block) => block.kind === "table")
      .map((block) => ({
        nodeIndex: block.source?.nodeIndex,
        rowRange: block.tableRowRange,
        heightPx: block.framePx.height,
      })),
  );

  assert.ok(tableSegments.length >= 2, "the table must span more than one page");
  assert.ok(
    tableSegments.every((segment) => segment.nodeIndex === 1),
    "every table segment points at the same source node",
  );
  assert.ok(
    tableSegments.every((segment) => segment.rowRange !== undefined),
    "split table segments must carry a row range",
  );
  assert.equal(tableSegments[0].rowRange.startRowIndex, 0);
  assert.equal(
    tableSegments[tableSegments.length - 1].rowRange.endRowIndex,
    24,
    "the row ranges must cover every table row",
  );
  for (let index = 1; index < tableSegments.length; index += 1) {
    assert.equal(
      tableSegments[index].rowRange.startRowIndex,
      tableSegments[index - 1].rowRange.endRowIndex,
      "row ranges must not overlap or leave gaps",
    );
  }

  for (const page of snapshot.pages) {
    for (const block of page.blocks) {
      assert.ok(
        Math.round(block.framePx.y + block.framePx.height) <=
          Math.round(page.contentBoxPx.y + page.contentBoxPx.height),
        `block ${block.source?.nodeIndex} must stay inside the page content box`,
      );
    }
  }

  // The estimator keeps the whole table on one page and overflows the paper.
  assert.deepEqual(legacyPageStartNodeIndexes(legacy), [0, 1, 2]);
  assert.deepEqual(pageStartNodeIndexes(snapshot), [0, 1, 1]);
});

test("demo.docx paginates to the Word/LibreOffice baseline page count", async () => {
  const file = readFileSync(
    new URL("../../apps/demo/public/samples/demo.docx", import.meta.url),
  );
  const bytes = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength);
  const { model } = await core.wasmBuildDocModelFromBytes(bytes);
  const layout = core.resolveDocumentLayout(model);
  const options = {
    pageWidth: layout.pageWidthPx,
    pageHeight: layout.pageHeightPx,
    margin: layout.marginsPx.top,
  };

  const snapshot = core.buildLayoutSnapshot(model, options);
  const legacy = core.layoutDocument(model, options);

  // Issue #25: Word, LibreOffice and the upstream playground all render this
  // sample as 9 pages; the estimator path produced 6.
  assert.equal(snapshot.pages.length, 9);
  assert.equal(legacy.length, 6);
});

test("snapshot page segments match buildDocumentPageNodeSegments boundaries", () => {
  const model = modelWith([
    paragraph("First page", {
      sourceXml: '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
    }),
    table(
      Array.from({ length: 16 }, (_, rowIndex) => [`Row ${rowIndex}`]),
      { rowStyle: { heightTwips: 1440, heightRule: "exact" } },
    ),
    paragraph("Tail"),
  ]);

  const snapshot = core.buildLayoutSnapshot(model, PAGE_OPTIONS);

  // Same callback bundle the adapter uses; every helper is public API.
  const segments = core.buildDocumentPageNodeSegments(
    model,
    PAGE_CONTENT_HEIGHT_PX,
    PAGE_OPTIONS.pageWidth - PAGE_OPTIONS.margin * 2,
    {
      estimateDocNodeHeightPx: core.estimateDocNodeHeightPx,
      paragraphHasVisibleText: core.paragraphHasVisibleText,
      estimateParagraphHeightPx: core.estimateParagraphHeightPx,
      estimateParagraphLineHeightPx: core.estimateParagraphLineHeightPx,
      estimateTableRowHeightsPx: core.estimateTableRowHeightsPx,
      paragraphLineCountWithinWidth: core.paragraphLineCountWithinWidth,
      paragraphWidowControlEnabled: core.paragraphWidowControlEnabled,
      paragraphCanSplitAcrossPages: core.paragraphCanSplitAcrossPages,
      paragraphIsStructuralSectionBreakSpacer:
        core.paragraphIsStructuralSectionBreakSpacer,
    },
    undefined,
    [
      {
        startNodeIndex: 0,
        pageContentWidthPx: PAGE_OPTIONS.pageWidth - PAGE_OPTIONS.margin * 2,
        pageContentHeightPx: PAGE_CONTENT_HEIGHT_PX,
        pageContentHeightMultiplier: 1,
      },
    ],
    { allowParagraphLineSplitting: false },
  );

  assert.equal(snapshot.pages.length, segments.length);
  assert.deepEqual(
    segments.map((pageSegments) =>
      core.resolveDocumentPageSegmentStartNodeIndex(pageSegments) ?? null,
    ),
    pageStartNodeIndexes(snapshot),
  );
});
