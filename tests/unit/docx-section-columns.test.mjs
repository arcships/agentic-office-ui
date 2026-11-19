// Issue #98 — DOCX multi-column sections: parsing, authoring and fidelity.
//
// Every block names the bullet it covers so the mapping stays checkable:
//
//   1. w:cols parse/write ......................... readSectionColumns / writeSectionColumns
//   2. section command keeps all XML copies ....... setSectionColumns
//   3. column change survives export/reload ....... serializeDocx + buildDocModelFromBytes
//   4. fixture renders in flow order with Word's
//      page count ................................ sections-columns.docx snapshot
//
// The responsive rendering assertions (1440 / 390 px) and the toolbar control
// live in tests/component/docx-section-columns.test.mjs.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const { strFromU8, unzipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromDocxCore.resolve("@arcships/docx-core")).href
);

const SAMPLES = new URL("../../apps/demo/public/samples/", import.meta.url);
const CORPUS = new URL("../../test-data/corpus/", import.meta.url);

function sampleBytes(name) {
  const bytes = readFileSync(new URL(name, SAMPLES));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function loadSample(name) {
  return core.importDocxBuffer(sampleBytes(name), { useWorker: false });
}

function corpusBytes(name) {
  const bytes = readFileSync(new URL(name, CORPUS));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function loadCorpus(name) {
  return core.importDocxBuffer(corpusBytes(name), { useWorker: false });
}

function documentXmlOf(exported) {
  const archive = unzipSync(new Uint8Array(exported));
  return strFromU8(archive["word/document.xml"]);
}

function colsTag(xml) {
  return xml.match(/<w:cols\b[^>]*>/)?.[0] ?? "";
}

function sectionPropertiesOf(model, sectionIndex) {
  return core.resolveSectionPropertiesXml(model, sectionIndex);
}

// ── 1. w:cols parse / write ─────────────────────────────────────────

test("readSectionColumns reads count, gap and explicit column widths", () => {
  assert.deepEqual(core.readSectionColumns(undefined), {
    count: 1,
    spaceTwips: 720,
  });
  assert.deepEqual(
    core.readSectionColumns('<w:cols w:num="2" w:space="360"/>'),
    { count: 2, spaceTwips: 360 },
  );
  assert.deepEqual(
    core.readSectionColumns(
      '<w:cols w:num="2" w:space="240" w:equalWidth="0">' +
        '<w:col w:w="3000" w:space="240"/><w:col w:w="3000"/>' +
        "</w:cols>",
    ),
    { count: 2, spaceTwips: 240, widthsTwips: [3000, 3000] },
  );
  // Word's default gap applies when the attribute is omitted; a malformed
  // count falls back to one column instead of producing a zero-width body.
  assert.deepEqual(core.readSectionColumns("<w:cols/>"), {
    count: 1,
    spaceTwips: 720,
  });
  assert.equal(core.readSectionColumns('<w:cols w:num="x"/>').count, 1);
});

test("writeSectionColumns keeps every other sectPr child and drops stale geometry", () => {
  const sectionProperties =
    '<w:sectPr><w:headerReference w:type="default" r:id="rId9"/>' +
    '<w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>' +
    '<w:cols w:num="2" w:space="360" w:equalWidth="0" w:sep="1">' +
    '<w:col w:w="3000" w:space="360"/><w:col w:w="3000"/>' +
    "</w:cols>" +
    "<w:docGrid w:linePitch=\"360\"/></w:sectPr>";

  const three = core.writeSectionColumns(sectionProperties, { count: 3 });
  assert.match(
    three,
    /<w:cols w:num="3" w:space="360" w:sep="1"\/>/,
    "the new count uses equal columns and keeps the separator",
  );
  assert.doesNotMatch(three, /<w:col\b/);
  assert.doesNotMatch(three, /w:equalWidth/);
  assert.match(three, /<w:headerReference\b[^>]*r:id="rId9"/);
  assert.match(three, /<w:pgSz w:w="12240" w:h="15840"\/>/);
  assert.match(three, /<w:docGrid w:linePitch="360"\/>/);

  const one = core.writeSectionColumns(sectionProperties, { count: 1 });
  assert.match(one, /<w:cols w:space="360" w:sep="1"\/>/);
  assert.doesNotMatch(one, /w:num=/);
  assert.doesNotMatch(one, /<w:col\b/);

  // No-op writes keep the original string (and therefore no undo entry).
  const two = core.writeSectionColumns(sectionProperties, { count: 2 });
  assert.equal(two, sectionProperties);

  // A gap-only edit keeps explicit per-column widths and the start tag.
  const customWidths =
    '<w:sectPr><w:cols w:num="2" w:space="360">' +
    '<w:col w:w="3000" w:space="360"/><w:col w:w="3000"/>' +
    "</w:cols></w:sectPr>";
  const regap = core.writeSectionColumns(customWidths, { spaceTwips: 720 });
  assert.match(regap, /<w:cols w:num="2" w:space="720">/);
  assert.match(regap, /<w:col w:w="3000" w:space="360"\/>/);
  assert.deepEqual(core.readSectionColumns(regap), {
    count: 2,
    spaceTwips: 720,
    widthsTwips: [3000, 3000],
  });

  // A bare fragment is accepted the same way the other sectPr writers.
  assert.equal(
    core.writeSectionColumns('<w:cols w:num="2" w:space="720"/>', { count: 3 }),
    '<w:sectPr><w:cols w:num="3" w:space="720"/></w:sectPr>',
  );
  // Out-of-range counts clamp to OOXML's 45 columns.
  assert.match(
    core.writeSectionColumns(sectionProperties, { count: 99 }),
    /w:num="45"/,
  );
});

// ── 2. section command ──────────────────────────────────────────────

test("setSectionColumns edits only the addressed section and keeps every XML copy in step", async () => {
  const { model } = await loadSample("sections-columns.docx");
  const sectionsBefore = model.metadata.sections.map(
    (section) => section.sectionPropertiesXml,
  );
  const bodyBefore = model.metadata.sectionPropertiesXml;

  const firstTwoColumns = core.setSectionColumns(model, 0, { count: 2 });
  assert.notEqual(firstTwoColumns, model, "the command copies the model");
  assert.match(colsTag(sectionPropertiesOf(firstTwoColumns, 0)), /w:num="2"/);
  assert.equal(
    firstTwoColumns.metadata.sections[1].sectionPropertiesXml,
    sectionsBefore[1],
    "a later section keeps its page setup",
  );
  assert.equal(
    firstTwoColumns.metadata.sectionPropertiesXml,
    bodyBefore,
    "the body-level sectPr belongs to the final section",
  );
  assert.deepEqual(
    core
      .resolvePageSetup(firstTwoColumns)
      .map((section) => section.columns.count),
    [2, 1, 2],
    "resolvePageSetup reports each section's own w:cols",
  );
  // The paragraph that ends section 1 carries the same fragment for export.
  const sectionEndParagraph = firstTwoColumns.nodes
    .slice(0, 26)
    .reverse()
    .find((node) => node.type === "paragraph");
  assert.match(
    sectionEndParagraph?.sectionBreak?.sectionPropertiesXml ?? "",
    /<w:cols w:num="2"/,
  );
  // The source model was not mutated.
  assert.equal(model.metadata.sections[0].sectionPropertiesXml, sectionsBefore[0]);

  const finalOneColumn = core.setSectionColumns(firstTwoColumns, 2, { count: 1 });
  assert.doesNotMatch(
    colsTag(sectionPropertiesOf(finalOneColumn, 2)),
    /w:num=/,
  );
  assert.equal(
    finalOneColumn.metadata.sectionPropertiesXml,
    finalOneColumn.metadata.sections[2].sectionPropertiesXml,
    "the final section mirrors into the body-level sectPr",
  );
  assert.match(
    colsTag(sectionPropertiesOf(finalOneColumn, 0)),
    /w:num="2"/,
    "the earlier edit survives",
  );
});

test("setSectionColumns creates the implicit first section of a blank document", () => {
  const model = core.createBlankDocumentModel();
  const next = core.setSectionColumns(model, 0, { count: 2 });

  assert.match(
    next.metadata.sectionPropertiesXml ?? "",
    /<w:cols w:num="2" w:space="720"\/>/,
    "the body-level sectPr carries the new column count",
  );
  assert.equal(next.metadata.sections?.length, 1);
  assert.equal(
    core.readSectionColumns(next.metadata.sections[0].sectionPropertiesXml).count,
    2,
  );
  assert.equal(
    core.readSectionColumns(model.metadata.sectionPropertiesXml).count,
    1,
    "the input model is untouched",
  );
});

test("setSectionColumns rejects a no-op and preserves explicit widths when only the gap changes", async () => {
  const { model } = await loadSample("sections-columns.docx");
  assert.equal(
    core.setSectionColumns(model, 2, { count: 2 }),
    model,
    "the final section is already two columns",
  );

  const narrowGap = core.setSectionColumns(model, 2, { spaceTwips: 360 });
  assert.match(colsTag(sectionPropertiesOf(narrowGap, 2)), /w:space="360"/);
});

// ── 3. export / reload ──────────────────────────────────────────────

test("column changes survive export and reload", async () => {
  const { package: pkg, model } = await loadSample("sections-columns.docx");
  let next = core.setSectionColumns(model, 0, { count: 2 });
  next = core.setSectionColumns(next, 2, { count: 3 });

  const exported = await core.serializeDocx(next, pkg);
  const documentXml = documentXmlOf(exported);
  assert.equal((documentXml.match(/<w:sectPr\b/gi) ?? []).length, 3);
  const columnsTags = [...documentXml.matchAll(/<w:cols\b[^>]*>/gi)].map(
    (match) => match[0],
  );
  assert.equal(columnsTags.length, 3);
  assert.match(columnsTags[0], /w:num="2"/);
  assert.match(columnsTags[2], /w:num="3"/);

  const reimported = (await core.buildDocModelFromBytes(exported)).model;
  assert.deepEqual(
    reimported.metadata.sections.map(
      (section) => core.readSectionColumns(section.sectionPropertiesXml).count,
    ),
    [2, 1, 3],
  );
  assert.equal(
    core.readSectionColumns(reimported.metadata.sectionPropertiesXml).count,
    3,
  );
  // The layout loader sees the same three-column geometry Word would use.
  assert.equal(
    core.resolvePaginationSectionGeometry(reimported, {
      pageWidthPx: 816,
      pageHeightPx: 1056,
      marginPx: 72,
    })[2].columnCount,
    3,
  );
});

// ── 4. corpus fixture ───────────────────────────────────────────────

test("the two-column corpus fixture reads in column-major order on one page", async () => {
  const { model } = await loadCorpus("docx-two-columns.docx");
  assert.equal(model.metadata.sections.length, 1);
  assert.deepEqual(
    core.readSectionColumns(model.metadata.sections[0].sectionPropertiesXml),
    { count: 2, spaceTwips: 720 },
    "the fixture declares a two-column section with Word's default gap",
  );

  const snapshot = core.buildLayoutSnapshot(model, {});
  assert.equal(snapshot.pages.length, 1, "eight one-line paragraphs fit on one page");
  const page = snapshot.pages[0];
  assert.equal(page.columnCount, 2);
  assert.equal(page.columnGapPx, 48);
  assert.deepEqual(page.columnWidthsPx, [288, 288]);
  // Word balances the final section's columns: paragraphs 1-4 fill the first
  // column, 5-8 the second, and the reading order is column-major.
  assert.deepEqual(
    page.blocks.map((block) => block.source?.nodeIndex),
    [0, 1, 2, 3, 4, 5, 6, 7],
  );
  assert.deepEqual(
    page.blocks.map((block) => block.source?.columnIndex),
    [0, 0, 0, 0, 1, 1, 1, 1],
  );
  assert.ok(
    page.blocks.every((block) => block.paragraphLineRange === undefined),
    "every paragraph stays whole in its column",
  );
});

test("the section fixture keeps column count, page count and text order", async () => {
  const { model } = await loadSample("sections-columns.docx");
  const snapshot = core.buildLayoutSnapshot(model, {});

  assert.equal(snapshot.pages.length, 4, "the fixture paginates into four pages");
  assert.deepEqual(
    snapshot.pages.map((page) => page.columnCount),
    [1, 1, 1, 2],
    "only the final section is two-column",
  );

  const finalPage = snapshot.pages[3];
  assert.equal(finalPage.columnWidthsPx.length, 2);
  assert.equal(finalPage.columnWidthsPx[0], finalPage.columnWidthsPx[1]);
  // Flow order: every paragraph of the section appears once, first column
  // first, the way Word reads a two-column page.
  const flowNodeIndexes = finalPage.blocks.map(
    (block) => block.source?.nodeIndex,
  );
  assert.deepEqual(
    flowNodeIndexes,
    Array.from({ length: 12 }, (_, index) => 29 + index),
  );
  assert.deepEqual(
    [...new Set(finalPage.blocks.map((block) => block.source?.columnIndex))],
    [0, 1],
  );
  assert.ok(
    finalPage.blocks.every((block) => block.paragraphLineRange === undefined),
    "short paragraphs stay whole in their column",
  );
});

test("an explicit column break starts the next column", () => {
  const model = core.createBlankDocumentModel();
  model.metadata.sections = [
    {
      startNodeIndex: 0,
      sectionPropertiesXml:
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>' +
        '<w:cols w:num="2" w:space="720"/></w:sectPr>',
      headerSections: [],
      footerSections: [],
    },
  ];
  model.metadata.sectionPropertiesXml = model.metadata.sections[0].sectionPropertiesXml;
  model.nodes = [0, 1, 2, 3].map((index) => ({
    type: "paragraph",
    children: [{ type: "text", text: `Item ${index}` }],
    ...(index === 1
      ? {
          sourceXml:
            '<w:p><w:r><w:t>Item 1</w:t></w:r>' +
            '<w:r><w:br w:type="column"/></w:r>' +
            '<w:r><w:t>after the break</w:t></w:r></w:p>',
        }
      : {}),
  }));

  const layout = core.paginateDocumentLayout(model, {
    pageWidthPx: 816,
    pageHeightPx: 1056,
    marginPx: 72,
    measurementFeedback: {
      paragraphHeightsPxByNodeIndex: { 0: 20, 1: 20, 2: 20, 3: 20 },
    },
  });
  assert.equal(layout.pages.length, 1);
  assert.deepEqual(
    layout.pages[0].columns.map((column) =>
      column.map((segment) => segment.nodeIndex),
    ),
    [[0], [1, 2, 3]],
    "the paragraph carrying the break starts column two",
  );
});
