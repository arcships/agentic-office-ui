// Issue #122 — DOCX section page numbering and parity blank pages.
//
// The viewer used to derive page numbers per page (a counter that reset on
// `w:pgNumType/@w:start` plus a `w:fmt` read straight off the section) and had
// no model at all for the blank page an `oddPage`/`evenPage` section break
// inserts. This suite pins the per-section tracker that replaced it:
//
//   1. format normalisation and label formatting (`decimal` … `lowerRoman`)
//   2. numbering: restart on `w:start`, continuation without it
//   3. parity blank pages: when a break needs one, when it does not, and that
//      the engine's own blank page is the one the tracker reports
//   4. the `PAGE`/`NUMPAGES`/`SECTIONPAGES` field values per page
//   5. the two-section acceptance fixture (decimal, then `lowerRoman` restart at
//      1 with an `oddPage` start) through the real pagination engine
//
// Page counts are forced with a trailing `<w:br w:type="page"/>` rather than
// tuned heights, so these fixtures do not drift when the estimator changes.

import assert from "node:assert/strict";
import test from "node:test";

// Import the built package entry directly (the `exports["."]` target), so this
// test binds to the checkout it lives in.
const core = await import(
  new URL("../../packages/docx-core/dist/index.js", import.meta.url).href
);

const SNAPSHOT_OPTIONS = { pageWidth: 400, pageHeight: 300, margin: 50 };

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

/** A paragraph that ends its page, so the next node starts a new one. */
function paragraphEndingPage(text) {
  return {
    type: "paragraph",
    children: [{ type: "text", text }],
    sourceXml: `<w:p><w:r><w:t>${text}</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>`,
  };
}

/** `w:sectPr` for a section: `w:type`, optional `w:pgNumType`, page geometry. */
function sectionPropertiesXml({ type, format, start } = {}) {
  const typeXml = type ? `<w:type w:val="${type}"/>` : "";
  const pageNumberAttributes = [
    format === undefined ? "" : ` w:fmt="${format}"`,
    start === undefined ? "" : ` w:start="${start}"`,
  ].join("");
  const pageNumberXml =
    format === undefined && start === undefined
      ? ""
      : `<w:pgNumType${pageNumberAttributes}/>`;
  return (
    `<w:sectPr>${typeXml}${pageNumberXml}` +
    '<w:pgSz w:w="6000" w:h="9000"/>' +
    '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"' +
    ' w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>'
  );
}

function sectionModel(sectionInputs, nodes) {
  return {
    nodes,
    metadata: {
      sourceParts: 0,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
      sections: sectionInputs.map((input) => ({
        startNodeIndex: input.startNodeIndex,
        sectionPropertiesXml: sectionPropertiesXml(input),
        headerSections: [],
        footerSections: [],
      })),
    },
  };
}

/** Physical page list from the real pagination engine, with node starts. */
function pagesFromSnapshot(model) {
  const snapshot = core.buildLayoutSnapshot(model, SNAPSHOT_OPTIONS);
  return {
    snapshot,
    pages: snapshot.pages.map((page) => {
      const block = page.blocks.find(
        (candidate) =>
          typeof candidate.source?.nodeIndex === "number" &&
          candidate.source.nodeIndex >= 0
      );
      return { startNodeIndex: block?.source?.nodeIndex };
    }),
  };
}

function trackerFor(model, pages) {
  return core.createSectionPageNumberTracker({
    sections: core.resolveDocumentSectionsFromMetadata(model.metadata),
    pages,
    nodeCount: model.nodes.length,
  });
}

function labelsFor(tracker, pageCount) {
  return Array.from({ length: pageCount }, (_, index) =>
    tracker.pageLabelForPage(index)
  );
}

function blockCounts(snapshot) {
  return snapshot.pages.map((page) => page.blocks.length);
}

// ── 1. format normalisation ─────────────────────────────────────────

test("page-number formats normalise to ST_NumberFormat casing", () => {
  assert.equal(core.normalizeDocxPageNumberFormat(undefined), "decimal");
  assert.equal(core.normalizeDocxPageNumberFormat(""), "decimal");
  assert.equal(core.normalizeDocxPageNumberFormat("decimal"), "decimal");
  assert.equal(core.normalizeDocxPageNumberFormat("decimalZero"), "decimalZero");
  assert.equal(core.normalizeDocxPageNumberFormat("LowerRoman"), "lowerRoman");
  assert.equal(core.normalizeDocxPageNumberFormat("LOWERROMAN"), "lowerRoman");
  assert.equal(core.normalizeDocxPageNumberFormat("upperRoman"), "upperRoman");
  assert.equal(core.normalizeDocxPageNumberFormat("lowerLetter"), "lowerLetter");
  assert.equal(core.normalizeDocxPageNumberFormat("upperLetter"), "upperLetter");
  assert.equal(core.normalizeDocxPageNumberFormat("none"), "none");
  // A format this module does not model keeps Word's decimal default label.
  assert.equal(core.normalizeDocxPageNumberFormat("chineseCounting"), "decimal");
});

// ── 2. numbering: restart and continuation ──────────────────────────

test("a section without w:start continues the previous section's numbering", () => {
  const tracker = core.createSectionPageNumberTracker({
    sections: [
      { startNodeIndex: 0, sectionPropertiesXml: sectionPropertiesXml({ type: "nextPage" }) },
      { startNodeIndex: 2, sectionPropertiesXml: sectionPropertiesXml({ type: "nextPage" }) },
      { startNodeIndex: 4, sectionPropertiesXml: sectionPropertiesXml({ type: "nextPage" }) },
    ],
    pages: [
      { startNodeIndex: 0 },
      { startNodeIndex: 1 },
      { startNodeIndex: 2 },
      { startNodeIndex: 3 },
      { startNodeIndex: 4 },
    ],
    nodeCount: 5,
  });

  assert.deepEqual(labelsFor(tracker, 5), ["1", "2", "3", "4", "5"]);
  assert.deepEqual(
    tracker.ranges.map((range) => [range.sectionIndex, range.startNumber, range.endNumber]),
    [
      [0, 1, 2],
      [1, 3, 4],
      [2, 5, 5],
    ]
  );
  assert.deepEqual(
    tracker.ranges.map((range) => range.restarts),
    [false, false, false]
  );
});

test("w:start restarts numbering and w:fmt formats every label of the section", () => {
  const cases = [
    { format: "lowerRoman", start: 1, labels: ["i", "ii", "iii"] },
    { format: "upperRoman", start: 4, labels: ["IV", "V", "VI"] },
    { format: "lowerLetter", start: 1, labels: ["a", "b", "c"] },
    { format: "upperLetter", start: 26, labels: ["Z", "AA", "AB"] },
    { format: "decimalZero", start: 8, labels: ["08", "09", "10"] },
    { format: "none", start: 1, labels: ["", "", ""] },
    { format: undefined, start: 3, labels: ["3", "4", "5"] },
  ];

  for (const testCase of cases) {
    const tracker = core.createSectionPageNumberTracker({
      sections: [
        {
          startNodeIndex: 0,
          sectionPropertiesXml: sectionPropertiesXml({
            format: testCase.format,
            start: testCase.start,
          }),
        },
      ],
      pages: [{ startNodeIndex: 0 }, { startNodeIndex: 1 }, { startNodeIndex: 2 }],
      nodeCount: 3,
    });

    const description = `format ${String(testCase.format)} start ${testCase.start}`;
    assert.deepEqual(labelsFor(tracker, 3), testCase.labels, description);
    assert.deepEqual(
      [0, 1, 2].map((pageIndex) => tracker.pageNumberForPage(pageIndex)),
      [testCase.start, testCase.start + 1, testCase.start + 2],
      description
    );
    assert.equal(tracker.ranges[0].restarts, true, description);
    assert.equal(tracker.ranges[0].formatXml, testCase.format, description);
    assert.equal(
      tracker.ranges[0].format,
      core.normalizeDocxPageNumberFormat(testCase.format),
      description
    );
  }
});

// ── 3. parity blank pages ───────────────────────────────────────────

test("oddPage inserts a blank page when the next page number would be even", () => {
  const model = sectionModel(
    [
      { startNodeIndex: 0, type: "nextPage" },
      { startNodeIndex: 1, type: "oddPage", format: "lowerRoman", start: 1 },
    ],
    [paragraph("cover"), paragraph("appendix")]
  );
  const { snapshot, pages } = pagesFromSnapshot(model);
  assert.deepEqual(
    blockCounts(snapshot),
    [1, 0, 1],
    "the engine keeps a blank page before the odd-parity section"
  );

  const tracker = trackerFor(model, pages);
  assert.equal(tracker.pageCount, 3);
  assert.deepEqual(labelsFor(tracker, 3), ["1", "2", "i"]);
  assert.deepEqual(
    [0, 1, 2].map((pageIndex) => tracker.sectionIndexForPage(pageIndex)),
    [0, 0, 1]
  );
  assert.deepEqual(
    [0, 1, 2].map((pageIndex) => tracker.isParityBlankPage(pageIndex)),
    [false, true, false]
  );
  assert.equal(tracker.parityBlankPages.length, 1);
  const blank = tracker.parityBlankPages[0];
  assert.deepEqual(
    {
      kind: blank.kind,
      pageIndex: blank.pageIndex,
      sectionIndex: blank.sectionIndex,
      requiredParity: blank.requiredParity,
      ownerSectionIndex: blank.ownerSectionIndex,
    },
    {
      kind: "section-parity-blank-page",
      pageIndex: 1,
      sectionIndex: 1,
      requiredParity: "odd",
      ownerSectionIndex: 0,
    }
  );
  assert.equal(core.isSectionParityBlankPage(blank), true);
  // The blank page completes the range of the section that ends there.
  assert.deepEqual(tracker.ranges[0].parityBlankPageIndexes, [1]);
  assert.equal(tracker.ranges[0].endNumber, 2);
});

test("oddPage needs no blank page when the section already starts on an odd page", () => {
  const model = sectionModel(
    [
      { startNodeIndex: 0, type: "nextPage" },
      { startNodeIndex: 2, type: "oddPage", format: "lowerRoman", start: 1 },
    ],
    [paragraphEndingPage("cover"), paragraph("contents"), paragraph("appendix")]
  );
  const { snapshot, pages } = pagesFromSnapshot(model);
  assert.deepEqual(blockCounts(snapshot), [1, 1, 1]);

  const tracker = trackerFor(model, pages);
  assert.deepEqual(labelsFor(tracker, 3), ["1", "2", "i"]);
  assert.deepEqual(tracker.parityBlankPages, []);
  assert.deepEqual(
    [0, 1, 2].map((pageIndex) => tracker.isParityBlankPage(pageIndex)),
    [false, false, false]
  );
});

test("evenPage inserts a blank page when the next page number would be odd", () => {
  const model = sectionModel(
    [
      { startNodeIndex: 0, type: "nextPage" },
      { startNodeIndex: 2, type: "evenPage" },
    ],
    [paragraphEndingPage("cover"), paragraph("contents"), paragraph("appendix")]
  );
  const { snapshot, pages } = pagesFromSnapshot(model);
  assert.deepEqual(blockCounts(snapshot), [1, 1, 0, 1]);

  const tracker = trackerFor(model, pages);
  assert.deepEqual(labelsFor(tracker, 4), ["1", "2", "3", "4"]);
  assert.deepEqual(
    [0, 1, 2, 3].map((pageIndex) => tracker.isParityBlankPage(pageIndex)),
    [false, false, true, false]
  );
  assert.equal(tracker.parityBlankPages.length, 1);
  assert.equal(tracker.parityBlankPages[0].requiredParity, "even");
  assert.equal(tracker.parityBlankPages[0].ownerSectionIndex, 0);
});

test("a page with body content is never reported as a parity blank page", () => {
  // Same section sequence as the blank-page case, but every page carries
  // content: no page may be reclassified as a parity blank.
  const tracker = core.createSectionPageNumberTracker({
    sections: [
      { startNodeIndex: 0, sectionPropertiesXml: sectionPropertiesXml({ type: "nextPage" }) },
      { startNodeIndex: 2, sectionPropertiesXml: sectionPropertiesXml({ type: "oddPage" }) },
    ],
    pages: [{ startNodeIndex: 0 }, { startNodeIndex: 1 }, { startNodeIndex: 2 }],
    nodeCount: 3,
  });
  assert.deepEqual(tracker.parityBlankPages, []);
});

test("createSectionParityBlankPage validates its input", () => {
  assert.throws(
    () =>
      core.createSectionParityBlankPage({
        pageIndex: -1,
        sectionIndex: 1,
        requiredParity: "odd",
      }),
    /pageIndex must be a non-negative integer/
  );
  assert.throws(
    () =>
      core.createSectionParityBlankPage({
        pageIndex: 1,
        sectionIndex: 0.5,
        requiredParity: "odd",
      }),
    /sectionIndex must be a non-negative integer/
  );
  assert.throws(
    () =>
      core.createSectionParityBlankPage({
        pageIndex: 1,
        sectionIndex: 1,
        requiredParity: "first",
      }),
    /requiredParity must be "even" or "odd"/
  );

  const blank = core.createSectionParityBlankPage({
    pageIndex: 2,
    sectionIndex: 1,
    requiredParity: "even",
  });
  assert.equal(blank.ownerSectionIndex, 1, "owner defaults to the requiring section");
  assert.equal(core.isSectionParityBlankPage(blank), true);
  assert.equal(core.isSectionParityBlankPage({ ...blank, kind: "other" }), false);
  assert.equal(core.isSectionParityBlankPage(null), false);
  assert.equal(core.isSectionParityBlankPage("section-parity-blank-page"), false);
});

// ── 4. field values ─────────────────────────────────────────────────

test("PAGE, NUMPAGES and SECTIONPAGES resolve per section", () => {
  const model = sectionModel(
    [
      { startNodeIndex: 0, type: "nextPage" },
      { startNodeIndex: 2, type: "evenPage", format: "lowerRoman", start: 1 },
      { startNodeIndex: 3, type: "nextPage", format: "upperRoman", start: 5 },
    ],
    [
      paragraphEndingPage("cover"),
      paragraph("contents"),
      paragraph("appendix"),
      paragraph("glossary"),
    ]
  );
  const { snapshot, pages } = pagesFromSnapshot(model);
  assert.deepEqual(blockCounts(snapshot), [1, 1, 0, 1, 1]);

  const tracker = trackerFor(model, pages);
  assert.equal(tracker.pageCount, 5);
  assert.deepEqual(labelsFor(tracker, 5), ["1", "2", "3", "i", "V"]);

  const fields = Array.from({ length: 5 }, (_, pageIndex) =>
    tracker.fieldsForPage(pageIndex)
  );
  assert.deepEqual(
    fields.map((entry) => entry.pageNumber),
    [1, 2, 3, 1, 5]
  );
  assert.deepEqual(
    fields.map((entry) => entry.pageLabel),
    ["1", "2", "3", "i", "V"]
  );
  assert.deepEqual(
    fields.map((entry) => entry.pageNumberFormat),
    [undefined, undefined, undefined, "lowerRoman", "upperRoman"]
  );
  // The parity blank page belongs to the first section, so it counts there.
  assert.deepEqual(
    fields.map((entry) => entry.sectionPages),
    [3, 3, 3, 1, 1]
  );
  assert.deepEqual(
    fields.map((entry) => entry.totalPages),
    [5, 5, 5, 5, 5]
  );

  assert.equal(tracker.pageLabelForPage(99), undefined);
  assert.equal(tracker.fieldsForPage(-1), undefined);
  assert.equal(tracker.sectionIndexForPage(1.5), undefined);
});

// ── 5. acceptance fixture ───────────────────────────────────────────

test("a decimal section followed by a lowerRoman restart on an odd page labels like Word", () => {
  // Section 1: two pages, decimal numbering from 1.
  // Section 2: `oddPage` + `lowerRoman` restart at 1. The first section already
  // ends on page 2, so page 3 is free: no blank page, and the section's pages
  // read `i`, `ii`.
  const model = sectionModel(
    [
      { startNodeIndex: 0, type: "nextPage" },
      { startNodeIndex: 2, type: "oddPage", format: "lowerRoman", start: 1 },
    ],
    [
      paragraphEndingPage("cover"),
      paragraph("contents"),
      paragraphEndingPage("chapter one"),
      paragraph("chapter two"),
    ]
  );

  const { snapshot, pages } = pagesFromSnapshot(model);
  assert.deepEqual(
    blockCounts(snapshot),
    [1, 1, 1, 1],
    "no blank page: the odd-parity section already starts on page 3"
  );

  const tracker = trackerFor(model, pages);
  assert.deepEqual(labelsFor(tracker, tracker.pageCount), ["1", "2", "i", "ii"]);
  assert.deepEqual(
    tracker.ranges.map((range) => ({
      sectionIndex: range.sectionIndex,
      startPageIndex: range.startPageIndex,
      endPageIndex: range.endPageIndex,
      startNumber: range.startNumber,
      endNumber: range.endNumber,
      format: range.format,
      restarts: range.restarts,
    })),
    [
      {
        sectionIndex: 0,
        startPageIndex: 0,
        endPageIndex: 2,
        startNumber: 1,
        endNumber: 2,
        format: "decimal",
        restarts: false,
      },
      {
        sectionIndex: 1,
        startPageIndex: 2,
        endPageIndex: 4,
        startNumber: 1,
        endNumber: 2,
        format: "lowerRoman",
        restarts: true,
      },
    ]
  );
  assert.deepEqual(
    tracker.ranges.map((range) => range.nodeRange),
    [
      { startNodeIndex: 0, endNodeIndex: 2 },
      { startNodeIndex: 2, endNodeIndex: 4 },
    ]
  );
  assert.deepEqual(tracker.parityBlankPages, []);

  // The same document with a one-page first section must start the second
  // section on page 3 as well — now through a blank page 2 that continues the
  // first section's numbering, so the section's pages still read `i`, `ii`.
  const withBlank = sectionModel(
    [
      { startNodeIndex: 0, type: "nextPage" },
      { startNodeIndex: 1, type: "oddPage", format: "lowerRoman", start: 1 },
    ],
    [
      paragraph("cover"),
      paragraphEndingPage("chapter one"),
      paragraph("chapter two"),
    ]
  );
  const blankFixture = pagesFromSnapshot(withBlank);
  assert.deepEqual(blockCounts(blankFixture.snapshot), [1, 0, 1, 1]);
  const blankTracker = trackerFor(withBlank, blankFixture.pages);
  assert.deepEqual(labelsFor(blankTracker, blankTracker.pageCount), [
    "1",
    "2",
    "i",
    "ii",
  ]);
  assert.deepEqual(
    [0, 1, 2, 3].map((pageIndex) => blankTracker.isParityBlankPage(pageIndex)),
    [false, true, false, false]
  );
});

// ── 6. layout estimator parity ──────────────────────────────────────

test("layoutDocument attributes a parity blank page to the section that ends there", () => {
  const model = sectionModel(
    [
      { startNodeIndex: 0, type: "nextPage" },
      { startNodeIndex: 1, type: "oddPage", format: "lowerRoman", start: 1 },
    ],
    [paragraph("cover"), paragraph("appendix")]
  );

  const pages = core.layoutDocument(model, {
    pageWidth: SNAPSHOT_OPTIONS.pageWidth,
    pageHeight: SNAPSHOT_OPTIONS.pageHeight,
    margin: SNAPSHOT_OPTIONS.margin,
  });
  const blankIndex = pages.findIndex((page) => page.blocks.length === 0);
  assert.equal(blankIndex, 1, "one blank page is inserted before the section");
  assert.equal(
    pages[blankIndex].sectionIndex,
    0,
    "the blank page keeps the ending section's header/footer and numbering"
  );
  assert.equal(pages[blankIndex].number, 2);
  assert.equal(pages[blankIndex + 1].sectionIndex, 1);
  assert.equal(pages[blankIndex + 1].number, 3);
});
