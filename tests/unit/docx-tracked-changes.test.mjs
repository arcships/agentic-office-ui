import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url)
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);
const vueDocx = await import(
  new URL("../../packages/vue-docx/dist/index.js", import.meta.url).href
);
const { mount, textContent, walk } = await import(
  "../component/vue-test-renderer.mjs"
);

const REVISION_ATTRS =
  'w:id="7" w:author="Ada" w:date="2024-01-01T00:00:00Z"';

function paragraph(children, sourceXml, extra = {}) {
  return { type: "paragraph", children, sourceXml, ...extra };
}

function tableRow(text, rowProperties = "", cellProperties = "") {
  return (
    `<w:tr>${rowProperties}<w:tc>${cellProperties}` +
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc></w:tr>`
  );
}

function tableModel(rowXmls, sourceXml) {
  return {
    type: "table",
    rows: rowXmls.map((text) => ({
      type: "table-row",
      cells: [
        {
          type: "table-cell",
          nodes: [
            {
              type: "paragraph",
              children: [{ type: "text", text }],
            },
          ],
        },
      ],
    })),
    sourceXml,
  };
}

function positionedChanges(changes) {
  return changes.map((change, index) => ({
    annotation: {
      id: `change:${change.id}`,
      location: change.location,
      trackedChange: change,
      ordinal: index + 1,
    },
    anchorX: 0,
    anchorY: 20 + index * 24,
    top: 20 + index * 24,
    heightPx: 24,
    connectorLane: 0,
  }));
}

test("insertion highlight follows merged tab+text model children", () => {
  const sourceXml =
    "<w:p>" +
    `<w:ins ${REVISION_ATTRS}><w:r><w:tab/><w:t>Heading</w:t></w:r></w:ins>` +
    '<w:r><w:t xml:space="preserve"> tail</w:t></w:r>' +
    "</w:p>";
  const merged = paragraph(
    [
      { type: "text", text: "\tHeading" },
      { type: "text", text: " tail" },
    ],
    sourceXml
  );

  const markup = core.resolveParagraphTrackedMarkup(merged);
  assert.equal(markup.changes.length, 1);
  assert.equal(markup.inlineChangeByVisibleChildIndex.length, 1);
  assert.equal(
    markup.inlineChangeByVisibleChildIndex[0]?.id,
    "insertion-7"
  );

  const separate = paragraph(
    [
      { type: "text", text: "\t" },
      { type: "text", text: "Heading" },
      { type: "text", text: " tail" },
    ],
    sourceXml
  );
  // The resolver cache is keyed by sourceXml; this test intentionally feeds the
  // same XML with a different model-child shape.
  core.paragraphTrackedMarkupBySourceXml.clear();
  const separateMarkup = core.resolveParagraphTrackedMarkup(separate);
  assert.equal(separateMarkup.inlineChangeByVisibleChildIndex[0]?.kind, "insertion");
  assert.equal(separateMarkup.inlineChangeByVisibleChildIndex[1]?.kind, "insertion");
  assert.equal(separateMarkup.inlineChangeByVisibleChildIndex[2], undefined);
});

test("comment ranges use the model-child cursor, skipping form fields", () => {
  const sourceXml =
    "<w:p>" +
    "<w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent>" +
    "<w:r><w:tab/><w:t>field value</w:t></w:r>" +
    "</w:sdtContent></w:sdt>" +
    '<w:commentRangeStart w:id="5"/>' +
    "<w:r><w:t>after</w:t></w:r>" +
    '<w:commentRangeEnd w:id="5"/>' +
    "</w:p>";
  const node = paragraph(
    [
      { type: "form-field", fieldType: "text", sourceKind: "sdt", value: "field value" },
      { type: "text", text: "after" },
    ],
    sourceXml
  );

  const markup = core.resolveParagraphCommentMarkup(node);
  assert.deepEqual(markup?.commentIdsByVisibleChildIndex[0], [5]);
  assert.equal(markup?.commentIdsByVisibleChildIndex[1], undefined);
});

test("deleted paragraph mark produces a gutter entry and inline pilcrow", () => {
  const node = paragraph(
    [{ type: "text", text: "First" }],
    "<w:p><w:pPr><w:rPr>" +
      `<w:del ${REVISION_ATTRS}/>` +
      "</w:rPr></w:pPr>" +
      "<w:r><w:t>First</w:t></w:r></w:p>",
    { paragraphMarkDeleted: true }
  );

  const markup = core.resolveParagraphTrackedMarkup(node);
  assert.equal(markup.paragraphMarkChange?.kind, "deletion");
  assert.equal(markup.paragraphMarkChange?.author, "Ada");
  assert.equal(markup.paragraphMarkChange?.text, "Paragraph mark");

  const changes = core.collectTrackedChangesFromModel({ nodes: [node] });
  assert.equal(changes.length, 1);
  assert.equal(changes[0].kind, "deletion");
  assert.equal(changes[0].text, "Paragraph mark");

  const rendered = vueDocx.renderParagraphRuns(
    node,
    "paragraph-mark",
    "light",
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    { showTrackedChanges: true }
  );
  const html = vueDocx.renderStaticHtml(rendered);
  assert.match(html, /data-docx-paragraph-mark="true"/);
  assert.match(html, /¶/);
  assert.match(html, /data-docx-tracked-change="deletion"/);

  core.paragraphTrackedMarkupBySourceXml.clear();
  const inserted = paragraph(
    [{ type: "text", text: "Merged" }],
    "<w:p><w:pPr><w:rPr>" +
      `<w:ins ${REVISION_ATTRS}/>` +
      "</w:rPr></w:pPr>" +
      "<w:r><w:t>Merged</w:t></w:r></w:p>"
  );
  const insertedMarkup = core.resolveParagraphTrackedMarkup(inserted);
  assert.equal(insertedMarkup.paragraphMarkChange?.kind, "insertion");
  assert.equal(insertedMarkup.changes[0]?.text, "Paragraph mark");
});

test("deleted table row is collected once and rendered as a deleted row", async () => {
  const sourceXml =
    "<w:tbl><w:tblPr/>" +
    "<w:tblGrid><w:gridCol w:w=\"100\"/></w:tblGrid>" +
    tableRow("keep") +
    tableRow(
      "gone",
      `<w:trPr><w:del ${REVISION_ATTRS}/></w:trPr>`
    ) +
    "</w:tbl>";
  const table = tableModel(["keep", "gone"], sourceXml);

  const markup = core.resolveTableTrackedMarkup(table);
  assert.equal(markup.rowChangeByIndex.get(1)?.kind, "deletion");
  assert.equal(markup.rowChangeByIndex.get(1)?.author, "Ada");

  const changes = core.collectTrackedChangesFromModel({ nodes: [table] });
  const rowDeletions = changes.filter(
    (change) => change.kind === "deletion" && change.location.kind === "table-cell"
  );
  assert.equal(rowDeletions.length, 1);
  assert.equal(rowDeletions[0].location.rowIndex, 1);

  const mounted = await mount(vueDocx.DocxTableHost, {
    table,
    tableIndex: 0,
    showTrackedChanges: true,
  });
  const rows = walk(mounted.root).filter((node) => node.type === "tr");
  assert.equal(rows.length, 2);
  assert.match(rows[1].props.class ?? "", /docx-table-row--deletion/);
  assert.equal(rows[1].props["data-docx-tracked-change"], "deletion");
  assert.match(rows[1].props.style.backgroundColor, /rgba\(185, 28, 28/);
  mounted.app.unmount();
});

test("row-height revision is anchored at the real row with a row label", async () => {
  const rowProperties =
    "<w:trPr>" +
    '<w:trHeight w:val="300"/>' +
    `<w:trPrChange ${REVISION_ATTRS}><w:trPr><w:trHeight w:val="200"/></w:trPr></w:trPrChange>` +
    "</w:trPr>";
  const rowXmls = Array.from({ length: 6 }, (_, index) =>
    tableRow(`row ${index}`, index === 4 ? rowProperties : "")
  );
  const table = tableModel(
    Array.from({ length: 6 }, (_, index) => `row ${index}`),
    `<w:tbl><w:tblPr/>${rowXmls.join("")}</w:tbl>`
  );

  const changes = core.collectTrackedChangesFromModel({ nodes: [table] });
  assert.equal(changes.length, 1);
  assert.equal(changes[0].kind, "row-format-change");
  assert.equal(changes[0].location.kind, "table-cell");
  assert.equal(changes[0].location.rowIndex, 4);
  assert.equal(changes[0].location.cellIndex, 0);
  assert.match(changes[0].text, /Row formatting/);
  assert.match(changes[0].text, /row height/);

  const mounted = await mount(vueDocx.DocxTrackedChangeGutter, {
    positioned: positionedChanges(changes),
    zoomFactor: 1,
    pageHeightPx: 1000,
    documentTheme: "light",
  });
  assert.match(textContent(mounted.root), /Row formatting/);
  mounted.app.unmount();
});

test("cell insertions/deletions are parsed and anchored to the real cell", () => {
  const sourceXml =
    "<w:tbl><w:tblPr/>" +
    tableRow("keep") +
    tableRow(
      "inserted",
      "",
      `<w:tcPr><w:cellIns ${REVISION_ATTRS}/></w:tcPr>`
    ) +
    tableRow(
      "deleted",
      "",
      `<w:tcPr><w:cellDel ${REVISION_ATTRS}/></w:tcPr>`
    ) +
    "</w:tbl>";
  const table = tableModel(["keep", "inserted", "deleted"], sourceXml);

  const markup = core.resolveTableTrackedMarkup(table);
  assert.equal(markup.cellChangeByIndex.get("1:0")?.kind, "insertion");
  assert.equal(markup.cellChangeByIndex.get("2:0")?.kind, "deletion");

  const changes = core.collectTrackedChangesFromModel({ nodes: [table] });
  assert.deepEqual(
    changes.map((change) => [change.kind, change.location.rowIndex]),
    [
      ["insertion", 1],
      ["deletion", 2],
    ]
  );
});

test("gutter labels and accents reuse the core helpers for moves and table kinds", async () => {
  const changes = [
    {
      id: "move-from-1",
      kind: "move-from",
      author: "Ada",
      text: "moved text",
      nodeIndex: 0,
      location: { kind: "paragraph", nodeIndex: 0 },
    },
    {
      id: "move-to-1",
      kind: "move-to",
      author: "Ada",
      text: "moved text",
      nodeIndex: 0,
      location: { kind: "paragraph", nodeIndex: 0 },
    },
    {
      id: "table-format-1",
      kind: "table-format-change",
      author: "Ada",
      text: "Table formatting: width",
      nodeIndex: 0,
      location: { kind: "paragraph", nodeIndex: 0 },
    },
  ];

  const mounted = await mount(vueDocx.DocxTrackedChangeGutter, {
    positioned: positionedChanges(changes),
    zoomFactor: 1,
    pageHeightPx: 1000,
    documentTheme: "light",
  });

  const text = textContent(mounted.root);
  assert.match(text, /Moved from/);
  assert.match(text, /Moved to/);
  assert.match(text, /Table formatting/);
  assert.doesNotMatch(text, /move-from|move-to/);

  for (const kind of ["move-from", "move-to", "table-format-change"]) {
    assert.equal(core.trackedChangeKindLabel(kind), {
      "move-from": "Moved from",
      "move-to": "Moved to",
      "table-format-change": "Table formatting",
    }[kind]);
  }

  const cards = walk(mounted.root).filter((node) =>
    Boolean(node.props?.["data-docx-gutter-card-id"])
  );
  assert.equal(cards.length, 3);
  assert.match(cards[0].props.class, /docx-gutter-card--move/);
  assert.equal(cards[0].props["data-docx-gutter-balloon"], "true");
  assert.equal(cards[1].props["data-docx-gutter-balloon"], "false");
  const moveBadge = walk(cards[0]).find((node) =>
    String(node.props?.class ?? "").includes("docx-gutter-card-index")
  );
  assert.equal(moveBadge?.props.style.backgroundColor, "#70ad47");
  mounted.app.unmount();
});

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

/** Mirrors the inline revision markup of apps/demo/public/samples/review-comments.docx. */
const REVISION_PARAGRAPH_XML =
  "<w:p>" +
  '<w:r><w:t xml:space="preserve">Response time changed from </w:t></w:r>' +
  '<w:del w:id="1" w:author="Alex Reviewer" w:date="2026-07-11T00:00:00Z">' +
  "<w:r><w:delText>five business days</w:delText></w:r></w:del>" +
  '<w:r><w:t xml:space="preserve"> to </w:t></w:r>' +
  '<w:ins w:id="2" w:author="Morgan Editor" w:date="2026-07-11T00:05:00Z">' +
  "<w:r><w:t>two business days</w:t></w:r></w:ins>" +
  "<w:r><w:t>.</w:t></w:r>" +
  "</w:p>";

function revisionParagraphModel() {
  return {
    ...core.createBlankDocumentModel(),
    nodes: [
      {
        type: "paragraph",
        children: [
          { type: "text", text: "Response time changed from " },
          { type: "text", text: " to " },
          { type: "text", text: "two business days" },
          { type: "text", text: "." },
        ],
        sourceXml: REVISION_PARAGRAPH_XML,
      },
    ],
  };
}

function paragraphText(paragraph) {
  return paragraph.children.map((child) => child.text ?? "").join("");
}

function changesIn(model) {
  return core.collectTrackedChangesFromModel(model);
}

function changeIdOf(model, kind) {
  const change = changesIn(model).find((candidate) => candidate.kind === kind);
  assert.ok(change, `expected a ${kind} revision`);
  return change.id;
}

function decisionTableRowModel() {
  const cellParagraph = (text, sourceXml) => ({
    type: "paragraph",
    children: [{ type: "text", text }],
    sourceXml,
  });
  return {
    ...core.createBlankDocumentModel(),
    nodes: [
      {
        type: "table",
        rows: [0, 1, 2, 3].map((rowIndex) => ({
          cells: [
            {
              nodes: [
                cellParagraph(
                  rowIndex === 3 ? "Row three" : `Row ${rowIndex}`,
                  rowIndex === 3
                    ? "<w:p><w:ins w:id=\"9\" w:author=\"Kim Editor\" w:date=\"2026-07-12T00:00:00Z\"><w:r><w:t>Row three</w:t></w:r></w:ins></w:p>"
                    : `<w:p><w:r><w:t>Row ${rowIndex}</w:t></w:r></w:p>`
                ),
              ],
            },
          ],
        })),
        sourceXml:
          "<w:tbl><w:tblPr/><w:tr><w:tc><w:p><w:r><w:t>Row 0</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
      },
    ],
  };
}

// ---------------------------------------------------------------------------
// Accept / reject inline revisions
// ---------------------------------------------------------------------------

test("accepting a single insertion keeps the text and leaves other revisions intact", () => {
  const model = revisionParagraphModel();
  const insertionId = changeIdOf(model, "insertion");

  const result = core.acceptTrackedChangesInModel(model, { changeIds: [insertionId] });

  assert.deepEqual(result.resolvedChangeIds, [insertionId]);
  assert.equal(result.model.nodes[0].children.length, 4);
  assert.equal(
    paragraphText(result.model.nodes[0]),
    "Response time changed from  to two business days."
  );
  assert.doesNotMatch(result.model.nodes[0].sourceXml, /<w:ins\b/u);
  assert.match(result.model.nodes[0].sourceXml, /<w:del\b/u);
  assert.deepEqual(
    changesIn(result.model).map((change) => change.kind),
    ["deletion"]
  );
});

test("rejecting a single insertion removes the inserted runs and the markup", () => {
  const model = revisionParagraphModel();
  const insertionId = changeIdOf(model, "insertion");

  const result = core.rejectTrackedChangesInModel(model, { changeIds: [insertionId] });

  assert.deepEqual(result.resolvedChangeIds, [insertionId]);
  assert.equal(
    paragraphText(result.model.nodes[0]),
    "Response time changed from  to ."
  );
  assert.doesNotMatch(result.model.nodes[0].sourceXml, /<w:ins\b/u);
  assert.match(result.model.nodes[0].sourceXml, /<w:del\b/u);
});

test("rejecting a deletion restores the deleted text at its original position", () => {
  const model = revisionParagraphModel();
  const deletionId = changeIdOf(model, "deletion");

  const result = core.rejectTrackedChangesInModel(model, { changeIds: [deletionId] });

  assert.equal(
    paragraphText(result.model.nodes[0]),
    "Response time changed from five business days to two business days."
  );
  const sourceXml = result.model.nodes[0].sourceXml;
  assert.doesNotMatch(sourceXml, /<w:del\b/u);
  assert.doesNotMatch(sourceXml, /<w:delText\b/u);
  assert.match(sourceXml, /five business days/u);
  assert.match(sourceXml, /<w:t xml:space="preserve">five business days<\/w:t>/u);
});

test("accepting a deletion drops the deleted text and keeps the visible text", () => {
  const model = revisionParagraphModel();
  const deletionId = changeIdOf(model, "deletion");

  const result = core.acceptTrackedChangesInModel(model, { changeIds: [deletionId] });

  assert.equal(
    paragraphText(result.model.nodes[0]),
    "Response time changed from  to two business days."
  );
  assert.doesNotMatch(result.model.nodes[0].sourceXml, /<w:del\b/u);
});

test("accept-all clears every revision without changing the visible text", () => {
  const model = revisionParagraphModel();
  const before = paragraphText(model.nodes[0]);

  const result = core.acceptTrackedChangesInModel(model);

  assert.equal(result.resolvedChangeIds.length, 2);
  assert.equal(paragraphText(result.model.nodes[0]), before);
  assert.deepEqual(changesIn(result.model), []);
});

test("reject-all restores the pre-revision text", () => {
  const model = revisionParagraphModel();

  const result = core.rejectTrackedChangesInModel(model);

  assert.equal(
    paragraphText(result.model.nodes[0]),
    "Response time changed from five business days to ."
  );
  assert.deepEqual(changesIn(result.model), []);
});

test("a decision on an already-resolved revision is a no-op", () => {
  const model = revisionParagraphModel();
  const accepted = core.acceptTrackedChangesInModel(model).model;
  const insertionId = changeIdOf(model, "insertion");

  const result = core.acceptTrackedChangesInModel(accepted, { changeIds: [insertionId] });

  assert.equal(result.model, accepted);
  assert.deepEqual(result.resolvedChangeIds, []);
});

test("reports revisions whose source markup was already dropped by an edit", () => {
  const model = revisionParagraphModel();
  const insertionId = changeIdOf(model, "insertion");
  const edited = core.cloneDocModel(model);
  edited.nodes[0].sourceXml = undefined;

  const result = core.acceptTrackedChangesInModel(edited, { changeIds: [insertionId] });

  // Without source markup the collector no longer reports the revision at all,
  // which is exactly the data loss this API has to surface.
  assert.deepEqual(result.resolvedChangeIds, []);
  assert.equal(result.model, edited);
});

// ---------------------------------------------------------------------------
// Table row revisions
// ---------------------------------------------------------------------------

test("accepts an insertion anchored in row 3 of a table", () => {
  const model = decisionTableRowModel();
  const change = changesIn(model).find((candidate) => candidate.kind === "insertion");
  assert.ok(change);
  assert.equal(change.location.rowIndex, 3);

  const result = core.acceptTrackedChangesInModel(model, { changeIds: [change.id] });
  const table = result.model.nodes[0];

  assert.doesNotMatch(table.rows[3].cells[0].nodes[0].sourceXml, /<w:ins\b/u);
  assert.equal(table.rows[3].cells[0].nodes[0].children[0].text, "Row three");
  assert.deepEqual(changesIn(result.model), []);
});

test("restricting a decision to a table row range skips other rows", () => {
  const model = decisionTableRowModel();
  const change = changesIn(model).find((candidate) => candidate.kind === "insertion");
  assert.ok(change);

  const skipped = core.rejectTrackedChangesInModel(model, {
    nodeIndex: 0,
    tableRowRange: { startRowIndex: 0, endRowIndex: 3 },
  });
  assert.equal(skipped.model, model);

  const applied = core.rejectTrackedChangesInModel(model, {
    nodeIndex: 0,
    tableRowRange: { startRowIndex: 3, endRowIndex: 4 },
  });
  assert.deepEqual(applied.resolvedChangeIds, [change.id]);
  assert.deepEqual(applied.model.nodes[0].rows[3].cells[0].nodes[0].children, []);
});

// ---------------------------------------------------------------------------
// Formatting revisions
// ---------------------------------------------------------------------------

test("rejecting a run format change restores the previous run style", () => {
  const model = {
    ...core.createBlankDocumentModel(),
    nodes: [
      {
        type: "paragraph",
        children: [{ type: "text", text: "Formatted", style: { bold: true } }],
        sourceXml:
          "<w:p><w:r><w:rPr><w:b/>" +
          '<w:rPrChange w:id="7" w:author="Alex Reviewer"><w:rPr><w:i/></w:rPr></w:rPrChange>' +
          "</w:rPr><w:t>Formatted</w:t></w:r></w:p>",
      },
    ],
  };
  const change = changesIn(model).find((candidate) => candidate.kind === "format-change");
  assert.ok(change, "expected a run format revision");

  const rejected = core.rejectTrackedChangesInModel(model, { changeIds: [change.id] });

  assert.equal(rejected.model.nodes[0].children[0].style.bold, undefined);
  assert.equal(rejected.model.nodes[0].children[0].style.italic, true);
  assert.doesNotMatch(rejected.model.nodes[0].sourceXml, /<w:rPrChange\b/u);
  assert.deepEqual(changesIn(rejected.model), []);
});

test("accepting a run format change keeps the current formatting", () => {
  const model = {
    ...core.createBlankDocumentModel(),
    nodes: [
      {
        type: "paragraph",
        children: [{ type: "text", text: "Formatted", style: { bold: true } }],
        sourceXml:
          "<w:p><w:r><w:rPr><w:b/>" +
          '<w:rPrChange w:id="7" w:author="Alex Reviewer"><w:rPr><w:i/></w:rPr></w:rPrChange>' +
          "</w:rPr><w:t>Formatted</w:t></w:r></w:p>",
      },
    ],
  };
  const change = changesIn(model).find((candidate) => candidate.kind === "format-change");
  assert.ok(change);

  const accepted = core.acceptTrackedChangesInModel(model, { changeIds: [change.id] });

  assert.equal(accepted.model.nodes[0].children[0].style.bold, true);
  assert.equal(accepted.model.nodes[0].children[0].style.italic, undefined);
  assert.deepEqual(accepted.clearedOnlyChangeIds, []);
});

test("rejecting a paragraph format change restores the previous alignment", () => {
  const model = {
    ...core.createBlankDocumentModel(),
    nodes: [
      {
        type: "paragraph",
        children: [{ type: "text", text: "Aligned" }],
        style: { align: "right" },
        sourceXml:
          "<w:p><w:pPr><w:jc w:val=\"right\"/>" +
          '<w:pPrChange w:id="11" w:author="Alex Reviewer"><w:pPr><w:jc w:val="center"/></w:pPr></w:pPrChange>' +
          "</w:pPr><w:r><w:t>Aligned</w:t></w:r></w:p>",
      },
    ],
  };
  const change = changesIn(model).find(
    (candidate) => candidate.kind === "paragraph-format-change"
  );
  assert.ok(change, "expected a paragraph format revision");

  const rejected = core.rejectTrackedChangesInModel(model, { changeIds: [change.id] });

  assert.equal(rejected.model.nodes[0].style.align, "center");
  assert.doesNotMatch(rejected.model.nodes[0].sourceXml, /<w:pPrChange\b/u);
});

test("rejecting a table property revision clears the markup and reports it", () => {
  const model = {
    ...core.createBlankDocumentModel(),
    nodes: [
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                nodes: [
                  {
                    type: "paragraph",
                    children: [{ type: "text", text: "Cell" }],
                    sourceXml: "<w:p><w:r><w:t>Cell</w:t></w:r></w:p>",
                  },
                ],
              },
            ],
          },
        ],
        sourceXml:
          "<w:tbl><w:tblPr><w:jc w:val=\"center\"/>" +
          '<w:tblPrChange w:id="21" w:author="Alex Reviewer"><w:jc w:val="left"/></w:tblPrChange>' +
          "</w:tblPr>" +
          "<w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
      },
    ],
  };
  const change = changesIn(model).find(
    (candidate) => candidate.kind === "table-format-change"
  );
  assert.ok(change, "expected a table property revision");

  const rejected = core.rejectTrackedChangesInModel(model, { changeIds: [change.id] });

  assert.deepEqual(rejected.resolvedChangeIds, [change.id]);
  assert.deepEqual(rejected.clearedOnlyChangeIds, [change.id]);
  assert.doesNotMatch(rejected.model.nodes[0].sourceXml, /<w:tblPrChange\b/u);
});

// ---------------------------------------------------------------------------
// Gutter placement helpers
// ---------------------------------------------------------------------------
test("counts the paragraphs an edit stripped of revision markup", () => {
  const before = revisionParagraphModel();
  const after = core.cloneDocModel(before);
  after.nodes[0].sourceXml = undefined;

  assert.equal(core.countParagraphsLosingRevisionMarkup(before, after), 1);
  assert.equal(core.countParagraphsLosingRevisionMarkup(before, before), 0);

  // A decision keeps `sourceXml`; only a text edit clears it.
  const accepted = core.acceptTrackedChangesInModel(before).model;
  assert.equal(core.countParagraphsLosingRevisionMarkup(before, accepted), 0);

  const plain = core.cloneDocModel(before);
  plain.nodes[0].sourceXml = "<w:p><w:r><w:t>Plain</w:t></w:r></w:p>";
  assert.equal(core.countParagraphsLosingRevisionMarkup(plain, after), 0);
});

test("gutter policy hides insertions because they already render inline", () => {
  const model = revisionParagraphModel();
  const changes = changesIn(model);
  const insertion = changes.find((change) => change.kind === "insertion");
  const deletion = changes.find((change) => change.kind === "deletion");

  assert.equal(core.trackedChangeUsesGutterBalloon(insertion), false);
  assert.equal(core.trackedChangeUsesGutterBalloon(deletion), true);
});

test("a revision inside row 3 of a table spanning pages is owned by its own page", () => {
  const model = decisionTableRowModel();
  const change = changesIn(model).find((candidate) => candidate.kind === "insertion");
  assert.ok(change);

  const pageNodeSegmentsByPage = [
    [{ nodeIndex: 0, tableRowRange: { startRowIndex: 0, endRowIndex: 2 } }],
    [{ nodeIndex: 0, tableRowRange: { startRowIndex: 2, endRowIndex: 4 } }],
  ];

  assert.equal(
    core.trackedChangeBelongsToPageSegments(change.location, pageNodeSegmentsByPage[0]),
    false
  );
  assert.equal(
    core.trackedChangeBelongsToPageSegments(change.location, pageNodeSegmentsByPage[1]),
    true
  );

  const pageIndexById = core.resolveGutterAnnotationPageIndexes(
    changesIn(model),
    pageNodeSegmentsByPage
  );
  assert.equal(pageIndexById.get(change.id), 1);
});

test("a revision in a paragraph split across pages is owned by one page", () => {
  const model = revisionParagraphModel();
  const changes = changesIn(model);
  const pageNodeSegmentsByPage = [
    [{ nodeIndex: 0, paragraphLineRange: { startLineIndex: 0, endLineIndex: 1 } }],
    [{ nodeIndex: 0, paragraphLineRange: { startLineIndex: 1, endLineIndex: 3 } }],
  ];

  const pageIndexById = core.resolveGutterAnnotationPageIndexes(changes, pageNodeSegmentsByPage);

  for (const change of changes) {
    assert.equal(pageIndexById.get(change.id), 0);
  }
});
