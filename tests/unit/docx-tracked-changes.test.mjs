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
