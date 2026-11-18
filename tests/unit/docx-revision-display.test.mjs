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

const REVISION_ATTRS =
  'w:id="7" w:author="Ada" w:date="2024-01-01T00:00:00Z"';

const DELETION_XML =
  "<w:p>" +
  '<w:r><w:t xml:space="preserve">Keep </w:t></w:r>' +
  `<w:del ${REVISION_ATTRS}><w:r><w:t>gone</w:t></w:r></w:del>` +
  `<w:ins ${REVISION_ATTRS}><w:r><w:t>new</w:t></w:r></w:ins>` +
  '<w:r><w:t xml:space="preserve"> tail</w:t></w:r>' +
  "</w:p>";

function revisionParagraph() {
  return {
    type: "paragraph",
    children: [
      { type: "text", text: "Keep " },
      { type: "text", text: "gone" },
      { type: "text", text: "new" },
      { type: "text", text: " tail" },
    ],
    sourceXml: DELETION_XML,
  };
}

function revisionModel() {
  return {
    ...core.createBlankDocumentModel(),
    nodes: [revisionParagraph()],
  };
}

function paragraphText(paragraph) {
  return paragraph.children
    .filter((child) => child.type === "text")
    .map((child) => child.text)
    .join("");
}

test("reviewDisplayModel keeps deleted runs inline but removes them for balloons/none", () => {
  const model = revisionModel();

  assert.equal(
    core.reviewDisplayModel(model, "inline"),
    model,
    "inline display does not clone or rewrite the model"
  );

  for (const display of ["balloons", "none"]) {
    const review = core.reviewDisplayModel(model, display);
    assert.notEqual(review, model, `${display} returns a display model`);
    assert.equal(paragraphText(review.nodes[0]), "Keep new tail");
    assert.match(review.nodes[0].sourceXml, /<w:ins\b/u);
    assert.doesNotMatch(review.nodes[0].sourceXml, /<w:del\b/u);

    // The editing model keeps every revision for export and accept/reject.
    assert.equal(paragraphText(model.nodes[0]), "Keep gonenew tail");
    assert.match(model.nodes[0].sourceXml, /<w:del\b/u);
    assert.match(model.nodes[0].sourceXml, /<w:ins\b/u);
  }
});

test("reviewDisplayModel preserves insertion markup for balloons", () => {
  const model = revisionModel();
  const review = core.reviewDisplayModel(model, "balloons");
  const markup = core.resolveParagraphTrackedMarkup(review.nodes[0]);

  assert.equal(markup.inlineChangeByVisibleChildIndex[0], undefined);
  assert.equal(markup.inlineChangeByVisibleChildIndex[1]?.kind, "insertion");
  assert.equal(markup.inlineChangeByVisibleChildIndex[2], undefined);
});

test("deleted w:delText runs keep the paragraph and source XML untouched", () => {
  const paragraph = {
    type: "paragraph",
    children: [
      { type: "text", text: "Keep " },
      { type: "text", text: " tail" },
    ],
    sourceXml:
      "<w:p><w:r><w:t xml:space=\"preserve\">Keep </w:t></w:r>" +
      `<w:del ${REVISION_ATTRS}><w:r><w:delText>gone</w:delText></w:r></w:del>` +
      '<w:r><w:t xml:space="preserve"> tail</w:t></w:r></w:p>',
  };

  assert.equal(
    core.reviewDisplayParagraph(paragraph, "balloons"),
    paragraph,
    "w:delText never becomes a model child, so nothing is hidden"
  );
});

test("display text offsets follow the text flow balloons and none render", () => {
  const paragraph = revisionParagraph();
  const ranges = core.hiddenParagraphTextRanges(paragraph, "balloons");

  assert.deepEqual(ranges, [{ start: 5, end: 9 }]);
  assert.deepEqual(
    [...core.hiddenParagraphChildIndexes(paragraph, "balloons")],
    [1]
  );
  assert.equal(core.displayTextOffsetForSourceOffset(ranges, 5), 5);
  assert.equal(core.displayTextOffsetForSourceOffset(ranges, 9), 5);
  assert.equal(core.displayTextOffsetForSourceOffset(ranges, 14), 10);
  assert.equal(core.sourceTextOffsetForDisplayOffset(ranges, 5), 5);
  assert.equal(core.sourceTextOffsetForDisplayOffset(ranges, 10), 14);

  assert.deepEqual(core.hiddenParagraphTextRanges(paragraph, "inline"), []);
});

test("comment ranges are remapped onto the display text", () => {
  const paragraph = {
    ...revisionParagraph(),
    commentRanges: [{ commentId: 3, start: 0, end: 14 }],
  };
  const review = core.reviewDisplayParagraph(paragraph, "balloons");

  assert.deepEqual(review.commentRanges, [{ commentId: 3, start: 0, end: 10 }]);
  assert.deepEqual(paragraph.commentRanges, [{ commentId: 3, start: 0, end: 14 }]);
});

test("reviewDisplayModel strips deleted runs inside table cells and header stories", () => {
  const cellParagraph = revisionParagraph();
  const headerParagraph = revisionParagraph();
  const model = {
    ...core.createBlankDocumentModel(),
    nodes: [
      {
        type: "table",
        rows: [
          {
            cells: [{ nodes: [cellParagraph] }],
          },
        ],
      },
    ],
    metadata: {
      ...core.createBlankDocumentModel().metadata,
      headerSections: [
        { partName: "word/header1.xml", nodes: [headerParagraph] },
      ],
    },
  };

  const review = core.reviewDisplayModel(model, "none");
  assert.equal(
    paragraphText(review.nodes[0].rows[0].cells[0].nodes[0]),
    "Keep new tail"
  );
  assert.equal(
    paragraphText(review.metadata.headerSections[0].nodes[0]),
    "Keep new tail"
  );
  assert.equal(paragraphText(cellParagraph), "Keep gonenew tail");
});

test("export still writes every revision after a balloons display pass", async () => {
  const model = revisionModel();
  core.reviewDisplayModel(model, "balloons");

  const xml = await core.modelToDocumentXml(model);
  assert.match(xml, /<w:del\b/u);
  assert.match(xml, /<w:ins\b/u);
  assert.match(xml, /gone/u);
  assert.match(xml, /new/u);
});

test("paragraph mark deletions disappear from balloons/none display text", () => {
  const paragraph = {
    type: "paragraph",
    children: [{ type: "text", text: "First" }],
    paragraphMarkDeleted: true,
    sourceXml:
      "<w:p><w:pPr><w:rPr>" +
      `<w:del ${REVISION_ATTRS}/>` +
      "</w:rPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>",
  };

  assert.equal(core.reviewDisplayParagraph(paragraph, "inline"), paragraph);
  for (const display of ["balloons", "none"]) {
    const review = core.reviewDisplayParagraph(paragraph, display);
    assert.equal(review.paragraphMarkDeleted, undefined);
    assert.equal(paragraph.paragraphMarkDeleted, true);
  }
});
