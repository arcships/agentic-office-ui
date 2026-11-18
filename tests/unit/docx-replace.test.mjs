import assert from "node:assert/strict";
import { test } from "node:test";

const docxCore = await import(
  new URL("../../packages/docx-core/dist/index.js", import.meta.url).href
);

const {
  createBlankDocumentModel,
  paragraphText,
  replaceTextInDocxParagraphs,
  replaceTextInParagraphLocation,
  replaceTextRangeInParagraph,
} = docxCore;

function paragraph(children) {
  return { type: "paragraph", children };
}

function textRun(text, style) {
  return style ? { type: "text", text, style } : { type: "text", text };
}

function plainParagraph(text) {
  return paragraph([textRun(text)]);
}

function modelWith(nodes) {
  return { ...createBlankDocumentModel(), nodes };
}

/** Compare the style keys a replace may touch; clones carry extra undefined keys. */
function styleOf(child) {
  const style = child.style ?? {};
  return {
    bold: style.bold,
    italic: style.italic,
  };
}

const BOLD = { bold: true };
const ITALIC = { italic: true };

test("replaceTextRangeInParagraph splits runs and keeps the surrounding styles", () => {
  const target = paragraph([
    textRun("alpha ", BOLD),
    textRun("needle", ITALIC),
    textRun(" omega", BOLD),
  ]);

  assert.equal(replaceTextRangeInParagraph(target, 6, 12, "pin"), true);
  assert.deepEqual(
    target.children.map((child) => ({ text: child.text, style: styleOf(child) })),
    [
      { text: "alpha ", style: { bold: true, italic: undefined } },
      { text: "pin", style: { bold: undefined, italic: true } },
      { text: " omega", style: { bold: true, italic: undefined } },
    ],
    "the replacement takes the style of the matched text and the neighbours keep theirs",
  );
  assert.equal(target.sourceXml, undefined, "the edited paragraph must be re-serialized");
});

test("replaceTextRangeInParagraph handles ranges that start mid-run and cross runs", () => {
  const target = paragraph([textRun("alpha ", BOLD), textRun("needle", ITALIC)]);

  assert.equal(replaceTextRangeInParagraph(target, 3, 9, "X"), true);
  assert.deepEqual(
    target.children.map((child) => [child.text, styleOf(child)]),
    [
      ["alpX", { bold: true, italic: undefined }],
      ["dle", { bold: undefined, italic: true }],
    ],
    "equal-styled neighbours merge, so the run list does not grow",
  );

  // A deletion keeps the prefix and the suffix, and the empty result still has
  // one text run so the paragraph stays editable.
  const deleted = paragraph([textRun("needle")]);
  assert.equal(replaceTextRangeInParagraph(deleted, 0, 6, ""), true);
  assert.deepEqual(deleted.children.map((child) => child.text), [""]);
});

test("replaceTextRangeInParagraph keeps links and non-text anchors in place", () => {
  const image = { type: "image", src: "data:image/png;base64,AAAA" };
  const target = paragraph([
    textRun("go "),
    { ...textRun("here", undefined), link: "https://example.com" },
    image,
    textRun(" now"),
  ]);

  assert.equal(replaceTextRangeInParagraph(target, 3, 7, "there"), true);
  assert.deepEqual(
    target.children.map((child) => child.type === "text"
      ? { text: child.text, link: child.link }
      : { type: child.type }),
    [
      { text: "go ", link: undefined },
      { text: "there", link: "https://example.com" },
      { type: "image" },
      { text: " now", link: undefined },
    ],
    "the replaced text inherits the link and the image anchor survives",
  );
});

test("replaceTextRangeInParagraph remaps comment range offsets", () => {
  const target = paragraph([textRun("alpha needle omega")]);
  target.commentRanges = [{ commentId: 1, start: 13, end: 18, anchorText: "omega" }];

  assert.equal(replaceTextRangeInParagraph(target, 6, 12, "a much longer replacement"), true);
  assert.equal(paragraphText(target), "alpha a much longer replacement omega");
  assert.deepEqual(
    target.commentRanges?.map(({ start, end, anchorText }) => ({ start, end, anchorText })),
    [{ start: 32, end: 37, anchorText: "omega" }],
  );
});

test("replaceTextRangeInParagraph is a no-op for identical text and clamps offsets", () => {
  const target = paragraph([textRun("needle")]);
  assert.equal(replaceTextRangeInParagraph(target, 0, 6, "needle"), false);
  assert.equal(replaceTextRangeInParagraph(target, 99, 120, ""), false);

  const clamped = paragraph([textRun("needle")]);
  assert.equal(replaceTextRangeInParagraph(clamped, -5, 3, "NEE"), true);
  assert.equal(paragraphText(clamped), "NEEdle");
});

test("replaceTextInParagraphLocation addresses body, table-cell and nested-table paragraphs", () => {
  const nested = {
    type: "table",
    rows: [{ cells: [{ nodes: [plainParagraph("deep needle")] }] }],
  };
  const model = modelWith([
    plainParagraph("body needle"),
    {
      type: "table",
      rows: [{ cells: [{ nodes: [plainParagraph("cell needle"), nested] }] }],
    },
  ]);

  const body = replaceTextInParagraphLocation(model, {
    location: { kind: "paragraph", nodeIndex: 0 },
    start: 5,
    end: 11,
    replacement: "pin",
  });
  assert.equal(paragraphText(body.model.nodes[0]), "body pin");
  assert.equal(
    paragraphText(model.nodes[0]),
    "body needle",
    "the input model is untouched (copy-on-write)",
  );

  const cell = replaceTextInParagraphLocation(model, {
    location: { kind: "table-cell", tableIndex: 1, rowIndex: 0, cellIndex: 0, paragraphIndex: 0 },
    start: 5,
    end: 11,
    replacement: "pin",
  });
  assert.equal(paragraphText(cell.model.nodes[1].rows[0].cells[0].nodes[0]), "cell pin");
  assert.equal(cell.model.nodes[1].sourceXml, undefined, "the owning table is re-serialized");

  const nestedResult = replaceTextInParagraphLocation(model, {
    location: {
      kind: "nested-table-cell",
      tableIndex: 1,
      path: [{ rowIndex: 0, cellIndex: 0, contentIndex: 1 }],
      rowIndex: 0,
      cellIndex: 0,
      paragraphIndex: 0,
    },
    start: 5,
    end: 11,
    replacement: "pin",
  });
  const nestedTable = nestedResult.model.nodes[1].rows[0].cells[0].nodes[1];
  assert.equal(paragraphText(nestedTable.rows[0].cells[0].nodes[0]), "deep pin");
  assert.equal(
    paragraphText(nestedResult.model.nodes[1].rows[0].cells[0].nodes[0]),
    "cell needle",
    "a nested edit must not touch the sibling paragraph",
  );

  const missing = replaceTextInParagraphLocation(model, {
    location: { kind: "table-cell", tableIndex: 9, rowIndex: 0, cellIndex: 0, paragraphIndex: 0 },
    start: 0,
    end: 3,
    replacement: "x",
  });
  assert.equal(missing, undefined);
});

test("replaceTextInDocxParagraphs applies one snapshot's offsets per paragraph", () => {
  const model = modelWith([
    plainParagraph("needle needle"),
    plainParagraph("second needle"),
  ]);

  const result = replaceTextInDocxParagraphs(model, [
    { location: { kind: "paragraph", nodeIndex: 0 }, start: 0, end: 6, replacement: "pin" },
    { location: { kind: "paragraph", nodeIndex: 0 }, start: 7, end: 13, replacement: "pin" },
    { location: { kind: "paragraph", nodeIndex: 1 }, start: 7, end: 13, replacement: "pin" },
  ]);

  assert.equal(result.applied, 3);
  assert.equal(paragraphText(result.model.nodes[0]), "pin pin");
  assert.equal(paragraphText(result.model.nodes[1]), "second pin");
  assert.equal(
    paragraphText(model.nodes[0]),
    "needle needle",
    "the source model is not mutated",
  );
});
