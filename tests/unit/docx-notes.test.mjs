// Issue #32 — DOCX footnote/endnote numbering + note content.
//
// Markers, note bodies and page placement must all come from the run model and
// from w:footnotePr / w:endnotePr, never from w:id or paragraph.sourceXml.

import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);

const WORD_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const SECTION_PROPS =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>';

function documentXml(body) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${WORD_NS}><w:body>${body}<w:sectPr>${SECTION_PROPS}</w:sectPr></w:body></w:document>`;
}

function notesXml(tag, notes) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${tag} ${WORD_NS}>${notes}</w:${tag}>`;
}

function settingsXml(body) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings ${WORD_NS}>${body}</w:settings>`;
}

function noteXml(tag, id, bodyXml) {
  return `<w:${tag} w:id="${id}">${bodyXml}</w:${tag}>`;
}

function paragraphXml(text, reference) {
  const referenceXml = reference
    ? reference.kind === "footnote"
      ? `<w:r><w:footnoteReference w:id="${reference.id}"/></w:r>`
      : `<w:r><w:endnoteReference w:id="${reference.id}"/></w:r>`
    : "";
  return `<w:p><w:r><w:t>${text}</w:t></w:r>${referenceXml}</w:p>`;
}

function sectionBreakParagraph(propsXml = SECTION_PROPS) {
  return `<w:p><w:pPr><w:sectPr>${propsXml}</w:sectPr></w:pPr><w:r><w:t>section break</w:t></w:r></w:p>`;
}

async function buildModel({ body, footnotes, endnotes, settings }) {
  let pkg = core.createMinimalDocxPackage(documentXml(body));
  if (footnotes) {
    pkg = core.withPart(pkg, {
      name: "word/footnotes.xml",
      content: notesXml("footnotes", footnotes),
    });
  }
  if (endnotes) {
    pkg = core.withPart(pkg, {
      name: "word/endnotes.xml",
      content: notesXml("endnotes", endnotes),
    });
  }
  if (settings) {
    pkg = core.withPart(pkg, {
      name: "word/settings.xml",
      content: settingsXml(settings),
    });
  }
  return core.buildDocModel(pkg);
}

function pageOf(nodeIndexes) {
  return nodeIndexes.map((nodeIndex) => ({ nodeIndex }));
}

function labels(entries) {
  return entries.map((entry) => [entry.note.id, entry.label]);
}

function allPageLabels(pages, kind) {
  return pages.flatMap((page, pageIndex) =>
    page[kind === "footnote" ? "footnotes" : "endnotes"].map((entry) => [
      pageIndex,
      entry.note.id,
      entry.label,
    ]),
  );
}

const firstNote = noteXml(
  "footnote",
  3,
  "<w:p><w:r><w:t>First note</w:t></w:r></w:p>",
);
const secondNote = noteXml(
  "footnote",
  7,
  "<w:p><w:r><w:t>Second note</w:t></w:r></w:p>",
);

test("note references are read from the run model and survive sourceXml loss", async () => {
  const model = await buildModel({
    body: paragraphXml("Alpha", { kind: "footnote", id: 3 }),
    footnotes: firstNote,
  });
  const paragraph = model.nodes[0];

  assert.ok(paragraph.sourceXml, "imported paragraph keeps its source XML");
  assert.deepEqual(core.paragraphNoteReferences(paragraph), [
    { kind: "footnote", id: 3 },
  ]);

  // Editing clears sourceXml; the note must stay attached through the model.
  const edited = { ...paragraph, sourceXml: undefined };
  assert.deepEqual(core.paragraphReferencedNoteIds(edited, "footnote"), [3]);
  assert.deepEqual(core.nodeReferencedNoteIds(edited, "footnote"), [3]);
});

test("sourceXml-only models still resolve note references (compatibility fallback)", () => {
  const paragraph = {
    type: "paragraph",
    sourceXml:
      '<w:p><w:r><w:t>Legacy</w:t></w:r><w:r><w:footnoteReference w:id="9"/></w:r></w:p>',
    children: [{ type: "text", text: "Legacy" }],
  };
  assert.deepEqual(core.paragraphReferencedNoteIds(paragraph, "footnote"), [9]);
});

test("marker and note numbering follow document order, not w:id", async () => {
  const model = await buildModel({
    body:
      paragraphXml("Alpha", { kind: "footnote", id: 3 }) +
      paragraphXml("Beta", { kind: "footnote", id: 7 }),
    footnotes: firstNote + secondNote,
  });

  const [page] = core.buildDocumentPageNotes(model, [pageOf([0, 1])]);
  assert.deepEqual(labels(page.footnotes), [
    [3, "1"],
    [7, "2"],
  ]);
  // The in-text marker uses the same map as the note body.
  assert.equal(page.footnoteLabels.get(3), "1");
  assert.equal(page.footnoteLabels.get(7), "2");
  assert.deepEqual([...core.buildDocumentNoteNumbering(model).footnote], [
    [3, 1],
    [7, 2],
  ]);
});

test("notes are attached to the page that holds their first reference", async () => {
  const model = await buildModel({
    body:
      paragraphXml("Alpha", { kind: "footnote", id: 3 }) +
      paragraphXml("Beta", { kind: "footnote", id: 7 }),
    footnotes: firstNote + secondNote,
  });

  const pages = core.buildDocumentPageNotes(model, [pageOf([0]), pageOf([1])]);
  assert.deepEqual(labels(pages[0].footnotes), [[3, "1"]]);
  assert.deepEqual(labels(pages[1].footnotes), [[7, "2"]]);
  assert.equal(pages[0].footnoteLabels.has(7), false);
  assert.equal(pages[1].footnoteLabels.get(7), "2");
});

test("editing a paragraph keeps its marker, its page note and its numbering", async () => {
  const model = await buildModel({
    body:
      paragraphXml("Alpha", { kind: "footnote", id: 3 }) +
      paragraphXml("Beta", { kind: "footnote", id: 7 }),
    footnotes: firstNote + secondNote,
  });

  for (const text of ["Alpha updated", "New Alpha", "Alph", ""]) {
    const updated = core.updateParagraphText(model, 0, text);
    assert.equal(updated.nodes[0].sourceXml, undefined, "edit clears sourceXml");
    assert.deepEqual(
      core.paragraphNoteReferences(updated.nodes[0]),
      [{ kind: "footnote", id: 3 }],
      `marker survives editing to ${JSON.stringify(text)}`,
    );

    const [page] = core.buildDocumentPageNotes(updated, [pageOf([0, 1])]);
    assert.deepEqual(labels(page.footnotes), [
      [3, "1"],
      [7, "2"],
    ]);
    assert.equal(page.footnoteLabels.get(3), "1");
  }
});

test("footnote bodies keep their tables", async () => {
  const model = await buildModel({
    body: paragraphXml("Alpha", { kind: "footnote", id: 7 }),
    footnotes: noteXml(
      "footnote",
      7,
      "<w:p><w:r><w:t>Second note</w:t></w:r></w:p>" +
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
        "<w:tr><w:tc><w:p><w:r><w:t>Note table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
    ),
  });

  const note = model.metadata.footnotes.find((entry) => entry.id === 7);
  assert.ok(note, "footnote definition is parsed");
  assert.ok(
    note.nodes.some((node) => node.type === "table"),
    "the note body keeps its table node",
  );
});

test("numRestart=eachSect restarts footnote numbering in every section", async () => {
  const body =
    paragraphXml("S1", { kind: "footnote", id: 5 }) +
    sectionBreakParagraph() +
    paragraphXml("S2", { kind: "footnote", id: 9 });
  const footnotes =
    noteXml("footnote", 5, "<w:p><w:r><w:t>First note</w:t></w:r></w:p>") +
    noteXml("footnote", 9, "<w:p><w:r><w:t>Second note</w:t></w:r></w:p>");

  const continuousModel = await buildModel({ body, footnotes });
  const continuousPages = core.buildDocumentPageNotes(continuousModel, [
    pageOf([0]),
    pageOf([1]),
    pageOf([2]),
  ]);
  assert.deepEqual(allPageLabels(continuousPages, "footnote"), [
    [0, 5, "1"],
    [2, 9, "2"],
  ]);

  const restartingModel = await buildModel({
    body,
    footnotes,
    settings: '<w:footnotePr><w:numRestart w:val="eachSect"/></w:footnotePr>',
  });
  assert.equal(restartingModel.metadata.sections.length, 2);
  const restartingPages = core.buildDocumentPageNotes(restartingModel, [
    pageOf([0]),
    pageOf([1]),
    pageOf([2]),
  ]);
  assert.deepEqual(allPageLabels(restartingPages, "footnote"), [
    [0, 5, "1"],
    [2, 9, "1"],
  ]);
});

test("numRestart=eachPage restarts numbering on every page", async () => {
  const model = await buildModel({
    body:
      paragraphXml("Alpha", { kind: "footnote", id: 3 }) +
      paragraphXml("Beta", { kind: "footnote", id: 7 }),
    footnotes: firstNote + secondNote,
    settings: '<w:footnotePr><w:numRestart w:val="eachPage"/></w:footnotePr>',
  });

  const pages = core.buildDocumentPageNotes(model, [pageOf([0]), pageOf([1])]);
  assert.deepEqual(allPageLabels(pages, "footnote"), [
    [0, 3, "1"],
    [1, 7, "1"],
  ]);
  assert.equal(pages[0].footnoteLabels.get(3), "1");
  assert.equal(pages[1].footnoteLabels.get(7), "1");
});

test("numFmt and numStart drive the marker/note label format", async () => {
  const model = await buildModel({
    body:
      paragraphXml("Alpha", { kind: "footnote", id: 3 }) +
      paragraphXml("Beta", { kind: "footnote", id: 7 }),
    footnotes: firstNote + secondNote,
    settings:
      '<w:footnotePr><w:numFmt w:val="lowerRoman"/><w:numStart w:val="3"/></w:footnotePr>',
  });

  const [page] = core.buildDocumentPageNotes(model, [pageOf([0, 1])]);
  assert.deepEqual(labels(page.footnotes), [
    [3, "iii"],
    [7, "iv"],
  ]);
  assert.equal(page.footnoteLabels.get(3), "iii");
});

test("endnotes default to the end of the document", async () => {
  const model = await buildModel({
    body:
      paragraphXml("Alpha", { kind: "endnote", id: 4 }) +
      sectionBreakParagraph() +
      paragraphXml("Beta"),
    endnotes: noteXml("endnote", 4, "<w:p><w:r><w:t>An endnote</w:t></w:r></w:p>"),
  });

  const pages = core.buildDocumentPageNotes(model, [
    pageOf([0]),
    pageOf([1]),
    pageOf([2]),
  ]);
  assert.deepEqual(allPageLabels(pages, "endnote"), [[2, 4, "i"]]);
  // The page that prints the body still gets the marker label ("i"), not 1.
  assert.equal(pages[2].endnoteLabels.get(4), "i");
});

test("endnotePr pos=sectEnd places endnotes on the last page of their section", async () => {
  const model = await buildModel({
    body:
      paragraphXml("Alpha", { kind: "endnote", id: 4 }) +
      sectionBreakParagraph() +
      paragraphXml("Beta"),
    endnotes: noteXml("endnote", 4, "<w:p><w:r><w:t>An endnote</w:t></w:r></w:p>"),
    settings: '<w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr>',
  });

  const pages = core.buildDocumentPageNotes(model, [
    pageOf([0]),
    pageOf([1]),
    pageOf([2]),
  ]);
  assert.deepEqual(allPageLabels(pages, "endnote"), [[1, 4, "i"]]);
});

test("section-level notePr overrides the document defaults", async () => {
  const sectionFootnotePr =
    "<w:footnotePr><w:numFmt w:val=\"upperLetter\"/><w:numStart w:val=\"2\"/></w:footnotePr>";
  const model = await buildModel({
    body:
      paragraphXml("S1", { kind: "footnote", id: 5 }) +
      sectionBreakParagraph(`${sectionFootnotePr}${SECTION_PROPS}`) +
      paragraphXml("S2", { kind: "footnote", id: 9 }),
    footnotes:
      noteXml("footnote", 5, "<w:p><w:r><w:t>First note</w:t></w:r></w:p>") +
      noteXml("footnote", 9, "<w:p><w:r><w:t>Second note</w:t></w:r></w:p>"),
    settings: '<w:footnotePr><w:numFmt w:val="decimal"/></w:footnotePr>',
  });

  const pages = core.buildDocumentPageNotes(model, [
    pageOf([0]),
    pageOf([1]),
    pageOf([2]),
  ]);
  // Section 1 numbers with upperLetter from 2 ("B"); section 2 falls back to the
  // document default decimal format while the counter continues ("3").
  assert.deepEqual(allPageLabels(pages, "footnote"), [
    [0, 5, "B"],
    [2, 9, "3"],
  ]);
});
