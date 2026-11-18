// Issue #32 — rendered footnote/endnote markers and note bodies.
//
// The marker number in the text and the number next to the note body must come
// from the same numbering source, and note bodies must keep tables.

import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const core = await importFromDemo("@arcships/docx-core");
const { DocxViewer, useDocxEditor } = await importFromDemo("@arcships/vue-docx");

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

async function buildModel({ body, footnotes, endnotes }) {
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
  return core.buildDocModel(pkg);
}

function markerLabels(root) {
  const labels = [];
  for (const node of walk(root)) {
    const html = String(node.props?.innerHTML ?? "");
    for (const match of html.matchAll(
      /data-docx-note-marker="true"[^>]*>([^<]*)</g,
    )) {
      labels.push(match[1]);
    }
  }
  return labels;
}

function htmlToText(html) {
  return String(html ?? "")
    .replace(/<[^>]*>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&amp;/gi, "&")
    .replace(/\s+/g, " ")
    .trim();
}

/** Text of a subtree: rendered `v-html` hosts plus plain text nodes. */
function domText(root) {
  const parts = [];
  for (const node of walk(root)) {
    if (node.props?.innerHTML) parts.push(htmlToText(node.props.innerHTML));
    if (node.type === "#text" || node.type === "#static") parts.push(node.text);
  }
  return parts.join(" ");
}

function noteArticles(section) {
  if (!section) return [];
  return walk(section).filter((node) => node.props?.["data-note-id"] !== undefined);
}

function noteLabel(article) {
  const marker = walk(article).find(
    (node) => node.props?.["data-docx-note-body-marker"] === "true",
  );
  return marker ? textContent(marker) : undefined;
}

async function render(model) {
  const mounted = await mount(DocxViewer, { model, showToolbar: false });
  return mounted;
}

const footnoteBodies =
  noteXml("footnote", 3, "<w:p><w:r><w:t>First note</w:t></w:r></w:p>") +
  noteXml(
    "footnote",
    7,
    "<w:p><w:r><w:t>Second note</w:t></w:r></w:p>" +
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      "<w:tr><w:tc><w:p><w:r><w:t>Note table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
  );

test("DOCX markers and footnote bodies share one numbering (ids 3 and 7 render 1 and 2)", async () => {
  const model = await buildModel({
    body:
      paragraphXml("Alpha", { kind: "footnote", id: 3 }) +
      paragraphXml("Beta", { kind: "footnote", id: 7 }),
    footnotes: footnoteBodies,
  });

  const mounted = await render(model);
  try {
    assert.deepEqual(markerLabels(mounted.root), ["1", "2"]);

    const section = findByTestId(mounted.root, "docx-footnotes");
    const articles = noteArticles(section);
    assert.deepEqual(
      articles.map((article) => article.props["data-note-id"]),
      [3, 7],
    );
    assert.deepEqual(articles.map(noteLabel), ["1", "2"]);
    assert.match(domText(section), /First note/);
    assert.match(domText(section), /Second note/);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("DOCX footnote bodies render tables", async () => {
  const model = await buildModel({
    body: paragraphXml("Alpha", { kind: "footnote", id: 7 }),
    footnotes: footnoteBodies,
  });

  const mounted = await render(model);
  try {
    const section = findByTestId(mounted.root, "docx-footnotes");
    const [article] = noteArticles(section);
    assert.ok(article, "the referenced footnote is rendered");
    assert.equal(article.props["data-note-id"], 7);
    assert.match(domText(article), /Note table cell/);
    assert.ok(
      walk(article).some((node) => node.type === "table"),
      "the table inside the footnote body is rendered",
    );
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("DOCX endnote markers and bodies share the endnote numbering", async () => {
  const model = await buildModel({
    body:
      paragraphXml("Alpha", { kind: "endnote", id: 4 }) +
      paragraphXml("Beta", { kind: "endnote", id: 8 }),
    endnotes:
      noteXml("endnote", 4, "<w:p><w:r><w:t>First endnote</w:t></w:r></w:p>") +
      noteXml("endnote", 8, "<w:p><w:r><w:t>Second endnote</w:t></w:r></w:p>"),
  });

  const mounted = await render(model);
  try {
    assert.deepEqual(markerLabels(mounted.root), ["i", "ii"]);

    const section = findByTestId(mounted.root, "docx-endnotes");
    const articles = noteArticles(section);
    assert.deepEqual(
      articles.map((article) => article.props["data-note-id"]),
      [4, 8],
    );
    assert.deepEqual(articles.map(noteLabel), ["i", "ii"]);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("DOCX editing a paragraph with a footnote keeps marker and note body", async () => {
  const model = await buildModel({
    body:
      paragraphXml("Alpha", { kind: "footnote", id: 3 }) +
      paragraphXml("Beta", { kind: "footnote", id: 7 }),
    footnotes: footnoteBodies,
  });

  const edited = core.updateParagraphText(model, 0, "Alpha edited");
  assert.equal(edited.nodes[0].sourceXml, undefined);

  const mounted = await render(edited);
  try {
    assert.deepEqual(markerLabels(mounted.root), ["1", "2"]);
    const section = findByTestId(mounted.root, "docx-footnotes");
    assert.deepEqual(
      noteArticles(section).map(noteLabel),
      ["1", "2"],
    );
    assert.match(domText(section), /First note/);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

// ── Remaining issue #32 items ──────────────────────────────────────

function tableCellParagraphXml(text, reference) {
  return (
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr><w:tc><w:p><w:r><w:t>${text}</w:t></w:r>` +
    `<w:r><w:${reference.kind}Reference w:id="${reference.id}"/></w:r></w:p></w:tc></w:tr>` +
    "</w:tbl>"
  );
}

test("DOCX a footnote inside a table cell renders its marker and body", async () => {
  const model = await buildModel({
    body: tableCellParagraphXml("Cell text", { kind: "footnote", id: 3 }),
    footnotes: noteXml(
      "footnote",
      3,
      "<w:p><w:r><w:t>Cell note</w:t></w:r></w:p>" +
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
        "<w:tr><w:tc><w:p><w:r><w:t>Cell note table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
    ),
  });

  const mounted = await render(model);
  try {
    assert.deepEqual(markerLabels(mounted.root), ["1"]);
    const section = findByTestId(mounted.root, "docx-footnotes");
    const [article] = noteArticles(section);
    assert.equal(noteLabel(article), "1");
    assert.match(domText(article), /Cell note/);
    assert.ok(
      walk(article).some((node) => node.type === "table"),
      "the footnote table renders next to the cell marker",
    );
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

function nestedCellTableModel() {
  const cellParagraph = {
    type: "paragraph",
    children: [
      { type: "text", text: "Nested text" },
      { type: "text", text: "", noteReference: { kind: "footnote", id: 3 } },
    ],
  };
  const nestedTable = {
    type: "table",
    rows: [
      { type: "table-row", cells: [{ type: "table-cell", nodes: [cellParagraph] }] },
    ],
  };
  const hostTable = {
    type: "table",
    rows: [{ type: "table-row", cells: [{ type: "table-cell", nodes: [nestedTable] }] }],
  };
  const base = core.createBlankDocumentModel();
  return {
    ...base,
    nodes: [hostTable],
    metadata: {
      ...base.metadata,
      footnotes: [
        {
          id: 3,
          text: "Nested cell note",
          nodes: [
            {
              type: "paragraph",
              children: [{ type: "text", text: "Nested cell note" }],
            },
          ],
        },
      ],
    },
  };
}

function nestedCellParagraph(model) {
  return model.nodes[0].rows[0].cells[0].nodes[0].rows[0].cells[0].nodes[0];
}

async function mountController(model) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () => vue.h("div", { "data-testid": "notes-controller" }, controller.status);
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

test("DOCX editing a nested table cell keeps its footnote reference", async () => {
  const model = nestedCellTableModel();
  const mounted = await mountController(model);
  try {
    mounted.controller.commitTableCellTextByRuntimeKey(
      "body-cell-nested-table-0-0-0-0",
      0,
      0,
      "Nested edited",
    );
    await vue.nextTick();

    const paragraph = nestedCellParagraph(mounted.controller.model);
    const text = paragraph.children
      .filter((child) => child.type === "text")
      .map((child) => child.text)
      .join("");
    assert.equal(text, "Nested edited");
    assert.deepEqual(core.paragraphNoteReferences(paragraph), [
      { kind: "footnote", id: 3 },
    ]);

    const pageNotes = core.buildDocumentPageNotes(mounted.controller.model, [
      [{ nodeIndex: 0 }],
    ]);
    assert.deepEqual(
      pageNotes[0].footnotes.map((entry) => [entry.note.id, entry.label]),
      [[3, "1"]],
    );
    assert.equal(pageNotes[0].footnoteLabels.get(3), "1");
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});
