// Bookmark navigation index (issue #148).
//
// Internal `#bookmark` links stayed dead because nothing mapped a bookmark name
// to the body node that carries it. These tests pin the mapping that the viewer
// navigation and the REF/PAGEREF field resolvers now share:
//
//   1. `w:bookmarkStart` names become model-backed targets .......... direct model
//   2. lookups accept `#name`, URL-escapes and any case ............. normalization
//   3. table-cell bookmarks keep their table index and cell position  table walk
//   4. imported documents expose the same targets ................... real DOCX import

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromDocxCore.resolve("@arcships/docx-core")).href
);
const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function bookmarkParagraph(name, text = "Target paragraph") {
  return {
    type: "paragraph",
    sourceXml:
      `<w:p><w:bookmarkStart w:id="1" w:name="${name}"/>` +
      `<w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`,
    children: [{ type: "text", text }],
  };
}

function tableWithBookmark(name) {
  return {
    type: "table",
    rows: [
      {
        type: "table-row",
        cells: [
          { type: "table-cell", nodes: [paragraph("before")] },
          {
            type: "table-cell",
            nodes: [bookmarkParagraph(name, "Cell target")],
          },
        ],
      },
    ],
  };
}

function modelWithNodes(nodes) {
  return { ...core.createBlankDocumentModel(), nodes };
}

test("normalizeDocxBookmarkName accepts href fragments, escapes and any case", () => {
  assert.equal(core.normalizeDocxBookmarkName("#Section_1"), "section_1");
  assert.equal(core.normalizeDocxBookmarkName("SECTION_1"), "section_1");
  assert.equal(core.normalizeDocxBookmarkName("  #My%20Section  "), "my section");
  // A malformed escape stays a literal name instead of breaking navigation.
  assert.equal(core.normalizeDocxBookmarkName("#Bad%E0%A4%A"), "bad%e0%a4%a");
  assert.equal(core.normalizeDocxBookmarkName(undefined), "");
});

test("createDocxBookmarkIndex maps paragraph and table bookmarks by name", () => {
  const model = modelWithNodes([
    {
      ...bookmarkParagraph("_Toc100", "Contents entry"),
      sourceXml:
        '<w:p><w:bookmarkStart w:id="1" w:name="_Toc100"/>' +
        '<w:bookmarkStart w:id="2" w:name="Section_1"/>' +
        '<w:bookmarkStart w:id="3" w:name="_GoBack"/>' +
        '<w:r><w:t>Contents entry</w:t></w:r></w:p>',
    },
    bookmarkParagraph("section_1", "Duplicate target"),
    tableWithBookmark("Table_Target"),
  ]);

  const index = core.createDocxBookmarkIndex(model);
  assert.deepEqual(
    index.targets.map((target) => target.name),
    ["_Toc100", "Section_1", "Table_Target"],
    "_GoBack is not a navigation target and a duplicate name is indexed once",
  );

  const section = index.get("#SECTION_1");
  assert.equal(section?.nodeIndex, 0, "lookup is case-insensitive and accepts the href form");
  assert.equal(section?.text, "Contents entry");
  assert.equal(index.has("section_1"), true);
  assert.equal(index.has("_GoBack"), false);
  assert.equal(index.get("missing"), undefined);

  const table = index.get("table_target");
  assert.equal(table?.nodeIndex, 2, "a cell bookmark addresses its owning table node");
  assert.deepEqual(table?.tableCell, {
    rowIndex: 0,
    cellIndex: 1,
    paragraphIndex: 0,
  });
  assert.equal(table?.text, "Cell target");
});

test("imported DOCX bookmark anchors resolve to the carrying paragraph", async () => {
  const source = readFileSync(new URL("demo.docx", samples));
  const archive = unzipSync(new Uint8Array(source));
  let documentXml = strFromU8(archive["word/document.xml"]);
  const injected =
    '<w:p><w:bookmarkStart w:id="91" w:name="Imported_Target"/>' +
    '<w:r><w:t>Imported target</w:t></w:r><w:bookmarkEnd w:id="91"/></w:p>' +
    '<w:p><w:hyperlink w:anchor="Imported_Target" w:history="1">' +
    '<w:r><w:t>Jump to imported target</w:t></w:r></w:hyperlink></w:p>';
  documentXml = documentXml.replace("<w:body>", `<w:body>${injected}`);
  archive["word/document.xml"] = strToU8(documentXml);

  const imported = await core.importDocxBuffer(zipSync(archive).buffer, {
    useWorker: false,
  });

  assert.equal(imported.model.nodes[1].type, "paragraph");
  assert.match(
    imported.model.nodes[1].children[0].link,
    /^#Imported_Target$/,
    "w:hyperlink/@w:anchor imports as an internal href",
  );
  assert.equal(imported.model.nodes[1].children[0].text, "Jump to imported target");

  const index = core.createDocxBookmarkIndex(imported.model);
  const target = index.get(imported.model.nodes[1].children[0].link);
  assert.equal(target?.nodeIndex, 0, "the bookmark lives on the preceding paragraph");
  assert.equal(target?.text, "Imported target");
});
