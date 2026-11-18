// OpenDocument text import — issue #144.
//
// Upstream react-docx opens the ODF text family in view-only mode. These tests
// exercise the local reader with the paired packaged/flat fixtures: the model
// must carry headings, inline formatting, lists, a spanning table, an inline
// picture and page setup, must report the model as OpenDocument, and must keep
// the DOCX entry points (parseDocx / export) from pretending ODF is Word.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const { strToU8, zipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);

const FIXTURES = new URL("./../fixtures/docx/opendocument/", import.meta.url);

function fixtureBytes(name) {
  const bytes = readFileSync(new URL(name, FIXTURES));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function paragraphs(nodes) {
  return nodes.flatMap((node) =>
    node.type === "paragraph"
      ? [node]
      : node.rows.flatMap((row) => row.cells.flatMap((cell) => paragraphs(cell.nodes)))
  );
}

function paragraphText(paragraph) {
  return paragraph.children
    .map((child) => (child.type === "text" ? child.text : ""))
    .join("");
}

function images(nodes) {
  return paragraphs(nodes).flatMap((paragraph) =>
    paragraph.children.filter((child) => child.type === "image")
  );
}

test("a packaged .odt builds a model without a Word package", async () => {
  const { package: pkg, model } = await core.buildDocModelFromBytes(
    fixtureBytes("sample-html-features.odt")
  );

  assert.equal(pkg.parts.size, 0, "OpenDocument text has no OOXML parts");
  assert.equal(pkg.binaryAssets.size, 0);
  assert.equal(core.isOpenDocumentModel(model), true);
  assert.deepEqual(model.metadata.sourcePartNames, { document: "content.xml" });
  assert.equal(model.metadata.documentOpenTag, undefined);
  assert.match(model.metadata.sectionPropertiesXml, /<w:pgSz w:w="11906" w:h="16838"/);
  assert.match(model.metadata.sectionPropertiesXml, /w:pgMar w:top="567"/);

  const body = paragraphs(model.nodes);
  assert.ok(body.every((paragraph) => paragraph.sourceXml === undefined));
  assert.equal(body.find((paragraph) => paragraphText(paragraph) === "First heading")?.style?.headingLevel, 1);
  assert.equal(body.find((paragraph) => paragraphText(paragraph) === "Second level")?.style?.headingLevel, 2);
  assert.equal(body.find((paragraph) => paragraphText(paragraph) === "Last paragraph.")?.style?.align, "left");

  const plain = body.find((paragraph) => paragraphText(paragraph).startsWith("Plain "));
  assert.ok(plain);
  const bold = plain.children.find((child) => child.type === "text" && child.text === "bold");
  assert.equal(bold?.style?.bold, true);
  const italic = plain.children.find((child) => child.type === "text" && child.text === "italic");
  assert.equal(italic?.style?.italic, true);
  const underlined = plain.children.find((child) => child.type === "text" && child.text === "under");
  assert.equal(underlined?.style?.underline, true);
  const link = plain.children.find((child) => child.type === "text" && child.text === "link");
  assert.equal(link?.link, "https://example.com/x");
  assert.equal(link?.style?.color, "#000080", "the ODF Internet link style keeps its colour");

  const bullet = body.find((paragraph) => paragraphText(paragraph) === "Bullet one");
  assert.equal(bullet?.style?.numbering?.ilvl, 0);
  const nested = body.find((paragraph) => paragraphText(paragraph) === "Nested bullet");
  assert.equal(nested?.style?.numbering?.ilvl, 1);
  const numbered = body.find((paragraph) => paragraphText(paragraph) === "Number one");
  assert.ok(numbered?.style?.numbering);
  assert.notEqual(numbered?.style?.numbering?.numId, bullet?.style?.numbering?.numId);
  assert.ok(model.metadata.numberingDefinitions?.abstracts.length >= 2);

  const table = model.nodes.find((node) => node.type === "table");
  assert.ok(table);
  assert.deepEqual(table.style?.columnWidthsTwips, [900, 872, 943]);
  assert.deepEqual(
    table.rows.map((row) => row.cells.length),
    [3, 2, 3],
    "a horizontal span keeps its logical width without a covered cell",
  );
  assert.equal(table.rows[1].cells[0].style?.gridSpan, 2);
  assert.equal(table.rows[1].cells[1].style?.rowSpan, 2);
  assert.equal(table.rows[2].cells[2].style?.vMergeContinuation, true);
  assert.deepEqual(
    table.rows.flatMap((row) => row.cells.flatMap((cell) => paragraphs(cell.nodes))).map(paragraphText),
    ["Head A", "Head B", "Head C", "Wide cell", "Tall cell", "Left", "Middle", ""],
  );

  const [image] = images(model.nodes);
  assert.equal(image.contentType, "image/png");
  assert.ok(image.data instanceof Uint8Array);
  assert.deepEqual(Array.from(image.data.subarray(0, 8)), [137, 80, 78, 71, 13, 10, 26, 10]);
  assert.match(image.src, /^data:image\/png;base64,/);
  assert.equal(image.widthPx, 80.016);
  assert.ok(Math.abs(image.heightPx - 39.984) < 0.001);
});

test("the flat .fodt form reads into the same text and picture", async () => {
  const packaged = await core.buildDocModelFromBytes(fixtureBytes("sample-html-features.odt"));
  const flat = await core.buildDocModelFromBytes(fixtureBytes("sample-html-features.fodt"));

  assert.equal(core.isOpenDocumentModel(flat.model), true);
  assert.deepEqual(
    paragraphs(flat.model.nodes).map(paragraphText),
    paragraphs(packaged.model.nodes).map(paragraphText),
  );
  const [image] = images(flat.model.nodes);
  assert.equal(image.contentType, "image/png");
  assert.match(image.src, /^data:image\/png;base64,/);
});

test("text frames and page fields survive the flat OpenDocument reader", async () => {
  const { model } = await core.buildDocModelFromBytes(
    fixtureBytes("frames-and-fields.fodt")
  );

  const frames = images(model.nodes).filter((image) => image.syntheticTextBox);
  assert.ok(frames.length >= 1, "the floating text frame becomes a synthetic frame");
  const floating = frames.find((image) => image.floating);
  assert.ok(floating);
  assert.equal(floating.floating?.wrapType, "square");
  assert.deepEqual(
    floating.textBox?.paragraphs.map((paragraph) =>
      paragraph.segments.map((segment) => segment.text).join("")
    ),
    ["Floating frame text"],
  );
  assert.match(floating.textBoxSourceXml, /^<w:txbxContent>/);
  assert.ok((floating.widthPx ?? 0) > 0);
  assert.ok((floating.heightPx ?? 0) > 0);

  const footer = model.metadata.footerSections[0];
  assert.equal(footer?.partName, "styles.xml/footer1");
  assert.equal(paragraphs(footer.nodes).map(paragraphText).join(""), "Page 1 of 1");
});

test("the runtime import path returns the OpenDocument model and empty package", async () => {
  const result = await core.importDocxBuffer(
    fixtureBytes("sample-html-features.odt"),
    { useWorker: false },
  );
  assert.equal(result.source, "main-thread");
  assert.equal(result.package.parts.size, 0);
  assert.equal(core.isOpenDocumentModel(result.model), true);
  assert.ok(result.timings.totalMs >= 0);
});

test("OpenDocument models survive cloning and the Word entry point refuses them", async () => {
  const { model } = await core.buildDocModelFromBytes(fixtureBytes("sample-html-features.odt"));
  const clone = core.cloneDocModel(model);
  assert.equal(core.isOpenDocumentModel(clone), true);
  assert.deepEqual(clone.metadata.sourcePartNames, { document: "content.xml" });
  assert.equal(core.isOpenDocumentModel(core.createBlankDocumentModel()), false);

  await assert.rejects(
    core.parseDocx(fixtureBytes("sample-html-features.odt")),
    /OpenDocument text has no Word package/,
  );
  await assert.rejects(
    core.exportDocxBuffer(model),
    (error) => error.code === "INVALID_MODEL" && /view-only/.test(error.message),
  );
});

test("a non-text OpenDocument package is rejected instead of being read as Word", async () => {
  const spreadsheet = strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<office:document
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  office:version="1.3"
  office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
  <office:body><office:spreadsheet><table:table table:name="Sheet1"/></office:spreadsheet></office:body>
</office:document>`);
  await assert.rejects(
    core.buildDocModelFromBytes(spreadsheet),
    /not a text document/,
  );

  const archive = zipSync({
    mimetype: strToU8("application/vnd.oasis.opendocument.spreadsheet"),
    "content.xml": spreadsheet,
  });
  await assert.rejects(
    core.buildDocModelFromBytes(archive),
    /not a text document/,
  );
});

test("the OpenDocument reader enforces input and XML safety budgets", async () => {
  const odt = fixtureBytes("sample-html-features.odt");
  await assert.rejects(
    core.importDocxBuffer(odt, { useWorker: false, limits: { maxInputBytes: 32 } }),
    (error) => error.code === "LIMIT_EXCEEDED" && error.limit === "maxInputBytes",
  );

  const malicious = strToU8(`<?xml version="1.0"?>
<!DOCTYPE office:document [ <!ENTITY xxe SYSTEM "file:///etc/passwd"> ]>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  office:mimetype="application/vnd.oasis.opendocument.text">
  <office:body><office:text><text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">&xxe;</text:p></office:text></office:body>
</office:document>`);
  await assert.rejects(
    core.buildDocModelFromBytes(malicious),
    /document type or entity|禁止的文档类型或实体声明/i,
  );
});

