// Issue #72 — the 0.8 fixture corpus must contain the structures that the
// individual viewer tests cannot synthesize cheaply (a real chart part inside a
// DOCX, SmartArt and a video placeholder in a deck, an OLE compound file for a
// macro-enabled workbook, an AcroForm and an OCG in a PDF). This file opens the
// fixtures as packages and fails when a regeneration silently drops a feature.
//
// The heavier, parser-specific assertions live in `scripts/verify_test_materials.py`
// (python-docx/openpyxl/pypdf) and in the format suites that consume the same
// files. Here the contract is the package structure and the feature tags in the
// manifest/coverage matrix, so a broken fixture is caught before the browser
// suites spend time loading it.

import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const { strFromU8, unzipSync } = requireFromXlsxCore("fflate");
const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const docxCore = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href,
);

const SAMPLES = new URL("../../apps/demo/public/samples/", import.meta.url);
const MANIFEST = new URL("../../test-data/manifest.json", import.meta.url);
const COVERAGE = new URL("../../test-data/corpus-coverage.json", import.meta.url);

const NEW_FIXTURES = [
  "docx-embedded-objects.docx",
  "docx-fields-controls.docx",
  "xlsx-macros.xlsm",
  "xlsx-protected-1904.xlsx",
  "xlsx-chart-families.xlsx",
  "xlsx-negative-databar-icons.xlsx",
  "xlsx-large-scale.xlsx",
  "deck-structure.pptx",
  "pdf-acroform-annotations.pdf",
  "pdf-1000-pages.pdf",
];

function sampleBytes(name) {
  return new Uint8Array(readFileSync(new URL(name, SAMPLES)));
}

function sampleZip(name, filter) {
  const archive = unzipSync(sampleBytes(name), filter ? { filter } : undefined);
  return {
    names: Object.keys(archive),
    text(part) {
      const bytes = archive[part];
      assert.ok(bytes, `${name}: missing package part ${part}`);
      return strFromU8(bytes);
    },
  };
}

function manifestEntry(name) {
  const manifest = JSON.parse(readFileSync(MANIFEST, "utf8"));
  const entry = manifest.entries.find((candidate) => candidate.file === name);
  assert.ok(entry, `manifest is missing ${name}`);
  return entry;
}

test("the 0.8 fixtures carry feature tags and a coverage-matrix test", () => {
  const manifest = JSON.parse(readFileSync(MANIFEST, "utf8"));
  const coverage = JSON.parse(readFileSync(COVERAGE, "utf8"));
  assert.equal(manifest.coverage, "test-data/corpus-coverage.json");
  assert.equal(manifest.entries.length, coverage.fixtures.length);
  assert.equal(coverage.fixtures.length, 38);

  const byFile = new Map(coverage.fixtures.map((item) => [item.file, item]));
  for (const name of NEW_FIXTURES) {
    const entry = manifestEntry(name);
    assert.ok(Array.isArray(entry.features) && entry.features.length > 0, `${name}: features`);
    const item = byFile.get(name);
    assert.ok(item, `${name}: missing from coverage matrix`);
    assert.deepEqual(item.features, entry.features, `${name}: feature tags drifted`);
    assert.ok(item.issues.includes(72), `${name}: coverage must link issue #72`);
    assert.ok(
      item.tests.includes("tests/unit/fixture-corpus.test.mjs"),
      `${name}: fixture is not referenced by this test`,
    );
    assert.ok(
      item.tests.includes("scripts/verify_test_materials.py"),
      `${name}: fixture is not checked by material verification`,
    );
    for (const script of item.tests) {
      assert.ok(existsSync(new URL(`../../${script}`, import.meta.url)), `${name}: missing ${script}`);
    }
  }
});

test("DOCX fixtures carry text boxes, shapes, SmartArt, charts, OLE, OMML, fields and moves", () => {
  const embedded = sampleZip("docx-embedded-objects.docx");
  const document = embedded.text("word/document.xml");
  assert.equal((document.match(/<w:txbxContent>/g) ?? []).length, 1, "text box content");
  assert.equal((document.match(/<wps:wsp\b/g) ?? []).length, 2, "text box plus shape");
  assert.equal((document.match(/<dgm:relIds\b/g) ?? []).length, 1, "SmartArt relationship block");
  assert.equal((document.match(/<c:chart\b/g) ?? []).length, 1, "chart frame");
  assert.match(document, /<w:object\b[\s\S]*<o:OLEObject\b/, "OLE object wrapper");
  assert.equal((document.match(/<m:oMathPara>/g) ?? []).length, 1, "display equation");
  for (const part of [
    "word/charts/chart1.xml",
    "word/diagrams/data1.xml",
    "word/diagrams/layout1.xml",
    "word/diagrams/quickStyle1.xml",
    "word/diagrams/colors1.xml",
    "word/embeddings/oleObject1.bin",
    "word/media/ole-preview.png",
  ]) {
    assert.ok(embedded.names.includes(part), `missing ${part}`);
  }
  const ole = unzipSync(sampleBytes("docx-embedded-objects.docx"))["word/embeddings/oleObject1.bin"];
  assert.equal(Buffer.from(ole.subarray(0, 8)).toString("hex"), "d0cf11e0a1b11ae1");
  assert.match(embedded.text("word/charts/chart1.xml"), /<c:barChart>/);
  const rels = embedded.text("word/_rels/document.xml.rels");
  for (const target of ["charts/chart1.xml", "diagrams/data1.xml", "embeddings/oleObject1.bin"]) {
    assert.ok(rels.includes(target), `missing relationship target ${target}`);
  }

  const fields = sampleZip("docx-fields-controls.docx");
  const body = fields.text("word/document.xml");
  assert.equal((body.match(/<w:sdt\b/g) ?? []).length, 2, "block and inline content controls");
  assert.equal((body.match(/<w:moveFrom\b/g) ?? []).length, 1, "moveFrom revision");
  assert.equal((body.match(/<w:moveTo\b/g) ?? []).length, 1, "moveTo revision");
  assert.ok(body.includes(" PAGE \\* roman "), "PAGE roman field");
  assert.ok(body.includes(" NUMPAGES \\* ALPHA "), "NUMPAGES letter field");
  assert.ok(body.includes(' STYLEREF "Heading 1" \\* ROMAN '), "STYLEREF roman field");
  assert.equal((fields.text("word/footnotes.xml").match(/<w:footnote w:id="/g) ?? []).length, 2);
  const tableStart = body.indexOf("<w:tbl>");
  const tableEnd = body.indexOf("</w:tbl>", tableStart);
  const footnote = body.indexOf('<w:footnoteReference w:id="2"/>');
  assert.ok(tableStart !== -1 && tableEnd !== -1 && footnote > tableStart && footnote < tableEnd,
    "the table footnote anchor must stay inside the table");
});

test("the DOCX importer resolves the embedded objects and table footnote", async () => {
  const bytes = readFileSync(new URL("docx-embedded-objects.docx", SAMPLES));
  const { model } = await docxCore.importDocxBuffer(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    { useWorker: false },
  );
  const children = model.nodes.flatMap((node) => node.children ?? []);
  const drawingKinds = children
    .filter((child) => child.type === "drawing")
    .map((child) => child.kind);
  assert.deepEqual(
    [...new Set(drawingKinds)].sort(),
    ["chart", "oleObject", "shape", "smartArt"],
  );
  assert.ok(
    children.some((child) => child.type === "image" && child.syntheticTextBox),
    "the text box must import as a synthetic text box",
  );
  const math = children.find((child) => child.type === "math");
  assert.equal(math?.display, true);
  assert.match(math?.text ?? "", /x\^2\+a\/b/);

  const fieldBytes = readFileSync(new URL("docx-fields-controls.docx", SAMPLES));
  const fieldImport = await docxCore.importDocxBuffer(
    fieldBytes.buffer.slice(fieldBytes.byteOffset, fieldBytes.byteOffset + fieldBytes.byteLength),
    { useWorker: false },
  );
  const footnotes = fieldImport.model.metadata.footnotes ?? [];
  assert.equal(footnotes.length, 2);
  assert.ok(footnotes.some((note) => note.text.includes("anchored inside a table cell")));
});

test("XLSX chart-family fixture carries every plot family and both combo axis modes", () => {
  const xlsx = sampleZip("xlsx-chart-families.xlsx");
  const chartParts = xlsx.names.filter((name) => /^xl\/charts\/chart\d+\.xml$/.test(name));
  assert.equal(chartParts.length, 12);
  const chartXml = chartParts.map((part) => xlsx.text(part)).join("\n");
  for (const family of [
    "barChart",
    "bar3DChart",
    "lineChart",
    "pieChart",
    "doughnutChart",
    "scatterChart",
    "radarChart",
    "areaChart",
  ]) {
    assert.ok(chartXml.includes(`<${family}>`) || chartXml.includes(`<c:${family}>`), `missing ${family}`);
  }
  const combos = chartParts
    .map((part) => xlsx.text(part))
    .filter((xml) => xml.includes("barChart") && xml.includes("lineChart"));
  assert.equal(combos.length, 2, "one-axis and two-axis combo charts");
  assert.equal(combos.filter((xml) => (xml.match(/<valAx>/g) ?? []).length >= 2).length, 1);
});

test("XLSX conditional-format fixture uses negative data bars and 4/5-icon sets", () => {
  const xlsx = sampleZip("xlsx-negative-databar-icons.xlsx");
  const bars = xlsx.text("xl/worksheets/sheet1.xml");
  assert.match(bars, /type="dataBar"/);
  assert.ok((bars.match(/<v>-\d+<\/v>/g) ?? []).length >= 3, "negative data values");
  assert.match(bars, /cfvo type="num" val="-50"/);
  assert.match(bars, /cfvo type="percentile"/);
  assert.match(bars, /type="expression"/);
  const icons = xlsx.text("xl/worksheets/sheet2.xml");
  assert.match(icons, /iconSet="5Rating"/);
  assert.match(icons, /iconSet="4Arrows"/);
  assert.match(icons, /cfvo type="percent"/);
});

test("XLSX protection and 1904-date fixture keeps both passwords and the epoch", () => {
  const xlsx = sampleZip("xlsx-protected-1904.xlsx");
  const workbook = xlsx.text("xl/workbook.xml");
  assert.match(workbook, /date1904="1"/);
  assert.match(workbook, /workbookProtection[^>]*workbookPassword="9721"[^>]*lockStructure="1"/);
  const sheet = xlsx.text("xl/worksheets/sheet2.xml");
  assert.match(sheet, /sheetProtection[^>]*sheet="1"/);
  assert.match(sheet, /sheetProtection[^>]*password="9721"/);
  assert.match(sheet, /formatCells="0"/);
});

test("XLSX macro fixture is an inert, macro-enabled package", () => {
  const xlsx = sampleZip("xlsx-macros.xlsm");
  const contentTypes = xlsx.text("[Content_Types].xml");
  assert.match(contentTypes, /application\/vnd\.ms-excel\.sheet\.macroEnabled\.main\+xml/);
  assert.match(contentTypes, /application\/vnd\.ms-office\.vbaProject/);
  assert.match(xlsx.text("xl/_rels/workbook.xml.rels"), /vbaProject\.bin/);
  const vba = unzipSync(sampleBytes("xlsx-macros.xlsm"))["xl/vbaProject.bin"];
  assert.equal(Buffer.from(vba.subarray(0, 8)).toString("hex"), "d0cf11e0a1b11ae1");
  const text = Buffer.from(vba).toString("latin1");
  assert.ok(text.includes("Inert Agentic Office UI fixture text"), "stub marker");
  assert.ok(!text.includes("Sub Auto_Open"), "the fixture must not contain runnable macro code");
});

test("the large XLSX fixture reaches 100k rows and 10k columns", () => {
  const rows = sampleZip("xlsx-large-scale.xlsx", (file) => file.name === "xl/worksheets/sheet1.xml");
  const rowSheet = rows.text("xl/worksheets/sheet1.xml");
  assert.match(rowSheet, /<row r="100001"><c r="A100001" t="n"><v>100000<\/v><\/c><c r="B100001" t="n"><v>200000<\/v><\/c><\/row>/);
  const columns = sampleZip("xlsx-large-scale.xlsx", (file) => file.name === "xl/worksheets/sheet2.xml");
  const columnSheet = columns.text("xl/worksheets/sheet2.xml");
  assert.match(columnSheet, /<c r="NTP2" t="n"><v>10000<\/v><\/c>/);
  assert.match(columnSheet, /<c r="NTP1" t="inlineStr"><is><t>C10000<\/t><\/is><\/c>/);
});

test("PPTX fixture carries a master, layouts, placeholders, SmartArt, video and CJK runs", () => {
  const pptx = sampleZip("deck-structure.pptx");
  assert.ok(pptx.names.includes("ppt/slideMasters/slideMaster1.xml"));
  assert.equal(pptx.names.filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name)).length, 3);
  assert.equal(pptx.names.filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length, 3);
  const master = pptx.text("ppt/slideMasters/slideMaster1.xml");
  assert.match(master, /<p:ph type="title"/);
  assert.match(master, /<p:ph type="body"/);
  assert.match(master, /<p:ph type="ftr"/);
  const slide = pptx.text("ppt/slides/slide2.xml");
  assert.match(slide, /<dgm:relIds/);
  assert.match(slide, /typeface="Microsoft YaHei"/);
  assert.match(slide, /[\u4e00-\u9fff]/);
  assert.ok(pptx.names.includes("ppt/diagrams/drawing1.xml"), "SmartArt drawing fallback");
  const video = pptx.text("ppt/slides/slide3.xml");
  assert.match(video, /<a:videoFile/);
  assert.match(video, /p14:media/);
  assert.match(video, /<p:ph type="pic"/);
  assert.ok(pptx.names.includes("ppt/media/video1.mp4"));
  assert.ok(pptx.names.includes("ppt/media/poster.png"));
});

test("PDF fixtures expose AcroForm, annotations, attachments, OCG, signature and 1000 pages", () => {
  const form = Buffer.from(sampleBytes("pdf-acroform-annotations.pdf"));
  for (const token of [
    "/AcroForm",
    "/Highlight",
    "/FreeText",
    "/EmbeddedFiles",
    "/OCProperties",
    "/DocMDP",
    "/ByteRange",
  ]) {
    assert.ok(form.includes(Buffer.from(token)), `AcroForm fixture is missing ${token}`);
  }
  const large = Buffer.from(sampleBytes("pdf-1000-pages.pdf"));
  assert.equal((large.toString("latin1").match(/\/Type\s*\/Page\b/g) ?? []).length, 1000);
  assert.ok(large.length < 2 * 1024 * 1024, "1000-page fixture must stay below 2 MiB");
});
