// Issue #33 — the DOCX model could not represent text boxes, shapes, SmartArt,
// charts, embedded OLE objects, 3D models or equations, and drop caps were
// parsed but never drawn. Every block names the bullet it covers:
//
//   1. node model: text boxes, shapes/groups/WordArt ......... import
//   2. charts carry the chart part's cached data ............. chart part
//   3. SmartArt, OLE objects and 3D models ................... import
//   4. equations become a math node with linear text ......... m:oMath
//   5. export preserves unsupported objects after an edit .... serialize/reload
//   6. text box edits reach the preserved drawing ............ textBoxText
//   7. header/footer and package exports keep the objects .... serializeDocModel
//   8. placeholder render inputs + chart preview ............. helpers
//   9. drop caps resolve their frame metrics ................. resolveDropCap*

import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromDocxCore.resolve("@arcships/docx-core")).href
);
const { strFromU8, unzipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);

const WORD_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml"';
const SECTION_PROPS =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>';

// ── Fixtures ───────────────────────────────────────────────────────

function documentXml(body) {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<w:document ${WORD_NS}><w:body>${body}<w:sectPr>${SECTION_PROPS}</w:sectPr></w:body></w:document>`
  );
}

function chartDrawing(relationshipId = "rIdChart") {
  return (
    '<w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/>' +
    '<wp:docPr id="11" name="Chart 1" descr="Quarterly revenue"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    `<c:chart r:id="${relationshipId}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
  );
}

function shapeDrawing() {
  return (
    '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" ' +
    'simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>360000</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="914400" cy="457200"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="12" name="Callout"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp><wps:cNvSpPr/><wps:spPr><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>' +
    '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>' +
    '<a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></wps:spPr>' +
    '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
  );
}

function smartArtDrawing() {
  return (
    '<w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="13" name="Diagram 1"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
    '<dgm:relIds r:dm="rIdDm" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  );
}

function oleObjectRun() {
  return (
    '<w:r><w:object w:dxaOrig="1440" w:dyaOrig="1440">' +
    '<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:24pt;height:24pt"/>' +
    '<o:OLEObject Type="Embed" ProgID="Equation.3" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1234"/>' +
    '</w:object></w:r>'
  );
}

function model3dDrawing() {
  return (
    '<w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="14" name="3D Model"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp><wps:spPr/><a:model3d><a:blip r:embed="rIdPoster"/></a:model3d></wps:wsp>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  );
}

function textBoxDrawing() {
  return (
    '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
    '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" ' +
    'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="1828800" cy="457200"/><wp:wrapNone/><wp:docPr id="2" name="Text Box 2"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr/><wps:txbx><w:txbxContent>' +
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Quarterly</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>Review pack</w:t></w:r></w:p>' +
    '</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>' +
    '</wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>' +
    '<w:p><w:r><w:t>Quarterly</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict>' +
    '</mc:Fallback></mc:AlternateContent></w:r>'
  );
}

function mathRun() {
  return (
    '<m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>' +
    '<m:r><m:t>+</m:t></m:r><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e>' +
    '<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>'
  );
}

const CHART_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart>' +
  '<c:title><c:tx><c:rich><a:p><a:r><a:t>Quarterly revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>' +
  '<c:plotArea><c:barChart><c:barDir val="col"/>' +
  '<c:ser><c:idx val="0"/>' +
  '<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>' +
  '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/>' +
  '<c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>' +
  '<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/>' +
  '<c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt></c:numCache></c:numRef></c:val>' +
  '</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>';

const POSTER_PNG = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d,
]);

function packageWith(documentXmlValue, { chart = false, poster = false } = {}) {
  const base = core.createMinimalDocxPackage(documentXmlValue);
  const relationships = [
    '<Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>',
    '<Relationship Id="rIdPoster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/model3d-poster.png"/>',
  ].join("");
  const parts = new Map(base.parts);
  parts.set("word/_rels/document.xml.rels", {
    name: "word/_rels/document.xml.rels",
    content:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
  });
  if (chart) {
    parts.set("word/charts/chart1.xml", {
      name: "word/charts/chart1.xml",
      content: CHART_XML,
    });
  }
  const binaryAssets = new Map(base.binaryAssets);
  if (poster) {
    binaryAssets.set("word/media/model3d-poster.png", POSTER_PNG);
  }
  return { parts, binaryAssets };
}

function paragraphChildren(model, nodeIndex = 0) {
  const node = model.nodes[nodeIndex];
  assert.equal(node.type, "paragraph", `node ${nodeIndex} is a paragraph`);
  return node.children;
}

function documentXmlOf(bytes) {
  return strFromU8(unzipSync(new Uint8Array(bytes))["word/document.xml"]);
}

function countOf(xml, pattern) {
  return [...xml.matchAll(pattern)].length;
}

// ── 1. Text boxes, shapes, groups and WordArt ──────────────────────
test("a text box keeps parsed content on the model instead of raw XML only", async () => {
  const pkg = packageWith(documentXml(`<w:p><w:r><w:t>Before</w:t></w:r>${textBoxDrawing()}</w:p>`));
  const model = await core.buildDocModel(pkg);
  const children = paragraphChildren(model);
  const image = children.find((child) => child.type === "image");
  assert.ok(image, "the text box imports as the synthetic image run");
  assert.equal(image.syntheticTextBox, true);
  assert.ok(image.textBox, "the model carries the text box content");
  assert.equal(image.textBox.paragraphs.length, 2);
  assert.equal(image.textBox.paragraphs[0].align, "center");
  assert.equal(image.textBox.paragraphs[0].segments[0].style?.bold, true);
  assert.equal(core.textBoxContentText(image.textBox), "Quarterly\nReview pack");
  assert.equal(core.syntheticTextBoxTextValue(image), "Quarterly\nReview pack");
  assert.ok(image.textBoxSourceXml, "the original drawing is kept for the export");
});

test("a shape becomes an anchored drawing node with fill, stroke and z-order", async () => {
  const pkg = packageWith(documentXml(`<w:p>${shapeDrawing()}</w:p>`));
  const model = await core.buildDocModel(pkg);
  const drawing = paragraphChildren(model).find((child) => child.type === "drawing");
  assert.ok(drawing, "the shape is a drawing node");
  assert.equal(drawing.kind, "shape");
  assert.equal(drawing.name, "Callout");
  assert.equal(drawing.widthPx, 96);
  assert.equal(drawing.heightPx, 48);
  assert.equal(drawing.floating?.wrapType, "square");
  assert.equal(drawing.floating?.xPx, 37.795);
  assert.equal(drawing.floating?.yPx, 48);
  assert.equal(drawing.floating?.zIndex, 251658240);
  assert.equal(drawing.zIndex, 251658240);
  assert.equal(drawing.shape?.preset, "roundRect");
  assert.equal(drawing.shape?.fillColor, "#00FF00");
  assert.equal(drawing.shape?.strokeColor, "#FF0000");
  assert.match(drawing.sourceXml, /<wps:wsp>/);
});

test("a shape group imports as one drawing object", async () => {
  const groupXml =
    '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="15" name="Group 1"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
    '<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr/><wps:wsp><wps:spPr><a:prstGeom prst="rect"/></wps:spPr><wps:bodyPr/></wps:wsp>' +
    '</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const model = await core.buildDocModel(packageWith(documentXml(`<w:p>${groupXml}</w:p>`)));
  const drawing = paragraphChildren(model).find((child) => child.type === "drawing");
  assert.ok(drawing);
  assert.equal(drawing.kind, "group");
});

// ── 2. Charts carry the chart part's data ──────────────────────────
test("a chart drawing carries the chart part's cached data", async () => {
  const pkg = packageWith(documentXml(`<w:p>${chartDrawing()}</w:p>`), { chart: true });
  const model = await core.buildDocModel(pkg);
  const chart = paragraphChildren(model).find((child) => child.type === "drawing");
  assert.ok(chart, "the chart is a drawing node, not a missing image");
  assert.equal(chart.kind, "chart");
  assert.equal(chart.alt, "Quarterly revenue");
  assert.equal(chart.chart?.chartType, "barChart");
  assert.equal(chart.chart?.title, "Quarterly revenue");
  assert.deepEqual(chart.chart?.categories, ["Q1", "Q2"]);
  assert.equal(chart.chart?.series?.[0].name, "Revenue");
  assert.deepEqual(chart.chart?.series?.[0].values, [10, 30]);
  assert.equal(core.chartPlaceholderText(chart.chart), "Quarterly revenue · barChart · 1 series");
  const preview = core.chartPreviewDataUri(chart.chart, 200, 120);
  assert.ok(preview?.startsWith("data:image/svg+xml"));
  assert.match(decodeURIComponent(preview), /<rect[^>]*fill="#4472C4"/);
});

// ── 3. SmartArt, OLE objects and 3D models ────────────────────────
test("SmartArt and embedded OLE objects become drawing nodes", async () => {
  const pkg = packageWith(
    documentXml(`<w:p>${smartArtDrawing()}</w:p><w:p>${oleObjectRun()}</w:p>`),
  );
  const model = await core.buildDocModel(pkg);
  const smartArt = paragraphChildren(model, 0).find((child) => child.type === "drawing");
  const ole = paragraphChildren(model, 1).find((child) => child.type === "drawing");
  assert.equal(smartArt?.kind, "smartArt");
  assert.equal(smartArt?.name, "Diagram 1");
  assert.equal(ole?.kind, "oleObject");
  assert.equal(ole?.ole?.progId, "Equation.3");
  assert.match(ole?.sourceXml ?? "", /<o:OLEObject/);
});

test("a 3D model keeps its poster frame and object markup", async () => {
  const pkg = packageWith(documentXml(`<w:p>${model3dDrawing()}</w:p>`), { poster: true });
  const model = await core.buildDocModel(pkg);
  const drawing = paragraphChildren(model).find((child) => child.type === "drawing");
  assert.equal(drawing?.kind, "model3d");
  assert.match(drawing?.src ?? "", /^data:image\/png;base64,/);
  assert.match(drawing?.sourceXml ?? "", /<a:model3d>/);
});

// ── 4. Equations ───────────────────────────────────────────────────
test("an equation becomes a math node with linear text", async () => {
  const pkg = packageWith(documentXml(`<w:p><w:r><w:t>Where</w:t></w:r>${mathRun()}</w:p>`));
  const model = await core.buildDocModel(pkg);
  const math = paragraphChildren(model).find((child) => child.type === "math");
  assert.ok(math, "the equation is a math node");
  assert.equal(math.text, "a/b+x^2");
  assert.equal(math.display, undefined, "an inline m:oMath stays inline");
  assert.match(math.sourceXml, /<m:oMath>/);
  assert.equal(core.linearizeOmml(math.sourceXml), "a/b+x^2");
});

test("a display equation is marked as a block equation", async () => {
  const pkg = packageWith(
    documentXml(`<w:p><m:oMathPara>${mathRun()}</m:oMathPara></w:p>`),
  );
  const model = await core.buildDocModel(pkg);
  const math = paragraphChildren(model).find((child) => child.type === "math");
  assert.equal(math?.display, true);
});

// ── 5. Export preserves objects an edit would drop ────────────────
test("an edited paragraph keeps its chart, shape and equation on export", async () => {
  const body =
    `<w:p><w:r><w:t>Chart</w:t></w:r>${chartDrawing()}</w:p>` +
    `<w:p><w:r><w:t>Shape</w:t></w:r>${shapeDrawing()}</w:p>` +
    `<w:p><w:r><w:t>Math</w:t></w:r>${mathRun()}</w:p>`;
  const pkg = packageWith(documentXml(body), { chart: true });
  const model = await core.buildDocModel(pkg);

  // An edit clears `sourceXml`; the engine rebuilds each paragraph from the
  // model, which has no slot for these objects.
  const edited = core.updateParagraphText(
    core.updateParagraphText(core.updateParagraphText(model, 0, "Chart edited"), 1, "Shape edited"),
    2,
    "Math edited",
  );
  edited.nodes[2].sourceXml = undefined;

  const xml = await core.modelToDocumentXml(edited, pkg);
  assert.match(xml, /Chart edited/);
  assert.equal(countOf(xml, /<c:chart\b/g), 1, "the chart is written back");
  assert.equal(countOf(xml, /<wps:wsp>/g), 1, "the shape is written back");
  assert.equal(countOf(xml, /<m:oMath>/g), 1, "the equation is written back");
  assert.equal(countOf(xml, /<w:drawing>/g), 2);
  assert.match(xml, /<m:t>a<\/m:t>/);

  // Export → reload keeps the objects and the edited text.
  const reopened = (await core.buildDocModelFromBytes(await core.serializeDocx(edited, pkg))).model;
  assert.equal(core.paragraphText(reopened.nodes[0]), "Chart edited");
  const chartAfter = reopened.nodes[0].children.find((child) => child.type === "drawing");
  assert.equal(chartAfter?.kind, "chart");
  assert.deepEqual(chartAfter?.chart?.series?.[0].values, [10, 30]);
  const shapeAfter = reopened.nodes[1].children.find((child) => child.type === "drawing");
  assert.equal(shapeAfter?.kind, "shape");
  const mathAfter = reopened.nodes[2].children.find((child) => child.type === "math");
  assert.equal(mathAfter?.text, "a/b+x^2");
});

test("an untouched paragraph is written verbatim, not patched twice", async () => {
  const model = await core.buildDocModel(
    packageWith(documentXml(`<w:p>${shapeDrawing()}</w:p>`)),
  );
  const xml = await core.modelToDocumentXml(model);
  assert.equal(countOf(xml, /<wps:wsp>/g), 1);
  assert.equal(countOf(xml, /<w:drawing>/g), 1);
});

// ── 6. Text box edits reach the preserved drawing ─────────────────
test("editing a text box writes the new text into the preserved drawing", async () => {
  const pkg = packageWith(
    documentXml(`<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p>${textBoxDrawing()}</w:p>`),
  );
  const model = await core.buildDocModel(pkg);
  const image = paragraphChildren(model, 1).find((child) => child.type === "image");
  assert.ok(image?.textBox);

  const edited = structuredClone(model);
  const editedImage = edited.nodes[1].children.find((child) => child.type === "image");
  // `invalidateImageSource` clears `sourceXml`; `textBoxSourceXml` survives.
  editedImage.sourceXml = undefined;
  editedImage.textBoxText = "Updated box";
  edited.nodes[1].sourceXml = undefined;

  const xml = await core.modelToDocumentXml(edited, pkg);
  assert.match(xml, /<w:txbxContent>/);
  assert.match(xml, /<w:t xml:space="preserve">Updated box<\/w:t>/);
  assert.doesNotMatch(xml, />Quarterly</);
  assert.match(xml, /<w:jc w:val="center"\/>/, "the paragraph properties survive");
  assert.match(xml, /<w:b\/>/, "the run properties survive");

  // Export → reload keeps a text box whose text was edited.
  const reopened = (await core.buildDocModelFromBytes(await core.serializeDocx(edited, pkg))).model;
  const reopenedImage = paragraphChildren(reopened, 1).find((child) => child.type === "image");
  assert.equal(reopenedImage?.syntheticTextBox, true);
  assert.equal(
    core.syntheticTextBoxTextValue(reopenedImage),
    "Updated box",
    "the edited text is what the reopened model reads back",
  );
});

test("a drawing inside a table cell survives a cell edit", async () => {
  const tableXml =
    '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>Cell text</w:t></w:r></w:p>' +
    `<w:p>${shapeDrawing()}</w:p></w:tc></w:tr></w:tbl>`;
  const pkg = packageWith(documentXml(tableXml));
  const model = await core.buildDocModel(pkg);
  const table = model.nodes[0];
  assert.equal(table.type, "table");
  const cellParagraph = table.rows[0].cells[0].nodes[1];
  assert.ok(cellParagraph.children.some((child) => child.type === "drawing"));

  const edited = core.updateTableCellParagraphText(model, 0, 0, 0, 0, "Cell edited");
  const xml = await core.modelToDocumentXml(edited, pkg);
  assert.match(xml, /Cell edited/);
  assert.equal(countOf(xml, /<wps:wsp>/g), 1, "the cell drawing is written back");

  const reopened = (await core.buildDocModelFromBytes(await core.serializeDocx(edited, pkg))).model;
  const reopenedCellParagraph = reopened.nodes[0].rows[0].cells[0].nodes[1];
  assert.equal(reopenedCellParagraph.children.some((child) => child.type === "drawing"), true);
});

// ── 7. Header/footer and package exports ──────────────────────────
test("a shaped header keeps its drawing object through the package export", async () => {
  const model = await core.buildDocModel(
    packageWith(documentXml("<w:p><w:r><w:t>Body</w:t></w:r></w:p>")),
  );
  let next = core.addHeaderFooter(model, 0, "header", { referenceType: "default" });
  const partName = next.metadata.headerSections.at(-1).partName;
  const headerParagraph = {
    type: "paragraph",
    children: [
      { type: "text", text: "Header" },
      {
        type: "drawing",
        kind: "shape",
        name: "Header rule",
        widthPx: 200,
        heightPx: 12,
        sourceXml:
          '<w:drawing><wp:inline><wp:extent cx="1905000" cy="114300"/><wp:docPr id="21" name="Header rule"/>' +
          '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
          '<wps:wsp><wps:spPr><a:prstGeom prst="rect"/></wps:spPr><wps:bodyPr/></wps:wsp>' +
          '</a:graphicData></a:graphic></wp:inline></w:drawing>',
      },
    ],
  };
  next = structuredClone(next);
  next.metadata.headerSections = next.metadata.headerSections.map((section) =>
    section.partName === partName ? { ...section, nodes: [headerParagraph] } : section,
  );
  next.metadata.sections = next.metadata.sections.map((section) => ({
    ...section,
    headerSections: section.headerSections.map((entry) =>
      entry.partName === partName ? { ...entry, nodes: [headerParagraph] } : entry,
    ),
  }));

  const archive = unzipSync(new Uint8Array(await core.serializeDocx(next)));
  assert.equal(countOf(strFromU8(archive[partName]), /<w:drawing>/g), 1);

  const packageExport = await core.serializeDocModel(next);
  assert.equal(
    countOf(packageExport.parts.get(partName).content, /<wps:wsp>/g),
    1,
    "the package-returning serializer patches story parts",
  );

  // Reopening the exported package parses the header shape back into the model.
  const reopened = (await core.buildDocModelFromBytes(await core.serializeDocx(next))).model;
  const reopenedHeader = reopened.metadata.headerSections.find(
    (section) => section.partName === partName,
  );
  assert.ok(reopenedHeader, "the header story is reimported");
  const reopenedDrawing = reopenedHeader.nodes
    .flatMap((node) => (node.type === "paragraph" ? node.children : []))
    .find((child) => child.type === "drawing");
  assert.equal(reopenedDrawing?.kind, "shape");
  assert.equal(reopenedDrawing?.name, "Header rule");
});

test("the engine bridge drops drawing nodes without losing them on export", async () => {
  const model = await core.buildDocModel(
    packageWith(documentXml(`<w:p>${shapeDrawing()}</w:p>`)),
  );
  // `docModelToWasmJson` is the WASM boundary; an unknown child type would make
  // the engine reject the whole model.
  const json = core.docModelToWasmJson(model);
  assert.doesNotMatch(json, /"type":"drawing"/);
  assert.match(json, /"type":"paragraph"/);
  assert.ok((await core.modelToDocumentXml(model)).includes("<w:drawing>"));
});

// ── 8. Placeholder render inputs ──────────────────────────────────
test("drawing placeholders size, anchor and label from the model", () => {
  const floatingShape = {
    type: "drawing",
    kind: "shape",
    widthPx: 120,
    heightPx: 60,
    sourceXml: "<w:drawing/>",
    floating: {
      xPx: 24,
      yPx: 12,
      wrapType: "none",
      behindDocument: true,
      zIndex: 3,
    },
    shape: { fillColor: "#FFCC00", strokeColor: "#003366", strokeWidthPx: 2 },
  };
  assert.deepEqual(core.drawingObjectSizePx(floatingShape), {
    widthPx: 120,
    heightPx: 60,
  });
  const style = core.drawingObjectPlaceholderStyle(floatingShape);
  assert.equal(style.position, "absolute");
  assert.equal(style.left, "24px");
  assert.equal(style.top, "12px");
  assert.equal(style.zIndex, -1, "behind-text objects paint behind the text");
  assert.equal(core.drawingObjectIsBehindText(floatingShape), true);
  const frame = core.drawingObjectFrameStyle(floatingShape);
  assert.equal(frame.backgroundColor, "#FFCC00");
  assert.equal(frame.borderColor, "#003366");
  assert.equal(frame.borderWidth, "2px");

  const inline = core.drawingObjectPlaceholderStyle({
    type: "drawing",
    kind: "smartArt",
    sourceXml: "<w:drawing/>",
  });
  assert.equal(inline.display, "inline-block");
  assert.equal(inline.width, "320px", "a drawing without an extent keeps a default frame");
  assert.equal(core.drawingObjectIsBehindText({ type: "drawing", kind: "shape", sourceXml: "" }), false);
});

test("the chart preview stays empty without cached values", () => {
  assert.equal(core.chartPreviewDataUri(undefined), undefined);
  assert.equal(core.chartPreviewDataUri({ chartType: "barChart" }), undefined);
  assert.equal(core.chartPlaceholderText({ chartType: "pieChart" }), "pieChart");
});

test("drawing import helpers tolerate unknown content", () => {
  assert.equal(core.paragraphDrawingObjects("<w:p><w:r><w:t>text</w:t></w:r></w:p>").length, 0);
  assert.equal(core.classifyDrawingXml("<w:drawing><wp:inline/></w:drawing>"), undefined);
  assert.equal(core.linearizeOmml(undefined), "");
  assert.equal(core.textBoxContentText(undefined), undefined);
});

test("WordArt and text box shapes classify by their own markup", () => {
  const wordArt =
    '<w:drawing><wp:inline><wp:docPr id="31" name="WordArt 1"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp><wps:spPr><a:prstGeom prst="rect"/></wps:spPr>' +
    '<wps:bodyPr><a:prstTxWarp prst="textArchUp"/></wps:bodyPr></wps:wsp>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing>';
  assert.equal(core.classifyDrawingXml(wordArt), "wordArt");

  const textBox =
    '<w:drawing><wp:inline><wp:docPr id="32" name="Text Box 3"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>Label</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
    '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>';
  assert.equal(core.classifyDrawingXml(textBox), "textBox");
  assert.equal(
    core.textBoxContentText(core.parseTextBoxContent(textBox)),
    "Label",
  );

  const picture =
    '<w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';
  assert.equal(core.classifyDrawingXml(picture), "picture");
});

// ── 9. Drop-cap metrics ───────────────────────────────────────────
test("drop caps resolve their frame size from the paragraph", () => {
  const paragraph = {
    type: "paragraph",
    style: { dropCap: { type: "drop", lines: 3 } },
    children: [
      { type: "text", text: "Dropped capital", style: { fontSizePt: 12 } },
    ],
  };
  const style = { fontSizePt: 12 };
  const fontSize = core.resolveDropCapFontSizePx(style, 24, 3);
  const height = core.resolveDropCapVisualHeightPx(style, 24, 3);
  assert.ok(height >= 72, `three lines of 24px make the frame (${height})`);
  assert.ok(fontSize > 0, "the letter keeps a measurable size");
  assert.equal(core.paragraphDropCap(paragraph)?.lines, 3);

  const fromXml = core.paragraphDropCap({
    type: "paragraph",
    sourceXml:
      '<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="2" w:wrap="around"/></w:pPr>' +
      "<w:r><w:t>D</w:t></w:r></w:p>",
    children: [{ type: "text", text: "D" }],
  });
  assert.equal(fromXml?.type, "drop");
  assert.equal(fromXml?.lines, 2);
  assert.equal(fromXml?.wrap, "around");
});
