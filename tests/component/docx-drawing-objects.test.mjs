// Issue #33 — rendered output for drawing objects, equations and drop caps.
//
// The unit test proves the model carries the objects and that export writes them
// back. This file proves the viewer paints them:
//
//   1. a chart drawing renders a preview with the chart's data .... chart
//   2. shapes/SmartArt/OLE/3D models render labelled placeholders .. placeholders
//   3. a text box drawing renders its text content ................ text box
//   4. an equation renders its linear text ........................ math
//   5. a drop cap renders as a floated, oversized first letter ..... drop cap
//   6. a drop-cap paragraph keeps the rest of its text ............. rest text

import assert from "node:assert/strict";
import { test } from "node:test";

import { importFromDemo, mount, walk } from "./vue-test-renderer.mjs";

const { DocxViewerRoot, renderParagraphRuns, renderStaticHtml } = await importFromDemo(
  "@arcships/vue-docx",
);
const core = await importFromDemo("@arcships/docx-core");

function renderParagraph(paragraph, options) {
  const rendered = renderParagraphRuns(
    paragraph,
    options?.key ?? "drawing",
    "light",
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    options,
  );
  const vnodes = Array.isArray(rendered) ? rendered : [rendered];
  return renderStaticHtml(vnodes);
}

function paragraph(children, style) {
  return { type: "paragraph", children, ...(style ? { style } : {}) };
}

function drawing(overrides) {
  return {
    type: "drawing",
    kind: "shape",
    sourceXml: "<w:drawing/>",
    ...overrides,
  };
}

const CHART = {
  chartType: "barChart",
  title: "Quarterly revenue",
  categories: ["Q1", "Q2"],
  series: [{ name: "Revenue", values: [10, 30] }],
};

// ── 1. Charts ──────────────────────────────────────────────────────
test("a chart drawing renders a preview with its data", () => {
  const html = renderParagraph(
    paragraph([drawing({ kind: "chart", widthPx: 240, heightPx: 140, chart: CHART })]),
  );
  assert.match(html, /data-docx-drawing="true"/);
  assert.match(html, /data-docx-drawing-kind="chart"/);
  assert.match(html, /data-docx-drawing-preview="true"/);
  assert.match(html, /src="data:image\/svg\+xml/);
  assert.match(html, /role="img"/);
  assert.match(html, /aria-label="Chart"/);
});

// ── 2. Placeholders ────────────────────────────────────────────────
test("shapes, SmartArt, embedded objects and 3D models render placeholders", () => {
  const html = renderParagraph(
    paragraph([
      drawing({ kind: "shape", name: "Callout" }),
      drawing({ kind: "smartArt", name: "Process" }),
      drawing({ kind: "oleObject", ole: { progId: "Equation.3" } }),
      drawing({ kind: "model3d", src: "data:image/png;base64,AAAA" }),
    ]),
  );
  assert.match(html, /data-docx-drawing-kind="shape"/);
  assert.match(html, /data-docx-drawing-kind="smartArt"/);
  assert.match(html, /data-docx-drawing-kind="oleObject"/);
  assert.match(html, /data-docx-drawing-kind="model3d"/);
  assert.match(html, /data-docx-drawing-placeholder="true"/);
  assert.match(html, />Callout</);
  assert.match(html, />Process</);
  assert.match(html, /Embedded object \(Equation\.3\)/);
  // The 3D model paints its poster frame instead of a label.
  assert.match(html, /src="data:image\/png;base64,AAAA"/);
});

test("a floating behind-text drawing carries its anchor position", () => {
  const html = renderParagraph(
    paragraph([
      drawing({
        kind: "shape",
        floating: { xPx: 20, yPx: 30, wrapType: "none", behindDocument: true },
        widthPx: 100,
        heightPx: 40,
      }),
    ]),
  );
  assert.match(html, /data-docx-drawing-behind="true"/);
  assert.match(html, /position:absolute/);
  assert.match(html, /left:20px/);
  assert.match(html, /top:30px/);
  assert.match(html, /z-index:-1/);
});

// ── 3. Text box drawings ───────────────────────────────────────────
test("a text box drawing renders its model paragraphs", () => {
  const html = renderParagraph(
    paragraph([
      drawing({
        kind: "textBox",
        textBox: {
          paragraphs: [
            {
              align: "center",
              lineHeightPx: 20,
              segments: [{ text: "Quarterly", style: { bold: true } }],
            },
            { lineHeightPx: 18, segments: [{ text: "Review pack" }] },
          ],
        },
      }),
    ]),
  );
  assert.match(html, /data-docx-drawing-textbox="true"/);
  assert.match(html, /data-docx-drawing-textbox-paragraph="true"/);
  assert.match(html, />Quarterly</);
  assert.match(html, />Review pack</);
  assert.match(html, /font-weight:700/);
  assert.match(html, /text-align:center/);
});

// ── 4. Equations ───────────────────────────────────────────────────
test("an equation renders its linear text", () => {
  const html = renderParagraph(
    paragraph([{ type: "math", text: "a/b+x^2", sourceXml: "<m:oMath/>" }]),
  );
  assert.match(html, /data-docx-math="true"/);
  assert.match(html, /role="math"/);
  assert.match(html, />a\/b\+x\^2</);
  assert.match(html, /font-style:italic/);
});

test("a display equation renders as a centred block", () => {
  const html = renderParagraph(
    paragraph([
      { type: "math", text: "y=mx+b", display: true, sourceXml: "<m:oMathPara/>" },
    ]),
  );
  assert.match(html, /data-docx-math-display="true"/);
  assert.match(html, /display:block/);
});

// ── 5. Drop caps ───────────────────────────────────────────────────
test("a drop cap renders as a floated, oversized first letter", () => {
  const html = renderParagraph(
    paragraph(
      [
        { type: "text", text: "Dropped capital starts the article.", style: { fontSizePt: 12 } },
      ],
      { dropCap: { type: "drop", lines: 3 } },
    ),
  );
  assert.match(html, /data-docx-drop-cap="true"/);
  assert.match(html, /data-docx-drop-cap-lines="3"/);
  assert.match(html, /float:left/);
  // The first letter is painted once, by the drop-cap span.
  assert.equal([...html.matchAll(/>D</g)].length, 1);
  const fontSize = Number(/font-size:(\d+)px/.exec(html)?.[1]);
  const lineHeight = Number(/line-height:(\d+)px/.exec(html)?.[1]);
  // Three lines of a 12pt paragraph: the letter spans the whole frame and is
  // much larger than the 16px body text.
  assert.ok(fontSize >= 32, `the letter is scaled (${fontSize}px)`);
  assert.ok(lineHeight >= fontSize, `the frame holds the letter (${lineHeight}px)`);
  assert.match(html, /ropped capital starts the article\./);
});

test("the rest of a drop-cap paragraph keeps its text", () => {
  const html = renderParagraph(
    paragraph(
      [
        { type: "text", text: "W" },
        { type: "text", text: "ith the rest of the sentence." },
      ],
      { dropCap: { type: "drop", lines: 2 } },
    ),
  );
  assert.match(html, /data-docx-drop-cap="true"/);
  assert.match(html, />ith the rest of the sentence\.</);
  assert.equal([...html.matchAll(/>W</g)].length, 1);
});

test("a drop cap can be switched off for host surfaces", () => {
  const html = renderParagraph(
    paragraph(
      [{ type: "text", text: "Plain capital" }],
      { dropCap: { type: "drop", lines: 2 } },
    ),
    { renderDropCaps: false },
  );
  assert.doesNotMatch(html, /data-docx-drop-cap=/);
  assert.match(html, /Plain capital/);
});

// ── 6. Import → page render, end to end ────────────────────────────
const WORD_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';

function chartPartPackage() {
  const body =
    '<w:p><w:r><w:t>Shape and chart</w:t></w:r>' +
    '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="41" name="Callout"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp><wps:spPr><a:prstGeom prst="rect"/></wps:spPr><wps:bodyPr/></wps:wsp>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' +
    '<w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="42" name="Chart 1"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>' +
    '<w:p><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></w:p>' +
    '<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="2"/></w:pPr>' +
    '<w:r><w:t>Dropped capital</w:t></w:r></w:p>';
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document ${WORD_NS}><w:body>${body}` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
    '</w:body></w:document>';
  const base = core.createMinimalDocxPackage(documentXml);
  const parts = new Map(base.parts);
  parts.set("word/_rels/document.xml.rels", {
    name: "word/_rels/document.xml.rels",
    content:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>' +
      '</Relationships>',
  });
  parts.set("word/charts/chart1.xml", {
    name: "word/charts/chart1.xml",
    content:
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:numCache>' +
      '<c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt>' +
      '</c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>',
  });
  return { parts, binaryAssets: new Map() };
}

function withContainerWidth(t, widthPx) {
  const originalWidth = Object.getOwnPropertyDescriptor(Object.prototype, "clientWidth");
  const originalHeight = Object.getOwnPropertyDescriptor(Object.prototype, "clientHeight");
  Object.defineProperty(Object.prototype, "clientWidth", {
    configurable: true,
    get: () => widthPx,
  });
  Object.defineProperty(Object.prototype, "clientHeight", {
    configurable: true,
    get: () => 1200,
  });
  t.after(() => {
    if (originalWidth) Object.defineProperty(Object.prototype, "clientWidth", originalWidth);
    if (originalHeight) Object.defineProperty(Object.prototype, "clientHeight", originalHeight);
  });
}

test("an imported document renders its drawings, equation and drop cap", async (t) => {
  withContainerWidth(t, 1440);
  const model = await core.buildDocModel(chartPartPackage());
  const mounted = await mount(DocxViewerRoot, { model, zoomScale: 100 });
  t.after(() => mounted.app.unmount());

  const html = walk(mounted.root)
    .map((node) => node.props?.innerHTML ?? node.text ?? "")
    .join("");
  assert.match(html, /data-docx-drawing-kind="shape"/);
  assert.match(html, /data-docx-drawing-kind="chart"/);
  assert.match(html, /data-docx-drawing-preview="true"/, "the chart preview is drawn from its data");
  assert.match(html, /data-docx-math="true"/);
  assert.match(html, />x=1</);
  assert.match(html, /data-docx-drop-cap="true"/);
  assert.match(html, />ropped capital</, "the drop cap keeps the rest of the word");
});
