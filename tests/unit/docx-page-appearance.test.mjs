import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);

test("DOCX page background resolves bare Word colours to CSS colours", () => {
  assert.equal(core.resolveDocumentBackgroundColor("FFF2CC"), "#FFF2CC");
  assert.equal(core.resolveDocumentBackgroundColor("#fff2cc"), "#fff2cc");
  assert.equal(core.resolveDocumentBackgroundColor("auto"), undefined);
  assert.equal(core.resolveDocumentBackgroundColor(""), undefined);
  assert.equal(core.resolveDocumentBackgroundColor("not-a-colour"), undefined);
});

test("DOCX pgBorders resolve per-edge style, offset, colour and shadow", () => {
  const sectionPropertiesXml = [
    "<w:sectPr>",
    '<w:pgBorders w:offsetFrom="page">',
    '<w:top w:val="double" w:sz="12" w:space="24" w:color="FF0000" w:shadow="1"/>',
    '<w:right w:val="single" w:sz="6" w:space="18" w:color="00FF00"/>',
    '<w:bottom w:val="dashed" w:sz="6" w:space="24" w:color="0000FF"/>',
    '<w:left w:val="single" w:sz="6" w:space="20" w:color="112233"/>',
    "</w:pgBorders>",
    "</w:sectPr>",
  ].join("");

  const borders = core.parseSectionPageBorders(sectionPropertiesXml);
  assert.ok(borders);
  assert.equal(borders.top.cssBorder, "2px double #FF0000");
  assert.equal(borders.top.offsetPx, 32, "24pt must convert to 32px");
  assert.equal(borders.top.shadow, true);
  assert.equal(borders.top.color, "#FF0000");
  assert.equal(borders.right.cssBorder, "1px solid #00FF00");
  assert.equal(borders.right.offsetPx, 24);
  assert.equal(borders.bottom.cssBorder, "1px dashed #0000FF");
  assert.equal(borders.left.offsetPx, 27, "20pt must round to 27px");
});

test("DOCX page pattern parses DrawingML, VML and shading fills", () => {
  const drawing = core.parseDocumentBackground([
    '<w:background w:color="DDEBF7">',
    "<w:drawing><wps:wsp><wps:spPr>",
    '<a:pattFill prst="pct20">',
    '<a:fgClr><a:srgbClr val="2F5597"/></a:fgClr>',
    '<a:bgClr><a:srgbClr val="DDEBF7"/></a:bgClr>',
    "</a:pattFill>",
    "</wps:spPr></wps:wsp></w:drawing>",
    "</w:background>",
  ].join(""));
  assert.equal(drawing.color, "#DDEBF7");
  assert.deepEqual(drawing.pattern, {
    preset: "pct20",
    foregroundColor: "#2F5597",
    backgroundColor: "#DDEBF7",
  });

  const vml = core.parseDocumentBackground(
    '<w:background><w:drawing><v:fill type="pattern" pattern="ltHorz" color="112233" color2="FFFFFF"/></w:drawing></w:background>'
  );
  assert.deepEqual(vml.pattern, {
    preset: "ltHorz",
    foregroundColor: "#112233",
    backgroundColor: "#FFFFFF",
  });

  const shading = core.parseDocumentBackground(
    '<w:background><w:shd w:val="pct30" w:color="FF0000" w:fill="FFFFFF"/></w:background>'
  );
  assert.deepEqual(shading.pattern, {
    preset: "pct30",
    foregroundColor: "#FF0000",
    backgroundColor: "#FFFFFF",
  });

  const style = core.pagePatternPaintStyle(drawing.pattern);
  assert.equal(style.backgroundColor, "#DDEBF7");
  assert.match(style.backgroundImage, /^radial-gradient\(circle, #2F5597/);
  assert.equal(style.backgroundSize, "22px 22px");
  assert.equal(style.backgroundRepeat, "repeat");

  const lineStyle = core.pagePatternPaintStyle(vml.pattern);
  assert.match(lineStyle.backgroundImage, /^repeating-linear-gradient\(0deg, #112233/);
});

test("DOCX page background pattern fixture reaches model metadata", async () => {
  const bytes = readFileSync(
    new URL("../fixtures/docx/page-background-pattern.docx", import.meta.url)
  );
  const { model } = await core.buildDocModelFromBytes(bytes);
  assert.equal(model.metadata.documentBackgroundColor, "#DDEBF7");
  assert.deepEqual(model.metadata.documentBackgroundPattern, {
    preset: "pct20",
    foregroundColor: "#2F5597",
    backgroundColor: "#DDEBF7",
  });
  assert.deepEqual(model.nodes[0].children[0], {
    type: "text",
    text: "Patterned page background",
  });

  const wasmResult = await core.wasmBuildDocModelFromBytes(bytes);
  assert.deepEqual(
    wasmResult.model.metadata.documentBackgroundPattern,
    model.metadata.documentBackgroundPattern,
    "the low-level WASM model helper must expose the pattern too"
  );
});
