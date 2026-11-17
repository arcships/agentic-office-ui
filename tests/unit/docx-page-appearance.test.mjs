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
