import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const { DOMParser } = requireFromXlsxCore("@xmldom/xmldom");

const ROOT = path.resolve(import.meta.dirname, "../..");
const TEXT_DIRECTION_PATH = path.join(ROOT, "packages/pptx-core/src/playback/text-direction.ts");
const MARKERS_PATH = path.join(ROOT, "packages/pptx-core/src/browser/renderer-markers.ts");

async function loadTextDirection() {
  const output = ts.transpileModule(readFileSync(TEXT_DIRECTION_PATH, "utf8"), {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
  return import(`data:text/javascript;base64,${Buffer.from(output).toString("base64")}`);
}

const { applyPptxParagraphDirection, collectPptxTextDirectionFeatures } = await loadTextDirection();

const SLIDE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:bodyPr rtlCol="1"/>
          <a:p><a:pPr algn="r" rtl="1"/><a:r><a:t>مرحبا بالعالم</a:t></a:r></a:p>
          <a:p><a:pPr rtl="0"/><a:r><a:t>Left to right</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

test("PPTX parser records RTL paragraph support and column-order limits", () => {
  const document = new DOMParser().parseFromString(SLIDE_XML, "application/xml");
  const features = collectPptxTextDirectionFeatures(document, 2);

  const paragraph = features.find((feature) => feature.feature === "text:rtl-paragraph");
  assert.ok(paragraph, "an a:pPr/@rtl='1' paragraph is discovered");
  assert.equal(paragraph.slideIndex, 2);
  assert.equal(paragraph.disposition, "strict", "the renderer applies the paragraph direction");
  assert.equal(paragraph.id, "text-rtl:2");

  const columns = features.find((feature) => feature.feature === "text:rtl-columns");
  assert.ok(columns, "a:bodyPr/@rtlCol is recorded as a limitation");
  assert.equal(columns.disposition, "static", "multi-column order is not mirrored");
  assert.match(columns.reason ?? "", /多栏/);
});

test("PPTX parser records nothing for left-to-right slides", () => {
  const document = new DOMParser().parseFromString(
    SLIDE_XML.replace('rtl="1"', 'rtl="0"').replace(' rtlCol="1"', ""),
    "application/xml",
  );
  assert.deepEqual(collectPptxTextDirectionFeatures(document, 0), []);
});

test("rendered paragraphs carry dir=rtl when the renderer sets direction: rtl", () => {
  function paragraph(style) {
    const element = {
      dataset: {},
      attributes: {},
      style: { direction: style },
      setAttribute(name, value) {
        this.attributes[name] = value;
      },
    };
    return element;
  }
  const rtl = paragraph("rtl");
  const ltr = paragraph("ltr");
  const none = paragraph("");
  const root = {
    querySelectorAll(selector) {
      assert.equal(selector, "[data-pptx-paragraph-index]");
      return [rtl, ltr, none];
    },
  };

  assert.equal(applyPptxParagraphDirection(root), 1, "one mirrored paragraph");
  assert.equal(rtl.attributes.dir, "rtl");
  assert.equal(rtl.dataset.pptxParagraphDirection, "rtl");
  assert.equal(ltr.attributes.dir, "ltr");
  assert.equal(ltr.dataset.pptxParagraphDirection, "ltr");
  assert.equal(none.attributes.dir, undefined, "paragraphs without a direction stay untouched");
  assert.equal(none.dataset.pptxParagraphDirection, undefined);
});

test("the renderer markers forward paragraph direction after a slide renders", () => {
  const markers = readFileSync(MARKERS_PATH, "utf8");
  assert.match(markers, /applyPptxParagraphDirection\(root\)/);
  assert.match(markers, /import \{ applyPptxParagraphDirection \} from "\.\.\/playback\/text-direction"/);
});
