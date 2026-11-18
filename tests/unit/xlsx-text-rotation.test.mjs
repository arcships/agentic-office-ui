import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import test from "node:test";
import { pathToFileURL } from "node:url";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;

const core = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);

const ROOT = path.resolve(import.meta.dirname, "../..");
const GRID_PAINT_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/grid-paint.ts");
const GRID_PAINT_SOURCE = readFileSync(GRID_PAINT_PATH, "utf8");

function loadTsModule(source) {
  const output = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
  return import(`data:text/javascript;base64,${Buffer.from(output).toString("base64")}`);
}

const paint = await loadTsModule(GRID_PAINT_SOURCE);
const {
  resolveCellTextLayout,
  resolveRotatedTextBounds,
  resolveRotatedTextCenter,
  resolveRotatedTextFitWidth,
  resolveShrinkToFitFontSize,
  resolveSpreadsheetTextRotation,
} = paint;

const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

/**
 * `charts-images.xlsx` with four extra cell formats and four header cells:
 * A1 rotated 45°, B1 90°, C1 180° and D1 shrink-to-fit, matching the fixtures
 * named in issue #141.
 */
function rotationWorkbookBytes() {
  const archive = unzipSync(new Uint8Array(readFileSync(new URL("charts-images.xlsx", samples))));
  const alignmentXfs = [
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1" pivotButton="0" quotePrefix="0" xfId="0"><alignment textRotation="45"/></xf>',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1" pivotButton="0" quotePrefix="0" xfId="0"><alignment textRotation="90"/></xf>',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1" pivotButton="0" quotePrefix="0" xfId="0"><alignment textRotation="180"/></xf>',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1" pivotButton="0" quotePrefix="0" xfId="0"><alignment shrinkToFit="1"/></xf>',
  ].join("");
  let styles = strFromU8(archive["xl/styles.xml"]);
  styles = styles
    .replace('<cellXfs count="5">', '<cellXfs count="9">')
    .replace("</cellXfs>", `${alignmentXfs}</cellXfs>`);
  assert.match(styles, /<cellXfs count="9">/, "the fixture must add four cell formats");
  archive["xl/styles.xml"] = strToU8(styles);

  let sheet = strFromU8(archive["xl/worksheets/sheet1.xml"]);
  sheet = sheet
    .replace(
      '<c r="A1" s="1" t="inlineStr"><is><t>Quarter</t></is></c>',
      '<c r="A1" s="5" t="inlineStr"><is><t>Rotated 45 header</t></is></c>',
    )
    .replace(
      '<c r="B1" s="1" t="inlineStr"><is><t>Revenue</t></is></c>',
      '<c r="B1" s="6" t="inlineStr"><is><t>Rotated 90 header</t></is></c>',
    )
    .replace(
      '<c r="C1" s="1" t="inlineStr"><is><t>Exceptions</t></is></c>',
      '<c r="C1" s="7" t="inlineStr"><is><t>Rotated 180 header</t></is></c>',
    )
    .replace(
      '<c r="D1" s="2" t="inlineStr"><is><t>&#36235;&#21183;</t></is></c>',
      '<c r="D1" s="8" t="inlineStr"><is><t>Shrink this long value</t></is></c>',
    );
  assert.match(sheet, /s="5"/, "A1 must carry the 45° style");
  assert.match(sheet, /s="8"/, "D1 must carry the shrink-to-fit style");
  archive["xl/worksheets/sheet1.xml"] = strToU8(sheet);
  return zipSync(archive);
}

test("XLSX parser keeps cell text rotation and shrink-to-fit for renderers", () => {
  const assets = core.parseWorkbookStructureAssets(rotationWorkbookBytes());
  const state = assets.sheetStatesByWorkbookSheetIndex[0];
  assert.ok(state, "the fixture sheet parses");
  assert.deepEqual(state.cellTextLayouts, {
    A1: { textRotation: 45 },
    B1: { textRotation: 90 },
    C1: { textRotation: 180 },
    D1: { shrinkToFit: true },
  });
  assert.equal(assets.styleById[5].alignment.textRotation, 45);
  assert.equal(assets.styleById[8].alignment.shrinkToFit, true);
});

test("Excel rotation maps to the canvas angle, including the >90° clockwise half", () => {
  assert.equal(resolveSpreadsheetTextRotation(45), -45, "45° Excel rotation is counter-clockwise");
  assert.equal(resolveSpreadsheetTextRotation(90), -90, "90° is vertical bottom-to-top");
  assert.equal(resolveSpreadsheetTextRotation(180), 90, "180° is vertical top-to-bottom");
  assert.equal(resolveSpreadsheetTextRotation(120), 30, "120° means 30° clockwise");
  assert.equal(resolveSpreadsheetTextRotation(91), 1, "91° means 1° clockwise");
  assert.equal(resolveSpreadsheetTextRotation(0), null, "zero rotation stays horizontal");
  assert.equal(resolveSpreadsheetTextRotation(181), null, "out-of-range values stay horizontal");
  assert.equal(resolveSpreadsheetTextRotation(255), null, "the legacy vertical-stack marker stays horizontal");
  assert.equal(resolveSpreadsheetTextRotation("45"), null, "non-numeric rotation is ignored");
});

test("cell text layout reads both engine and OOXML alignment spellings", () => {
  assert.deepEqual(resolveCellTextLayout({ rotation: 45, shrinkToFit: true }), {
    rotation: -45,
    shrinkToFit: true,
  });
  assert.deepEqual(resolveCellTextLayout({ textRotation: 180 }), {
    rotation: 90,
    shrinkToFit: false,
  });
  assert.deepEqual(resolveCellTextLayout({ shrinkToFit: true }), {
    rotation: null,
    shrinkToFit: true,
  });
  assert.deepEqual(resolveCellTextLayout(undefined), { rotation: null, shrinkToFit: false });
});

test("rotation swaps the occupied width and height of the text box", () => {
  const bounds = resolveRotatedTextBounds(100, 10, -90);
  assert.ok(Math.abs(bounds.width - 10) < 1e-9, "a 90° box takes the text height as its width");
  assert.ok(Math.abs(bounds.height - 100) < 1e-9, "a 90° box takes the text width as its height");

  const diagonal = resolveRotatedTextBounds(100, 10, -45);
  assert.ok(diagonal.width > 10 && diagonal.width < 100);
  assert.ok(diagonal.height > 10 && diagonal.height < 100);

  assert.equal(resolveRotatedTextFitWidth(80, 24, -90, 12), 24, "vertical text may be as tall as the row");
  assert.equal(resolveRotatedTextFitWidth(80, 24, 0, 12), 80, "horizontal text uses the column width");

  const center = resolveRotatedTextCenter(
    { height: 24, left: 0, top: 0, width: 80 },
    bounds,
    "right",
    "bottom",
  );
  assert.equal(center.x, 75, "right alignment anchors by the rotated width");
  assert.equal(center.y, -26, "bottom alignment anchors by the rotated height");
});

test("shrink-to-fit binary-searches down to the 0.4× floor", () => {
  const linear = (base) => (sizePx) => 100 * sizePx / base;
  assert.equal(
    resolveShrinkToFitFontSize(10, 200, linear(10)),
    10,
    "text that already fits keeps its size",
  );

  const fitted = resolveShrinkToFitFontSize(10, 50, linear(10));
  assert.ok(Math.abs(fitted - 5) < 0.2, `expected about 5px, got ${fitted}`);

  assert.equal(
    resolveShrinkToFitFontSize(10, 30, linear(10)),
    4,
    "text that cannot fit stops at the 0.4× floor",
  );
  assert.equal(
    resolveShrinkToFitFontSize(10, 0, linear(10)),
    10,
    "a zero-width cell keeps the base size instead of dividing by zero",
  );
});
