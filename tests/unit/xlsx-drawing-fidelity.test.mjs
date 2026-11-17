import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueXlsx = createRequire(
  new URL("../../packages/vue-xlsx/package.json", import.meta.url),
);
const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
const { strFromU8, strToU8, unzipSync, zipSync } = requireFromXlsxCore("fflate");
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;

const core = await import(
  pathToFileURL(requireFromVueXlsx.resolve("@arcships/xlsx-core")).href
);
const samples = new URL("../../apps/demo/public/samples/", import.meta.url);
const PX_TO_EMU = 9525;

function bytesFromFile(name) {
  return new Uint8Array(readFileSync(new URL(name, samples)));
}

test("anchored drawing geometry uses the grid's visible-column width model", () => {
  // Columns 0 and 2 are visible; column 1 is hidden and takes no space.
  const sheet = {
    colWidthOverridesPx: {},
    colWidths: [100, 50, 60],
    defaultColWidthPx: 80,
    defaultRowHeightPx: 24,
    rowHeightOverridesPx: {},
    rowHeights: [30, 40],
    visibleCols: [0, 2, 3],
    visibleRows: [0, 2],
  };
  const anchor = {
    from: { col: 3, colOffsetEmu: 0, row: 2, rowOffsetEmu: 0 },
    kind: "one-cell",
    sizeEmu: { cx: 10 * PX_TO_EMU, cy: 10 * PX_TO_EMU },
  };

  const rect = core.anchorToAbsoluteRect(anchor, sheet);
  assert.equal(core.emuToPixels(rect.x), 150, "cols 0 (100) + hidden 1 (0) + 2 (50)");
  assert.equal(core.emuToPixels(rect.y), 30, "row 0 (30) + hidden row 1 (0)");

  assert.equal(core.resolveSheetColumnWidthPx(sheet, 0), 100);
  assert.equal(core.resolveSheetColumnWidthPx(sheet, 1), 0, "hidden column is zero-width");
  assert.equal(core.resolveSheetColumnWidthPx(sheet, 3), 60);
  assert.equal(core.resolveSheetColumnWidthPx(sheet, 9), 80, "columns past the tracked range use the default");
  assert.equal(core.resolveSheetRowHeightPx(sheet, 1), 0, "hidden row is zero-height");
  assert.equal(core.resolveSheetRowHeightPx(sheet, 2), 40);

  const twoCell = core.anchorToAbsoluteRect({
    from: { col: 2, colOffsetEmu: 0, row: 0, rowOffsetEmu: 0 },
    kind: "two-cell",
    to: { col: 3, colOffsetEmu: 0, row: 2, rowOffsetEmu: 0 },
  }, sheet);
  assert.equal(core.emuToPixels(twoCell.x), 100);
  assert.equal(core.emuToPixels(twoCell.cx), 50, "two-cell width spans only visible columns");
  assert.equal(core.emuToPixels(twoCell.cy), 30);
});

test("picture parsing keeps rotation, flip and source crop from DrawingML", () => {
  const archive = unzipSync(bytesFromFile("charts-images.xlsx"));
  const drawingPath = "xl/drawings/drawing1.xml";
  let drawing = strFromU8(archive[drawingPath]);
  assert.match(drawing, /<pic>/, "fixture should contain a picture");

  drawing = drawing
    .replace(
      /(<a:blip[^>]*\/>)/,
      '$1<a:srcRect xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" l="10000" t="20000" r="10000" b="0"/>',
    )
    .replace(
      /<spPr><a:prstGeom xmlns:a="http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/main" prst="rect"\/><\/spPr>/,
      '<spPr><a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rot="2700000" flipH="1"><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect"/></spPr>',
    );

  archive[drawingPath] = strToU8(drawing);
  const assets = core.parseWorkbookImageAssets(zipSync(archive));
  const image = assets.imagesByWorkbookSheetIndex
    .flat()
    .find((entry) => entry.name === "Image 2");
  assert.ok(image, "the fixture picture should parse");
  assert.equal(image.rotationDeg, 45);
  assert.equal(image.flipH, true);
  assert.equal(image.flipV, false);
  assert.deepEqual(image.srcRect, { bottom: 0, left: 0.1, right: 0.1, top: 0.2 });
});

test("shape parsing keeps paragraph runs, alignment, insets and flip state", () => {
  const assets = core.parseWorkbookImageAssets(bytesFromFile("charts-images.xlsx"));
  const shape = assets.shapesByWorkbookSheetIndex
    .flat()
    .find((entry) => entry.name === "审批提醒");
  assert.ok(shape, "the fixture shape should parse");
  assert.equal(shape.paragraphs[0]?.align, "center");
  const run = shape.paragraphs[0]?.runs[0];
  assert.equal(run?.text, "审批提醒：请核对季度收入");
  assert.equal(run?.bold, true);
  assert.equal(shape.textBox?.verticalAlign, "middle");
  assert.equal(shape.textBox?.insetPx?.left, 9.6);
});
