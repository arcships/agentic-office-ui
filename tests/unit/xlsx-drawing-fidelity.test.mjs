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
const wasmPath = requireFromVueXlsx.resolve("@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm");
const wasmBytes = readFileSync(wasmPath);
core.setWasmSource(
  wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
);
const wasm = await core.getSheetsWasmModule();

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

function chartsArchiveWithDuplicates() {
  const archive = unzipSync(bytesFromFile("charts-images.xlsx"));
  const drawingPath = "xl/drawings/drawing1.xml";
  let drawing = strFromU8(archive[drawingPath]);
  const picture = drawing.match(/<oneCellAnchor>(?:(?!<\/oneCellAnchor>)[\s\S])*?<pic>(?:(?!<\/oneCellAnchor>)[\s\S])*?<\/oneCellAnchor>/)?.[0];
  const shape = drawing.match(/<xdr:twoCellAnchor[^>]*>(?:(?!<\/xdr:twoCellAnchor>)[\s\S])*?<\/xdr:twoCellAnchor>/)?.[0];
  assert.ok(picture, "fixture picture block");
  assert.ok(shape, "fixture shape block");
  const extraPicture = picture
    .replace('id="2"', 'id="20"')
    .replace('name="Image 2"', 'name="Image 20"');
  const extraShape = shape
    .replace('id="9002"', 'id="9003"')
    .replace('name="审批提醒"', 'name="Shape 20"');
  drawing = drawing.replace(shape, `${shape}${extraShape}`).replace(picture, `${picture}${extraPicture}`);
  archive[drawingPath] = strToU8(drawing);
  return archive;
}

test("image delete keeps surviving anchor identities stable", () => {
  const archive = chartsArchiveWithDuplicates();
  const assets = core.parseWorkbookImageAssets(zipSync(archive));
  const images = assets.imagesByWorkbookSheetIndex[0] ?? [];
  assert.equal(images.length, 2);
  const [first, second] = images;
  assert.equal(core.deleteWorkbookImage(assets, first.id), true);
  assert.equal(assets.imagesByWorkbookSheetIndex[0]?.length, 1);

  const nextAnchor = {
    kind: "absolute",
    positionEmu: { x: 42 * PX_TO_EMU, y: 24 * PX_TO_EMU },
    sizeEmu: { cx: 90 * PX_TO_EMU, cy: 60 * PX_TO_EMU },
  };
  assert.equal(core.updateWorkbookImageAnchor(assets, second.id, nextAnchor), true);

  const reparsed = core.parseWorkbookImageAssets(zipSync(assets.archive));
  const survivors = reparsed.imagesByWorkbookSheetIndex[0] ?? [];
  assert.equal(survivors.length, 1);
  assert.equal(survivors[0].name, "Image 20");
  assert.equal(survivors[0].anchor.kind, "absolute");
  assert.equal(survivors[0].anchor.positionEmu.x, 42 * PX_TO_EMU);
});

test("shape update and delete use a stable object identity", () => {
  const archive = chartsArchiveWithDuplicates();
  const assets = core.parseWorkbookImageAssets(zipSync(archive));
  const shapes = assets.shapesByWorkbookSheetIndex[0] ?? [];
  assert.equal(shapes.length, 2);
  const [first, second] = shapes;
  assert.equal(core.deleteWorkbookShape(assets, first.id), true);

  const nextAnchor = {
    from: { col: 6, colOffsetEmu: 0, row: 13, rowOffsetEmu: 0 },
    kind: "two-cell",
    to: { col: 8, colOffsetEmu: 0, row: 15, rowOffsetEmu: 0 },
  };
  assert.equal(core.updateWorkbookShapeAnchor(assets, second.id, nextAnchor), true);
  const reparsed = core.parseWorkbookImageAssets(zipSync(assets.archive));
  const survivors = reparsed.shapesByWorkbookSheetIndex[0] ?? [];
  assert.equal(survivors.length, 1);
  assert.equal(survivors[0].name, "Shape 20");
  assert.equal(survivors[0].anchor.kind, "two-cell");
  assert.equal(survivors[0].anchor.from.col, 6);
});

test("chart delete removes the drawing frame and keeps the archive exportable", () => {
  const bytes = bytesFromFile("charts-images.xlsx");
  const workbook = wasm.Workbook.fromBytes(bytes);
  try {
    const imageAssets = core.parseWorkbookImageAssets(bytes);
    const chartAssets = core.loadWorkbookChartAssets(workbook, imageAssets, new Map([[0, 0]]));
    const chart = chartAssets.chartsByWorkbookSheetIndex[0]?.[0];
    assert.ok(chart);
    assert.equal(core.deleteWorkbookChart(imageAssets, chartAssets, chart.id), true);
    assert.equal(chartAssets.chartsByWorkbookSheetIndex[0]?.length, 0);
    assert.doesNotMatch(strFromU8(imageAssets.archive["xl/drawings/drawing1.xml"]), /<c:chart[\s>]/);
    const merged = core.mergeWorkbookImageAssets(workbook.saveXlsxBytes(), imageAssets, imageAssets.sheetOrigins);
    assert.ok(merged.byteLength > 0);
  } finally {
    workbook.free();
  }
});

test("form-control CRUD rewrites sheet controls and ctrlProps", () => {
  const assets = core.parseWorkbookImageAssets(bytesFromFile("charts-images.xlsx"));
  const input = {
    anchor: {
      kind: "absolute",
      positionEmu: { x: 30 * PX_TO_EMU, y: 40 * PX_TO_EMU },
      sizeEmu: { cx: 120 * PX_TO_EMU, cy: 24 * PX_TO_EMU },
    },
    checked: false,
    kind: "checkbox",
    label: "Approve",
    linkedCell: "$C$3",
    name: "Approve",
  };
  const added = core.addWorkbookFormControl(assets, 0, input);
  assert.ok(added);
  assert.equal(added.shapeId > 9001, true);
  assert.equal(assets.formControlsByWorkbookSheetIndex[0]?.length, 2);

  let reparsed = core.parseWorkbookImageAssets(zipSync(assets.archive));
  let control = reparsed.formControlsByWorkbookSheetIndex[0]?.find((entry) => entry.shapeId === added.shapeId);
  assert.ok(control, "added control must survive a reparse");
  assert.equal(control.kind, "checkbox");
  assert.equal(control.linkedCell, "$C$3");

  assert.equal(core.updateWorkbookFormControl(assets, added.id, {
    anchor: {
      kind: "absolute",
      positionEmu: { x: 55 * PX_TO_EMU, y: 66 * PX_TO_EMU },
      sizeEmu: { cx: 100 * PX_TO_EMU, cy: 30 * PX_TO_EMU },
    },
    checked: true,
    linkedCell: "$D$4",
  }), true);
  reparsed = core.parseWorkbookImageAssets(zipSync(assets.archive));
  control = reparsed.formControlsByWorkbookSheetIndex[0]?.find((entry) => entry.shapeId === added.shapeId);
  assert.ok(control);
  assert.equal(control.checked, true);
  assert.equal(control.linkedCell, "$D$4");
  assert.equal(control.anchor.kind, "absolute");
  assert.equal(control.anchor.positionEmu.x, 55 * PX_TO_EMU);

  assert.equal(core.deleteWorkbookFormControl(assets, added.id), true);
  reparsed = core.parseWorkbookImageAssets(zipSync(assets.archive));
  assert.equal(
    reparsed.formControlsByWorkbookSheetIndex[0]?.some((entry) => entry.shapeId === added.shapeId),
    false,
  );
});

test("anchored rect cache reuses content rects until sheet or zoom changes", () => {
  const sheet = {
    colWidths: [100],
    defaultColWidthPx: 100,
    defaultRowHeightPx: 20,
    rowHeights: [20],
    visibleCols: [0],
    visibleRows: [0],
  };
  const anchor = {
    kind: "absolute",
    positionEmu: { x: PX_TO_EMU, y: 2 * PX_TO_EMU },
    sizeEmu: { cx: 10 * PX_TO_EMU, cy: 5 * PX_TO_EMU },
  };
  const cache = core.createAnchoredRectCache();
  const first = cache.resolve(anchor, sheet, 100);
  const second = cache.resolve(anchor, sheet, 100);
  assert.equal(first, second, "same anchor/sheet/zoom returns the cached rect instance");
  const zoomed = cache.resolve(anchor, sheet, 50);
  assert.notEqual(first, zoomed);
  assert.equal(zoomed.width, first.width / 2);
});
