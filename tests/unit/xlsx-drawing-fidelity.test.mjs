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

function assetsWithShapeGeometry(prst, { adjustments, connector = false, line } = {}) {
  const archive = unzipSync(bytesFromFile("charts-images.xlsx"));
  const drawingPath = "xl/drawings/drawing1.xml";
  let drawing = strFromU8(archive[drawingPath]);
  const avLst = adjustments
    ? `<a:avLst>${Object.entries(adjustments)
        .map(([name, value]) => `<a:gd name="${name}" fmla="val ${value}"/>`)
        .join("")}</a:avLst>`
    : "<a:avLst/>";
  assert.match(drawing, /prst="roundRect"/, "fixture shape should use preset geometry");
  drawing = drawing.replace(
    /<a:prstGeom prst="roundRect"><a:avLst\/><\/a:prstGeom>/,
    `<a:prstGeom prst="${prst}">${avLst}</a:prstGeom>`,
  );
  if (line) {
    drawing = drawing.replace('<a:ln w="19050">', `<a:ln w="19050">${line}`);
  }
  if (connector) {
    drawing = drawing.replaceAll("<xdr:sp>", "<xdr:cxnSp>").replaceAll("</xdr:sp>", "</xdr:cxnSp>");
  }
  archive[drawingPath] = strToU8(drawing);
  return core.parseWorkbookImageAssets(zipSync(archive));
}

test("preset shape geometry builds Excel preset paths with guide adjustments", () => {
  const rect = core.buildPresetShapePath("rect", undefined, { height: 50, width: 100 });
  assert.deepEqual(rect, {
    path: "M0,0 L100,0 L100,50 L0,50 Z",
    viewBox: { height: 50, width: 100 },
  });

  const rightArrow = core.buildPresetShapePath("rightArrow", undefined, { height: 80, width: 160 });
  assert.equal(rightArrow.path, "M0,20 L120,20 L120,0 L160,40 L120,80 L120,60 L0,60 Z");
  assert.deepEqual(rightArrow.viewBox, { height: 80, width: 160 });
  assert.equal(
    core.buildPresetShapePath("rightArrow", { adj1: 25000, adj2: 50000 }, { height: 80, width: 160 }).path,
    "M0,30 L120,30 L120,0 L160,40 L120,80 L120,50 L0,50 Z",
    "a:avLst guides must feed the preset builders",
  );

  assert.equal(
    core.buildPresetShapePath("chevron", undefined, { height: 80, width: 160 }).path,
    "M0,0 L120,0 L160,40 L120,80 L0,80 L40,40 Z",
  );
  assert.equal(
    core.buildPresetShapePath("chevron", { adj: 25000 }, { height: 80, width: 160 }).path,
    "M0,0 L140,0 L160,40 L140,80 L0,80 L20,40 Z",
  );

  const star5 = core.buildPresetShapePath("star5", undefined, { height: 200, width: 200 });
  assert.equal((star5.path.match(/L/g) ?? []).length, 9, "ten star vertices: one move plus nine lines");
  assert.match(star5.path, /^M100,0 /);
  assert.match(star5.path, / Z$/);

  const callout = core.buildPresetShapePath("callout1", undefined, { height: 80, width: 160 });
  const borderCallout = core.buildPresetShapePath("borderCallout1", undefined, { height: 80, width: 160 });
  assert.equal(callout.path, borderCallout.path, "callout aliases share the border callout outline");
  assert.equal((callout.path.match(/M/g) ?? []).length, 2, "callout body plus the leader line");
  for (const index of [2, 3]) {
    const aliased = core.buildPresetShapePath(`callout${index}`, undefined, { height: 80, width: 160 });
    assert.equal(
      aliased.path,
      core.buildPresetShapePath(`borderCallout${index}`, undefined, { height: 80, width: 160 }).path,
    );
    assert.match(aliased.path, /M0,0 L160,0 L160,80 L0,80 Z M/, `callout${index} keeps the body and leader`);
  }

  assert.equal(
    core.buildPresetShapePath("flowChartDecision", undefined, { height: 80, width: 160 }).path,
    "M80,0 L160,40 L80,80 L0,40 Z",
  );
  assert.equal(
    core.buildPresetShapePath("FLOWCHARTDECISION", undefined, { height: 80, width: 160 }).path,
    core.buildPresetShapePath("flowChartDecision", undefined, { height: 80, width: 160 }).path,
    "prst names are matched case-insensitively",
  );

  assert.equal(core.buildPresetShapePath("not-a-preset", undefined, { height: 80, width: 160 }), undefined);
  assert.equal(core.buildPresetShapePath("custom", undefined, { height: 80, width: 160 }), undefined);
  assert.equal(core.buildPresetShapePath(undefined, undefined, { height: 80, width: 160 }), undefined);
  assert.deepEqual(
    core.buildPresetShapePath("rightArrow", undefined, { height: 0, width: 0 }).viewBox,
    { height: 1, width: 1 },
    "zero-sized connectors keep a drawable view box",
  );

  assert.equal(core.isStrokeOnlyShapeGeometry("bentConnector3"), true);
  assert.equal(core.isStrokeOnlyShapeGeometry("Line"), true);
  assert.equal(core.isStrokeOnlyShapeGeometry("rightArrow"), false);
});

test("shape parsing emits preset svg paths instead of falling back to rectangles", () => {
  const assets = assetsWithShapeGeometry("rightArrow", { adjustments: { adj1: 60000, adj2: 40000 } });
  const shape = assets.shapesByWorkbookSheetIndex
    .flat()
    .find((entry) => entry.name === "审批提醒");
  assert.ok(shape);
  assert.equal(shape.geometry, "rightArrow");
  assert.equal(shape.svgPath, "M0,12 L168,12 L168,0 L192,30 L168,60 L168,48 L0,48 Z");
  assert.deepEqual(shape.svgViewBox, { height: 60, width: 192 });
  assert.deepEqual(shape.geometryAdjustments, { adj1: 60000, adj2: 40000 });
});

test("connector shapes keep their dash, arrow ends and stroke-only fill", () => {
  const assets = assetsWithShapeGeometry("bentConnector3", {
    connector: true,
    line: '<a:prstDash val="dash"/><a:headEnd type="triangle"/><a:tailEnd type="stealth"/>',
  });
  const shape = assets.shapesByWorkbookSheetIndex
    .flat()
    .find((entry) => entry.name === "审批提醒");
  assert.ok(shape);
  assert.equal(shape.geometry, "bentConnector3");
  assert.equal(shape.fill?.none, true, "connectors must not be filled");
  assert.equal(shape.stroke?.dash, "dash");
  assert.equal(shape.stroke?.headEndType, "triangle");
  assert.equal(shape.stroke?.tailEndType, "stealth");
  assert.match(shape.svgPath, /^M0,0 L/, "connectors use the preset connector path");
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
