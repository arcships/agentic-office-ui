import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

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
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);
const samples = new URL("../../apps/demo/public/samples/", import.meta.url);
const PX_TO_EMU = 9525;
const wasmPath = requireFromXlsxCore.resolve(
  "@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm",
);
const wasmBytes = readFileSync(wasmPath);
core.setWasmSource(
  wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
);
const wasm = await core.getSheetsWasmModule();

function bytesFromFile(name) {
  return new Uint8Array(readFileSync(new URL(name, samples)));
}

/**
 * Adds a second, drawing-free worksheet to the charts-images fixture so
 * cross-sheet paste has a target that starts without a drawing part.
 */
function twoSheetFixture() {
  const archive = unzipSync(bytesFromFile("charts-images.xlsx"));
  archive["xl/workbook.xml"] = strToU8(
    strFromU8(archive["xl/workbook.xml"]).replace(
      "</sheets>",
      '<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Target" sheetId="3" r:id="rIdTarget"/></sheets>',
    ),
  );
  archive["xl/_rels/workbook.xml.rels"] = strToU8(
    strFromU8(archive["xl/_rels/workbook.xml.rels"]).replace(
      "</Relationships>",
      '<Relationship Id="rIdTarget" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
    ),
  );
  archive["xl/worksheets/sheet2.xml"] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
  );
  return zipSync(archive, { level: 6 });
}

function loadFixture() {
  const bytes = twoSheetFixture();
  const workbook = wasm.Workbook.fromBytes(bytes);
  const assets = core.parseWorkbookImageAssets(bytes);
  const chartAssets = core.loadWorkbookChartAssets(
    workbook,
    assets,
    new Map([[0, 0], [1, 1]]),
    false,
  );
  return { assets, bytes, chartAssets, workbook };
}

function oneCellAnchor(col, row, widthPx = 160, heightPx = 90) {
  return {
    from: { col, colOffsetEmu: 0, row, rowOffsetEmu: 0 },
    kind: "one-cell",
    sizeEmu: { cx: widthPx * PX_TO_EMU, cy: heightPx * PX_TO_EMU },
  };
}

function drawingFor(assets, chartAssets, kind) {
  if (kind === "image") return assets.imagesByWorkbookSheetIndex[0][0];
  if (kind === "shape") return assets.shapesByWorkbookSheetIndex[0][0];
  return chartAssets.chartsByWorkbookSheetIndex[0][0];
}

test("drawing clipboard items carry model, source sheet, relative anchor and size", () => {
  const { assets, chartAssets } = loadFixture();
  const sheetState = assets.sheetStatesByWorkbookSheetIndex[0];
  const drawings = ["image", "shape", "chart"].map((kind) => ({
    kind,
    model: drawingFor(assets, chartAssets, kind),
  }));
  const items = core.createDrawingClipboardGroup(drawings, sheetState);

  assert.deepEqual(items.map((item) => item.kind), ["image", "shape", "chart"]);
  for (const item of items) {
    assert.equal(item.sourceSheetId, 0);
    assert.equal(item.sourceId, item.model.id);
    assert.ok(item.size.cx > 0 && item.size.cy > 0, `${item.kind} keeps a rendered size`);
    assert.ok(item.relativeAnchor.x >= 0 && item.relativeAnchor.y >= 0);
  }
  const originX = Math.min(...drawings.map(({ model }) => (
    core.anchorToAbsoluteRect(model.anchor, sheetState).x
  )));
  const originY = Math.min(...drawings.map(({ model }) => (
    core.anchorToAbsoluteRect(model.anchor, sheetState).y
  )));
  for (const [index, item] of items.entries()) {
    const rect = core.anchorToAbsoluteRect(drawings[index].model.anchor, sheetState);
    assert.equal(item.relativeAnchor.x, rect.x - originX);
    assert.equal(item.relativeAnchor.y, rect.y - originY);
  }
});

test("cloning an image, chart and shape pastes them onto another sheet", () => {
  const { assets, chartAssets, workbook } = loadFixture();
  try {
    const clipboard = core.createDrawingClipboardGroup(
      ["image", "chart", "shape"].map((kind) => ({
        kind,
        model: drawingFor(assets, chartAssets, kind),
      })),
      assets.sheetStatesByWorkbookSheetIndex[0],
    );
    const sourceIds = clipboard.map((item) => item.sourceId);
    const clones = clipboard.map((item, index) => core.cloneWorkbookDrawing(
      assets,
      chartAssets,
      {
        id: item.sourceId,
        kind: item.kind,
        workbookSheetIndex: item.sourceSheetId,
      },
      1,
      oneCellAnchor(1 + index, 2 + index),
      { visibleSheetIndex: 1 },
    ));

    assert.ok(clones.every(Boolean), "every drawing clones");
    assert.deepEqual(
      assets.imagesByWorkbookSheetIndex[1].map((image) => image.id),
      [clones[0].id],
    );
    assert.deepEqual(
      chartAssets.chartsByWorkbookSheetIndex[1].map((chart) => chart.id),
      [clones[1].id],
    );
    assert.deepEqual(
      assets.shapesByWorkbookSheetIndex[1].map((shape) => shape.id),
      [clones[2].id],
    );
    assert.equal(new Set([...sourceIds, ...clones.map((clone) => clone.id)]).size, 6);
    assert.equal(clones[1].model.chartPath.startsWith("xl/charts/chartClipboard"), true);
    assert.notEqual(clones[1].model.chartPath, drawingFor(assets, chartAssets, "chart").chartPath);
    assert.deepEqual(
      clones[1].model.series.map((series) => series.id),
      clones[1].model.series.map((_, index) => `${clones[1].id}-series-${index}`),
    );
    assert.equal(assets.sheetOrigins[1].attachments.length, 1, "target gains a drawing attachment");

    const exported = core.mergeWorkbookImageAssets(
      workbook.saveXlsxBytes(),
      assets,
      assets.sheetOrigins,
    );
    const reparsed = core.parseWorkbookImageAssets(exported);
    const reparsedWorkbook = wasm.Workbook.fromBytes(exported);
    try {
      const reparsedCharts = core.loadWorkbookChartAssets(
        reparsedWorkbook,
        reparsed,
        new Map([[0, 0], [1, 1]]),
        false,
      );
      assert.equal(reparsed.imagesByWorkbookSheetIndex[1].length, 1);
      assert.equal(reparsed.shapesByWorkbookSheetIndex[1].length, 1);
      assert.equal(reparsedCharts.chartsByWorkbookSheetIndex[1].length, 1);
      assert.equal(reparsed.imagesByWorkbookSheetIndex[0].length, 1, "source picture survives");
      assert.equal(reparsed.shapesByWorkbookSheetIndex[0].length, 1, "source shape survives");
      assert.equal(reparsedCharts.chartsByWorkbookSheetIndex[0].length, 1, "source chart survives");
      assert.equal(
        reparsed.imagesByWorkbookSheetIndex[1][0].anchor.from.col,
        1,
        "pasted picture keeps the target anchor",
      );
      assert.equal(reparsedCharts.chartsByWorkbookSheetIndex[1][0].anchor.from.col, 2);
    } finally {
      reparsedWorkbook.free();
    }
  } finally {
    workbook.free();
  }
});

test("pasted chart owns a separate chart part and relationship ids are remapped", () => {
  const { assets, chartAssets, workbook } = loadFixture();
  try {
    const source = drawingFor(assets, chartAssets, "chart");
    const clone = core.cloneWorkbookDrawing(
      assets,
      chartAssets,
      { id: source.id, kind: "chart", workbookSheetIndex: 0 },
      0,
      oneCellAnchor(10, 20),
    );
    assert.ok(clone);
    const drawingPath = chartAssets.chartOriginsById.get(clone.id).drawingPath;
    const drawingXml = strFromU8(assets.archive[drawingPath]);
    const relationshipsXml = strFromU8(
      assets.archive[`xl/drawings/_rels/${drawingPath.split("/").pop()}.rels`],
    );
    assert.match(drawingXml, /r:id="rIdClipboard\d+"/);
    assert.match(relationshipsXml, /Id="rIdClipboard\d+"/);
    assert.match(relationshipsXml, /chartClipboard1\.xml/);
    const sourceXmlBefore = strFromU8(assets.archive[source.chartPath]);
    assert.equal(core.changeWorkbookChartType(assets, chartAssets, clone.id, "Pie"), true);
    assert.equal(
      strFromU8(assets.archive[source.chartPath]),
      sourceXmlBefore,
      "editing the pasted chart leaves the source part alone",
    );
    assert.match(strFromU8(assets.archive[clone.model.chartPath]), /pieChart/);
    assert.notEqual(
      core.deleteWorkbookChart(assets, chartAssets, clone.id),
      false,
      "the clone can be deleted on its own",
    );
    assert.equal(chartAssets.chartsByWorkbookSheetIndex[0].length, 1);
    assert.ok(assets.archive[source.chartPath], "the source chart part is untouched");
  } finally {
    workbook.free();
  }
});

test("clone converts the anchor kind when the target geometry asks for it", () => {
  const { assets, chartAssets, workbook } = loadFixture();
  try {
    const source = drawingFor(assets, chartAssets, "image");
    assert.equal(source.anchor.kind, "one-cell");
    const twoCell = {
      from: { col: 2, colOffsetEmu: 0, row: 4, rowOffsetEmu: 0 },
      kind: "two-cell",
      to: { col: 6, colOffsetEmu: 0, row: 9, rowOffsetEmu: 0 },
    };
    const clone = core.cloneWorkbookDrawing(
      assets,
      chartAssets,
      { id: source.id, kind: "image", workbookSheetIndex: 0 },
      1,
      twoCell,
    );
    assert.ok(clone);
    assert.deepEqual(clone.model.anchor, twoCell);
    const drawingPath = assets.imageOriginsById.get(clone.id).drawingPath;
    const drawingXml = strFromU8(assets.archive[drawingPath]);
    assert.match(drawingXml, /twoCellAnchor/);
    assert.match(drawingXml, /r:embed="rIdClipboard\d+"/);
    const reparsed = core.parseWorkbookImageAssets(zipSync(assets.archive));
    assert.equal(reparsed.imagesByWorkbookSheetIndex[1].length, 1);
    assert.equal(reparsed.imagesByWorkbookSheetIndex[1][0].anchor.kind, "two-cell");
  } finally {
    workbook.free();
  }
});

test("removeWorkbookDrawing moves a cut picture, shape and chart without touching siblings", () => {
  const { assets, chartAssets, workbook } = loadFixture();
  try {
    const clones = ["image", "shape", "chart"].map((kind, index) => core.cloneWorkbookDrawing(
      assets,
      chartAssets,
      {
        id: drawingFor(assets, chartAssets, kind).id,
        kind,
        workbookSheetIndex: 0,
      },
      1,
      oneCellAnchor(index, 5),
    ));
    assert.ok(clones.every(Boolean));

    for (const clone of clones) {
      assert.equal(core.removeWorkbookDrawing(
        assets,
        chartAssets,
        { id: clone.id, kind: clone.kind, workbookSheetIndex: 1 },
      ), true, `${clone.kind} is removed`);
    }
    assert.equal(assets.imagesByWorkbookSheetIndex[1].length, 0);
    assert.equal(assets.shapesByWorkbookSheetIndex[1].length, 0);
    assert.equal(chartAssets.chartsByWorkbookSheetIndex[1].length, 0);
    assert.equal(assets.imagesByWorkbookSheetIndex[0].length, 1);
    assert.equal(assets.shapesByWorkbookSheetIndex[0].length, 1);
    assert.equal(chartAssets.chartsByWorkbookSheetIndex[0].length, 1);

    const exported = core.mergeWorkbookImageAssets(
      workbook.saveXlsxBytes(),
      assets,
      assets.sheetOrigins,
    );
    const reparsed = core.parseWorkbookImageAssets(exported);
    assert.equal(reparsed.imagesByWorkbookSheetIndex[1].length, 0);
    assert.equal(reparsed.shapesByWorkbookSheetIndex[1].length, 0);
  } finally {
    workbook.free();
  }
});
