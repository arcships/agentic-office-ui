import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

import {
  CHART_SOURCE_NOTE_TEXT,
  CHART_USER_SHAPES_PATH,
  buildChartMediaWorkbookBytes,
} from "../fixtures/chart-embedded-media.mjs";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;

const { strFromU8, strToU8, unzipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);

const wasmPath = requireFromXlsxCore.resolve(
  "@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm",
);
const wasmBytes = readFileSync(wasmPath);
core.setWasmSource(
  wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
);
const wasm = await core.getSheetsWasmModule();

const CHART_PATH = "xl/charts/chart1.xml";
const CHART_RELS_PATH = "xl/charts/_rels/chart1.xml.rels";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const PICTURE_REL = `${REL_NS}/image`;
const SHAPES_REL = "http://schemas.microsoft.com/office/2011/relationships/chartUserShapes";
const PNG_BYTES = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
  0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
]);

function loadFixture(options) {
  const bytes = buildChartMediaWorkbookBytes(options);
  const workbook = wasm.Workbook.fromBytes(bytes);
  const assets = core.parseWorkbookImageAssets(bytes);
  const chartAssets = core.loadWorkbookChartAssets(
    workbook,
    assets,
    new Map([[0, 0], [1, 1]]),
  );
  const chart = (chartAssets.chartsByWorkbookSheetIndex[0] ?? [])[0];
  assert.ok(chart, "fixture workbook keeps its worksheet chart");
  return { assets, bytes, chart, chartAssets, workbook };
}

function relationships(items) {
  return `<Relationships xmlns="${PKG_REL_NS}">${items.map((item) => (
    `<Relationship Id="${item.id}" Type="${item.type}" Target="${item.target}"${item.mode ? ` TargetMode="${item.mode}"` : ""}/>`
  )).join("")}</Relationships>`;
}

function miniArchive({ chartRels = [], parts = {} } = {}) {
  return {
    [CHART_PATH]: strToU8(
      '<chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart"><chart/></chartSpace>',
    ),
    [CHART_RELS_PATH]: strToU8(relationships(chartRels)),
    ...parts,
  };
}

function decodeDataUrl(dataUrl) {
  const match = /^data:([^;]+);base64,(.+)$/.exec(dataUrl ?? "");
  assert.ok(match, `expected a base64 data URL, got ${String(dataUrl).slice(0, 40)}`);
  return { bytes: Buffer.from(match[2], "base64"), mime: match[1] };
}

test("chart parts expose picture fills and chartUserShapes overlays", () => {
  const { chart, workbook } = loadFixture();
  try {
    assert.ok(chart.images, "chart model carries the chart part images");
    const logo = chart.images.rIdLogo;
    const decoded = decodeDataUrl(logo);
    assert.equal(decoded.mime, "image/png");
    assert.deepEqual([...decoded.bytes.subarray(0, 8)], [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

    assert.equal(chart.series[0].fillImage, logo, "series picture fill resolves to the image");
    assert.equal(chart.series[0].color, undefined, "a picture-filled series has no solid fill colour");
    const pointStyle = chart.series[0].dataPointStyles?.find((entry) => entry.index === 1);
    assert.ok(pointStyle, "the data point picture fill is parsed");
    assert.equal(pointStyle.fillImage, logo);
    assert.equal(pointStyle.color, undefined);

    assert.equal(chart.chartAreaFillImage, logo);
    assert.equal(chart.plotAreaFillImage, logo);

    assert.ok(chart.drawings, "chart model carries the chartUserShapes part");
    const drawing = chart.drawings.rIdShapes;
    assert.ok(drawing);
    assert.match(drawing.xml, /chartUserShapes|wsDr/);
    const shapeImage = decodeDataUrl(drawing.images.rIdImg);
    assert.equal(shapeImage.mime, "image/png");

    assert.equal(drawing.shapes.length, 2);
    const note = drawing.shapes[0];
    assert.equal(note.name, "Source note");
    assert.equal(note.description, "Chart source note");
    assert.equal(note.geometry, "roundRect");
    assert.equal(note.geometryAdjustments.adj, 12000);
    assert.equal(note.fill.color, "#ffffff");
    assert.equal(note.stroke.color, "#1f2937");
    assert.ok(note.stroke.widthPx > 0);
    assert.equal(note.textBox.verticalAlign, "middle");
    assert.equal(note.paragraphs.length, 1);
    assert.equal(note.paragraphs[0].runs[0].text, CHART_SOURCE_NOTE_TEXT);
    assert.equal(note.paragraphs[0].runs[0].bold, true);
    assert.equal(note.paragraphs[0].runs[0].fontSizePt, 10);
    for (const key of ["left", "top", "width", "height"]) {
      assert.ok(note.rect[key] > 0 && note.rect[key] < 1, `note rect.${key} is a chart fraction`);
    }
    assert.ok(note.rect.left + note.rect.width <= 1);
    assert.ok(note.rect.top + note.rect.height <= 1);

    const badge = drawing.shapes[1];
    assert.equal(badge.name, "Chart badge");
    assert.equal(badge.geometry, "rect");
    assert.equal(badge.fill.image, drawing.images.rIdImg, "picture shapes keep their blip fill");
    assert.equal(badge.paragraphs.length, 0);

    assert.equal(chart.series[0].fillImage, chart.images.rIdLogo);
  } finally {
    workbook.free();
  }
});

test("chart areas without chartUserShapes still load picture fills", () => {
  const { chart, workbook } = loadFixture({ skipShapesPart: true });
  try {
    assert.ok(chart.images?.rIdLogo);
    assert.equal(chart.series[0].fillImage, chart.images.rIdLogo);
    assert.equal(chart.drawings, undefined);
  } finally {
    workbook.free();
  }
});

test("chart embedded media honours the byte caps", () => {
  assert.equal(core.CHART_EMBEDDED_IMAGE_MAX_BYTES, 72 * 1024 * 1024);
  assert.equal(core.CHART_EMBEDDED_IMAGE_TOTAL_MAX_BYTES, 100 * 1024 * 1024);
  assert.equal(core.CHART_USER_SHAPES_MAX_BYTES, 8 * 1024 * 1024);

  const padding = new Uint8Array(64);
  const oversized = new Uint8Array([...PNG_BYTES, ...padding]);
  const archive = miniArchive({
    chartRels: [{ id: "rIdImage", type: PICTURE_REL, target: "../media/logo.png" }],
    parts: { "xl/media/logo.png": oversized },
  });

  assert.equal(
    core.readChartEmbeddedMedia(archive, CHART_PATH, { maxImageBytes: 16 }).images,
    undefined,
    "an image over the per-part cap is dropped",
  );
  const included = core.readChartEmbeddedMedia(archive, CHART_PATH, { maxImageBytes: oversized.length });
  assert.ok(included.images?.rIdImage, "the same image is kept under a larger cap");

  const totalArchive = miniArchive({
    chartRels: [
      { id: "rIdOne", type: PICTURE_REL, target: "../media/one.png" },
      { id: "rIdTwo", type: PICTURE_REL, target: "../media/two.png" },
    ],
    parts: {
      "xl/media/one.png": PNG_BYTES,
      "xl/media/two.png": PNG_BYTES,
    },
  });
  const limited = core.readChartEmbeddedMedia(totalArchive, CHART_PATH, {
    maxTotalImageBytes: PNG_BYTES.length,
  });
  assert.deepEqual(Object.keys(limited.images ?? {}), ["rIdOne"], "the total cap stops later images");

  const shapesArchive = miniArchive({
    chartRels: [{ id: "rIdShapes", type: SHAPES_REL, target: "chartUserShapes1.xml" }],
    parts: {
      [CHART_USER_SHAPES_PATH]: strToU8(
        '<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"><twoCellAnchor><from><x>0</x><y>0</y></from><to><x>100</x><y>100</y></to><sp><nvSpPr><cNvPr id="1" name="n"/></nvSpPr><spPr><prstGeom prst="rect"/></spPr><txBody><p><r><t>hello</t></r></p></txBody></sp></twoCellAnchor></wsDr>',
      ),
    },
  });
  assert.equal(
    Object.keys(core.readChartEmbeddedMedia(shapesArchive, CHART_PATH, { maxUserShapesBytes: 16 }).drawings ?? {}).length,
    0,
    "a chartUserShapes part over the cap is dropped",
  );
  const shapes = core.readChartEmbeddedMedia(shapesArchive, CHART_PATH, { maxUserShapesBytes: 4096 });
  assert.equal(Object.keys(shapes.drawings ?? {}).length, 1);
});

test("chart media rejects external, escaped and non-image parts", () => {
  const escapedArchive = miniArchive({
    chartRels: [
      { id: "rIdExternal", type: PICTURE_REL, target: "https://evil.example/logo.png", mode: "External" },
      { id: "rIdFile", type: PICTURE_REL, target: "file:///etc/passwd" },
      { id: "rIdSvg", type: PICTURE_REL, target: "../media/logo.svg" },
      { id: "rIdHtml", type: PICTURE_REL, target: "../media/logo.png" },
      { id: "rIdShapes", type: SHAPES_REL, target: "https://evil.example/shapes.xml", mode: "External" },
    ],
    parts: {
      "xl/media/logo.svg": strToU8('<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>'),
      "xl/media/logo.png": strToU8("<html><script>alert(1)</script></html>"),
    },
  });
  const media = core.readChartEmbeddedMedia(escapedArchive, CHART_PATH);
  assert.equal(media.images, undefined, "external, scheme, SVG and spoofed parts never become images");
  assert.equal(media.drawings, undefined);

  assert.equal(core.sniffChartImageMime(PNG_BYTES), "image/png");
  assert.equal(core.sniffChartImageMime(new Uint8Array([0xff, 0xd8, 0xff, 0xe0])), "image/jpeg");
  assert.equal(core.sniffChartImageMime(strToU8("GIF89a......")), "image/gif");
  assert.equal(core.sniffChartImageMime(strToU8("RIFF....WEBPVP8 ")), "image/webp");
  assert.equal(core.sniffChartImageMime(strToU8("BM......")), "image/bmp");
  assert.equal(core.sniffChartImageMime(strToU8("II*\0......")), "image/tiff");
  assert.equal(core.sniffChartImageMime(strToU8("MM\0*......")), "image/tiff");
  assert.equal(core.sniffChartImageMime(strToU8("<svg/>")), null);

  assert.deepEqual(core.readChartEmbeddedMedia({}, CHART_PATH), {});
  assert.deepEqual(core.readChartEmbeddedMedia(miniArchive(), CHART_PATH), {});
});

test("chartUserShapes parsing keeps anchors, presets and text styles", () => {
  const xml = `<cdr:wsDr xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL_NS}">
    <cdr:absoluteAnchor>
      <cdr:pos x="500000" y="250000"/>
      <cdr:ext cx="2000000" cy="1000000"/>
      <cdr:sp>
        <cdr:nvSpPr><cdr:cNvPr id="7" name="Note" descr="Note description"/><cdr:cNvSpPr/></cdr:nvSpPr>
        <cdr:spPr>
          <a:noFill/>
          <a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/></a:ln>
          <a:prstGeom prst="chevron"><a:avLst><a:gd name="adj" fmla="val 30000"/></a:avLst></a:prstGeom>
        </cdr:spPr>
        <cdr:txBody>
          <a:bodyPr anchor="b" anchorCtr="1" lIns="45720" tIns="22860"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"><a:defRPr sz="900"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:defRPr></a:pPr>
            <a:r><a:rPr lang="en-US" sz="1400" b="1" i="1" u="sng"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>Total</a:t></a:r>
            <a:r><a:rPr lang="en-US" sz="1000"/><a:t> revenue</a:t></a:r>
            <a:br/>
            <a:r><a:t>second line</a:t></a:r>
          </a:p>
        </cdr:txBody>
      </cdr:sp>
      <cdr:clientData/>
    </cdr:absoluteAnchor>
    <cdr:oneCellAnchor>
      <cdr:from><cdr:x>0</cdr:x><cdr:y>0</cdr:y></cdr:from>
      <cdr:ext cx="4000000" cy="2000000"/>
      <cdr:sp>
        <cdr:nvSpPr><cdr:cNvPr id="8" name="Custom"/></cdr:nvSpPr>
        <cdr:spPr><a:custGeom><a:pathLst/></a:custGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></cdr:spPr>
        <cdr:txBody><a:bodyPr/><a:p><a:r><a:t>custom</a:t></a:r></a:p></cdr:txBody>
      </cdr:sp>
      <cdr:clientData/>
    </cdr:oneCellAnchor>
  </cdr:wsDr>`;
  const shapes = core.parseChartUserShapes(xml, {
    extentEmu: { cx: 4000000, cy: 2000000 },
    themePalette: { colorsByIndex: { 4: "#4f81bd" } },
  });

  assert.equal(shapes.length, 2);
  const note = shapes[0];
  assert.equal(note.name, "Note");
  assert.equal(note.description, "Note description");
  assert.equal(note.geometry, "chevron");
  assert.equal(note.geometryAdjustments.adj, 30000);
  assert.equal(note.fill.none, true);
  assert.equal(note.stroke.color, "#ff0000");
  assert.equal(note.stroke.dash, "dash");
  assert.equal(note.stroke.widthPx, 25400 / 9525);
  assert.equal(note.textBox.horizontalAlign, "center");
  assert.equal(note.textBox.verticalAlign, "bottom");
  assert.equal(note.textBox.insetPx.left, 45720 / 9525);
  assert.equal(note.textBox.insetPx.top, 22860 / 9525);
  assert.deepEqual(note.rect, {
    height: 0.5,
    left: 0.125,
    top: 0.125,
    width: 0.5,
  });
  assert.equal(note.paragraphs.length, 1);
  const firstParagraph = note.paragraphs[0];
  assert.equal(firstParagraph.align, "center");
  assert.deepEqual(
    firstParagraph.runs.map((run) => run.text),
    ["Total", " revenue", "\n", "second line"],
  );
  assert.equal(firstParagraph.runs[0].bold, true);
  assert.equal(firstParagraph.runs[0].italic, true);
  assert.equal(firstParagraph.runs[0].underline, true);
  assert.equal(firstParagraph.runs[0].color, "#112233");
  assert.equal(firstParagraph.runs[0].fontFamily, "Aptos");
  assert.equal(firstParagraph.runs[0].fontSizePt, 14);
  assert.equal(firstParagraph.runs[1].fontSizePt, 10);
  assert.equal(firstParagraph.runs[1].color, "#445566", "runs inherit the paragraph default colour");
  assert.equal(firstParagraph.runs[3].color, "#445566");
  assert.equal(firstParagraph.runs[3].fontSizePt, 9, "runs inherit the paragraph default size");

  const custom = shapes[1];
  assert.equal(custom.geometry, "custom");
  assert.equal(custom.fill.color, "#4f81bd", "theme scheme colours resolve");
  assert.deepEqual(custom.rect, { height: 1, left: 0, top: 0, width: 1 });

  assert.deepEqual(core.parseChartUserShapes("not xml"), []);
});

test("exporting and reopening keeps chart pictures and user shapes", () => {
  const { assets, chart, workbook } = loadFixture();
  try {
    const exported = core.mergeWorkbookImageAssets(
      workbook.saveXlsxBytes(),
      assets,
      assets.sheetOrigins,
    );
    const archive = unzipSync(exported);
    assert.ok(archive["xl/media/chart-logo.png"], "chart picture media survives export");
    assert.ok(archive[CHART_USER_SHAPES_PATH], "chartUserShapes part survives export");
    assert.match(strFromU8(archive["xl/charts/_rels/chart1.xml.rels"]), /chartUserShapes/);
    assert.match(strFromU8(archive["xl/charts/chart1.xml"]), /r:embed="rIdLogo"/);
    assert.match(
      strFromU8(archive["[Content_Types].xml"]),
      /chartUserShapes1\.xml/,
    );

    const reparsed = core.parseWorkbookImageAssets(exported);
    const reopenedWorkbook = wasm.Workbook.fromBytes(exported);
    try {
      const chartAssets = core.loadWorkbookChartAssets(
        reopenedWorkbook,
        reparsed,
        new Map([[0, 0], [1, 1]]),
      );
      const reopened = (chartAssets.chartsByWorkbookSheetIndex[0] ?? [])[0];
      assert.ok(reopened, "the chart is still a worksheet chart after reload");
      assert.equal(reopened.images.rIdLogo, chart.images.rIdLogo);
      assert.equal(reopened.series[0].fillImage, chart.series[0].fillImage);
      assert.equal(reopened.chartAreaFillImage, chart.chartAreaFillImage);
      assert.equal(reopened.plotAreaFillImage, chart.plotAreaFillImage);
      const drawing = reopened.drawings.rIdShapes;
      assert.ok(drawing, "chartUserShapes is reparsed after reload");
      assert.equal(drawing.images.rIdImg, chart.drawings.rIdShapes.images.rIdImg);
      assert.equal(drawing.shapes.length, 2);
      assert.equal(drawing.shapes[0].paragraphs[0].runs[0].text, CHART_SOURCE_NOTE_TEXT);
      assert.equal(drawing.shapes[1].fill.image, chart.drawings.rIdShapes.images.rIdImg);
      // The overlay is anchored in the chart's drawing space, so the fractions
      // follow whichever chart anchor the reopened package exposes.
      const extent = core.anchorToAbsoluteRect(
        reopened.anchor,
        reparsed.sheetStatesByWorkbookSheetIndex[reopened.workbookSheetIndex] ?? null,
      );
      assert.equal(drawing.shapes[0].rect.left, 340000 / extent.cx);
      assert.equal(drawing.shapes[0].rect.top, 1900000 / extent.cy);
      assert.equal(drawing.shapes[0].rect.width, (3200000 - 340000) / extent.cx);
      assert.equal(drawing.shapes[0].rect.height, (2320000 - 1900000) / extent.cy);
    } finally {
      reopenedWorkbook.free();
    }
  } finally {
    workbook.free();
  }
});
