import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

import { MemoChartSvg } from "../../packages/vue-xlsx/dist/chart.js";
import {
  CHART_SOURCE_NOTE_TEXT,
  buildChartMediaWorkbookBytes,
} from "../fixtures/chart-embedded-media.mjs";
import { importFromDemo, mount, textContent, walk } from "./vue-test-renderer.mjs";

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

const palette = {
  border: "#cbd5e1",
  mutedText: "#64748b",
  surface: "#ffffff",
  text: "#111827",
};
const rect = { height: 320, left: 0, top: 0, width: 480 };

const AREA_IMAGE = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAAB";
const PLOT_IMAGE = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAAC";
const SERIES_IMAGE = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAAD";
const POINT_IMAGE = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAAE";
const SHAPE_IMAGE = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAAF";

const SOURCE_NOTE = "source: Internal revenue dashboard";

function noteShape(overrides = {}) {
  return {
    fill: { color: "#ffffff" },
    geometry: "roundRect",
    geometryAdjustments: { adj: 12000 },
    name: "Source note",
    paragraphs: [{
      align: "left",
      runs: [{ bold: true, color: "#111827", fontSizePt: 10, text: SOURCE_NOTE }],
    }],
    rect: { height: 0.15, left: 0.05, top: 0.7, width: 0.5 },
    stroke: { color: "#1f2937", widthPx: 1.25 },
    textBox: { horizontalAlign: "left", insetPx: { bottom: 4.8, left: 9.6, right: 9.6, top: 4.8 }, verticalAlign: "middle" },
    ...overrides,
  };
}

function badgeShape(overrides = {}) {
  return {
    fill: { image: SHAPE_IMAGE },
    geometry: "rect",
    name: "Chart badge",
    paragraphs: [],
    rect: { height: 0.2, left: 0.8, top: 0.1, width: 0.2 },
    stroke: { hidden: true },
    ...overrides,
  };
}

function mediaChart(overrides = {}) {
  return {
    anchor: {
      from: { col: 0, colOffsetEmu: 0, row: 0, rowOffsetEmu: 0 },
      kind: "one-cell",
      sizeEmu: { cx: 4572000, cy: 2857500 },
    },
    axes: [],
    chartAreaFillImage: AREA_IMAGE,
    chartType: "ColumnClustered",
    id: "chart-media",
    legend: null,
    plotAreaFillImage: PLOT_IMAGE,
    series: [
      {
        categories: ["Q1", "Q2"],
        dataPointStyles: [{ fillImage: POINT_IMAGE, index: 1 }],
        dataPoints: [],
        fillImage: SERIES_IMAGE,
        id: "series-1",
        name: "Revenue",
        values: [10, 20],
      },
    ],
    sheetIndex: 0,
    workbookSheetIndex: 0,
    zIndex: 0,
    drawings: {
      rIdShapes: {
        images: { rIdImg: SHAPE_IMAGE },
        shapes: [noteShape(), badgeShape()],
        xml: '<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"/>',
      },
    },
    ...overrides,
  };
}

async function renderChart(chart) {
  const { root, warnings } = await mount(MemoChartSvg, { chart, palette, rect });
  assert.deepEqual(warnings, []);
  return root;
}

function nodesOfType(root, type) {
  return walk(root).filter((node) => node.type === type);
}

function patternIds(root) {
  return nodesOfType(root, "pattern").map((node) => node.props.id);
}

function fillNodes(root) {
  return walk(root)
    .filter((node) => node.props?.fill != null)
    .map((node) => ({ fill: node.props.fill, type: node.type }));
}

test("picture fills render as SVG patterns for the chart area, plot, series and points", async () => {
  const root = await renderChart(mediaChart());

  const defs = walk(root).find((node) => node.props?.["data-xlsx-chart-image-defs"] != null);
  assert.ok(defs, "the chart defines image-fill patterns");
  assert.deepEqual(patternIds(root), [
    "xlsx-chart-image-chart-media-chart-area",
    "xlsx-chart-image-chart-media-plot-area",
    "xlsx-chart-image-chart-media-series-0",
    "xlsx-chart-image-chart-media-series-0-point-1",
  ]);
  for (const pattern of nodesOfType(root, "pattern")) {
    const image = walk(pattern).find((node) => node.type === "image");
    assert.ok(image, `${pattern.props.id} carries an <image>`);
    assert.equal(image.props.width, "1");
    assert.equal(image.props.height, "1");
    assert.equal(image.props.preserveAspectRatio, "none");
    assert.match(image.props.href, /^data:image\/png;base64,/);
  }

  const fills = fillNodes(root);
  assert.ok(
    fills.some((entry) => entry.fill === "url(#xlsx-chart-image-chart-media-chart-area)"),
    "the chart area paints its picture fill",
  );
  assert.ok(
    fills.some((entry) => entry.fill === "url(#xlsx-chart-image-chart-media-plot-area)"),
    "the plot area paints its picture fill",
  );

  const bars = nodesOfType(root, "rect").filter((node) => node.props?.["data-xlsx-chart-series-index"] === "0");
  assert.equal(bars.length, 2);
  assert.equal(bars[0].props.fill, "url(#xlsx-chart-image-chart-media-series-0)");
  assert.equal(bars[1].props.fill, "url(#xlsx-chart-image-chart-media-series-0-point-1)");
});

test("chartUserShapes overlay renders text boxes and pictures above the plot", async () => {
  const root = await renderChart(mediaChart());

  const shapesGroup = walk(root).find((node) => node.props?.["data-xlsx-chart-user-shapes"] != null);
  assert.ok(shapesGroup, "user shapes paint as one overlay group");
  const shapes = walk(shapesGroup).filter((node) => node.props?.["data-xlsx-chart-shape"] != null);
  assert.deepEqual(shapes.map((node) => node.props["data-xlsx-chart-shape"]), ["roundRect", "rect"]);

  const noteText = walk(shapesGroup)
    .filter((node) => node.type === "text")
    .map((node) => textContent(node));
  assert.ok(noteText.some((text) => text.includes(SOURCE_NOTE)), "the source note text is painted");
  const noteGroup = shapes[0];
  const noteTextNode = walk(noteGroup).find((node) => node.type === "text");
  assert.match(noteTextNode.props["text-anchor"], /start/);
  // 320 × 0.7 + inset (4.8) + half of the 15 %-height box → inside the shape.
  assert.ok(Number(noteTextNode.props.y) > 320 * 0.7);
  assert.ok(Number(noteTextNode.props.y) < 320 * 0.85);

  const notePath = walk(noteGroup).find((node) => node.type === "path" && node.props?.fill != null);
  assert.match(notePath.props.d, /^M /);
  assert.equal(notePath.props.fill, "#ffffff");
  assert.equal(notePath.props.stroke, "#1f2937");

  const badgeGroup = shapes[1];
  const badgeImage = walk(badgeGroup).find((node) => node.type === "image");
  assert.ok(badgeImage, "the picture shape paints its image");
  assert.equal(badgeImage.props.href, SHAPE_IMAGE);
  assert.equal(badgeImage.props.x, 480 * 0.8);
  assert.equal(badgeImage.props.y, 320 * 0.1);
  assert.equal(badgeImage.props.width, 480 * 0.2);
  assert.equal(badgeImage.props.height, 320 * 0.2);
  const badgePath = walk(badgeGroup).find((node) => node.type === "path" && node.props?.fill != null);
  assert.equal(
    badgePath.props.d,
    `M ${480 * 0.8} ${320 * 0.1} H ${480} V ${320 * 0.3} H ${480 * 0.8} Z`,
  );
  assert.equal(badgePath.props.fill, "none", "the image paints the shape body");

  // The overlay belongs to the same SVG as the plot, after the plot elements.
  const svg = nodesOfType(root, "svg")[0];
  const children = svg.children.filter((child) => child.type !== "#text" && child.type !== "#comment");
  assert.ok(children.indexOf(shapesGroup) > children.findIndex((child) => child.type === "rect"));
});

test("invalid picture sources are never painted", async () => {
  const root = await renderChart(mediaChart({
    chartAreaFillImage: "javascript:alert(1)",
    plotAreaFillImage: "data:text/html;base64,PHNjcmlwdD4=",
    series: [{
      categories: ["Q1", "Q2"],
      dataPointStyles: [{ fillImage: "data:image/svg+xml;base64,PHN2Zy8+", index: 1 }],
      dataPoints: [],
      fillImage: "https://evil.example/logo.png",
      id: "series-1",
      name: "Revenue",
      values: [10, 20],
    }],
    drawings: {
      rIdShapes: {
        images: { rIdImg: "data:text/html;base64,PHNjcmlwdD4=" },
        shapes: [noteShape({ fill: { image: "data:text/html;base64,PHNjcmlwdD4=" } }), badgeShape({ fill: { image: "javascript:alert(1)" } })],
        xml: "<wsDr/>",
      },
    },
  }));

  assert.deepEqual(patternIds(root), [], "unsafe fill sources do not become patterns");
  assert.equal(nodesOfType(root, "image").length, 0, "unsafe shape sources do not become images");
  assert.ok(
    walk(root).some((node) => node.type === "text" && textContent(node).includes(SOURCE_NOTE)),
    "the text box still renders without its picture fill",
  );
  const bars = nodesOfType(root, "rect").filter((node) => node.props?.["data-xlsx-chart-series-index"] === "0");
  assert.equal(bars.length, 2);
  for (const bar of bars) {
    assert.doesNotMatch(String(bar.props.fill), /^url\(#/);
  }
});

test("charts without embedded media keep their previous SVG", async () => {
  const root = await renderChart(mediaChart({
    chartAreaFillImage: undefined,
    drawings: undefined,
    plotAreaFillImage: undefined,
    series: [{
      categories: ["Q1", "Q2"],
      dataPoints: [],
      id: "series-1",
      name: "Revenue",
      values: [10, 20],
    }],
  }));
  assert.equal(walk(root).some((node) => node.props?.["data-xlsx-chart-image-defs"] != null), false);
  assert.equal(walk(root).some((node) => node.props?.["data-xlsx-chart-user-shapes"] != null), false);
  const bars = nodesOfType(root, "rect").filter((node) => node.props?.["data-xlsx-chart-series-index"] === "0");
  assert.equal(bars.length, 2);
  for (const bar of bars) {
    assert.doesNotMatch(String(bar.props.fill), /^url\(#/);
    assert.match(String(bar.props.fill), /^#[0-9a-f]{6}$/i);
  }
});

test("a chartUserShapes picture with a shape clip path paints a clipped image", async () => {
  const root = await renderChart(mediaChart({
    drawings: {
      rIdShapes: {
        images: { rIdImg: SHAPE_IMAGE },
        shapes: [badgeShape()],
        xml: "<wsDr/>",
      },
    },
  }));
  const badgeImage = nodesOfType(root, "image").find((node) => node.props.href === SHAPE_IMAGE);
  assert.ok(badgeImage);
  assert.match(String(badgeImage.props["clip-path"]), /^url\(#xlsx-chart-shape-clip-chart-media-0\)$/);
  const clipPath = walk(root).find((node) => node.type === "clipPath");
  assert.equal(clipPath.props.id, "xlsx-chart-shape-clip-chart-media-0");
  assert.match(walk(clipPath).find((node) => node.type === "path").props.d, /^M /);
});

test("a real workbook with a logo fill and a source text box renders both", async () => {
  const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
  const core = await import(
    pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href,
  );
  const wasmBytes = readFileSync(
    requireFromXlsxCore.resolve("@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm"),
  );
  setWasmSource(
    wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
  );
  const wasm = await core.getSheetsWasmModule();
  const bytes = buildChartMediaWorkbookBytes();
  const workbook = wasm.Workbook.fromBytes(bytes);
  try {
    const assets = core.parseWorkbookImageAssets(bytes);
    const chartAssets = core.loadWorkbookChartAssets(
      workbook,
      assets,
      new Map([[0, 0], [1, 1]]),
    );
    const chart = (chartAssets.chartsByWorkbookSheetIndex[0] ?? [])[0];
    assert.ok(chart);

    const { root, warnings } = await mount(MemoChartSvg, { chart, palette, rect });
    assert.deepEqual(warnings, []);
    assert.ok(
      walk(root).some((node) => node.type === "text" && textContent(node).includes(CHART_SOURCE_NOTE_TEXT)),
      "the chartUserShapes note is painted from the package",
    );
    const images = nodesOfType(root, "image");
    assert.ok(images.length >= 2, "the series fill and the badge picture both paint an <image>");
    for (const image of images) {
      assert.match(String(image.props.href), /^data:image\/png;base64,/);
    }
    const patterns = nodesOfType(root, "pattern");
    assert.ok(patterns.length >= 4, "chart area, plot area, series and point patterns exist");
    const bars = nodesOfType(root, "rect").filter((node) => node.props?.["data-xlsx-chart-series-index"] === "0");
    assert.equal(bars.length, 4);
    assert.match(String(bars[0].props.fill), /^url\(#xlsx-chart-image-chart-0-0-series-0\)$/);
    assert.match(String(bars[1].props.fill), /^url\(#xlsx-chart-image-chart-0-0-series-0-point-1\)$/);
  } finally {
    workbook.free();
  }
});
