import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import test from "node:test";
import { pathToFileURL } from "node:url";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  walk,
  waitFor,
} from "./vue-test-renderer.mjs";

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

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const {
  XlsxPrintPreview,
  useXlsxViewerController,
  useXlsxViewerThumbnails,
} = await importFromDemo("@arcships/vue-xlsx");

const ROOT = path.resolve(import.meta.dirname, "../..");
const GRID_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxGrid.vue");
const GRID_SOURCE = readFileSync(GRID_PATH, "utf8");
const GRID_SCRIPT = [...GRID_SOURCE.matchAll(/<script(?:\s+setup)?[^>]*>([\s\S]*?)<\/script>/g)]
  .map((match) => match[1])
  .join("\n");
const GRID_AST = ts.createSourceFile(
  GRID_PATH,
  GRID_SCRIPT,
  ts.ScriptTarget.Latest,
  true,
  ts.ScriptKind.TS,
);

function gridFunctionText(name) {
  let found;
  function visit(node) {
    if (ts.isFunctionDeclaration(node) && node.name?.text === name) {
      found = node;
      return;
    }
    ts.forEachChild(node, visit);
  }
  visit(GRID_AST);
  assert.ok(found, `missing ${name} in XlsxGrid`);
  return found.getText(GRID_AST);
}

function transpileModuleSource(source) {
  return ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
}

function dataUrl(source) {
  return `data:text/javascript;base64,${Buffer.from(transpileModuleSource(source)).toString("base64")}`;
}

function loadTsModule(source) {
  return import(dataUrl(source));
}

const PAINT_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/grid-paint.ts");
const PAINT_SOURCE = readFileSync(PAINT_PATH, "utf8");
const RICH_TEXT_LAYOUT_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/rich-text-layout.ts");
const RICH_TEXT_LAYOUT_SOURCE = readFileSync(RICH_TEXT_LAYOUT_PATH, "utf8");

const paint = await loadTsModule(PAINT_SOURCE);
const { resolveSpreadsheetTextRotation } = paint;

const drawTextModule = await loadTsModule([
  "let __iconGeometry = null;",
  "function resolveIconGeometry() { return __iconGeometry; }",
  "export function setIconGeometry(value) { __iconGeometry = value; }",
  `import { wrapRichTextRuns, measureRichTextSegments, richTextBaselineOffset, richTextFontScale } from "${dataUrl(RICH_TEXT_LAYOUT_SOURCE)}";`,
  `import { resolveRotatedTextBounds, resolveRotatedTextCenter, resolveRotatedTextFitWidth, resolveShrinkToFitFontSize } from "${dataUrl(PAINT_SOURCE)}";`,
  "function resolveWorkbookColor(color) { return color?.rgb ? `#${String(color.rgb).slice(2)}` : null; }",
  gridFunctionText("truncateText"),
  gridFunctionText("wrapCellText"),
  gridFunctionText("baseCellFont"),
  gridFunctionText("richTextRunFont"),
  gridFunctionText("drawRichTextRuns"),
  gridFunctionText("richTextNaturalWidth"),
  gridFunctionText("drawRotatedPlainCellText"),
  gridFunctionText("drawRotatedRichTextRuns"),
  gridFunctionText("drawCellText"),
  "export { drawCellText, wrapCellText };",
].join("\n"));
const { drawCellText } = drawTextModule;

let wasmConfigured = false;
function configureWasm() {
  if (wasmConfigured) return;
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  );
  wasmConfigured = true;
}
configureWasm();

const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

/** `charts-images.xlsx` with 45°, 90°, 180° and shrink-to-fit header cells. */
function rotationWorkbookBuffer() {
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
  archive["xl/worksheets/sheet1.xml"] = strToU8(sheet);
  const patched = zipSync(archive);
  return patched.buffer.slice(patched.byteOffset, patched.byteOffset + patched.byteLength);
}

async function mountController(options = {}) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({ useWorker: false, ...options });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, controller };
}

async function mountRotationController(t, options = {}) {
  const mounted = await mountController({
    file: rotationWorkbookBuffer(),
    fileName: "text-rotation.xlsx",
    ...options,
  });
  t.after(() => mounted.app.unmount());
  await waitFor(
    () => mounted.controller.sourceState === "ready" || mounted.controller.error,
    "XLSX controller ready state",
  );
  assert.ifError(mounted.controller.error);
  return mounted;
}

function createTextContext() {
  const calls = [];
  const stack = [];
  const state = { angle: 0, tx: 0, ty: 0 };
  const ctx = {
    fillStyle: "",
    font: "",
    lineWidth: 1,
    strokeStyle: "",
    textAlign: "left",
    textBaseline: "alphabetic",
    beginPath() {},
    clip() {},
    lineTo() {},
    moveTo() {},
    rect(x, y, width, height) { calls.push({ type: "rect", x, y, width, height }); },
    fillText(text, x, y) {
      calls.push({
        angle: state.angle,
        font: this.font,
        text,
        tx: state.tx,
        ty: state.ty,
        type: "fillText",
        x,
        y,
      });
    },
    measureText(text) {
      const size = Number.parseFloat(/(\d+(?:\.\d+)?)px/u.exec(this.font)?.[1] ?? "12");
      return { width: String(text).length * size * 0.6 };
    },
    restore() {
      const previous = stack.pop();
      if (previous) Object.assign(state, previous);
    },
    rotate(angle) {
      state.angle += angle;
      calls.push({ angle, type: "rotate" });
    },
    save() { stack.push({ ...state }); },
    stroke() {},
    translate(x, y) {
      state.tx += x;
      state.ty += y;
      calls.push({ type: "translate", x, y });
    },
  };
  return { calls, ctx };
}

function baseTextState(overrides = {}) {
  return {
    conditional: null,
    style: {
      alignment: "left",
      backgroundColor: "#ffffff",
      color: "#18181b",
      fontFamily: "Arial",
      fontSizePx: 12,
      fontStyle: "normal",
      fontWeight: "400",
      rotation: null,
      shrinkToFit: false,
      strike: false,
      underline: false,
      vertical: "bottom",
      wrap: false,
    },
    textColor: "#18181b",
    value: "Header",
    ...overrides,
  };
}

test("canvas text rotates 45°, 90° and 180° inside the cell clip at 100% and 200%", () => {
  for (const fontSizePx of [12, 24]) {
    for (const [excelRotation, expected] of [[45, -Math.PI / 4], [90, -Math.PI / 2], [180, Math.PI / 2]]) {
      const { ctx, calls } = createTextContext();
      drawCellText(ctx, 0, 0, 80, 24, baseTextState({
        style: {
          ...baseTextState().style,
          fontSizePx,
          rotation: resolveSpreadsheetTextRotation(excelRotation),
        },
        value: "Rotated header",
      }), "#2563eb");

      const rotate = calls.find((call) => call.type === "rotate");
      assert.ok(rotate, `${excelRotation}° at ${fontSizePx}px must call rotate`);
      assert.ok(
        Math.abs(rotate.angle - expected) < 1e-9,
        `${excelRotation}° at ${fontSizePx}px rotates to ${expected} rad`,
      );
      const clip = calls.find((call) => call.type === "rect");
      assert.deepEqual(
        { height: clip.height, width: clip.width, x: clip.x, y: clip.y },
        { height: 22, width: 76, x: 2, y: 1 },
        "rotated text stays clipped to its own cell so neighbours are never painted over",
      );
      assert.ok(
        calls.some((call) => call.type === "fillText" && call.text === "Rotated header"),
        "the rotated value still paints in full and relies on the cell clip",
      );
    }
  }
});

test("shrink-to-fit binary-searches the canvas font and keeps the full value", () => {
  const { ctx, calls } = createTextContext();
  drawCellText(ctx, 0, 0, 80, 24, baseTextState({
    style: { ...baseTextState().style, fontSizePx: 16, shrinkToFit: true },
    value: "1234567890",
  }), "#2563eb");

  const fill = calls.find((call) => call.type === "fillText");
  assert.ok(fill, "the value paints");
  const fontPx = Number.parseFloat(/(\d+(?:\.\d+)?)px/u.exec(fill.font)?.[1] ?? "");
  assert.ok(fontPx >= 6.4 - 1e-9, `shrink-to-fit stops at 0.4× (got ${fontPx}px)`);
  assert.ok(fontPx < 16, `the font shrinks below the base size (got ${fontPx}px)`);
  assert.equal(fill.text, "1234567890", "fitted text is not truncated");
  assert.doesNotMatch(fill.text, /\u2026/u);

  const zoomed = createTextContext();
  drawCellText(zoomed.ctx, 0, 0, 80, 24, baseTextState({
    style: { ...baseTextState().style, fontSizePx: 32, shrinkToFit: true },
    value: "1234567890",
  }), "#2563eb");
  const zoomedFill = zoomed.calls.find((call) => call.type === "fillText");
  const zoomedPx = Number.parseFloat(/(\d+(?:\.\d+)?)px/u.exec(zoomedFill.font)?.[1] ?? "");
  assert.ok(zoomedPx >= 12.8, "at 200% the 0.4× floor scales with the zoomed base font");
  assert.ok(zoomedPx <= 16, `200% still fits the narrow column (got ${zoomedPx}px)`);
});

function createThumbnailCanvas() {
  const calls = [];
  const stack = [];
  const state = { angle: 0, tx: 0, ty: 0 };
  const context = {
    fillStyle: "",
    font: "",
    lineWidth: 1,
    strokeStyle: "",
    textAlign: "left",
    textBaseline: "top",
    beginPath() {},
    clip() {},
    fillRect() {},
    lineTo() {},
    moveTo() {},
    rect(x, y, width, height) { calls.push({ type: "rect", x, y, width, height }); },
    restore() {
      const previous = stack.pop();
      if (previous) Object.assign(state, previous);
    },
    rotate(angle) {
      state.angle += angle;
      calls.push({ angle, type: "rotate" });
    },
    save() { stack.push({ ...state }); },
    setTransform() {},
    stroke() {},
    translate(x, y) {
      state.tx += x;
      state.ty += y;
      calls.push({ type: "translate", x, y });
    },
    fillText(text, x, y) {
      calls.push({
        angle: state.angle,
        font: this.font,
        text,
        tx: state.tx,
        ty: state.ty,
        type: "fillText",
        x,
        y,
      });
    },
    measureText(text) {
      const size = Number.parseFloat(/(\d+(?:\.\d+)?)px/u.exec(this.font)?.[1] ?? "8");
      return { width: String(text).length * size * 0.6 };
    },
  };
  return {
    calls,
    canvas: {
      getContext: () => context,
      height: 0,
      style: {},
      width: 0,
    },
  };
}

test("XLSX thumbnails rotate and shrink the same header cells", async (t) => {
  const { controller } = await mountRotationController(t);
  const api = useXlsxViewerThumbnails(() => controller);
  const { calls, canvas } = createThumbnailCanvas();

  assert.equal(api.paintThumbnail(0, canvas), true, "the thumbnail paints");
  const rotations = calls.filter((call) => call.type === "rotate").map((call) => call.angle);
  assert.equal(rotations.length, 3, "three header cells rotate");
  assert.ok(Math.abs(rotations[0] - -Math.PI / 4) < 1e-9);
  assert.ok(Math.abs(rotations[1] - -Math.PI / 2) < 1e-9);
  assert.ok(Math.abs(rotations[2] - Math.PI / 2) < 1e-9);

  const clips = calls.filter((call) => call.type === "rect");
  assert.deepEqual(
    clips.map((clip) => ({ height: clip.height, width: clip.width })),
    [
      { height: 24, width: 80 },
      { height: 24, width: 80 },
      { height: 24, width: 80 },
    ],
    "each rotated thumbnail cell clips to its own rectangle",
  );

  const shrink = calls.find((call) => call.type === "fillText" && call.text === "Shrink this long value");
  assert.ok(shrink, "the shrink-to-fit value paints in full");
  const shrinkFont = Number.parseFloat(/(\d+(?:\.\d+)?)px/u.exec(shrink.font)?.[1] ?? "");
  assert.ok(shrinkFont < 8, `the thumbnail shrinks below the 8px base (got ${shrinkFont}px)`);
});

test("XLSX print preview rotates and shrinks the same header cells", async (t) => {
  const { controller } = await mountRotationController(t);
  assert.equal(controller.openPrintPreview(), true);

  const { app, root } = await mount(XlsxPrintPreview, { controller });
  t.after(() => app.unmount());

  const cells = walk(root).filter((node) => node.type === "td");
  const findCell = (text) => cells.find((cell) => textContent(cell).trim() === text);

  const header45 = findCell("Rotated 45 header");
  assert.ok(header45, "the 45° cell prints");
  assert.equal(header45.props.style.position, "relative", "rotated cells anchor the absolute text span");
  const span45 = walk(header45).find((node) => node.type === "span" && textContent(node) === "Rotated 45 header");
  assert.match(span45.props.style.transform, /rotate\(-45deg\)/u);

  const header90 = findCell("Rotated 90 header");
  const span90 = walk(header90).find((node) => node.type === "span" && textContent(node) === "Rotated 90 header");
  assert.match(span90.props.style.transform, /rotate\(-90deg\)/u);

  const header180 = findCell("Rotated 180 header");
  const span180 = walk(header180).find((node) => node.type === "span" && textContent(node) === "Rotated 180 header");
  assert.match(span180.props.style.transform, /rotate\(90deg\)/u);

  const shrink = findCell("Shrink this long value");
  assert.ok(shrink, "the shrink-to-fit cell prints");
  const shrinkSpan = walk(shrink).find(
    (node) => node.type === "span" && textContent(node) === "Shrink this long value",
  );
  const fittedPx = Number.parseFloat(shrinkSpan.props.style.fontSize);
  const basePx = Number.parseFloat(shrink.props.style.fontSize);
  assert.ok(fittedPx < basePx, "the DOM path sets a smaller font size for shrink-to-fit");
  assert.ok(fittedPx >= basePx * 0.4 - 1e-9, "the DOM path keeps the 0.4× floor");
});

test("XLSX grid resolution keeps rotation when the worker snapshot lacks alignment", () => {
  assert.match(GRID_SOURCE, /resolveCellTextLayout\(/);
  assert.match(GRID_SOURCE, /sheet\.cellTextLayouts\?\.\[cellAddressToA1\(cell\)\]/);
  assert.match(GRID_SOURCE, /resolveRotatedTextFitWidth\(/);
  assert.match(GRID_SOURCE, /resolveShrinkToFitFontSize\(/);
  assert.match(GRID_SOURCE, /drawRotatedPlainCellText\(/);
  assert.match(GRID_SOURCE, /drawRotatedRichTextRuns\(/);
  assert.doesNotMatch(GRID_SOURCE, /textRotation === 255\) return -90/);
});
