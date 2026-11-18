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
const RICH_TEXT_LAYOUT_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/rich-text-layout.ts");
const GRID_PAINT_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/grid-paint.ts");
const GRID_PAINT_SOURCE = readFileSync(GRID_PAINT_PATH, "utf8");
const RICH_TEXT_PLAIN = "Quarterly2";
const RICH_TEXT_RUNS_XML = [
  "<r><rPr><b/></rPr><t>Quarter</t></r>",
  '<r><rPr><color rgb="FFFF0000"/></rPr><t>ly</t></r>',
  '<r><rPr><vertAlign val="superscript"/></rPr><t>2</t></r>',
].join("");

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

/** `charts-images.xlsx` with A1 inline-rich and B1 shared-string-rich. */
function richTextWorkbookBuffer() {
  const bytes = new Uint8Array(
    readFileSync(new URL("../../apps/demo/public/samples/charts-images.xlsx", import.meta.url)),
  );
  const archive = unzipSync(bytes);
  let sheet = strFromU8(archive["xl/worksheets/sheet1.xml"]);
  sheet = sheet
    .replace(
      '<c r="A1" s="1" t="inlineStr"><is><t>Quarter</t></is></c>',
      `<c r="A1" s="1" t="inlineStr"><is>${RICH_TEXT_RUNS_XML}</is></c>`,
    )
    .replace(
      '<c r="B1" s="1" t="inlineStr"><is><t>Revenue</t></is></c>',
      '<c r="B1" s="1" t="s"><v>0</v></c>',
    );
  archive["xl/worksheets/sheet1.xml"] = strToU8(sheet);
  archive["xl/sharedStrings.xml"] = strToU8(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">'
      + `<si>${RICH_TEXT_RUNS_XML}</si></sst>`,
  );
  archive["xl/_rels/workbook.xml.rels"] = strToU8(
    strFromU8(archive["xl/_rels/workbook.xml.rels"]).replace(
      "</Relationships>",
      '<Relationship Id="rIdSharedStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
    ),
  );
  archive["[Content_Types].xml"] = strToU8(
    strFromU8(archive["[Content_Types].xml"]).replace(
      "</Types>",
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
    ),
  );
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

async function mountRichTextController(t, options = {}) {
  const mounted = await mountController({
    file: richTextWorkbookBuffer(),
    fileName: "rich-text.xlsx",
    ...options,
  });
  t.after(() => mounted.app.unmount());
  await waitFor(
    () => mounted.controller.sourceState === "ready" || mounted.controller.error,
    "XLSX controller ready state",
  );
  assert.ifError(mounted.controller.error);
  assert.equal(mounted.controller.sourceState, "ready");
  return mounted;
}

// ── Canvas harness ─────────────────────────────────────────────────────

function gridSource() {
  const source = readFileSync(GRID_PATH, "utf8");
  return {
    path: GRID_PATH,
    text: [...source.matchAll(/<script(?:\s+setup)?[^>]*>([\s\S]*?)<\/script>/gu)]
      .map((match) => match[1])
      .join("\n"),
  };
}

const GRID_SOURCE = gridSource();
const GRID_AST = ts.createSourceFile(
  GRID_SOURCE.path,
  GRID_SOURCE.text,
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

function transpile(source) {
  return ts.transpileModule(source, {
    compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2022 },
  }).outputText;
}

function dataUrl(source) {
  return `data:text/javascript;base64,${Buffer.from(transpile(source)).toString("base64")}`;
}

const richTextLayout = await import(dataUrl(readFileSync(RICH_TEXT_LAYOUT_PATH, "utf8")));
const drawTextModule = await import(dataUrl([
  "let __iconGeometry = null;",
  "function resolveIconGeometry() { return __iconGeometry; }",
  `import { wrapRichTextRuns, measureRichTextSegments, richTextBaselineOffset, richTextFontScale } from "${dataUrl(readFileSync(RICH_TEXT_LAYOUT_PATH, "utf8"))}";`,
  `import { resolveRotatedTextBounds, resolveRotatedTextCenter, resolveRotatedTextFitWidth, resolveShrinkToFitFontSize } from "${dataUrl(GRID_PAINT_SOURCE)}";`,
  "function resolveWorkbookColor(color) {",
  "  const value = color?.rgb;",
  "  if (!value) return null;",
  '  const hex = String(value).replace(/^#/, "");',
  '  return `#${(hex.length === 8 ? hex.slice(2) : hex).toLowerCase()}`;',
  "}",
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
].join("\n")));
const { drawCellText } = drawTextModule;
const { measureRichTextSegments, truncateRichTextRuns } = richTextLayout;

function createTextContext() {
  const calls = [];
  let font = "12px Arial";
  const ctx = {
    fillStyle: "",
    lineWidth: 1,
    strokeStyle: "",
    textAlign: "left",
    textBaseline: "alphabetic",
    get font() {
      return font;
    },
    set font(value) {
      font = value;
    },
    beginPath() {},
    clip() {},
    lineTo() {},
    moveTo() {},
    restore() {},
    save() {},
    stroke() {},
    rect() {},
    fillText(text, x, y) {
      calls.push({ fillStyle: this.fillStyle, font, text, x, y });
    },
    measureText(text) {
      // Advance proportional to the active font size, so a shrunken
      // superscript run really measures narrower than the base font.
      const size = Number.parseFloat(/(\d+(?:\.\d+)?)px/u.exec(font)?.[1] ?? "12");
      return { width: String(text).length * size * 0.6 };
    },
  };
  return { calls, ctx };
}

function baseTextState(overrides = {}) {
  return {
    conditional: null,
    runs: undefined,
    style: {
      alignment: "left",
      backgroundColor: "#ffffff",
      color: "#18181b",
      fontFamily: "Arial",
      fontSizePx: 12,
      fontStyle: "normal",
      fontWeight: "400",
      strike: false,
      underline: false,
      vertical: "bottom",
      wrap: false,
    },
    textColor: "#18181b",
    themePalette: { colorsByIndex: {} },
    value: RICH_TEXT_PLAIN,
    ...overrides,
  };
}

const RICH_RUNS = [
  { text: "Quarter", font: { bold: true } },
  { text: "ly", font: { color: { rgb: "FFFF0000" } } },
  { text: "2", font: { verticalAlign: "superscript" } },
];

// ── Snapshot / search ──────────────────────────────────────────────────

test("XLSX controller exposes rich-text runs and keeps the plain value", async (t) => {
  const { controller } = await mountRichTextController(t);

  const sheet = controller.activeSheet;
  assert.ok(sheet.richTextCells, "the sheet carries the parsed runs");
  assert.deepEqual(
    sheet.richTextCells.A1.map((run) => run.text),
    ["Quarter", "ly", "2"],
    "the inline rich cell keeps all three runs",
  );
  assert.deepEqual(
    sheet.richTextCells.B1.map((run) => run.text),
    ["Quarter", "ly", "2"],
    "the shared-string rich cell keeps its runs too",
  );

  const runs = controller.getCellRuns({ col: 0, row: 0 });
  assert.ok(runs, "the controller exposes the runs of a cell");
  assert.equal(runs[0].font.bold, true);
  assert.deepEqual(runs[1].font.color, { rgb: "#ff0000" });
  assert.equal(runs[2].font.verticalAlign, "superscript");
  assert.equal(controller.getCellDisplayValue({ col: 0, row: 0 }), RICH_TEXT_PLAIN);
  assert.equal(
    controller.getCellRuns({ col: 2, row: 0 }),
    undefined,
    "a plain cell has no runs",
  );

  assert.deepEqual(
    controller.findCells(RICH_TEXT_PLAIN).map((cell) => `${cell.row}:${cell.col}`),
    ["0:0", "0:1"],
    "search still matches the concatenated text",
  );
  assert.deepEqual(
    controller.findCells("Quarter").map((cell) => `${cell.row}:${cell.col}`),
    ["0:0", "0:1"],
    "a run prefix is searchable as part of the plain string",
  );
});

test("XLSX editing a rich cell drops its runs", async (t) => {
  const { controller } = await mountRichTextController(t);
  assert.ok(controller.getCellRuns({ col: 0, row: 0 }), "the cell starts rich");

  assert.equal(controller.commitCellInput({ col: 0, row: 0 }, "Edited"), true);
  await vue.nextTick();
  assert.equal(controller.getCellDisplayValue({ col: 0, row: 0 }), "Edited");
  assert.equal(
    controller.getCellRuns({ col: 0, row: 0 }),
    undefined,
    "runs must not paint over the new value",
  );
});

// ── Canvas renderer ────────────────────────────────────────────────────

test("XLSX click draws every rich run with its own font, colour and offset", () => {
  const { calls, ctx } = createTextContext();
  drawCellText(ctx, 0, 0, 160, 24, baseTextState({ runs: RICH_RUNS }), "#2563eb");

  const drawn = calls.filter((call) => call.text);
  assert.deepEqual(
    drawn.map((call) => call.text),
    ["Quarter", "ly", "2"],
    "each run reaches the canvas separately",
  );
  assert.match(drawn[0].font, /700/u, "the bold run keeps its weight");
  assert.equal(drawn[0].fillStyle, "#18181b", "the bold run inherits the cell colour");
  assert.equal(drawn[1].fillStyle, "#ff0000", "the red run paints in its own colour");
  assert.match(drawn[2].font, /8px/u, "the superscript run shrinks to 70% of 12px");
  assert.ok(
    drawn[2].y < drawn[0].y,
    "the superscript run sits above the text baseline",
  );
  assert.equal(drawn[0].y, drawn[1].y, "runs on the same line share the baseline");
});

test("XLSX cell text measures rich runs by their own widths, not the plain string", () => {
  const runWidth = measureRichTextSegments(RICH_RUNS, (text, font) => {
    const scale = font?.verticalAlign === "superscript" ? 0.7 : 1;
    return String(text).length * 12 * scale * 0.6;
  });
  const plainWidth = String(RICH_TEXT_PLAIN).length * 12 * 0.6;
  assert.equal(runWidth, 69.84, "run widths add up per run");
  assert.ok(runWidth < plainWidth, "the shrunken superscript makes the cell narrower");

  // 80px cell, 5px padding each side: the run sum fits, the plain string does not.
  const { calls, ctx } = createTextContext();
  drawCellText(ctx, 0, 0, 80, 24, baseTextState({ runs: RICH_RUNS }), "#2563eb");
  assert.equal(
    calls.some((call) => String(call.text).includes("\u2026")),
    false,
    "the run sum fits the cell, so nothing is truncated",
  );
  assert.deepEqual(calls.map((call) => call.text), ["Quarter", "ly", "2"]);

  const narrow = createTextContext();
  drawCellText(narrow.ctx, 0, 0, 70, 24, baseTextState({ runs: RICH_RUNS }), "#2563eb");
  const truncated = narrow.calls.map((call) => call.text).join("");
  assert.match(truncated, /\u2026/u, "a narrow cell truncates the laid-out runs");
  assert.ok(
    truncated.length < RICH_TEXT_PLAIN.length + 1,
    "the truncated line is shorter than the full text",
  );

  const segments = truncateRichTextRuns(RICH_RUNS, (text, font) => {
    const scale = font?.verticalAlign === "superscript" ? 0.7 : 1;
    return String(text).length * 12 * scale * 0.6;
  }, 70);
  assert.equal(segments.map((segment) => segment.text).join(""), RICH_TEXT_PLAIN);
});

// ── DOM renderer ───────────────────────────────────────────────────────

test("XLSX print preview renders one span per rich run", async (t) => {
  const { controller } = await mountRichTextController(t);
  assert.equal(controller.openPrintPreview(), true);

  const { app, root } = await mount(XlsxPrintPreview, { controller });
  t.after(() => app.unmount());

  const richCell = walk(root)
    .filter((node) => node.type === "td")
    .find((node) => textContent(node).trim() === RICH_TEXT_PLAIN);
  assert.ok(richCell, "the printed page keeps the concatenated cell text");

  const spans = walk(richCell).filter(
    (node) => node.type === "span" && !String(node.props?.class ?? "").includes("xlsx-print-page__rich-text"),
  );
  assert.deepEqual(
    spans.map((node) => textContent(node)),
    ["Quarter", "ly", "2"],
    "the DOM path emits one span per run",
  );
  assert.equal(spans[0].props.style.fontWeight, "700", "the bold run keeps its weight");
  assert.equal(spans[1].props.style.color, "#ff0000", "the red run keeps its colour");
  assert.equal(spans[2].props.style.verticalAlign, "super", "the superscript run is raised");
  assert.equal(
    Number(spans[2].props.style.lineHeight),
    richTextLayout.RICH_TEXT_VERTICAL_ALIGN_SCALE,
    "DOM and canvas share the superscript scale",
  );
  assert.equal(
    spans[0].props.style.verticalAlign,
    undefined,
    "a run without vertical alignment stays on the baseline",
  );
});

// ── Thumbnail renderer ─────────────────────────────────────────────────

function createThumbnailCanvas() {
  const calls = [];
  const context = {
    fillStyle: "",
    font: "",
    lineWidth: 1,
    strokeStyle: "",
    textAlign: "left",
    textBaseline: "alphabet",
    beginPath() {},
    clip() {},
    fillRect() {},
    lineTo() {},
    moveTo() {},
    rect() {},
    restore() {},
    save() {},
    setTransform() {},
    stroke() {},
    fillText(text, x, y) {
      calls.push({ fillStyle: this.fillStyle, font: this.font, text, x, y });
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

test("XLSX thumbnails paint rich runs with their own fonts and colours", async (t) => {
  const { controller } = await mountRichTextController(t);
  const api = useXlsxViewerThumbnails(() => controller);
  const { calls, canvas } = createThumbnailCanvas();

  assert.equal(api.paintThumbnail(0, canvas), true, "the thumbnail paints");

  // Rich cells of row 1: A1 (inline) and B1 (shared string). Other cells draw
  // plain text with the 8px base font or a header font.
  // Rich cells of row 1: A1 (inline) and B1 (shared string). Row 1 starts at the
  // 24px header with the usual 2px cell inset; the superscript sits above it.
  const baseline = 24 + 2;
  const superscriptBaseline = baseline + richTextLayout.richTextBaselineOffset(
    { verticalAlign: "superscript" },
    8,
  );
  const runCalls = calls.filter(
    (call) => ["Quarter", "ly", "2"].includes(call.text)
      && (call.y === baseline || call.y === superscriptBaseline),
  );
  assert.deepEqual(
    runCalls.map((call) => call.text),
    ["Quarter", "ly", "2", "Quarter", "ly", "2"],
    "the thumbnail draws every run of both rich cells separately",
  );
  assert.match(runCalls[0].font, /^bold 8px/u, "the bold run keeps its weight in the thumbnail");
  assert.equal(runCalls[0].fillStyle, "#333", "the bold run inherits the thumbnail text colour");
  assert.equal(runCalls[1].fillStyle, "#ff0000", "the red run keeps its colour in the thumbnail");
  assert.match(runCalls[2].font, /^5\.6px/u, "the superscript run shrinks below the 8px base");
  assert.ok(
    runCalls[2].y < runCalls[0].y,
    "the superscript run sits above the thumbnail baseline",
  );
  assert.deepEqual(
    runCalls.slice(3).map((call) => ({ fillStyle: call.fillStyle, font: call.font, text: call.text, y: call.y })),
    runCalls.slice(0, 3).map((call) => ({ fillStyle: call.fillStyle, font: call.font, text: call.text, y: call.y })),
    "the inline and shared-string rich cells render identically",
  );
});
