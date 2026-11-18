import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";

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
const PAINT_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/grid-paint.ts");
const PAINT_SOURCE = readFileSync(PAINT_PATH, "utf8");
const CONDITIONAL_PATH = path.join(ROOT, "packages/vue-xlsx/src/render/conditional-format.ts");
const CONDITIONAL_SOURCE = readFileSync(CONDITIONAL_PATH, "utf8");

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

const RICH_TEXT_LAYOUT_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/rich-text-layout.ts");
const RICH_TEXT_LAYOUT_SOURCE = readFileSync(RICH_TEXT_LAYOUT_PATH, "utf8");

const paintModule = await loadTsModule(PAINT_SOURCE);
const conditionalModule = await loadTsModule(CONDITIONAL_SOURCE);
const richTextModule = await loadTsModule(RICH_TEXT_LAYOUT_SOURCE);
// `drawCellText` is the shared renderer used by both `drawCellContent` and the
// overflow pass. Transpile it with a stub icon resolver and a colour stub so the
// canvas calls can be asserted without a real Canvas implementation.
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
const { drawCellText, setIconGeometry, wrapCellText } = drawTextModule;
const { truncateRichTextRuns, wrapRichTextRuns } = richTextModule;

const {
  MAX_CANVAS_AREA_PX,
  MAX_CANVAS_DIMENSION_PX,
  intersectsViewport,
  readDevicePixelRatio,
  resolveAxisResizeSize,
  resolveCanvasPixelRatio,
  resolveOverflowTextWidth,
} = paintModule;
const {
  createConditionalEnvironment,
  resolveIconBucket,
  resolveIconGeometry,
} = conditionalModule;

test("XLSX canvas backing stores stay inside browser limits at any DPR", () => {
  assert.equal(resolveCanvasPixelRatio(800, 600, 2), 2);
  assert.equal(resolveCanvasPixelRatio(800, 600, 1), 1);
  assert.equal(resolveCanvasPixelRatio(Number.NaN, 600, Number.NaN), 1);

  const ratio = resolveCanvasPixelRatio(20_000, 10_000, 4);
  assert.ok(ratio < 1, "a huge canvas lowers the effective ratio instead of failing");
  assert.ok(20_000 * ratio <= MAX_CANVAS_DIMENSION_PX, "width stays under the canvas edge limit");
  assert.ok(10_000 * ratio <= MAX_CANVAS_DIMENSION_PX, "height stays under the canvas edge limit");
  assert.ok(
    20_000 * 10_000 * ratio * ratio <= MAX_CANVAS_AREA_PX * 1.000001,
    "backing-store area stays under the per-canvas limit",
  );

  globalThis.window = { devicePixelRatio: 3 };
  assert.equal(readDevicePixelRatio(), 3, "DPR is read live, not frozen at module load");
  delete globalThis.window;
  assert.equal(readDevicePixelRatio(), 1);
});

test("merged-region viewport intersection keeps clips whose anchor scrolled away", () => {
  const viewport = { bottom: 400, left: 100, right: 700, top: 100 };
  assert.equal(
    intersectsViewport({ height: 50, left: -200, top: 120, width: 400 }, viewport),
    true,
    "a merge starting left of the viewport still overlaps it",
  );
  assert.equal(
    intersectsViewport({ height: 400, left: 200, top: -200, width: 100 }, viewport),
    true,
    "a merge starting above the viewport still overlaps it",
  );
  assert.equal(
    intersectsViewport({ height: 50, left: -200, top: 120, width: 200 }, viewport),
    false,
    "a merge entirely left of the viewport stays hidden",
  );
  assert.equal(
    intersectsViewport({ height: 50, left: 702, top: 120, width: 100 }, viewport),
    false,
  );
});

test("Excel text overflow uses neighbour widths only up to the first occupied cell", () => {
  assert.equal(resolveOverflowTextWidth(80, []), 80);
  assert.equal(resolveOverflowTextWidth(80, [80, 120]), 280);
  assert.equal(resolveOverflowTextWidth(80, [Number.NaN, -10, 40]), 120);
});

function createTextContext() {
  const calls = [];
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
    restore() {},
    save() {},
    stroke() {},
    rect(x, y, width, height) { calls.push({ type: "rect", x, y, width, height }); },
    fillText(text, x, y) { calls.push({ type: "fillText", text, x, y }); },
    measureText(text) { return { width: String(text).length * 6 }; },
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
      strike: false,
      underline: false,
      vertical: "bottom",
      wrap: false,
    },
    textColor: "#18181b",
    value: "Spills across empty cells",
    ...overrides,
  };
}

test("cell text honours hard line breaks from Alt+Enter", () => {
  const { ctx, calls } = createTextContext();
  assert.deepEqual(
    wrapCellText(ctx, "first\nsecond", 70, 3),
    ["first", "second"],
    "wrap text breaks at the hard newline and keeps every line",
  );
  assert.deepEqual(
    wrapCellText(ctx, "first\nsecond", 70, 1),
    ["first"],
    "without wrap text only the first hard line is visible",
  );
  assert.deepEqual(
    wrapCellText(ctx, "abcdefghijklmnop\nsecond", 30, 3),
    ["abcde", "fghij", "k"],
    "soft wrapping applies inside each hard line and the line cap still wins",
  );
  assert.deepEqual(
    wrapCellText(ctx, "a\n", 70, 3),
    ["a", ""],
    "a trailing hard break keeps its empty line",
  );

  // The renderer draws the hard lines one below the other when wrap text is on.
  let state = baseTextState({ style: { ...baseTextState().style, wrap: true }, value: "first\nsecond" });
  drawCellText(ctx, 0, 0, 80, 40, state, "#2563eb");
  assert.deepEqual(
    calls.filter((call) => call.type === "fillText").map((call) => call.text),
    ["first", "second"],
    "both hard lines reach the canvas",
  );

  // Wrap text off is Excel's first-line-only display.
  const plain = createTextContext();
  state = baseTextState({ value: "first\nsecond" });
  drawCellText(plain.ctx, 0, 0, 80, 24, state, "#2563eb");
  assert.deepEqual(
    plain.calls.filter((call) => call.type === "fillText").map((call) => call.text),
    ["first"],
    "a cell without wrap text shows only the first line",
  );
});

test("drawCellText spills text across empty neighbours and still clips to the cell", () => {
  let { ctx, calls } = createTextContext();
  const value = "Spills across empty cells";
  drawCellText(ctx, 0, 0, 80, 24, baseTextState({ overflowWidth: 240, value }), "#2563eb");
  const spillClip = calls.find((call) => call.type === "rect");
  assert.equal(spillClip.width, 236, "the clip follows the overflow width, not the cell width");
  assert.equal(
    calls.find((call) => call.type === "fillText")?.text,
    value,
    "the label is not truncated while it fits the spill width",
  );

  ({ ctx, calls } = createTextContext());
  drawCellText(ctx, 0, 0, 80, 24, baseTextState({ value }), "#2563eb");
  assert.equal(calls.find((call) => call.type === "rect").width, 76);
  assert.match(
    calls.find((call) => call.type === "fillText")?.text ?? "",
    /\u2026/,
    "without an empty neighbour the text still truncates inside its own cell",
  );
});

test("icon sets honour showValue and reserve text width for the glyph", () => {
  setIconGeometry({ bucket: 4, color: "#16a34a", glyph: "\u25b2", iconX: 66, size: 12, textWidth: 62 });

  let { ctx, calls } = createTextContext();
  drawCellText(ctx, 0, 0, 80, 24, baseTextState({
    conditional: { environment: {}, kind: "iconSet", showValue: false, value: 100 },
  }), "#2563eb");
  assert.deepEqual(
    calls.filter((call) => call.type === "fillText").map((call) => call.text),
    ["\u25b2"],
    "showValue:false draws only the icon",
  );

  ({ ctx, calls } = createTextContext());
  drawCellText(ctx, 0, 0, 80, 24, baseTextState({
    conditional: { environment: {}, kind: "iconSet", showValue: true, value: 100 },
  }), "#2563eb");
  assert.equal(calls.filter((call) => call.type === "fillText").length, 2, "the value and icon both draw");
  assert.equal(
    calls.find((call) => call.type === "rect").width,
    58,
    "value clipping stops before the icon column",
  );
});

test("row / column resize divides by the same zoom basis used for painting", () => {
  assert.equal(resolveAxisResizeSize(80, 10, 100, 32), 90);
  assert.equal(
    resolveAxisResizeSize(80, 10, 25, 32),
    120,
    "at 25 % zoom a 10 px drag is 40 sheet pixels, matching the rendered width",
  );
  assert.equal(resolveAxisResizeSize(80, 10, 10, 32), 180);
  assert.equal(resolveAxisResizeSize(10, -100, 100, 32), 32);
});

const FIVE_ARROW_RULE = {
  cfvos: [
    { type: "percent", value: 0 },
    { type: "percent", value: 20 },
    { type: "percent", value: 40 },
    { type: "percent", value: 60 },
    { type: "percent", value: 80 },
  ],
  icons: Array.from({ length: 5 }, (_, iconId) => ({ iconId, iconSet: "5Arrows" })),
  kind: "iconSet",
  priority: 1,
  ranges: [{ end: { col: 0, row: 0 }, start: { col: 0, row: 0 } }],
  reverse: false,
};

function iconEnvironment(rule, values = [0, 25, 50, 75, 100]) {
  return createConditionalEnvironment(rule, values, () => undefined);
}

test("conditional-format icon buckets follow the rule thresholds", () => {
  const environment = iconEnvironment(FIVE_ARROW_RULE);
  assert.equal(resolveIconBucket(environment, -1), 0);
  assert.equal(resolveIconBucket(environment, 0), 0);
  assert.equal(resolveIconBucket(environment, 19), 0);
  assert.equal(resolveIconBucket(environment, 20), 1);
  assert.equal(resolveIconBucket(environment, 79), 3);
  assert.equal(resolveIconBucket(environment, 80), 4);
  assert.equal(resolveIconBucket(environment, 100), 4);

  const reversed = iconEnvironment({ ...FIVE_ARROW_RULE, reverse: true });
  assert.equal(resolveIconBucket(reversed, 100), 0, "reverse flips the icon order");
});

test("conditional-format icon geometry reserves space so icons never overlap values", () => {
  const environment = iconEnvironment(FIVE_ARROW_RULE);
  const geometry = resolveIconGeometry(environment, 100, 120, 24);
  assert.ok(geometry);
  assert.equal(geometry.bucket, 4);
  assert.ok(geometry.iconX > 0);
  assert.ok(
    geometry.textWidth < 120,
    "right-aligned values must stop before the icon column",
  );
  assert.ok(geometry.textWidth <= geometry.iconX);
});

test("conditional-format icon sets honour percentiles and the selected scheme", () => {
  const rule = {
    cfvos: [
      { type: "min" },
      { type: "percentile", value: 50 },
      { type: "percentile", value: 80 },
    ],
    icons: Array.from({ length: 3 }, (_, iconId) => ({ iconId, iconSet: "3TrafficLights1" })),
    kind: "iconSet",
    priority: 1,
    ranges: [{ end: { col: 0, row: 0 }, start: { col: 0, row: 0 } }],
  };
  const values = Array.from({ length: 100 }, (_, index) => index + 1);
  const environment = iconEnvironment(rule, values);
  assert.equal(resolveIconBucket(environment, 1), 0);
  assert.equal(resolveIconBucket(environment, 50), 0);
  assert.equal(resolveIconBucket(environment, 81), 2);
  assert.equal(environment.icons[2].glyph, "\u25cf", "traffic lights use filled circles");

  const arrows = iconEnvironment(
    { ...rule, icons: rule.icons.map((icon) => ({ ...icon, iconSet: "3Arrows" })) },
    values,
  );
  assert.equal(arrows.icons[2].glyph, "\u25b2", "arrow sets use directional glyphs");

  const gray = iconEnvironment(
    { ...rule, icons: rule.icons.map((icon) => ({ ...icon, iconSet: "3ArrowsGray" })) },
    values,
  );
  assert.notEqual(gray.icons[2].color, arrows.icons[2].color, "gray icon sets have their own palette");
});

test("XlsxGrid wires merged clipping, overflow, icon layout, DPR and zoom into paint", () => {
  assert.match(GRID_SOURCE, /for \(const region of paintedMergeRegions\)/);
  assert.match(GRID_SOURCE, /intersectsViewport\(/);
  assert.match(GRID_SOURCE, /overflowTextTasks/);
  assert.match(GRID_SOURCE, /drawCellText\(ctx, task\.x, task\.y, task\.width, task\.height, task\.state, "#2563eb"\)/);
  assert.match(GRID_SOURCE, /resolveOverflowTextWidth\(/);
  assert.match(
    GRID_SOURCE,
    /style\.wrap \|\| style\.rotation != null \|\| style\.shrinkToFit \|\| style\.alignment !== trailingAlignment/,
  );
  assert.match(GRID_SOURCE, /conditionalHidesValue/);

  assert.match(GRID_SOURCE, /resolveIconGeometry\(/);
  assert.match(GRID_SOURCE, /conditionalIcon\?\.textWidth/);
  assert.match(GRID_SOURCE, /drawCellContent\(/);
  assert.match(GRID_SOURCE, /function drawCellText\(/);
  assert.doesNotMatch(GRID_SOURCE, /conditional\.ratio < \.34/);

  assert.match(GRID_SOURCE, /resolveCanvasPixelRatio\(/);
  assert.match(GRID_SOURCE, /readDevicePixelRatio\(/);
  assert.match(GRID_SOURCE, /Math\.min\(MAX_CANVAS_DIMENSION_PX, Math\.ceil\(safeWidth \* pixelRatio\)\)/);
  assert.match(GRID_SOURCE, /ctx\.setTransform\(pixelRatio, 0, 0, pixelRatio, 0, 0\)/);
  assert.match(GRID_SOURCE, /matchMedia\(`\(resolution: \$\{readDevicePixelRatio\(\)\}dppx\)`\)/);
  assert.match(GRID_SOURCE, /devicePixelRatioRevision\.value \+= 1/);

  assert.match(GRID_SOURCE, /resolveAxisResizeSize\(/);
  assert.doesNotMatch(GRID_SOURCE, /Math\.max\(0\.5, zoomScale\.value \/ 100\)/);
});
