import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import test, { after } from "node:test";
import { pathToFileURL } from "node:url";

// Regression coverage for issue #61: XLSX conditional formatting must follow rule
// priority, cfvo bounds (num/percent/percentile/formula), data-bar geometry and
// colours (min/maxLength, negative bars, zero axis) and the real icon set.
// The workbook is the shared review-corpus fixture: data bars with fixed numeric
// bounds, a colour scale with a percentile middle, a 4-icon set and negative data.

const ROOT = path.resolve(import.meta.dirname, "../..");
const CORPUS = path.join(ROOT, "test-data/corpus/xlsx-conditional-formats.xlsx");
const MODULE_PATH = path.join(ROOT, "packages/vue-xlsx/src/render/conditional-format.ts");
const GRID_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxGrid.vue");
const GRID_SOURCE = readFileSync(GRID_PATH, "utf8");

const requireFromXlsxCore = createRequire(path.join(ROOT, "packages/xlsx-core/package.json"));
const requireFromVueXlsx = createRequire(path.join(ROOT, "packages/vue-xlsx/package.json"));
const ts = requireFromXlsxCore("typescript");

const { DOMParser, Node, XMLSerializer } = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= DOMParser;
globalThis.Node ??= Node;
globalThis.XMLSerializer ??= XMLSerializer;

const core = await import(
  pathToFileURL(requireFromVueXlsx.resolve("@arcships/xlsx-core")).href
);

async function loadConditionalFormatModule() {
  const source = readFileSync(MODULE_PATH, "utf8");
  const output = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
  return import(`data:text/javascript;base64,${Buffer.from(output).toString("base64")}`);
}

const conditionalFormat = await loadConditionalFormatModule();

const corpusBytes = new Uint8Array(readFileSync(CORPUS));
const structure = core.parseWorkbookStructureAssets(corpusBytes);
const rules = structure.sheetStatesByWorkbookSheetIndex[0].conditionalFormatRules;
const resolveColor = (color) => core.resolveWorkbookColor(color) ?? undefined;

const wasmPath = requireFromVueXlsx.resolve(
  "@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm",
);
const wasmBytes = readFileSync(wasmPath);
core.setWasmSource(
  wasmBytes.buffer.slice(
    wasmBytes.byteOffset,
    wasmBytes.byteOffset + wasmBytes.byteLength,
  ),
);
const wasm = await core.getSheetsWasmModule();
const workbook = wasm.Workbook.fromBytes(corpusBytes);
const worksheet = workbook.getSheet(0);
after(() => workbook.free());

function ruleValues(rule) {
  const values = [];
  for (const range of rule.ranges) {
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      for (let col = range.start.col; col <= range.end.col; col += 1) {
        const text = String(worksheet.getFormattedValueAt(row, col)).trim();
        if (!text) continue;
        const value = Number(text.replace(/[,％%]/g, ""));
        if (Number.isFinite(value)) values.push(value);
      }
    }
  }
  return values;
}

function rulesOfKind(kind, priority) {
  return rules.find((rule) => rule.kind === kind && rule.priority === priority);
}

function clone(rule) {
  return structuredClone(rule);
}

function assertClose(actual, expected, message, tolerance = 1e-9) {
  assert.ok(
    Math.abs(actual - expected) <= tolerance,
    `${message}: expected ${expected}, received ${actual}`,
  );
}

test("corpus exposes conditional rules with their priority and cfvos", () => {
  assert.deepEqual(
    rules.map((rule) => [rule.kind, rule.priority]),
    [
      ["dataBar", 1],
      ["colorScale", 2],
      ["iconSet", 3],
      ["dataBar", 4],
    ],
    "test-data/corpus/xlsx-conditional-formats.xlsx keeps four rules with distinct priorities",
  );
  assert.deepEqual(
    rules[0].cfvos.map((cfvo) => [cfvo.type, cfvo.value]),
    [["num", 0], ["num", 100]],
    "data bars keep their fixed numeric bounds",
  );
  assert.deepEqual(
    rules[1].cfvos.map((cfvo) => [cfvo.type, cfvo.value]),
    [["min", undefined], ["percentile", 50], ["max", undefined]],
    "colour scales keep the percentile middle",
  );
  assert.deepEqual(
    rules[2].cfvos.map((cfvo) => cfvo.value),
    [0, 25, 50, 75],
    "icon sets keep all four thresholds",
  );
  assert.deepEqual(
    rules[2].icons.map((icon) => [icon.iconSet, icon.iconId]),
    [["4TrafficLights", 0], ["4TrafficLights", 1], ["4TrafficLights", 2], ["4TrafficLights", 3]],
    "icon ids survive the parse",
  );
  assert.deepEqual(
    rules[3].cfvos.map((cfvo) => cfvo.value),
    [-3, 3],
    "the negative data bar keeps its signed bounds",
  );

  const aValues = ruleValues(rules[0]);
  assert.deepEqual([aValues[0], aValues[aValues.length - 1]], [2, 60]);
  const bValues = ruleValues(rules[3]);
  assert.deepEqual([Math.min(...bValues), Math.max(...bValues)], [-3, 3]);
});

test("overlapping rules resolve in priority order, not document order", () => {
  assert.deepEqual(
    conditionalFormat.conditionalRulesForCell(rules, { row: 0, col: 0 }).map((rule) => rule.priority),
    [1, 2],
    "the data bar (priority 1) wins over the colour scale (priority 2)",
  );
  assert.deepEqual(
    conditionalFormat.conditionalRulesForCell(rules, { row: 0, col: 1 }).map((rule) => rule.priority),
    [3, 4],
    "the icon set (priority 3) wins over the negative data bar (priority 4)",
  );
  assert.deepEqual(
    conditionalFormat.sortConditionalRules([...rules].reverse()).map((rule) => rule.priority),
    [1, 2, 3, 4],
    "sorting ignores the source order of the rule list",
  );
  assert.deepEqual(
    conditionalFormat.conditionalRulesForCell(rules, { row: 200, col: 9 }),
    [],
    "cells outside every range have no rule",
  );
  const tieA = { ...clone(rulesOfKind("dataBar", 1)), priority: 7 };
  const tieB = { ...clone(rulesOfKind("colorScale", 2)), priority: 7 };
  assert.deepEqual(
    conditionalFormat.sortConditionalRules([tieA, tieB]).map((rule) => rule.kind),
    ["dataBar", "colorScale"],
    "equal priorities keep document order",
  );

  assert.match(
    GRID_SOURCE,
    /conditionalRulesForCell\(cell\)\[0\]/,
    "XlsxGrid draws the winning rule only",
  );
  assert.match(
    GRID_SOURCE,
    /if \(!valueText\.trim\(\)\) return null;/,
    "blank cells stay unformatted",
  );
  assert.doesNotMatch(
    GRID_SOURCE,
    /\.ratio\s*<\s*\.34/,
    "XlsxGrid no longer buckets icons by a hardcoded ratio",
  );
});

test("data-bar bounds come from cfvos instead of the range min/max", () => {
  const rule = rulesOfKind("dataBar", 1);
  const values = ruleValues(rule);
  const environment = conditionalFormat.createConditionalEnvironment(rule, values, resolveColor);

  assert.deepEqual(
    { high: environment.bounds.high, low: environment.bounds.low },
    { high: 100, low: 0 },
    "0…100 cfvos define the scale, not the 2…60 data range",
  );
  assertClose(conditionalFormat.resolveConditionalRatio(20, environment.bounds), .2, "ratio for 20");
  assertClose(conditionalFormat.resolveConditionalRatio(40, environment.bounds), .4, "ratio for 40");
  assertClose(conditionalFormat.resolveConditionalRatio(60, environment.bounds), .6, "ratio for 60");
  assert.equal(conditionalFormat.resolveConditionalRatio(-10, environment.bounds), 0, "values below the low bound clamp");
  assert.equal(conditionalFormat.resolveConditionalRatio(200, environment.bounds), 1, "values above the high bound clamp");

  const bar20 = conditionalFormat.resolveDataBarGeometry(environment, 20, 100);
  const bar40 = conditionalFormat.resolveDataBarGeometry(environment, 40, 100);
  const bar120 = conditionalFormat.resolveDataBarGeometry(environment, 120, 100);
  assertClose(bar20.barWidth, 18, "20 on a 0…100 scale over a 100px bar with maxLength 90");
  assertClose(bar40.barWidth, 36, "40 is twice the 20 bar");
  assertClose(bar40.barWidth, bar20.barWidth * 2, "bar length follows the absolute scale");
  assertClose(bar120.barWidth, 90, "values above the scale stop at maxLength");
  assert.equal(bar40.showAxis, false, "an all-positive scale has no zero axis");
});

test("data bars honour minLength/maxLength", () => {
  const fixed = clone(rulesOfKind("dataBar", 1));
  fixed.minLength = 20;
  fixed.maxLength = 50;
  const fixedEnvironment = conditionalFormat.createConditionalEnvironment(
    fixed,
    ruleValues(fixed),
    resolveColor,
  );
  assertClose(
    conditionalFormat.resolveDataBarGeometry(fixedEnvironment, 0, 100).barWidth,
    0,
    "a value at the fixed minimum draws no bar",
  );
  assertClose(
    conditionalFormat.resolveDataBarGeometry(fixedEnvironment, 10, 100).barWidth,
    20,
    "a short bar is lifted to minLength",
  );
  assertClose(
    conditionalFormat.resolveDataBarGeometry(fixedEnvironment, 100, 100).barWidth,
    50,
    "a full bar stops at maxLength",
  );

  const automatic = clone(rulesOfKind("dataBar", 1));
  automatic.cfvos = [{ type: "min" }, { type: "max" }];
  const automaticEnvironment = conditionalFormat.createConditionalEnvironment(
    automatic,
    ruleValues(automatic),
    resolveColor,
  );
  assertClose(
    conditionalFormat.resolveDataBarGeometry(automaticEnvironment, 2, 100).barWidth,
    10,
    "the automatic minimum keeps the default minLength",
  );
  assertClose(
    conditionalFormat.resolveDataBarGeometry(automaticEnvironment, 60, 100).barWidth,
    90,
    "the automatic maximum reaches the default maxLength",
  );
});

test("negative data bars grow left from the zero axis with their own colour", () => {
  const rule = clone(rulesOfKind("dataBar", 4));
  const values = ruleValues(rule);
  const plainEnvironment = conditionalFormat.createConditionalEnvironment(rule, values, resolveColor);
  assert.equal(
    plainEnvironment.dataBar.negativeFillColor,
    "#dc2626",
    "negative bars default to a negative colour instead of reusing the positive one",
  );

  rule.negativeFillColor = { rgb: "#1d4ed8" };
  rule.border = true;
  const environment = conditionalFormat.createConditionalEnvironment(rule, values, resolveColor);
  assert.equal(environment.dataBar.negativeFillColor, "#1d4ed8");

  const negative = conditionalFormat.resolveDataBarGeometry(environment, -1, 100);
  const positive = conditionalFormat.resolveDataBarGeometry(environment, 1, 100);
  assertClose(negative.axisX, 50, "the zero axis sits in the middle of a -3…3 scale");
  assert.equal(negative.negative, true);
  assert.equal(negative.showAxis, true);
  assertClose(negative.barX + negative.barWidth, negative.axisX, "negative bars end at the axis");
  assertClose(negative.barWidth, 30, "-1 is a third of the negative span");
  assert.equal(positive.negative, false);
  assertClose(positive.barX, positive.axisX, "positive bars start at the axis");
  assertClose(positive.barWidth, 30, "+1 is a third of the positive span");

  assert.match(
    GRID_SOURCE,
    /geometry\.negative \? style\.negativeFillColor : style\.color/,
    "XlsxGrid paints negative bars with negativeFillColor",
  );
});

test("icon sets use their cfvo thresholds and the parsed icon ids", () => {
  const corpusRule = rulesOfKind("iconSet", 3);
  const bValues = ruleValues(corpusRule);
  const corpusEnvironment = conditionalFormat.createConditionalEnvironment(
    corpusRule,
    bValues,
    resolveColor,
  );
  assert.deepEqual(corpusEnvironment.iconThresholds, [0, 25, 50, 75]);
  assert.equal(corpusEnvironment.icons.length, 4, "4TrafficLights keeps four icons");
  assert.deepEqual(
    [-3, -1, 0, 3].map((value) => conditionalFormat.resolveIconBucket(corpusEnvironment, value)),
    [0, 0, 0, 0],
    "the corpus thresholds start at 0 and jump to 25, so -3…3 stay in the lowest bucket "
      + "(a thirds-of-the-range mapping would have promoted 0 and 3 to bucket 1)",
  );
  assert.notEqual(
    corpusEnvironment.icons[0].color,
    corpusEnvironment.icons[1].color,
    "each parsed icon id maps to its own traffic light",
  );

  const arrows = clone(corpusRule);
  arrows.cfvos = [
    { type: "percent", value: 0 },
    { type: "percent", value: 25 },
    { type: "percent", value: 50 },
    { type: "percent", value: 75 },
  ];
  arrows.icons = [
    { iconId: 0, iconSet: "4Arrows" },
    { iconId: 1, iconSet: "4Arrows" },
    { iconId: 2, iconSet: "4Arrows" },
    { iconId: 3, iconSet: "4Arrows" },
  ];
  const arrowEnvironment = conditionalFormat.createConditionalEnvironment(
    arrows,
    bValues,
    resolveColor,
  );
  assert.deepEqual(
    [-3, -1, 0, 3].map((value) => conditionalFormat.resolveIconGeometry(arrowEnvironment, value, 80, 24).bucket),
    [0, 1, 2, 3],
    "percent thresholds split the -3…3 range into the four buckets",
  );
  assert.deepEqual(
    [-3, -1, 0, 3].map((value) => conditionalFormat.resolveIconGeometry(arrowEnvironment, value, 80, 24).glyph),
    ["▼", "◢", "◥", "▲"],
    "the glyph comes from icons[].iconSet/iconId",
  );

  const reversed = clone(arrows);
  reversed.reverse = true;
  const reversedEnvironment = conditionalFormat.createConditionalEnvironment(
    reversed,
    bValues,
    resolveColor,
  );
  assert.deepEqual(
    [-3, 3].map((value) => conditionalFormat.resolveIconGeometry(reversedEnvironment, value, 80, 24).glyph),
    ["▲", "▼"],
    "reverse flips the threshold mapping, not just the glyph table",
  );
});

test("icon sets reserve space so the glyph never overlaps the value", () => {
  const rule = rulesOfKind("iconSet", 3);
  const environment = conditionalFormat.createConditionalEnvironment(
    rule,
    ruleValues(rule),
    resolveColor,
  );
  const layout = conditionalFormat.resolveIconGeometry(environment, -3, 80, 24);
  assert.equal(layout.size, 18);
  assert.equal(layout.iconX, 60);
  assert.equal(layout.textWidth, 58);
  assert.ok(
    layout.textWidth + layout.size + 2 <= 80,
    "text and icon together fit inside the cell",
  );

  assert.match(GRID_SOURCE, /resolveIconGeometry\(conditional\.environment/);
  assert.match(
    GRID_SOURCE,
    /ctx\.rect\(colX \+ 2, rowY \+ 1, Math\.max\(0, textAreaWidth - 4\)/,
    "the text clip uses the width left over by the icon",
  );
  assert.match(
    GRID_SOURCE,
    /colX \+ textAreaWidth - padding/,
    "right-aligned values stop before the icon",
  );
});

test("colour scales interpolate between their cfvo stops", () => {
  const rule = rulesOfKind("colorScale", 2);
  const values = ruleValues(rule);
  const environment = conditionalFormat.createConditionalEnvironment(rule, values, resolveColor);
  assert.deepEqual(
    environment.colorScale.colors,
    ["#f8696b", "#ffeb84", "#63be7b"],
  );
  assert.deepEqual(environment.colorScale.stops, [0, .5, 1]);
  assert.equal(
    conditionalFormat.resolveColorScaleFill(environment, 31),
    "#ffeb84",
    "the percentile 50 stop sits at the middle colour",
  );

  // Skewed data proves the stop follows the percentile, not 50% of the range.
  const skewed = clone(rule);
  skewed.cfvos = [{ type: "min" }, { type: "percentile", value: 25 }, { type: "max" }];
  const skewedEnvironment = conditionalFormat.createConditionalEnvironment(
    skewed,
    [0, 10, 20, 30],
    resolveColor,
  );
  assert.deepEqual(skewedEnvironment.colorScale.stops, [0, .25, 1]);
  assert.equal(
    conditionalFormat.resolveColorScaleFill(skewedEnvironment, 7.5),
    "#ffeb84",
    "7.5 is the 25th percentile of 0…30",
  );
  assert.notEqual(
    conditionalFormat.resolveColorScaleFill(skewedEnvironment, 3.75),
    "#ffeb84",
    "the midpoint of the range is no longer hardcoded to the middle colour",
  );
});

test("formula cfvos use their literal or fall back to the automatic bound", () => {
  const base = clone(rulesOfKind("dataBar", 1));
  const literal = clone(base);
  literal.cfvos = [{ type: "formula", value: "=20" }, { type: "num", value: 40 }];
  const literalEnvironment = conditionalFormat.createConditionalEnvironment(
    literal,
    ruleValues(literal),
    resolveColor,
  );
  assert.deepEqual(
    { high: literalEnvironment.bounds.high, low: literalEnvironment.bounds.low },
    { high: 40, low: 20 },
    "a literal formula cfvo is a fixed bound",
  );

  const reference = clone(base);
  reference.cfvos = [{ type: "formula", value: "$B$1" }, { type: "num", value: 40 }];
  const referenceEnvironment = conditionalFormat.createConditionalEnvironment(
    reference,
    ruleValues(reference),
    resolveColor,
  );
  assert.equal(referenceEnvironment.bounds.low, 2, "unresolvable formulas fall back to the range min");
  assert.equal(referenceEnvironment.bounds.lowAutomatic, true);
});

test("parsed formula cfvos keep the formula text", async () => {
  const { unzipSync, zipSync } = requireFromVueXlsx("fflate");
  const files = unzipSync(corpusBytes);
  const sheetPath = "xl/worksheets/sheet1.xml";
  const sheetXml = Buffer.from(files[sheetPath]).toString("utf8");
  const formulaSheet = sheetXml.replace(
    '<cfvo type="num" val="0"/><cfvo type="num" val="100"/>',
    '<cfvo type="formula" val="=25"/><cfvo type="num" val="100"/>',
  );
  assert.notEqual(formulaSheet, sheetXml, "the corpus data bar cfvo was rewritten");
  files[sheetPath] = new TextEncoder().encode(formulaSheet);

  const assets = core.parseWorkbookStructureAssets(zipSync(files));
  const [dataBarRule] = assets.sheetStatesByWorkbookSheetIndex[0].conditionalFormatRules;
  assert.deepEqual(dataBarRule.cfvos[0], { type: "formula", value: "=25" });
  assert.equal(
    formulaSheet.includes('type="formula"'),
    true,
    "the rewritten workbook still declares the formula cfvo",
  );
});
