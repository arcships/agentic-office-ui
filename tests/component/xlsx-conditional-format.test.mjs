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

test("data bars honour border, axis and gradient settings", () => {
  const plain = rulesOfKind("dataBar", 4);
  const plainEnvironment = conditionalFormat.createConditionalEnvironment(
    plain,
    ruleValues(plain),
    resolveColor,
  );
  assert.equal(plainEnvironment.dataBar.border, false, "bars are unbordered unless the rule asks");
  assert.equal(plainEnvironment.dataBar.gradient, true, "Excel data bars default to a gradient");

  const styled = clone(plain);
  styled.axisColor = { rgb: "#aabbcc" };
  styled.border = true;
  styled.borderColor = { rgb: "#112233" };
  styled.gradient = false;
  styled.negativeBorderColor = { rgb: "#445566" };
  styled.negativeFillColor = { rgb: "#778899" };
  const environment = conditionalFormat.createConditionalEnvironment(
    styled,
    ruleValues(styled),
    resolveColor,
  );
  assert.equal(environment.dataBar.axisColor, "#aabbcc");
  assert.equal(environment.dataBar.border, true);
  assert.equal(environment.dataBar.borderColor, "#112233");
  assert.equal(environment.dataBar.gradient, false);
  assert.equal(environment.dataBar.negativeBorderColor, "#445566");
  assert.equal(environment.dataBar.negativeFillColor, "#778899");

  styled.negativeBarBorderColorSameAsPositive = true;
  const sameBorder = conditionalFormat.createConditionalEnvironment(
    styled,
    ruleValues(styled),
    resolveColor,
  );
  assert.equal(
    sameBorder.dataBar.negativeBorderColor,
    sameBorder.dataBar.borderColor,
    "negativeBarBorderColorSameAsPositive reuses the positive border colour",
  );

  assert.match(GRID_SOURCE, /if \(style\.gradient\)/, "the grid draws the gradient");
  assert.match(GRID_SOURCE, /if \(style\.border\)/, "the grid draws the bar border");
  assert.match(GRID_SOURCE, /style\.axisColor/, "the grid draws the zero axis");
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
    /ctx\.rect\((?:colX|x) \+ 2, (?:rowY|y) \+ 1, Math\.max\(0, textAreaWidth - 4\)/,
    "the text clip uses the width left over by the icon",
  );
  assert.match(
    GRID_SOURCE,
    /(?:colX|x) \+ textAreaWidth - padding/,
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

// Reads the corpus the way the grid does: display string in, number or `null`
// (empty) out; `undefined` stands for a cell the renderer cannot read yet.
function readCorpusCell(cell) {
  const text = String(worksheet.getFormattedValueAt(cell.row, cell.col)).trim();
  if (!text) return null;
  const value = Number(text.replace(/[,％%]/g, ""));
  return Number.isFinite(value) ? value : null;
}

function withFormulaBounds(rule, formulas) {
  const cloneRule = clone(rule);
  cloneRule.cfvos = formulas.map((value) => ({ type: "formula", value }));
  return cloneRule;
}

test("formula cfvos evaluate real workbook references", () => {
  const base = rulesOfKind("dataBar", 1);
  // B1 is outside the A1:A30 rule range: the threshold comes from another cell.
  const reference = withFormulaBounds(base, ["=$B$1", "=100"]);
  const environment = conditionalFormat.createConditionalEnvironment(
    reference,
    ruleValues(reference),
    resolveColor,
    readCorpusCell,
  );
  assert.deepEqual(
    { high: environment.bounds.high, low: environment.bounds.low },
    { high: 100, low: -2 },
    "$B$1 (-2) is the evaluated low bound, not the literal text or the range min",
  );
  assert.equal(environment.bounds.lowAutomatic, false, "an evaluated formula is a fixed bound");
  assertClose(
    conditionalFormat.resolveConditionalRatio(20, environment.bounds),
    22 / 102,
    "the data bar scales from the evaluated formula bound",
  );

  const relative = withFormulaBounds(base, ["=B5", "=A30"]);
  const relativeEnvironment = conditionalFormat.createConditionalEnvironment(
    relative,
    ruleValues(relative),
    resolveColor,
    readCorpusCell,
  );
  assert.deepEqual(
    { high: relativeEnvironment.bounds.high, low: relativeEnvironment.bounds.low },
    { high: 60, low: 2 },
    "relative references resolve against the same sheet",
  );

  const iconRule = withFormulaBounds(rulesOfKind("iconSet", 3), ["=B1", "=B2", "=B3", "=B4"]);
  iconRule.icons = [
    { iconId: 0, iconSet: "4Arrows" },
    { iconId: 1, iconSet: "4Arrows" },
    { iconId: 2, iconSet: "4Arrows" },
    { iconId: 3, iconSet: "4Arrows" },
  ];
  const iconEnvironment = conditionalFormat.createConditionalEnvironment(
    iconRule,
    ruleValues(iconRule),
    resolveColor,
    readCorpusCell,
  );
  assert.deepEqual(
    iconEnvironment.iconThresholds,
    [-2, -1, 0, 1],
    "icon-set cfvo formulas feed the bucket thresholds",
  );
  assert.deepEqual(
    [-3, -1, 0, 3].map((value) => conditionalFormat.resolveIconGeometry(iconEnvironment, value, 80, 24).bucket),
    [0, 1, 2, 3],
    "buckets follow the evaluated thresholds",
  );

  const scale = withFormulaBounds(rulesOfKind("colorScale", 2), ["=B1", "=B5", "=A30"]);
  const scaleEnvironment = conditionalFormat.createConditionalEnvironment(
    scale,
    ruleValues(scale),
    resolveColor,
    readCorpusCell,
  );
  assertClose(scaleEnvironment.colorScale.stops[0], 0, "B1 (-2) is the low stop");
  assertClose(scaleEnvironment.colorScale.stops[1], 4 / 62, "B5 (2) is the middle stop of -2…60");
  assertClose(scaleEnvironment.colorScale.stops[2], 1, "A30 (60) is the high stop");
  assert.equal(
    conditionalFormat.resolveColorScaleFill(scaleEnvironment, 2),
    "#ffeb84",
    "the middle colour sits on the evaluated formula stop",
  );

  assert.match(
    GRID_SOURCE,
    /createConditionalEnvironment\([\s\S]{0,400}readConditionalCellValue,/,
    "XlsxGrid hands the cfvo formula evaluator a cell reader",
  );
  assert.match(
    GRID_SOURCE,
    /requestConditionalFormulaCell\(sheet\.workbookSheetIndex, cell\)/,
    "Worker-backed grids hydrate formula references they have not received yet",
  );
});

test("formula cfvos support arithmetic, ranges and functions", () => {
  const base = rulesOfKind("dataBar", 1);
  const cases = [
    ["=(B5*2)+10", 14],
    ["=SUM(A1:A3)", 12],
    ["=SUM(A1:A3)/2", 6],
    ["=MAX($B$1:$B$5)", 2],
    ["=MIN(B1:B5)", -2],
    ["=AVERAGE(B1:B5)", 0],
    ["=MEDIAN(A1:A5)", 6],
    ["=COUNT(A1:A30)", 30],
    ["=IF(B5>0,50,10)", 50],
    ["=IF(B1>0,50,10)", 10],
    ["=ABS(B1)", 2],
    ["=ROUND(A1/3,2)", 0.67],
    ["=PERCENTILE(A1:A30,0.5)", 31],
    ["=A1^2", 4],
    ["=50%*100", 50],
  ];
  for (const [formula, expected] of cases) {
    assert.equal(
      conditionalFormat.evaluateConditionalFormula(formula, readCorpusCell),
      expected,
      `${formula} evaluates to ${expected}`,
    );
    const rule = withFormulaBounds(base, [formula, "=100"]);
    const environment = conditionalFormat.createConditionalEnvironment(
      rule,
      ruleValues(rule),
      resolveColor,
      readCorpusCell,
    );
    assertClose(
      environment.bounds.low,
      expected,
      `${formula} becomes the cfvo bound`,
    );
  }
});

test("formula cfvos fall back to the automatic bound when unresolvable", () => {
  const base = rulesOfKind("dataBar", 1);
  const stats = conditionalFormat.createConditionalRuleStats(ruleValues(base));

  assert.equal(
    conditionalFormat.evaluateConditionalFormula("=B1"),
    undefined,
    "a reference without a cell reader cannot be evaluated",
  );
  assert.equal(
    conditionalFormat.evaluateConditionalFormula("='Other Sheet'!A1", readCorpusCell),
    undefined,
    "cross-sheet references stay unsupported",
  );
  assert.equal(
    conditionalFormat.evaluateConditionalFormula("=B1", () => undefined),
    undefined,
    "an unavailable cell fails instead of reading as zero",
  );
  assert.equal(
    conditionalFormat.evaluateConditionalFormula("=B1", () => null),
    0,
    "an empty cell reads as zero like Excel",
  );
  assert.equal(
    conditionalFormat.evaluateConditionalFormula("=IF(A1)", readCorpusCell),
    undefined,
    "a malformed formula reports no value instead of throwing",
  );
  assert.equal(
    conditionalFormat.evaluateConditionalFormula("=MAX(A1:Z5000)", () => 1),
    undefined,
    "an over-budget range fails instead of evaluating as an empty aggregate",
  );

  const unavailable = withFormulaBounds(base, ["=B1", "=100"]);
  const environment = conditionalFormat.createConditionalEnvironment(
    unavailable,
    ruleValues(unavailable),
    resolveColor,
    () => undefined,
  );
  assert.equal(environment.bounds.low, stats.min, "an unavailable formula keeps the range min");
  assert.equal(environment.bounds.lowAutomatic, true);
});

test("parsed formula cfvos keep and evaluate their formula", async () => {
  const { unzipSync, zipSync } = requireFromVueXlsx("fflate");
  const files = unzipSync(corpusBytes);
  const sheetPath = "xl/worksheets/sheet1.xml";
  const sheetXml = Buffer.from(files[sheetPath]).toString("utf8");
  // A real workbook rule whose data-bar bound is another cell.
  const formulaSheet = sheetXml.replace(
    '<cfvo type="num" val="0"/><cfvo type="num" val="100"/>',
    '<cfvo type="formula" val="=$B$1"/><cfvo type="num" val="100"/>',
  );
  assert.notEqual(formulaSheet, sheetXml, "the corpus data bar cfvo was rewritten");
  files[sheetPath] = new TextEncoder().encode(formulaSheet);

  const assets = core.parseWorkbookStructureAssets(zipSync(files));
  const [dataBarRule] = assets.sheetStatesByWorkbookSheetIndex[0].conditionalFormatRules;
  assert.deepEqual(dataBarRule.cfvos[0], { type: "formula", value: "=$B$1" });

  const environment = conditionalFormat.createConditionalEnvironment(
    dataBarRule,
    ruleValues(dataBarRule),
    resolveColor,
    readCorpusCell,
  );
  assert.equal(environment.bounds.low, -2, "the parsed formula is evaluated against B1");
  assert.equal(environment.bounds.lowAutomatic, false);
  assertClose(
    conditionalFormat.resolveConditionalRatio(20, environment.bounds),
    22 / 102,
    "the parsed rule scales from the evaluated cell",
  );
});
