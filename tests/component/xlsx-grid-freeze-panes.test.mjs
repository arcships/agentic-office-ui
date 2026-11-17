import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";

const ROOT = path.resolve(import.meta.dirname, "../..");
const GRID_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxGrid.vue");
const PANE_GEOMETRY_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/grid-pane-geometry.ts");
const GRID_SOURCE = readFileSync(GRID_PATH, "utf8");
const SCRIPT_SOURCE = [...GRID_SOURCE.matchAll(/<script(?:\s+setup)?[^>]*>([\s\S]*?)<\/script>/g)]
  .map((match) => match[1])
  .join("\n");
const AST = ts.createSourceFile(
  GRID_PATH,
  SCRIPT_SOURCE,
  ts.ScriptTarget.Latest,
  true,
  ts.ScriptKind.TS,
);

async function loadPaneGeometry() {
  // The module is dependency-free apart from type imports, so it can be
  // transpiled and executed directly.
  const output = ts.transpileModule(readFileSync(PANE_GEOMETRY_PATH, "utf8"), {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
  return import(`data:text/javascript;base64,${Buffer.from(output).toString("base64")}`);
}

const geometry = await loadPaneGeometry();

function descendants(node, predicate) {
  const matches = [];
  function visit(current) {
    if (predicate(current)) matches.push(current);
    ts.forEachChild(current, visit);
  }
  visit(node);
  return matches;
}

function functionNode(name) {
  return descendants(
    AST,
    (node) => ts.isFunctionDeclaration(node) && node.name?.text === name,
  )[0];
}

function declaration(name) {
  return descendants(
    AST,
    (node) =>
      ts.isVariableDeclaration(node) &&
      ts.isIdentifier(node.name) &&
      node.name.text === name,
  )[0];
}

function functionText(name) {
  const node = functionNode(name);
  assert.ok(node, `XlsxGrid is missing ${name}`);
  return node.getText(AST);
}

function calledNames(node) {
  return descendants(node, ts.isCallExpression).map((call) => call.expression.getText(AST));
}

// The issue's repro fixture: row 1 and column A frozen, 80 × 24 cells, and a
// viewport scrolled right and down.
const COL_WIDTH = 80;
const ROW_HEIGHT = 24;
const SCROLL_LEFT = 200;
const SCROLL_TOP = 100;

test("freeze-pane point mapping keeps frozen strips in content space", () => {
  // Frozen column 0..80, frozen row 0..24.
  assert.deepEqual(
    geometry.resolvePaneContentPoint(30, 10, SCROLL_LEFT, SCROLL_TOP, COL_WIDTH, ROW_HEIGHT),
    { pane: "frozen-corner", contentX: 30, contentY: 10 },
  );
  assert.deepEqual(
    geometry.resolvePaneContentPoint(120, 10, SCROLL_LEFT, SCROLL_TOP, COL_WIDTH, ROW_HEIGHT),
    { pane: "frozen-rows", contentX: 120 + SCROLL_LEFT, contentY: 10 },
  );
  assert.deepEqual(
    geometry.resolvePaneContentPoint(30, 60, SCROLL_LEFT, SCROLL_TOP, COL_WIDTH, ROW_HEIGHT),
    { pane: "frozen-cols", contentX: 30, contentY: 60 + SCROLL_TOP },
  );
  assert.deepEqual(
    geometry.resolvePaneContentPoint(120, 60, SCROLL_LEFT, SCROLL_TOP, COL_WIDTH, ROW_HEIGHT),
    { pane: "main", contentX: 120 + SCROLL_LEFT, contentY: 60 + SCROLL_TOP },
  );
  // Without a freeze the whole body is the scrolling pane.
  assert.deepEqual(
    geometry.resolvePaneContentPoint(120, 60, SCROLL_LEFT, SCROLL_TOP, 0, 0),
    { pane: "main", contentX: 120 + SCROLL_LEFT, contentY: 60 + SCROLL_TOP },
  );
});

test("freeze-pane axes subtract scroll only outside the frozen strip", () => {
  assert.equal(geometry.paneViewOffset(30, SCROLL_LEFT, true), 30);
  assert.equal(geometry.paneViewOffset(120, SCROLL_LEFT, false), 120 - SCROLL_LEFT);
  assert.equal(geometry.paneAxisContentOffset(78, SCROLL_LEFT, COL_WIDTH), 78);
  assert.equal(geometry.paneAxisContentOffset(90, SCROLL_LEFT, COL_WIDTH), 90 + SCROLL_LEFT);
  // A zero extent must never claim points that scrolled past the origin.
  assert.equal(geometry.paneAxisContentOffset(0, SCROLL_LEFT, 0), SCROLL_LEFT);
});

test("freeze-pane view points place overlays on the pinned cell", () => {
  const base = {
    frozenCols: 1,
    frozenRows: 1,
    scrollLeft: SCROLL_LEFT,
    scrollTop: SCROLL_TOP,
  };
  // Frozen corner cell A1 keeps its content position.
  assert.deepEqual(
    geometry.resolvePaneViewPoint({ ...base, contentX: 0, contentY: 0, displayCol: 0, displayRow: 0 }),
    { x: 0, y: 0 },
  );
  // Column A stays pinned while its row scrolls.
  assert.deepEqual(
    geometry.resolvePaneViewPoint({ ...base, contentX: 0, contentY: 96, displayCol: 0, displayRow: 5 }),
    { x: 0, y: 96 - SCROLL_TOP },
  );
  // A cell in the scrolling pane follows both scroll offsets.
  assert.deepEqual(
    geometry.resolvePaneViewPoint({ ...base, contentX: 400, contentY: 96, displayCol: 5, displayRow: 5 }),
    { x: 400 - SCROLL_LEFT, y: 96 - SCROLL_TOP },
  );
});

test("XLSX hit testing resolves the freeze pane before mapping to a cell", () => {
  const resolve = functionText("resolvePanePoint");
  assert.match(resolve, /resolvePaneContentPoint\(/, "hit tests must share the pane mapping");
  assert.match(resolve, /container\.scrollLeft/);
  assert.match(resolve, /container\.scrollTop/);
  assert.match(resolve, /frozen\.width/);
  assert.match(resolve, /frozen\.height/);

  for (const name of ["hitTestCell", "hitTestSheetPoint"]) {
    const node = functionNode(name);
    const text = node.getText(AST);
    assert.ok(calledNames(node).includes("resolvePanePoint"), `${name} must go through resolvePanePoint`);
    assert.doesNotMatch(text, /container\.scroll(Left|Top)/, `${name} must not add scroll unconditionally`);
  }

  const axis = functionNode("hitTestAxis");
  const axisText = axis.getText(AST);
  const axisCalls = descendants(
    axis,
    (node) => ts.isCallExpression(node) && node.expression.getText(AST) === "paneAxisContentOffset",
  );
  assert.equal(axisCalls.length, 2, "both header strips must resolve their pane");
  assert.match(axisCalls[0].getText(AST), /container\.scrollLeft/);
  assert.match(axisCalls[1].getText(AST), /container\.scrollTop/);
  assert.doesNotMatch(axisText, /localX \+ container\.scrollLeft/);
  assert.doesNotMatch(axisText, /localY \+ container\.scrollTop/);
});

test("XLSX headers paint a pinned strip and a scrolling strip", () => {
  const columns = functionText("paintColHeaders");
  assert.match(
    columns,
    /paneViewOffset\(getColOffsetSum\(displayIndex\), scrollLeft\.value, displayIndex < frozen\.cols\)/,
    "a frozen column label must ignore scrollLeft",
  );
  assert.doesNotMatch(columns, /translate\(-scrollLeft/, "the strip can no longer shift as a whole");
  assert.match(columns, /ctx\.rect\(frozen\.width, 0, Math\.max\(0, cw - frozen\.width\), HEADER_HEIGHT\)/);
  assert.match(columns, /for \(let c = 0; c < frozen\.cols && c < widths\.length; c\+\+\) paintColumn\(c\);?/);
  assert.match(columns, /Math\.max\(colRange\.start, frozen\.cols\)/);
  assert.equal(
    descendants(
      functionNode("paintColHeaders"),
      (node) => ts.isCallExpression(node) && node.expression.getText(AST) === "ctx.fillText",
    ).length,
    1,
    "one label routine serves both strips",
  );

  const rows = functionText("paintRowHeaders");
  assert.match(
    rows,
    /paneViewOffset\(getRowOffsetSum\(displayIndex\), scrollTop\.value, displayIndex < frozen\.rows\)/,
    "a frozen row label must ignore scrollTop",
  );
  assert.doesNotMatch(rows, /translate\(0, -scrollTop/);
  assert.match(rows, /ctx\.rect\(0, frozen\.height, ROW_HEADER_WIDTH, Math\.max\(0, rh - frozen\.height\)\)/);
  assert.match(rows, /for \(let r = 0; r < frozen\.rows && r < heights\.length; r\+\+\) paintRow\(r\);?/);
  assert.match(rows, /Math\.max\(rowRange\.start, frozen\.rows\)/);
});

test("XLSX frozen panes reuse the full cell renderer", () => {
  const content = `${functionText("drawCellContent")}\n${functionText("drawCellText")}`;
  assert.match(
    functionText("drawCellContent"),
    /drawCellText\(ctx, x, y, width, height, state/,
    "drawCellContent must delegate text so the overflow pass can share it",
  );
  for (const [fragment, label] of [
    [/conditional\?\.kind === "dataBar"/, "data bars"],
    [/drawSparkline\(ctx, sparkline/, "sparklines"],
    [/#2563eb/, "hyperlink colour"],
    [/style\.vertical === "top"/, "vertical alignment"],
    [/style\.wrap/, "text wrapping"],
    [/#dc2626/, "comment marker"],
    [/state\.gridLines/, "gridline outlines"],
    [/drawCellBorders\(ctx, style/, "borders"],
  ]) {
    assert.match(content, fragment, `drawCellContent must keep ${label}`);
  }

  const body = functionText("paintBody");
  assert.match(body, /drawCellContent\(ctx, colX, rowY, colW, rowHResolved/, "the scrolling pass uses the shared renderer");

  const frozen = functionText("paintFrozenPanes");
  assert.match(frozen, /drawCellContent\(ctx, x, y, width, height/, "the frozen pass uses the shared renderer");
  assert.match(frozen, /paneViewOffset\(getColOffsetSum\(displayCol\), scrollLeft\.value, frozenX\)/);
  assert.match(frozen, /paneViewOffset\(getRowOffsetSum\(displayRow\), scrollTop\.value, frozenY\)/);
  assert.match(
    frozen,
    /gridLines: sheet\.showGridLines !== false/,
    "the frozen strips draw their own gridlines because the scrolled ones land elsewhere",
  );
  assert.doesNotMatch(
    frozen,
    /ctx\.fillText|textBaseline|drawSparkline\(|conditional\?\.|#2563eb|#dc2626|drawCellBorders\(ctx, style/,
    "the frozen pass must not carry its own reduced cell renderer",
  );
});

test("XLSX selection outlines are painted per freeze pane", () => {
  const body = functionText("paintBody");
  assert.ok(
    body.indexOf("paintFrozenPanes(") < body.indexOf("paintSelectionBorders("),
    "outlines must be drawn after the frozen cells",
  );

  const borders = functionText("paintSelectionBorders");
  assert.match(borders, /ctx\.clip\(\)/);
  assert.match(
    borders,
    /translate\(pane\.scrollX \? -scrollLeft\.value : 0, pane\.scrollY \? -scrollTop\.value : 0\)/,
    "each pane translates only the axes that scroll",
  );
  assert.doesNotMatch(
    functionText("paintFrozenPanes"),
    /activeCell/,
    "the frozen pass no longer owns a separate active-cell outline",
  );
});

test("XLSX inline editor and comment popover follow the frozen cell", () => {
  for (const name of ["editInputStyle", "hoveredComment"]) {
    const node = declaration(name);
    assert.ok(node, `XlsxGrid is missing ${name}`);
    const text = node.getText(AST);
    assert.match(text, /resolvePaneViewPoint\(\{/, `${name} must resolve the pane`);
    assert.match(text, /frozenCols: frozen\.cols/);
    assert.match(text, /frozenRows: frozen\.rows/);
    assert.doesNotMatch(text, /-\s*scrollLeft\.value/, `${name} must not subtract scroll unconditionally`);
    assert.doesNotMatch(text, /-\s*scrollTop\.value/, `${name} must not subtract scroll unconditionally`);
  }
});

test("XLSX scrollToCell leaves pinned cells where they are", () => {
  const source = functionText("scrollToCell");
  assert.match(source, /frozenPaneMetrics\(\)/);
  assert.match(source, /displayCell\.col >= frozen\.cols/, "a pinned column never scrolls into view");
  assert.match(source, /displayCell\.row >= frozen\.rows/, "a pinned row never scrolls into view");
  assert.match(source, /bodyWidth - frozen\.width/, "the scrolling pane is narrower than the body");
  assert.match(source, /bodyHeight - frozen\.height/);
});

test("XLSX frozen extent reports the pinned display columns and rows", () => {
  const metrics = functionText("frozenPaneMetrics");
  assert.match(metrics, /cols \+= 1/);
  assert.match(metrics, /rows \+= 1/);
  assert.match(metrics, /return \{ width, height, cols, rows \}/);
  // The header strips and hit tests all key off the same extent.
  assert.match(functionText("paintColHeaders"), /frozenPaneMetrics\(\)/);
  assert.match(functionText("paintRowHeaders"), /frozenPaneMetrics\(\)/);
  assert.match(functionText("hitTestAxis"), /frozenPaneMetrics\(\)/);
});
