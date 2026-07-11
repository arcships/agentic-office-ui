import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";


const ROOT = path.resolve(import.meta.dirname, "../..");
const GRID_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxGrid.vue");
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

async function loadGeometryHelpers() {
  const names = [
    "createOffsetIndex",
    "findDisplayIndexAtOffset",
    "resolveVisibleRange",
    "resolveDisplayIndices",
    "createSheetToDisplayIndex",
    "resolveAdjacentSheetCell",
    "resolveAxisResizeRelease",
  ];
  const source = names.map((name) => {
    const node = functionNode(name);
    assert.ok(node, `missing ${name}`);
    return node.getText(AST);
  }).join("\n");
  const output = ts.transpileModule(
    `${source}\nexport { ${names.join(", ")} };`,
    {
      compilerOptions: {
        module: ts.ModuleKind.ESNext,
        target: ts.ScriptTarget.ES2022,
      },
    },
  ).outputText;
  return import(`data:text/javascript;base64,${Buffer.from(output).toString("base64")}`);
}

const geometry = await loadGeometryHelpers();

test("XlsxGrid cumulative offsets and binary lookup preserve exact cell boundaries", () => {
  const offsets = geometry.createOffsetIndex([10, 25, 5]);
  assert.deepEqual(offsets, [0, 10, 35, 40]);

  for (const [offset, expected] of [
    [-0.01, -1],
    [0, 0],
    [9.99, 0],
    [10, 1],
    [34.99, 1],
    [35, 2],
    [39.99, 2],
    [40, -1],
  ]) {
    assert.equal(geometry.findDisplayIndexAtOffset(offsets, offset), expected, `offset ${offset}`);
  }

  assert.deepEqual(geometry.resolveVisibleRange(offsets, -240, 34), { start: 0, end: 1 });
  assert.deepEqual(geometry.resolveVisibleRange(offsets, 10, 35), { start: 1, end: 2 });
  assert.deepEqual(geometry.resolveVisibleRange(offsets, 36, 300), { start: 2, end: 2 });
  assert.deepEqual(geometry.resolveVisibleRange([0], 0, 100), { start: 0, end: -1 });
});

test("XlsxGrid boundary click selects the row or column unless the pointer actually drags", () => {
  assert.equal(geometry.resolveAxisResizeRelease(false, "pointerup"), "select");
  assert.equal(geometry.resolveAxisResizeRelease(true, "pointerup"), "resize");
  assert.equal(geometry.resolveAxisResizeRelease(false, "pointercancel"), "cancel");

  const pointerDown = functionNode("onPointerDown").getText(AST);
  const pointerMove = functionNode("onPointerMove").getText(AST);
  const pointerUp = functionNode("onPointerUp").getText(AST);
  assert.match(pointerDown, /moved:\s*false/);
  assert.match(pointerMove, /AXIS_RESIZE_DRAG_THRESHOLD/);
  assert.match(pointerUp, /releaseAction === "select"/);
  assert.match(pointerUp, /selectAxisRange\(resizing\)/);
});

test("XlsxGrid returns keyboard focus after committing or cancelling cell editing", () => {
  assert.match(GRID_SOURCE, /@keydown\.enter\.stop\.prevent="commitEdit"/);
  assert.match(GRID_SOURCE, /@keydown\.escape\.stop\.prevent="cancelEdit"/);
  assert.match(SCRIPT_SOURCE, /function commitEdit\(\)[\s\S]*nextTick\(\(\) => containerRef\.value\?\.focus\(\)\)/);
  assert.match(SCRIPT_SOURCE, /function cancelEdit\(\)[\s\S]*nextTick\(\(\) => containerRef\.value\?\.focus\(\)\)/);
});

test("XlsxGrid display indices keep hidden row and column workbook addresses", () => {
  const rows = geometry.resolveDisplayIndices([0, 2, 5], 3);
  const cols = geometry.resolveDisplayIndices([1, 4], 2);
  assert.deepEqual(rows, [0, 2, 5]);
  assert.deepEqual(cols, [1, 4]);
  assert.deepEqual(geometry.resolveDisplayIndices(undefined, 3), [0, 1, 2]);

  const rowMap = geometry.createSheetToDisplayIndex(rows);
  const colMap = geometry.createSheetToDisplayIndex(cols);
  assert.equal(rowMap.get(5), 2);
  assert.equal(rowMap.has(1), false);
  assert.equal(colMap.get(4), 1);
  assert.deepEqual(
    geometry.resolveAdjacentSheetCell(
      { row: 2, col: 1 },
      0,
      1,
      rows,
      cols,
      rowMap,
      colMap,
    ),
    { row: 2, col: 4 },
    "Tab skips a hidden column but preserves the workbook row",
  );
  assert.deepEqual(
    geometry.resolveAdjacentSheetCell(
      { row: 2, col: 1 },
      1,
      0,
      rows,
      cols,
      rowMap,
      colMap,
    ),
    { row: 5, col: 1 },
    "Arrow navigation skips a hidden row",
  );

  assert.match(functionNode("hitTestCell").getText(AST), /getSheetCell\(displayRow, displayCol\)/);
  assert.match(functionNode("hitTestCell").getText(AST), /container\.scrollLeft/);
  assert.equal(descendants(functionNode("hitTestCell"), ts.isForStatement).length, 0);
  assert.match(functionNode("paintBody").getText(AST), /getSheetCell\(r, c\)/);
  assert.match(declaration("editInputStyle").getText(AST), /getDisplayCell\(cell\)/);
  assert.match(functionNode("commitAndTab").getText(AST), /getAdjacentCell\(editingCell\.value, 0, 1\)/);
});

test("XlsxGrid owns a real scroll surface and coalesces scroll painting by animation frame", () => {
  assert.match(GRID_SOURCE, /data-testid="xlsx-grid-scroll-space"/);
  assert.match(GRID_SOURCE, /class="xlsx-grid__viewport"/);
  assert.match(GRID_SOURCE, /position:\s*sticky/);

  const scheduleSource = functionNode("schedulePaint").getText(AST);
  const scrollSource = functionNode("onScroll").getText(AST);
  assert.match(scheduleSource, /paintFrame !== null/);
  assert.match(scheduleSource, /requestAnimationFrame/);
  assert.match(scheduleSource, /paintGrid\(\)/);
  assert.match(scrollSource, /schedulePaint\(\)/);
  assert.doesNotMatch(scrollSource, /paintGrid\(\)/);
});

test("XlsxGrid canvas dimensions are guarded against scroll-only rewrites", () => {
  const resize = functionNode("resizeCanvas");
  assert.ok(resize);
  const assignments = descendants(resize, (node) =>
    ts.isBinaryExpression(node) &&
    node.operatorToken.kind === ts.SyntaxKind.EqualsToken &&
    ts.isPropertyAccessExpression(node.left) &&
    ts.isIdentifier(node.left.expression) &&
    node.left.expression.text === "canvas" &&
    ["width", "height"].includes(node.left.name.text)
  );
  assert.equal(assignments.length, 2);
  for (const assignment of assignments) {
    let parent = assignment.parent;
    let guarded = false;
    while (parent && parent !== resize) {
      if (ts.isIfStatement(parent)) {
        guarded = true;
        break;
      }
      parent = parent.parent;
    }
    assert.equal(guarded, true, `${assignment.getText(AST)} must be conditional`);
  }
});

test("XlsxGrid hydrates visible worker rows without falling back to a main-thread workbook", () => {
  const hydrate = functionNode("hydrateVisibleWorkerRows");
  const display = functionNode("getGridCellDisplayValue");
  assert.ok(hydrate);
  assert.ok(display);
  assert.match(hydrate.getText(AST), /controller\.getRowsBatchAsync/);
  assert.match(hydrate.getText(AST), /WORKER_ROW_BATCH_SIZE/);
  assert.match(hydrate.getText(AST), /workerCellCache\.set/);
  assert.match(display.getText(AST), /controller\.isWorkerBacked/);
  assert.match(functionNode("paintBody").getText(AST), /getGridCellDisplayValue\(cell\)/);
  assert.match(functionNode("schedulePaint").getText(AST), /hydrateVisibleWorkerRows\(\)/);
});
