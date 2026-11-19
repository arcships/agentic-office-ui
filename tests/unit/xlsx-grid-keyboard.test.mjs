import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";

const ROOT = path.resolve(import.meta.dirname, "../..");
const KEYBOARD_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/grid-keyboard.ts");
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

function loadTsModule(source) {
  const output = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
  return import(`data:text/javascript;base64,${Buffer.from(output).toString("base64")}`);
}

const keyboard = await loadTsModule(readFileSync(KEYBOARD_PATH, "utf8"));

function functionNode(name) {
  const matches = [];
  function visit(node) {
    if (ts.isFunctionDeclaration(node) && node.name?.text === name) matches.push(node);
    ts.forEachChild(node, visit);
  }
  visit(AST);
  return matches[0];
}

const ROW_HEIGHT = 24;
// 480 px grid body minus the 24 px column-header strip.
const BODY_HEIGHT = 456;

test("PageUp/PageDown step by one viewport of rows, keeping a row of overlap", () => {
  const heights = Array.from({ length: 200 }, () => ROW_HEIGHT);
  assert.equal(keyboard.resolvePageRowStep(heights, BODY_HEIGHT), 18);

  assert.equal(keyboard.resolvePageRowStep(heights, ROW_HEIGHT * 3), 2, "three visible rows move by two");
  assert.equal(keyboard.resolvePageRowStep(heights, 10), 1, "a viewport shorter than a row still moves");
  assert.equal(keyboard.resolvePageRowStep(heights, 0), 1, "an unmeasured viewport still moves one row");
  assert.equal(keyboard.resolvePageRowStep([], BODY_HEIGHT), 1, "a sheet without rows still moves");

  // Hidden rows report zero and must not eat viewport space.
  const withHidden = Array.from({ length: 200 }, (_, index) => (index % 2 === 0 ? 0 : ROW_HEIGHT));
  assert.equal(keyboard.resolvePageRowStep(withHidden, ROW_HEIGHT * 5), 4);
  assert.equal(
    keyboard.resolvePageRowStep(Array.from({ length: 10 }, () => 0), BODY_HEIGHT),
    1,
    "an all-hidden axis never divides by zero",
  );
});

test("Ctrl+Arrow stops at the end of a run, skips gaps and falls back to the axis edge", () => {
  //  0  1  2  3  4  5
  //  x  x  .  x  .  .
  const data = [true, true, false, true, false, false];
  const occupied = (index) => data[index] === true;

  assert.equal(keyboard.resolveJumpTarget(0, 1, data.length, occupied), 1, "run stops on its last cell");
  assert.equal(keyboard.resolveJumpTarget(1, 1, data.length, occupied), 3, "a gap jumps to the next entry");
  assert.equal(keyboard.resolveJumpTarget(3, 1, data.length, occupied), 5, "no entry left lands on the axis edge");
  assert.equal(keyboard.resolveJumpTarget(4, 1, data.length, occupied), 5, "an empty tail lands on the axis edge");
  assert.equal(keyboard.resolveJumpTarget(5, -1, data.length, occupied), 3, "walking back finds the last entry");
  assert.equal(keyboard.resolveJumpTarget(0, -1, data.length, occupied), 0, "the axis start is a wall");

  const empty = () => false;
  assert.equal(keyboard.resolveJumpTarget(2, 1, data.length, empty), 5);
  assert.equal(keyboard.resolveJumpTarget(2, -1, data.length, empty), 0);

  const full = () => true;
  assert.equal(keyboard.resolveJumpTarget(0, 1, data.length, full), 5, "a full axis has no inner edge");

  assert.equal(keyboard.resolveJumpTarget(3, 1, 0, occupied), 3, "an empty axis keeps the cursor");
});

test("horizontal jumps follow the plain-arrow direction on LTR and RTL sheets", () => {
  assert.equal(keyboard.resolveHorizontalJumpDirection("ArrowLeft", false), -1);
  assert.equal(keyboard.resolveHorizontalJumpDirection("ArrowRight", false), 1);
  assert.equal(
    keyboard.resolveHorizontalJumpDirection("ArrowLeft", true),
    1,
    "an RTL sheet mirrors the column axis, so ← points at higher display indices",
  );
  assert.equal(keyboard.resolveHorizontalJumpDirection("ArrowRight", true), -1);
});

test("Ctrl+End / Ctrl+A clamp the used extent to the displayed rows and columns", () => {
  assert.equal(
    keyboard.resolveLastUsedDisplayIndex([0, 1, 2, 3], 2, 3),
    2,
    "the used end wins while it is displayed",
  );
  assert.equal(
    keyboard.resolveLastUsedDisplayIndex([0, 1, 3, 4], 2, 3),
    1,
    "a hidden trailing row cannot carry the target past the used range",
  );
  assert.equal(
    keyboard.resolveLastUsedDisplayIndex([0, 1, 2], 99, 2),
    2,
    "a used end past the display window stops on the last displayed index",
  );
  assert.equal(keyboard.resolveLastUsedDisplayIndex([5, 6], 2, 1), 0);
  assert.equal(keyboard.resolveLastUsedDisplayIndex([0, 1], -1, 1), 1, "an empty sheet keeps the fallback");
  assert.equal(keyboard.resolveLastUsedDisplayIndex([], 4, 7), 7);
});

test("range identity drives the Ctrl+A escalation and the Escape collapse", () => {
  const a1 = { row: 0, col: 0 };
  const c3 = { row: 2, col: 2 };
  assert.equal(keyboard.isSameCell(a1, { row: 0, col: 0 }), true);
  assert.equal(keyboard.isSameCell(a1, { row: 0, col: 1 }), false);
  assert.equal(keyboard.isSameCell(a1, null), false);

  assert.equal(keyboard.isSameRange({ start: a1, end: c3 }, { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } }), true);
  assert.equal(keyboard.isSameRange({ start: a1, end: c3 }, { start: a1, end: a1 }), false);
  assert.equal(keyboard.isSameRange(null, { start: a1, end: c3 }), false);
});

test("onGridKeydown wires the Excel navigation keys to the grid helpers", () => {
  const keydown = functionNode("onGridKeydown");
  assert.ok(keydown, "missing onGridKeydown");
  const source = keydown.getText(AST);

  assert.match(source, /event\.key === "Escape"[\s\S]*cancelCut[\s\S]*collapseSelectionToActiveCell\(\)/);
  assert.match(source, /event\.key\.toLocaleLowerCase\(\) === "a"[\s\S]*selectUsedRange\(\)/);
  assert.match(source, /isArrowKey\(event\.key\)[\s\S]*jumpToEdge\(event\.key, event\.shiftKey\)/);
  assert.match(source, /event\.key === "PageDown" \|\| event\.key === "PageUp"[\s\S]*moveCellByPage/);
  assert.match(source, /event\.key === "Home" \|\| event\.key === "End"[\s\S]*moveToSheetEdge[\s\S]*moveToRowEdge/);
  assert.match(source, /Enter: \(\) => \{[\s\S]*event\.shiftKey[\s\S]*moveCell\(-1, 0, false\)/);

  const escape = source.slice(source.indexOf('"Escape"'));
  assert.ok(
    escape.indexOf("cancelCut") < escape.indexOf("collapseSelectionToActiveCell"),
    "a pending cut is cancelled before the selection collapses",
  );

  assert.match(functionNode("selectUsedRange").getText(AST), /selectAxisRange\(\{ actualIndex: 0, kind: "corner" \}\)/);
  assert.match(functionNode("moveCellByPage").getText(AST), /resolvePageRowStep\(effectiveRowHeights\.value, bodyViewportSize\(\)\.height\)/);
  assert.match(functionNode("jumpToEdge").getText(AST), /resolveJumpTarget\(/);
  assert.match(
    functionNode("jumpToEdge").getText(AST),
    /resolveHorizontalJumpDirection\(key, isRightToLeft\.value\)/,
    "Ctrl+Arrow must reuse the plain-arrow direction on RTL sheets",
  );
  assert.match(functionNode("usedDisplayEdge").getText(AST), /resolveLastUsedDisplayIndex\(/);
  assert.match(functionNode("collapseSelectionToActiveCell").getText(AST), /selectCell\(cell\)/);
  assert.match(functionNode("moveToRowEdge").getText(AST), /isCellOccupied\(candidate\)/);
});

test("the inline editor commits up with Shift+Enter and fills a multi-cell selection on Enter", () => {
  const editKeydown = functionNode("onEditKeydown").getText(AST);
  assert.match(editKeydown, /event\.key === "Enter"\)[\s\S]*event\.shiftKey[\s\S]*commitAndMoveCell\(-1, 0\)/);
  assert.match(editKeydown, /else commitEditFillingSelection\(\)/);

  // Tab, blur and commit-and-move only write the active cell; Enter is the
  // one confirm that fills a selected range.
  for (const name of ["commitEdit", "commitAndTab", "commitAndMoveCell"]) {
    const source = functionNode(name).getText(AST);
    assert.match(source, /commitEditingDraft\(/, `${name} must write through the shared draft commit`);
    assert.doesNotMatch(source, /commitEditingDraft\([^)]*true/, `${name} must not fill the selection`);
  }
  const draft = functionNode("commitEditingDraft").getText(AST);
  assert.match(draft, /fillSelection = false/);
  assert.match(draft, /fillSelection && hasRangeSelection\(\)/);
  assert.match(draft, /commitSelectionInput/);
  assert.match(draft, /commitCellInput\(cell, editingValue\.value\)/);
  assert.match(
    functionNode("commitEditFillingSelection").getText(AST),
    /commitEditingDraft\(editingCell\.value, true\)/,
  );

  assert.match(GRID_SOURCE, /@blur="commitEdit"/);
  assert.match(GRID_SOURCE, /@keydown="onGridKeydown"/);
  assert.match(GRID_SOURCE, /@keydown="onEditKeydown"/);
});

test("the editing domain writes a filled range as one history entry", () => {
  const editingPath = path.join(ROOT, "packages/vue-xlsx/src/composables/editing.ts");
  const source = readFileSync(editingPath, "utf8");
  assert.match(source, /function commitSelectionInput\(input: string\)/);
  assert.match(source, /function writeCellValue\(/);
  const selectionInput = source.slice(
    source.indexOf("function commitSelectionInput"),
    source.indexOf("function setCellFormula"),
  );
  assert.match(selectionInput, /isMergedSecondary/);
  assert.match(selectionInput, /writeCellValue\(cell, text\)/);
  assert.match(selectionInput, /recordRangeEditHistory\(mutations, normalized, active\)/);
  assert.doesNotMatch(selectionInput, /recordCellEditHistory/, "a fill is one range entry, not one per cell");

  const setCellValue = source.slice(source.indexOf("function setCellValue"), source.indexOf("function commitSelectionInput"));
  assert.match(setCellValue, /writeCellValue\(cell, value\)/);
  assert.match(setCellValue, /recordCellEditHistory\(cell, before, after\)/);

  assert.match(
    readFileSync(path.join(ROOT, "packages/vue-xlsx/src/composables/useXlsxViewerController.ts"), "utf8"),
    /commitSelectionInput: editingDomain\.commitSelectionInput/,
  );
  assert.match(
    readFileSync(path.join(ROOT, "packages/xlsx-core/src/types/worksheet-types.ts"), "utf8"),
    /commitSelectionInput: \(text: string\) => number/,
  );
});
