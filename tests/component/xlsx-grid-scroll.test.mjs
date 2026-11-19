import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";
import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const { XlsxViewer } = await importFromDemo("@arcships/vue-xlsx");


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
  assert.match(GRID_SOURCE, /@keydown="onEditKeydown"/);
  assert.match(GRID_SOURCE, /@compositionstart="onEditCompositionStart"/);
  assert.match(GRID_SOURCE, /@compositionend="onEditCompositionEnd"/);
  assert.match(SCRIPT_SOURCE, /function commitEdit\(\)[\s\S]*nextTick\(\(\) => containerRef\.value\?\.focus\(\)\)/);
  assert.match(SCRIPT_SOURCE, /function cancelEdit\(\)[\s\S]*nextTick\(\(\) => containerRef\.value\?\.focus\(\)\)/);
  assert.match(functionNode("onEditKeydown").getText(AST), /isComposingKeyboardEvent\(event\)/);
  assert.match(functionNode("onEditCompositionEnd").getText(AST), /nextTick\(commitEditFillingSelection\)/);
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

  const hitTestSource = functionNode("hitTestCell").getText(AST);
  assert.match(hitTestSource, /getSheetCell\(displayRow, displayCol\)/);
  assert.match(hitTestSource, /resolvePanePoint\(/, "hit testing resolves the freeze pane before adding scroll");
  assert.doesNotMatch(hitTestSource, /container\.scroll(Left|Top)/);
  assert.equal(descendants(functionNode("hitTestCell"), ts.isForStatement).length, 0);
  assert.match(functionNode("paintBody").getText(AST), /getSheetCell\(r, c\)/);
  assert.match(declaration("editInputStyle").getText(AST), /paneOverlayRect\(/);
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

function installWheelEnvironment() {
  const previousWindow = globalThis.window;
  const previousResizeObserver = globalThis.ResizeObserver;
  const previousDocument = globalThis.Document;
  const previousShadowRoot = globalThis.ShadowRoot;
  globalThis.window = {
    devicePixelRatio: 1,
    requestAnimationFrame: () => 1,
    cancelAnimationFrame() {},
    addEventListener() {},
    removeEventListener() {},
  };
  globalThis.ResizeObserver = class ResizeObserver {
    observe() {}
    disconnect() {}
  };
  globalThis.Document = class Document {};
  globalThis.ShadowRoot = class ShadowRoot {};
  return () => {
    if (previousWindow === undefined) delete globalThis.window;
    else globalThis.window = previousWindow;
    if (previousResizeObserver === undefined) delete globalThis.ResizeObserver;
    else globalThis.ResizeObserver = previousResizeObserver;
    if (previousDocument === undefined) delete globalThis.Document;
    else globalThis.Document = previousDocument;
    if (previousShadowRoot === undefined) delete globalThis.ShadowRoot;
    else globalThis.ShadowRoot = previousShadowRoot;
  };
}

function createWheelController() {
  const sheet = {
    name: "Sheet1",
    workbookSheetIndex: 0,
    visibleRows: Array.from({ length: 200 }, (_, index) => index),
    visibleCols: Array.from({ length: 50 }, (_, index) => index),
    mergedRegions: [],
    colWidths: [],
    rowHeights: [],
    colWidthOverridesPx: [],
    rowHeightOverridesPx: [],
    colStyleIds: [],
    rowStyleIds: [],
    styleById: new Map(),
    hyperlinks: [],
    comments: [],
    sparklines: [],
    conditionalFormatRules: [],
    maxUsedRow: 0,
    maxUsedCol: 0,
    showGridLines: true,
    getCellStyleAt: () => null,
  };
  return vue.reactive({
    activeSheet: sheet,
    activeSheetIndex: 0,
    activeTab: null,
    activeTabIndex: 0,
    activeCell: { row: 0, col: 0 },
    activeCellAddress: "A1",
    selection: {
      start: { row: 0, col: 0 },
      end: { row: 0, col: 0 },
    },
    sheets: [sheet],
    tabs: [],
    tables: [],
    charts: [],
    images: [],
    shapes: [],
    formControls: [],
    isChartsLoading: false,
    maxConcurrentImageDecodes: 1,
    selectedChartId: null,
    selectedChartElement: null,
    selectedImageId: null,
    selectedShapeId: null,
    revision: 1,
    zoomScale: 100,
    isLoading: false,
    error: null,
    isWorkerBacked: false,
    readOnly: false,
    selectedFormula: "",
    selectedFormulaTarget: null,
    selectedValue: "",
    displayFileName: "wheel.xlsx",
    canDownload: false,
    canExport: false,
    canZoomIn: true,
    canZoomOut: true,
    getCellDisplayValue: () => "",
    getCellFormula: () => "",
    commitCellInput() {},
    getActiveWorksheet: () => null,
    getRowsBatchAsync: async () => [],
    setCellValue() {},
    selectCell() {},
    selectRange() {},
    clearSelectedCells() {},
    copySelectionToClipboard() {},
    pasteFromClipboard() {},
    resizeColumn() {},
    resizeRow() {},
    setActiveSheetIndex() {},
    selectChart() {},
    clearSelectedChart() {},
    selectChartElement() {},
    moveChartBy() {},
    resizeChartBy() {},
    deleteChart() {},
    selectImage() {},
    clearSelectedImage() {},
    moveImageBy() {},
    resizeImageBy() {},
    deleteImage() {},
    reportImageDecodeError() {},
    selectShape() {},
    clearSelectedShape() {},
    moveShapeBy() {},
    resizeShapeBy() {},
    deleteShape() {},
    setZoomScale(value) {
      this.zoomScale = value;
      this.revision += 1;
    },
  });
}

function nodesWithClass(root, className) {
  return walk(root).filter((node) =>
    String(node?.props?.class || "")
      .split(/\s+/)
      .filter(Boolean)
      .includes(className),
  );
}

test("XlsxViewer Ctrl/Cmd+wheel zoom keeps the cell under the cursor anchored", async (t) => {
  const restoreEnvironment = installWheelEnvironment();
  const controller = createWheelController();
  const mounted = await mount(XlsxViewer, {
    controller,
    showDefaultToolbar: false,
    showRibbon: false,
    showFormulaBar: false,
    showImages: false,
  });
  t.after(() => {
    mounted.app.unmount();
    restoreEnvironment();
  });

  const body = nodesWithClass(mounted.root, "xlsx-viewer__body")[0];
  const grid = findByTestId(mounted.root, "xlsx-grid");
  assert.ok(body, "viewer body renders");
  assert.ok(grid, "grid renders");
  grid.getBoundingClientRect = () => ({ left: 0, top: 0, right: 640, bottom: 480, width: 640, height: 480 });
  grid.scrollLeft = 0;
  grid.scrollTop = 0;
  await vue.nextTick();

  const ctrlWheel = {
    type: "wheel",
    deltaY: -200,
    deltaMode: 0,
    ctrlKey: true,
    metaKey: false,
    clientX: 248,
    clientY: 84,
    defaultPrevented: false,
    preventDefault() { this.defaultPrevented = true; },
  };
  body.dispatchEvent(ctrlWheel);
  assert.equal(ctrlWheel.defaultPrevented, true, "ctrl+wheel must prevent browser page zoom");
  await vue.nextTick();
  await vue.nextTick();
  assert.equal(controller.zoomScale, 200);
  // C3 (row 2 / column 2) is at 50% of its 80x24 cell before zoom. After a
  // 200% zoom the same 50% point is 200px/60px from the scroll origin.
  assert.equal(grid.scrollLeft, 200, "anchor column stays under the cursor");
  assert.equal(grid.scrollTop, 60, "anchor row stays under the cursor");

  const metaWheel = {
    ...ctrlWheel,
    ctrlKey: false,
    metaKey: true,
    deltaY: 200,
    defaultPrevented: false,
  };
  body.dispatchEvent(metaWheel);
  assert.equal(metaWheel.defaultPrevented, true, "cmd/meta+wheel must also stay inside the viewer");
  await vue.nextTick();
  await vue.nextTick();
  assert.ok(controller.zoomScale < 200, "cmd/meta+wheel zooms back out");

  const plainWheel = {
    ...ctrlWheel,
    ctrlKey: false,
    metaKey: false,
    defaultPrevented: false,
  };
  const plainZoom = controller.zoomScale;
  body.dispatchEvent(plainWheel);
  assert.equal(plainWheel.defaultPrevented, false, "plain wheel keeps native scrolling");
  assert.equal(controller.zoomScale, plainZoom, "plain wheel does not change zoom");
});
