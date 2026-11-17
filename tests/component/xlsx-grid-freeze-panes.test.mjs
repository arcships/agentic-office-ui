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
const { XlsxGrid } = await importFromDemo("@arcships/vue-xlsx");

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
  for (const name of ["editInputStyle", "hoveredComment", "validationHint", "validationDropdownTriggerStyle", "validationDropdownStyle", "validationFeedbackView"]) {
    const node = declaration(name);
    assert.ok(node, `XlsxGrid is missing ${name}`);
    const text = node.getText(AST);
    assert.match(text, /paneOverlayRect\(/, `${name} must resolve and clip the pane`);
    assert.doesNotMatch(text, /-\s*scrollLeft\.value/, `${name} must not subtract scroll unconditionally`);
    assert.doesNotMatch(text, /-\s*scrollTop\.value/, `${name} must not subtract scroll unconditionally`);
  }
});

test("XLSX scrollToCell leaves pinned cells where they are", () => {
  const source = functionText("scrollToCell");
  assert.match(source, /frozenPaneMetrics\(\)/);
  assert.match(source, /displayCell\.col >= frozen\.cols/, "a pinned column never scrolls into view");
  assert.match(source, /displayCell\.row >= frozen\.rows/, "a pinned row never scrolls into view");
  assert.match(source, /colX - container\.scrollLeft < frozen\.width/, "the freeze line hides the cell's left edge");
  assert.match(source, /rowY - container\.scrollTop < frozen\.height/);
  assert.match(source, /colX \+ colW - container\.scrollLeft > bodyWidth/);
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


// ── Rendered behaviour ───────────────────────────────────────────────
// The geometry above is pure; the rest mounts the real component against a
// recording 2D context that honours save/restore, translate and clip, so each
// issue bullet is checked against what actually reaches the canvas and where
// the DOM overlays land.

const HEADER_HEIGHT = 24;
const ROW_HEADER_WIDTH = 48;
const BODY_WIDTH = 592;
const BODY_HEIGHT = 456;
const SCROLL_X = 150;
const SCROLL_Y = 100;
const SELECTION_FILL = "#123456";
const SELECTION_BORDER = "#654321";

function frozenSheet(overrides = {}) {
  const values = { "0:0": "42", "0:3": "D1", "1:1": "B2" };
  return {
    name: "Sheet1",
    visibility: "visible",
    workbookSheetIndex: 0,
    freezePanes: { row: 1, col: 1 },
    visibleRows: Array.from({ length: 40 }, (_, index) => index),
    visibleCols: Array.from({ length: 20 }, (_, index) => index),
    rowHeights: Array.from({ length: 40 }, () => ROW_HEIGHT),
    colWidths: Array.from({ length: 20 }, () => COL_WIDTH),
    rowHeightOverridesPx: {},
    colWidthOverridesPx: {},
    rowStyleIds: { 0: 1 },
    colStyleIds: {},
    styleById: { 1: { font: { bold: true }, alignment: { vertical: "top" } } },
    namedCellStyleByName: {},
    tableStyleByName: {},
    cachedFormulaValues: {},
    themePalette: {},
    defaultRowHeightPx: ROW_HEIGHT,
    defaultColWidthPx: COL_WIDTH,
    hasHorizontalMerges: false,
    hasVerticalMerges: false,
    maxHorizontalMergeEndCol: 0,
    maxVerticalMergeEndRow: 0,
    minUsedCol: 0,
    minUsedRow: 0,
    maxUsedCol: 3,
    maxUsedRow: 1,
    rowCount: 40,
    colCount: 20,
    showGridLines: true,
    mergedRegions: [],
    hyperlinks: [{ cell: { row: 0, col: 0 }, target: "https://example.com/report" }],
    comments: [{ cell: { row: 0, col: 0 }, author: "Ada", text: "frozen note" }],
    dataValidations: [],
    sparklines: [],
    conditionalFormatRules: [
      {
        kind: "dataBar",
        priority: 1,
        ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }],
        cfvos: [],
        showValue: true,
      },
    ],
    ...overrides,
    values,
  };
}

function createController(sheet) {
  const controller = vue.reactive({
    activeSheet: sheet,
    activeCell: { row: 0, col: 0 },
    selection: null,
    revision: 1,
    zoomScale: 100,
    tables: [],
    isWorkerBacked: false,
    readOnly: false,
    isEditingLocked: false,
    sheets: [sheet],
    dataValidationFeedback: null,
    cellValidation: null,
    getCellDisplayValue(cell) {
      return sheet.values[`${cell.row}:${cell.col}`] ?? "";
    },
    getCellFormula() {
      return undefined;
    },
    getCellDataValidation() {
      return controller.cellValidation;
    },
    getDataValidationListOptions() {
      return [];
    },
    getActiveWorksheet() {
      return null;
    },
    setZoomScale(value) {
      controller.zoomScale = value;
    },
    setActiveSheetIndex() {},
    selectCell(cell) {
      controller.activeCell = cell;
    },
    selectRange(range) {
      controller.selection = range;
    },
    commitCellInput(cell, value) {
      sheet.values[`${cell.row}:${cell.col}`] = value;
    },
    resizeColumn() {},
    resizeRow() {},
    setCellValue() {},
    clearSelectedCells() {},
    copySelectionToClipboard() {},
    pasteFromClipboard() {},
    dismissDataValidationFeedback() {},
  });
  return controller;
}

const CONTEXT_PROPERTIES = new Set(["fillStyle", "strokeStyle", "font", "textAlign", "textBaseline", "lineWidth"]);

/**
 * Minimal 2D context: records draw calls that survive the current clip, so a
 * pass that paints outside its clip region is observably silent. Operations are
 * tagged with the canvas class rather than the node object: Vue clones the
 * hoisted static canvases, so node identity is not stable across renders.
 */
function createRecordingContext(canvas, operations) {
  const style = { fillStyle: "#000000", strokeStyle: "#000000", font: "10px sans-serif", textAlign: "start", textBaseline: "alphabetic", lineWidth: 1 };
  const snapshot = () => ({ ...style });
  const canvasName = () => String(canvas.props?.class ?? canvas.type)
    .split(/\s+/)
    .find((name) => name.startsWith("xlsx-grid__")) ?? String(canvas.type);
  let transform = { x: 0, y: 0 };
  let clip = null;
  let pathRects = [];
  let pathPoints = [];
  const stack = [];

  const canvasSize = () => ({ width: Number(canvas.width) || 0, height: Number(canvas.height) || 0 });
  const deviceRect = (x, y, width, height) => {
    const left = x + transform.x;
    const top = y + transform.y;
    return { bottom: top + height, left, right: left + width, top };
  };
  const intersect = (left, right) => {
    const region = {
      bottom: Math.min(left.bottom, right.bottom),
      left: Math.max(left.left, right.left),
      right: Math.min(left.right, right.right),
      top: Math.max(left.top, right.top),
    };
    return region.right > region.left && region.bottom > region.top ? region : null;
  };
  const visible = (rect) => (clip ? clip.some((region) => intersect(region, rect) !== null) : true);
  const record = (name, args) => operations.push({ canvas: canvasName(), name, args, ...snapshot() });

  return new Proxy(style, {
    get(_target, property) {
      if (typeof property !== "string") return undefined;
      switch (property) {
        case "measureText":
          return (value) => ({ width: String(value).length * 7 });
        case "createLinearGradient":
        case "createRadialGradient":
          return () => ({ addColorStop() {} });
        case "save":
          return () => stack.push({ clip: clip ? clip.map((region) => ({ ...region })) : null, transform: { ...transform } });
        case "restore":
          return () => {
            const saved = stack.pop();
            if (!saved) return;
            clip = saved.clip;
            transform = saved.transform;
          };
        case "setTransform":
          return (_a, _b, _c, _d, e, f) => { transform = { x: e || 0, y: f || 0 }; };
        case "translate":
          return (x, y) => { transform = { x: transform.x + x, y: transform.y + y }; };
        case "beginPath":
          return () => { pathRects = []; pathPoints = []; };
        case "rect":
          return (x, y, width, height) => { pathRects.push(deviceRect(x, y, width, height)); };
        case "moveTo":
        case "lineTo":
          return (x, y) => {
            const rect = deviceRect(x, y, 1, 1);
            pathPoints.push(rect);
            if (visible(rect)) record(property, [x, y]);
          };
        case "arc":
          return (x, y) => {
            const rect = deviceRect(x, y, 1, 1);
            pathPoints.push(rect);
            if (visible(rect)) record("arc", [x, y]);
          };
        case "clip":
          return () => {
            const size = canvasSize();
            const base = clip ?? [{ bottom: size.height, left: 0, right: size.width, top: 0 }];
            clip = base
              .flatMap((region) => pathRects.map((rect) => intersect(region, rect)))
              .filter(Boolean);
          };
        case "fillRect":
          return (x, y, width, height) => { if (visible(deviceRect(x, y, width, height))) record("fillRect", [x, y, width, height]); };
        case "strokeRect":
          return (x, y, width, height) => { if (visible(deviceRect(x, y, width, height))) record("strokeRect", [x, y, width, height]); };
        case "fillText":
          return (text, x, y) => {
            const fontSize = Number.parseFloat(style.font) || 12;
            const rect = deviceRect(x, y - fontSize, String(text).length * 7, fontSize + 2);
            if (visible(rect)) record("fillText", [text, x, y]);
          };
        case "stroke":
        case "fill":
          return () => {
            const bounds = pathPoints.reduce((acc, point) => (acc
              ? {
                bottom: Math.max(acc.bottom, point.bottom),
                left: Math.min(acc.left, point.left),
                right: Math.max(acc.right, point.right),
                top: Math.min(acc.top, point.top),
              }
              : { ...point }), null);
            if (bounds && visible(bounds)) record(property, []);
          };
        default:
          return (...args) => { record(property, args); };
      }
    },
    set(_target, property, value) {
      style[property] = value;
      if (CONTEXT_PROPERTIES.has(property)) record(`set:${property}`, [value]);
      return true;
    },
  });
}

function installBrowserEnvironment(t) {
  const originalGetContext = Object.getOwnPropertyDescriptor(Object.prototype, "getContext");
  const previousWindow = globalThis.window;
  const previousResizeObserver = globalThis.ResizeObserver;
  const operations = [];
  const frames = [];

  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value(kind) {
      if (this.type !== "canvas" || kind !== "2d") return null;
      if (Object.hasOwn(this, "__gridContext")) return this.__gridContext;
      Object.defineProperty(this, "__gridContext", {
        configurable: true,
        value: createRecordingContext(this, operations),
      });
      return this.__gridContext;
    },
  });

  globalThis.window = {
    devicePixelRatio: 1,
    requestAnimationFrame(callback) {
      frames.push(callback);
      return frames.length;
    },
    cancelAnimationFrame() {},
    addEventListener() {},
    removeEventListener() {},
  };
  globalThis.ResizeObserver = class ResizeObserver {
    constructor(callback) {
      this.callback = callback;
    }
    observe() {
      this.callback([{ contentRect: { width: 640, height: 480 } }]);
    }
    disconnect() {}
  };

  const cleanup = () => {
    if (originalGetContext) Object.defineProperty(Object.prototype, "getContext", originalGetContext);
    else delete Object.prototype.getContext;
    if (previousWindow === undefined) delete globalThis.window;
    else globalThis.window = previousWindow;
    if (previousResizeObserver === undefined) delete globalThis.ResizeObserver;
    else globalThis.ResizeObserver = previousResizeObserver;
  };

  const flush = async () => {
    for (let pass = 0; pass < 5 && frames.length; pass += 1) {
      const pending = frames.splice(0, frames.length);
      // Keep the operations of the last paint only; watchers may schedule a
      // follow-up frame and assertions should describe one rendered state.
      operations.length = 0;
      for (const callback of pending) callback();
      await vue.nextTick();
    }
  };

  return { operations, flush, cleanup };
}

async function mountFrozenGrid(t, sheetOverrides = {}) {
  const environment = installBrowserEnvironment(t);
  const sheet = frozenSheet(sheetOverrides);
  const controller = createController(sheet);
  let api;
  const Harness = vue.defineComponent({
    setup() {
      const gridRef = vue.ref(null);
      vue.watch(gridRef, (value) => { api = value; }, { flush: "sync" });
      return () => vue.h(XlsxGrid, {
        ref: gridRef,
        controller,
        selectionColor: SELECTION_BORDER,
        selectionFillColor: SELECTION_FILL,
      });
    },
  });
  const mounted = await mount(Harness);
  t.after(() => {
    mounted.app.unmount();
    environment.cleanup();
  });
  // Canvas host nodes get a style bag before the first paint writes sizes.
  for (const node of walk(mounted.root)) {
    if (node.type === "canvas") node.style = {};
  }

  const treeContainer = findByTestId(mounted.root, "xlsx-grid");
  assert.ok(treeContainer, "grid container must render");
  // Vue clones the hoisted static subtree, so the element the component
  // measures (`scrollContainer`) is not the node the tree walk returns. Drive
  // the component through its own element and stub both.
  const container = api.scrollContainer ?? treeContainer;
  for (const node of new Set([treeContainer, container])) {
    node.getBoundingClientRect = () => ({ left: 0, top: 0, width: 640, height: 480 });
    Object.defineProperty(node, "clientWidth", { configurable: true, value: 640 });
    Object.defineProperty(node, "clientHeight", { configurable: true, value: 480 });
    node.scrollLeft = 0;
    node.scrollTop = 0;
    node.style = {};
    node.focus = () => {};
  }
  await environment.flush();
  const canvasOperations = (className) => {
    const canvas = walk(mounted.root).find(
      (node) => node.type === "canvas" && String(node.props?.class).includes(className),
    );
    assert.ok(canvas, `missing ${className} canvas`);
    return environment.operations.filter((operation) => operation.canvas === className);
  };

  const scrollTo = async (scrollLeft, scrollTop) => {
    container.scrollLeft = scrollLeft;
    container.scrollTop = scrollTop;
    environment.operations.length = 0;
    // Host nodes do not dispatch scroll events, so drive the handler directly.
    container.props.onScroll?.({ target: container });
    await environment.flush();
  };

  const find = (testId) =>
    walk(mounted.root).find((node) => node.props?.["data-testid"] === testId)
    ?? walk(container).find((node) => node.props?.["data-testid"] === testId);

  return { mounted, container, controller, api, sheet, canvasOperations, scrollTo, find, flush: environment.flush };
}

test("XLSX frozen cells are painted by the frozen pass alone", async (t) => {
  const env = await mountFrozenGrid(t);
  env.controller.selection = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };

  await env.scrollTo(SCROLL_X, SCROLL_Y);
  const scrolled = env.canvasOperations("xlsx-grid__body")
    .filter((operation) => operation.name === "fillRect" && operation.fillStyle === SELECTION_FILL);
  assert.deepEqual(
    scrolled.map((operation) => [operation.args[0], operation.args[1]]),
    [[1, 1]],
    "the scrolling pass must not repaint the pinned corner cell at scrolled coordinates",
  );

  await env.scrollTo(0, 0);
  const unscrolled = env.canvasOperations("xlsx-grid__body")
    .filter((operation) => operation.name === "fillRect" && operation.fillStyle === SELECTION_FILL);
  assert.equal(unscrolled.length, 1, "an unpinned scroll still paints a pinned cell exactly once");
});

test("XLSX frozen cells keep their formatting when repainted", async (t) => {
  const env = await mountFrozenGrid(t);
  await env.scrollTo(SCROLL_X, SCROLL_Y);
  const body = env.canvasOperations("xlsx-grid__body");

  const dataBars = body.filter(
    (operation) =>
      operation.name === "fillRect" &&
      operation.args[1] === 4 &&
      operation.args[2] > 1 &&
      operation.args[3] === ROW_HEIGHT - 8,
  );
  assert.equal(dataBars.length, 1, "the pinned cell paints exactly one data bar");
  assert.ok(dataBars[0].args[2] > 0, "the data bar keeps its width");

  const link = body.find(
    (operation) =>
      operation.name === "fillText" &&
      operation.args[0] === "42" &&
      operation.args[1] === 5 &&
      operation.args[2] === 14,
  );
  assert.ok(link, "pinned text keeps its cell-relative position and top alignment");
  assert.equal(link.fillStyle, "#2563eb", "hyperlink colour");
  assert.match(link.font, /700 12px/, "workbook font");

  assert.ok(
    body.some(
      (operation) =>
        operation.name === "strokeRect" &&
        operation.args[0] === 0.5 &&
        operation.args[1] === 0.5 &&
        operation.args[2] === COL_WIDTH - 1,
    ),
    "pinned cells draw their own gridline box",
  );
  assert.ok(
    body.some(
      (operation) =>
        operation.name === "moveTo" &&
        operation.args[0] === COL_WIDTH - 9 &&
        operation.args[1] === 1 &&
        operation.fillStyle === "#dc2626",
    ),
    "the pinned cell keeps its comment marker",
  );
});

test("XLSX frozen rows stay pinned while the columns scroll", async (t) => {
  const env = await mountFrozenGrid(t);
  await env.scrollTo(SCROLL_X, SCROLL_Y);
  const body = env.canvasOperations("xlsx-grid__body");
  const frozenRowText = body.filter(
    (operation) => operation.name === "fillText" && operation.args[0] === "D1",
  );
  assert.equal(frozenRowText.length, 1, "the pinned row repaints its scrolled columns once");
  assert.equal(frozenRowText[0].args[1], 3 * COL_WIDTH - SCROLL_X + 5);
  assert.equal(frozenRowText[0].args[2], 14, "the pinned row keeps its y");
});

test("XLSX column headers stay above their pinned column while scrolled", async (t) => {
  const env = await mountFrozenGrid(t);
  await env.scrollTo(SCROLL_X, SCROLL_Y);
  const labels = env.canvasOperations("xlsx-grid__col-header")
    .filter((operation) => operation.name === "fillText");

  const pinned = labels.find((operation) => operation.args[0] === "A");
  assert.ok(pinned, "column A keeps its header");
  assert.equal(pinned.args[1], COL_WIDTH / 2, "the pinned label ignores scrollLeft");
  const scrolled = labels.find((operation) => operation.args[0] === "F");
  assert.ok(scrolled, "column F is painted in the scrolling strip");
  assert.equal(scrolled.args[1], 5 * COL_WIDTH - SCROLL_X + COL_WIDTH / 2);
});

test("XLSX row headers stay beside their pinned row while scrolled", async (t) => {
  const env = await mountFrozenGrid(t);
  await env.scrollTo(SCROLL_X, SCROLL_Y);
  const labels = env.canvasOperations("xlsx-grid__row-header")
    .filter((operation) => operation.name === "fillText");

  const pinned = labels.find((operation) => operation.args[0] === "1");
  assert.ok(pinned, "row 1 keeps its header");
  assert.equal(pinned.args[2], ROW_HEIGHT / 2, "the pinned label ignores scrollTop");
  const scrolled = labels.find((operation) => operation.args[0] === "7");
  assert.ok(scrolled, "row 7 is painted in the scrolling strip");
  assert.equal(scrolled.args[2], 6 * ROW_HEIGHT - SCROLL_Y + ROW_HEIGHT / 2);
});

test("XLSX hit tests select the pinned cell and header under a freeze", async (t) => {
  const env = await mountFrozenGrid(t);
  await env.scrollTo(SCROLL_X, SCROLL_Y);

  assert.deepEqual(
    env.api.hitTestCell(ROW_HEADER_WIDTH + 30, HEADER_HEIGHT + 10),
    { row: 0, col: 0 },
    "a click on the pinned corner cell selects that cell",
  );
  assert.deepEqual(
    env.api.hitTestCell(ROW_HEADER_WIDTH + 3 * COL_WIDTH - SCROLL_X + 10, HEADER_HEIGHT + 5 * ROW_HEIGHT - SCROLL_Y + 10),
    { row: 5, col: 3 },
    "the scrolling pane still adds the scroll offset",
  );
  assert.deepEqual(
    env.api.hitTestAxis(ROW_HEADER_WIDTH + 30, 10),
    { actualIndex: 0, displayIndex: 0, kind: "column", resize: false },
    "the header above a pinned column reports that column",
  );
  assert.deepEqual(
    env.api.hitTestAxis(ROW_HEADER_WIDTH + COL_WIDTH - 2, 10),
    { actualIndex: 0, displayIndex: 0, kind: "column", resize: true },
    "the pinned column's right edge stays a resize handle",
  );
});

test("XLSX inline editor lands on a frozen cell", async (t) => {
  const env = await mountFrozenGrid(t);
  await env.scrollTo(SCROLL_X, SCROLL_Y);

  env.container.props.onDblclick({ clientX: ROW_HEADER_WIDTH + 30, clientY: HEADER_HEIGHT + 10 });
  await vue.nextTick();
  const editor = env.find("xlsx-cell-editor");
  assert.ok(editor, "double-clicking a frozen cell opens the inline editor");
  assert.equal(editor.props.style.left, `${ROW_HEADER_WIDTH}px`);
  assert.equal(editor.props.style.top, `${HEADER_HEIGHT}px`);
  assert.equal(editor.props.style.width, `${COL_WIDTH}px`);
  assert.equal(editor.props.style.height, `${ROW_HEIGHT}px`);
});

test("XLSX inline editor pulls a cell out from under the frozen column", async (t) => {
  const env = await mountFrozenGrid(t);
  await env.scrollTo(SCROLL_X, SCROLL_Y);

  // B1 sits in the frozen row but its column is scrolled underneath column A.
  env.controller.selectCell({ row: 1, col: 1 });
  await vue.nextTick();
  env.container.props.onKeydown({ key: "F2", preventDefault() {}, stopPropagation() {} });
  await vue.nextTick();
  await env.flush();

  assert.equal(env.container.scrollLeft, 0, "editing scrolls the cell clear of the freeze line");
  const editor = env.find("xlsx-cell-editor");
  assert.ok(editor, "the inline editor opens");
  assert.equal(editor.props.style.left, `${ROW_HEADER_WIDTH + COL_WIDTH}px`);
  assert.equal(editor.props.style.top, `${HEADER_HEIGHT + ROW_HEIGHT}px`);
});

test("XLSX comment popover pins to a frozen cell", async (t) => {
  const env = await mountFrozenGrid(t);
  await env.scrollTo(SCROLL_X, SCROLL_Y);

  env.container.props.onPointermove({
    clientX: ROW_HEADER_WIDTH + 30,
    clientY: HEADER_HEIGHT + 10,
    pointerId: 1,
    pointerType: "mouse",
  });
  await vue.nextTick();
  const popover = env.find("xlsx-comment-popover");
  assert.ok(popover, "hovering the pinned comment shows the popover");
  assert.equal(popover.props.style.left, `${ROW_HEADER_WIDTH + COL_WIDTH + 6}px`);
  assert.equal(popover.props.style.top, `${HEADER_HEIGHT + 4}px`);
});

test("XLSX validation overlays never cover the frozen strip", async (t) => {
  const env = await mountFrozenGrid(t);
  await env.scrollTo(SCROLL_X, SCROLL_Y);

  env.controller.cellValidation = {
    ranges: [{ start: { row: 1, col: 2 }, end: { row: 1, col: 2 } }],
    validationType: "whole",
    showInputMessage: true,
    inputMessage: "Enter a value",
    inputTitle: "Input",
  };
  // Display column 2 starts under the pinned column at this scroll offset.
  env.controller.selectCell({ row: 1, col: 2 });
  await vue.nextTick();

  const hint = env.find("xlsx-validation-hint");
  assert.ok(hint, "the input-message hint renders");
  assert.equal(
    hint.props.style.left,
    `${ROW_HEADER_WIDTH + COL_WIDTH}px`,
    "the hint stops at the freeze line instead of covering column A",
  );
});

test("pane overlay boxes are clipped to the pane that owns the cell", () => {
  const base = {
    frozenCols: 1,
    frozenRows: 1,
    frozenWidth: COL_WIDTH,
    frozenHeight: ROW_HEIGHT,
    viewportWidth: BODY_WIDTH,
    viewportHeight: BODY_HEIGHT,
    scrollLeft: SCROLL_X,
    scrollTop: SCROLL_Y,
  };

  assert.deepEqual(
    geometry.resolvePaneOverlayRect({
      ...base, contentX: 0, contentY: 0, displayCol: 0, displayRow: 0, width: COL_WIDTH, height: ROW_HEIGHT,
    }),
    { left: 0, top: 0, width: COL_WIDTH, height: ROW_HEIGHT },
    "a frozen cell keeps its pinned box",
  );
  assert.deepEqual(
    geometry.resolvePaneOverlayRect({
      ...base, contentX: 4 * COL_WIDTH, contentY: 6 * ROW_HEIGHT, displayCol: 4, displayRow: 6, width: COL_WIDTH, height: ROW_HEIGHT,
    }),
    { left: 4 * COL_WIDTH - SCROLL_X, top: 6 * ROW_HEIGHT - SCROLL_Y, width: COL_WIDTH, height: ROW_HEIGHT },
    "a fully visible scrolling cell keeps its box",
  );
  // Cell C1 (content 160..240) starts under the pinned column at scroll 150:
  // the box snaps to the freeze line and keeps only its visible part.
  assert.deepEqual(
    geometry.resolvePaneOverlayRect({
      ...base, contentX: 2 * COL_WIDTH, contentY: 0, displayCol: 2, displayRow: 0, width: COL_WIDTH, height: ROW_HEIGHT,
    }),
    { left: COL_WIDTH, top: 0, width: 2 * COL_WIDTH - SCROLL_X + COL_WIDTH - COL_WIDTH, height: ROW_HEIGHT },
    "a half-hidden cell snaps to the freeze line",
  );
  assert.deepEqual(
    geometry.resolvePaneOverlayRect({
      ...base, contentX: 4 * COL_WIDTH, contentY: 6 * ROW_HEIGHT, displayCol: 4, displayRow: 6, width: 0, height: 0,
    }),
    { left: 4 * COL_WIDTH - SCROLL_X, top: 6 * ROW_HEIGHT - SCROLL_Y, width: 0, height: 0 },
    "point anchors keep their pane",
  );
});

test("XLSX selection outlines are drawn once per pane at the pinned coordinates", async (t) => {
  const env = await mountFrozenGrid(t);
  env.controller.selection = { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } };
  await env.scrollTo(SCROLL_X, SCROLL_Y);
  const body = env.canvasOperations("xlsx-grid__body");

  const outlines = body.filter(
    (operation) =>
      operation.name === "strokeRect" &&
      operation.args[0] === 1 &&
      operation.args[1] === 1 &&
      operation.args[2] === 2 * COL_WIDTH - 2 &&
      operation.args[3] === 2 * ROW_HEIGHT - 2 &&
      operation.strokeStyle === SELECTION_BORDER,
  );
  assert.equal(outlines.length, 1, "the outline is drawn once, at the pinned coordinates");
  assert.equal(
    body.filter(
      (operation) =>
        operation.name === "strokeRect" &&
        operation.strokeStyle === SELECTION_BORDER &&
        (operation.args[0] < 0 || operation.args[1] < 0),
    ).length,
    0,
    "no outline fragment is painted in the scrolled space",
  );
});
