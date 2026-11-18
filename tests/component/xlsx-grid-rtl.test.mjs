import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const { XlsxGrid } = await importFromDemo("@arcships/vue-xlsx");

const HEADER_HEIGHT = 24;
const ROW_HEADER_WIDTH = 48;
const COL_WIDTH = 80;
const ROW_HEIGHT = 24;
const BODY_WIDTH = 592;
const BODY_HEIGHT = 456;
const SCROLL_X = 150;
// The grid fills sheets up to at least 50 display columns.
const LAYOUT_WIDTH = 50 * COL_WIDTH;
const SELECTION_FILL = "#123456";

function rightToLeftSheet(overrides = {}) {
  return {
    name: "Sheet1",
    visibility: "visible",
    workbookSheetIndex: 0,
    rightToLeft: true,
    freezePanes: { row: 1, col: 1 },
    visibleRows: Array.from({ length: 40 }, (_, index) => index),
    visibleCols: Array.from({ length: 20 }, (_, index) => index),
    rowHeights: Array.from({ length: 40 }, () => ROW_HEIGHT),
    colWidths: Array.from({ length: 20 }, () => COL_WIDTH),
    rowHeightOverridesPx: {},
    colWidthOverridesPx: {},
    rowStyleIds: {},
    colStyleIds: {},
    styleById: {},
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
    maxUsedCol: 5,
    maxUsedRow: 1,
    rowCount: 40,
    colCount: 20,
    showGridLines: true,
    mergedRegions: [],
    hyperlinks: [],
    comments: [],
    dataValidations: [],
    sparklines: [],
    conditionalFormatRules: [],
    ...overrides,
    values: { "0:0": "A1", "0:1": "B1", "0:5": "F1", "1:0": "A2" },
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
    commitCellInput() {},
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

/** Recording 2D context: draw calls are tagged with the canvas and final transform. */
function createRecordingContext(canvas, operations) {
  const style = { fillStyle: "#000000", strokeStyle: "#000000", font: "10px sans-serif", textAlign: "start", textBaseline: "alphabetic", lineWidth: 1 };
  const snapshot = () => ({ ...style });
  const canvasName = () => String(canvas.props?.class ?? canvas.type)
    .split(/\s+/)
    .find((name) => name.startsWith("xlsx-grid__")) ?? String(canvas.type);
  let transform = { x: 0, y: 0 };
  let clip = null;
  let pathRects = [];
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
          return () => { pathRects = []; };
        case "rect":
          return (x, y, width, height) => { pathRects.push(deviceRect(x, y, width, height)); };
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
          return () => record(property, []);
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
      operations.length = 0;
      for (const callback of pending) callback();
      await vue.nextTick();
    }
  };

  return { operations, flush, cleanup };
}

async function mountRightToLeftGrid(t, sheetOverrides = {}) {
  const environment = installBrowserEnvironment(t);
  const sheet = rightToLeftSheet(sheetOverrides);
  const controller = createController(sheet);
  let api;
  const Harness = vue.defineComponent({
    setup() {
      const gridRef = vue.ref(null);
      vue.watch(gridRef, (value) => { api = value; }, { flush: "sync" });
      return () => vue.h(XlsxGrid, {
        ref: gridRef,
        controller,
        selectionFillColor: SELECTION_FILL,
      });
    },
  });
  const mounted = await mount(Harness);
  t.after(() => {
    mounted.app.unmount();
    environment.cleanup();
  });
  for (const node of walk(mounted.root)) {
    if (node.type === "canvas") node.style = {};
  }

  const treeContainer = findByTestId(mounted.root, "xlsx-grid");
  assert.ok(treeContainer, "grid container must render");
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

  const scrollTo = async (scrollLeft, scrollTop = 0) => {
    container.scrollLeft = scrollLeft;
    container.scrollTop = scrollTop;
    environment.operations.length = 0;
    container.props.onScroll?.({ target: container });
    await environment.flush();
  };
  const canvasOperations = (className) => {
    const canvas = walk(mounted.root).find(
      (node) => node.type === "canvas" && String(node.props?.class).includes(className),
    );
    assert.ok(canvas, `missing ${className} canvas`);
    return environment.operations.filter((operation) => operation.canvas === className);
  };
  return { mounted, container, controller, api, sheet, canvasOperations, scrollTo };
}

test("an RTL sheet paints column A at the right edge", async (t) => {
  const env = await mountRightToLeftGrid(t);
  const grid = findByTestId(env.mounted.root, "xlsx-grid");
  assert.equal(grid.props["data-right-to-left"], "true", "the grid exposes the sheet direction");

  const labels = env.canvasOperations("xlsx-grid__col-header")
    .filter((operation) => operation.name === "fillText");
  const a = labels.find((operation) => operation.args[0] === "A");
  const b = labels.find((operation) => operation.args[0] === "B");
  const f = labels.find((operation) => operation.args[0] === "F");
  assert.ok(a && b && f, "A, B and F headers are painted");
  assert.equal(a.args[1], LAYOUT_WIDTH - COL_WIDTH + COL_WIDTH / 2, "column A sits above the pinned right-hand column");
  assert.equal(b.args[1], LAYOUT_WIDTH - 2 * COL_WIDTH + COL_WIDTH / 2, "column B is to its left");
  assert.ok(a.args[1] > b.args[1] && b.args[1] > f.args[1], "headers run right to left");

  const body = env.canvasOperations("xlsx-grid__body");
  const a1 = body.find((operation) => operation.name === "fillText" && operation.args[0] === "A1");
  const b1 = body.find((operation) => operation.name === "fillText" && operation.args[0] === "B1");
  assert.ok(a1 && b1, "A1 and B1 are painted");
  assert.ok(a1.args[1] > b1.args[1], "A1 is to the right of B1");
  assert.equal(a1.args[1], LAYOUT_WIDTH - 5, "right-aligned general text starts at the right edge");
});

test("RTL hit testing selects the cell under the pointer", async (t) => {
  const env = await mountRightToLeftGrid(t);

  assert.deepEqual(
    env.api.hitTestCell(ROW_HEADER_WIDTH + BODY_WIDTH - 10, HEADER_HEIGHT + 10),
    { row: 0, col: 0 },
    "the right-hand pinned cell is A1",
  );
  assert.deepEqual(
    env.api.hitTestCell(ROW_HEADER_WIDTH + BODY_WIDTH - COL_WIDTH - 10, HEADER_HEIGHT + 10),
    { row: 0, col: 1 },
    "the next cell to the left is B1",
  );
  assert.deepEqual(
    env.api.hitTestCell(ROW_HEADER_WIDTH + BODY_WIDTH - 10, HEADER_HEIGHT + ROW_HEIGHT + 10),
    { row: 1, col: 0 },
    "rows are not mirrored",
  );
  assert.deepEqual(
    env.api.hitTestAxis(ROW_HEADER_WIDTH + BODY_WIDTH - 10, 10),
    { actualIndex: 0, displayIndex: 0, kind: "column", resize: false },
    "column A owns the right-hand header",
  );
  const point = env.api.hitTestSheetPoint(
    ROW_HEADER_WIDTH + BODY_WIDTH - 10,
    HEADER_HEIGHT + 10,
  );
  assert.equal(point?.col, 0);
  assert.equal(point?.xOffset, 0.875, "the fraction is measured from the visual left edge");
});

test("RTL frozen panes pin column A while the other columns scroll", async (t) => {
  const env = await mountRightToLeftGrid(t);
  await env.scrollTo(SCROLL_X);
  const labels = env.canvasOperations("xlsx-grid__col-header")
    .filter((operation) => operation.name === "fillText");
  const a = labels.find((operation) => operation.args[0] === "A");
  const f = labels.find((operation) => operation.args[0] === "F");
  assert.ok(a && f, "A stays pinned and F scrolls into view");
  assert.equal(a.args[1], LAYOUT_WIDTH - COL_WIDTH + COL_WIDTH / 2, "the pinned label ignores scrollLeft");
  assert.equal(
    f.args[1],
    LAYOUT_WIDTH - 6 * COL_WIDTH + COL_WIDTH / 2,
    "F follows the mirrored scroll",
  );

  const body = env.canvasOperations("xlsx-grid__body");
  const a1 = body.find((operation) => operation.name === "fillText" && operation.args[0] === "A1");
  const f1 = body.find((operation) => operation.name === "fillText" && operation.args[0] === "F1");
  assert.ok(a1 && f1, "the pinned and scrolled cells are painted");
  assert.equal(a1.args[1], LAYOUT_WIDTH - 5, "A1 stays at the right edge");
  assert.equal(
    f1.args[1],
    LAYOUT_WIDTH - 6 * COL_WIDTH + COL_WIDTH - 5,
    "F1 follows the mirrored scroll",
  );

  assert.deepEqual(
    env.api.hitTestCell(ROW_HEADER_WIDTH + BODY_WIDTH - 10, HEADER_HEIGHT + 10),
    { row: 0, col: 0 },
    "clicks on the pinned strip still select A1",
  );
  const fCenter = ROW_HEADER_WIDTH + (BODY_WIDTH - 5 * COL_WIDTH - COL_WIDTH + SCROLL_X) + COL_WIDTH / 2;
  assert.deepEqual(
    env.api.hitTestCell(fCenter, HEADER_HEIGHT + 10),
    { row: 0, col: 5 },
    "clicks in the scrolling pane select the mirrored column",
  );
});

test("scrollToCell keeps a pinned column visible and scrolls the rest", async (t) => {
  const env = await mountRightToLeftGrid(t);
  await env.scrollTo(SCROLL_X);

  env.api.scrollToCell({ row: 0, col: 1 });
  await vue.nextTick();
  assert.equal(env.container.scrollLeft, 0, "column B is pulled out from under the pinned column A");

  env.api.scrollToCell({ row: 0, col: 20 });
  await vue.nextTick();
  assert.equal(
    env.container.scrollLeft,
    20 * COL_WIDTH + COL_WIDTH - BODY_WIDTH,
    "a later column scrolls until its right edge clears the freeze line",
  );
});

test("the selection outline and fill handle follow the mirrored corner", async (t) => {
  const env = await mountRightToLeftGrid(t);
  env.controller.selection = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };
  await env.scrollTo(0);

  const fills = env.canvasOperations("xlsx-grid__body")
    .filter((operation) => operation.name === "fillRect" && operation.fillStyle === SELECTION_FILL);
  assert.deepEqual(
    fills.map((operation) => [operation.args[0], operation.args[1]]),
    [[LAYOUT_WIDTH - COL_WIDTH + 1, 1]],
    "A1's selection tint starts at the pinned right-hand column",
  );

  const handle = walk(env.mounted.root).find(
    (node) => node.props?.["data-testid"] === "xlsx-fill-handle",
  );
  assert.ok(handle, "the selection renders its fill handle");
  assert.equal(
    handle.props.style.left,
    `${ROW_HEADER_WIDTH + BODY_WIDTH - COL_WIDTH}px`,
    "the handle is clipped at the right-hand freeze line",
  );
});
