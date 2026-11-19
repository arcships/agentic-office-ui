import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  keydownEvent,
  mount,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

// Only the browser plumbing the host renderer cannot provide is stubbed. The
// grid, the controller and every DOM node asserted below are real.
if (typeof Object.prototype.getContext !== "function") {
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value: () => null,
  });
}

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
// Parsed sheet features (conditional formats, comments, hyperlinks) come from
// the archive's XML, so the DOM renderer tests that read them need the parser.
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { XlsxGrid, XlsxViewer, useXlsxViewerController } = await importFromDemo(
  "@arcships/vue-xlsx",
);

const ROW_HEADER_WIDTH = 48;
const HEADER_HEIGHT = 24;
const VIEWPORT = { height: 600, width: 1_600 };

function configureWasm() {
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}

/** The demo's sales table has values, headers and a merged title. */
function workbookBuffer() {
  const bytes = readFileSync(
    new URL("../../apps/demo/public/samples/sales-table.xlsx", import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function installGridEnvironment() {
  const previous = {
    Document: globalThis.Document,
    ResizeObserver: globalThis.ResizeObserver,
    ShadowRoot: globalThis.ShadowRoot,
    document: globalThis.document,
    window: globalThis.window,
  };
  // Animation frames are queued so a test can run the grid's paint frame when
  // it needs the reactive scroll offsets and the rebuilt DOM window.
  const frames = [];
  globalThis.window = {
    addEventListener() {},
    cancelAnimationFrame() {},
    devicePixelRatio: 1,
    removeEventListener() {},
    requestAnimationFrame: (callback) => frames.push(callback),
    setInterval: globalThis.setInterval,
    setTimeout: globalThis.setTimeout,
    clearInterval: globalThis.clearInterval,
    clearTimeout: globalThis.clearTimeout,
  };
  globalThis.ResizeObserver = class ResizeObserver {
    constructor(callback) {
      this.callback = callback;
    }
    observe() {
      this.callback([{ contentRect: { height: VIEWPORT.height, width: VIEWPORT.width } }]);
    }
    unobserve() {}
    disconnect() {}
  };
  // v-model's runtime-dom directive probes `rootNode instanceof Document`.
  globalThis.Document = class Document {};
  globalThis.ShadowRoot = class ShadowRoot {};
  const flushFrames = () => {
    while (frames.length) {
      const callback = frames.shift();
      callback(0);
    }
  };
  const restore = () => {
    for (const [name, value] of Object.entries(previous)) {
      if (value === undefined) delete globalThis[name];
      else globalThis[name] = value;
    }
  };
  return { flushFrames, restore };
}

async function mountGrid(t, options = {}) {
  const { flushFrames, restore } = installGridEnvironment();
  const renderMode = vue.ref(options.renderMode ?? "dom");
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: options.file ?? workbookBuffer(),
        fileName: "dom-render.xlsx",
        skipXmlParsing: options.skipXmlParsing ?? true,
        useWorker: false,
      });
      return () => vue.h(XlsxGrid, {
        controller,
        getCellStyle: options.getCellStyle,
        readOnly: options.readOnly === true,
        renderMode: renderMode.value,
      });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "XLSX DOM-render controller ready",
  );
  assert.ifError(controller.error);
  const grid = findByTestId(mounted.root, "xlsx-grid");
  assert.ok(grid, "grid root renders");
  grid.getBoundingClientRect = () => ({
    bottom: VIEWPORT.height,
    height: VIEWPORT.height,
    left: 0,
    right: VIEWPORT.width,
    top: 0,
    width: VIEWPORT.width,
  });
  grid.style = {};
  grid.clientWidth = VIEWPORT.width;
  grid.clientHeight = VIEWPORT.height;
  grid.scrollLeft = 0;
  grid.scrollTop = 0;
  grid.focus = () => {};
  grid.setPointerCapture = () => {};
  grid.hasPointerCapture = () => true;
  grid.releasePointerCapture = () => {};
  t.after(() => mounted.app.unmount());
  t.after(restore);
  await paint({ grid, flushFrames });
  return { ...mounted, controller, flushFrames, grid, renderMode };
}

/** Runs one grid paint frame and lets Vue commit the resulting DOM window. */
async function paint({ grid, flushFrames }) {
  // The host renderer exposes handlers as props; call the grid's scroll
  // listener the way Vue would.
  grid.props.onScroll({ type: "scroll" });
  flushFrames();
  await vue.nextTick();
}

function cells(root) {
  return walk(root).filter(
    (node) => node.props?.role === "gridcell" && node.props?.["data-cell"],
  );
}

function cell(root, address) {
  return cells(root).find((node) => node.props?.["data-cell"] === address);
}

function rows(root) {
  return walk(root).filter((node) => node.props?.role === "row");
}

function canvasNodes(root) {
  return walk(root).filter((node) => node.type === "canvas");
}

function hiddenA11yRow(root) {
  return walk(root).find((node) => node.props?.class === "xlsx-grid__a11y");
}

function pointerDown(init = {}) {
  return {
    button: 0,
    pointerId: 1,
    preventDefault() {
      this.defaultPrevented = true;
    },
    stopPropagation() {
      this.propagationStopped = true;
    },
    type: "pointerdown",
    ...init,
  };
}

configureWasm();

test("renderMode='dom' replaces the canvases with semantic gridcells for the visible window", async (t) => {
  const viewer = await mountGrid(t);
  const { root, controller } = viewer;

  assert.deepEqual(viewer.warnings, [], "the DOM layer renders without Vue warnings");
  assert.equal(canvasNodes(root).length, 0, "DOM mode paints no canvases");
  const layer = findByTestId(root, "xlsx-grid-dom");
  assert.ok(layer, "the DOM layer renders");
  assert.equal(findByTestId(root, "xlsx-grid").props["data-render-mode"], "dom");

  const rendered = cells(root);
  assert.ok(rendered.length > 0, "the visible window has cells");
  assert.equal(rendered.length, rows(root).reduce((total, row) => total + row.children.filter(
    (node) => node.props?.role === "gridcell",
  ).length, 0), "every gridcell lives inside a role=row");

  const a1 = cell(root, "A1");
  assert.ok(a1, "A1 is in the window");
  assert.equal(textContent(a1), controller.getCellDisplayValue({ row: 0, col: 0 }));
  assert.equal(a1.props["aria-rowindex"], 1);
  assert.equal(a1.props["aria-colindex"], 1);
  assert.equal(a1.parent.props.role, "row");

  const columnHeaders = findByTestId(root, "xlsx-grid-dom-column-headers");
  assert.ok(columnHeaders, "column headers are DOM labels");
  assert.match(textContent(columnHeaders), /A/);
  assert.match(textContent(columnHeaders), /B/);
  const rowHeaders = findByTestId(root, "xlsx-grid-dom-row-headers");
  assert.ok(rowHeaders, "row headers are DOM labels");
  assert.match(textContent(rowHeaders), /1/);
  assert.ok(findByTestId(root, "xlsx-grid-dom-corner"), "the corner box is DOM too");

  // The active cell is a real gridcell: it owns the aria-activedescendant id
  // and the hidden mirror is not rendered twice.
  controller.selectCell({ row: 0, col: 0 });
  await vue.nextTick();
  await paint(viewer);
  const active = cell(root, "A1");
  assert.ok(active.props.id, "the active cell carries the activedescendant id");
  assert.equal(
    findByTestId(root, "xlsx-grid").props["aria-activedescendant"],
    active.props.id,
  );
  assert.equal(active.props["aria-selected"], "true");
  assert.equal(Boolean(hiddenA11yRow(root)), false);

  // Only the visible window is materialised: a row far below the viewport has
  // no element even though the sheet declares it.
  assert.equal(Boolean(cell(root, "A500")), false);
  assert.equal(Boolean(cell(root, "A2000")), false);
});

test("the DOM window scrolls with the grid and keeps real text nodes selectable", async (t) => {
  const { root, grid, flushFrames, controller } = await mountGrid(t);
  const firstText = findByTestId(root, "xlsx-grid-dom-cell");
  assert.ok(firstText, "a first cell is rendered");

  grid.scrollTop = 800;
  grid.scrollLeft = 400;
  await paint({ grid, flushFrames });

  assert.equal(Boolean(cell(root, "A2")), false, "the first scrolling row left the window");
  const pinnedRowHeaders = walk(root)
    .filter((node) => node.props?.["data-row-header"] !== undefined && node.props?.style)
    .filter((node) => node.props.style.top === "0px")
    .map((node) => textContent(node));
  assert.ok(pinnedRowHeaders.includes("1"), "row 1 lives in the frozen header strip");
  const visibleAddresses = cells(root).map((node) => node.props["data-cell"]);
  assert.ok(visibleAddresses.length > 0);
  assert.ok(
    visibleAddresses.some((address) => !/^A\d+$/.test(address)),
    "the scrolled window covers further columns",
  );
  const cellText = cells(root).find((node) => textContent(node).length > 0);
  assert.ok(cellText, "a visible cell carries its display value");
  assert.equal(
    walk(cellText).some((node) => node.type === "#text" && node.text.length > 0),
    true,
    "cell text is a real text node, so browser find and native selection see it",
  );
  const textNode = walk(cellText).find((node) => node.props?.style?.userSelect === "text");
  assert.ok(textNode, "the cell text span opts into native selection");
  assert.equal(textNode.props.style.userSelect, "text");
  assert.ok(controller.getCellDisplayValue({ row: 0, col: 0 }));
});

test("DOM cells reuse the canvas style resolution for fonts, fills, alignment and borders", async (t) => {
  const viewer = await mountGrid(t);
  const { root, controller } = viewer;
  const target = { row: 1, col: 1 };
  controller.selectRange({ start: target, end: target });
  controller.setRangeStyle(
    { start: target, end: target },
    {
      alignment: { horizontal: "right", vertical: "center", wrapText: true },
      border: {
        bottom: { color: { colorType: "rgb", hex: "FF0000" }, style: "medium" },
        top: { color: { colorType: "rgb", hex: "0000FF" }, style: "dashed" },
      },
      fill: { color: { colorType: "rgb", hex: "FFF2CC" }, fillType: "solid" },
      font: {
        bold: true,
        color: { colorType: "rgb", hex: "C00000" },
        italic: true,
        name: "Calibri",
        size: 15,
        strikethrough: true,
        underline: "single",
      },
      numberFormat: { formatString: "0.00%", formatType: "custom" },
    },
  );
  controller.setCellValue(target, "0.5");
  await vue.nextTick();
  await paint(viewer);

  const styled = cell(root, "B2");
  assert.ok(styled, "B2 is rendered");
  const style = styled.props.style;
  assert.equal(textContent(styled), "50.00%", "the DOM cell shows the formatted value");
  assert.equal(style.backgroundColor, "#fff2cc");
  assert.equal(style.color, "#c00000");
  assert.equal(style.fontWeight, "700");
  assert.equal(style.fontStyle, "italic");
  assert.equal(style.fontFamily, "\"Calibri\", -apple-system, BlinkMacSystemFont, \"Segoe UI\", sans-serif");
  assert.equal(style.fontSize, "20px", "15 pt at 96 dpi, zoomed to 100%");
  assert.equal(style.justifyContent, "flex-end");
  assert.equal(style.alignItems, "center");
  assert.equal(style.textAlign, "right");
  assert.match(style.textDecoration, /underline/);
  assert.match(style.textDecoration, /line-through/);
  assert.match(style.borderBottom, /^2px solid #ff0000$/, "medium border width and colour");
  assert.match(style.borderTop, /^1px dashed #0000ff$/);
  assert.equal(styled.props["aria-selected"], "true");
  assert.equal(style.boxShadow.includes("#d9d9d9"), true, "gridlines use the shared grid colour");

  const textStyle = walk(styled).find((node) => node.props?.style?.whiteSpace)?.props.style;
  assert.equal(textStyle.whiteSpace, "pre-wrap", "wrap text keeps the workbook's wrapping");
  assert.equal(textStyle.userSelect, "text", "the DOM text stays natively selectable");
});

function columnIndex(letters) {
  let index = 0;
  for (const character of letters) index = index * 26 + character.charCodeAt(0) - 64;
  return index - 1;
}

function addressOf(cell) {
  const match = /^([A-Z]+)(\d+)$/.exec(cell.props["data-cell"]);
  assert.ok(match, "cell has an A1 address");
  return { col: columnIndex(match[1]), row: Number(match[2]) - 1 };
}

test("DOM mode carries conditional data bars, colour scales and icon sets", async (t) => {
  const viewer = await mountGrid(t, { skipXmlParsing: false });
  const { root, controller } = viewer;
  for (let row = 5; row <= 9; row += 1) {
    for (let col = 0; col <= 2; col += 1) {
      controller.setCellValue({ col, row }, String(row - 4));
    }
  }
  const range = (col) => ({ end: { col, row: 9 }, start: { col, row: 5 } });
  assert.equal(
    controller.addConditionalFormatRule({ colors: ["638EC6"], kind: "dataBar", ranges: [range(0)] }),
    true,
  );
  assert.equal(
    controller.addConditionalFormatRule({ colors: ["F8696B", "FFEB84", "63BE7B"], kind: "colorScale", ranges: [range(1)] }),
    true,
  );
  assert.equal(
    controller.addConditionalFormatRule({ iconSet: "3Arrows", kind: "iconSet", ranges: [range(2)] }),
    true,
  );
  await vue.nextTick();
  await paint(viewer);

  const bar = walk(cell(root, "A6")).find(
    (node) => node.props?.["data-testid"] === "xlsx-grid-dom-data-bar",
  );
  assert.ok(bar, "the data bar is a DOM layer inside the cell");
  assert.equal(bar.props.style.left, "3px");
  assert.equal(bar.props.style.top, "4px");
  assert.match(bar.props.style.height, /^\d+px$/);
  assert.ok(Number.parseFloat(bar.props.style.width) > 0, "the bar has a positive width");

  const scaleCell = cell(root, "B6");
  assert.match(
    scaleCell.props.style.backgroundColor,
    /^(#|rgb\()/,
    "the colour scale fills the cell",
  );
  assert.notEqual(scaleCell.props.style.backgroundColor, "#ffffff");

  const iconCell = cell(root, "C6");
  const icon = walk(iconCell).find(
    (node) => node.props?.["data-testid"] === "xlsx-grid-dom-icon",
  );
  assert.ok(icon, "the icon-set glyph renders");
  assert.ok(["▼", "▶", "▲"].includes(textContent(icon)), "the glyph comes from the 3-arrow set");
  assert.match(
    walk(iconCell).find((node) => node.props?.style?.whiteSpace).props.style.maxWidth,
    /^\d+px$/,
    "the text column stops before the icon",
  );
});

test("host getCellStyle overrides reach the DOM cells", async (t) => {
  const viewer = await mountGrid(t, {
    getCellStyle: (cellAddress) => (cellAddress.row === 1 && cellAddress.col === 1
      ? { backgroundColor: "#123456", color: "#abcdef", fontWeight: "900", textAlign: "center" }
      : undefined),
  });
  const b2 = cell(viewer.root, "B2");
  assert.ok(b2, "B2 is rendered");
  assert.equal(b2.props.style.backgroundColor, "#123456");
  assert.equal(b2.props.style.color, "#abcdef");
  assert.equal(b2.props.style.fontWeight, "900");
  assert.equal(b2.props.style.justifyContent, "center");
});

test("an active cell outside the DOM window keeps the hidden accessibility mirror", async (t) => {
  const viewer = await mountGrid(t);
  const { root, controller, grid } = viewer;
  controller.selectCell({ row: 499, col: 0 });
  await vue.nextTick();
  await paint(viewer);

  assert.equal(Boolean(cell(root, "A500")), false, "A500 is outside the rendered window");
  const hidden = hiddenA11yRow(root);
  assert.ok(hidden, "the off-screen active cell falls back to the hidden mirror");
  const mirrorCell = hidden.children.find((node) => node.props?.role === "gridcell");
  assert.ok(mirrorCell?.props.id);
  assert.equal(grid.props["aria-activedescendant"], mirrorCell.props.id);
  assert.equal(hidden.props["aria-rowindex"], 500, "the mirror carries the off-screen cell's row");
  assert.equal(mirrorCell.props["aria-colindex"], 1);
});

test("switching renderMode keeps the scroll position and the selection", async (t) => {
  const viewer = await mountGrid(t);
  const { root, controller, grid, flushFrames, renderMode } = viewer;
  grid.scrollTop = 900;
  grid.scrollLeft = 250;
  await paint(viewer);

  // The DOM window starts at the scrolled rows, not at A1.
  assert.equal(Boolean(cell(root, "A2")), false, "DOM mode is scrolled before the switch");
  const first = cells(root)[0];
  assert.ok(first, "the scrolled window has cells");
  const start = addressOf(first);
  controller.selectRange({ start, end: { col: start.col + 1, row: start.row + 1 } });
  await vue.nextTick();
  await paint(viewer);
  const selectedAddress = first.props["data-cell"];
  const selectionBefore = {
    end: { ...controller.selection.end },
    start: { ...controller.selection.start },
  };
  assert.equal(cell(root, selectedAddress).props["aria-selected"], "true");

  renderMode.value = "canvas";
  await vue.nextTick();
  flushFrames();
  await vue.nextTick();
  assert.equal(findByTestId(root, "xlsx-grid-dom"), undefined, "canvas mode removes the DOM layer");
  assert.ok(canvasNodes(root).length >= 4, "canvas mode paints the four canvases");
  assert.deepEqual(controller.selection, selectionBefore, "selection survives DOM → canvas");
  assert.equal(grid.scrollTop, 900, "the scroller is not recreated by the switch");
  assert.equal(grid.scrollLeft, 250);

  renderMode.value = "dom";
  await vue.nextTick();
  flushFrames();
  await vue.nextTick();
  assert.equal(Boolean(cell(root, "A2")), false, "the stored scroll position survived the round trip");
  assert.deepEqual(controller.selection, selectionBefore, "selection survives canvas → DOM");
  assert.equal(
    cell(root, selectedAddress).props["aria-selected"],
    "true",
    "the selection is expressed as aria-selected gridcells in the new window",
  );
  assert.ok(
    cells(root).filter((node) => node.props["aria-selected"] === "true").length >= 4,
    "the whole selected range is highlighted",
  );
});

test("Alt+drag leaves the native text selection alone in DOM mode", async (t) => {
  const { root, controller, grid } = await mountGrid(t);
  controller.selectCell({ row: 0, col: 0 });
  await vue.nextTick();
  let captured = 0;
  grid.setPointerCapture = () => {
    captured += 1;
  };
  const a1 = {
    clientX: ROW_HEADER_WIDTH + 20,
    clientY: HEADER_HEIGHT + 10,
  };

  const plain = pointerDown(a1);
  grid.props.onPointerdown(plain);
  assert.equal(captured, 1, "a plain drag keeps the spreadsheet range selection");

  const altDrag = pointerDown({ ...a1, altKey: true });
  grid.props.onPointerdown(altDrag);
  assert.equal(captured, 1, "Alt+drag is not captured by the grid");
  assert.equal(altDrag.defaultPrevented, undefined, "the browser keeps the selection default");
});

async function mountViewer(t, options = {}) {
  const { flushFrames, restore } = installGridEnvironment();
  const renderMode = vue.ref(options.renderMode);
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(),
        fileName: "dom-viewer.xlsx",
        skipXmlParsing: true,
        useWorker: false,
      });
      return () => vue.h(XlsxViewer, {
        controller,
        renderMode: renderMode.value,
        showDefaultToolbar: false,
        showFormulaBar: false,
        showImages: false,
        showRibbon: false,
      });
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "XLSX viewer controller ready",
  );
  assert.ifError(controller.error);
  const grid = findByTestId(mounted.root, "xlsx-grid");
  assert.ok(grid, "viewer renders the grid");
  grid.getBoundingClientRect = () => ({
    bottom: VIEWPORT.height,
    height: VIEWPORT.height,
    left: 0,
    right: VIEWPORT.width,
    top: 0,
    width: VIEWPORT.width,
  });
  grid.style = {};
  grid.clientWidth = VIEWPORT.width;
  grid.clientHeight = VIEWPORT.height;
  grid.scrollLeft = 0;
  grid.scrollTop = 0;
  t.after(() => mounted.app.unmount());
  t.after(restore);
  await paint({ grid, flushFrames });
  return { ...mounted, controller, flushFrames, grid, renderMode };
}

test("Ctrl/Cmd+C over native selected text is left to the browser in DOM mode", async (t) => {
  const viewer = await mountGrid(t);
  viewer.controller.selectCell({ row: 0, col: 0 });
  await vue.nextTick();
  const textNode = walk(cells(viewer.root)[0]).find((node) => node.type === "#text");
  assert.ok(textNode, "a cell text node is available for a native selection");
  viewer.grid.contains = (node) => node === textNode;
  globalThis.getSelection = () => ({ anchorNode: textNode, isCollapsed: false, rangeCount: 1 });
  try {
    const { event, preventedCount } = keydownEvent("c", { ctrlKey: true });
    viewer.grid.props.onKeydown(event);
    assert.equal(preventedCount(), 0, "the browser copy command keeps the selected text");

    const plain = keydownEvent("c", { ctrlKey: true, target: viewer.root });
    globalThis.getSelection = () => ({ anchorNode: null, isCollapsed: true, rangeCount: 0 });
    viewer.grid.props.onKeydown(plain.event);
    assert.equal(plain.preventedCount(), 1, "without a text selection Ctrl+C still copies the cell range");
  } finally {
    delete globalThis.getSelection;
    delete viewer.grid.contains;
  }
});

test("XlsxViewer keeps canvas as the default and forwards renderMode", async (t) => {
  const viewer = await mountViewer(t);
  assert.equal(viewer.grid.props["data-render-mode"], "canvas", "the default renderer is canvas");
  assert.equal(Boolean(findByTestId(viewer.root, "xlsx-grid-dom")), false);
  assert.ok(canvasNodes(viewer.root).length >= 4);

  viewer.renderMode.value = "dom";
  await vue.nextTick();
  viewer.flushFrames();
  await vue.nextTick();
  assert.equal(viewer.grid.props["data-render-mode"], "dom");
  assert.ok(findByTestId(viewer.root, "xlsx-grid-dom"), "the viewer passes renderMode to the grid");
  assert.ok(cells(viewer.root).length > 0);
});
