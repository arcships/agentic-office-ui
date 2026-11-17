import assert from "node:assert/strict";
import { test } from "node:test";

import { importFromDemo, mount, textContent, vue, walk, waitFor } from "./vue-test-renderer.mjs";

const { XlsxChartOverlay, XlsxDrawingLayer, XlsxImageLayer } = await importFromDemo("@arcships/vue-xlsx");

const PX_TO_EMU = 9525;

function sheetData(overrides = {}) {
  const visibleCols = Array.from({ length: 10 }, (_, index) => index);
  const visibleRows = Array.from({ length: 20 }, (_, index) => index);
  return {
    colWidthOverridesPx: {},
    colWidths: visibleCols.map(() => 80),
    defaultColWidthPx: 80,
    defaultRowHeightPx: 24,
    freezePanes: null,
    name: "Sheet1",
    rowHeightOverridesPx: {},
    rowHeights: visibleRows.map(() => 24),
    showGridLines: true,
    visibleCols,
    visibleRows,
    ...overrides,
  };
}

function absoluteAnchor(xPx, yPx, widthPx = 120, heightPx = 80) {
  return {
    kind: "absolute",
    positionEmu: { x: xPx * PX_TO_EMU, y: yPx * PX_TO_EMU },
    sizeEmu: { cx: widthPx * PX_TO_EMU, cy: heightPx * PX_TO_EMU },
  };
}

function image(id, xPx, yPx, overrides = {}) {
  return {
    anchor: absoluteAnchor(xPx, yPx),
    id,
    mimeType: "image/png",
    name: id,
    sheetIndex: 0,
    src: `data:image/png;base64,${id}`,
    workbookSheetIndex: 0,
    zIndex: 0,
    ...overrides,
  };
}

function imageController(overrides = {}) {
  const calls = { delete: [], move: [], resize: [], select: [], shapeMove: [], shapeResize: [] };
  const controller = {
    activeSheet: sheetData(),
    charts: [],
    clearSelectedChart() {},
    clearSelectedImage() {},
    clearSelectedShape() {},
    deleteChart(id) { calls.delete.push(["chart", id]); },
    deleteImage(id) { calls.delete.push(["image", id]); },
    deleteShape(id) { calls.delete.push(["shape", id]); },
    formControls: [],
    images: [],
    isChartsLoading: false,
    maxConcurrentImageDecodes: 4,
    moveChartBy() {},
    moveImageBy(id, dx, dy) { calls.move.push([id, dx, dy]); },
    moveShapeBy(id, dx, dy) { calls.shapeMove.push([id, dx, dy]); },
    readOnly: false,
    reportImageDecodeError() {},
    resizeChartBy() {},
    resizeImageBy(id, handle, dx, dy) { calls.resize.push([id, handle, dx, dy]); },
    resizeShapeBy(id, handle, dx, dy) { calls.shapeResize.push([id, handle, dx, dy]); },
    selectedChartElement: null,
    selectedChartId: null,
    selectedImageId: null,
    selectedShapeId: null,
    selectChart(id) { this.selectedChartId = id; },
    selectImage(id) { calls.select.push(id); this.selectedImageId = id; },
    selectShape(id) { this.selectedShapeId = id; },
    setCellValue() {},
    shapes: [],
    zoomScale: 100,
    ...overrides,
  };
  return { calls, controller };
}

function installWindowListeners() {
  const listeners = new Map();
  const previous = globalThis.window;
  globalThis.window = {
    addEventListener(type, listener) { listeners.set(type, listener); },
    removeEventListener(type, listener) {
      if (listeners.get(type) === listener) listeners.delete(type);
    },
  };
  return {
    listeners,
    restore() {
      globalThis.window = previous;
    },
  };
}

function pointerEvent(overrides = {}) {
  return {
    button: 0,
    clientX: 100,
    clientY: 50,
    currentTarget: { focus() {}, setPointerCapture() {} },
    pointerId: 1,
    preventDefault() {},
    stopPropagation() {},
    ...overrides,
  };
}

function findItems(root, testId) {
  return walk(root).filter((node) => node.props?.["data-testid"] === testId);
}

test("image items honour workbook zIndex, rotation, flip and srcRect crop", async () => {
  const { controller } = imageController({
    images: [
      image("front", 0, 0, {
        flipH: true,
        rotationDeg: 45,
        srcRect: { bottom: 0, left: 0.1, right: 0.1, top: 0.2 },
        zIndex: 7,
      }),
      image("back", 200, 0, { zIndex: 2 }),
    ],
    selectedImageId: "front",
  });
  const mounted = await mount(XlsxImageLayer, { controller, showImages: true });
  try {
    const [front, back] = findItems(mounted.root, "xlsx-image-item");
    assert.equal(front.props.style.zIndex, 7, "selected state must not replace the parsed z-order");
    assert.equal(back.props.style.zIndex, 2);
    assert.equal(front.props.style.transform, "rotate(45deg) scaleX(-1)");

    const img = walk(front).find((node) => node.type === "img");
    assert.ok(img, "the cropped image should render");
    assert.equal(img.props.style.width, "125%");
    assert.equal(img.props.style.left, "-12.5%");
    assert.equal(img.props.style.height, "125%");
    assert.equal(img.props.style.top, "-25%");
  } finally {
    mounted.app.unmount();
  }
});

test("dragging an image previews locally and commits exactly one move", async () => {
  const { calls, controller } = imageController({ images: [image("photo", 100, 50)] });
  const windowMock = installWindowListeners();
  const mounted = await mount(XlsxImageLayer, { controller, showImages: true });
  try {
    const [item] = findItems(mounted.root, "xlsx-image-item");
    item.props.onPointerdown(pointerEvent());
    assert.deepEqual(calls.select, ["photo"], "pointerdown selects the object");

    windowMock.listeners.get("pointermove")({ clientX: 140, clientY: 70, pointerId: 1 });
    await vue.nextTick();
    assert.equal(item.props.style.left, "140px", "drag preview moves without touching the workbook");
    assert.equal(item.props.style.top, "70px");
    assert.deepEqual(calls.move, [], "no controller mutation while the pointer is down");

    windowMock.listeners.get("pointerup")({ clientX: 140, clientY: 70, pointerId: 1 });
    assert.deepEqual(calls.move, [["photo", 40, 20]]);
    assert.deepEqual(calls.resize, []);
    assert.equal(windowMock.listeners.has("pointermove"), false, "gesture listeners are released");
  } finally {
    mounted.app.unmount();
    windowMock.restore();
  }
});

test("selected images expose eight resize handles and commit one resize", async () => {
  const { calls, controller } = imageController({
    images: [image("photo", 100, 50)],
    selectedImageId: "photo",
  });
  const windowMock = installWindowListeners();
  const mounted = await mount(XlsxImageLayer, { controller, showImages: true });
  try {
    const [item] = findItems(mounted.root, "xlsx-image-item");
    const handles = findItems(item, "xlsx-image-resize-handle");
    assert.equal(handles.length, 8);
    assert.deepEqual(
      handles.map((node) => node.props["data-handle"]).sort(),
      ["e", "n", "ne", "nw", "s", "se", "sw", "w"],
    );

    const handle = handles.find((node) => node.props["data-handle"] === "se");
    handle.props.onPointerdown(pointerEvent({ clientX: 220, clientY: 130 }));
    windowMock.listeners.get("pointermove")({ clientX: 260, clientY: 170, pointerId: 1 });
    windowMock.listeners.get("pointerup")({ clientX: 260, clientY: 170, pointerId: 1 });
    assert.deepEqual(calls.resize, [["photo", "se", 40, 40]]);
    assert.deepEqual(calls.move, []);
  } finally {
    mounted.app.unmount();
    windowMock.restore();
  }
});

test("arrow keys nudge the selected image (shift = 10px)", async () => {
  const { calls, controller } = imageController({
    images: [image("photo", 100, 50)],
    selectedImageId: "photo",
  });
  const mounted = await mount(XlsxImageLayer, { controller, showImages: true });
  try {
    const [item] = findItems(mounted.root, "xlsx-image-item");
    let prevented = false;
    item.props.onKeydown({
      key: "ArrowRight",
      shiftKey: true,
      preventDefault() { prevented = true; },
      stopPropagation() {},
    });
    assert.equal(prevented, true);
    assert.deepEqual(calls.move, [["photo", 10, 0]]);
  } finally {
    mounted.app.unmount();
  }
});

test("selected charts expose handles and share the drag/resize gesture", async () => {
  const { calls, controller } = imageController({
    charts: [{
      anchor: absoluteAnchor(100, 50, 200, 120),
      chartType: "Unsupported",
      id: "chart-1",
      name: "Chart",
      series: [],
      sheetIndex: 0,
      workbookSheetIndex: 0,
    }],
    selectedChartId: "chart-1",
  });
  controller.moveChartBy = (id, dx, dy) => calls.move.push([id, dx, dy]);
  controller.resizeChartBy = (id, handle, dx, dy) => calls.resize.push([id, handle, dx, dy]);
  const windowMock = installWindowListeners();
  const mounted = await mount(XlsxChartOverlay, { controller });
  try {
    const [item] = findItems(mounted.root, "xlsx-chart-item");
    assert.equal(findItems(item, "xlsx-chart-resize-handle").length, 8);
    item.props.onPointerdown(pointerEvent({ clientX: 120, clientY: 60 }));
    windowMock.listeners.get("pointermove")({ clientX: 135, clientY: 75, pointerId: 1 });
    windowMock.listeners.get("pointerup")({ clientX: 135, clientY: 75, pointerId: 1 });
    assert.deepEqual(calls.move, [["chart-1", 15, 15]]);
  } finally {
    mounted.app.unmount();
    windowMock.restore();
  }
});

test("chart SVG still renders through the memoised renderer", async () => {
  const chart = {
    anchor: absoluteAnchor(100, 50, 220, 140),
    axes: [],
    chartType: "ColumnClustered",
    editable: true,
    id: "chart-1",
    name: "Revenue",
    series: [{
      categories: ["Q1", "Q2"],
      id: "series-1",
      name: "Revenue",
      values: [10, 20],
    }],
    sheetIndex: 0,
    workbookSheetIndex: 0,
  };
  const { controller } = imageController({ charts: [chart] });
  const scrollLeft = vue.ref(0);
  const Wrapper = vue.defineComponent({
    name: "ChartOverlayScrollHarness",
    setup() {
      return () => vue.h(XlsxChartOverlay, { controller, scrollLeft: scrollLeft.value });
    },
  });
  const mounted = await mount(Wrapper);
  try {
    await waitFor(
      () => walk(mounted.root).find((node) => node.type === "svg"),
      "the chart SVG to render",
    );
    // A scroll frame recreates the rect object but not its size; the memoised
    // SVG must survive the parent update without warnings or a remount.
    scrollLeft.value = 40;
    await vue.nextTick();
    assert.ok(walk(mounted.root).find((node) => node.type === "svg"));
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("shape text renders run styling, paragraph alignment, insets, flip and zIndex", async () => {
  const shape = {
    anchor: absoluteAnchor(100, 50, 160, 80),
    flipV: true,
    geometry: "roundRect",
    id: "shape-1",
    name: "Note",
    paragraphs: [{
      align: "right",
      runs: [{
        bold: true,
        color: "#ff0000",
        fontSizePt: 12,
        italic: true,
        text: "Review",
        underline: true,
      }],
    }],
    rotationDeg: 30,
    sheetIndex: 0,
    textBox: {
      insetPx: { bottom: 4, left: 6, right: 6, top: 4 },
      verticalAlign: "middle",
    },
    workbookSheetIndex: 0,
    zIndex: 4,
  };
  const { controller } = imageController({ shapes: [shape] });
  const mounted = await mount(XlsxDrawingLayer, { controller });
  try {
    const [node] = findItems(mounted.root, "xlsx-shape");
    assert.equal(node.props.style.zIndex, 4);
    assert.equal(node.props.style.transform, "rotate(30deg) scaleY(-1)");

    const paragraph = walk(node).find((entry) => entry.type === "p");
    assert.equal(paragraph.props.style.textAlign, "right");
    const span = walk(node).find((entry) => entry.type === "span");
    assert.equal(textContent(span), "Review");
    assert.equal(span.props.style.fontWeight, "700");
    assert.equal(span.props.style.fontStyle, "italic");
    assert.equal(span.props.style.color, "#ff0000");
    assert.equal(span.props.style.fontSize, "16px");
    assert.equal(span.props.style.textDecoration, "underline");

    const textBox = walk(node).find((entry) => entry.props?.style?.padding);
    assert.equal(textBox.props.style.padding, "4px 6px 4px 6px");
    assert.equal(textBox.props.style.justifyContent, "center");
  } finally {
    mounted.app.unmount();
  }
});

test("checkbox controls toggle and write TRUE/FALSE to their linked cell", async () => {
  const setCellValueCalls = [];
  const { controller } = imageController({
    formControls: [{
      anchor: absoluteAnchor(100, 50, 120, 24),
      checked: false,
      id: "control-1",
      kind: "checkbox",
      label: "Approve",
      linkedCell: "$B$2",
      name: "Check Box 1",
      sheetIndex: 0,
      workbookSheetIndex: 0,
      zIndex: 1,
    }],
    setCellValue(cell, value) { setCellValueCalls.push([cell, value]); },
  });
  const mounted = await mount(XlsxDrawingLayer, { controller });
  try {
    const [control] = findItems(mounted.root, "xlsx-form-control");
    assert.equal(control.props["data-control-checked"], "false");
    control.props.onClick({ stopPropagation() {} });
    await vue.nextTick();
    assert.deepEqual(setCellValueCalls, [[{ col: 1, row: 1 }, "TRUE"]]);
    assert.equal(control.props["data-control-checked"], "true");

    control.props.onClick({ stopPropagation() {} });
    await vue.nextTick();
    assert.deepEqual(setCellValueCalls.at(-1), [{ col: 1, row: 1 }, "FALSE"]);
    assert.equal(control.props["data-control-checked"], "false");
  } finally {
    mounted.app.unmount();
  }
});

test("radio controls sharing a linked cell clear their siblings", async () => {
  const base = {
    anchor: absoluteAnchor(100, 50, 120, 24),
    kind: "radio",
    sheetIndex: 0,
    workbookSheetIndex: 0,
    zIndex: 1,
  };
  const { controller } = imageController({
    formControls: [
      { ...base, checked: true, id: "radio-1", name: "Option 1", linkedCell: "$B$2" },
      { ...base, checked: false, id: "radio-2", name: "Option 2", linkedCell: "$B$2" },
    ],
  });
  const mounted = await mount(XlsxDrawingLayer, { controller });
  try {
    const controls = findItems(mounted.root, "xlsx-form-control");
    assert.deepEqual(controls.map((node) => node.props["data-control-checked"]), ["true", "false"]);
    controls[1].props.onClick({ stopPropagation() {} });
    await vue.nextTick();
    assert.deepEqual(controls.map((node) => node.props["data-control-checked"]), ["false", "true"]);
  } finally {
    mounted.app.unmount();
  }
});

function shape(id, xPx, yPx, overrides = {}) {
  return {
    anchor: absoluteAnchor(xPx, yPx, 160, 80),
    geometry: "rect",
    id,
    name: id,
    paragraphs: [{ runs: [{ text: id }] }],
    sheetIndex: 0,
    workbookSheetIndex: 0,
    zIndex: 4,
    ...overrides,
  };
}

test("Delete removes selected images, charts and shapes through the controller", async () => {
  const imageSetup = imageController({ images: [image("photo", 100, 50)], selectedImageId: "photo" });
  const imageMount = await mount(XlsxImageLayer, { controller: imageSetup.controller, showImages: true });
  try {
    const [item] = findItems(imageMount.root, "xlsx-image-item");
    item.props.onKeydown({ key: "Delete", preventDefault() {}, stopPropagation() {} });
    assert.deepEqual(imageSetup.calls.delete, [["image", "photo"]]);
  } finally {
    imageMount.app.unmount();
  }

  const chartSetup = imageController({
    charts: [{ anchor: absoluteAnchor(100, 50, 200, 120), chartType: "Unsupported", id: "chart-1", series: [], sheetIndex: 0, workbookSheetIndex: 0 }],
    selectedChartId: "chart-1",
  });
  const chartMount = await mount(XlsxChartOverlay, { controller: chartSetup.controller });
  try {
    const [item] = findItems(chartMount.root, "xlsx-chart-item");
    item.props.onKeydown({ key: "Delete", preventDefault() {}, stopPropagation() {} });
    assert.deepEqual(chartSetup.calls.delete, [["chart", "chart-1"]]);
  } finally {
    chartMount.app.unmount();
  }

  const shapeSetup = imageController({ selectedShapeId: "shape-1", shapes: [shape("shape-1", 100, 50)] });
  const shapeMount = await mount(XlsxDrawingLayer, { controller: shapeSetup.controller });
  try {
    const [item] = findItems(shapeMount.root, "xlsx-shape");
    item.props.onKeydown({ key: "Delete", preventDefault() {}, stopPropagation() {} });
    assert.deepEqual(shapeSetup.calls.delete, [["shape", "shape-1"]]);
  } finally {
    shapeMount.app.unmount();
  }
});

test("selected shapes drag, resize and nudge through the shape mutation APIs", async () => {
  const { calls, controller } = imageController({
    selectedShapeId: "shape-1",
    shapes: [shape("shape-1", 100, 50)],
  });
  const windowMock = installWindowListeners();
  const mounted = await mount(XlsxDrawingLayer, { controller });
  try {
    const [item] = findItems(mounted.root, "xlsx-shape");
    item.props.onPointerdown(pointerEvent());
    windowMock.listeners.get("pointermove")({ clientX: 130, clientY: 70, pointerId: 1 });
    await vue.nextTick();
    assert.equal(item.props.style.left, "130px");
    assert.deepEqual(calls.shapeMove, [], "shape preview must not mutate the workbook");
    windowMock.listeners.get("pointerup")({ clientX: 130, clientY: 70, pointerId: 1 });
    assert.deepEqual(calls.shapeMove, [["shape-1", 30, 20]]);

    const handles = findItems(item, "xlsx-shape-resize-handle");
    assert.equal(handles.length, 8);
    handles.find((node) => node.props["data-handle"] === "se")
      .props.onPointerdown(pointerEvent({ clientX: 200, clientY: 120 }));
    windowMock.listeners.get("pointermove")({ clientX: 225, clientY: 140, pointerId: 1 });
    windowMock.listeners.get("pointerup")({ clientX: 225, clientY: 140, pointerId: 1 });
    assert.deepEqual(calls.shapeResize, [["shape-1", "se", 25, 20]]);

    item.props.onKeydown({ key: "ArrowLeft", shiftKey: false, preventDefault() {}, stopPropagation() {} });
    assert.deepEqual(calls.shapeMove.at(-1), ["shape-1", -1, 0]);
  } finally {
    mounted.app.unmount();
    windowMock.restore();
  }
});

test("overlay layers cull anchored objects outside the visible viewport", async () => {
  const imageSetup = imageController({ images: [image("visible", 0, 0), image("offscreen", 1200, 900)] });
  const imageMount = await mount(XlsxImageLayer, {
    controller: imageSetup.controller,
    showImages: true,
    viewportHeight: 150,
    viewportWidth: 200,
  });
  try {
    assert.deepEqual(
      findItems(imageMount.root, "xlsx-image-item").map((node) => walk(node).find((child) => child.type === "img")?.props.alt),
      ["visible"],
    );
  } finally {
    imageMount.app.unmount();
  }

  const chartSetup = imageController({
    charts: [
      { anchor: absoluteAnchor(0, 0, 120, 80), chartType: "Unsupported", id: "chart-visible", series: [], sheetIndex: 0, workbookSheetIndex: 0 },
      { anchor: absoluteAnchor(1500, 900, 120, 80), chartType: "Unsupported", id: "chart-offscreen", series: [], sheetIndex: 0, workbookSheetIndex: 0 },
    ],
  });
  const chartMount = await mount(XlsxChartOverlay, {
    controller: chartSetup.controller,
    viewportHeight: 150,
    viewportWidth: 200,
  });
  try {
    assert.equal(findItems(chartMount.root, "xlsx-chart-item").length, 1);
  } finally {
    chartMount.app.unmount();
  }

  const shapeSetup = imageController({ shapes: [shape("shape-visible", 0, 0), shape("shape-offscreen", 1500, 900)] });
  const shapeMount = await mount(XlsxDrawingLayer, {
    controller: shapeSetup.controller,
    viewportHeight: 150,
    viewportWidth: 200,
  });
  try {
    assert.equal(findItems(shapeMount.root, "xlsx-shape").length, 1);
    assert.equal(findItems(shapeMount.root, "xlsx-shape")[0].props["data-shape-id"], "shape-visible");
  } finally {
    shapeMount.app.unmount();
  }
});
