import assert from "node:assert/strict";
import { test } from "node:test";

import { importFromDemo, mount, walk } from "./vue-test-renderer.mjs";

const { XlsxImageLayer } = await importFromDemo("@arcships/vue-xlsx");

const PX_TO_EMU = 9525;

function absoluteImage(id, xPx, yPx) {
  return {
    id,
    name: id,
    mimeType: "image/png",
    sheetIndex: 0,
    workbookSheetIndex: 0,
    src: `data:image/png;base64,${id}`,
    zIndex: 0,
    anchor: {
      kind: "absolute",
      positionEmu: { x: xPx * PX_TO_EMU, y: yPx * PX_TO_EMU },
      sizeEmu: { cx: 120 * PX_TO_EMU, cy: 80 * PX_TO_EMU },
    },
  };
}

test("XLSX image layer clips anchored objects per freeze pane and only scrolls the scrolling pane", async () => {
  const controller = {
    selectedImageId: null,
    zoomScale: 100,
    activeSheet: {
      name: "Sheet1",
      freezePanes: { row: 2, col: 0 },
      visibleRows: Array.from({ length: 20 }, (_, index) => index),
      visibleCols: Array.from({ length: 10 }, (_, index) => index),
      rowHeights: [],
      colWidths: [],
      defaultRowHeightPx: 24,
      defaultColWidthPx: 80,
      colWidthOverridesPx: {},
      rowHeightOverridesPx: {},
    },
    // 一张锚在冻结行区（y=0），一张锚在滚动区（y=240px）。
    images: [absoluteImage("frozen", 0, 0), absoluteImage("scrolling", 100, 240)],
    selectImage() {},
  };
  const mounted = await mount(XlsxImageLayer, {
    controller,
    showImages: true,
    scrollLeft: 30,
    scrollTop: 100,
  });
  try {
    const nodes = walk(mounted.root);
    const panes = nodes.filter((node) => node.props?.["data-grid-pane"]);
    assert.deepEqual(
      panes.map((pane) => pane.props["data-grid-pane"]).sort(),
      ["frozen-rows", "main"],
      "freezePanes {row:2, col:0} should produce a main pane plus a frozen-rows strip",
    );

    const paneByKey = new Map(panes.map((pane) => [pane.props["data-grid-pane"], pane]));
    const frozenPane = paneByKey.get("frozen-rows");
    const mainPane = paneByKey.get("main");
    // 冻结行高 2 × 24px；两个窗格在网格正文层内各占其位。
    assert.equal(frozenPane.props.style.height, "48px");
    assert.equal(frozenPane.props.style.top, "0px");
    assert.equal(mainPane.props.style.top, "48px");

    const itemsIn = (pane) =>
      walk(pane).filter((node) => node.props?.["data-testid"] === "xlsx-image-item");
    const [frozenItem] = itemsIn(frozenPane);
    const [scrollingItem] = itemsIn(mainPane);
    assert.ok(frozenItem, "the y=0 image belongs to the frozen-rows pane");
    assert.ok(scrollingItem, "the y=240 image belongs to the main pane");

    // 冻结条随水平滚动、不随垂直滚动：left = 0 − 30，top 固定 0。
    assert.equal(frozenItem.props.style.left, "-30px");
    assert.equal(frozenItem.props.style.top, "0px");
    // 滚动区两轴都补偿，且相对 main 窗格原点（去掉 48px 冻结高）。
    assert.equal(scrollingItem.props.style.left, "70px");
    assert.equal(scrollingItem.props.style.top, "92px");
  } finally {
    mounted.app.unmount();
  }
});

test("XLSX overlay panes follow resized frozen columns so anchored objects stay aligned", async () => {
  const controller = {
    selectedImageId: null,
    zoomScale: 100,
    activeSheet: {
      name: "Sheet1",
      freezePanes: { row: 0, col: 2 },
      visibleRows: Array.from({ length: 20 }, (_, index) => index),
      visibleCols: Array.from({ length: 10 }, (_, index) => index),
      rowHeights: [],
      colWidths: [],
      rowHeightOverridesPx: {},
      // The user dragged columns A and B wider; the grid paints those widths.
      colWidthOverridesPx: { 0: 120, 1: 60 },
      defaultRowHeightPx: 24,
      defaultColWidthPx: 80,
    },
    images: [absoluteImage("pinned", 0, 0), absoluteImage("scrolling", 400, 0)],
    selectImage() {},
  };
  const mounted = await mount(XlsxImageLayer, {
    controller,
    showImages: true,
    scrollLeft: 40,
    scrollTop: 0,
  });
  try {
    const panes = walk(mounted.root).filter((node) => node.props?.["data-grid-pane"]);
    const paneByKey = new Map(panes.map((pane) => [pane.props["data-grid-pane"], pane]));
    // 120 + 60: the resized widths, not 2 × the 80px default.
    assert.equal(
      paneByKey.get("frozen-cols")?.props.style.width,
      "180px",
      "the frozen columns pane matches the widths the grid paints",
    );
    assert.equal(paneByKey.get("main")?.props.style.left, "180px");

    const itemsIn = (pane) =>
      walk(pane).filter((node) => node.props?.["data-testid"] === "xlsx-image-item");
    // The x=400 object sits in the scrolling pane, which starts at the resized
    // freeze line (180px) and compensates the 40px scroll.
    const [scrollingItem] = itemsIn(paneByKey.get("main"));
    assert.equal(scrollingItem.props.style.left, "180px");
  } finally {
    mounted.app.unmount();
  }
});
