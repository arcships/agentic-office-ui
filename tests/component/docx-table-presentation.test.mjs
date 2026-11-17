import assert from "node:assert/strict";
import test from "node:test";

import { importFromDemo, mount, vue, walk } from "./vue-test-renderer.mjs";

const { DocxTableHost } = await importFromDemo("@arcships/vue-docx");
const { estimateTableRowHeightsPx, resolveTableRenderMetricsPx } =
  await importFromDemo("@arcships/docx-core");

const PAGE_CONTENT_WIDTH_PX = 600;

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function tableNode({ style, rows }) {
  return {
    type: "table",
    style,
    rows: rows.map((row) => ({
      type: "table-row",
      style: row.style,
      cells: row.cells.map((cell) => ({
        type: "table-cell",
        style: cell.style,
        nodes: cell.nodes ?? [paragraph("Cell")],
      })),
    })),
  };
}

function border(color, sizeEighthPt = 6) {
  return { type: "single", color, sizeEighthPt };
}

async function renderTable(table, props = {}) {
  const Harness = vue.defineComponent({
    setup() {
      return () =>
        vue.h(DocxTableHost, {
          table,
          tableIndex: 0,
          maxAvailableWidthPx: PAGE_CONTENT_WIDTH_PX,
          ...props,
        });
    },
  });
  const mounted = await mount(Harness);
  const nodes = walk(mounted.root);
  return {
    mounted,
    table: nodes.find((node) => node.type === "table"),
    rows: nodes.filter((node) => node.type === "tr"),
    cells: nodes.filter((node) => node.type === "td"),
    cols: nodes.filter((node) => node.type === "col"),
  };
}

test("DOCX narrow table keeps its model width instead of stretching to 100%", async () => {
  const table = tableNode({
    style: {
      widthTwips: 3000,
      columnWidthsTwips: [1500, 1500],
    },
    rows: [
      { cells: [{ nodes: [paragraph("left")] }, { nodes: [paragraph("right")] }] },
    ],
  });

  const rendered = await renderTable(table);

  assert.equal(rendered.table.props.style.width, "200px");
  assert.deepEqual(
    rendered.cols.map((col) => col.props.style.width),
    ["100px", "100px"]
  );
  rendered.mounted.app.unmount();
});

test("DOCX wide table clamps to the page content width and columns follow", async () => {
  const table = tableNode({
    // 12000 twips = 800px, wider than the 600px content box.
    style: { widthTwips: 12000, columnWidthsTwips: [6000, 6000] },
    rows: [
      { cells: [{ nodes: [paragraph("left")] }, { nodes: [paragraph("right")] }] },
    ],
  });

  const metrics = resolveTableRenderMetricsPx(table, PAGE_CONTENT_WIDTH_PX);
  const rendered = await renderTable(table);

  assert.equal(rendered.table.props.style.width, "600px");
  assert.equal(Math.round(metrics.tableWidthPx), 600);
  assert.ok(
    Math.abs(
      metrics.columnWidthsPx.reduce((sum, width) => sum + width, 0) - 600
    ) < 1
  );
  assert.deepEqual(
    rendered.cols.map((col) => col.props.style.width),
    metrics.columnWidthsPx.map((width) => `${Math.round(width)}px`)
  );
  rendered.mounted.app.unmount();
});

test("DOCX table without tblW or a grid keeps Word's auto-fit width", async () => {
  const table = tableNode({
    rows: [
      {
        cells: [
          { nodes: [paragraph("a")] },
          { nodes: [paragraph("b")] },
          { nodes: [paragraph("c")] },
        ],
      },
    ],
  });

  const metrics = resolveTableRenderMetricsPx(table, PAGE_CONTENT_WIDTH_PX);
  const rendered = await renderTable(table);

  assert.equal(metrics.autoFitWidth, true);
  assert.equal(rendered.table.props.style.width, "600px");
  assert.deepEqual(
    rendered.cols.map((col) => col.props.style.width),
    ["200px", "200px", "200px"]
  );
  rendered.mounted.app.unmount();
});

test("DOCX borderless table renders without the hard-coded grey grid", async () => {
  const table = tableNode({
    style: { widthTwips: 3000, columnWidthsTwips: [1500, 1500] },
    rows: [
      { cells: [{ nodes: [paragraph("a")] }, { nodes: [paragraph("b")] }] },
      { cells: [{ nodes: [paragraph("c")] }, { nodes: [paragraph("d")] }] },
    ],
  });

  const rendered = await renderTable(table);

  for (const cell of rendered.cells) {
    assert.equal(cell.props.style.borderTop, undefined);
    assert.equal(cell.props.style.borderRight, undefined);
    assert.equal(cell.props.style.borderBottom, undefined);
    assert.equal(cell.props.style.borderLeft, undefined);
  }
  assert.equal(rendered.cells[0].props.style.border, undefined);
  rendered.mounted.app.unmount();
});

test("DOCX table borders follow tblBorders inside/outside edges", async () => {
  const table = tableNode({
    style: {
      widthTwips: 3000,
      columnWidthsTwips: [1500, 1500],
      borders: {
        top: border("#ff0000"),
        bottom: border("#0000ff", 12),
        left: border("#00aa00"),
        right: border("#aa00aa"),
        insideH: border("#00ffff"),
        insideV: border("#333333"),
      },
    },
    rows: [
      { cells: [{ nodes: [paragraph("a")] }, { nodes: [paragraph("b")] }] },
      { cells: [{ nodes: [paragraph("c")] }, { nodes: [paragraph("d")] }] },
      { cells: [{ nodes: [paragraph("e")] }, { nodes: [paragraph("f")] }] },
    ],
  });

  const rendered = await renderTable(table);
  const styleOf = (index) => rendered.cells[index].props.style;

  assert.equal(styleOf(0).borderTop, "1px solid #ff0000");
  assert.equal(styleOf(0).borderLeft, "1px solid #00aa00");
  assert.equal(styleOf(0).borderRight, "1px solid #333333");
  assert.equal(styleOf(0).borderBottom, "1px solid #00ffff");

  assert.equal(styleOf(1).borderRight, "1px solid #aa00aa");
  assert.equal(styleOf(1).borderLeft, "1px solid #333333");
  assert.equal(styleOf(1).borderTop, "1px solid #ff0000");

  // Middle row: inside edges above and below.
  assert.equal(styleOf(2).borderTop, "1px solid #00ffff");
  assert.equal(styleOf(2).borderBottom, "1px solid #00ffff");

  // Last row: the table bottom border wins over insideH.
  assert.equal(styleOf(4).borderBottom, "2px solid #0000ff");
  assert.equal(styleOf(4).borderTop, "1px solid #00ffff");
  assert.equal(styleOf(5).borderBottom, "2px solid #0000ff");
  assert.equal(styleOf(5).borderRight, "1px solid #aa00aa");
  rendered.mounted.app.unmount();
});

test("DOCX cell padding merges tcMar per edge over tblCellMar", async () => {
  const table = tableNode({
    style: {
      widthTwips: 3000,
      columnWidthsTwips: [3000],
      // 60 twips = 4px on every edge.
      cellMarginTwips: {
        topTwips: 60,
        rightTwips: 60,
        bottomTwips: 60,
        leftTwips: 60,
      },
    },
    rows: [
      {
        cells: [
          {
            style: { marginTwips: { rightTwips: 300 } },
            nodes: [paragraph("one edge")],
          },
        ],
      },
    ],
  });

  const rendered = await renderTable(table);
  const style = rendered.cells[0].props.style;

  assert.equal(style.paddingTop, 4);
  assert.equal(style.paddingRight, 20);
  assert.equal(style.paddingBottom, 4);
  assert.equal(style.paddingLeft, 4);
  rendered.mounted.app.unmount();
});

test("DOCX merged cell keeps its grid span and borders the merged edges", async () => {
  const table = tableNode({
    style: {
      widthTwips: 3000,
      columnWidthsTwips: [1500, 1500],
      borders: {
        top: border("#ff0000"),
        bottom: border("#0000ff"),
        left: border("#00aa00"),
        right: border("#aa00aa"),
        insideH: border("#00ffff"),
        insideV: border("#333333"),
      },
    },
    rows: [
      {
        cells: [
          { style: { gridSpan: 2 }, nodes: [paragraph("merged header")] },
        ],
      },
      {
        cells: [{ nodes: [paragraph("left")] }, { nodes: [paragraph("right")] }],
      },
    ],
  });

  const rendered = await renderTable(table);
  const [merged, left, right] = rendered.cells;

  // #31 materializes the grid, #34 must still style the merged box from the model.
  assert.equal(merged.props.colspan, 2);
  assert.equal(merged.props.rowspan, undefined);
  assert.equal(merged.props.style.borderLeft, "1px solid #00aa00");
  assert.equal(merged.props.style.borderRight, "1px solid #aa00aa");
  assert.equal(merged.props.style.borderTop, "1px solid #ff0000");
  assert.equal(merged.props.style.borderBottom, "1px solid #00ffff");

  assert.equal(left.props.style.borderBottom, "1px solid #0000ff");
  assert.equal(right.props.style.borderBottom, "1px solid #0000ff");
  assert.equal(rendered.table.props.style.width, "200px");
  rendered.mounted.app.unmount();
});

test("DOCX exact row height clips while atLeast rows keep a minimum", async () => {
  const table = tableNode({
    style: { widthTwips: 3000, columnWidthsTwips: [3000] },
    rows: [
      {
        style: { heightTwips: 720, heightRule: "exact" },
        cells: [{ nodes: [paragraph("exact")] }],
      },
      {
        style: { heightTwips: 720, heightRule: "atLeast" },
        cells: [{ nodes: [paragraph("at least")] }],
      },
    ],
  });

  const rendered = await renderTable(table);
  const [exactStyle, atLeastStyle] = rendered.rows.map((row) => row.props.style);

  // `height` on a table row is a minimum, so exact rows must clip as well.
  assert.equal(exactStyle.height, "48px");
  assert.equal(exactStyle.overflow, "hidden");
  assert.equal(exactStyle.minHeight, undefined);

  assert.equal(atLeastStyle.minHeight, "48px");
  assert.equal(atLeastStyle.height, undefined);
  assert.equal(atLeastStyle.overflow, undefined);
  rendered.mounted.app.unmount();
});

test("DOCX row shading and cantSplit reach the rendered row", async () => {
  const table = tableNode({
    style: { widthTwips: 3000, columnWidthsTwips: [3000] },
    rows: [
      {
        style: { backgroundColor: "#ffe699", cantSplit: true },
        cells: [{ nodes: [paragraph("shaded")] }],
      },
      { cells: [{ nodes: [paragraph("plain")] }] },
    ],
  });

  const rendered = await renderTable(table);
  const [shaded, plain] = rendered.rows.map((row) => row.props.style);

  assert.equal(shaded.backgroundColor, "#ffe699");
  assert.equal(shaded.breakInside, "avoid");
  assert.equal(plain.backgroundColor, undefined);
  assert.equal(plain.breakInside, undefined);
  rendered.mounted.app.unmount();
});

test("DOCX diagonal cell borders render as an overlay gradient", async () => {
  const table = tableNode({
    style: { widthTwips: 3000, columnWidthsTwips: [3000] },
    rows: [
      {
        cells: [
          {
            style: { borders: { tl2br: border("#000000") } },
            nodes: [paragraph("diagonal")],
          },
        ],
      },
    ],
  });

  const rendered = await renderTable(table);
  const image = String(rendered.cells[0].props.style.backgroundImage ?? "");

  assert.match(image, /linear-gradient\(to bottom right/);
  assert.equal(rendered.cells[0].props.style.backgroundSize, "100% 100%");
  rendered.mounted.app.unmount();
});

test("DOCX row height estimate keeps the table cell margins on the missing edges", () => {
  const cellMarginTwips = {
    topTwips: 240,
    rightTwips: 60,
    bottomTwips: 240,
    leftTwips: 60,
  };
  const build = (marginTwips) =>
    tableNode({
      style: { widthTwips: 3000, columnWidthsTwips: [3000], cellMarginTwips },
      rows: [
        {
          cells: [
            { style: marginTwips ? { marginTwips } : undefined, nodes: [paragraph("Hi")] },
          ],
        },
      ],
    });

  // A cell that only overrides its right margin must still inherit the
  // table's top/bottom padding, so the estimate matches a full override.
  const partial = estimateTableRowHeightsPx(
    build({ rightTwips: 300 }),
    PAGE_CONTENT_WIDTH_PX
  );
  const full = estimateTableRowHeightsPx(
    build({ topTwips: 240, rightTwips: 300, bottomTwips: 240, leftTwips: 60 }),
    PAGE_CONTENT_WIDTH_PX
  );
  // Same cell without a table-level cell margin: 2 x 4px fallback padding.
  const withoutTableMargin = estimateTableRowHeightsPx(
    tableNode({
      style: { widthTwips: 3000, columnWidthsTwips: [3000] },
      rows: [
        {
          cells: [
            {
              style: { marginTwips: { rightTwips: 300 } },
              nodes: [paragraph("Hi")],
            },
          ],
        },
      ],
    }),
    PAGE_CONTENT_WIDTH_PX
  );

  assert.deepEqual(partial, full);
  // 2 x 240 twips = 32px of inherited vertical padding vs the 2 x 4px fallback.
  assert.ok(
    partial[0] - withoutTableMargin[0] >= 20,
    `expected the inherited 32px vertical padding, got ${partial[0]} vs ${withoutTableMargin[0]}`
  );
});
