import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const {
  DocxTableHost,
  DocxToolbar,
  findDocxSearchMatches,
  useDocxEditor,
} = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

function modelWithNodes(nodes) {
  return {
    ...createBlankDocumentModel(),
    nodes,
  };
}

function hasClass(node, className) {
  return String(node?.props?.class || "").split(/\s+/).includes(className);
}

async function mountController(model) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () => vue.h("div", { "data-testid": "controller-harness" });
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

test("DocxToolbar keeps export available while read-only blocks new and import", async () => {
  const calls = [];
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({
        starterModel: modelWithNodes([
          { type: "paragraph", children: [{ type: "text", text: "Read-only document" }] },
        ]),
      });
      controller.newDocument = () => calls.push("new");
      controller.importDocxFile = async () => calls.push("import");
      controller.exportDocx = () => calls.push("export");
      return () => vue.h(DocxToolbar, {
        controller,
        readOnly: true,
        currentPage: 1,
        totalPages: 1,
      });
    },
  });

  const mounted = await mount(Harness);
  const newButton = findByTestId(mounted.root, "editor-new");
  const importButton = findByTestId(mounted.root, "editor-import");
  const exportButton = findByTestId(mounted.root, "editor-export");

  assert.ok(newButton);
  assert.ok(importButton);
  assert.ok(exportButton);
  assert.equal(newButton.props.disabled, true);
  assert.equal(importButton.props.disabled, true);
  assert.equal(exportButton.props.disabled, undefined);

  exportButton.props.onClick();
  assert.deepEqual(calls, ["export"]);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DocxTableHost renders vertical merges once and commits editable cell text", async () => {
  const table = {
    type: "table",
    rows: [
      {
        cells: [
          {
            style: { rowSpan: 2 },
            nodes: [{ type: "paragraph", children: [{ type: "text", text: "Merged cell" }] }],
          },
          {
            nodes: [{ type: "paragraph", children: [{ type: "text", text: "First row" }] }],
          },
        ],
      },
      {
        cells: [
          {
            style: { vMergeContinuation: true },
            nodes: [{ type: "paragraph", children: [{ type: "text", text: "Hidden continuation" }] }],
          },
          {
            nodes: [{ type: "paragraph", children: [{ type: "text", text: "Second row" }] }],
          },
        ],
      },
    ],
  };
  const model = modelWithNodes([table]);
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () => vue.h(DocxTableHost, {
        table: controller.model.nodes[0],
        tableIndex: 0,
        editable: true,
        controller,
      });
    },
  });

  const mounted = await mount(Harness);
  const tableCells = walk(mounted.root).filter((node) => node.type === "td");
  assert.equal(tableCells.length, 3, "vertical merge continuation must not render another td");
  assert.equal(tableCells[0].props.rowspan, 2);
  assert.equal(
    walk(mounted.root).some((node) => String(node.text || "").includes("Hidden continuation")),
    false,
  );

  const tableHost = walk(mounted.root).find((node) => hasClass(node, "docx-table-host"));
  const cellParagraph = walk(mounted.root).find((node) =>
    node.props?.["data-docx-table-cell-paragraph-host"] === true &&
    node.props?.["data-docx-table-row-index"] === 0 &&
    node.props?.["data-docx-table-cell-index"] === 0
  );
  assert.ok(tableHost);
  assert.ok(cellParagraph);
  assert.equal(cellParagraph.props.contenteditable, "true");

  tableHost.querySelector = () => ({ textContent: "Merged cell updated" });
  cellParagraph.props.onFocus();
  cellParagraph.props.onBlur();
  await vue.nextTick();

  const committedTable = controller.model.nodes[0];
  assert.equal(committedTable.type, "table");
  assert.equal(committedTable.rows[0].cells[0].nodes[0].children[0].text, "Merged cell updated");
  assert.deepEqual(controller.selection, {
    kind: "table-cell",
    tableIndex: 0,
    rowIndex: 0,
    cellIndex: 0,
  });
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX surface search indexes an exact table-cell range without a block highlighter", async () => {
  const table = {
    type: "table",
    rows: [{
      cells: [{
        nodes: [{ type: "paragraph", children: [{ type: "text", text: "Needle inside table" }] }],
      }],
    }],
  };
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(DocxTableHost, {
        table,
        tableIndex: 0,
      });
    },
  });

  const matches = findDocxSearchMatches(modelWithNodes([table]), "needle");
  assert.equal(matches.length, 1);
  assert.deepEqual(matches[0].range, {
    start: {
      location: {
        kind: "table-cell",
        tableIndex: 0,
        rowIndex: 0,
        cellIndex: 0,
        paragraphIndex: 0,
      },
      offset: 0,
    },
    end: {
      location: {
        kind: "table-cell",
        tableIndex: 0,
        rowIndex: 0,
        cellIndex: 0,
        paragraphIndex: 0,
      },
      offset: 6,
    },
  });

  const mounted = await mount(Harness);
  const paragraph = walk(mounted.root).find((node) =>
    node.props?.["data-docx-table-cell-paragraph-host"] === true
  );
  assert.ok(paragraph);
  assert.equal(paragraph.props["data-docx-search-match"], undefined);
  assert.equal(paragraph.props["data-docx-search-active"], undefined);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("insertImageFile creates a visible floating square image with default dimensions", async () => {
  const mounted = await mountController(modelWithNodes([
    { type: "paragraph", children: [{ type: "text", text: "Image anchor" }] },
    { type: "paragraph", children: [{ type: "text", text: "Later selection" }] },
  ]));
  const file = new File(
    [new Uint8Array([137, 80, 78, 71])],
    "diagram.png",
    { type: "image/png" },
  );

  mounted.controller.selectParagraph(0);
  await vue.nextTick();
  mounted.controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 6 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 6 },
  });
  await vue.nextTick();

  let releaseFileRead;
  Object.defineProperty(file, "arrayBuffer", {
    configurable: true,
    value: () => new Promise((resolve) => {
      releaseFileRead = () => resolve(new Uint8Array([137, 80, 78, 71]).buffer);
    }),
  });
  const insertion = mounted.controller.insertImageFile(file);
  await Promise.resolve();
  mounted.controller.selectParagraph(1);
  releaseFileRead();
  await insertion;
  await vue.nextTick();

  const paragraph = mounted.controller.model.nodes[0];
  assert.equal(paragraph.type, "paragraph");
  const image = paragraph.children.find((child) => child.type === "image");
  assert.ok(image);
  assert.deepEqual(
    paragraph.children.map((child) => child.type === "text" ? child.text : child.type),
    ["Image ", "image", "anchor"],
    "the public insert command must place the image at the active caret",
  );
  assert.equal(
    mounted.controller.model.nodes[1].children.some((child) => child.type === "image"),
    false,
    "selection changes while reading the file must not move the insertion target",
  );
  assert.equal(image.widthPx, 240);
  assert.equal(image.heightPx, 160);
  assert.deepEqual(image.floating, {
    xPx: 72,
    yPx: 72,
    wrapType: "square",
    behindDocument: false,
    zIndex: 1,
  });
  assert.equal(mounted.controller.status, "Inserted image: diagram.png");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});
