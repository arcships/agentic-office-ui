import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
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
const { createBlankDocumentModel, wasmBuildDocModelFromBytes } =
  await importFromDemo("@arcships/docx-core");

function modelWithNodes(nodes) {
  return {
    ...createBlankDocumentModel(),
    nodes,
  };
}

function hasClass(node, className) {
  return String(node?.props?.class || "").split(/\s+/).includes(className);
}

function paragraphText(node) {
  if (!node || node.type !== "paragraph") return "";
  return node.children.map((child) => child.type === "text" ? child.text : "").join("");
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

test("table cell blur commits each paragraph without duplicating text or dropping run styles", async () => {
  const table = {
    type: "table",
    sourceXml: "<w:tbl/>",
    rows: [{
      cells: [{
        nodes: [
          {
            type: "paragraph",
            sourceXml: "<w:p first/>",
            children: [{ type: "text", text: "first", style: { bold: true } }],
          },
          {
            type: "paragraph",
            sourceXml: "<w:p second/>",
            children: [{ type: "text", text: "second", style: { italic: true } }],
          },
        ],
      }],
    }],
  };
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: modelWithNodes([table]) });
      return () => vue.h(DocxTableHost, {
        table: controller.model.nodes[0],
        tableIndex: 0,
        editable: true,
        controller,
      });
    },
  });

  const mounted = await mount(Harness);
  const cellParagraphs = walk(mounted.root).filter((node) =>
    node.props?.["data-docx-table-cell-paragraph-host"] === true
  );
  assert.equal(cellParagraphs.length, 2);

  cellParagraphs[0].props.onInput({ target: { textContent: "first edited" } });
  cellParagraphs[1].props.onInput({ target: { textContent: "second edited" } });
  cellParagraphs[0].props.onBlur({ target: { textContent: "first edited" } });
  await vue.nextTick();
  cellParagraphs[1].props.onBlur({ target: { textContent: "second edited" } });
  await vue.nextTick();

  const committed = controller.model.nodes[0];
  assert.equal(committed.type, "table");
  const paragraphs = committed.rows[0].cells[0].nodes;
  assert.deepEqual(paragraphs.map(paragraphText), ["first edited", "second edited"]);
  assert.equal(paragraphs[0].children[0].style.bold, true);
  assert.equal(paragraphs[1].children[0].style.italic, true);
  assert.equal(paragraphs[0].sourceXml, undefined);
  assert.equal(paragraphs[1].sourceXml, undefined);
  assert.equal(committed.sourceXml, undefined, "table XML must be rebuilt after a cell edit");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("flushing cell drafts commits every dirty paragraph in one model chain", async () => {
  const table = {
    type: "table",
    rows: [{
      cells: [{
        nodes: [
          { type: "paragraph", children: [{ type: "text", text: "first" }] },
          { type: "paragraph", children: [{ type: "text", text: "second" }] },
        ],
      }],
    }],
  };
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: modelWithNodes([table]) });
      return () => vue.h(DocxTableHost, {
        table: controller.model.nodes[0],
        tableIndex: 0,
        editable: true,
        controller,
      });
    },
  });
  const mounted = await mount(Harness);
  const cellParagraphs = walk(mounted.root).filter((node) =>
    node.props?.["data-docx-table-cell-paragraph-host"] === true
  );

  cellParagraphs[0].props.onInput({ target: { textContent: "first edited" } });
  cellParagraphs[1].props.onInput({ target: { textContent: "second edited" } });
  controller.flushTextDrafts();
  await vue.nextTick();

  const paragraphs = controller.model.nodes[0].rows[0].cells[0].nodes;
  assert.deepEqual(paragraphs.map(paragraphText), ["first edited", "second edited"]);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("whole-cell commit diffs multiline text, keeps run styles, and clears source XML", async () => {
  const table = {
    type: "table",
    sourceXml: "<w:tbl/>",
    rows: [{
      cells: [{
        nodes: [
          {
            type: "paragraph",
            sourceXml: "<w:p first/>",
            children: [{ type: "text", text: "one", style: { bold: true } }],
          },
          {
            type: "paragraph",
            sourceXml: "<w:p second/>",
            children: [{ type: "text", text: "two", style: { underline: true } }],
          },
        ],
      }],
    }],
  };
  const mounted = await mountController(modelWithNodes([table]));
  mounted.controller.commitTableCellText(0, 0, 0, "one edited\ntwo edited");
  await vue.nextTick();

  const committed = mounted.controller.model.nodes[0];
  const paragraphs = committed.rows[0].cells[0].nodes;
  assert.deepEqual(paragraphs.map(paragraphText), ["one edited", "two edited"]);
  assert.equal(paragraphs[0].children[0].style.bold, true);
  assert.equal(paragraphs[1].children[0].style.underline, true);
  assert.equal(paragraphs[0].sourceXml, undefined);
  assert.equal(paragraphs[1].sourceXml, undefined);
  assert.equal(committed.sourceXml, undefined);
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

test("insertImageFile drops the insertion when the document changes while reading the file", async () => {
  const mounted = await mountController(modelWithNodes([
    { type: "paragraph", children: [{ type: "text", text: "Image anchor" }] },
  ]));
  mounted.controller.selectParagraph(0);
  await vue.nextTick();
  mounted.controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 6 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 6 },
  });
  await vue.nextTick();

  const file = new File([new Uint8Array([137, 80, 78, 71])], "late.png", { type: "image/png" });
  let releaseFileRead;
  Object.defineProperty(file, "arrayBuffer", {
    configurable: true,
    value: () => new Promise((resolve) => {
      releaseFileRead = () => resolve(new Uint8Array([137, 80, 78, 71]).buffer);
    }),
  });

  const insertion = mounted.controller.insertImageFile(file);
  await Promise.resolve();
  mounted.controller.newDocument();
  releaseFileRead();
  await insertion;
  await vue.nextTick();

  assert.equal(
    mounted.controller.model.nodes.some((node) =>
      node.type === "paragraph" && node.children.some((child) => child.type === "image")
    ),
    false,
    "a document loaded during the file read must not receive the image",
  );
  assert.equal(mounted.controller.status, "Image insertion cancelled: document changed");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

function tableRowCells(root) {
  return walk(root)
    .filter((node) => node.type === "tr")
    .map((row) => row.children.filter((child) => child.type === "td"));
}

function paragraphNodeText(node) {
  return (node?.children ?? [])
    .map((child) => (child.type === "text" ? child.text : ""))
    .join("");
}

async function loadMergedTableFixture() {
  const bytes = readFileSync(
    new URL("../../test-data/corpus/docx-merged-table.docx", import.meta.url),
  );
  const input = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const { model } = await wasmBuildDocModelFromBytes(input);
  return model;
}

test("DOCX table grid carries vMerge rowSpan and drops continuation cells", async () => {
  const fixture = await loadMergedTableFixture();
  const table = fixture.nodes[1];
  assert.equal(table.type, "table");

  const mounted = await mount(DocxTableHost, { table, tableIndex: 1 });
  const rows = tableRowCells(mounted.root);

  assert.equal(rows.length, 3);
  assert.deepEqual(rows.map((cells) => cells.length), [3, 2, 3]);
  assert.equal(rows[0][0].props.rowspan, 2);
  assert.deepEqual(
    rows[1].map((cell) => cell.props["data-docx-table-cell-index"]),
    [1, 2],
    "the continuation cell must be covered by the restart rowSpan",
  );
  assert.equal(
    walk(mounted.root).filter((node) => node.type === "col").length,
    3,
  );
  const renderedHtml = walk(mounted.root)
    .map((node) => String(node.props?.innerHTML ?? ""))
    .join("");
  assert.match(renderedHtml, /Merged A1/);
  assert.equal(renderedHtml.includes("Hidden continuation"), false);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX table grid counts columns from gridSpan instead of the first row", async () => {
  const fixture = await loadMergedTableFixture();
  const table = fixture.nodes[3];
  assert.equal(table.type, "table");
  table.style = {
    ...(table.style ?? {}),
    columnWidthsTwips: [1500, 3000, 1500],
  };

  const mounted = await mount(DocxTableHost, { table, tableIndex: 3 });
  const rows = tableRowCells(mounted.root);
  const columns = walk(mounted.root).filter((node) => node.type === "col");

  assert.equal(rows.length, 2);
  assert.equal(rows[0].length, 2, "the wide first cell stays a single td");
  assert.equal(rows[0][0].props.colspan, 2);
  assert.equal(rows[1].length, 3);
  assert.equal(
    columns.length,
    3,
    "the grid must use the widest row, not the first row's cell count",
  );
  assert.deepEqual(
    columns.map((column) => column.props.style.width),
    ["100px", "200px", "100px"],
    "stored column widths must line up with the full grid",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX table grid degrades vMerge continuations across a page slice", async () => {
  const fixture = await loadMergedTableFixture();
  const table = fixture.nodes[1];

  const mounted = await mount(DocxTableHost, {
    table,
    tableIndex: 1,
    rowRange: { startRowIndex: 1, endRowIndex: 2 },
  });
  const rows = tableRowCells(mounted.root);

  assert.equal(rows.length, 1);
  assert.equal(
    rows[0].length,
    3,
    "a continuation row without its restart must keep all three columns",
  );
  assert.equal(rows[0][0].props.rowspan, undefined);
  assert.equal(
    String(rows[0][0].props.class).includes("docx-table-cell--merge-fallback"),
    true,
  );
  assert.deepEqual(
    rows[0].slice(1).map((cell) => cell.props["data-docx-table-cell-index"]),
    [1, 2],
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("DOCX nested tables use a unique runtime key and commit into the inner cell", async () => {
  const fixture = await loadMergedTableFixture();
  const model = modelWithNodes(fixture.nodes);
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () => vue.h(DocxTableHost, {
        table: controller.model.nodes[5],
        tableIndex: 5,
        editable: true,
        controller,
      });
    },
  });

  const mounted = await mount(Harness);
  const tableHosts = walk(mounted.root).filter((node) =>
    hasClass(node, "docx-table-host")
  );
  const outerHost = tableHosts.find((node) =>
    node.props?.["data-docx-table-index"] === 5
  );
  const nestedHost = tableHosts.find((node) =>
    String(node.props?.["data-docx-table-index"] ?? "").startsWith(
      "body-cell-nested-table-5-0-0-1"
    )
  );

  assert.ok(outerHost);
  assert.ok(nestedHost, "the nested table must expose its runtime key");
  assert.notEqual(
    nestedHost.props["data-docx-table-index"],
    outerHost.props["data-docx-table-index"],
  );

  const innerParagraph = walk(nestedHost).find((node) =>
    node.props?.["data-docx-table-cell-paragraph-host"] === true &&
    node.props?.["data-docx-table-row-index"] === 0 &&
    node.props?.["data-docx-table-cell-index"] === 0
  );
  assert.ok(innerParagraph);
  assert.equal(innerParagraph.props.contenteditable, "true");

  innerParagraph.props.onFocus();
  await vue.nextTick();
  assert.equal(
    controller.selection.tableRuntimeKey,
    nestedHost.props["data-docx-table-index"],
  );

  nestedHost.querySelector = () => ({ textContent: "Inner A1 updated" });
  innerParagraph.props.onBlur();
  await vue.nextTick();

  const outerTable = controller.model.nodes[5];
  assert.equal(
    paragraphNodeText(outerTable.rows[0].cells[0].nodes[0]),
    "Outer A1",
    "editing the nested cell must not overwrite the outer cell",
  );
  const innerTable = outerTable.rows[0].cells[0].nodes[1];
  assert.equal(innerTable.type, "table");
  assert.equal(
    paragraphNodeText(innerTable.rows[0].cells[0].nodes[0]),
    "Inner A1 updated",
  );
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});
