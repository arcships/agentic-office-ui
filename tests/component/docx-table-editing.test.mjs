import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  teleportTarget,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxContextMenu, DocxEditor, DocxTableHost, DocxToolbar, useDocxEditor } =
  await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

function deepFreeze(value, seen = new Set()) {
  if (!value || typeof value !== "object" || seen.has(value)) return value;
  seen.add(value);
  for (const child of Object.values(value)) deepFreeze(child, seen);
  return Object.freeze(value);
}

function starterModel(nodes) {
  return {
    ...createBlankDocumentModel(),
    nodes,
  };
}

async function mountController(model) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () => vue.h("div", { "data-testid": "controller-harness" }, controller.status);
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

function tableRowGridWidth(row) {
  return row.cells.reduce(
    (total, cell) => total + (cell.style?.gridSpan > 1 ? cell.style.gridSpan : 1),
    0
  );
}

function tableGridWidth(table) {
  return Math.max(1, ...table.rows.map(tableRowGridWidth));
}

function tableNode(rows, widths = [1000, 1000, 1000]) {
  return {
    type: "table",
    style: {
      layout: "fixed",
      widthTwips: widths.reduce((sum, width) => sum + width, 0),
      columnWidthsTwips: widths,
    },
    rows,
  };
}

function cell(text, style = {}) {
  return {
    type: "table-cell",
    style: { widthTwips: 1000, ...style },
    nodes: [{ type: "paragraph", children: [{ type: "text", text }] }],
  };
}

function cellTexts(tableCell) {
  return tableCell.nodes
    .filter((node) => node.type === "paragraph")
    .map((node) =>
      node.children
        .map((child) => (child.type === "text" ? child.text : ""))
        .join("")
    );
}

test("checkbox toggle writes checked and keeps the legacy value in sync", async () => {
  const source = deepFreeze(starterModel([
    {
      type: "paragraph",
      children: [
        { type: "form-field", fieldType: "checkbox", checked: false, value: "0" },
      ],
    },
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;
  const location = { kind: "paragraph", nodeIndex: 0, childIndex: 0 };

  controller.toggleFormCheckbox(location);
  await vue.nextTick();
  let field = controller.model.nodes[0].children[0];
  assert.equal(field.checked, true);
  assert.equal(field.value, "1");

  controller.toggleFormCheckbox(location);
  await vue.nextTick();
  field = controller.model.nodes[0].children[0];
  assert.equal(field.checked, false);
  assert.equal(field.value, "0");

  controller.setFormFieldValue(location, "1");
  await vue.nextTick();
  field = controller.model.nodes[0].children[0];
  assert.equal(field.checked, true);
  assert.equal(field.value, "1");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("table row and column commands keep gridSpan, vMerge and columnWidthsTwips in sync", async () => {
  const source = deepFreeze(starterModel([
    { type: "paragraph", children: [{ type: "text", text: "before" }] },
    tableNode([
      {
        type: "table-row",
        cells: [
          cell("A", { gridSpan: 2, widthTwips: 2000 }),
          cell("B"),
        ],
      },
      {
        type: "table-row",
        cells: [cell("C"), cell("D"), cell("E")],
      },
    ]),
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;
  const tableIndex = 1;

  controller.insertTableRow(tableIndex, 0, "below");
  await vue.nextTick();
  let table = controller.model.nodes[tableIndex];
  assert.equal(table.rows.length, 3);
  assert.equal(tableRowGridWidth(table.rows[1]), tableGridWidth(table));
  assert.equal(table.rows[1].cells.length, 2);
  assert.equal(table.rows[1].cells[0].style?.gridSpan, 2);

  controller.insertTableColumn(tableIndex, 0, "right");
  await vue.nextTick();
  table = controller.model.nodes[tableIndex];
  assert.equal(tableGridWidth(table), 4);
  assert.equal(table.style?.columnWidthsTwips?.length, 4);
  assert.equal(
    table.style.columnWidthsTwips.reduce((sum, width) => sum + width, 0),
    table.style.widthTwips
  );
  assert.equal(table.rows.every((row) => tableRowGridWidth(row) === 4), true);

  controller.deleteTableColumn(tableIndex, 0);
  await vue.nextTick();
  table = controller.model.nodes[tableIndex];
  assert.equal(tableGridWidth(table), 3);
  assert.equal(table.style?.columnWidthsTwips?.length, 3);
  assert.equal(table.rows.every((row) => tableRowGridWidth(row) === 3), true);

  controller.undo();
  await vue.nextTick();
  table = controller.model.nodes[tableIndex];
  assert.equal(table.style?.columnWidthsTwips?.length, 4);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("row insertion and deletion repair vertical merges", async () => {
  const source = deepFreeze(starterModel([
    tableNode([
      {
        type: "table-row",
        cells: [cell("merged", { rowSpan: 2 }), cell("top right")],
      },
      {
        type: "table-row",
        cells: [cell("", { vMergeContinuation: true }), cell("bottom right")],
      },
    ], [1000, 1000]),
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;

  controller.insertTableRow(0, 0, "below");
  await vue.nextTick();
  let table = controller.model.nodes[0];
  assert.equal(table.rows.length, 3);
  assert.equal(table.rows[0].cells[0].style?.rowSpan, 3);
  assert.equal(table.rows[1].cells[0].style?.vMergeContinuation, true);
  assert.equal(table.rows[2].cells[0].style?.vMergeContinuation, true);

  controller.deleteTableRow(0, 0);
  await vue.nextTick();
  table = controller.model.nodes[0];
  assert.equal(table.rows.length, 2);
  assert.equal(table.rows[0].cells[0].style?.vMergeContinuation, undefined);
  assert.equal(table.rows[1].cells[0].style?.vMergeContinuation, undefined);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("destructive table commands report refusals and leave a valid body", async () => {
  const source = deepFreeze(starterModel([
    tableNode([{ type: "table-row", cells: [cell("only")] }], [1000]),
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;

  controller.deleteTableRow(0, 0);
  await vue.nextTick();
  assert.match(controller.status, /only table row/i);
  assert.equal(controller.model.nodes[0].type, "table");

  controller.deleteTableColumn(0, 0);
  await vue.nextTick();
  assert.match(controller.status, /only table column/i);
  assert.equal(controller.model.nodes[0].type, "table");

  controller.selectTableCell(0, 0, 0);
  await vue.nextTick();
  controller.deleteTable(0);
  await vue.nextTick();
  assert.equal(controller.status, "Deleted table");
  assert.equal(controller.model.nodes.length, 1);
  assert.equal(controller.model.nodes[0].type, "paragraph");
  assert.deepEqual(controller.selection, { kind: "paragraph", nodeIndex: 0 });
  assert.equal(controller.activeTextRange, undefined);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("inserted tables have geometry and valid cell paragraphs", async () => {
  const mounted = await mountController(starterModel([
    { type: "paragraph", children: [{ type: "text", text: "body" }] },
  ]));
  const { controller } = mounted;

  controller.insertTable();
  await vue.nextTick();
  const table = controller.model.nodes[controller.model.nodes.length - 1];
  assert.equal(table.type, "table");
  assert.ok(table.style?.widthTwips > 0);
  assert.equal(table.style?.columnWidthsTwips?.length, 3);
  assert.ok(table.style?.borders?.top);
  assert.equal(
    table.style.columnWidthsTwips.reduce((sum, width) => sum + width, 0),
    table.style.widthTwips
  );
  for (const row of table.rows) {
    for (const tableCell of row.cells) {
      const paragraph = tableCell.nodes[0];
      assert.equal(paragraph.type, "paragraph");
      assert.ok(paragraph.children.length > 0);
    }
  }

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("merge and split table cells keep the grid rectangular", async () => {
  const source = deepFreeze(starterModel([
    tableNode([
      { type: "table-row", cells: [cell("A"), cell("B")] },
      { type: "table-row", cells: [cell("C"), cell("D")] },
    ], [1000, 1000]),
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;

  controller.mergeTableCells(0, {
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 1,
    focusCellIndex: 1,
  });
  await vue.nextTick();
  let table = controller.model.nodes[0];
  assert.equal(table.rows[0].cells.length, 1);
  assert.equal(table.rows[0].cells[0].style?.gridSpan, 2);
  assert.equal(table.rows[0].cells[0].style?.rowSpan, 2);
  assert.equal(table.rows[1].cells.length, 1);
  assert.equal(table.rows[1].cells[0].style?.vMergeContinuation, true);
  assert.deepEqual(cellTexts(table.rows[0].cells[0]), ["A", "B", "C", "D"]);
  assert.deepEqual(controller.selection, {
    kind: "table-cell",
    tableIndex: 0,
    rowIndex: 0,
    cellIndex: 0,
  });

  controller.splitTableCell(0, 0, 0);
  await vue.nextTick();
  table = controller.model.nodes[0];
  assert.equal(table.rows[0].cells.length, 2);
  assert.equal(table.rows[1].cells.length, 2);
  assert.equal(table.rows[0].cells[0].style?.gridSpan, undefined);
  assert.equal(table.rows[0].cells[0].style?.rowSpan, undefined);
  assert.equal(table.rows[1].cells[0].style?.vMergeContinuation, undefined);
  assert.deepEqual(cellTexts(table.rows[0].cells[0]), ["A", "B", "C", "D"]);

  const beforeRefusal = JSON.stringify(controller.model);
  controller.splitTableCell(0, 0, 0);
  await vue.nextTick();
  assert.match(controller.status, /not merged/i);
  assert.equal(JSON.stringify(controller.model), beforeRefusal);

  controller.mergeTableCells(0, {
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 0,
    focusCellIndex: 0,
  });
  await vue.nextTick();
  assert.match(controller.status, /single cell/i);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("moving a table keeps section boundaries attached to their nodes", async () => {
  const model = starterModel([
    { type: "paragraph", children: [{ type: "text", text: "section zero" }] },
    tableNode([{ type: "table-row", cells: [cell("moved")] }], [1000]),
    { type: "paragraph", children: [{ type: "text", text: "section one" }] },
  ]);
  model.metadata.sections = [
    { startNodeIndex: 0, headerSections: [], footerSections: [] },
    { startNodeIndex: 2, headerSections: [], footerSections: [] },
  ];
  const mounted = await mountController(deepFreeze(model));
  const { controller } = mounted;

  controller.moveTable(1, 3);
  await vue.nextTick();
  assert.deepEqual(
    controller.model.nodes.map((node) => node.type),
    ["paragraph", "paragraph", "table"]
  );
  assert.equal(controller.model.metadata.sections[1].startNodeIndex, 1);

  controller.moveTable(2, 1);
  await vue.nextTick();
  assert.deepEqual(
    controller.model.nodes.map((node) => node.type),
    ["paragraph", "table", "paragraph"]
  );
  assert.equal(controller.model.metadata.sections[1].startNodeIndex, 2);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("list commands resolve numbering definitions and respect their maximum level", async () => {
  const source = deepFreeze(starterModel([
    { type: "paragraph", children: [{ type: "text", text: "item" }] },
    { type: "paragraph", children: [{ type: "text", text: "other" }] },
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;

  controller.selectParagraph(0);
  await vue.nextTick();
  controller.toggleList("unordered");
  await vue.nextTick();
  const para = controller.model.nodes[0];
  assert.ok(para.style?.numbering);
  const definitions = controller.model.metadata.numberingDefinitions;
  assert.ok(definitions);
  const bulletInstance = definitions.instances.find(
    (instance) => instance.numId === para.style.numbering.numId
  );
  assert.ok(bulletInstance);
  const bulletAbstract = definitions.abstracts.find(
    (abstract) => abstract.abstractNumId === bulletInstance.abstractNumId
  );
  assert.equal(bulletAbstract?.levels[0].format, "bullet");

  controller.selectParagraph(1);
  await vue.nextTick();
  controller.toggleList("ordered");
  await vue.nextTick();
  const orderedPara = controller.model.nodes[1];
  assert.notEqual(
    orderedPara.style?.numbering?.numId,
    para.style.numbering.numId
  );

  const limitedDefinitions = {
    abstracts: [{
      abstractNumId: 10,
      levels: [
        { ilvl: 0, format: "decimal", text: "%1." },
        { ilvl: 1, format: "lowerLetter", text: "%2." },
        { ilvl: 2, format: "lowerRoman", text: "%3." },
      ],
    }],
    instances: [{ numId: 7, abstractNumId: 10 }],
  };
  const limitedModel = starterModel([
    {
      type: "paragraph",
      style: { numbering: { numId: 7, ilvl: 0 } },
      children: [{ type: "text", text: "limited" }],
    },
    tableNode([
      {
        type: "table-row",
        cells: [{
          type: "table-cell",
          nodes: [{
            type: "paragraph",
            style: { numbering: { numId: 7, ilvl: 0 } },
            children: [{ type: "text", text: "cell item" }],
          }],
        }],
      },
    ], [1000]),
  ]);
  limitedModel.metadata.numberingDefinitions = limitedDefinitions;
  const limited = await mountController(deepFreeze(limitedModel));
  const limitedController = limited.controller;

  limitedController.selectParagraph(0);
  await vue.nextTick();
  assert.equal(limitedController.adjustSelectedListDepth(1), true);
  assert.equal(limitedController.adjustSelectedListDepth(1), true);
  assert.equal(limitedController.adjustSelectedListDepth(1), false);
  assert.equal(limitedController.model.nodes[0].style.numbering.ilvl, 2);
  assert.match(limitedController.status, /Maximum list level is 3/);

  limitedController.selectTableCell(1, 0, 0);
  await vue.nextTick();
  assert.equal(limitedController.adjustSelectedListDepth(1), true);
  assert.equal(
    limitedController.model.nodes[1].rows[0].cells[0].nodes[0].style.numbering.ilvl,
    1
  );

  limited.app.unmount();
  assert.deepEqual(limited.warnings, []);
  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

function installFakeDocument(t) {
  const originalDocument = globalThis.document;
  const listeners = new Map();
  const createElement = (tagName) => ({
    tagName: String(tagName).toUpperCase(),
    style: {},
    children: [],
    innerHTML: "",
    appendChild(child) {
      this.children.push(child);
      return child;
    },
    setAttribute() {},
    removeAttribute() {},
    getAttribute() { return null; },
  });
  const fakeDocument = {
    createElement,
    createTextNode(text) {
      return { nodeType: 3, textContent: String(text) };
    },
    addEventListener(type, listener) {
      const entries = listeners.get(type) ?? new Set();
      entries.add(listener);
      listeners.set(type, entries);
    },
    removeEventListener(type, listener) {
      listeners.get(type)?.delete(listener);
    },
    dispatch(type) {
      for (const listener of listeners.get(type) ?? []) listener();
    },
  };
  globalThis.document = fakeDocument;
  t.after(() => {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
  });
  return fakeDocument;
}

function tableCellHosts(root) {
  return walk(root).filter(
    (node) =>
      node.type === "td" && node.props?.["data-docx-table-cell"] === true
  );
}

function tableCellHostAt(root, rowIndex, cellIndex) {
  return tableCellHosts(root).find(
    (node) =>
      node.props?.["data-docx-table-row-index"] === rowIndex &&
      node.props?.["data-docx-table-cell-index"] === cellIndex
  );
}

test("table host drag selection exposes a multi-cell range and keeps it active", async (t) => {
  const fakeDocument = installFakeDocument(t);
  const source = deepFreeze(starterModel([
    tableNode([
      { type: "table-row", cells: [cell("A"), cell("B")] },
      { type: "table-row", cells: [cell("C"), cell("D")] },
    ], [1000, 1000]),
  ]));
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: source });
      return () => vue.h(DocxTableHost, {
        table: controller.model.nodes[0],
        tableIndex: 0,
        editable: true,
        controller,
      });
    },
  });
  const mounted = await mount(Harness);

  const first = tableCellHostAt(mounted.root, 0, 0);
  const second = tableCellHostAt(mounted.root, 0, 1);
  const fourth = tableCellHostAt(mounted.root, 1, 1);
  assert.ok(first && second && fourth);

  await first.props.onMousedown({
    button: 0,
    shiftKey: false,
    target: null,
    preventDefault() {},
  });
  await vue.nextTick();
  await second.props.onMouseenter({ buttons: 1 });
  await vue.nextTick();

  assert.deepEqual(controller.selection, {
    kind: "table-cell",
    tableIndex: 0,
    rowIndex: 0,
    cellIndex: 1,
    cellRange: {
      tableIndex: 0,
      anchorRowIndex: 0,
      anchorCellIndex: 0,
      focusRowIndex: 0,
      focusCellIndex: 1,
      gridRange: { startColumnIndex: 0, endColumnIndex: 1 },
    },
  });
  assert.deepEqual(
    tableCellHosts(mounted.root)
      .filter((node) =>
        String(node.props?.class || "").split(/\s+/).includes("docx-table-cell--active")
      )
      .map((node) => [
        node.props["data-docx-table-row-index"],
        node.props["data-docx-table-cell-index"],
      ]),
    [[0, 0], [0, 1]]
  );

  fakeDocument.dispatch("mouseup");
  await fourth.props.onMousedown({
    button: 0,
    shiftKey: true,
    target: null,
    preventDefault() {},
  });
  await vue.nextTick();
  assert.deepEqual(controller.selection.cellRange, {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 1,
    focusCellIndex: 1,
    gridRange: { startColumnIndex: 0, endColumnIndex: 1 },
  });

  // A text range inside the selected rectangle must not collapse the range.
  controller.setActiveTextRange({
    start: { location: { kind: "table-cell", tableIndex: 0, rowIndex: 1, cellIndex: 1, paragraphIndex: 0 }, offset: 0 },
    end: { location: { kind: "table-cell", tableIndex: 0, rowIndex: 1, cellIndex: 1, paragraphIndex: 0 }, offset: 1 },
  });
  await vue.nextTick();
  assert.deepEqual(controller.selection.cellRange, {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 1,
    focusCellIndex: 1,
    gridRange: { startColumnIndex: 0, endColumnIndex: 1 },
  });

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("dragging across a 2x3 rectangle highlights every cell", async (t) => {
  const fakeDocument = installFakeDocument(t);
  const source = deepFreeze(starterModel([
    tableNode([
      { type: "table-row", cells: [cell("A"), cell("B"), cell("C")] },
      { type: "table-row", cells: [cell("D"), cell("E"), cell("F")] },
    ], [1000, 1000, 1000]),
  ]));
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: source });
      return () => vue.h(DocxTableHost, {
        table: controller.model.nodes[0],
        tableIndex: 0,
        editable: true,
        controller,
      });
    },
  });
  const mounted = await mount(Harness);

  await tableCellHostAt(mounted.root, 0, 0).props.onMousedown({
    button: 0,
    shiftKey: false,
    target: null,
    preventDefault() {},
  });
  await vue.nextTick();
  await tableCellHostAt(mounted.root, 1, 2).props.onMouseenter({ buttons: 1 });
  await vue.nextTick();

  assert.deepEqual(controller.selectedTableCellRange, {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 1,
    focusCellIndex: 2,
    gridRange: { startColumnIndex: 0, endColumnIndex: 2 },
  });
  assert.deepEqual(
    tableCellHosts(mounted.root)
      .filter((node) =>
        String(node.props?.class || "")
          .split(/\s+/)
          .includes("docx-table-cell--range-selected")
      )
      .map((node) => [
        node.props["data-docx-table-row-index"],
        node.props["data-docx-table-cell-index"],
      ]),
    [[0, 0], [0, 1], [0, 2], [1, 0], [1, 1], [1, 2]]
  );

  fakeDocument.dispatch("mouseup");
  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("table range formatting, copy and clear cover every highlighted cell", async () => {
  const source = deepFreeze(starterModel([
    tableNode([
      { type: "table-row", cells: [cell("A"), cell("B"), cell("C")] },
      { type: "table-row", cells: [cell("D"), cell("E"), cell("F")] },
    ], [1000, 1000, 1000]),
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;

  controller.selectTableCellRange(0, {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 1,
    focusCellIndex: 2,
  });
  await vue.nextTick();
  assert.deepEqual(controller.selectedTableCellRange?.gridRange, {
    startColumnIndex: 0,
    endColumnIndex: 2,
  });

  // Bold applies to every paragraph under the rectangle, not only the focus.
  controller.toggleBold();
  await vue.nextTick();
  const table = controller.model.nodes[0];
  for (const row of table.rows) {
    for (const tableCell of row.cells) {
      for (const paragraph of tableCell.nodes) {
        assert.equal(paragraph.children[0]?.style?.bold, true);
      }
    }
  }

  // Copy writes the selected matrix as TSV plus an HTML table.
  const clipboard = new Map();
  assert.equal(
    controller.writeSelectionToClipboardData({
      setData(type, value) {
        clipboard.set(type, value);
      },
    }),
    true
  );
  assert.equal(clipboard.get("text/plain"), "A\tB\tC\nD\tE\tF");
  assert.match(clipboard.get("text/html"), /^<table><tr><td>A<\/td>/);

  // Clearing the range empties all six cells through one undo entry.
  assert.equal(controller.clearTableSelection(), true);
  await vue.nextTick();
  const cleared = controller.model.nodes[0];
  for (const row of cleared.rows) {
    for (const tableCell of row.cells) {
      assert.deepEqual(cellTexts(tableCell), [""]);
    }
  }

  controller.undo();
  await vue.nextTick();
  const restored = controller.model.nodes[0];
  assert.deepEqual(
    restored.rows.map((row) => row.cells.map((tableCell) => cellTexts(tableCell)[0])),
    [["A", "B", "C"], ["D", "E", "F"]]
  );

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("whole-table selection spans the physical grid and highlights merged cells", async () => {
  const source = deepFreeze(starterModel([
    tableNode([
      { type: "table-row", cells: [cell("A", { gridSpan: 2 }), cell("C")] },
      { type: "table-row", cells: [cell("D"), cell("E"), cell("F")] },
    ], [1000, 1000, 1000]),
  ]));
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: source });
      return () => vue.h(DocxTableHost, {
        table: controller.model.nodes[0],
        tableIndex: 0,
        editable: true,
        controller,
      });
    },
  });
  const mounted = await mount(Harness);

  controller.selectTableCell(0, 1, 2);
  await vue.nextTick();
  controller.selectWholeTable();
  await vue.nextTick();
  assert.deepEqual(controller.selectedTableCellRange, {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 1,
    focusCellIndex: 2,
    gridRange: { startColumnIndex: 0, endColumnIndex: 2 },
  });

  const highlighted = tableCellHosts(mounted.root).filter((node) =>
    String(node.props?.class || "")
      .split(/\s+/)
      .includes("docx-table-cell--range-selected")
  );
  assert.deepEqual(
    highlighted.map((node) => [
      node.props["data-docx-table-row-index"],
      node.props["data-docx-table-cell-index"],
    ]),
    [[0, 0], [0, 1], [1, 0], [1, 1], [1, 2]]
  );
  const focusHighlights = tableCellHosts(mounted.root).filter((node) =>
    String(node.props?.class || "")
      .split(/\s+/)
      .includes("docx-table-cell--range-focus")
  );
  assert.deepEqual(
    focusHighlights.map((node) => [
      node.props["data-docx-table-row-index"],
      node.props["data-docx-table-cell-index"],
    ]),
    [[1, 2]]
  );

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("Escape collapses a table range to its anchor and Delete clears it", async (t) => {
  const originalWindow = globalThis.window;
  globalThis.window = {
    innerWidth: 1024,
    innerHeight: 768,
    addEventListener() {},
    removeEventListener() {},
  };
  t.after(() => {
    if (originalWindow === undefined) delete globalThis.window;
    else globalThis.window = originalWindow;
  });

  const source = deepFreeze(starterModel([
    tableNode([
      { type: "table-row", cells: [cell("A"), cell("B")] },
      { type: "table-row", cells: [cell("C"), cell("D")] },
    ], [1000, 1000]),
  ]));
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: source });
      return () => vue.h(DocxEditor, {
        editor: controller,
        editable: true,
        showToolbar: false,
        showThumbnails: false,
      });
    },
  });
  const mounted = await mount(Harness);
  const editorRoot = findByTestId(mounted.root, "docx-editor");
  assert.ok(editorRoot, "the editor root renders");

  let prevented = 0;
  const keydown = (key) => editorRoot.props.onKeydown({
    key,
    ctrlKey: false,
    metaKey: false,
    shiftKey: false,
    isComposing: false,
    defaultPrevented: false,
    target: { tagName: "DIV" },
    preventDefault() { prevented += 1; },
  });

  controller.selectTableCellRange(0, {
    tableIndex: 0,
    anchorRowIndex: 1,
    anchorCellIndex: 1,
    focusRowIndex: 0,
    focusCellIndex: 0,
  });
  await vue.nextTick();
  keydown("Escape");
  await vue.nextTick();
  assert.equal(prevented, 1);
  assert.deepEqual(controller.selection, {
    kind: "table-cell",
    tableIndex: 0,
    rowIndex: 1,
    cellIndex: 1,
  });

  controller.selectTableCellRange(0, {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 1,
    focusCellIndex: 1,
  });
  await vue.nextTick();
  keydown("Delete");
  await vue.nextTick();
  assert.equal(prevented, 2);
  const table = controller.model.nodes[0];
  for (const row of table.rows) {
    for (const tableCell of row.cells) {
      assert.deepEqual(cellTexts(tableCell), [""]);
    }
  }
  assert.deepEqual(controller.selection.cellRange, {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 1,
    focusCellIndex: 1,
    gridRange: { startColumnIndex: 0, endColumnIndex: 1 },
  });

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("context menu merges a selected range and splits the merged cell", async () => {
  const source = deepFreeze(starterModel([
    tableNode([
      { type: "table-row", cells: [cell("A"), cell("B")] },
      { type: "table-row", cells: [cell("C"), cell("D")] },
    ], [1000, 1000]),
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;
  const menu = await mount(DocxContextMenu, {
    context: {
      kind: "table",
      tableContext: { tableIndex: 0, rowIndex: 0, cellIndex: 0 },
    },
    controller,
    x: 10,
    y: 10,
  });
  const body = teleportTarget("body");
  await vue.nextTick();

  const mergeButton = findByTestId(body, "docx-context-menu-merge-cells");
  const splitButton = findByTestId(body, "docx-context-menu-split-cell");
  const selectTableButton = findByTestId(body, "docx-context-menu-select-table");
  const clearButton = findByTestId(body, "docx-context-menu-clear-cells");
  assert.ok(mergeButton, "context menu must expose Merge cells");
  assert.ok(splitButton, "context menu must expose Split cell");
  assert.ok(selectTableButton, "context menu must expose Select table");
  assert.ok(clearButton, "context menu must expose Clear contents");
  assert.equal(mergeButton.props.disabled, true, "merge needs a multi-cell range");
  assert.equal(clearButton.props.disabled, true, "clear needs a multi-cell range");

  controller.selectTableCellRange(0, {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 1,
    focusCellIndex: 1,
  });
  await vue.nextTick();
  assert.equal(mergeButton.props.disabled, false);
  assert.equal(clearButton.props.disabled, false);
  await selectTableButton.props.onClick();
  await vue.nextTick();
  assert.deepEqual(controller.selectedTableCellRange, {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 1,
    focusCellIndex: 1,
    gridRange: { startColumnIndex: 0, endColumnIndex: 1 },
  });
  await mergeButton.props.onClick();
  await vue.nextTick();

  let table = controller.model.nodes[0];
  assert.equal(table.rows[0].cells.length, 1);
  assert.equal(table.rows[0].cells[0].style?.gridSpan, 2);
  assert.equal(table.rows[0].cells[0].style?.rowSpan, 2);
  assert.equal(findByTestId(body, "docx-context-menu-split-cell").props.disabled, false);

  await findByTestId(body, "docx-context-menu-split-cell").props.onClick();
  await vue.nextTick();
  table = controller.model.nodes[0];
  assert.equal(table.rows[0].cells.length, 2);
  assert.equal(table.rows[1].cells.length, 2);

  menu.app.unmount();
  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("toolbar exposes row, column, merge and split commands for a table selection", async () => {
  const source = deepFreeze(starterModel([
    tableNode([
      { type: "table-row", cells: [cell("A"), cell("B")] },
      { type: "table-row", cells: [cell("C"), cell("D")] },
    ], [1000, 1000]),
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;
  const toolbar = await mount(DocxToolbar, { controller });

  controller.selectTableCell(0, 0, 0);
  await vue.nextTick();
  assert.ok(findByTestId(toolbar.root, "editor-table-actions"), "table actions appear for a table selection");
  assert.ok(findByTestId(toolbar.root, "editor-table-select-all"));
  const clearButton = findByTestId(toolbar.root, "editor-table-clear");
  assert.equal(clearButton.props.disabled, true, "clear needs a multi-cell range");

  await findByTestId(toolbar.root, "editor-table-select-all").props.onClick();
  await vue.nextTick();
  assert.deepEqual(controller.selectedTableCellRange, {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 1,
    focusCellIndex: 1,
    gridRange: { startColumnIndex: 0, endColumnIndex: 1 },
  });
  assert.equal(clearButton.props.disabled, false);
  await clearButton.props.onClick();
  await vue.nextTick();
  assert.deepEqual(
    controller.model.nodes[0].rows.map((row) => row.cells.map((tableCell) => cellTexts(tableCell)[0])),
    [["", ""], ["", ""]]
  );
  controller.undo();
  await vue.nextTick();
  assert.deepEqual(
    controller.model.nodes[0].rows.map((row) => row.cells.map((tableCell) => cellTexts(tableCell)[0])),
    [["A", "B"], ["C", "D"]]
  );

  controller.selectTableCell(0, 0, 0);
  await vue.nextTick();

  await findByTestId(toolbar.root, "editor-table-row-below").props.onClick();
  await vue.nextTick();
  assert.equal(controller.model.nodes[0].rows.length, 3);

  controller.selectTableCell(0, 0, 0);
  await vue.nextTick();
  await findByTestId(toolbar.root, "editor-table-column-right").props.onClick();
  await vue.nextTick();
  assert.equal(tableGridWidth(controller.model.nodes[0]), 3);

  controller.selectTableCellRange(0, {
    tableIndex: 0,
    anchorRowIndex: 0,
    anchorCellIndex: 0,
    focusRowIndex: 0,
    focusCellIndex: 1,
  });
  await vue.nextTick();
  const mergeButton = findByTestId(toolbar.root, "editor-table-merge");
  assert.equal(mergeButton.props.disabled, false);
  await mergeButton.props.onClick();
  await vue.nextTick();
  assert.equal(controller.model.nodes[0].rows[0].cells[0].style?.gridSpan, 2);

  const splitButton = findByTestId(toolbar.root, "editor-table-split");
  assert.equal(splitButton.props.disabled, false);
  await splitButton.props.onClick();
  await vue.nextTick();
  assert.equal(controller.model.nodes[0].rows[0].cells[0].style?.gridSpan, undefined);

  toolbar.app.unmount();
  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("table, list and form-field commands reuse untouched top-level nodes", async () => {
  const tableModel = starterModel([
    { type: "paragraph", children: [{ type: "text", text: "before" }] },
    tableNode([{ type: "table-row", cells: [cell("only")] }], [1000]),
    { type: "paragraph", children: [{ type: "text", text: "after" }] },
  ]);
  const tableMounted = await mountController(deepFreeze(tableModel));
  const tableController = tableMounted.controller;
  const untouchedParagraph = tableController.model.nodes[2];
  const beforeModel = tableController.model;
  const beforeJson = JSON.stringify(beforeModel);
  tableController.insertTableRow(1, 0, "below");
  await vue.nextTick();
  assert.strictEqual(tableController.model.nodes[2], untouchedParagraph);
  assert.equal(JSON.stringify(beforeModel), beforeJson);
  tableMounted.app.unmount();

  const listMounted = await mountController(deepFreeze(starterModel([
    { type: "paragraph", children: [{ type: "text", text: "item" }] },
    { type: "paragraph", children: [{ type: "text", text: "untouched" }] },
  ])));
  const listController = listMounted.controller;
  const listUntouched = listController.model.nodes[1];
  listController.selectParagraph(0);
  await vue.nextTick();
  listController.toggleList("unordered");
  await vue.nextTick();
  assert.strictEqual(listController.model.nodes[1], listUntouched);
  listMounted.app.unmount();

  const formMounted = await mountController(deepFreeze(starterModel([
    { type: "paragraph", children: [{ type: "form-field", fieldType: "checkbox", checked: false, value: "0" }] },
    { type: "paragraph", children: [{ type: "text", text: "untouched" }] },
  ])));
  const formController = formMounted.controller;
  const formUntouched = formController.model.nodes[1];
  formController.toggleFormCheckbox({ kind: "paragraph", nodeIndex: 0, childIndex: 0 });
  await vue.nextTick();
  assert.strictEqual(formController.model.nodes[1], formUntouched);
  formMounted.app.unmount();
});
