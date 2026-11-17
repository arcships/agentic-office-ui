import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
} from "./vue-test-renderer.mjs";

const { useDocxEditor } = await importFromDemo("@arcships/vue-docx");
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
