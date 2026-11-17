import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
} from "./vue-test-renderer.mjs";

const { useDocxEditor } = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

function paragraphText(node) {
  if (!node || node.type !== "paragraph") return "";
  return node.children
    .filter((child) => child.type === "text")
    .map((child) => child.text)
    .join("");
}

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

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function table(rows) {
  return {
    type: "table",
    rows: rows.map((cells) => ({
      type: "table-row",
      cells: cells.map((text) => ({
        type: "table-cell",
        nodes: [paragraph(text)],
      })),
    })),
  };
}

function cellTexts(tableNode) {
  return tableNode.rows.map((row) => row.cells.map((cell) => paragraphText(cell.nodes[0])));
}

function locationAt(nodeIndex, offset = 0) {
  return { location: { kind: "paragraph", nodeIndex }, offset };
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

test("row and column insertion honours the reference index without a double offset", async () => {
  const source = deepFreeze(starterModel([
    paragraph("body"),
    table([
      ["r0c0", "r0c1"],
      ["r1c0", "r1c1"],
      ["r2c0", "r2c1"],
    ]),
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;
  const sourceJson = JSON.stringify(source);

  // "Insert row below" on the first row of a 3-row table: the new row lands at
  // index 1, not at index 2 with a gap in between.
  controller.insertTableRow(1, 0, "below");
  await vue.nextTick();
  let target = controller.model.nodes[1];
  assert.equal(target.rows.length, 4, "the table gains exactly one row");
  assert.deepEqual(cellTexts(target), [
    ["r0c0", "r0c1"],
    ["", ""],
    ["r1c0", "r1c1"],
    ["r2c0", "r2c1"],
  ]);
  assert.equal(JSON.stringify(source), sourceJson, "the frozen source model stays untouched");

  // "Insert row above" on the same reference row keeps it as the reference.
  controller.insertTableRow(1, 2, "above");
  await vue.nextTick();
  target = controller.model.nodes[1];
  assert.deepEqual(cellTexts(target), [
    ["r0c0", "r0c1"],
    ["", ""],
    ["", ""],
    ["r1c0", "r1c1"],
    ["r2c0", "r2c1"],
  ]);

  // "Insert column right" of the first column of a 2-column table.
  controller.insertTableColumn(1, 0, "right");
  await vue.nextTick();
  target = controller.model.nodes[1];
  assert.equal(target.rows[0].cells.length, 3, "the table gains exactly one column");
  assert.deepEqual(cellTexts(target), [
    ["r0c0", "", "r0c1"],
    ["", "", ""],
    ["", "", ""],
    ["r1c0", "", "r1c1"],
    ["r2c0", "", "r2c1"],
  ]);
  assert.equal(controller.canUndo, true);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("paragraph-level commands apply to every paragraph of the selection", async () => {
  const source = deepFreeze(starterModel([
    paragraph("alpha"),
    paragraph("bravo"),
    paragraph("charlie"),
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;

  controller.selectParagraph(0);
  await vue.nextTick();
  controller.setActiveTextRange({ start: locationAt(0, 1), end: locationAt(2, 4) });
  await vue.nextTick();

  controller.setAlignment("center");
  await vue.nextTick();
  assert.deepEqual(
    controller.model.nodes.map((node) => node.style?.align),
    ["center", "center", "center"],
    "alignment reaches all three selected paragraphs",
  );

  controller.setLineSpacing(2);
  await vue.nextTick();
  assert.deepEqual(
    controller.model.nodes.map((node) => node.style?.spacing?.lineTwips),
    [480, 480, 480],
  );

  controller.setHeading(2);
  await vue.nextTick();
  assert.deepEqual(
    controller.model.nodes.map((node) => node.style?.headingLevel),
    [2, 2, 2],
  );

  controller.toggleList("unordered");
  await vue.nextTick();
  assert.deepEqual(
    controller.model.nodes.map((node) => node.style?.numbering?.numId),
    [1, 1, 1],
    "the list toggle applies to the whole selection",
  );

  // A collapsed caret still commands exactly one paragraph.
  controller.selectParagraph(1);
  await vue.nextTick();
  controller.setActiveTextRange({ start: locationAt(1, 2), end: locationAt(1, 2) });
  await vue.nextTick();
  controller.setAlignment("right");
  await vue.nextTick();
  assert.deepEqual(
    controller.model.nodes.map((node) => node.style?.align),
    ["center", "right", "center"],
  );

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("insertTable lands at the caret instead of the end of the document", async () => {
  const buildSource = () => deepFreeze(starterModel([
    paragraph("alpha"),
    paragraph("beta gamma"),
    paragraph("delta"),
  ]));

  const middle = await mountController(buildSource());
  middle.controller.selectParagraph(1);
  await vue.nextTick();
  middle.controller.setActiveTextRange({ start: locationAt(1, 5), end: locationAt(1, 5) });
  await vue.nextTick();
  middle.controller.insertTable();
  await vue.nextTick();

  assert.deepEqual(
    middle.controller.model.nodes.map((node) => node.type),
    ["paragraph", "paragraph", "table", "paragraph", "paragraph"],
    "the table splits the caret paragraph instead of appending at the end",
  );
  assert.deepEqual(
    middle.controller.model.nodes.map(paragraphText),
    ["alpha", "beta ", "", "gamma", "delta"],
  );
  const inserted = middle.controller.model.nodes[2];
  assert.equal(inserted.rows.length, 3);
  assert.equal(inserted.rows[0].cells.length, 3);
  middle.app.unmount();
  assert.deepEqual(middle.warnings, []);

  const start = await mountController(buildSource());
  start.controller.selectParagraph(1);
  await vue.nextTick();
  start.controller.setActiveTextRange({ start: locationAt(1, 0), end: locationAt(1, 0) });
  await vue.nextTick();
  start.controller.insertTable();
  await vue.nextTick();
  assert.deepEqual(
    start.controller.model.nodes.map((node) => node.type),
    ["paragraph", "table", "paragraph", "paragraph"],
    "a caret at the paragraph start inserts the table before it",
  );
  start.app.unmount();
  assert.deepEqual(start.warnings, []);

  const end = await mountController(buildSource());
  end.controller.selectParagraph(1);
  await vue.nextTick();
  end.controller.setActiveTextRange({ start: locationAt(1, 10), end: locationAt(1, 10) });
  await vue.nextTick();
  end.controller.insertTable();
  await vue.nextTick();
  assert.deepEqual(
    end.controller.model.nodes.map((node) => node.type),
    ["paragraph", "paragraph", "table", "paragraph"],
    "a caret at the paragraph end inserts the table after it",
  );
  end.app.unmount();
  assert.deepEqual(end.warnings, []);
});

test("diagonal border presets report unsupported instead of silently doing nothing", async () => {
  const source = deepFreeze(starterModel([
    paragraph("alpha"),
    table([
      ["a", "b"],
      ["c", "d"],
    ]),
  ]));
  const mounted = await mountController(source);
  const { controller } = mounted;

  controller.selectParagraph(0);
  await vue.nextTick();
  const beforeParagraphPreset = JSON.stringify(controller.model);
  controller.applyBorderPreset("diagonal-down");
  await vue.nextTick();
  assert.match(controller.status, /Unsupported: diagonal paragraph borders/);
  assert.equal(
    JSON.stringify(controller.model),
    beforeParagraphPreset,
    "an unsupported preset must not modify the model",
  );

  controller.selectTableCell(1, 0, 0);
  await vue.nextTick();
  controller.applyBorderPreset("diagonal-up");
  await vue.nextTick();
  assert.match(controller.status, /Unsupported: diagonal table-cell borders/);
  assert.equal(JSON.stringify(controller.model), beforeParagraphPreset);

  // Supported presets keep working on the same target.
  controller.applyBorderPreset("all");
  await vue.nextTick();
  assert.ok(controller.model.nodes[1].style?.borders?.top, "supported presets still apply");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});
