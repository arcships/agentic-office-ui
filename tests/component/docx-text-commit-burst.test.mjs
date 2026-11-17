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
  return node?.type === "paragraph"
    ? node.children.map((child) => (child.type === "text" ? child.text : "")).join("")
    : "";
}

function modelWithParagraphs(texts) {
  return {
    ...createBlankDocumentModel(),
    nodes: texts.map((text) => ({
      type: "paragraph",
      children: [{ type: "text", text }],
    })),
  };
}

function modelWithTable() {
  const base = createBlankDocumentModel();
  return {
    ...base,
    nodes: [
      { type: "paragraph", children: [{ type: "text", text: "before" }] },
      {
        type: "table",
        rows: [
          {
            type: "table-row",
            cells: [
              {
                type: "table-cell",
                nodes: [
                  { type: "paragraph", children: [{ type: "text", text: "cell a" }] },
                ],
              },
              {
                type: "table-cell",
                nodes: [
                  { type: "paragraph", children: [{ type: "text", text: "cell b" }] },
                ],
              },
            ],
          },
        ],
      },
    ],
  };
}

function modelWithHeader() {
  const base = createBlankDocumentModel();
  return {
    ...base,
    nodes: [{ type: "paragraph", children: [{ type: "text", text: "body" }] }],
    metadata: {
      ...base.metadata,
      headerSections: [
        {
          partName: "word/header1.xml",
          referenceType: "default",
          nodes: [{ type: "paragraph", children: [{ type: "text", text: "Header" }] }],
        },
      ],
    },
  };
}

async function mountEditor(starterModel) {
  let editor;
  const Harness = vue.defineComponent({
    setup() {
      editor = useDocxEditor({ starterModel });
      return () => vue.h("div", String(editor.canUndo));
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, editor };
}

async function countUndos(editor, limit = 100) {
  let count = 0;
  while (editor.canUndo && count < limit) {
    editor.undo();
    count += 1;
    await vue.nextTick();
  }
  return count;
}

test("consecutive keystrokes in one paragraph produce one undo entry", async () => {
  const { app, editor } = await mountEditor(modelWithParagraphs(["start", "other"]));

  try {
    for (let index = 0; index < 20; index += 1) {
      editor.commitParagraphText(0, `start ${index}`);
      await vue.nextTick();
    }

    assert.equal(paragraphText(editor.model.nodes[0]), "start 19");
    assert.equal(editor.canUndo, true);

    editor.undo();
    await vue.nextTick();
    assert.equal(
      paragraphText(editor.model.nodes[0]),
      "start",
      "one undo restores the paragraph as it was before the burst",
    );
    assert.equal(editor.canUndo, false, "the whole burst is one undo entry");

    editor.redo();
    await vue.nextTick();
    assert.equal(paragraphText(editor.model.nodes[0]), "start 19");
  } finally {
    app.unmount();
  }
});

test("keystrokes in different paragraphs open separate undo entries", async () => {
  const { app, editor } = await mountEditor(modelWithParagraphs(["first", "second"]));

  try {
    editor.commitParagraphText(0, "first typed");
    await vue.nextTick();
    editor.commitParagraphText(1, "second typed");
    await vue.nextTick();

    assert.equal(await countUndos(editor), 2);

    editor.redo();
    await vue.nextTick();
    assert.equal(paragraphText(editor.model.nodes[0]), "first typed");
    assert.equal(paragraphText(editor.model.nodes[1]), "second");
  } finally {
    app.unmount();
  }
});

test("a command between keystrokes starts a new undo entry", async () => {
  const { app, editor } = await mountEditor(modelWithParagraphs(["alpha"]));

  try {
    editor.commitParagraphText(0, "alpha one");
    await vue.nextTick();
    editor.selectParagraph(0);
    editor.setAlignment("center");
    await vue.nextTick();
    editor.commitParagraphText(0, "alpha two");
    await vue.nextTick();

    assert.equal(await countUndos(editor), 3);
  } finally {
    app.unmount();
  }
});

test("a selection change between keystrokes starts a new undo entry", async () => {
  const { app, editor } = await mountEditor(modelWithParagraphs(["alpha", "beta"]));

  try {
    editor.commitParagraphText(0, "alpha one");
    await vue.nextTick();
    editor.setActiveTextRange({
      start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
      end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 5 },
    });
    await vue.nextTick();
    editor.commitParagraphText(0, "alpha two");
    await vue.nextTick();

    assert.equal(await countUndos(editor), 2);
  } finally {
    app.unmount();
  }
});

test("moving the caret to another paragraph starts a new undo entry", async () => {
  const { app, editor } = await mountEditor(modelWithParagraphs(["alpha", "beta"]));

  try {
    editor.commitParagraphText(0, "alpha one");
    await vue.nextTick();
    editor.setActiveTextRange({
      start: { location: { kind: "paragraph", nodeIndex: 1 }, offset: 2 },
      end: { location: { kind: "paragraph", nodeIndex: 1 }, offset: 2 },
    });
    await vue.nextTick();
    editor.commitParagraphText(0, "alpha two");
    await vue.nextTick();

    assert.equal(await countUndos(editor), 2);
  } finally {
    app.unmount();
  }
});

test("the caret following committed text keeps the burst open", async () => {
  const { app, editor } = await mountEditor(modelWithParagraphs(["alpha"]));

  try {
    editor.commitParagraphText(0, "alpha one");
    await vue.nextTick();
    // Typing moves the caret inside the same paragraph: still one burst.
    editor.setActiveTextRange({
      start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 9 },
      end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 9 },
    });
    await vue.nextTick();
    editor.commitParagraphText(0, "alpha one more");
    await vue.nextTick();

    assert.equal(await countUndos(editor), 1);
  } finally {
    app.unmount();
  }
});

test("endTextCommitBurst closes the burst (blur boundary)", async () => {
  const { app, editor } = await mountEditor(modelWithParagraphs(["alpha"]));

  try {
    editor.commitParagraphText(0, "alpha one");
    await vue.nextTick();
    editor.endTextCommitBurst();
    editor.commitParagraphText(0, "alpha two");
    await vue.nextTick();

    assert.equal(await countUndos(editor), 2);
  } finally {
    app.unmount();
  }
});

test("undo closes the burst so later typing is a new undoable step", async () => {
  const { app, editor } = await mountEditor(modelWithParagraphs(["alpha"]));

  try {
    editor.commitParagraphText(0, "alpha one");
    await vue.nextTick();
    editor.commitParagraphText(0, "alpha two");
    await vue.nextTick();

    editor.undo();
    await vue.nextTick();
    assert.equal(paragraphText(editor.model.nodes[0]), "alpha");
    assert.equal(editor.canRedo, true);

    editor.redo();
    await vue.nextTick();
    assert.equal(paragraphText(editor.model.nodes[0]), "alpha two");

    // Typing after redo must open a new entry instead of folding into the
    // restored burst entry.
    editor.commitParagraphText(0, "alpha three");
    await vue.nextTick();
    assert.equal(editor.canRedo, false, "typing after redo clears redo");

    editor.undo();
    await vue.nextTick();
    assert.equal(paragraphText(editor.model.nodes[0]), "alpha two");
    editor.undo();
    await vue.nextTick();
    assert.equal(paragraphText(editor.model.nodes[0]), "alpha");
    assert.equal(editor.canUndo, false);
  } finally {
    app.unmount();
  }
});

test("table cell typing coalesces per cell", async () => {
  const { app, editor } = await mountEditor(modelWithTable());

  try {
    for (let index = 0; index < 5; index += 1) {
      editor.commitTableCellText(1, 0, 0, `cell a ${index}`);
      await vue.nextTick();
    }
    editor.commitTableCellText(1, 0, 1, "cell b typed");
    await vue.nextTick();

    const table = editor.model.nodes[1];
    assert.equal(table.rows[0].cells[0].nodes[0].children[0].text, "cell a 4");
    assert.equal(table.rows[0].cells[1].nodes[0].children[0].text, "cell b typed");

    editor.undo();
    await vue.nextTick();
    assert.equal(
      editor.model.nodes[1].rows[0].cells[1].nodes[0].children[0].text,
      "cell b",
      "the second cell was its own undo entry",
    );
    editor.undo();
    await vue.nextTick();
    assert.equal(
      editor.model.nodes[1].rows[0].cells[0].nodes[0].children[0].text,
      "cell a",
      "the five commits in the first cell were one entry",
    );
    assert.equal(editor.canUndo, false);
  } finally {
    app.unmount();
  }
});

test("header typing coalesces and stays separate from body typing", async () => {
  const { app, editor } = await mountEditor(modelWithHeader());

  try {
    editor.commitSectionParagraphText(
      { region: "header", partName: "word/header1.xml", nodeIndex: 0 },
      "Header one",
    );
    await vue.nextTick();
    editor.commitSectionParagraphText(
      { region: "header", partName: "word/header1.xml", nodeIndex: 0 },
      "Header two",
    );
    await vue.nextTick();
    editor.commitParagraphText(0, "body typed");
    await vue.nextTick();

    const headerText = () =>
      editor.model.metadata.headerSections[0].nodes[0].children[0].text;
    assert.equal(headerText(), "Header two");

    editor.undo();
    await vue.nextTick();
    assert.equal(paragraphText(editor.model.nodes[0]), "body", "body typing undone");
    assert.equal(headerText(), "Header two", "header burst untouched by body undo");

    editor.undo();
    await vue.nextTick();
    assert.equal(headerText(), "Header", "the header burst is one entry");
    assert.equal(editor.canUndo, false);
  } finally {
    app.unmount();
  }
});
