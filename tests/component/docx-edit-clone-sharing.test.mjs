import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
} from "./vue-test-renderer.mjs";

const { useDocxEditor } = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

const IMAGE_BYTES = 2 * 1024 * 1024;
const PAGE_COUNT = 200;

function imageRun() {
  return {
    type: "image",
    src: "data:image/png;base64,AAAA",
    partName: "word/media/image1.png",
    contentType: "image/png",
    widthPx: 240,
    heightPx: 160,
    floating: { xPx: 72, yPx: 72, wrapType: "square" },
    data: new Uint8Array(IMAGE_BYTES),
  };
}

function createTableNode() {
  return {
    type: "table",
    rows: [
      {
        type: "table-row",
        cells: [
          {
            type: "table-cell",
            nodes: [
              { type: "paragraph", children: [{ type: "text", text: "cell" }] },
            ],
          },
        ],
      },
    ],
  };
}

/**
 * A long document with a header section, an image, and a table — every shape the
 * interactive commands edit.
 */
function createLongModel({ pages = PAGE_COUNT } = {}) {
  const base = createBlankDocumentModel();
  const nodes = [];
  for (let index = 0; index < pages; index += 1) {
    const children = [{ type: "text", text: `Paragraph ${index} body text` }];
    if (index === 1) children.push(imageRun());
    nodes.push({ type: "paragraph", children });
  }
  nodes.splice(2, 0, createTableNode());

  return {
    ...base,
    nodes,
    metadata: {
      ...base.metadata,
      headerSections: [
        {
          partName: "word/header1.xml",
          referenceType: "default",
          nodes: [{ type: "paragraph", children: [{ type: "text", text: "Header" }] }],
        },
        {
          partName: "word/header2.xml",
          referenceType: "first",
          nodes: [{ type: "paragraph", children: [{ type: "text", text: "First" }] }],
        },
      ],
    },
  };
}

function imageBytes(model) {
  const bytes = [];
  const visit = (nodes) => {
    for (const node of nodes) {
      if (node.type === "paragraph") {
        for (const child of node.children) {
          if (child.type === "image" && child.data) bytes.push(child.data);
        }
      } else {
        for (const row of node.rows) {
          for (const cell of row.cells) visit(cell.nodes);
        }
      }
    }
  };
  visit(model.nodes);
  return bytes;
}

function changedNodeIndexes(before, after) {
  const changed = [];
  for (let index = 0; index < before.nodes.length; index += 1) {
    if (before.nodes[index] !== after.nodes[index]) changed.push(index);
  }
  return changed;
}

async function mountEditor(starterModel) {
  let editor;
  const Harness = vue.defineComponent({
    setup() {
      editor = useDocxEditor({ starterModel });
      return () => vue.h("div", `${editor.canUndo}`);
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, editor };
}

test("interactive commands edit a bounded node set and keep image bytes shared", async () => {
  const starterModel = createLongModel();
  const { app, editor } = await mountEditor(starterModel);

  try {
    const originalBytes = imageBytes(editor.model);
    const untouchedHeaderNode = editor.model.metadata.headerSections[1].nodes[0];
    assert.equal(originalBytes.length, 1);

    const commands = [
      [
        "commitParagraphText",
        () => editor.commitParagraphText(0, "typed"),
        () => [0],
      ],
      [
        "dev commitParagraphText",
        () => editor.commitParagraphText(0, "typed again"),
        () => [0],
      ],
      [
        "commitTableCellText",
        () => editor.commitTableCellText(2, 0, 0, "cell edited"),
        () => [2],
      ],
      [
        "setAlignment",
        () => {
          editor.selectParagraph(3);
          editor.setAlignment("center");
        },
        () => [3],
      ],
      [
        "toggleBold",
        () => {
          editor.selectParagraph(3);
          editor.setActiveTextRange({
            start: { location: { kind: "paragraph", nodeIndex: 3 }, offset: 0 },
            end: { location: { kind: "paragraph", nodeIndex: 3 }, offset: 4 },
          });
          editor.toggleBold();
        },
        () => [3],
      ],
      [
        "appendParagraph",
        () => editor.appendParagraph("appended"),
        () => [],
      ],
      [
        "resizeImage",
        () => editor.resizeImage({ kind: "paragraph", nodeIndex: 1, childIndex: 1 }, 200, 120),
        () => [1],
      ],
      [
        "insertTableRow",
        () => editor.insertTableRow(2, 0),
        () => [2],
      ],
      [
        "commitSectionParagraphText",
        () => editor.commitSectionParagraphText(
          { region: "header", partName: "word/header1.xml", nodeIndex: 0 },
          "Header edited",
        ),
        () => [],
      ],
    ];

    for (const [name, run, expectedChanged] of commands) {
      const before = editor.model;
      run();
      await vue.nextTick();
      const after = editor.model;

      assert.notEqual(after, before, `${name} must commit a new model`);
      const changed = changedNodeIndexes(before, after);
      assert.ok(
        changed.length < PAGE_COUNT,
        `${name} rebuilt ${changed.length} of ${before.nodes.length} body nodes`,
      );
      const expected = new Set(expectedChanged(before, after));
      for (const index of changed) {
        assert.ok(expected.has(index), `${name} unexpectedly rebuilt body node ${index}`);
      }
      const expectedBytes = imageBytes(before);
      const nextBytes = imageBytes(after);
      assert.equal(
        nextBytes.length,
        expectedBytes.length,
        `${name} changed the image run count`,
      );
      nextBytes.forEach((buffer, imageIndex) => {
        assert.equal(
          buffer,
          expectedBytes[imageIndex],
          `${name} copied image bytes`,
        );
      });
    }

    assert.equal(editor.model.nodes.length, PAGE_COUNT + 2);
    assert.equal(
      editor.model.nodes.at(-1).children[0].text,
      "appended",
    );
    assert.equal(
      editor.model.metadata.headerSections[0].nodes[0].children[0].text,
      "Header edited",
    );
    assert.equal(
      editor.model.metadata.headerSections[1].nodes[0],
      untouchedHeaderNode,
      "nodes of an untouched header section stay shared",
    );
    assert.equal(editor.canUndo, true);

    editor.undo();
    await vue.nextTick();
    assert.equal(
      editor.model.metadata.headerSections[0].nodes[0].children[0].text,
      "Header",
    );
    assert.deepEqual(imageBytes(editor.model), originalBytes);
  } finally {
    app.unmount();
  }
});

test("typing and formatting a long document never copies image bytes", async () => {
  const starterModel = createLongModel();
  const { app, editor } = await mountEditor(starterModel);

  try {
    const bytes = imageBytes(editor.model)[0];
    const startedAt = performance.now();
    let rebuilt = 0;

    for (let index = 0; index < 100; index += 1) {
      const before = editor.model;
      editor.commitParagraphText(0, `typed ${index}`);
      await vue.nextTick();
      rebuilt += changedNodeIndexes(before, editor.model).length;
      assert.equal(
        imageBytes(editor.model)[0],
        bytes,
        `keystroke ${index} copied image bytes`,
      );
    }

    const elapsedMs = performance.now() - startedAt;
    assert.equal(rebuilt, 100, `100 keystrokes rebuilt ${rebuilt} body nodes`);
    assert.ok(
      elapsedMs < 15_000,
      `100 keystrokes took ${elapsedMs.toFixed(0)} ms`,
    );
    assert.equal(editor.model.nodes.length, PAGE_COUNT + 1);
  } finally {
    app.unmount();
  }
});
