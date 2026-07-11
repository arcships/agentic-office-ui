import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
} from "./vue-test-renderer.mjs";

const { useDocxEditor } = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

function estimateBytes(value) {
  return new TextEncoder().encode(JSON.stringify(value)).byteLength;
}

function modelWithText(text) {
  return {
    ...createBlankDocumentModel(),
    nodes: [{ type: "paragraph", children: [{ type: "text", text }] }],
  };
}

function paragraphText(controller) {
  const paragraph = controller.model.nodes[0];
  return paragraph?.type === "paragraph"
    ? paragraph.children.map((child) => child.type === "text" ? child.text : "").join("")
    : "";
}

async function commitSeries(controller, prefix, count, textLength) {
  for (let index = 0; index < count; index += 1) {
    controller.commitParagraphText(
      0,
      `${prefix}-${index}:`.padEnd(textLength, String.fromCharCode(65 + index)),
    );
    await vue.nextTick();
  }
}

async function countUndos(controller) {
  let count = 0;
  while (controller.canUndo && count < 50) {
    controller.undo();
    count += 1;
    await vue.nextTick();
  }
  return count;
}

test("useDocxEditor applies byte budgets per instance and preserves undo semantics", async () => {
  const selection = { kind: "paragraph", nodeIndex: 0 };
  const largeTextLength = 4 * 1024;
  const oneLargeSnapshotBytes = estimateBytes({
    model: modelWithText("A".repeat(largeTextLength)),
    selection,
  });
  let lowBudgetEditor;
  let highBudgetEditor;
  const Harness = vue.defineComponent({
    setup() {
      lowBudgetEditor = useDocxEditor({
        historyMaxBytes: oneLargeSnapshotBytes + 64,
        historyMaxEntries: 20,
        starterModel: modelWithText("initial"),
      });
      highBudgetEditor = useDocxEditor({
        historyMaxBytes: 2 * 1024 * 1024,
        historyMaxEntries: 20,
        starterModel: modelWithText("initial"),
      });
      return () => vue.h("div", `${lowBudgetEditor.canUndo}:${highBudgetEditor.canUndo}`);
    },
  });
  const mounted = await mount(Harness);

  try {
    await commitSeries(lowBudgetEditor, "low", 6, largeTextLength);
    await commitSeries(highBudgetEditor, "high", 6, largeTextLength);

    const lowUndoCount = await countUndos(lowBudgetEditor);
    const highUndoCount = await countUndos(highBudgetEditor);
    assert.ok(lowUndoCount >= 1, "the latest low-budget edit must remain undoable");
    assert.ok(lowUndoCount < highUndoCount);
    assert.equal(highUndoCount, 6);
  } finally {
    mounted.app.unmount();
  }
});

test("useDocxEditor keeps one oversized undo item and clears redo on a new edit", async () => {
  const hugeText = "超".repeat(20 * 1024);
  let editor;
  const Harness = vue.defineComponent({
    setup() {
      editor = useDocxEditor({
        historyMaxBytes: 512,
        historyMaxEntries: 10,
        starterModel: modelWithText(hugeText),
      });
      return () => vue.h("div", `${editor.canUndo}:${editor.canRedo}`);
    },
  });
  const mounted = await mount(Harness);

  try {
    editor.commitParagraphText(0, "small");
    await vue.nextTick();
    assert.equal(editor.canUndo, true);
    editor.undo();
    await vue.nextTick();
    assert.equal(paragraphText(editor), hugeText);
    assert.equal(editor.canRedo, true);

    editor.commitParagraphText(0, "new branch");
    await vue.nextTick();
    assert.equal(editor.canRedo, false);
    assert.equal(editor.canUndo, true);
  } finally {
    mounted.app.unmount();
  }
});
