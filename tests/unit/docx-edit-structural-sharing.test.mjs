import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);

const PAGE_COUNT = 200;
const IMAGE_BYTES = 5 * 1024 * 1024;

function imageRun(index) {
  return {
    type: "image",
    src: `data:image/png;base64,image-${index}`,
    partName: `word/media/image${index}.png`,
    contentType: "image/png",
    widthPx: 640,
    heightPx: 480,
    data: new Uint8Array(IMAGE_BYTES),
  };
}

function createLongModel({ pages = PAGE_COUNT, imageCount = 4 } = {}) {
  const base = core.createBlankDocumentModel();
  const nodes = [];
  for (let index = 0; index < pages; index += 1) {
    const children = [{ type: "text", text: `Paragraph ${index} body text` }];
    if (imageCount > 0 && index % Math.ceil(pages / imageCount) === 0) {
      children.push(imageRun(index / Math.ceil(pages / imageCount)));
    }
    nodes.push({ type: "paragraph", children });
  }
  return { ...base, nodes };
}

function createTableModel() {
  const base = core.createBlankDocumentModel();
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
                  {
                    type: "paragraph",
                    children: [{ type: "text", text: "cell text" }],
                  },
                ],
              },
            ],
          },
        ],
      },
      { type: "paragraph", children: [{ type: "text", text: "after" }] },
    ],
  };
}

function textRunsOf(model) {
  const runs = [];
  for (const node of model.nodes) {
    if (node.type !== "paragraph") continue;
    for (const child of node.children) runs.push(child);
  }
  return runs;
}

function serializeBytes(value) {
  return new TextEncoder().encode(JSON.stringify(value)).byteLength;
}

function countSerializations(run) {
  const originalStringify = JSON.stringify;
  let serializations = 0;
  JSON.stringify = function countedStringify(...args) {
    serializations += 1;
    return originalStringify.apply(this, args);
  };
  try {
    run();
  } finally {
    JSON.stringify = originalStringify;
  }
  return serializations;
}

test("cloneDocModel stays isolated while edit-path node clones share image bytes", () => {
  const model = createLongModel({ pages: 2, imageCount: 1 });
  const cloned = core.cloneDocModel(model);

  assert.notEqual(cloned, model);
  assert.notEqual(cloned.nodes[0], model.nodes[0]);
  assert.notEqual(cloned.nodes[0].children[0], model.nodes[0].children[0]);
  assert.notEqual(
    cloned.nodes[0].children[1].data,
    model.nodes[0].children[1].data,
    "the full clone keeps its own copy of the encoded bytes",
  );
  assert.deepEqual(
    cloned.nodes[0].children[1].data,
    model.nodes[0].children[1].data,
    "the copied bytes have the same content",
  );
  assert.equal(textRunsOf(cloned).length, textRunsOf(model).length);

  const editable = core.cloneDocNode(model.nodes[0]);
  assert.notEqual(editable, model.nodes[0]);
  assert.equal(
    editable.children[1].data,
    model.nodes[0].children[1].data,
    "the edit-path node clone must not duplicate image bytes",
  );
});

test("cloneDocModelShallow shares nodes and metadata until a node is replaced", () => {
  const model = createLongModel({ pages: 4, imageCount: 1 });
  const shell = core.cloneDocModelShallow(model);

  assert.notEqual(shell, model);
  assert.notEqual(shell.nodes, model.nodes);
  assert.notEqual(shell.metadata, model.metadata);
  assert.equal(shell.metadata.warnings, model.metadata.warnings);
  assert.equal(shell.nodes[1], model.nodes[1], "untouched nodes stay shared");
  assert.equal(shell.nodes[2], model.nodes[2]);

  const editable = core.cloneDocNode(shell.nodes[0]);
  assert.notEqual(editable, model.nodes[0]);
  assert.notEqual(editable.children, model.nodes[0].children);
  assert.equal(
    editable.children[1].data,
    model.nodes[0].children[1].data,
    "editing a node must not duplicate its image bytes",
  );

  shell.nodes[0] = editable;
  editable.children = [{ type: "text", text: "edited" }];

  assert.equal(model.nodes[0].children[0].text, "Paragraph 0 body text");
  assert.equal(shell.nodes[1], model.nodes[1]);
});

test("single-paragraph edit commands rebuild one node and keep siblings shared", () => {
  const model = createLongModel({ pages: 6, imageCount: 2 });

  const cases = [
    ["updateParagraphText", core.updateParagraphText(model, 0, "typed")],
    ["setParagraphAlignment", core.setParagraphAlignment(model, 0, "center")],
    ["applyRunStyle", core.applyRunStyle(model, 0, 0, { bold: true })],
  ];

  for (const [command, next] of cases) {
    assert.notEqual(next, model, `${command} must return a new model`);
    assert.notEqual(next.nodes[0], model.nodes[0], `${command} must clone the edited node`);
    for (let index = 1; index < model.nodes.length; index += 1) {
      assert.equal(
        next.nodes[index],
        model.nodes[index],
        `${command} must keep node ${index} shared`,
      );
    }
    assert.equal(next.metadata.warnings, model.metadata.warnings);
    assert.equal(
      model.nodes[0].children[0].text,
      "Paragraph 0 body text",
      `${command} must not mutate the source paragraph`,
    );
  }

  const typed = core.updateParagraphText(model, 0, "typed");
  assert.equal(typed.nodes[0].children[0].text, "typed");
  assert.equal(
    typed.nodes[3].children[1].data,
    model.nodes[3].children[1].data,
    "image runs on untouched nodes keep sharing their bytes",
  );
});

test("table-cell text edits clone only the edited table", () => {
  const model = createTableModel();
  const next = core.updateTableCellText(model, 1, 0, 0, "changed");

  assert.equal(next.nodes[0], model.nodes[0], "sibling paragraphs stay shared");
  assert.equal(next.nodes[2], model.nodes[2]);
  assert.notEqual(next.nodes[1], model.nodes[1], "the edited table is cloned");
  assert.equal(
    next.nodes[1].rows[0].cells[0].nodes[0].children[0].text,
    "changed",
  );
  assert.equal(
    model.nodes[1].rows[0].cells[0].nodes[0].children[0].text,
    "cell text",
    "the source table is untouched",
  );
});

test("typing 100 characters into a 200-page document rebuilds one node per keystroke", () => {
  const model = createLongModel();
  let current = model;
  let rebuiltNodes = 0;
  let copiedImageBytes = 0;
  const startedAt = performance.now();

  for (let index = 0; index < 100; index += 1) {
    const previous = current;
    current = core.updateParagraphText(previous, 0, `typed ${index}`);
    rebuiltNodes += current.nodes.filter(
      (node, nodeIndex) => node !== previous.nodes[nodeIndex],
    ).length;
    for (let nodeIndex = 1; nodeIndex < previous.nodes.length; nodeIndex += 1) {
      const before = previous.nodes[nodeIndex];
      const after = current.nodes[nodeIndex];
      if (before.type !== "paragraph" || after.type !== "paragraph") continue;
      for (let runIndex = 0; runIndex < before.children.length; runIndex += 1) {
        const beforeRun = before.children[runIndex];
        const afterRun = after.children[runIndex];
        if (beforeRun.type === "image" && afterRun.type === "image" && beforeRun.data !== afterRun.data) {
          copiedImageBytes += afterRun.data?.byteLength ?? 0;
        }
      }
    }
  }
  const elapsedMs = performance.now() - startedAt;

  assert.equal(rebuiltNodes, 100, "each keystroke must rebuild exactly the edited node");
  assert.equal(copiedImageBytes, 0, "typing must not copy image bytes");
  assert.equal(current.nodes[0].children[0].text, "typed 99");
  assert.equal(current.nodes[1], model.nodes[1], "untouched paragraphs stay shared");
  assert.ok(
    elapsedMs < 5_000,
    `100 keystrokes on a ${PAGE_COUNT}-page document took ${elapsedMs.toFixed(0)} ms`,
  );
});

test("history budgets measure each retained snapshot once", () => {
  const selection = { kind: "paragraph", nodeIndex: 0 };
  const pushes = 60;

  function runHistory(maxEntries) {
    let state = core.createEditorStateV2({
      model: createLongModel({ pages: 20, imageCount: 0 }),
      selection,
      historyMaxEntries: maxEntries,
      historyMaxBytes: 32 * 1024 * 1024,
    });
    let model = state.model;
    const deltas = [];
    let serializations = 0;
    const originalStringify = JSON.stringify;
    JSON.stringify = function countedStringify(...args) {
      serializations += 1;
      return originalStringify.apply(this, args);
    };
    try {
      for (let index = 0; index < pushes; index += 1) {
        const before = serializations;
        model = core.updateParagraphText(model, 0, `keystroke ${index}`);
        state = core.applyEditorTransactionV2(state, { model, selection });
        deltas.push(serializations - before);
      }
    } finally {
      JSON.stringify = originalStringify;
    }
    return { state, serializations, deltas };
  }

  const small = runHistory(4);
  const large = runHistory(64);

  assert.equal(small.state.history.past.length, 4);
  assert.equal(large.state.history.past.length, pushes);
  assert.ok(
    Math.max(...large.deltas) <= 5,
    `a push re-serialised ${Math.max(...large.deltas)} times with ${pushes} retained entries`,
  );
  assert.ok(
    Math.max(...large.deltas) <= Math.max(...small.deltas) + 1,
    `serialisations grew with retained entries: small ${small.serializations}, large ${large.serializations}`,
  );
  assert.equal(
    large.state.history.pastBytes,
    large.state.history.past.reduce(
      (total, entry) => total + serializeBytes(entry),
      0,
    ),
    "the cached estimate must match a full serialisation of each entry",
  );
});
