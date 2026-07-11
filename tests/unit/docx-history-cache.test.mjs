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

function modelWithText(text) {
  return {
    ...core.createBlankDocumentModel(),
    nodes: [{ type: "paragraph", children: [{ type: "text", text }] }],
  };
}

function textFromState(state) {
  return state.model.nodes[0]?.children[0]?.text;
}

function estimateBytes(value) {
  return new TextEncoder().encode(JSON.stringify(value)).byteLength;
}

function applyText(state, text) {
  return core.applyEditorTransactionV2(state, {
    type: "replace-text",
    model: modelWithText(text),
  });
}

test("DOCX core history evicts deterministically by count and estimated bytes", () => {
  const selection = { kind: "paragraph", nodeIndex: 0 };
  const largeText = "A".repeat(4 * 1024);
  const oneLargeSnapshotBytes = estimateBytes({
    model: modelWithText(largeText),
    selection,
  });
  let state = core.createEditorStateV2({
    historyMaxBytes: oneLargeSnapshotBytes + 64,
    historyMaxEntries: 3,
    model: modelWithText("initial"),
    selection,
  });

  for (const marker of ["A", "B", "C", "D"]) {
    state = applyText(state, marker.repeat(4 * 1024));
  }

  assert.equal(state.history.past.length, 1);
  assert.ok(state.history.pastBytes <= state.history.maxBytes);
  assert.equal(
    state.history.pastBytes,
    state.history.past.reduce(
      (total, entry) => total + estimateBytes(entry),
      0,
    ),
  );

  state = core.undoEditorStateV2(state);
  assert.equal(textFromState(state), "C".repeat(4 * 1024));
  assert.equal(state.history.future.length, 1);
  assert.ok(state.history.futureBytes <= state.history.maxBytes);

  state = core.redoEditorStateV2(state);
  assert.equal(textFromState(state), "D".repeat(4 * 1024));
  state = core.undoEditorStateV2(state);
  state = applyText(state, "new branch");
  assert.equal(state.history.future.length, 0, "a new edit must clear redo history");
  assert.equal(state.history.futureBytes, 0);
});

test("DOCX core history retains the newest oversized snapshot and isolates instances", () => {
  const selection = { kind: "paragraph", nodeIndex: 0 };
  const hugeText = "超".repeat(20 * 1024);
  let oversized = core.createEditorStateV2({
    historyMaxBytes: 512,
    historyMaxEntries: 10,
    model: modelWithText(hugeText),
    selection,
  });
  oversized = applyText(oversized, "small");
  assert.equal(oversized.history.past.length, 1);
  assert.ok(oversized.history.pastBytes > oversized.history.maxBytes);
  oversized = core.undoEditorStateV2(oversized);
  assert.equal(textFromState(oversized), hugeText);

  let low = core.createEditorStateV2({
    historyMaxBytes: 1024 * 1024,
    historyMaxEntries: 1,
    model: modelWithText("initial"),
    selection,
  });
  let high = core.createEditorStateV2({
    historyMaxBytes: 1024 * 1024,
    historyMaxEntries: 4,
    model: modelWithText("initial"),
    selection,
  });
  for (let index = 0; index < 5; index += 1) {
    low = applyText(low, `low-${index}`);
    high = applyText(high, `high-${index}`);
  }
  assert.equal(low.history.past.length, 1);
  assert.equal(high.history.past.length, 4);
});

test("DOCX thumbnail surface cache combines byte/count LRU and owns its state", () => {
  const first = new core.DocxThumbnailSurfaceCache({
    estimateBytes: (surface) => surface.bytes,
    maxBytes: 8,
    maxEntries: 3,
  });
  first.set("a", { id: "a", bytes: 4 });
  first.set("b", { id: "b", bytes: 4 });
  assert.equal(first.get("a")?.id, "a", "get must refresh LRU order");
  first.set("c", { id: "c", bytes: 4 });

  assert.equal(first.get("b"), undefined);
  assert.equal(first.get("a")?.id, "a");
  assert.equal(first.get("c")?.id, "c");
  assert.equal(first.size, 2);
  assert.equal(first.estimatedBytes, 8);

  first.set("oversized", { id: "oversized", bytes: 9 });
  assert.equal(first.get("oversized"), undefined);
  assert.equal(first.estimatedBytes, 8);

  const second = new core.DocxThumbnailSurfaceCache({
    estimateBytes: (surface) => surface.bytes,
    maxBytes: 20,
    maxEntries: 5,
  });
  second.set("a", { id: "second-a", bytes: 10 });
  assert.equal(second.get("a")?.id, "second-a");
  assert.equal(first.get("a")?.id, "a", "cache instances must not share entries");

  assert.equal(first.delete("a"), true);
  assert.equal(first.estimatedBytes, 4);
  first.clear();
  assert.equal(first.size, 0);
  assert.equal(first.estimatedBytes, 0);

  const compatibleCountOnlyCache = new core.DocxThumbnailSurfaceCache(1);
  compatibleCountOnlyCache.set("first", 1);
  compatibleCountOnlyCache.set("second", 2);
  assert.equal(compatibleCountOnlyCache.get("first"), undefined);
  assert.equal(compatibleCountOnlyCache.get("second"), 2);
});
