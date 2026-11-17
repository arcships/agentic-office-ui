#!/usr/bin/env node
// DOCX edit/history benchmark for issue #37.
//
// Measures the cost of one edit commit on a long document:
//   A. keystroke commits (updateParagraphText) — how much of the document is
//      rebuilt per keystroke and whether image bytes are copied;
//   B. paragraph style toggles (applyRunStyle) — same accounting;
//   C. history pushes (createEditorStateV2 + applyEditorTransactionV2) — how
//      many times the retained snapshots are JSON-serialised per push;
//   D. layout passes — whether the layout copies image bytes.
//
// It imports the DOCX core *source* modules (Node type stripping) so it can run
// in any checkout, including one without a build or node_modules:
//
//   node scripts/bench/docx-edit-history.mjs
//   node scripts/bench/docx-edit-history.mjs --json
//
// The numbers are wall-clock timings from this machine; treat them as a
// before/after ratio, not as a portable budget.

import { register } from "node:module";
import path from "node:path";
import process from "node:process";
import { pathToFileURL } from "node:url";

const root = path.resolve(import.meta.dirname, "../..");

register(
  `data:text/javascript,${encodeURIComponent(`
    export async function resolve(specifier, context, nextResolve) {
      try {
        return await nextResolve(specifier, context);
      } catch (error) {
        if (specifier.startsWith(".") && !specifier.endsWith(".ts")) {
          return nextResolve(specifier + ".ts", context);
        }
        throw error;
      }
    }
  `)}`,
  import.meta.url,
);

const importSource = (relative) =>
  import(pathToFileURL(path.join(root, "packages/docx-core/src", relative)).href);

const { updateParagraphText } = await importSource("editor/paragraph-ops.ts");
const { applyRunStyle } = await importSource("editor/run-style-ops.ts");
const { createEditorStateV2, applyEditorTransactionV2 } = await importSource(
  "editor/helpers/state.ts",
);
const { layoutDocument } = await importSource("layout/layout-engine.ts");

const PAGE_COUNT = 200;
const IMAGE_COUNT = 4;
const IMAGE_BYTES = 5 * 1024 * 1024;
const KEYSTROKES = 100;
const STYLE_TOGGLES = 100;
const HISTORY_PUSHES = 60;
const HISTORY_ENTRY_TEXT = 300;
const LAYOUT_PASSES = 20;

function createModel({ pageCount = PAGE_COUNT, textLength = 180, imageCount = IMAGE_COUNT } = {}) {
  const nodes = [];
  for (let index = 0; index < pageCount; index += 1) {
    const children = [
      {
        type: "text",
        text: `Paragraph ${index}: ${"lorem ipsum ".repeat(Math.ceil(textLength / 12)).slice(0, textLength)}`,
      },
    ];
    if (index % Math.ceil(pageCount / Math.max(1, imageCount)) === 0 && imageCount > 0) {
      const imageIndex = Math.min(imageCount - 1, Math.floor(index / Math.ceil(pageCount / imageCount)));
      children.push({
        type: "image",
        src: `data:image/png;base64,image-${imageIndex}`,
        partName: `word/media/image${imageIndex}.png`,
        contentType: "image/png",
        widthPx: 640,
        heightPx: 480,
        data: new Uint8Array(IMAGE_BYTES),
      });
    }
    nodes.push({ type: "paragraph", children });
  }

  return {
    nodes,
    metadata: {
      sourceParts: 1,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
    },
  };
}

function imageRunsOf(model) {
  const runs = [];
  for (const node of model.nodes) {
    if (node.type !== "paragraph") continue;
    for (const child of node.children) {
      if (child.type === "image" && child.data) runs.push(child.data);
    }
  }
  return runs;
}

function rebuiltNodeIndexes(previous, next) {
  const limit = Math.max(previous.nodes.length, next.nodes.length);
  const rebuilt = [];
  for (let index = 0; index < limit; index += 1) {
    if (previous.nodes[index] !== next.nodes[index]) rebuilt.push(index);
  }
  return rebuilt;
}

function copiedImageBytes(previous, next) {
  const previousBytes = imageRunsOf(previous);
  const nextBytes = imageRunsOf(next);
  let copied = 0;
  for (let index = 0; index < nextBytes.length; index += 1) {
    if (nextBytes[index] !== previousBytes[index]) copied += nextBytes[index].byteLength;
  }
  return copied;
}

function median(values) {
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)] ?? 0;
}

function measure(iterations, run) {
  const samples = [];
  for (let index = 0; index < iterations; index += 1) {
    const startedAt = performance.now();
    run(index);
    samples.push(performance.now() - startedAt);
  }
  const total = samples.reduce((sum, value) => sum + value, 0);
  return {
    totalMs: Number(total.toFixed(2)),
    meanMs: Number((total / iterations).toFixed(3)),
    medianMs: Number(median(samples).toFixed(3)),
    maxMs: Number(Math.max(...samples).toFixed(3)),
  };
}

function benchmarkKeystrokes() {
  let model = createModel();
  let rebuilt = 0;
  let copiedBytes = 0;
  const timing = measure(KEYSTROKES, (index) => {
    const next = updateParagraphText(model, 0, `typed ${index}`);
    rebuilt += rebuiltNodeIndexes(model, next).length;
    copiedBytes += copiedImageBytes(model, next);
    model = next;
  });

  return {
    scenario: "keystroke commit (updateParagraphText)",
    docPages: PAGE_COUNT,
    imageBytes: IMAGE_COUNT * IMAGE_BYTES,
    iterations: KEYSTROKES,
    ...timing,
    rebuiltNodesPerOp: Number((rebuilt / KEYSTROKES).toFixed(2)),
    copiedImageBytesPerOp: Math.round(copiedBytes / KEYSTROKES),
  };
}

function benchmarkStyleToggles() {
  let model = createModel();
  let rebuilt = 0;
  let copiedBytes = 0;
  const timing = measure(STYLE_TOGGLES, (index) => {
    const next = applyRunStyle(model, 0, 0, { bold: index % 2 === 0 });
    rebuilt += rebuiltNodeIndexes(model, next).length;
    copiedBytes += copiedImageBytes(model, next);
    model = next;
  });

  return {
    scenario: "style toggle (applyRunStyle)",
    iterations: STYLE_TOGGLES,
    ...timing,
    rebuiltNodesPerOp: Number((rebuilt / STYLE_TOGGLES).toFixed(2)),
    copiedImageBytesPerOp: Math.round(copiedBytes / STYLE_TOGGLES),
  };
}

function benchmarkHistory(mode) {
  const selection = { kind: "paragraph", nodeIndex: 0 };
  let state = createEditorStateV2({
    model: createModel({ imageCount: 0, textLength: HISTORY_ENTRY_TEXT }),
    selection,
    historyMaxEntries: 100,
    historyMaxBytes: 32 * 1024 * 1024,
  });

  const originalStringify = JSON.stringify;
  let serializations = 0;
  JSON.stringify = function countedStringify(...args) {
    serializations += 1;
    return originalStringify.apply(this, args);
  };

  let model = state.model;
  let timing;
  try {
    timing = measure(HISTORY_PUSHES, (index) => {
      model = updateParagraphText(
        model,
        0,
        `${mode}:${index}:${String(index).repeat(HISTORY_ENTRY_TEXT).slice(0, HISTORY_ENTRY_TEXT)}`,
      );
      state = applyEditorTransactionV2(state, { model, selection });
    });
  } finally {
    JSON.stringify = originalStringify;
  }

  return {
    scenario: "history push (createEditorStateV2 + applyEditorTransactionV2)",
    iterations: HISTORY_PUSHES,
    retainedEntries: state.history.past.length,
    ...timing,
    serializations,
    serializationsPerPush: Number((serializations / HISTORY_PUSHES).toFixed(2)),
    retainedBytes: state.history.pastBytes,
  };
}

function benchmarkLayout() {
  const model = createModel({ imageCount: 4 });
  const originalBytes = imageRunsOf(model);
  let sharedBytes = 0;

  const timing = measure(LAYOUT_PASSES, () => {
    const pages = layoutDocument(model, { pageWidth: 816, pageHeight: 1056 });
    for (const page of pages) {
      for (const block of page.blocks) {
        for (const run of block.runs ?? []) {
          if (run.kind !== "image" || !run.data) continue;
          if (originalBytes.includes(run.data)) sharedBytes += run.data.byteLength;
        }
      }
    }
  });

  return {
    scenario: "layout pass (layoutDocument)",
    iterations: LAYOUT_PASSES,
    ...timing,
    sharedImageBytesPerPass: Math.round(sharedBytes / LAYOUT_PASSES),
  };
}

const results = [
  benchmarkKeystrokes(),
  benchmarkStyleToggles(),
  benchmarkHistory("history"),
  benchmarkLayout(),
];

if (process.argv.includes("--json")) {
  console.log(JSON.stringify({ results }, null, 2));
} else {
  console.log(`DOCX edit/history benchmark — ${PAGE_COUNT}-page document, ${IMAGE_COUNT * IMAGE_BYTES / (1024 * 1024)} MB of images`);
  console.log("");
  const columns = [
    "scenario",
    "meanMs",
    "medianMs",
    "maxMs",
    "rebuiltNodesPerOp",
    "copiedImageBytesPerOp",
    "serializationsPerPush",
    "sharedImageBytesPerPass",
  ];
  console.log(columns.join("\t"));
  for (const result of results) {
    console.log(columns.map((column) => result[column] ?? "-").join("\t"));
  }
}
