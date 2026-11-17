import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);
const runtimeEntry = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core/runtime")).href
);
const vueDocx = await import(
  new URL("../../packages/vue-docx/dist/index.js", import.meta.url).href
);
const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

function arrayBufferFromFile(name) {
  const bytes = readFileSync(new URL(name, samples));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function withWordMainContentType(input, contentType) {
  const archive = unzipSync(new Uint8Array(input));
  const contentTypes = strFromU8(archive["[Content_Types].xml"]);
  archive["[Content_Types].xml"] = strToU8(contentTypes.replace(
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    contentType,
  ));
  const bytes = zipSync(archive);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function collectText(value, result = []) {
  if (!value || typeof value !== "object") return result;
  if (value.type === "text" && typeof value.text === "string") {
    result.push(value.text);
  }
  for (const child of Array.isArray(value) ? value : Object.values(value)) {
    collectText(child, result);
  }
  return result;
}

function workerResult(request, marker) {
  return {
    data: {
      id: request.id,
      type: "success",
      package: { parts: new Map(), binaryAssets: new Map() },
      model: {
        nodes: [{ id: marker, type: "paragraph", children: [] }],
        metadata: {
          sourceParts: 0,
          warnings: [],
          headerSections: [],
          footerSections: [],
          paragraphStyles: [],
        },
      },
      timings: { totalMs: 1, parseMs: 0.5, buildModelMs: 0.5 },
    },
  };
}

test("DOCX surface search returns exact UTF-16 ranges for body and direct table paragraphs", () => {
  const paragraph = (text) => ({
    type: "paragraph",
    children: [{ type: "text", text }],
  });
  const model = {
    nodes: [
      paragraph("Alpha alpha ALPHA"),
      {
        type: "table",
        rows: [{
          type: "table-row",
          cells: [{
            type: "table-cell",
            nodes: [
              paragraph("literal a.b and A.B"),
              {
                type: "table",
                rows: [{
                  type: "table-row",
                  cells: [{ type: "table-cell", nodes: [paragraph("nested a.b")] }],
                }],
              },
            ],
          }],
        }],
      },
    ],
    metadata: {
      sourceParts: 0,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
    },
  };

  const bodyMatches = vueDocx.findDocxSearchMatches(model, "alpha");
  assert.deepEqual(
    bodyMatches.map((match) => [match.range.start.offset, match.range.end.offset]),
    [[0, 5], [6, 11], [12, 17]],
  );
  assert.deepEqual(bodyMatches[0].range.start.location, { kind: "paragraph", nodeIndex: 0 });

  const tableMatches = vueDocx.findDocxSearchMatches(model, "a.b");
  assert.equal(tableMatches.length, 2);
  assert.deepEqual(tableMatches[0].range.start.location, {
    kind: "table-cell",
    tableIndex: 1,
    rowIndex: 0,
    cellIndex: 0,
    paragraphIndex: 0,
  });
  assert.equal(tableMatches.every((match) => match.nodeIndex === 1), true);
});

class ControlledDocxWorker {
  constructor({ delay = 0, marker = "worker", respond = true } = {}) {
    this.delay = delay;
    this.marker = marker;
    this.respond = respond;
    this.listeners = new Map();
    this.messages = [];
    this.terminated = false;
    this.timer = null;
  }

  addEventListener(type, listener) {
    this.listeners.set(type, [...(this.listeners.get(type) || []), listener]);
  }

  removeEventListener(type, listener) {
    this.listeners.set(
      type,
      (this.listeners.get(type) || []).filter((item) => item !== listener),
    );
  }

  postMessage(request) {
    this.messages.push(request);
    if (!this.respond) return;
    this.timer = setTimeout(() => {
      if (this.terminated) return;
      for (const listener of this.listeners.get("message") || []) {
        listener(workerResult(request, this.marker));
      }
    }, this.delay);
  }

  terminate() {
    this.terminated = true;
    if (this.timer) clearTimeout(this.timer);
  }
}

async function withWorkerGlobal(run) {
  const original = globalThis.Worker;
  globalThis.Worker = ControlledDocxWorker;
  try {
    return await run();
  } finally {
    if (original === undefined) delete globalThis.Worker;
    else globalThis.Worker = original;
  }
}

test("DOCX parses a normal fixture and preserves content through roundtrip", async () => {
  const input = arrayBufferFromFile("invoice-table.docx");
  const first = await core.wasmBuildDocModelFromBytes(input);
  const firstText = collectText(first.model).join(" ");
  assert.match(firstText, /INVOICE INV-2026-0705/);
  assert.ok(first.model.nodes.length > 0);

  const exported = await core.wasmSerializeDocx(first.model, first.package);
  assert.ok(exported.byteLength > 0);
  assert.deepEqual([...new Uint8Array(exported).slice(0, 2)], [0x50, 0x4b]);

  const second = await core.wasmBuildDocModelFromBytes(exported);
  const secondText = collectText(second.model).join(" ");
  assert.match(secondText, /INVOICE INV-2026-0705/);
  assert.equal(second.model.nodes.length, first.model.nodes.length);
});

test("DOCX importer accepts macro-enabled documents and Word templates as read-only inputs", async () => {
  const source = arrayBufferFromFile("invoice-table.docx");
  for (const [format, contentType] of [
    ["docm", "application/vnd.ms-word.document.macroEnabled.main+xml"],
    ["dotx", "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"],
    ["dotm", "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"],
  ]) {
    const result = await core.importDocxBuffer(withWordMainContentType(source, contentType), {
      useWorker: false,
    });
    assert.ok(result.model.nodes.length > 0, `${format} should produce a document model`);
  }
});

test("DOCX rejects empty and corrupted input with a stable parse error", async () => {
  for (const input of [new ArrayBuffer(0), arrayBufferFromFile("corrupted.docx")]) {
    await assert.rejects(
      core.importDocxBuffer(input, { useWorker: false }),
      (error) => error?.name === "DocxImportError" && error?.code === "PARSE_FAILED",
    );
  }
});

test("DOCX runtime entry exposes its public import error constructor", () => {
  assert.equal(runtimeEntry.DocxImportError, core.DocxImportError);
  const error = new runtimeEntry.DocxImportError("PARSE_FAILED", "broken document");
  assert.equal(error.name, "DocxImportError");
  assert.equal(error.code, "PARSE_FAILED");
});

test("DOCX page margin styles include CSS pixel units", () => {
  assert.deepEqual(
    core.pageMarginPaddingStyle({ top: 56, right: 48, bottom: 40, left: 32 }),
    {
      paddingTop: "56px",
      paddingRight: "48px",
      paddingBottom: "40px",
      paddingLeft: "32px",
    },
  );
});

test("DOCX tracked revision styles stay valid and distinguishable", () => {
  const baseStyle = core.runStyleToCss({}, "light");
  assert.equal(baseStyle.textDecoration, "none");

  const insertion = core.trackedInlineStyle(
    baseStyle,
    { id: "ins-1", kind: "insertion" },
    "light",
  );
  const deletion = core.trackedDeletedStyle("light");
  const insertionDark = core.trackedInlineStyle(
    baseStyle,
    { id: "ins-1", kind: "insertion" },
    "dark",
  );
  const deletionDark = core.trackedDeletedStyle("dark");

  assert.equal(insertion.textDecoration, "underline");
  assert.equal(deletion.textDecoration, "line-through");
  assert.notEqual(insertion.color, deletion.color);
  assert.notEqual(insertionDark.color, deletionDark.color);
  assert.equal(insertion.textDecorationColor, insertion.color);
  assert.equal(deletion.textDecorationColor, deletion.color);

  // "none" is exclusive in CSS; it must never be combined with a decoration.
  assert.equal(core.mergeTextDecorations("none", "underline"), "underline");
  assert.equal(core.mergeTextDecorations("none", "line-through"), "line-through");
  assert.equal(
    core.mergeTextDecorations("underline", "line-through"),
    "underline line-through",
  );
  assert.equal(core.mergeTextDecorations("none", "none"), "none");

  const underlinedDeletion = core.trackedDeletedStyle("light", { underline: true });
  assert.match(underlinedDeletion.textDecoration, /\bline-through\b/);
  assert.match(underlinedDeletion.textDecoration, /\bunderline\b/);
  assert.doesNotMatch(underlinedDeletion.textDecoration, /\bnone\b/);

  // The gutter accent helper should follow the same insertion/deletion split.
  assert.notEqual(
    core.trackedChangeAccentColor("insertion", "light"),
    core.trackedChangeAccentColor("deletion", "light"),
  );
  assert.notEqual(
    core.trackedChangeAccentColor("insertion", "dark"),
    core.trackedChangeAccentColor("deletion", "dark"),
  );
});

test("DOCX abort terminates the active Worker and rejects as ABORTED", async () => {
  await withWorkerGlobal(async () => {
    const worker = new ControlledDocxWorker({ respond: false });
    const runtime = core.createDocxRuntime({
      createWorker: () => worker,
      wasmUrl: "https://unit.invalid/docx.wasm",
    });
    const controller = new AbortController();
    const pending = runtime.importDocxBuffer(new ArrayBuffer(8), {
      signal: controller.signal,
    });
    controller.abort();
    await assert.rejects(
      pending,
      (error) => error?.name === "AbortError" && error?.code === "ABORTED",
    );
    assert.equal(worker.terminated, true);
    runtime.dispose();
  });
});

test("DOCX loader cancels the older request and accepts only the latest result", async () => {
  await withWorkerGlobal(async () => {
    const workers = [];
    const runtime = core.createDocxRuntime({
      createWorker: () => {
        const index = workers.length;
        const worker = new ControlledDocxWorker({
          delay: index === 0 ? 40 : 0,
          marker: index === 0 ? "old" : "latest",
        });
        workers.push(worker);
        return worker;
      },
      wasmUrl: "https://unit.invalid/docx.wasm",
    });
    const loader = runtime.createLoader();
    const oldRequest = loader.load(new ArrayBuffer(8));
    const latestRequest = loader.load(new ArrayBuffer(8));

    await assert.rejects(
      oldRequest,
      (error) => error?.name === "AbortError" && error?.code === "ABORTED",
    );
    const latest = await latestRequest;
    assert.equal(latest.source, "worker");
    assert.equal(latest.model.nodes[0].id, "latest");
    assert.equal(workers[0].terminated, true);

    loader.dispose();
    runtime.dispose();
  });
});

test("DOCX runtime loads URL sources through shared policy and one task identity", async () => {
  await withWorkerGlobal(async () => {
    const diagnostics = [];
    const requests = [];
    const runtime = core.createDocxRuntime({
      createWorker: () => new ControlledDocxWorker({ marker: "url-source" }),
      wasmUrl: "https://assets.example/docx.wasm",
      urlPolicy: {
        enabled: true,
        baseUrl: "https://app.example/viewer/",
        allowRelativeUrl: true,
        allowedProtocols: ["https:"],
        allowedOrigins: ["https://app.example"],
        fetch: async (url, init) => {
          requests.push({ url, init });
          const bytes = new Uint8Array(arrayBufferFromFile("invoice-table.docx"));
          return {
            ok: true,
            status: 200,
            url,
            headers: { get: (name) => name.toLowerCase() === "content-length" ? String(bytes.byteLength) : null },
            arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
          };
        },
      },
      limits: { maxInputBytes: 1024 * 1024 },
      onDiagnostic: (event) => diagnostics.push(event),
    });
    const result = await runtime.loadSource({ kind: "url", url: "invoice-table.docx?token=secret" });
    assert.equal(result.model.nodes[0].id, "url-source");
    assert.equal(requests.length, 1);
    assert.equal(requests[0].init.redirect, "follow");
    assert.equal(requests[0].init.credentials, "omit");
    const taskIds = new Set(diagnostics.map((event) => event.requestId));
    assert.equal(taskIds.size, 1);
    assert.match([...taskIds][0], new RegExp(`^${runtime.id}:1$`));
    assert.equal(JSON.stringify(diagnostics).includes("secret"), false);
    runtime.dispose();
  });
});

test("DOCX runtime rejects a dangerous URL before fetch", async () => {
  let requestCount = 0;
  const runtime = core.createDocxRuntime({
    urlPolicy: {
      enabled: true,
      baseUrl: "https://app.example/",
      allowRelativeUrl: true,
      allowedProtocols: ["https:"],
      allowedOrigins: ["https://app.example"],
      fetch: async () => { requestCount += 1; throw new Error("must not fetch"); },
    },
  });
  await assert.rejects(
    runtime.loadSource({ kind: "url", url: "javascript:alert(1)" }),
    (error) => error?.name === "DocxImportError" && error?.code === "SOURCE_NOT_ALLOWED",
  );
  assert.equal(requestCount, 0);
  runtime.dispose();
});

test("DOCX runtime snapshots limits instead of reading later caller mutations", async () => {
  await withWorkerGlobal(async () => {
    const limits = { maxInputBytes: 8, maxArchiveEntries: 7, maxDocxPages: 3 };
    const worker = new ControlledDocxWorker({ marker: "snapshotted-config" });
    const runtime = core.createDocxRuntime({
      limits,
      createWorker: () => worker,
      wasmUrl: "https://assets.example/docx.wasm",
    });
    limits.maxInputBytes = 1;
    limits.maxArchiveEntries = 1;
    limits.maxDocxPages = 1;
    const result = await runtime.importDocxBuffer(new ArrayBuffer(4));
    assert.equal(result.model.nodes[0].id, "snapshotted-config");
    assert.equal(Object.isFrozen(runtime.limits), true);
    assert.equal(runtime.limits.maxArchiveEntries, 7);
    assert.equal(runtime.limits.maxDocxPages, 3);
    assert.equal(worker.messages[0].limits.maxArchiveEntries, 7);
    runtime.dispose();
  });
});

test("DOCX Worker timeout is structured and terminates the blocked Worker", async () => {
  await withWorkerGlobal(async () => {
    const worker = new ControlledDocxWorker({ respond: false });
    const runtime = core.createDocxRuntime({
      createWorker: () => worker,
      wasmUrl: "https://assets.example/docx.wasm",
      limits: { maxInputBytes: 8, maxParseMs: 5 },
    });
    await assert.rejects(
      runtime.importDocxBuffer(new ArrayBuffer(4)),
      (error) => error?.code === "TIMEOUT" &&
        error?.phase === "parse" &&
        error?.limit === "maxParseMs" &&
        error?.actual > error?.allowed,
    );
    assert.equal(worker.terminated, true);
    runtime.dispose();
  });
});

// ---------------------------------------------------------------------------
// Comment collection (issue #30)
// ---------------------------------------------------------------------------

function commentParagraph(text, sourceXml, extra = {}) {
  return {
    type: "paragraph",
    children: [{ type: "text", text }],
    sourceXml,
    ...extra,
  };
}

function commentRangeXml(text, commentId) {
  return `<w:p><w:commentRangeStart w:id="${commentId}"/><w:r><w:t>${text}</w:t></w:r><w:commentRangeEnd w:id="${commentId}"/><w:r><w:commentReference w:id="${commentId}"/></w:r></w:p>`;
}

function commentModel({
  nodes,
  comments,
  headerSections = [],
  footerSections = [],
  sections = undefined,
}) {
  return {
    nodes,
    metadata: {
      sourceParts: 1,
      warnings: [],
      headerSections,
      footerSections,
      paragraphStyles: [],
      sections,
      comments,
    },
  };
}

function commentByIdMap(comments) {
  return new Map(comments.map((comment) => [comment.commentId, comment]));
}

test("DOCX comment collection includes threaded replies and keeps one anchor per thread", () => {
  const model = commentModel({
    nodes: [commentParagraph("Quarterly revenue", commentRangeXml("Quarterly revenue", 0))],
    comments: [
      { id: 0, author: "Priya", text: "Please confirm the revenue figure." },
      { id: 1, author: "Sam", text: "Confirmed against the ledger.", parentId: 0 },
      { id: 2, author: "Priya", text: "Thanks!", parentId: 0 },
    ],
  });

  const comments = core.collectCommentsFromModel(model);
  assert.deepEqual(comments.map((comment) => comment.commentId), [0, 1, 2]);
  assert.deepEqual(comments.map((comment) => comment.parentId), [undefined, 0, 0]);
  for (const comment of comments) {
    assert.deepEqual(comment.location, { kind: "paragraph", nodeIndex: 0 });
    assert.equal(comment.anchorText, "Quarterly revenue");
  }
});

test("DOCX comment collection includes header and footer comments with a region tag", () => {
  const headerParagraph = commentParagraph(
    "Acme Corp",
    commentRangeXml("Acme Corp", 7),
  );
  const footerParagraph = commentParagraph(
    "Confidential",
    commentRangeXml("Confidential", 8),
  );
  const model = commentModel({
    nodes: [commentParagraph("Body", undefined)],
    comments: [
      { id: 7, author: "Priya", text: "Header note" },
      { id: 8, author: "Sam", text: "Footer note" },
    ],
    headerSections: [{ partName: "word/header1.xml", nodes: [headerParagraph] }],
    footerSections: [{ partName: "word/footer1.xml", nodes: [footerParagraph] }],
  });

  const comments = core.collectCommentsFromModel(model);
  assert.equal(comments.length, 2);
  const header = comments.find((comment) => comment.commentId === 7);
  const footer = comments.find((comment) => comment.commentId === 8);
  assert.equal(header?.region, "header");
  assert.equal(header?.partName, "word/header1.xml");
  assert.equal(header?.anchorText, "Acme Corp");
  assert.equal(footer?.region, "footer");
  assert.equal(footer?.partName, "word/footer1.xml");
  assert.equal(footer?.anchorText, "Confidential");
});

test("DOCX comment collection survives edits that clear paragraph sourceXml", () => {
  const model = commentModel({
    nodes: [commentParagraph("Quarterly revenue", commentRangeXml("Quarterly revenue", 0))],
    comments: [{ id: 0, author: "Priya", text: "Please confirm." }],
  });
  assert.equal(core.collectCommentsFromModel(model).length, 1);

  const edited = core.updateParagraphText(model, 0, "Revised revenue summary");
  assert.equal(edited.nodes[0].sourceXml, undefined);

  const comments = core.collectCommentsFromModel(edited);
  assert.equal(comments.length, 1);
  assert.deepEqual(comments[0].location, { kind: "paragraph", nodeIndex: 0 });
  assert.equal(comments[0].anchorText, "Quarterly revenue");
});

test("DOCX table comments use non-negative paragraph indexes and stay on the anchored row's page", () => {
  const nestedTable = {
    type: "table",
    rows: [
      {
        type: "table-row",
        cells: [
          {
            type: "table-cell",
            nodes: [
              commentParagraph(
                "Nested comment",
                commentRangeXml("Nested comment", 4),
              ),
            ],
          },
        ],
      },
    ],
  };
  const model = commentModel({
    nodes: [
      {
        type: "table",
        rows: [
          { type: "table-row", cells: [{ type: "table-cell", nodes: [commentParagraph("Top", undefined)] }] },
          {
            type: "table-row",
            cells: [
              {
                type: "table-cell",
                nodes: [commentParagraph("Direct", undefined), nestedTable],
              },
            ],
          },
        ],
      },
    ],
    comments: [{ id: 4, author: "Priya", text: "Nested table note" }],
  });

  const comments = core.collectCommentsFromModel(model);
  assert.equal(comments.length, 1);
  assert.equal(comments[0].location.kind, "table-cell");
  assert.equal(comments[0].location.rowIndex, 1);
  assert.ok(comments[0].location.paragraphIndex >= 0);
  assert.equal(comments[0].anchorText, "Nested comment");

  const resolved = core.getParagraphAtLocation(model, comments[0].location);
  assert.equal(core.paragraphText(resolved.paragraph), "Nested comment");

  const firstPage = [{ nodeIndex: 0, tableRowRange: { startRowIndex: 0, endRowIndex: 1 } }];
  const secondPage = [{ nodeIndex: 0, tableRowRange: { startRowIndex: 1, endRowIndex: 2 } }];
  assert.equal(core.trackedChangeBelongsToPageSegments(comments[0].location, firstPage), false);
  assert.equal(core.trackedChangeBelongsToPageSegments(comments[0].location, secondPage), true);
});

test("DOCX resolved comment threads are detected for replies", () => {
  const model = commentModel({
    nodes: [commentParagraph("Quarterly revenue", commentRangeXml("Quarterly revenue", 0))],
    comments: [
      { id: 0, author: "Priya", text: "Please confirm.", resolved: true },
      { id: 1, author: "Sam", text: "Confirmed.", parentId: 0 },
    ],
  });
  const comments = core.collectCommentsFromModel(model);
  const byId = commentByIdMap(comments);
  assert.equal(core.commentThreadIsResolved(byId.get(0), byId), true);
  assert.equal(core.commentThreadIsResolved(byId.get(1), byId), true);
});
