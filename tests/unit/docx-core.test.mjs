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
const UTIF = requireFromDocxCore("utif");
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

function tiffDataUri(width, height, [red, green, blue]) {
  const rgba = new Uint8Array(width * height * 4);
  for (let index = 0; index < rgba.length; index += 4) {
    rgba[index] = red;
    rgba[index + 1] = green;
    rgba[index + 2] = blue;
    rgba[index + 3] = 255;
  }
  const tiff = new Uint8Array(UTIF.encodeImage(rgba, width, height));
  return `data:image/tiff;base64,${Buffer.from(tiff).toString("base64")}`;
}

function docxWithAnchoredImage(anchorAttributes) {
  const archive = unzipSync(new Uint8Array(arrayBufferFromFile("report-with-image.docx")));
  let documentXml = strFromU8(archive["word/document.xml"]);
  const inline = documentXml.match(
    /<w:drawing><wp:inline[\s\S]*?<\/wp:inline><\/w:drawing>/
  )[0];
  const extent = inline.match(/<wp:extent cx="(\d+)" cy="(\d+)"/);
  const anchor = `<w:drawing><wp:anchor ${anchorAttributes}>` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>720000</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>360000</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${extent[1]}" cy="${extent[2]}"/>` +
    `<wp:effectExtent l="19050" t="19050" r="19050" b="19050"/>` +
    `<wp:wrapSquare wrapText="bothSides"/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr><pic:cNvPr id="0" name="invoice.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rId11"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${extent[1]}" cy="${extent[2]}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing>`;
  documentXml = documentXml.replace(inline, anchor);
  archive["word/document.xml"] = strToU8(documentXml);
  const bytes = zipSync(archive);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function findFirstImage(model) {
  let found;
  const visit = (nodes) => {
    for (const node of nodes) {
      if (node.type === "paragraph") {
        for (const child of node.children) {
          if (child.type === "image" && !found) found = child;
        }
      } else if (node.type === "table") {
        for (const row of node.rows) {
          for (const cell of row.cells) visit(cell.nodes);
        }
      }
    }
  };
  visit(model.nodes);
  return found;
}

function withWordMainContentType(input, contentType, { vbaProject } = {}) {
  const archive = unzipSync(new Uint8Array(input));
  const contentTypes = strFromU8(archive["[Content_Types].xml"]);
  archive["[Content_Types].xml"] = strToU8(contentTypes.replace(
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    contentType,
  ));
  if (vbaProject !== undefined) {
    archive["word/vbaProject.bin"] = vbaProject;
  }
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

/** Text, tables and image geometry of a model, independent of run boundaries. */
function modelProjection(model) {
  const projection = { paragraphs: [], tables: [], images: [] };
  const visit = (nodes) => {
    for (const node of nodes) {
      if (node.type === "paragraph") {
        projection.paragraphs.push(collectText(node).join(""));
        for (const child of node.children) {
          if (child.type === "image") {
            projection.images.push({
              widthPx: child.widthPx ?? null,
              heightPx: child.heightPx ?? null,
              wrapType: child.floating?.wrapType ?? null,
              alt: child.alt ?? null,
            });
          }
        }
      } else if (node.type === "table") {
        projection.tables.push(
          node.rows.map((row) =>
            row.cells.map((cell) => collectText(cell).join("")),
          ),
        );
        visit(node.rows.flatMap((row) => row.cells.flatMap((cell) => cell.nodes)));
      }
    }
  };
  visit(model.nodes);
  return projection;
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

test("DOCX surface search returns exact UTF-16 ranges for body, direct and nested table paragraphs", () => {
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
  assert.equal(tableMatches.length, 3);
  assert.deepEqual(tableMatches[0].range.start.location, {
    kind: "table-cell",
    tableIndex: 1,
    rowIndex: 0,
    cellIndex: 0,
    paragraphIndex: 0,
  });
  assert.deepEqual(tableMatches[2].range.start.location, {
    kind: "nested-table-cell",
    tableIndex: 1,
    path: [{ rowIndex: 0, cellIndex: 0, contentIndex: 1 }],
    rowIndex: 0,
    cellIndex: 0,
    paragraphIndex: 0,
    nestingPath: "0-0-1",
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

  // A ZIP signature proves nothing about the body. Open the exported package
  // and compare the text, tables and images before and after re-import.
  const archive = unzipSync(new Uint8Array(exported));
  assert.ok(archive["word/document.xml"], "exported package must contain word/document.xml");
  assert.match(strFromU8(archive["word/document.xml"]), /<w:tbl>/);
  assert.ok(
    archive["[Content_Types].xml"],
    "exported package must contain [Content_Types].xml",
  );

  const second = await core.wasmBuildDocModelFromBytes(exported);
  const secondText = collectText(second.model).join(" ");
  assert.match(secondText, /INVOICE INV-2026-0705/);
  assert.equal(second.model.nodes.length, first.model.nodes.length);
  assert.deepEqual(modelProjection(second.model), modelProjection(first.model));
});

test("DOCX export round trip keeps images and their geometry", async () => {
  const first = await core.wasmBuildDocModelFromBytes(
    arrayBufferFromFile("report-with-image.docx"),
  );
  const exported = await core.wasmSerializeDocx(first.model, first.package);
  const archive = unzipSync(new Uint8Array(exported));
  assert.ok(
    Object.keys(archive).some((name) => name.startsWith("word/media/")),
    "exported package must keep the image part",
  );
  const second = await core.wasmBuildDocModelFromBytes(exported);
  assert.deepEqual(modelProjection(second.model), modelProjection(first.model));
  assert.equal(
    modelProjection(second.model).images.length,
    1,
    "the round trip must not lose or duplicate the image",
  );
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

test("DOCX main-thread parsing enforces maxParseMs instead of returning a late result", async () => {
  await assert.rejects(
    core.importDocxBuffer(arrayBufferFromFile("invoice-table.docx"), {
      useWorker: false,
      limits: { maxParseMs: 1 },
    }),
    (error) => error?.name === "DocxImportError" &&
      error?.code === "TIMEOUT" &&
      error?.phase === "parse" &&
      error?.limit === "maxParseMs" &&
      error?.allowed === 1 &&
      error?.actual > error?.allowed,
  );
});

test("DOCX macro and template exports preserve content types and macro parts", async () => {
  const source = arrayBufferFromFile("invoice-table.docx");
  const macroBytes = new Uint8Array([0x4d, 0x5a, 0x90, 0x00, 0x03, 0x00, 0x00, 0x00]);
  for (const [format, contentType, hasMacro] of [
    ["docm", "application/vnd.ms-word.document.macroEnabled.main+xml", true],
    ["dotx", "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", false],
    ["dotm", "application/vnd.ms-word.template.macroEnabledTemplate.main+xml", true],
  ]) {
    const input = withWordMainContentType(source, contentType, {
      vbaProject: hasMacro ? macroBytes : undefined,
    });
    const { model, package: pkg } = await core.wasmBuildDocModelFromBytes(input);
    const exported = await core.serializeDocx(model, core.wasmPackageToMaps(pkg));
    const archive = unzipSync(new Uint8Array(exported));
    assert.match(
      strFromU8(archive["[Content_Types].xml"]),
      new RegExp(contentType.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")),
      `${format} export must keep its container content type`,
    );
    if (hasMacro) {
      assert.deepEqual(
        [...archive["word/vbaProject.bin"]],
        [...macroBytes],
        `${format} export must keep the macro part`,
      );
    } else {
      assert.equal(archive["word/vbaProject.bin"], undefined);
    }
  }
});

test("DOCX edit and export roundtrip preserves styles, notes, comments and image bytes", async () => {
  for (const sample of ["review-comments.docx", "report-with-image.docx"]) {
    const first = await core.wasmBuildDocModelFromBytes(arrayBufferFromFile(sample));
    const firstPackage = core.wasmPackageToMaps(first.package);
    const images = [...firstPackage.binaryAssets].filter(([name]) => name.includes("/media/"));
    const styleJson = JSON.stringify(first.model.metadata.paragraphStyles);
    const commentCount = core.collectCommentsFromModel(first.model).length;
    const trackedCount = core.collectTrackedChangesFromModel(first.model).length;
    const footnoteCount = first.model.metadata.footnotes?.length ?? 0;
    const endnoteCount = first.model.metadata.endnotes?.length ?? 0;

    const edited = core.updateParagraphText(first.model, 0, "Roundtrip fidelity paragraph");
    const exported = await core.serializeDocx(edited, firstPackage);
    const second = await core.wasmBuildDocModelFromBytes(exported);
    const secondPackage = core.wasmPackageToMaps(second.package);

    assert.match(collectText(second.model).join(" "), /Roundtrip fidelity paragraph/, sample);
    assert.equal(JSON.stringify(second.model.metadata.paragraphStyles), styleJson, sample);
    assert.equal(core.collectCommentsFromModel(second.model).length, commentCount, sample);
    assert.equal(core.collectTrackedChangesFromModel(second.model).length, trackedCount, sample);
    assert.equal(second.model.metadata.footnotes?.length ?? 0, footnoteCount, sample);
    assert.equal(second.model.metadata.endnotes?.length ?? 0, endnoteCount, sample);
    for (const [name, bytes] of images) {
      const reimported = secondPackage.binaryAssets.get(name);
      assert.ok(reimported, `${sample}: ${name} must survive export`);
      assert.deepEqual([...reimported], [...bytes], `${sample}: ${name} bytes must survive export`);
    }
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

test("DOCX runtime shares one latest-wins slot across top-level imports", async () => {
  await withWorkerGlobal(async () => {
    const workers = [];
    const runtime = core.createDocxRuntime({
      createWorker: () => {
        const index = workers.length;
        const worker = new ControlledDocxWorker({
          delay: index === 0 ? 40 : 0,
          marker: ["old", "latest", "after"][index] ?? "extra",
        });
        workers.push(worker);
        return worker;
      },
      wasmUrl: "https://unit.invalid/docx.wasm",
    });

    const oldRequest = runtime.importDocxBuffer(new ArrayBuffer(8));
    const latestRequest = runtime.loadSource({ kind: "bytes", bytes: new ArrayBuffer(8) });
    await assert.rejects(
      oldRequest,
      (error) => error?.name === "AbortError" && error?.code === "ABORTED",
    );
    const latest = await latestRequest;
    assert.equal(latest.model.nodes[0].id, "latest");
    assert.equal(workers[0].terminated, true);

    const after = await runtime.importDocxBuffer(new ArrayBuffer(8));
    assert.equal(after.model.nodes[0].id, "after");
    assert.equal(workers.length, 3);
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

test("DOCX TIFF conversion is cached per source and target size", () => {
  const originalDecode = UTIF.decode;
  let decodeCount = 0;
  const countingDecode = (...args) => {
    decodeCount += 1;
    return originalDecode(...args);
  };
  // UTIF.decode carries its codec helpers as static properties and reads them
  // back through `UTIF.decode`; keep them on the counting wrapper.
  Object.assign(countingDecode, originalDecode);
  UTIF.decode = countingDecode;

  try {
    const firstSrc = tiffDataUri(16, 16, [220, 40, 40]);
    const secondSrc = tiffDataUri(16, 16, [40, 80, 220]);
    const first = { type: "image", src: firstSrc, contentType: "image/tiff", widthPx: 16, heightPx: 16 };
    const second = { type: "image", src: secondSrc, contentType: "image/tiff", widthPx: 16, heightPx: 16 };

    const firstPng = core.resolveRenderableImageSource(first);
    assert.match(firstPng, /^data:image\/png;base64,/);
    assert.equal(core.resolveRenderableImageSource({ ...first }), firstPng);
    assert.equal(core.resolveRenderableImageSource(first), firstPng);
    assert.equal(decodeCount, 1);

    const secondPng = core.resolveRenderableImageSource(second);
    assert.match(secondPng, /^data:image\/png;base64,/);
    assert.notEqual(secondPng, firstPng);
    assert.equal(core.resolveRenderableImageSource(second), secondPng);
    assert.equal(decodeCount, 2);

    // The target size is part of the cache key, so a different render size
    // produces its own conversion instead of reusing the wrong pixels.
    core.resolveRenderableImageSource({ ...first, widthPx: 8, heightPx: 8 });
    assert.equal(decodeCount, 3);
  } finally {
    UTIF.decode = originalDecode;
  }
});

test("DOCX floating anchor parses layoutInCell, allowOverlap, locked, simplePos and effectExtent", async () => {
  const input = docxWithAnchoredImage(
    'distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
    'relativeHeight="251658240" behindDoc="0" locked="1" layoutInCell="0" allowOverlap="1"'
  );
  const { model } = await core.importDocxBuffer(input, { useWorker: false });
  const image = findFirstImage(model);
  assert.ok(image?.floating);
  assert.equal(image.floating.layoutInCell, false);
  assert.equal(image.floating.allowOverlap, true);
  assert.equal(image.floating.lockedAnchor, true);
  assert.equal(image.floating.simplePos, false);
  assert.deepEqual(image.floating.effectExtent, {
    leftPx: 2,
    topPx: 2,
    rightPx: 2,
    bottomPx: 2,
  });
});

test("DOCX simplePos anchors resolve against the page origin", () => {
  const style = core.absoluteFloatingImageStyle(
    {
      type: "image",
      floating: {
        xPx: 120,
        yPx: 60,
        horizontalRelativeTo: "paragraph",
        verticalRelativeTo: "paragraph",
        simplePos: true,
      },
    },
    {
      pageOriginLeft: 10,
      pageOriginTop: 20,
      paragraphOriginLeft: 1000,
      paragraphOriginTop: 1000,
    }
  );
  assert.equal(style.left, 130);
  assert.equal(style.top, 80);
});

test("DOCX effectExtent expands the wrapped-image exclusion zone", () => {
  const geometry = core.resolveDualWrappedFloatingImageGeometry(
    {
      type: "image",
      widthPx: 100,
      heightPx: 80,
      floating: {
        xPx: 150,
        yPx: 50,
        wrapType: "square",
        horizontalRelativeTo: "column",
        verticalRelativeTo: "paragraph",
        effectExtent: { leftPx: 5, topPx: 6, rightPx: 7, bottomPx: 8 },
      },
    },
    400
  );
  assert.ok(geometry);
  assert.deepEqual(geometry.exclusion, {
    left: 145,
    right: 257,
    top: 44,
    bottom: 138,
  });
});

function furnitureParagraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function furnitureModel(headerNodes = [], footerNodes = []) {
  return {
    nodes: Array.from({ length: 30 }, (_, index) =>
      furnitureParagraph(`body ${index}`),
    ),
    metadata: {
      sourceParts: 0,
      warnings: [],
      headerSections: headerNodes.length
        ? [{ partName: "header1.xml", referenceType: "default", nodes: headerNodes }]
        : [],
      footerSections: footerNodes.length
        ? [{ partName: "footer1.xml", referenceType: "default", nodes: footerNodes }]
        : [],
      paragraphStyles: [],
      sectionPropertiesXml:
        '<w:sectPr><w:pgMar w:top="720" w:bottom="720" w:header="720" w:footer="720"/></w:sectPr>',
    },
  };
}

const furnitureLayoutOptions = {
  pageWidth: 400,
  pageHeight: 400,
  margin: 48,
  minLineHeight: 20,
  paragraphSpacing: 10,
};

test("DOCX pagination reserves a tall header instead of growing the page", () => {
  const baselinePages = core.layoutDocument(
    furnitureModel(),
    furnitureLayoutOptions,
  );
  const headerModel = furnitureModel(
    Array.from({ length: 4 }, (_, index) =>
      furnitureParagraph(`header ${index}`),
    ),
  );
  const headerPages = core.layoutDocument(headerModel, furnitureLayoutOptions);

  assert.ok(headerPages.length > baselinePages.length);
  // Four 20px lines + three 8px gaps + 48px header distance - 48px top margin.
  assert.equal(headerPages[0].blocks[0].y, 152);
  assert.equal(headerPages[0].blocks[0].height, 20);
});

test("DOCX pagination reserves a tall footer so body lines stay above it", () => {
  const baselinePages = core.layoutDocument(
    furnitureModel(),
    furnitureLayoutOptions,
  );
  const footerModel = furnitureModel(
    [],
    Array.from({ length: 4 }, (_, index) =>
      furnitureParagraph(`footer ${index}`),
    ),
  );
  const footerPages = core.layoutDocument(footerModel, furnitureLayoutOptions);

  assert.ok(footerPages.length > baselinePages.length);
  const lastPage = footerPages[footerPages.length - 1];
  const lastBlock = lastPage.blocks[lastPage.blocks.length - 1];
  // Body content bottom is pageHeight - margin - footer reserve (104px).
  assert.ok(lastBlock.y + lastBlock.height <= 248);
});

test("DOCX layout snapshot inherits header/footer pagination reserves", () => {
  const model = furnitureModel(
    Array.from({ length: 4 }, (_, index) =>
      furnitureParagraph(`header ${index}`),
    ),
  );
  const snapshot = core.buildLayoutSnapshot(model, furnitureLayoutOptions);
  const firstBlock = snapshot.pages[0].blocks[0];
  assert.equal(firstBlock.framePx.y, 124);
  const resolved = core.resolveDocumentForLayout(model, furnitureLayoutOptions);
  assert.equal(resolved.layout.contentBoxPx.y, 124);
  assert.equal(resolved.layout.contentBoxPx.height, 1028);

  const baseline = core.buildLayoutSnapshot(
    furnitureModel(),
    furnitureLayoutOptions,
  );
  assert.ok(snapshot.pages[0].contentBoxPx.y > baseline.pages[0].contentBoxPx.y);
  assert.ok(
    snapshot.pages[0].contentBoxPx.height < baseline.pages[0].contentBoxPx.height,
  );
});

test("DOCX pagination reserves footnote and endnote height", () => {
  const base = furnitureModel();
  const noteReference = {
    type: "paragraph",
    children: [
      { type: "text", text: "", noteReference: { kind: "footnote", id: 1 } },
    ],
    sourceXml: '<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>',
  };
  const noteNodes = Array.from({ length: 12 }, (_, index) =>
    furnitureParagraph(`note body ${index}`),
  );
  const endnoteReference = {
    type: "paragraph",
    children: [
      { type: "text", text: "", noteReference: { kind: "endnote", id: 2 } },
    ],
    sourceXml: '<w:p><w:r><w:endnoteReference w:id="2"/></w:r></w:p>',
  };
  const noteModel = {
    ...base,
    nodes: [
      noteReference,
      endnoteReference,
      ...Array.from({ length: 79 }, (_, index) =>
        furnitureParagraph(`body ${index}`),
      ),
    ],
    metadata: {
      ...base.metadata,
      footnotes: [{ id: 1, text: "", nodes: noteNodes }],
      endnotes: [{ id: 2, text: "", nodes: noteNodes }],
    },
  };

  const reserved = core.buildLayoutSnapshot(noteModel, furnitureLayoutOptions);
  const unreserved = core.buildLayoutSnapshot(noteModel, {
    ...furnitureLayoutOptions,
    reserveNoteHeight: false,
  });

  assert.ok(
    reserved.pages.length > unreserved.pages.length,
    "the note body height must move body content to the next page",
  );
  assert.ok(
    reserved.pages[0].blocks.length < unreserved.pages[0].blocks.length,
    "the page that owns the note must reserve the note body height",
  );
});

class ControlledExportWorker {
  static instances = [];

  constructor() {
    this.listeners = new Map();
    this.messages = [];
    this.terminated = false;
    ControlledExportWorker.instances.push(this);
  }

  addEventListener(type, listener) {
    this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
  }

  removeEventListener(type, listener) {
    this.listeners.set(
      type,
      (this.listeners.get(type) ?? []).filter((candidate) => candidate !== listener),
    );
  }

  postMessage(message) {
    this.messages.push(message);
  }

  terminate() {
    this.terminated = true;
  }

  emit(type, data) {
    for (const listener of this.listeners.get(type) ?? []) {
      listener({ data });
    }
  }
}

test("DOCX export Worker receives binary Uint8Array model data and reports progress", async () => {
  const { model, package: pkg } = await core.wasmBuildDocModelFromBytes(
    arrayBufferFromFile("report-with-image.docx"),
  );
  const normalized = core.normalizeDocModel(model);
  const basePackage = core.wasmPackageToMaps(pkg);
  const image = findFirstImage(normalized);
  assert.ok(image?.data instanceof Uint8Array, "model image bytes must stay binary");
  assert.ok(
    basePackage.binaryAssets.get("word/media/image1.png") instanceof Uint8Array,
    "base package image bytes must stay binary",
  );

  const worker = new ControlledExportWorker();
  const progress = [];
  const pending = core.exportDocxBuffer(normalized, basePackage, {
    workerFactory: () => worker,
    onProgress: (event) => progress.push(event),
  });

  assert.equal(worker.messages.length, 1);
  const request = worker.messages[0];
  assert.equal(request.type, "export-docx");
  assert.ok(
    findFirstImage(request.model)?.data instanceof Uint8Array,
    "the Worker request must not JSON-expand image bytes into number arrays",
  );
  assert.ok(
    request.basePackage.binaryAssets.get("word/media/image1.png") instanceof Uint8Array,
    "the Worker request must keep package binary assets typed",
  );
  assert.equal(JSON.stringify(request).includes('"data":['), false);

  const output = new Uint8Array([0x50, 0x4b, 0x03, 0x04]).buffer;
  worker.emit("message", {
    id: request.id,
    type: "progress",
    phase: "serializing",
    percent: 10,
  });
  worker.emit("message", { id: request.id, type: "success", bytes: output });

  assert.equal(await pending, output);
  assert.deepEqual(progress.map((event) => event.phase), [
    "starting",
    "serializing",
    "completed",
  ]);
  assert.equal(worker.terminated, true);
});

test("DOCX export Worker timeout terminates the Worker and reports maxExportMs", async () => {
  const worker = new ControlledExportWorker();
  await assert.rejects(
    core.exportDocxBuffer(core.createBlankDocumentModel(), undefined, {
      workerFactory: () => worker,
      limits: { maxExportMs: 5 },
    }),
    (error) => error?.name === "DocxExportError" &&
      error?.code === "TIMEOUT" &&
      error?.phase === "export" &&
      error?.limit === "maxExportMs" &&
      error?.allowed === 5 &&
      error?.actual > error?.allowed,
  );
  assert.equal(worker.terminated, true);
});

test("DOCX export Worker cancellation terminates the Worker", async () => {
  const worker = new ControlledExportWorker();
  const controller = new AbortController();
  const pending = core.exportDocxBuffer(core.createBlankDocumentModel(), undefined, {
    workerFactory: () => worker,
    signal: controller.signal,
  });
  controller.abort();
  await assert.rejects(
    pending,
    (error) => error?.name === "AbortError" && error?.code === "ABORTED",
  );
  assert.equal(worker.terminated, true);
});

test("DOCX export refuses a synchronous UI-thread fallback without a Worker", async () => {
  const originalDocument = globalThis.document;
  const originalWorker = globalThis.Worker;
  globalThis.document = {};
  delete globalThis.Worker;
  try {
    await assert.rejects(
      core.exportDocxBuffer(core.createBlankDocumentModel()),
      (error) => error?.name === "DocxExportError" &&
        error?.code === "WORKER_UNAVAILABLE",
    );
  } finally {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
  }
});

test("DOCX import refuses synchronous main-thread parsing in UI environments", async () => {
  const originalDocument = globalThis.document;
  const originalWorker = globalThis.Worker;
  const input = arrayBufferFromFile("invoice-table.docx");
  globalThis.document = {};
  try {
    await assert.rejects(
      core.importDocxBuffer(input, { useWorker: false }),
      (error) => error?.name === "DocxImportError" && error?.code === "WORKER_UNAVAILABLE",
    );

    class ThrowingWorker {
      constructor() {
        throw new Error("Worker blocked");
      }
    }
    globalThis.Worker = ThrowingWorker;
    const diagnostics = [];
    await assert.rejects(
      core.importDocxBuffer(input, {
        allowMainThreadFallback: true,
        mainThreadFallbackMaxBytes: 1024,
        onDiagnostic: (event) => diagnostics.push(event),
      }),
      (error) => error?.name === "DocxImportError" && error?.code === "WORKER_UNAVAILABLE",
    );
    assert.equal(
      diagnostics.some((event) =>
        event.type === "main-thread-start" || event.type === "main-thread-fallback"
      ),
      false,
      "a UI environment must never enter the synchronous parse path",
    );
  } finally {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
  }
});

test("DOCX headless export remains available without a Worker", async () => {
  const { model, package: pkg } = await core.wasmBuildDocModelFromBytes(
    arrayBufferFromFile("report-with-image.docx"),
  );
  const output = await core.exportDocxBuffer(
    core.normalizeDocModel(model),
    core.wasmPackageToMaps(pkg),
  );
  assert.ok(output instanceof ArrayBuffer);
  assert.deepEqual([...new Uint8Array(output).slice(0, 2)], [0x50, 0x4b]);
  const roundtrip = await core.wasmBuildDocModelFromBytes(output);
  assert.ok(roundtrip.model.nodes.length > 0);
});
