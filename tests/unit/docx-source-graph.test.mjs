// Issue #131 — the DOCX document source graph was missing: hosts could not
// tell which package parts and preserved features an import read, and the
// vendored WASM predates `inspect_document_source`. The port keeps the upstream
// contract in TypeScript:
//
//   1. package parts + byte ranges + parsed roots
//   2. binary assets with null roots
//   3. preserved_only / approximate_rendering feature inventory
//   4. malformed XML is a part error, malformed packages reject cleanly
//   5. onSourceGraph fires from importDocxBuffer and DocxRuntime

import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromDocxCore.resolve("@arcships/docx-core")).href
);
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);
const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const DIAGRAM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";

function arrayBuffer(bytes) {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function sampleBytes(name) {
  return new Uint8Array(readFileSync(new URL(name, samples)));
}

function documentXml(body, extraNamespaces = "") {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${WORD_NS}" xmlns:m="${MATH_NS}" ${extraNamespaces}><w:body>${body}</w:body></w:document>`;
}

function packageBytes(document, { media = true } = {}) {
  const files = {
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
    ),
    "word/document.xml": strToU8(document),
  };
  if (media) files["word/media/picture.bin"] = new Uint8Array([1, 2, 3, 4]);
  return zipSync(files);
}

function expectedDocumentId(parts, assets = []) {
  const hash = createHash("sha256");
  const length = (value) => {
    const bytes = Buffer.alloc(8);
    bytes.writeBigUInt64LE(BigInt(value));
    return bytes;
  };
  for (const [name, content] of [...parts].sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0)) {
    const nameBytes = Buffer.from(name, "utf8");
    const contentBytes = Buffer.from(content, "utf8");
    hash.update(length(nameBytes.length));
    hash.update(nameBytes);
    hash.update(length(contentBytes.length));
    hash.update(contentBytes);
  }
  for (const [name, bytes] of [...assets].sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0)) {
    const nameBytes = Buffer.from(name, "utf8");
    hash.update(length(nameBytes.length));
    hash.update(nameBytes);
    hash.update(length(bytes.byteLength));
    hash.update(bytes);
  }
  return hash.digest("hex");
}

function findElement(node, name) {
  if (!node) return undefined;
  if (node.kind === "element" && node.name === name) return node;
  for (const child of node.children ?? []) {
    const found = findElement(child, name);
    if (found) return found;
  }
  return undefined;
}

function findAttribute(node, name, namespace) {
  if (!node) return undefined;
  if (node.kind === "element") {
    const attribute = node.attributes?.find(
      (entry) => entry.name === name && entry.namespace === namespace,
    );
    if (attribute) return attribute;
  }
  for (const child of node.children ?? []) {
    const found = findAttribute(child, name, namespace);
    if (found) return found;
  }
  return undefined;
}

function findText(node, text) {
  if (!node) return undefined;
  if (node.kind === "text" && node.text === text) return node;
  for (const child of node.children ?? []) {
    const found = findText(child, text);
    if (found) return found;
  }
  return undefined;
}

function partContent(bytes, name) {
  return strFromU8(unzipSync(bytes)[name]);
}

function slice(bytes, start, end) {
  return new TextDecoder().decode(bytes.subarray(start, end));
}

test("inspectDocumentSource reports package parts, byte ranges and binary assets", async () => {
  const bytes = sampleBytes("demo.docx");
  const graph = await core.inspectDocumentSource(bytes);

  assert.equal(graph.schemaVersion, 1);
  assert.match(graph.documentId, /^[0-9a-f]{64}$/);
  assert.equal(graph.parts.length, 18);
  assert.equal(graph.assets.length, 1);
  assert.deepEqual(
    graph.parts.map((part) => part.name),
    [...graph.parts.map((part) => part.name)].sort(),
  );

  const documentPart = graph.parts.find((part) => part.name === "word/document.xml");
  assert.ok(documentPart);
  assert.equal(documentPart.error, null);
  assert.equal(documentPart.byteLength, Buffer.byteLength(partContent(bytes, "word/document.xml")));
  assert.ok(documentPart.root);
  assert.equal(documentPart.root.kind, "document");
  assert.equal(documentPart.root.byteStart, 0);
  assert.equal(documentPart.root.byteEnd, documentPart.byteLength);
  assert.deepEqual(documentPart.root.attributes, []);

  const raw = new TextEncoder().encode(partContent(bytes, "word/document.xml"));
  const body = findElement(documentPart.root, "body");
  assert.ok(body);
  assert.equal(slice(raw, body.byteStart, body.byteEnd).slice(0, 8), "<w:body>");

  const title = findText(documentPart.root, "MASTER SERVICES AGREEMENT");
  assert.ok(title);
  assert.equal(slice(raw, title.byteStart, title.byteEnd), "MASTER SERVICES AGREEMENT");

  const xmlSpace = findAttribute(documentPart.root, "space", XML_NS);
  assert.equal(xmlSpace?.namespace, XML_NS);

  const asset = graph.assets.find((entry) => entry.name === "docProps/thumbnail.jpeg");
  assert.ok(asset);
  assert.equal(asset.byteLength, 8324);
  assert.equal(asset.root, null);
  assert.equal(asset.error, null);

  const second = await core.inspectDocumentSource(bytes);
  assert.equal(second.documentId, graph.documentId);

  const knownId = expectedDocumentId(
    [...graph.parts].map((part) => [
      part.name,
      partContent(bytes, part.name),
    ]),
    [...graph.assets].map((asset) => [asset.name, unzipSync(bytes)[asset.name]]),
  );
  assert.equal(graph.documentId, knownId);
});

test("inspectDocumentSource inventories preserved and approximate features", async () => {
  const document = documentXml(
    "<w:p><w:r><w:object/></w:r></w:p>" +
      "<w:p><w:r><w:object/></w:r></w:p>" +
      "<w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>" +
      `<w:p><w:r><w:drawing><a:graphic><a:graphicData uri="${DIAGRAM_NS}">` +
      "<dgm:dataModel/></a:graphicData></a:graphic></w:drawing></w:r></w:p>",
    `xmlns:a="${DRAWING_NS}" xmlns:dgm="${DIAGRAM_NS}"`,
  );
  const bytes = packageBytes(document);
  const graph = await core.inspectDocumentSource(bytes);

  assert.deepEqual(graph.preservedFeatures, [
    {
      partName: "word/document.xml",
      feature: "diagram",
      occurrences: 1,
      support: "preserved_only",
    },
    {
      partName: "word/document.xml",
      feature: "equation",
      occurrences: 1,
      support: "approximate_rendering",
    },
    {
      partName: "word/document.xml",
      feature: "object",
      occurrences: 2,
      support: "preserved_only",
    },
  ]);

  const documentPart = graph.parts.find((part) => part.name === "word/document.xml");
  const raw = new TextEncoder().encode(partContent(bytes, "word/document.xml"));
  const object = findElement(documentPart.root, "object");
  assert.ok(object);
  assert.equal(slice(raw, object.byteStart, object.byteEnd), "<w:object/>");
  const equation = findElement(documentPart.root, "oMath");
  assert.ok(equation);
  assert.match(slice(raw, equation.byteStart, equation.byteEnd), /^<m:oMath>/);
  const diagram = findElement(documentPart.root, "dataModel");
  assert.ok(diagram);
  assert.equal(slice(raw, diagram.byteStart, diagram.byteEnd), "<dgm:dataModel/>");
});

test("source graph byte offsets use UTF-8 bytes and survive non-ASCII text", async () => {
  const document = documentXml(
    '<w:p><w:r><w:t xml:space="preserve">é😀 &amp; done</w:t></w:r></w:p>',
  );
  const bytes = packageBytes(document, { media: false });
  const graph = await core.inspectDocumentSource(bytes);
  const part = graph.parts.find((entry) => entry.name === "word/document.xml");
  const raw = new TextEncoder().encode(partContent(bytes, "word/document.xml"));
  const text = findText(part.root, "é😀 & done");
  assert.ok(text);
  assert.equal(slice(raw, text.byteStart, text.byteEnd), "é😀 &amp; done");
  assert.equal(text.byteEnd - text.byteStart, Buffer.byteLength("é😀 &amp; done"));
  assert.equal(part.root.byteEnd, raw.byteLength);
});

test("malformed XML is isolated to its part instead of failing the package", async () => {
  const broken = `<?xml version="1.0"?>
<w:document xmlns:w="${WORD_NS}"><w:body></w:document>`;
  const bytes = packageBytes(broken);
  const graph = await core.inspectDocumentSource(bytes);
  const documentPart = graph.parts.find((entry) => entry.name === "word/document.xml");
  assert.ok(documentPart);
  assert.equal(documentPart.root, null);
  assert.match(documentPart.error, /closing tag/i);
  assert.deepEqual(graph.preservedFeatures, []);

  const rels = graph.parts.find((entry) => entry.name === "_rels/.rels");
  assert.equal(rels?.error, null);
  assert.ok(rels?.root);

  await assert.rejects(
    core.inspectDocumentSource(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x00])),
    /Invalid DOCX ZIP/,
  );

  const openDocument = zipSync({
    mimetype: strToU8("application/vnd.oasis.opendocument.text"),
    "content.xml": strToU8("<office:document-content/>"),
  });
  await assert.rejects(
    core.inspectDocumentSource(openDocument),
    /OpenDocument text has no Word package/,
  );
});

test("onSourceGraph fires from import options and a configured runtime", async () => {
  const bytes = packageBytes(documentXml("<w:p><w:r><w:t>graph</w:t></w:r></w:p>"));
  const direct = [];
  await core.importDocxBuffer(arrayBuffer(bytes), {
    useWorker: false,
    onSourceGraph: (graph) => direct.push(graph),
  });
  assert.equal(direct.length, 1);
  assert.equal(direct[0].schemaVersion, 1);
  assert.ok(direct[0].parts.some((part) => part.name === "word/document.xml"));

  const runtimeGraphs = [];
  const runtime = core.createDocxRuntime({
    onSourceGraph: (graph) => runtimeGraphs.push(graph),
  });
  try {
    await runtime.importDocxBuffer(arrayBuffer(bytes), { useWorker: false });
    assert.equal(runtimeGraphs.length, 1);
    assert.equal(runtimeGraphs[0].documentId, direct[0].documentId);
  } finally {
    runtime.dispose();
  }
});

test("the demo diagnostics panel renders preserved and approximate features", () => {
  const source = readFileSync(
    new URL("../../apps/demo/src/pages/DocxViewerPage.vue", import.meta.url),
    "utf8",
  );
  assert.match(source, /onSourceGraph:\s*recordSourceGraph/);
  assert.match(source, /data-testid="docx-source-graph"/);
  assert.match(source, /data-support="preserved_only"/);
  assert.match(source, /data-support="approximate_rendering"/);
  assert.match(source, /data-testid="docx-source-parts"/);
});

test("the Worker builds the graph only when a callback asks for it", async () => {
  const bytes = packageBytes(documentXml("<w:p><w:r><w:t>worker graph</w:t></w:r></w:p>"));
  const graph = await core.inspectDocumentSource(bytes);
  const originalWorker = globalThis.Worker;
  const workers = [];

  class FakeWorker {
    constructor() {
      this.listeners = new Map();
      this.messages = [];
      this.terminated = false;
      workers.push(this);
    }
    addEventListener(type, listener) {
      this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
    }
    removeEventListener(type, listener) {
      this.listeners.set(type, (this.listeners.get(type) ?? []).filter((item) => item !== listener));
    }
    postMessage(request) {
      this.messages.push(request);
      queueMicrotask(() => {
        const response = {
          id: request.id,
          type: "success",
          package: { parts: {}, binaryAssets: {} },
          model: { nodes: [], metadata: {} },
          timings: { totalMs: 0, parseMs: 0, buildModelMs: 0 },
          ...(request.sourceGraph ? { sourceGraph: graph } : {}),
        };
        for (const listener of this.listeners.get("message") ?? []) {
          listener({ data: response });
        }
      });
    }
    terminate() {
      this.terminated = true;
    }
  }

  globalThis.Worker = FakeWorker;
  try {
    const received = [];
    const runtime = core.createDocxRuntime({
      createWorker: () => new FakeWorker(),
      wasmUrl: "https://unit.invalid/docx.wasm",
      onSourceGraph: (value) => received.push(value),
    });
    try {
      const result = await runtime.importDocxBuffer(arrayBuffer(bytes));
      assert.equal(result.source, "worker");
      assert.equal(workers[0].messages[0].sourceGraph, true);
      assert.equal(received.length, 1);
      assert.equal(received[0].documentId, graph.documentId);
    } finally {
      runtime.dispose();
    }

    const withoutCallback = core.createDocxRuntime({
      createWorker: () => new FakeWorker(),
      wasmUrl: "https://unit.invalid/docx.wasm",
    });
    try {
      await withoutCallback.importDocxBuffer(arrayBuffer(bytes));
      assert.equal(workers[1].messages[0].sourceGraph, undefined);
    } finally {
      withoutCallback.dispose();
    }
  } finally {
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
  }
});
