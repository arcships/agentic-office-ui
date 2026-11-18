import assert from "node:assert/strict";
import { test } from "node:test";

const pdf = await import(new URL("../../packages/vue-pdf/dist/index.js", import.meta.url).href);

function pdfBytes(text) {
  const bytes = new TextEncoder().encode(text);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function pdfTask(value) {
  return {
    abort() {},
    async toPromise() {
      return value;
    },
  };
}

function nativeDocument(id, pageCount) {
  return {
    id,
    pageCount,
    pages: Array.from({ length: pageCount }, (_, index) => ({
      index,
      size: { width: 600, height: 800 },
      rotation: 0,
    })),
  };
}

function installEngine(runtime, { pageCount, bookmarks = [], documentId = "native-nav" }) {
  runtime.engine = {
    openDocumentBuffer() {
      return pdfTask(nativeDocument(documentId, pageCount));
    },
    getBookmarks() {
      return pdfTask({ bookmarks });
    },
    closeDocument() {
      return pdfTask(true);
    },
    destroy() {
      return pdfTask(true);
    },
  };
}

async function open(runtime, text, pageCount, bookmarks = []) {
  installEngine(runtime, { pageCount, bookmarks });
  return runtime.openDocument(pdfBytes(text));
}

test("PDF runtime exposes the bookmark tree with page and URI targets", async () => {
  const runtime = pdf.createPdfRenderRuntime({ runtimeId: "outline" });
  const document = await open(
    runtime,
    "%PDF-1.4\n%%EOF",
    4,
    [
      {
        title: "  Chapter one  ",
        target: { type: "destination", destination: { pageIndex: 0 } },
        children: [
          {
            title: "Section 1.1",
            target: { type: "action", action: { type: 1, destination: { pageIndex: 2 } } },
          },
          {
            title: "External",
            target: { type: "action", action: { type: 3, uri: " https://example.com/manual " } },
          },
        ],
      },
      {
        title: "Out of range",
        target: { type: "destination", destination: { pageIndex: 99 } },
      },
      { title: "Untargeted" },
      { title: "Remote", target: { type: "action", action: { type: 2, destination: { pageIndex: 1 } } } },
      { title: "   " },
    ],
  );

  assert.deepEqual(await runtime.getOutline(document), [
    {
      title: "Chapter one",
      pageIndex: 0,
      children: [
        { title: "Section 1.1", pageIndex: 2, children: [] },
        { title: "External", url: "https://example.com/manual", children: [] },
      ],
    },
    { title: "Out of range", children: [] },
    { title: "Untargeted", children: [] },
    { title: "Remote", children: [] },
  ]);
  await runtime.closeDocument(document);
  await runtime.dispose();
});

test("PDF runtime reads page label ranges from an inline number tree", async () => {
  const runtime = pdf.createPdfRenderRuntime({ runtimeId: "labels-inline" });
  const document = await open(
    runtime,
    "%PDF-1.4\n5 0 obj\n<< /Type /Catalog /PageLabels << /Nums [0 << /S /r >> 2 << /S /D /St 5 /P (A-) >>] >> >>\nendobj\n%%EOF",
    4,
  );

  assert.deepEqual(await runtime.getPageLabels(document), ["i", "ii", "A-5", "A-6"]);
  await runtime.closeDocument(document);
  await runtime.dispose();

  const laterStart = pdf.createPdfRenderRuntime({ runtimeId: "labels-later-start" });
  const laterDocument = await open(
    laterStart,
    "%PDF-1.4\n<< /PageLabels << /Nums [1 << /S /D /St 5 >>] >> >>\n%%EOF",
    3,
  );
  assert.deepEqual(
    await laterStart.getPageLabels(laterDocument),
    ["1", "5", "6"],
    "pages before the first range keep their decimal number",
  );
  await laterStart.closeDocument(laterDocument);
  await laterStart.dispose();
});

test("PDF runtime follows indirect page label trees, kids and letter styles", async () => {
  const runtime = pdf.createPdfRenderRuntime({ runtimeId: "labels-kids" });
  const document = await open(
    runtime,
    [
      "%PDF-1.4",
      "1 0 obj << /Type /Catalog /PageLabels 2 0 R >> endobj",
      "2 0 obj << /Kids [3 0 R 4 0 R] >> endobj",
      "3 0 obj << /Nums [0 << /S /A >>] >> endobj",
      "4 0 obj << /Nums [3 << /S /a /St 27 /P <41> >>] >> endobj",
      "%%EOF",
    ].join("\n"),
    5,
  );

  assert.deepEqual(await runtime.getPageLabels(document), ["A", "B", "C", "Aaa", "Aab"]);
  await runtime.closeDocument(document);
  await runtime.dispose();
});

test("PDF runtime falls back to decimal page labels without a usable tree", async () => {
  const runtime = pdf.createPdfRenderRuntime({ runtimeId: "labels-fallback" });
  const document = await open(runtime, "%PDF-1.4\n%%EOF", 3);
  assert.deepEqual(await runtime.getPageLabels(document), ["1", "2", "3"]);
  await runtime.closeDocument(document);
  await runtime.dispose();
});

test("PDF runtime rejects navigation reads for a closed document and an aborted signal", async () => {
  const runtime = pdf.createPdfRenderRuntime({ runtimeId: "labels-closed" });
  const document = await open(runtime, "%PDF-1.4\n%%EOF", 2);

  const controller = new AbortController();
  controller.abort();
  await assert.rejects(
    runtime.getPageLabels(document, { signal: controller.signal }),
    (error) => error.name === "AbortError",
  );

  await runtime.closeDocument(document);
  await assert.rejects(runtime.getOutline(document), /no longer open/u);
  await assert.rejects(runtime.getPageLabels(document), /no longer open/u);
  await runtime.dispose();
});

test("PDF runtimes without bookmark support keep the capability optional", async () => {
  const runtime = {
    runtimeId: "injected",
    async openDocument() {
      return {
        id: "injected-document",
        pageCount: 1,
        pages: [{ index: 0, width: 600, height: 800, rotation: 0 }],
      };
    },
    async renderPage() {
      return new Blob(["page"], { type: "image/png" });
    },
  };
  const document = await runtime.openDocument();
  assert.equal(typeof runtime.getOutline, "undefined");
  assert.equal(typeof runtime.getPageLabels, "undefined");
  assert.equal(document.pageCount, 1);
});
