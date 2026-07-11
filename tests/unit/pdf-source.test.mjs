import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

const pdf = await import(new URL("../../packages/vue-pdf/dist/index.js", import.meta.url).href);
const fixture = readFileSync(new URL("../../apps/demo/public/samples/sample.pdf", import.meta.url));

test("PDF URL loading shares controlled fetch and strips query secrets", async () => {
  const calls = [];
  const result = await pdf.loadVerifiedPdfSource(
    { kind: "url", url: "sample.pdf?token=secret" },
    {
      enabled: true,
      baseUrl: "https://app.example/viewer/",
      allowRelativeUrl: true,
      allowedProtocols: ["https:"],
      allowedOrigins: ["https://app.example"],
      fetch: async (url, init) => {
        calls.push({ url, init });
        return {
          ok: true,
          status: 200,
          url,
          headers: { get: (name) => name.toLowerCase() === "content-type" ? "application/pdf" : null },
          arrayBuffer: async () => fixture.buffer.slice(fixture.byteOffset, fixture.byteOffset + fixture.byteLength),
        };
      },
    },
  );
  assert.equal(result.fileName, "sample.pdf");
  assert.equal(result.resolvedUrl, "https://app.example/viewer/sample.pdf");
  assert.equal(calls[0].init.redirect, "error");
  assert.equal(calls[0].init.credentials, "omit");
  assert.equal(JSON.stringify(result).includes("secret"), false);
});

test("PDF dangerous URL is rejected before fetch with the stable shared code", async () => {
  let requestCount = 0;
  await assert.rejects(
    pdf.loadVerifiedPdfSource(
      { kind: "url", url: "javascript:alert(1)" },
      {
        enabled: true,
        baseUrl: "https://app.example/",
        allowRelativeUrl: true,
        allowedProtocols: ["https:"],
        allowedOrigins: ["https://app.example"],
        fetch: async () => { requestCount += 1; throw new Error("must not fetch"); },
      },
    ),
    (error) => error?.name === "PdfSourceError" && error?.code === "SOURCE_NOT_ALLOWED",
  );
  assert.equal(requestCount, 0);
});

test("PDF defaults to a 50 MiB limit and rejects a local Blob before reading it", async () => {
  assert.equal(pdf.DEFAULT_PDF_MAX_FILE_SIZE, 50 * 1024 * 1024);
  let reads = 0;
  const actual = pdf.DEFAULT_PDF_MAX_FILE_SIZE + 1;
  const oversizedBlob = {
    size: actual,
    type: "application/pdf",
    async arrayBuffer() {
      reads += 1;
      throw new Error("oversized Blob must not be read");
    },
  };
  await assert.rejects(
    pdf.loadVerifiedPdfSource({ kind: "blob", blob: oversizedBlob }, undefined),
    (error) => {
      assert.equal(error?.name, "PdfSourceError");
      assert.equal(error?.code, "PDF_TOO_LARGE");
      assert.equal(error?.actual, actual);
      assert.equal(error?.allowed, pdf.DEFAULT_PDF_MAX_FILE_SIZE);
      assert.match(error?.message, new RegExp(`${actual}.*${pdf.DEFAULT_PDF_MAX_FILE_SIZE}`));
      return true;
    },
  );
  assert.equal(reads, 0);
});

test("PDF custom size limit rejects URL Content-Length before reading the response body", async () => {
  let bodyReads = 0;
  await assert.rejects(
    pdf.loadVerifiedPdfSource(
      { kind: "url", url: "https://files.example/large.pdf" },
      {
        enabled: true,
        baseUrl: "https://files.example/",
        allowedProtocols: ["https:"],
        allowedOrigins: ["https://files.example"],
        fetch: async (url) => ({
          ok: true,
          status: 200,
          url,
          headers: {
            get: (name) => {
              if (name.toLowerCase() === "content-length") return "9";
              if (name.toLowerCase() === "content-type") return "application/pdf";
              return null;
            },
          },
          arrayBuffer: async () => {
            bodyReads += 1;
            return new Uint8Array(9).buffer;
          },
        }),
      },
      { maxFileSize: 8 },
    ),
    (error) => {
      assert.equal(error?.code, "PDF_TOO_LARGE");
      assert.equal(error?.actual, 9);
      assert.equal(error?.allowed, 8);
      return true;
    },
  );
  assert.equal(bodyReads, 0);
});

test("PDF custom size limit checks URL bytes when Content-Length is absent", async () => {
  let bodyReads = 0;
  const bytes = new TextEncoder().encode("%PDF-1.4\n");
  await assert.rejects(
    pdf.loadVerifiedPdfSource(
      { kind: "url", url: "https://files.example/chunked.pdf" },
      {
        enabled: true,
        baseUrl: "https://files.example/",
        allowedProtocols: ["https:"],
        allowedOrigins: ["https://files.example"],
        fetch: async (url) => ({
          ok: true,
          status: 200,
          url,
          headers: {
            get: (name) => name.toLowerCase() === "content-type" ? "application/pdf" : null,
          },
          arrayBuffer: async () => {
            bodyReads += 1;
            return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
          },
        }),
      },
      { maxFileSize: 8 },
    ),
    (error) => {
      assert.equal(error?.code, "PDF_TOO_LARGE");
      assert.equal(error?.actual, bytes.byteLength);
      assert.equal(error?.allowed, 8);
      return true;
    },
  );
  assert.equal(bodyReads, 1);
});

test("PDF loader keeps the existing AbortSignal argument compatible", async () => {
  const controller = new AbortController();
  controller.abort();
  let reads = 0;
  await assert.rejects(
    pdf.loadVerifiedPdfSource(
      {
        kind: "blob",
        blob: {
          size: 8,
          type: "application/pdf",
          async arrayBuffer() {
            reads += 1;
            return new TextEncoder().encode("%PDF-1.4").buffer;
          },
        },
      },
      undefined,
      controller.signal,
    ),
    (error) => error?.code === "ABORTED",
  );
  assert.equal(reads, 0);
});
