import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

const pdf = await import(new URL("../../packages/vue-pdf/dist/index.js", import.meta.url).href);

function fakeNativeDocument(id) {
  return {
    id,
    pageCount: 1,
    pages: [{ index: 0, size: { width: 600, height: 800 }, rotation: 0 }],
  };
}

function installFakeEngine(runtime, passwords, acceptedPassword) {
  runtime.engine = {
    openDocumentBuffer(_file, options) {
      passwords.push(options?.password);
      return {
        abort() {},
        async toPromise() {
          if (options?.password !== acceptedPassword) {
            const error = new Error("Task rejected: password required");
            // FPDF_ERR_PASSWORD, matching PdfErrorCode.Password in @embedpdf/models.
            error.reason = { code: 4, message: "FPDF_LoadMemDocument failed" };
            throw error;
          }
          return fakeNativeDocument(`fake-native-${passwords.length}`);
        },
      };
    },
    destroy() {
      return { abort() {}, async toPromise() { return true; } };
    },
  };
}

function documentMetadata(overrides = {}) {
  return {
    id: "runtime-document-not-open",
    pageCount: 1,
    pages: [{
      index: 0,
      width: overrides.width ?? 600,
      height: overrides.height ?? 800,
      rotation: overrides.rotation ?? 0,
      ...(overrides.contentHash ? { contentHash: overrides.contentHash } : {}),
    }],
  };
}

test("PDF render parameters reject non-finite scale, dpr and rotation before the engine is touched", async () => {
  const runtime = pdf.createPdfRenderRuntime({ runtimeId: "limits" });
  const document = documentMetadata();

  for (const [options, code] of [
    [{ scale: Number.NaN, rotation: 0 }, "INVALID_RENDER_OPTIONS"],
    [{ scale: Number.POSITIVE_INFINITY, rotation: 0 }, "INVALID_RENDER_OPTIONS"],
    [{ scale: 0, rotation: 0 }, "INVALID_RENDER_OPTIONS"],
    [{ scale: 1, rotation: 0, dpr: Number.NaN }, "INVALID_RENDER_OPTIONS"],
    [{ scale: 1, rotation: 0, dpr: 0.5 }, "INVALID_RENDER_OPTIONS"],
    [{ scale: 1, rotation: 45 }, "INVALID_RENDER_OPTIONS"],
  ]) {
    await assert.rejects(
      runtime.renderPage(document, 0, options),
      (error) => error instanceof pdf.PdfRenderError
        && error.code === code
        && error.name === "PdfRenderError",
      JSON.stringify(options),
    );
  }
  await runtime.dispose();
});

test("PDF render budget rejects huge scale and dpr combinations without allocating", async () => {
  const runtime = pdf.createPdfRenderRuntime({ runtimeId: "budget" });
  const document = documentMetadata();

  for (const options of [
    { scale: 40, rotation: 0, dpr: 4 },
    { scale: 4, rotation: 0, dpr: 4 },
    { scale: 1, rotation: 0, dpr: 8 },
  ]) {
    await assert.rejects(
      runtime.renderPage(document, 0, options),
      (error) => error?.code === "PAGE_TOO_LARGE",
      JSON.stringify(options),
    );
  }
  await assert.rejects(
    runtime.renderThumbnail(document, 0, { scale: 100 }),
    (error) => error?.code === "PAGE_TOO_LARGE",
  );
  await runtime.dispose();
});

test("PDF render budget accepts a valid request and only then discovers the document is closed", async () => {
  const runtime = pdf.createPdfRenderRuntime({ runtimeId: "valid" });
  await assert.rejects(
    runtime.renderPage(documentMetadata(), 0, { scale: 1.5, rotation: 90, dpr: 2 }),
    /PDF document is no longer open/u,
  );
  await runtime.dispose();
});

test("PDF page parameters swap dimensions for quarter-turn rotations", () => {
  const parameters = pdf.resolvePdfPageRenderParameters(
    { width: 600, height: 800, rotation: 0 },
    { scale: 2, rotation: 90, dpr: 2 },
  );
  assert.deepEqual(parameters, {
    scaleFactor: 2,
    rotation: 90,
    dpr: 2,
    pixelWidth: 3200,
    pixelHeight: 2400,
  });
});

test("PDF runtime leaves URL.createObjectURL untouched across ownership lifecycle", async () => {
  const original = URL.createObjectURL;
  const runtime = pdf.createPdfRenderRuntime({ runtimeId: "no-global-patch" });
  assert.equal(URL.createObjectURL, original);
  await runtime.dispose();
  assert.equal(URL.createObjectURL, original);
});

test("PDF runtime rejects empty document bytes before creating an engine", async () => {
  const runtime = pdf.createPdfRenderRuntime({ runtimeId: "empty" });
  await assert.rejects(
    runtime.openDocument(new ArrayBuffer(0)),
    (error) => error instanceof TypeError && /non-empty ArrayBuffer/u.test(error.message),
  );
  await runtime.dispose();
});

test("PDF runtime prompts for passwords, retries with the supplied value and reports ENCRYPTED on cancel", async () => {
  const runtime = pdf.createPdfRenderRuntime({ runtimeId: "password" });
  const passwords = [];
  installFakeEngine(runtime, passwords, "open-sesame");
  const prompts = [];
  const document = await runtime.openDocument(new ArrayBuffer(8), {
    async onPassword(request) {
      prompts.push(request.attempt);
      return request.attempt === 1 ? "wrong" : "open-sesame";
    },
  });
  assert.deepEqual(passwords, [undefined, "wrong", "open-sesame"]);
  assert.deepEqual(prompts, [1, 2]);
  assert.equal(document.pageCount, 1);
  await runtime.dispose();

  const cancelled = pdf.createPdfRenderRuntime({ runtimeId: "password-cancel" });
  const cancelledAttempts = [];
  installFakeEngine(cancelled, cancelledAttempts, "open-sesame");
  await assert.rejects(
    cancelled.openDocument(new ArrayBuffer(8), { onPassword: async () => undefined }),
    (error) => error instanceof pdf.PdfRenderError && error.code === "ENCRYPTED",
  );
  assert.deepEqual(cancelledAttempts, [undefined]);
  await cancelled.dispose();
});

test("PDF embed runtime packages are exact-pinned to the same release", () => {
  const manifest = JSON.parse(readFileSync(
    new URL("../../packages/vue-pdf/package.json", import.meta.url),
    "utf8",
  ));
  const dependencies = { ...manifest.dependencies, ...manifest.devDependencies };
  const embedPdf = Object.entries(dependencies)
    .filter(([name]) => name.startsWith("@embedpdf/"));
  assert.ok(embedPdf.length >= 14);
  for (const [name, version] of embedPdf) {
    assert.match(version, /^\d+\.\d+\.\d+$/u, `${name} must be exact-pinned`);
  }
  assert.deepEqual(
    new Set(embedPdf.map(([, version]) => version)),
    new Set(["2.14.4"]),
  );
});

test("PDF encrypted runtime errors map to the public ENCRYPTED load code", () => {
  const error = new pdf.PdfRenderError("ENCRYPTED", "PDF document is encrypted and requires a password.");
  const normalized = pdf.toPdfLoadError(error, "url", "https://files.example/secret.pdf?token=hidden");
  assert.equal(normalized.code, "ENCRYPTED");
  assert.equal(normalized.url, "https://files.example/secret.pdf");
  assert.equal(JSON.stringify(normalized).includes("hidden"), false);
});
