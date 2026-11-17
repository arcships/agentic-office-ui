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

test("PDF form field values map text, checkbox and choice widgets", () => {
  const options = [
    { label: "Alpha", selected: false },
    { label: "Beta", selected: true },
  ];
  assert.deepEqual(
    pdf.resolvePdfFormFieldValue({ fieldType: "text", fieldValue: "" }, "Ada"),
    { kind: "text", text: "Ada" },
  );
  assert.deepEqual(
    pdf.resolvePdfFormFieldValue({ fieldType: "checkbox" }, true),
    { kind: "checked", checked: true },
  );
  assert.deepEqual(
    pdf.resolvePdfFormFieldValue({ fieldType: "checkbox" }, "off"),
    { kind: "checked", checked: false },
  );
  assert.deepEqual(
    pdf.resolvePdfFormFieldValue({ fieldType: "combobox", options }, "Beta"),
    { kind: "selection", index: 1, isSelected: true },
  );
  assert.deepEqual(
    pdf.resolvePdfFormFieldValue({ fieldType: "listbox", options }, "0"),
    { kind: "selection", index: 0, isSelected: true },
  );
  assert.throws(
    () => pdf.resolvePdfFormFieldValue({ fieldType: "combobox", options }, "Gamma"),
    /does not match any option/,
  );
  assert.throws(
    () => pdf.resolvePdfFormFieldValue({ fieldType: "checkbox" }, "maybe"),
    /not a boolean/,
  );
  assert.throws(
    () => pdf.resolvePdfFormFieldValue({ fieldType: "radio", options }, true),
    /expects an option label or index/,
  );
});

function rawWidget(overrides = {}) {
  return {
    id: "widget-name",
    pageIndex: 0,
    type: 20,
    rect: { origin: { x: 10, y: 10 }, size: { width: 120, height: 20 } },
    flags: [],
    field: { type: 6, name: "name", alternateName: "", value: "", flag: 0 },
    ...overrides,
  };
}

test("PDF runtime writes form values and annotation appearances through cached engine objects", async () => {
  const runtime = pdf.createPdfRenderRuntime({ runtimeId: "form-writes" });
  const native = fakeNativeDocument("form-native");
  runtime.documents.set(native.id, native);
  const document = {
    id: native.id,
    pageCount: 1,
    pages: [{ index: 0, width: 600, height: 800, rotation: 0 }],
  };
  let widget = rawWidget();
  let note = {
    id: "note-1",
    pageIndex: 0,
    type: 1,
    rect: { origin: { x: 20, y: 20 }, size: { width: 30, height: 30 } },
    flags: [],
    contents: "old",
    color: "#ffff00",
    opacity: 1,
  };
  assert.ok(widget);
  assert.ok(note);
  const writes = [];
  const updates = [];
  runtime.engine = {
    getPageAnnotations() {
      return { abort() {}, async toPromise() { return [widget, note]; } };
    },
    setFormFieldValue(_doc, _page, annotation, value) {
      writes.push({ id: annotation.id, value });
      widget = { ...widget, field: { ...widget.field, value: value.kind === "text" ? value.text : "Yes" } };
      return { abort() {}, async toPromise() { return true; } };
    },
    updatePageAnnotation(_doc, _page, annotation, options) {
      updates.push({ id: annotation.id, contents: annotation.contents, color: annotation.color, options });
      note = { ...note, contents: annotation.contents, color: annotation.color };
      return { abort() {}, async toPromise() { return true; } };
    },
  };

  const written = await runtime.setFormFieldValue(document, 0, "widget-name", "Ada");
  assert.deepEqual(writes, [{ id: "widget-name", value: { kind: "text", text: "Ada" } }]);
  assert.equal(written.fieldType, "text");
  assert.equal(written.fieldValue, "Ada");

  const updated = await runtime.updateAnnotationAppearance(document, 0, "note-1", {
    contents: "reviewed",
    color: "#ff0000",
  });
  assert.equal(updates.at(-1).contents, "reviewed");
  assert.deepEqual(updates.at(-1).options, { regenerateAppearance: true });
  assert.equal(updated.contents, "reviewed");
  assert.equal(updated.color, "#ff0000");

  await assert.rejects(
    () => runtime.setFormFieldValue(document, 0, "missing", "Ada"),
    (error) => error.code === "ANNOTATION_NOT_FOUND",
  );
  await assert.rejects(
    () => runtime.setFormFieldValue(document, 0, "note-1", "Ada"),
    (error) => error.code === "ANNOTATION_NOT_WRITABLE",
  );
  await assert.rejects(
    () => runtime.updateAnnotationAppearance(document, 0, "note-1", { opacity: 2 }),
    (error) => error.code === "INVALID_RENDER_OPTIONS",
  );
  assert.equal(runtime.rawPageAnnotations.size, 1, "annotation reads are cached per page");
  await runtime.closeDocument(document);
  assert.equal(runtime.rawPageAnnotations.size, 0, "closing the document drops the annotation cache");
  await runtime.dispose();
});
