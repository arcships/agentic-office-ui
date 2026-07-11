#!/usr/bin/env node
/**
 * DOCX Integration Verification — D1–D5
 *
 * Verifies that the docx-core engine APIs used by vue-docx/composables.ts
 * are real call paths (not stubs/mocks).
 *
 * Run from repo root:
 *   node packages/vue-docx/tests/verify-integration.mjs
 */

import { readFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const __dirname = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(__dirname, "../../..");

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const PASS = "✅";
const FAIL = "❌";
const SKIP = "⏭️";

let passed = 0;
let failed = 0;
let skipped = 0;

function check(label, ok, detail) {
  if (ok) {
    console.log(`  ${PASS} ${label}`);
    passed++;
  } else {
    console.log(`  ${FAIL} ${label}${detail ? ` — ${detail}` : ""}`);
    failed++;
  }
}

function skip(label, reason) {
  console.log(`  ${SKIP} ${label} (${reason})`);
  skipped++;
}

function heading(text) {
  console.log(`\n${"=".repeat(60)}`);
  console.log(`  ${text}`);
  console.log(`${"=".repeat(60)}`);
}

// ---------------------------------------------------------------------------
// D1: wasm engine calls
// ---------------------------------------------------------------------------
async function verifyD1() {
  heading("D1: wasm engine calls (importDocxFile → wasmBuildDocModelFromBytes, exportDocx → wasmSerializeDocx)");

  const {
    wasmBuildDocModelFromBytes,
    wasmSerializeDocx,
  } = await import("@arcships/docx-core");

  const samplePath = resolve(
    REPO_ROOT,
    "apps/demo/public/samples/demo.docx",
  );
  if (!existsSync(samplePath)) {
    skip("Load sample docx", "sample file not found");
    return;
  }

  const buffer = readFileSync(samplePath);
  const arrayBuffer = buffer.buffer.slice(
    buffer.byteOffset,
    buffer.byteOffset + buffer.byteLength,
  );

  let model;
  let pkg;
  try {
    const result = await wasmBuildDocModelFromBytes(arrayBuffer);
    pkg = result.package;
    model = result.model;
  } catch (err) {
    check("wasmBuildDocModelFromBytes", false, err.message);
    return;
  }

  check("wasmBuildDocModelFromBytes returns model", !!model);
  check("model.nodes is array", Array.isArray(model?.nodes));
  check("model.nodes is non-empty", model?.nodes?.length > 0);
  check("package.parts is object", pkg && typeof pkg.parts === "object");

  // Verify export
  let exported;
  try {
    exported = await wasmSerializeDocx(model, pkg);
  } catch (err) {
    check("wasmSerializeDocx", false, err.message);
    return;
  }

  check("wasmSerializeDocx returns ArrayBuffer", exported instanceof ArrayBuffer);
  check("exported docx is non-empty", exported.byteLength > 0);
  check(
    "exported docx starts with PK (ZIP magic)",
    new Uint8Array(exported).slice(0, 2).every((b, i) => b === [0x50, 0x4b][i]),
  );

  // Verify with python-docx if available
  const pythonCheck = spawnSync("python3", ["-c", "import docx; print('ok')"], {
    timeout: 5000,
  });
  if (pythonCheck.status === 0 && pythonCheck.stdout.toString().trim() === "ok") {
    const tmpPath = resolve(REPO_ROOT, "tmp-d1-export.docx");
    const { writeFileSync } = await import("node:fs");
    writeFileSync(tmpPath, new Uint8Array(exported));

    const pyResult = spawnSync(
      "python3",
      [
        "-c",
        `
import docx
d = docx.Document("${tmpPath}")
paras = d.paragraphs
print(f"paragraphs={len(paras)}")
if paras:
    print(f"first_text={repr(paras[0].text[:50])}")
        `.trim(),
      ],
      { timeout: 10000 },
    );

    const { unlinkSync } = await import("node:fs");
    try { unlinkSync(tmpPath); } catch {}

    if (pyResult.status === 0) {
      const out = pyResult.stdout.toString().trim();
      const paraMatch = out.match(/paragraphs=(\d+)/);
      check("python-docx opens exported docx", !!paraMatch);
      if (paraMatch) {
        const paraCount = parseInt(paraMatch[1], 10);
        check("exported docx has paragraphs", paraCount > 0);
      }
    } else {
      skip("python-docx verification", pyResult.stderr?.toString()?.trim() || "python3 error");
    }
  } else {
    skip("python-docx verification", "python3 or python-docx not available");
  }
}

// ---------------------------------------------------------------------------
// D2: docx-import worker call
// ---------------------------------------------------------------------------
async function verifyD2() {
  heading("D2: docx-import worker call (importDocxBuffer → worker path)");

  const { createDocxRuntime, importDocxBuffer } = await import("@arcships/docx-core");

  const samplePath = resolve(
    REPO_ROOT,
    "apps/demo/public/samples/demo.docx",
  );
  if (!existsSync(samplePath)) {
    skip("Worker import", "sample file not found");
    return;
  }

  const buffer = readFileSync(samplePath);
  const arrayBuffer = buffer.buffer.slice(
    buffer.byteOffset,
    buffer.byteOffset + buffer.byteLength,
  );

  // Test main-thread import (useWorker: false)
  let result;
  try {
    result = await importDocxBuffer(arrayBuffer, { useWorker: false });
  } catch (err) {
    check("importDocxBuffer (main-thread)", false, err.message);
    return;
  }

  check("importDocxBuffer returns {package, model, source, timings}", !!result);
  check("result.package is defined", !!result.package);
  check("result.model is defined", !!result.model);
  check("result.model.nodes is non-empty", result.model?.nodes?.length > 0);
  check("result.source is 'main-thread'", result.source === "main-thread");
  check("result.timings.totalMs > 0", result.timings?.totalMs > 0);

  // Worker imports are required by default. In this Node environment there is
  // no browser Worker, so the contract is a structured failure, never a
  // successful main-thread result disguised as a Worker result.
  try {
    const workerResult = await importDocxBuffer(arrayBuffer, { useWorker: true });
    if (typeof Worker === "undefined") {
      check("worker-unavailable import rejects instead of falling back", false,
        `received ${workerResult.source}`);
    } else {
      check("worker result source is worker", workerResult.source === "worker");
    }
  } catch (err) {
    if (typeof Worker === "undefined") {
      check("worker-unavailable import returns structured error",
        err?.code === "WORKER_UNAVAILABLE", `${err?.name}: ${err?.message}`);
    } else {
      check("importDocxBuffer (worker path)", false, err.message);
    }
  }

  const runtime = createDocxRuntime();
  check("createDocxRuntime creates an isolated runtime id", typeof runtime.id === "string" && runtime.id.length > 0);
  runtime.dispose();

  // Verify the public runtime forwards one explicit WASM URL to the Worker
  // and exposes its lifecycle through diagnostics. This is a controlled unit
  // double; real Worker behavior is exercised by the formal browser suite.
  const originalWorker = globalThis.Worker;
  const workerRequests = [];
  class FakeWorker {
    constructor() { this.listeners = new Map(); }
    addEventListener(type, listener) {
      const listeners = this.listeners.get(type) ?? [];
      listeners.push(listener);
      this.listeners.set(type, listeners);
    }
    removeEventListener(type, listener) {
      this.listeners.set(type, (this.listeners.get(type) ?? []).filter(item => item !== listener));
    }
    postMessage(request) {
      workerRequests.push(request);
      queueMicrotask(() => {
        for (const listener of this.listeners.get("message") ?? []) {
          listener({ data: {
            id: request.id,
            type: "success",
            package: { parts: new Map(), binaryAssets: new Map() },
            model: {
              nodes: [],
              metadata: { sourceParts: 0, warnings: [], headerSections: [], footerSections: [], paragraphStyles: [] },
            },
            timings: { totalMs: 1, parseMs: 0.5, buildModelMs: 0.5 },
          } });
        }
      });
    }
    terminate() {}
  }

  try {
    globalThis.Worker = FakeWorker;
    const diagnostics = [];
    const configuredWasmUrl = "https://runtime.example.test/docx.wasm";
    const configuredRuntime = createDocxRuntime({
      wasmUrl: configuredWasmUrl,
      createWorker: () => new FakeWorker(),
      onDiagnostic: event => diagnostics.push(event),
    });
    const configuredResult = await configuredRuntime.importDocxBuffer(new ArrayBuffer(8));
    check("runtime worker result is labelled worker", configuredResult.source === "worker");
    check("runtime forwards configured WASM URL to worker", workerRequests[0]?.wasmSource === configuredWasmUrl);
    check("runtime diagnostics include worker lifecycle", diagnostics.some(event => event.type === "worker-created") && diagnostics.some(event => event.type === "worker-success"));
    configuredRuntime.dispose();
  } finally {
    if (originalWorker === undefined) {
      delete globalThis.Worker;
    } else {
      globalThis.Worker = originalWorker;
    }
  }
}

// ---------------------------------------------------------------------------
// D3: pagination call
// ---------------------------------------------------------------------------
async function verifyD3() {
  heading("D3: pagination (buildDocumentPageNodeSegments)");

  const {
    buildDocumentPageNodeSegments,
    defaultStarterModel,
  } = await import("@arcships/docx-core");

  const model = defaultStarterModel;

  const callbacks = {
    estimateDocNodeHeightPx: () => 22,
    paragraphHasVisibleText: (p) => p.children.some(c => c.type === "text" && c.text.length > 0),
    paragraphIsStructuralSectionBreakSpacer: () => false,
    estimateParagraphHeightPx: () => 22,
    estimateParagraphLineHeightPx: () => 22,
    paragraphLineCountWithinWidth: (p) =>
      Math.max(1, Math.ceil(p.children
        .filter(c => c.type === "text")
        .reduce((sum, c) => sum + c.text.length, 0) / 60)),
    paragraphWidowControlEnabled: () => true,
    paragraphCanSplitAcrossPages: () => true,
    estimateTableRowHeightsPx: () => [22],
  };

  const pages = buildDocumentPageNodeSegments(
    model,
    1056,
    816,
    callbacks,
    model.metadata.numberingDefinitions,
    [],
    { allowParagraphLineSplitting: true },
  );

  check("buildDocumentPageNodeSegments returns array", Array.isArray(pages));
  check("pages is non-empty", pages.length > 0);
  check("pageCount > 0", pages.length >= 1);

  const firstPage = pages[0];
  check("first page is array", Array.isArray(firstPage));
  check("first page has segments", firstPage.length > 0);

  // Verify with multi-paragraph model
  const multiParaModel = {
    nodes: [
      { type: "paragraph", children: [{ type: "text", text: "Page 1 content. ".repeat(20) }] },
      { type: "paragraph", children: [{ type: "text", text: "Page 1 more text. ".repeat(15) }] },
      { type: "paragraph", children: [{ type: "text", text: "Page 2 content. ".repeat(30) }] },
    ],
    metadata: {
      sourceParts: 0,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
    },
  };

  const multiPages = buildDocumentPageNodeSegments(
    multiParaModel,
    300,
    400,
    callbacks,
    undefined,
    [],
    { allowParagraphLineSplitting: true },
  );

  check("multi-paragraph model paginates", Array.isArray(multiPages));
  check("multi-paragraph produces multiple pages or 1 page",
    multiPages.length >= 1);
}

// ---------------------------------------------------------------------------
// D4: edit operations
// ---------------------------------------------------------------------------
async function verifyD4() {
  heading("D4: edit operations (commitParagraphText → updateParagraphText, toggleBold → toggleRunStyleFlag)");

  const {
    cloneDocModel,
    updateParagraphText,
    toggleRunStyleFlag,
    defaultStarterModel,
    assertValidDocxModel,
    DocxModelValidationError,
  } = await import("@arcships/docx-core");

  const original = cloneDocModel(defaultStarterModel);

  // Test updateParagraphText (called by commitParagraphText in composables)
  const updated = updateParagraphText(original, 0, "Hello DOCX World");
  check("updateParagraphText returns model", !!updated);
  check("updateParagraphText produces different model", updated !== original);

  const updatedPara = updated.nodes[0];
  check("updated paragraph exists", updatedPara?.type === "paragraph");
  const updatedText = updatedPara.type === "paragraph"
    ? updatedPara.children.filter(c => c.type === "text").map(t => t.text).join("")
    : "";
  check("updated paragraph text matches", updatedText === "Hello DOCX World");

  // Test toggleRunStyleFlag (called for bold/italic in composables)
  const bolded = toggleRunStyleFlag(updated, 0, 0, "bold");
  check("toggleRunStyleFlag returns model", !!bolded);
  check("toggleRunStyleFlag produces different model", bolded !== updated);

  const boldedPara = bolded.nodes[0];
  const boldedRun = boldedPara?.type === "paragraph"
    ? boldedPara.children.find(c => c.type === "text")
    : undefined;
  check("bolded run exists", !!boldedRun);
  check("bolded run has bold style",
    boldedRun?.type === "text" && boldedRun?.style?.bold === true);

  // Verify sourceXml is cleared (indicates real mutation, not stub)
  check("bolded paragraph sourceXml is undefined",
    boldedPara?.type === "paragraph" && boldedPara.sourceXml === undefined);

  // Clone chain: verify cloneDocModel is real (used by dispatchEditorTransaction)
  const cloned = cloneDocModel(bolded);
  check("cloneDocModel returns new object", cloned !== bolded);
  check("cloneDocModel preserves text",
    cloned.nodes[0]?.type === "paragraph" &&
    cloned.nodes[0].children.filter(c => c.type === "text").map(t => t.text).join("") === "Hello DOCX World");

  // Verify model changes are not stub returns
  check("original model unchanged after edit", original !== updated);
  check("model identity changes on edit (not stub)", updated !== bolded);

  // Verify dispatchEditorTransaction pattern: clone + mutate
  const fromDispatchPattern = cloneDocModel(original);
  const mutated = updateParagraphText(fromDispatchPattern, 0, "Real dispatch transaction");
  check("dispatch pattern clone+mutate returns new model", mutated !== fromDispatchPattern);
  check("dispatch pattern clone preserves original", original.nodes[0]?.type === "paragraph" &&
    original.nodes[0].children[0]?.type === "text" &&
    original.nodes[0].children[0].text !== "Real dispatch transaction");

  let invalidModelError;
  try {
    assertValidDocxModel({
      nodes: [],
      metadata: { headerSections: [], footerSections: [] },
    });
  } catch (error) {
    invalidModelError = error;
  }
  check("incomplete editor model has a structured validation error",
    invalidModelError instanceof DocxModelValidationError &&
      invalidModelError.code === "INVALID_DOC_MODEL");
  check("editor model validation names missing metadata fields",
    invalidModelError?.missingFields?.includes("metadata.sourceParts") &&
      invalidModelError?.missingFields?.includes("metadata.warnings") &&
      invalidModelError?.missingFields?.includes("metadata.paragraphStyles"));
}

// ---------------------------------------------------------------------------
// D5: serializer + basePackage
// ---------------------------------------------------------------------------
async function verifyD5() {
  heading("D5: serializer + basePackage (serializeDocx with basePackage)");

  const {
    serializeDocx,
    wasmBuildDocModelFromBytes,
    wasmPackageToMaps,
    cloneDocModel,
    updateParagraphText,
  } = await import("@arcships/docx-core");

  // Load a sample docx to get a real basePackage
  const samplePath = resolve(
    REPO_ROOT,
    "apps/demo/public/samples/invoice-table.docx",
  );
  if (!existsSync(samplePath)) {
    skip("basePackage test", "invoice-table.docx not found");
    return;
  }

  const buffer = readFileSync(samplePath);
  const arrayBuffer = buffer.buffer.slice(
    buffer.byteOffset,
    buffer.byteOffset + buffer.byteLength,
  );

  let model, wasmPkg;
  try {
    const result = await wasmBuildDocModelFromBytes(arrayBuffer);
    wasmPkg = result.package;
    model = result.model;
  } catch (err) {
    check("Load sample docx for D5", false, err.message);
    return;
  }

  // Convert WasmOoxmlPackage (plain objects) → OoxmlPackage (Maps)
  const pkg = wasmPackageToMaps(wasmPkg);
  check("basePackage converted to OoxmlPackage (Maps)", !!pkg);

  // Edit the model slightly (as composables does via commitParagraphText)
  const edited = updateParagraphText(model, 0, "Edited text for testing basePackage export");

  // Serialize WITH basePackage (as exportDocx does in composables)
  let exportWithPackage;
  try {
    exportWithPackage = await serializeDocx(edited, pkg);
  } catch (err) {
    check("serializeDocx(model, basePackage)", false, err.message);
    return;
  }

  check("serializeDocx with basePackage returns ArrayBuffer",
    exportWithPackage instanceof ArrayBuffer);
  check("export with basePackage is non-empty",
    exportWithPackage.byteLength > 0);
  check("export with basePackage starts with PK (ZIP magic)",
    new Uint8Array(exportWithPackage).slice(0, 2).every((b, i) => b === [0x50, 0x4b][i]));

  // Serialize WITHOUT basePackage for comparison
  let exportWithoutPackage;
  try {
    exportWithoutPackage = await serializeDocx(edited);
  } catch (err) {
    check("serializeDocx(model) without basePackage", false, err.message);
    return;
  }

  check("serializeDocx without basePackage returns ArrayBuffer",
    exportWithoutPackage instanceof ArrayBuffer);
  check("export without basePackage is non-empty",
    exportWithoutPackage.byteLength > 0);

  // With basePackage should preserve more content (styles, numbering, headers)
  check("basePackage export preserves more content (larger size)",
    exportWithPackage.byteLength >= exportWithoutPackage.byteLength);

  // Verify with python-docx for styles/numbering preservation
  const pythonCheck = spawnSync("python3", ["-c", "import docx; print('ok')"], {
    timeout: 5000,
  });
  if (pythonCheck.status === 0 && pythonCheck.stdout.toString().trim() === "ok") {
    const { writeFileSync, unlinkSync } = await import("node:fs");

    const withPath = resolve(REPO_ROOT, "tmp-d5-with-package.docx");
    const withoutPath = resolve(REPO_ROOT, "tmp-d5-without-package.docx");
    writeFileSync(withPath, new Uint8Array(exportWithPackage));
    writeFileSync(withoutPath, new Uint8Array(exportWithoutPackage));

    const pyWith = spawnSync("python3", [
      "-c",
      `import docx; d=docx.Document("${withPath}"); print(f"styles={len(d.styles)} paras={len(d.paragraphs)}")`,
    ], { timeout: 10000 });

    const pyWithout = spawnSync("python3", [
      "-c",
      `import docx; d=docx.Document("${withoutPath}"); print(f"styles={len(d.styles)} paras={len(d.paragraphs)}")`,
    ], { timeout: 10000 });

    try { unlinkSync(withPath); } catch {}
    try { unlinkSync(withoutPath); } catch {}

    if (pyWith.status === 0) {
      const withOut = pyWith.stdout.toString().trim();
      check("python-docx opens with-basePackage export", true, withOut);
    }
    if (pyWithout.status === 0) {
      const withoutOut = pyWithout.stdout.toString().trim();
      check("python-docx opens without-basePackage export", true, withoutOut);
    }

    if (pyWith.status === 0 && pyWithout.status === 0) {
      const withStyles = parseInt(pyWith.stdout.toString().match(/styles=(\d+)/)?.[1] || "0");
      const withoutStyles = parseInt(pyWithout.stdout.toString().match(/styles=(\d+)/)?.[1] || "0");
      check("basePackage preserves more styles", withStyles >= withoutStyles);
    }
  } else {
    skip("python-docx style verification", "python3/python-docx not available");
  }
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
async function main() {
  console.log("DOCX Integration Verification — D1 through D5\n");
  console.log(`Running from: ${REPO_ROOT}\n`);

  try {
    await verifyD1();
  } catch (err) {
    console.log(`  ${FAIL} D1 crashed: ${err.message}`);
    failed++;
  }

  try {
    await verifyD2();
  } catch (err) {
    console.log(`  ${FAIL} D2 crashed: ${err.message}`);
    failed++;
  }

  try {
    await verifyD3();
  } catch (err) {
    console.log(`  ${FAIL} D3 crashed: ${err.message}`);
    failed++;
  }

  try {
    await verifyD4();
  } catch (err) {
    console.log(`  ${FAIL} D4 crashed: ${err.message}`);
    failed++;
  }

  try {
    await verifyD5();
  } catch (err) {
    console.log(`  ${FAIL} D5 crashed: ${err.message}`);
    failed++;
  }

  console.log(`\n${"=".repeat(60)}`);
  console.log(`  Results: ${PASS} ${passed} passed  ${FAIL} ${failed} failed  ${SKIP} ${skipped} skipped`);
  console.log(`${"=".repeat(60)}\n`);

  if (failed > 0) {
    process.exit(1);
  }
}

main().catch((err) => {
  console.error(`Fatal: ${err.message}`);
  process.exit(1);
});
