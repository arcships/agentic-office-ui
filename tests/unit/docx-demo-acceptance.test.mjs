// Issue #58 — acceptance paths must not pass without exercising the product.
//
// Covered here without a browser:
//   1. the parity comparator rejects zero-page and partially mounted structures
//   2. the demo pages derive paper size from the document instead of hard-coding
//      Letter, and the A4/landscape/mixed fixture really carries those sizes
//   3. the editor page no longer exposes test-only table/image/format buttons;
//      acceptance uses the real toolbar buttons instead
//   4. the surface page loads through a latest-wins/abort gate
//
// The browser suites (docx_parity.py, docx_editing.py, e2e_workflows.py,
// race_workflows.py) exercise the same paths in a real Chromium.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

import { createTsModuleLoader } from "./ts-source-loader.mjs";

const requireFromDemo = createRequire(
  new URL("../../apps/demo/package.json", import.meta.url),
);
const typescript = requireFromDemo("typescript");
const loadDemoSource = createTsModuleLoader(
  new URL("../../apps/demo/src/", import.meta.url),
  typescript,
);
const parity = await import(loadDemoSource("lib/docx-parity.ts"));
const latestRequest = await import(loadDemoSource("lib/latest-request.ts"));
const core = await import(
  pathToFileURL(requireFromDemo.resolve("@arcships/docx-core")).href
);

const PAGES = new URL("../../apps/demo/src/pages/", import.meta.url);
const SAMPLES = new URL("../../apps/demo/public/samples/", import.meta.url);

function pageShape(index, overrides = {}) {
  return {
    index,
    nodes: [],
    text: `page ${index}`,
    tables: 0,
    cells: 0,
    images: 0,
    header: { text: "Fixture Header", images: [] },
    footer: { text: "Fixture Footer", images: [] },
    ...overrides,
  };
}

function structure(pages, meta = {}) {
  return parity.buildRenderStructure(pages, parity.EMPTY_INTERACTION(), {
    declaredPageCount: pages.length,
    scannedSteps: 3,
    ...meta,
  });
}

function sampleBytes(name) {
  const bytes = readFileSync(new URL(name, SAMPLES));
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

test("parity comparator fails when both sides render zero pages", () => {
  const empty = structure([]);
  const result = parity.compareRenderStructures(empty, empty);
  assert.equal(result.equal, false);
  assert.match(result.differences.join(" | "), /未渲染任何页面/);
});

test("parity comparator fails when only one side renders pages", () => {
  const viewer = structure([pageShape(0), pageShape(1)]);
  const editor = structure([]);
  const result = parity.compareRenderStructures(viewer, editor);
  assert.equal(result.equal, false);
  assert.match(result.differences.join(" | "), /Editor 未渲染任何页面/);
});

test("parity comparator fails when a declared page was never mounted", () => {
  const viewer = structure([pageShape(0), pageShape(1), pageShape(2)], {
    declaredPageCount: 5,
  });
  const editor = structure([pageShape(0), pageShape(1), pageShape(2)], {
    declaredPageCount: 3,
  });
  const result = parity.compareRenderStructures(viewer, editor);
  assert.equal(result.equal, false);
  assert.match(result.differences.join(" | "), /3\/5 页/);
});

test("parity comparator passes only for complete, identical structures", () => {
  const pages = [pageShape(0), pageShape(1)];
  const viewer = structure(pages, { scannedSteps: 1 });
  const editor = structure(pages, { scannedSteps: 9 });
  // Collection metadata is evidence, not part of the compared document.
  assert.deepEqual(parity.compareRenderStructures(viewer, editor), {
    equal: true,
    differences: [],
  });
});

test("parity comparator reports header and footer differences", () => {
  const viewer = structure([
    pageShape(0, { header: { text: "Header A", images: [] } }),
  ]);
  const editor = structure([
    pageShape(0, { header: { text: "", images: [] } }),
  ]);
  const result = parity.compareRenderStructures(viewer, editor);
  assert.equal(result.equal, false);
  assert.match(result.differences.join(" | "), /header\.text/);
});

test("latest-request gate aborts the previous request and rejects stale tokens", () => {
  const gate = latestRequest.createLatestRequestGate();
  const first = gate.begin();
  assert.equal(first.isCurrent(), true);
  const second = gate.begin();
  assert.equal(first.isCurrent(), false, "the first request must become stale");
  assert.equal(first.signal.aborted, true, "the first fetch must be aborted");
  assert.equal(second.isCurrent(), true);
  assert.equal(gate.isCurrent(first.token), false);
  assert.equal(gate.isCurrent(second.token), true);
  gate.abort();
  assert.equal(second.isCurrent(), false);
  assert.equal(second.signal.aborted, true);
});

test("demo DOCX pages do not hard-code Letter geometry", () => {
  for (const name of ["DocxViewerPage.vue", "DocxSurfacePage.vue"]) {
    const source = readFileSync(new URL(name, PAGES), "utf8");
    assert.doesNotMatch(
      source,
      /layout-options/,
      `${name} must let resolveDocumentLayout() read each section's w:pgSz`,
    );
    assert.doesNotMatch(
      source,
      /pageWidth:\s*816/,
      `${name} must not override the paper size with US Letter`,
    );
    assert.match(
      source,
      /page-sizes\.docx/,
      `${name} must offer the A4/landscape/mixed-section fixture`,
    );
  }
});

test("the mixed page-size fixture carries A4, landscape and Letter sections", async () => {
  const { model } = await core.buildDocModelFromBytes(sampleBytes("page-sizes.docx"));
  assert.equal(model.metadata.sections?.length, 3);
  const layouts = model.metadata.sections.map((section) =>
    core.parseSectionLayout(section.sectionPropertiesXml),
  );
  assert.deepEqual(
    layouts.map((layout) => [layout.pageWidthPx, layout.pageHeightPx]),
    [
      [794, 1123],
      [1123, 794],
      [816, 1056],
    ],
  );
  assert.deepEqual(
    [core.resolveDocumentLayout(model).pageWidthPx, core.resolveDocumentLayout(model).pageHeightPx],
    [794, 1123],
    "the demo must derive the first page from the document, not the host",
  );

  const headers = (model.metadata.headerSections ?? []).map((section) =>
    collectText(section).join(" "),
  );
  assert.deepEqual(headers, [
    "A4 Portrait Header",
    "A4 Landscape Header",
    "Letter Portrait Header",
  ]);
  const footers = (model.metadata.footerSections ?? []).map((section) =>
    collectText(section).join(" "),
  );
  assert.deepEqual(footers, [
    "A4 Portrait Footer",
    "A4 Landscape Footer",
    "Letter Portrait Footer",
  ]);
});

test("editor acceptance no longer drives table/image/format through test-only buttons", () => {
  const editor = readFileSync(new URL("DocxEditorPage.vue", PAGES), "utf8");
  for (const forbidden of [
    "editor-test-format-range",
    "editor-test-insert-table",
    "editor-test-insert-image",
    "setTableColumnWidths(",
    "setImageWrapMode(",
    "moveFloatingImage(",
  ]) {
    assert.equal(
      editor.includes(forbidden),
      false,
      `DocxEditorPage must not drive acceptance through ${forbidden}`,
    );
  }
  assert.match(editor, /editor-api-fixture-panel/);
  assert.match(editor, /editor-export-bytes-fixture/);

  const toolbar = readFileSync(
    new URL("../../packages/vue-docx/src/components/DocxToolbar.vue", import.meta.url),
    "utf8",
  );
  assert.match(toolbar, /data-testid="editor-insert-table"/);
  assert.match(toolbar, /data-testid="editor-insert-image"/);
});

test("parity page compares through the shared guard without a fixed scan cap", () => {
  const source = readFileSync(new URL("DocxParityPage.vue", PAGES), "utf8");
  assert.match(source, /compareRenderStructures/);
  assert.doesNotMatch(source, /Math\.min\(\s*64/, "the 64-step scan cap must stay gone");
  assert.match(source, /collectVirtualStructure\(editorHost\.value, viewer\.pageCount/);
});

test("surface demo loads through the latest-wins gate", () => {
  const surface = readFileSync(new URL("DocxSurfacePage.vue", PAGES), "utf8");
  assert.match(surface, /createLatestRequestGate/);
  assert.match(surface, /request\.signal\.addEventListener\(\"abort\"/);
  assert.match(surface, /data-testid="page-status"/);
  assert.match(surface, /data-testid="loaded-file"/);
});
