import assert from "node:assert/strict";
import { readFileSync, existsSync } from "node:fs";
import test from "node:test";

import { findByTestId, importFromDemo, mount, walk, waitFor } from "./vue-test-renderer.mjs";

const { DocxViewer } = await importFromDemo("@arcships/vue-docx");
const { PdfSurface, PdfViewer } = await importFromDemo("@arcships/vue-pdf");

const ROOT = new URL("../../", import.meta.url);

function read(relativePath) {
  return readFileSync(new URL(relativePath, ROOT), "utf8");
}

function assertPrintRule(source, selector, label) {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const rule = new RegExp(`${escaped}[\\s\\S]{0,180}?display:\\s*none`, "i");
  assert.match(source, rule, `${label} must be hidden for print`);
}

test("PDF viewer exposes dark mode, fluid height and container breakpoints", () => {
  const viewer = read("packages/vue-pdf/src/components/PdfViewer.vue");
  const types = read("packages/vue-pdf/src/types.ts");
  const surface = read("packages/vue-pdf/src/components/PdfSurface.vue");

  assert.match(types, /minHeight\?: string \| number/, "PdfViewerProps must expose an optional minHeight");
  assert.match(types, /isDark\?: boolean/, "PdfViewerProps must expose isDark");
  assert.match(viewer, /isDark: false/);
  assert.match(viewer, /pdf-viewer--dark/, "PdfViewer must render the dark root class");
  assert.doesNotMatch(
    viewer,
    /invert\(1\) hue-rotate\(180deg\)/,
    "dark mode must not invert page content by default (#199)",
  );
  assert.match(
    viewer,
    /--pdf-page-filter: var\(--pdf-dark-page-filter, var\(--pdf-page-softening\)\)/,
    "dark PDF pages keep the softening and only opt into a host filter",
  );
  assert.match(
    viewer,
    /--pdf-page-softening: blur\(0\.35px\)/,
    "page rasters get the upstream 0.35px softening (#199)",
  );
  assert.match(
    viewer,
    /filter: var\(--pdf-page-filter\)/,
    "the page image applies the softening token without changing its box",
  );
  assert.match(
    viewer,
    /--pdf-canvas-bg: color-mix\(in oklab, var\(--office-muted\) 30%, transparent\)/,
    "the canvas must be the upstream bg-muted/30 surface (#199)",
  );
  assert.match(
    viewer,
    /--pdf-bg: var\(--office-background\)/,
    "the viewer shell background follows the upstream bg-background token (#199)",
  );
  assert.doesNotMatch(
    viewer,
    /--pdf-canvas-bg: #(?:e5e7eb|020617)/,
    "the old canvas literals must be gone (#199)",
  );
  assert.match(
    viewer,
    /\.pdf-frame-wrap \{[^}]*padding: 24px/,
    "the canvas padding stays 24px at every container width (#199)",
  );
  assert.doesNotMatch(
    viewer,
    /\.pdf-frame-wrap \{ padding: 12px/,
    "the mobile 12px canvas padding override must be gone (#199)",
  );
  assert.match(
    viewer,
    /const showThumbnails = ref\(false\)/,
    "the thumbnails sidebar starts closed like upstream (#199)",
  );
  assert.match(viewer, /showStatusBar: false/, "the optional status bar stays opt-in (#199)");
  assert.doesNotMatch(viewer, /min-height: (?:620|540|220|180)px/, "viewer chrome must not use fixed min-heights that overflow short containers");
  assert.match(viewer, /container: pdf-viewer \/ inline-size/);
  assert.match(viewer, /@container pdf-viewer \(max-width: 760px\)/, "PDF breakpoints must follow the container, not the viewport");
  assert.doesNotMatch(viewer, /@media \(max-width: 760px\)/);
  assert.match(viewer, /@media print/);
  assertPrintRule(viewer, ".pdf-toolbar", "PDF toolbar");
  assert.match(viewer, /@page \{ size: \$\{/, "PDF print page size must follow the rendered page");
  assert.match(viewer, /printPageUrls/, "PDF print mode must cache every page bitmap");
  assert.match(viewer, /renderPrintPages/, "PDF print mode must render pages outside the lazy viewport");
  assert.match(viewer, /data-testid="pdf-print-pages"/, "PDF print mode must expose the complete print page set");

  assert.match(surface, /minHeight\?: string \| number/, "PdfSurface must expose an optional minHeight");
  assert.match(surface, /isDark\?: boolean/, "PdfSurface must expose isDark");
  assert.match(surface, /printMode/, "PdfSurface print mode must disable lazy viewport rasterization");
  assert.match(surface, /first: 0, last: doc\.pageCount - 1/, "PdfSurface print mode must render every page");
  assert.match(surface, /@page \{ size: \$\{/, "PDF surface print page size must follow the document layout");
  assertPrintRule(surface, ".pdf-surface__page-placeholder", "PDF placeholders");
  assert.match(surface, /pdf-surface--dark/);
  assert.match(surface, /--pdf-surface-page-filter/);
  assert.match(surface, /@media print/);
  assertPrintRule(surface, ".pdf-surface__interaction", "PDF surface interaction layer");
});

test("DOCX editor follows dark mode and prints every page", () => {
  const editor = read("packages/vue-docx/src/components/DocxEditor.vue");
  const toolbar = read("packages/vue-docx/src/components/DocxToolbar.vue");
  const viewer = read("packages/vue-docx/src/components/DocxViewer.vue");
  const viewerRoot = read("packages/vue-docx/src/components/DocxViewerRoot.vue");
  const thumbnailPanel = read("packages/vue-docx/src/components/DocxThumbnailPanel.vue");
  const viewerToolbar = read("packages/vue-docx/src/components/DocxViewerToolbar.vue");
  const gutter = read("packages/vue-docx/src/components/DocxTrackedChangeGutter.vue");
  const pageSurface = read("packages/vue-docx/src/components/DocxPageSurface.vue");

  assert.match(editor, /isDark\?: boolean/, "DocxEditor must expose isDark");
  assert.match(editor, /minHeight\?: string \| number/, "DocxEditor must expose an optional minHeight");
  assert.match(editor, /isDark: false/);
  assert.match(editor, /docx-editor-viewer--dark/);
  assert.match(editor, /:theme="isDarkTheme \? 'dark' : 'light'"/, "editor must forward its theme to the document surface");
  assert.match(editor, /container: docx-shell \/ inline-size/);
  assert.match(editor, /@media print/);
  assertPrintRule(editor, ".docx-toolbar-shell", "DOCX editor toolbar");

  assert.match(toolbar, /docx-toolbar-shell--dark/, "DOCX toolbar must support a dark shell");
  assert.match(toolbar, /@container docx-shell \(max-width: 900px\)/);
  assert.match(toolbar, /@container docx-shell \(max-width: 560px\)/);
  assert.doesNotMatch(toolbar, /@media \(max-width: (?:900|560)px\)/);

  assert.match(thumbnailPanel, /docx-thumbnail-panel--dark/);
  assert.match(viewer, /minHeight\?: string \| number/, "DocxViewer must expose an optional minHeight");
  assert.doesNotMatch(viewer, /min-height: 320px/, "DocxViewer must not impose a fixed minimum height");
  assert.match(viewer, /:is-dark="darkTheme"/, "DOCX viewer thumbnails must follow the dark shell");
  assert.match(viewer, /container: docx-shell \/ inline-size/);
  assert.match(viewer, /@media print/);
  assertPrintRule(viewer, ".docx-viewer-toolbar", "DOCX viewer toolbar");

  assert.match(viewerRoot, /matchMedia\("print"\)/, "DOCX pages must react to print media");
  assert.match(viewerRoot, /addEventListener\("beforeprint"/);
  assert.match(viewerRoot, /Number\.NEGATIVE_INFINITY/, "print mode must render pages before the viewport");
  assert.match(viewerRoot, /Number\.POSITIVE_INFINITY/, "print mode must render pages after the viewport");
  assert.match(viewerRoot, /breakAfter:/, "print mode must paginate between document pages");
  assert.match(viewerRoot, /@page \{ size: \$\{Math\.round\(widthPx!\)\}px/, "DOCX print page size must follow the document layout");
  assert.match(viewerRoot, /pageWidthPx/, "DOCX @page uses the paginated page width");
  assert.match(viewerRoot, /@media print/);

  assert.match(viewerToolbar, /@container docx-shell \(max-width: 720px\)/);
  assert.match(viewerToolbar, /@container docx-shell \(max-width: 460px\)/);
  assert.doesNotMatch(viewerToolbar, /@media \(max-width: (?:720|460)px\)/);
  assert.match(viewerRoot, /gutterScreenWidthPx/, "the annotation gutter footprint must be computed by the document shell");
  assert.match(viewerRoot, /containerWidth\.value - FIT_ZOOM_SIDE_PAD - gutterScreenWidthPx\.value/, "fit-width must reserve the gutter inside the measured container");
  assert.doesNotMatch(gutter, /@media \(max-width: 1100px\)/);
  assert.match(pageSurface, /@media print/);
  assert.match(pageSurface, /background: #fff !important/, "dark document pages must print as white paper");
});

test("XLSX and PPTX viewers use container breakpoints and print rules", () => {
  const xlsxViewer = read("packages/vue-xlsx/src/components/XlsxViewer.vue");
  const xlsxToolbar = read("packages/vue-xlsx/src/components/XlsxToolbar.vue");
  const xlsxRibbon = read("packages/vue-xlsx/src/components/XlsxRibbon.vue");
  const chartsheet = read("packages/vue-xlsx/src/components/XlsxChartsheetSurface.vue");
  const pptxStyle = read("packages/vue-pptx/src/style.css");
  const pptxViewer = read("packages/vue-pptx/src/PptxViewer.vue");

  assert.match(xlsxViewer, /minHeight\?: string \| number/, "XlsxViewer must expose an optional minHeight");
  assert.match(xlsxViewer, /@page \{ size: \$\{Math\.round\(width\)\}px/, "XLSX print page size must follow the sheet's used range");
  assert.doesNotMatch(chartsheet, /min-height: (?:360|320)px/, "chartsheet must not impose fixed minimum heights");
  assert.match(chartsheet, /aspect-ratio: 16 \/ 9/, "chartsheet height must follow its container width");
  assert.match(xlsxViewer, /@media print/);
  assertPrintRule(xlsxViewer, ".xlsx-toolbar", "XLSX toolbar");
  assert.match(xlsxToolbar, /@container xlsx-viewer \(max-width: 620px\)/);
  assert.doesNotMatch(xlsxToolbar, /@media \(max-width: 620px\)/);
  assert.match(xlsxRibbon, /@container xlsx-viewer \(max-width: 560px\)/);
  assert.doesNotMatch(xlsxRibbon, /@media \(max-width: 560px\)/);

  assert.match(pptxStyle, /container: pptx-viewer \/ inline-size/);
  assert.doesNotMatch(pptxStyle, /min-height: 320px/, "PPTX viewer must not impose a fixed minimum height");
  assert.match(pptxViewer, /minHeight\?: string \| number/, "PptxViewer must expose an optional minHeight");
  assert.match(pptxViewer, /@page \{ size: \$\{/, "PPTX print page size must follow the slide dimensions");
  assert.match(pptxStyle, /@container pptx-viewer \(max-width: 720px\)/);
  assert.doesNotMatch(pptxStyle, /@media \(max-width: 720px\)/);
  assert.match(pptxStyle, /@media print/);
  assertPrintRule(pptxStyle, ".pptx-viewer__toolbar", "PPTX toolbar");
});

test("shared print stylesheet covers every viewer", () => {
  const printPath = "packages/vue-ui/src/print.css";
  assert.equal(existsSync(new URL(printPath, ROOT)), true, "shared print stylesheet must exist");
  const printStyles = read(printPath);
  const vueUiIndex = read("packages/vue-ui/src/index.ts");

  for (const selector of [
    ".pdf-viewer",
    ".pdf-surface",
    ".docx-viewer",
    ".docx-editor-viewer",
    ".xlsx-viewer",
    ".pptx-viewer",
  ]) {
    assert.ok(printStyles.includes(selector), `shared print CSS must cover ${selector}`);
  }
  assert.match(printStyles, /@media print/);
  assert.match(printStyles, /min-height: 0 !important/);
  assert.match(vueUiIndex, /import "\.\/print\.css"/, "vue-ui must ship the shared print baseline");
});

test("viewers externalize the shared vue-ui dependency", () => {
  for (const configPath of [
    "packages/vue-docx/vite.config.ts",
    "packages/vue-xlsx/vite.config.ts",
    "packages/vue-pptx/vite.config.ts",
  ]) {
    assert.match(
      read(configPath),
      /"@arcships\/vue-ui"/,
      `${configPath} must externalize @arcships/vue-ui instead of bundling a second copy`,
    );
  }
});

function createPrintPdfRuntime(pageCount) {
  const document = {
    id: "print-pdf-document",
    pageCount,
    pages: Array.from({ length: pageCount }, (_, index) => ({
      index,
      width: 600 + index,
      height: 800 + index,
      rotation: 0,
    })),
  };
  const calls = { renderPage: [] };
  const runtime = {
    runtimeId: "print-pdf-runtime",
    async openDocument() { return document; },
    async renderPage(_document, pageIndex, options) {
      calls.renderPage.push({ pageIndex, options });
      return new Blob([`page:${pageIndex}`], { type: "image/png" });
    },
    async renderThumbnail() { return new Blob(["thumb"], { type: "image/png" }); },
    async getPageTextGeometry() { return { runs: [] }; },
    async getTextSlices(_document, slices) { return slices.map(() => ""); },
    async search() { return []; },
    async closeDocument() {},
    async dispose() {},
  };
  return { calls, document, runtime };
}

test("PDF print mode rasterizes every page while the media is print", async (t) => {
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  globalThis.window = {
    matchMedia: (query) => ({
      matches: query === "print",
      media: query,
      addEventListener() {},
      removeEventListener() {},
      addListener() {},
      removeListener() {},
    }),
    addEventListener() {},
    removeEventListener() {},
    print() {},
  };
  globalThis.document = {
    createElement: () => ({ dataset: {}, remove() {}, textContent: "" }),
    head: { append() {} },
  };
  t.after(() => {
    if (originalWindow === undefined) delete globalThis.window;
    else globalThis.window = originalWindow;
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
  });

  const fake = createPrintPdfRuntime(4);
  const bytes = new TextEncoder().encode("%PDF-1.4\n%%EOF");
  const mounted = await mount(PdfViewer, {
    source: {
      kind: "blob",
      blob: new Blob([bytes], { type: "application/pdf" }),
      fileName: "print-all.pdf",
    },
    runtime: fake.runtime,
    showToolbar: false,
  });

  try {
    await waitFor(
      () => fake.calls.renderPage.filter((call) => call.options.scale === 1.5).length === 4,
      "all PDF pages rasterized for print",
    );
    assert.deepEqual(
      fake.calls.renderPage.filter((call) => call.options.scale === 1.5).map((call) => call.pageIndex).sort(),
      [0, 1, 2, 3],
    );
    assert.equal(
      walk(mounted.root).filter((node) => node.props?.["data-testid"] === "pdf-print-page").length,
      4,
      "the print snapshot contains every page, not a placeholder",
    );
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("PDF surface print mode renders every page instead of lazy placeholders", async (t) => {
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  globalThis.window = {
    matchMedia: (query) => ({
      matches: query === "print",
      media: query,
      addEventListener() {},
      removeEventListener() {},
      addListener() {},
      removeListener() {},
    }),
    addEventListener() {},
    removeEventListener() {},
    print() {},
  };
  globalThis.document = {
    createElement: () => ({ dataset: {}, remove() {}, textContent: "" }),
    head: { append() {} },
  };
  t.after(() => {
    if (originalWindow === undefined) delete globalThis.window;
    else globalThis.window = originalWindow;
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
  });

  const fake = createPrintPdfRuntime(4);
  const bytes = new TextEncoder().encode("%PDF-1.4\n%%EOF");
  const mounted = await mount(PdfSurface, {
    source: {
      kind: "blob",
      blob: new Blob([bytes], { type: "application/pdf" }),
      fileName: "print-all-surface.pdf",
    },
    runtime: fake.runtime,
  });

  try {
    await waitFor(
      () => fake.calls.renderPage.length >= 4,
      "all PDF surface pages rasterized for print",
    );
    assert.deepEqual(
      fake.calls.renderPage.slice(0, 4).map((call) => call.pageIndex).sort(),
      [0, 1, 2, 3],
    );
    assert.equal(
      walk(mounted.root).filter((node) => node.props?.class === "pdf-surface__page-img").length,
      4,
      "every page has a rasterized image",
    );
    assert.equal(
      walk(mounted.root).filter((node) => node.props?.class === "pdf-surface__page-placeholder").length,
      0,
      "print mode leaves no placeholders",
    );
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("optional minHeight is applied to the viewer root", async () => {
  const docx = await mount(DocxViewer, { minHeight: 480 });
  const docxRoot = findByTestId(docx.root, "docx-viewer");
  assert.equal(docxRoot.props.style.minHeight, "480px");
  assert.deepEqual(docx.warnings, []);
  docx.app.unmount();

  const pdf = await mount(PdfViewer, { minHeight: "40vh", showToolbar: false });
  const pdfRoot = findByTestId(pdf.root, "pdf-viewer");
  assert.equal(pdfRoot.props.style.minHeight, "40vh");
  assert.deepEqual(pdf.warnings, []);
  pdf.app.unmount();
});
