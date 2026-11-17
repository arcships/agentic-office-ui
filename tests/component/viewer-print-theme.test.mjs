import assert from "node:assert/strict";
import { readFileSync, existsSync } from "node:fs";
import test from "node:test";

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

  assert.match(types, /isDark\?: boolean/, "PdfViewerProps must expose isDark");
  assert.match(viewer, /isDark: false/);
  assert.match(viewer, /pdf-viewer--dark/, "PdfViewer must render the dark root class");
  assert.match(viewer, /--pdf-page-filter: var\(--pdf-dark-page-filter, invert\(1\) hue-rotate\(180deg\)\)/, "dark PDF pages must be invertible but overridable");
  assert.doesNotMatch(viewer, /min-height: (?:620|540|220|180)px/, "viewer chrome must not use fixed min-heights that overflow short containers");
  assert.match(viewer, /container: pdf-viewer \/ inline-size/);
  assert.match(viewer, /@container pdf-viewer \(max-width: 760px\)/, "PDF breakpoints must follow the container, not the viewport");
  assert.doesNotMatch(viewer, /@media \(max-width: 760px\)/);
  assert.match(viewer, /@media print/);
  assertPrintRule(viewer, ".pdf-toolbar", "PDF toolbar");

  assert.match(surface, /isDark\?: boolean/, "PdfSurface must expose isDark");
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
  assert.match(viewer, /:is-dark="darkTheme"/, "DOCX viewer thumbnails must follow the dark shell");
  assert.match(viewer, /container: docx-shell \/ inline-size/);
  assert.match(viewer, /@media print/);
  assertPrintRule(viewer, ".docx-viewer-toolbar", "DOCX viewer toolbar");

  assert.match(viewerRoot, /matchMedia\("print"\)/, "DOCX pages must react to print media");
  assert.match(viewerRoot, /addEventListener\("beforeprint"/);
  assert.match(viewerRoot, /Number\.NEGATIVE_INFINITY/, "print mode must render pages before the viewport");
  assert.match(viewerRoot, /Number\.POSITIVE_INFINITY/, "print mode must render pages after the viewport");
  assert.match(viewerRoot, /breakAfter:/, "print mode must paginate between document pages");
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
  const pptxStyle = read("packages/vue-pptx/src/style.css");

  assert.match(xlsxViewer, /container: xlsx-viewer \/ inline-size/);
  assert.match(xlsxViewer, /@media print/);
  assertPrintRule(xlsxViewer, ".xlsx-toolbar", "XLSX toolbar");
  assert.match(xlsxToolbar, /@container xlsx-viewer \(max-width: 620px\)/);
  assert.doesNotMatch(xlsxToolbar, /@media \(max-width: 620px\)/);
  assert.match(xlsxRibbon, /@container xlsx-viewer \(max-width: 560px\)/);
  assert.doesNotMatch(xlsxRibbon, /@media \(max-width: 560px\)/);

  assert.match(pptxStyle, /container: pptx-viewer \/ inline-size/);
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
