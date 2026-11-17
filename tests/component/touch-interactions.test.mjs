import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

const ROOT = path.resolve(import.meta.dirname, "../..");

function source(relativePath) {
  return readFileSync(path.join(ROOT, relativePath), "utf8");
}

const PDF_SURFACE = source("packages/vue-pdf/src/components/PdfSurface.vue");
const PDF_VIEWER = source("packages/vue-pdf/src/components/PdfViewer.vue");
const DOCX_ROOT = source("packages/vue-docx/src/components/DocxViewerRoot.vue");
const DOCX_SURFACE = source("packages/vue-docx/src/components/DocxDocumentSurface.vue");
const DOCX_TOOLBAR = source("packages/vue-docx/src/components/DocxToolbar.vue");
const DOCX_VIEWER_TOOLBAR = source("packages/vue-docx/src/components/DocxViewerToolbar.vue");
const DOCX_MENU = source("packages/vue-docx/src/components/DocxContextMenu.vue");
const PPTX_STAGE = source("packages/vue-pptx/src/PptxStage.vue");
const PPTX_STYLE = source("packages/vue-pptx/src/style.css");
const XLSX_VIEWER = source("packages/vue-xlsx/src/components/XlsxViewer.vue");
const XLSX_SURFACE = source("packages/vue-xlsx/src/components/XlsxSheetSurface.vue");
const XLSX_GRID = source("packages/vue-xlsx/src/components/XlsxGrid.vue");
const XLSX_MENU = source("packages/vue-xlsx/src/components/XlsxContextMenu.vue");
const XLSX_TOOLBAR = source("packages/vue-xlsx/src/components/XlsxToolbar.vue");
const XLSX_RIBBON = source("packages/vue-xlsx/src/components/XlsxRibbon.vue");
const XLSX_FORMULA_BAR = source("packages/vue-xlsx/src/components/XlsxFormulaBar.vue");
const XLSX_SHEET_TABS = source("packages/vue-xlsx/src/components/XlsxSheetTabs.vue");

function assertPointerLifecycle(file, label) {
  assert.match(file, /createTouchGestureTracker/, `${label} uses the shared touch tracker`);
  assert.match(file, /addEventListener\("pointerdown", onTouchPointerDown, \{ passive: true \}\)/, `${label} listens for pointerdown`);
  assert.match(file, /addEventListener\("pointermove", onTouchPointerMove, \{ passive: true \}\)/, `${label} listens for pointermove`);
  assert.match(file, /addEventListener\("pointerup", onTouchPointerUp\)/, `${label} listens for pointerup`);
  assert.match(file, /addEventListener\("pointercancel", onTouchPointerCancel\)/, `${label} listens for pointercancel`);
  assert.match(file, /removeEventListener\("pointercancel", onTouchPointerCancel\)/, `${label} removes pointercancel`);
}

test("PDF surface opens its context menu on long press and suppresses the duplicate native menu", () => {
  assertPointerLifecycle(PDF_SURFACE, "PdfSurface");
  assert.match(PDF_SURFACE, /onLongPress: \(pointer\) => emitContextMenuAt\(pointer\.clientX, pointer\.clientY\)/);
  assert.match(PDF_SURFACE, /shouldSuppressContextMenu\(event\.clientX, event\.clientY\)/);
  assert.match(PDF_SURFACE, /touch-action: pan-x pan-y/);
});

test("PDF viewer pages gain pointer pinch and double-tap zoom", () => {
  assertPointerLifecycle(PDF_VIEWER, "PdfViewer");
  assert.match(PDF_VIEWER, /getZoom: \(\) => zoom\.value/);
  assert.match(PDF_VIEWER, /onDoubleTap: \(\) => \{ zoom\.value = nextDoubleTapZoom\(zoom\.value\) \}/);
  assert.match(PDF_VIEWER, /touch-action: pan-x pan-y/);
});

test("PPTX stage long-press resolves the slide/object under the finger", () => {
  assertPointerLifecycle(PPTX_STAGE, "PptxStage");
  assert.match(PPTX_STAGE, /function resolveHitAt\(rawTarget: unknown\)/);
  assert.match(PPTX_STAGE, /const hit = resolveHitAt\(pointer\.target\)/);
  assert.match(PPTX_STAGE, /emitStageContextMenu\(hit, pointer\.clientX, pointer\.clientY\)/);
  assert.match(PPTX_STAGE, /shouldSuppressContextMenu\(event\.clientX, event\.clientY\)/);
  assert.match(PPTX_STAGE, /touch-action: pan-x pan-y/);
  assert.match(PPTX_STYLE, /\.pptx-stage \{[^}]*touch-action: pan-x pan-y/);
});

test("DOCX viewer root opens the page context menu on long press", () => {
  assertPointerLifecycle(DOCX_ROOT, "DocxViewerRoot");
  assert.match(DOCX_ROOT, /onLongPress: \(pointer\) => emitContextMenuAt\(pointer\.clientX, pointer\.clientY, pointer\.target as EventTarget \| null\)/);
  assert.match(DOCX_ROOT, /shouldSuppressContextMenu\(event\.clientX, event\.clientY\)/);
  assert.match(DOCX_ROOT, /touch-action: pan-x pan-y/);
});

test("DOCX document surface adds pointer pinch without stealing editable selection", () => {
  assertPointerLifecycle(DOCX_SURFACE, "DocxDocumentSurface");
  assert.match(DOCX_SURFACE, /nextDoubleTapZoom/);
  assert.match(DOCX_SURFACE, /isEnabled: \(\) => selectionMode\.value === "content" && !props\.editable/);
  assert.match(DOCX_SURFACE, /onZoom: \(zoom, center\) => requestGestureZoom\(zoom, center\.clientX, center\.clientY\)/);
});

test("XLSX sheet surface opens the cell menu on long press and supports pointer pinch", () => {
  assertPointerLifecycle(XLSX_SURFACE, "XlsxSheetSurface");
  assert.match(XLSX_SURFACE, /isDoubleTapEnabled: \(\) => effectiveReadOnly\.value && controlledZoom\(\) !== undefined/);
  assert.match(XLSX_SURFACE, /ref="contextMenuRef"/);
  assert.match(XLSX_SURFACE, /:suppress-context-menu="touchGestures\.shouldSuppressContextMenu"/);
  assert.match(XLSX_SURFACE, /contextMenuRef\.value\?\.openAt\(pointer\.clientX, pointer\.clientY, pointer\.target\)/);
  assert.match(XLSX_SURFACE, /shouldSuppressContextMenu\(event\.clientX, event\.clientY\)/);
  assert.match(XLSX_SURFACE, /touch-action: pan-x pan-y/);
  assert.match(XLSX_GRID, /touch-action: pan-x pan-y/);
});

test("composite XlsxViewer gains gesture zoom, long press and pointercancel recovery", () => {
  assertPointerLifecycle(XLSX_VIEWER, "XlsxViewer");
  assert.match(XLSX_VIEWER, /enableGestureZoom\?: boolean/);
  assert.match(XLSX_VIEWER, /isDoubleTapEnabled: \(\) => effectiveReadOnly\.value/);
  assert.match(XLSX_VIEWER, /function requestViewerZoom/);
  assert.match(XLSX_VIEWER, /addEventListener\("wheel", onWheel, \{ passive: false \}\)/);
  assert.match(XLSX_VIEWER, /nextSurfaceZoom\(pendingZoom \?\? currentZoom\(\), event\.deltaY, event\.deltaMode\)/);
  assert.match(XLSX_VIEWER, /if \(!event\.ctrlKey && !event\.metaKey\) return;/);
  assert.match(XLSX_VIEWER, /captureZoomAnchor\(clientX, clientY\)/);
  assert.match(XLSX_VIEWER, /restoreZoomAnchor\(restore\.anchor\)/);
  assert.match(XLSX_VIEWER, /contextMenuRef\.value\?\.openAt\(pointer\.clientX, pointer\.clientY, pointer\.target\)/);
  assert.match(XLSX_VIEWER, /:suppress-context-menu="touchGestures\.shouldSuppressContextMenu"/);
  assert.match(XLSX_VIEWER, /touch-action: pan-x pan-y/);
});

test("context menus close on outside pointer down, not on the trailing click of a long press", () => {
  assert.match(XLSX_MENU, /class="xlsx-contextmenu-backdrop"[\s\S]*?@pointerdown="close"/);
  assert.match(DOCX_MENU, /class="docx-context-menu-backdrop"[\s\S]*?@pointerdown="close"/);
  assert.match(XLSX_MENU, /defineExpose\(\{ openAt, close \}\)/);
});

test("touch hit targets reach 44px on coarse pointers", () => {
  const targets = [
    [XLSX_TOOLBAR, "xlsx toolbar"],
    [XLSX_RIBBON, "xlsx ribbon"],
    [XLSX_FORMULA_BAR, "xlsx formula bar"],
    [XLSX_SHEET_TABS, "xlsx sheet tabs"],
    [XLSX_MENU, "xlsx context menu"],
    [DOCX_TOOLBAR, "docx editor toolbar"],
    [DOCX_VIEWER_TOOLBAR, "docx viewer toolbar"],
    [DOCX_MENU, "docx context menu"],
    [PDF_VIEWER, "pdf viewer toolbar"],
    [PPTX_STYLE, "pptx viewer toolbar"],
  ];
  for (const [file, label] of targets) {
    assert.match(file, /@media \(pointer: coarse\)/, `${label} has a coarse-pointer rule`);
    assert.match(file, /44px/, `${label} sizes touch targets to 44px`);
  }
});
