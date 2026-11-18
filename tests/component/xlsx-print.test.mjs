import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import test from "node:test";
import { pathToFileURL } from "node:url";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const {
  XlsxGrid,
  XlsxPrintPreview,
  XlsxToolbar,
  useXlsxViewerController,
} = await importFromDemo("@arcships/vue-xlsx");

const ROOT = path.resolve(import.meta.dirname, "../..");

function read(relativePath) {
  return readFileSync(path.join(ROOT, relativePath), "utf8");
}

function sampleBytes(name = "financial-model.xlsx") {
  return new Uint8Array(
    readFileSync(new URL(`../../apps/demo/public/samples/${name}`, import.meta.url)),
  );
}

function patchPrintMetadata(bytes = sampleBytes()) {
  const archive = unzipSync(bytes);
  let sheet = strFromU8(archive["xl/worksheets/sheet1.xml"]);
  sheet = sheet.replace(
    /<pageMargins[^>]*\/>/,
    '<pageMargins left="0.5" right="0.6" top="0.7" bottom="0.8" header="0.2" footer="0.3"/>',
  );
  sheet = sheet.replace(
    /<sheetPr([^>]*)>/,
    '<sheetPr$1><pageSetUpPr fitToPage="1"/>',
  );
  sheet = sheet.replace(
    "</worksheet>",
    [
      '<printOptions gridLines="1" headings="1" horizontalCentered="1"/>',
      '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="1" scale="85"/>',
      "</worksheet>",
    ].join(""),
  );
  archive["xl/worksheets/sheet1.xml"] = strToU8(sheet);
  let workbook = strFromU8(archive["xl/workbook.xml"]);
  workbook = workbook.replace(
    "</sheets>",
    "</sheets><definedNames><definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">Assumptions!$A$1:$D$20</definedName></definedNames>",
  );
  archive["xl/workbook.xml"] = strToU8(workbook);
  const patched = zipSync(archive);
  return patched.buffer.slice(patched.byteOffset, patched.byteOffset + patched.byteLength);
}

let wasmConfigured = false;
function configureWasm() {
  if (wasmConfigured) return;
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  setWasmSource(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  );
  wasmConfigured = true;
}
configureWasm();

async function mountController(options) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        useWorker: false,
        ...options,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, controller };
}

async function waitForReady(controller) {
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "XLSX controller ready state",
  );
  assert.ifError(controller.error);
  assert.equal(controller.sourceState, "ready");
}

test("controller exposes print layout, page-break preview and print actions", async (t) => {
  const { app, controller } = await mountController({
    file: patchPrintMetadata(),
    fileName: "print-layout.xlsx",
  });
  t.after(() => app.unmount());
  await waitForReady(controller);

  const layout = controller.printLayout;
  assert.ok(layout, "the active sheet must expose a print layout");
  assert.equal(layout.sheetName, "Assumptions");
  assert.equal(layout.paperSize.name, "A4");
  assert.equal(layout.pageSetup.orientation, "landscape");
  assert.equal(layout.pageSetup.fitToPage, true);
  assert.equal(layout.printAreas.length, 1);
  assert.deepEqual(layout.printAreas[0], {
    end: { col: 3, row: 19 },
    start: { col: 0, row: 0 },
  });
  assert.ok(layout.pages.length >= 1);
  assert.deepEqual(controller.printPages, layout.pages);
  assert.deepEqual(controller.getPrintLayout(), layout);

  assert.equal(controller.isPageBreakPreview, false);
  controller.setPageBreakPreview(true);
  await vue.nextTick();
  assert.equal(controller.isPageBreakPreview, true);
  controller.setPageBreakPreview(false);
  assert.equal(controller.isPageBreakPreview, false);

  assert.equal(controller.isPrintPreviewOpen, false);
  assert.equal(controller.openPrintPreview(), true);
  assert.equal(controller.isPrintPreviewOpen, true);
  controller.closePrintPreview();
  assert.equal(controller.isPrintPreviewOpen, false);

  const printed = [];
  const previousWindow = globalThis.window;
  globalThis.window = { print: () => printed.push("print") };
  t.after(() => {
    if (previousWindow === undefined) delete globalThis.window;
    else globalThis.window = previousWindow;
  });
  assert.equal(controller.print(), true);
  await vue.nextTick();
  await vue.nextTick();
  assert.deepEqual(printed, ["print"], "print() opens the preview and calls window.print");
  assert.equal(controller.isPrintPreviewOpen, true);
});

test("print preview renders one DOM block per computed page", async (t) => {
  const { app: controllerApp, controller } = await mountController({
    file: patchPrintMetadata(),
    fileName: "print-layout.xlsx",
  });
  t.after(() => controllerApp.unmount());
  await waitForReady(controller);
  controller.openPrintPreview();

  const { app, root } = await mount(XlsxPrintPreview, { controller });
  t.after(() => app.unmount());

  assert.ok(findByTestId(root, "xlsx-print-preview"), "print preview surface renders");
  const pagesRoot = findByTestId(root, "xlsx-print-pages");
  assert.ok(pagesRoot, "print preview exposes the page list");
  const pageNodes = walk(root).filter((node) => node.props?.["data-testid"] === "xlsx-print-page");
  assert.equal(pageNodes.length, controller.printPages.length);
  assert.equal(pageNodes[0].props["data-page-paper"], "A4");
  assert.ok(pageNodes[0].props["data-page-range"]);
  const printedCells = walk(root).filter((node) => node.type === "td");
  assert.ok(printedCells.length > 0, "printed pages render cell blocks");
  assert.ok(
    printedCells.some((cell) => textContent(cell).trim().length > 0),
    "printed pages render cell values",
  );
  assert.match(textContent(findByTestId(root, "xlsx-print-summary")), /A4 landscape/);
  assert.ok(findByTestId(root, "xlsx-print-print"), "print preview has a print button");
  assert.ok(findByTestId(root, "xlsx-print-close"), "print preview has a close button");
  const columnHeading = walk(root).find(
    (node) => String(node.props?.class).includes("xlsx-print-page__column-heading"),
  );
  assert.ok(columnHeading, "printOptions.headings renders column headings");
  assert.equal(textContent(columnHeading), "A");

  const printed = [];
  const previousWindow = globalThis.window;
  globalThis.window = { print: () => printed.push("print") };
  t.after(() => {
    if (previousWindow === undefined) delete globalThis.window;
    else globalThis.window = previousWindow;
  });
  findByTestId(root, "xlsx-print-print").props.onClick();
  await vue.nextTick();
  await vue.nextTick();
  assert.deepEqual(printed, ["print"]);
});

test("page-break preview renders the dashed overlay in the grid", async (t) => {
  const { app: controllerApp, controller } = await mountController({
    file: patchPrintMetadata(),
    fileName: "print-layout.xlsx",
  });
  t.after(() => controllerApp.unmount());
  await waitForReady(controller);
  controller.setPageBreakPreview(true);

  const previousWindow = globalThis.window;
  const previousResizeObserver = globalThis.ResizeObserver;
  const previousGetContext = Object.getOwnPropertyDescriptor(Object.prototype, "getContext");
  globalThis.window = {
    cancelAnimationFrame() {},
    devicePixelRatio: 1,
    requestAnimationFrame: () => 1,
  };
  globalThis.ResizeObserver = class ResizeObserver {
    disconnect() {}
    observe() {}
  };
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value: () => null,
  });

  const { app, root } = await mount(XlsxGrid, { controller });
  t.after(() => {
    app.unmount();
    if (previousWindow === undefined) delete globalThis.window;
    else globalThis.window = previousWindow;
    if (previousResizeObserver === undefined) delete globalThis.ResizeObserver;
    else globalThis.ResizeObserver = previousResizeObserver;
    if (previousGetContext) Object.defineProperty(Object.prototype, "getContext", previousGetContext);
    else delete Object.prototype.getContext;
  });
  await vue.nextTick();
  assert.ok(findByTestId(root, "xlsx-page-break-preview"), "page-break overlay mounts in preview mode");
});

test("toolbar exposes page-break preview and print-preview controls", async (t) => {
  const { app: controllerApp, controller } = await mountController({
    file: patchPrintMetadata(),
    fileName: "print-layout.xlsx",
  });
  t.after(() => controllerApp.unmount());
  await waitForReady(controller);

  const { app, root } = await mount(XlsxToolbar, { controller });
  t.after(() => app.unmount());
  const pageBreakButton = findByTestId(root, "xlsx-toolbar-page-breaks");
  assert.ok(pageBreakButton, "toolbar must expose the page-break toggle");
  pageBreakButton.props.onClick();
  await vue.nextTick();
  assert.equal(controller.isPageBreakPreview, true);
  assert.equal(pageBreakButton.props["aria-pressed"], true);

  const previewButton = findByTestId(root, "xlsx-toolbar-print-preview");
  assert.ok(previewButton, "toolbar must expose print preview");
  previewButton.props.onClick();
  await vue.nextTick();
  assert.equal(controller.isPrintPreviewOpen, true);
});

test("viewer, grid and print preview ship print CSS and page-break overlay hooks", () => {
  const viewer = read("packages/vue-xlsx/src/components/XlsxViewer.vue");
  const preview = read("packages/vue-xlsx/src/components/XlsxPrintPreview.vue");
  const grid = read("packages/vue-xlsx/src/components/XlsxGrid.vue");
  const toolbar = read("packages/vue-xlsx/src/components/XlsxToolbar.vue");

  assert.match(viewer, /<XlsxPrintPreview/);
  assert.match(viewer, /@media print/);
  assert.match(viewer, /@page \{/);
  assert.match(viewer, /\.xlsx-viewer--print-preview \.xlsx-viewer__body > :not\(\.xlsx-print-preview\) \{ display: none !important; \}/);
  assert.match(viewer, /xlsx-viewer--print-preview/);

  assert.match(preview, /data-testid="xlsx-print-pages"/);
  assert.match(preview, /data-testid="xlsx-print-page"/);
  assert.match(preview, /buildPrintPageGrid/);
  assert.match(preview, /@media print/);
  assert.match(preview, /break-after: page/);
  assert.match(preview, /page-break-after: always/);
  assert.match(preview, /transform: none !important/);

  assert.match(grid, /data-testid="xlsx-page-break-preview"/);
  assert.match(grid, /pageBreakOverlay/);
  assert.match(grid, /xlsx-grid__page-break-line--column/);
  assert.match(grid, /xlsx-grid__page-break-line--row/);
  assert.match(grid, /page\.range\.start/);

  assert.match(toolbar, /data-testid="xlsx-toolbar-page-breaks"/);
  assert.match(toolbar, /data-testid="xlsx-toolbar-print-preview"/);
});
