// Issue #171 — XLSX page setup authoring end to end.
//
// The dialog test (`xlsx-page-setup-dialog.test.mjs`) drives the form through
// happy-dom. This file drives the thing the form writes: the controller API,
// the undo entry, the exported bytes and the print preview the user sees.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const { strFromU8, unzipSync } = await import(
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
  XlsxPrintPreview,
  useXlsxViewerController,
} = await importFromDemo("@arcships/vue-xlsx");

function sampleBuffer(name = "financial-model.xlsx") {
  const bytes = readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
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
      controller = useXlsxViewerController({ useWorker: false, ...options });
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

function installDownloadCapture(t) {
  const downloads = [];
  const blobs = new Map();
  const originalDocument = globalThis.document;
  const originalCreateObjectUrl = URL.createObjectURL;
  const originalRevokeObjectUrl = URL.revokeObjectURL;
  let nextObjectUrl = 0;

  globalThis.document = {
    body: { append() {} },
    createElement() {
      return {
        href: "",
        download: "",
        click() {
          downloads.push({ blob: blobs.get(this.href), fileName: this.download });
        },
        remove() {},
      };
    },
  };
  URL.createObjectURL = (blob) => {
    const objectUrl = `blob:test-${nextObjectUrl += 1}`;
    blobs.set(objectUrl, blob);
    return objectUrl;
  };
  URL.revokeObjectURL = () => {};

  t.after(() => {
    if (originalDocument === undefined) delete globalThis.document;
    else globalThis.document = originalDocument;
    if (originalCreateObjectUrl === undefined) delete URL.createObjectURL;
    else URL.createObjectURL = originalCreateObjectUrl;
    if (originalRevokeObjectUrl === undefined) delete URL.revokeObjectURL;
    else URL.revokeObjectURL = originalRevokeObjectUrl;
  });

  return downloads;
}

async function exportedArchive(downloads) {
  await waitFor(() => downloads.length > 0, "exported workbook download");
  const bytes = new Uint8Array(await downloads.at(-1).blob.arrayBuffer());
  return unzipSync(bytes);
}

test("setPageSetup writes the sheet, one undo step and the exported bytes", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController({
    file: sampleBuffer(),
    fileName: "financial-model.xlsx",
  });
  t.after(() => app.unmount());
  await waitForReady(controller);

  const before = { ...controller.activeSheet.pageSetup };
  const beforeMargins = { ...controller.activeSheet.pageMargins };
  assert.equal(before.orientation, "portrait", "the sample starts as portrait Letter");

  assert.equal(
    controller.setPageSetup({
      pageMargins: { left: 0.5, right: 0.6, top: 1.25 },
      pageSetup: { orientation: "landscape", paperSize: 9, scale: 120 },
      printAreas: [{ end: { col: 3, row: 19 }, start: { col: 0, row: 0 } }],
      printOptions: { gridLines: true, headings: true },
    }),
    true,
  );
  await vue.nextTick();

  const sheet = controller.activeSheet;
  assert.equal(sheet.pageSetup.orientation, "landscape");
  assert.equal(sheet.pageSetup.paperSize, 9);
  assert.equal(sheet.pageSetup.scale, 120);
  assert.equal(sheet.pageMargins.top, 1.25);
  assert.equal(sheet.pageMargins.left, 0.5);
  assert.equal(
    sheet.pageMargins.bottom,
    beforeMargins.bottom,
    "untouched margins keep the loaded value",
  );
  assert.equal(sheet.printOptions.gridLines, true);
  assert.equal(sheet.printOptions.headings, true);
  assert.deepEqual(sheet.printAreas, [{ end: { col: 3, row: 19 }, start: { col: 0, row: 0 } }]);

  // One command, one undo entry.
  assert.equal(controller.canUndo, true);
  assert.equal(
    controller.setPageSetup({ pageSetup: { orientation: "landscape" } }),
    false,
    "an unchanged patch is not a history entry",
  );

  // The export carries the authored page setup and the print area definition.
  controller.exportXlsx();
  const archive = await exportedArchive(downloads);
  const worksheet = strFromU8(archive["xl/worksheets/sheet1.xml"]);
  assert.match(worksheet, /<pageSetup[^>]*orientation="landscape"/);
  assert.match(worksheet, /<pageSetup[^>]*paperSize="9"/);
  assert.match(worksheet, /<pageSetup[^>]*scale="120"/);
  assert.match(worksheet, /<pageMargins[^>]*left="0.5"/);
  assert.match(worksheet, /<pageMargins[^>]*top="1.25"/);
  assert.match(worksheet, /<printOptions[^>]*gridLines="1"/);
  assert.match(worksheet, /<printOptions[^>]*headings="1"/);
  const workbook = strFromU8(archive["xl/workbook.xml"]);
  assert.match(
    workbook,
    /<definedName name="_xlnm.Print_Area" localSheetId="0">'?Assumptions'?!\$A\$1:\$D\$20<\/definedName>/,
    "the print area is an absolute, sheet-scoped name",
  );

  // Undo returns the worksheet to what it was before the dialog.
  controller.undo();
  await waitFor(() => controller.activeSheet?.pageSetup.orientation === "portrait", "undo restore");
  assert.equal(controller.activeSheet.pageSetup.orientation, "portrait");
  assert.equal(controller.activeSheet.pageSetup.paperSize, before.paperSize);
  assert.equal(controller.activeSheet.pageMargins.top, beforeMargins.top);
  assert.deepEqual(controller.activeSheet.printAreas, []);
});

test("fit-to-page authoring keeps an automatic height and clears the print area", async (t) => {
  const { app, controller } = await mountController({
    file: sampleBuffer(),
    fileName: "financial-model.xlsx",
  });
  t.after(() => app.unmount());
  await waitForReady(controller);

  assert.equal(
    controller.setPageSetup({
      pageSetup: { fitToHeight: 0, fitToPage: true, fitToWidth: 1 },
    }),
    true,
  );
  await vue.nextTick();
  assert.equal(controller.activeSheet.pageSetup.fitToPage, true);
  assert.equal(controller.activeSheet.pageSetup.fitToWidth, 1);
  assert.equal(controller.activeSheet.pageSetup.fitToHeight, 0);
  assert.equal(controller.printLayout.fitToPage, true);
  assert.ok(controller.printLayout.pages.length >= 1);

  // Setting a print area and clearing it again removes the definition.
  controller.setPageSetup({ printAreas: [{ end: { col: 2, row: 9 }, start: { col: 0, row: 0 } }] });
  await vue.nextTick();
  assert.equal(controller.activeSheet.printAreas.length, 1);
  assert.equal(controller.setPageSetup({ printAreas: [] }), true);
  await vue.nextTick();
  assert.deepEqual(controller.activeSheet.printAreas, []);
});

test("the print preview renders the authored paper and orientation", async (t) => {
  const { app: controllerApp, controller } = await mountController({
    file: sampleBuffer(),
    fileName: "financial-model.xlsx",
  });
  t.after(() => controllerApp.unmount());
  await waitForReady(controller);

  const portraitWidth = controller.printLayout.pages[0].widthPx;
  const portraitHeight = controller.printLayout.pages[0].heightPx;
  assert.ok(portraitWidth < portraitHeight, "the sample starts portrait");

  assert.equal(
    controller.setPageSetup({
      pageSetup: { orientation: "landscape", paperSize: 9 },
    }),
    true,
  );
  await vue.nextTick();

  const page = controller.printLayout.pages[0];
  assert.equal(page.paperSizeName, "A4");
  assert.ok(page.widthPx > page.heightPx, "landscape swaps the paper axes");
  assert.ok(page.widthPx > portraitWidth, "the A4 long edge is wider than the Letter short edge");

  const { app, root } = await mount(XlsxPrintPreview, { controller });
  t.after(() => app.unmount());
  assert.ok(findByTestId(root, "xlsx-print-preview"), "print preview renders");
  assert.match(textContent(findByTestId(root, "xlsx-print-summary")), /A4 landscape/);
});

test("page setup authoring is a no-op on a read-only sheet", async (t) => {
  const { app, controller } = await mountController({
    file: sampleBuffer(),
    fileName: "financial-model.xlsx",
    readOnly: true,
  });
  t.after(() => app.unmount());
  await waitForReady(controller);

  assert.equal(controller.readOnly, true);
  assert.equal(
    controller.setPageSetup({ pageSetup: { orientation: "landscape" } }),
    false,
    "a read-only viewer refuses the mutation",
  );
  assert.equal(controller.activeSheet.pageSetup.orientation, "portrait");
});
