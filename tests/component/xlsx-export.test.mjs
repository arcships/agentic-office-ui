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
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { useXlsxViewerController, XlsxToolbar } = await importFromDemo("@arcships/vue-xlsx");

function sampleBytes(name) {
  const bytes = readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  );
  return new Uint8Array(bytes);
}

function sampleBuffer(name) {
  const bytes = sampleBytes(name);
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

function installDownloadCapture(t) {
  const downloads = [];
  const blobs = new Map();
  const originalDocument = globalThis.document;
  const originalCreateObjectUrl = URL.createObjectURL;
  const originalRevokeObjectUrl = URL.revokeObjectURL;
  let nextObjectUrl = 0;

  globalThis.document = {
    body: { append() {} },
    createElement(tagName) {
      return {
        tagName,
        href: "",
        download: "",
        click() {
          downloads.push({
            blob: blobs.get(this.href),
            fileName: this.download,
          });
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

function patchWorkbookPart(bytes, partName, replace) {
  const archive = unzipSync(bytes);
  archive[partName] = strToU8(replace(strFromU8(archive[partName])));
  return zipSync(archive);
}

test("XLSX export waits for a loaded source and ignores loading state", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController({
    deferLoadingAboveBytes: 1,
    file: sampleBuffer("charts-images.xlsx"),
    fileName: "charts-images.xlsx",
    readOnly: true,
    skipXmlParsing: true,
  });
  t.after(() => app.unmount());

  await waitFor(() => controller.isLoadDeferred, "deferred XLSX load");
  assert.equal(controller.workbook, null);
  assert.equal(controller.canExport, false);

  controller.continueDeferredLoad();
  assert.equal(controller.sourceState, "loading");
  assert.equal(controller.canExport, false);
  assert.equal(controller.workbook, null);
  controller.exportXlsx();
  assert.equal(downloads.length, 0, "export must be ignored until the load commits");

  await waitForReady(controller);
  assert.notEqual(controller.workbook, null);
  assert.equal(controller.canExport, true);

  controller.exportXlsx();
  await vue.nextTick();
  assert.equal(downloads.length, 1);
  assert.match(downloads[0].fileName, /\.xlsx$/);
});

test("XLSX export warns when an xlsm/xltm source loses its macros", async (t) => {
  const diagnostics = [];
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController({
    file: sampleBuffer("charts-images.xlsx"),
    fileName: "macro-book.xlsm",
    onDiagnostic: (diagnostic) => diagnostics.push(diagnostic),
    skipXmlParsing: true,
  });
  t.after(() => app.unmount());

  await waitForReady(controller);
  assert.equal(controller.lastExportWarning, null);

  controller.exportXlsx();
  await vue.nextTick();

  assert.match(controller.lastExportWarning ?? "", /VBA|宏/);
  assert.equal(downloads.length, 1);
  assert.equal(downloads[0].fileName, "macro-book.xlsx");
  const warning = diagnostics.find((diagnostic) => diagnostic.type === "export-warning");
  assert.ok(warning, "macro loss must be reported through diagnostics");
  assert.match(warning.message ?? "", /VBA|宏/);

  const toolbar = await mount(XlsxToolbar, { controller });
  t.after(() => toolbar.app.unmount());
  const warningNode = findByTestId(toolbar.root, "xlsx-export-warning");
  assert.ok(warningNode, "the toolbar must render the export warning");
  assert.match(textContent(warningNode), /VBA|宏/);
});

test("XLSX export sanitizes double-escaped entities in every XML part", async (t) => {
  const downloads = installDownloadCapture(t);
  const patched = patchWorkbookPart(
    patchWorkbookPart(
      sampleBytes("charts-images.xlsx"),
      "xl/worksheets/sheet1.xml",
      (xml) => xml.replace(
        "<t>Quarter</t>",
        "<t>Cell &amp;quot;value&amp;quot; and &amp;lt;tag&amp;gt;</t>",
      ),
    ),
    "xl/styles.xml",
    (xml) => xml.replace(
      'val="Calibri"',
      'val="Calibri &amp;quot;Style&amp;quot;"',
    ),
  );
  const { app, controller } = await mountController({
    file: patched.buffer.slice(patched.byteOffset, patched.byteOffset + patched.byteLength),
    fileName: "double-escaped.xlsx",
    skipXmlParsing: true,
  });
  t.after(() => app.unmount());

  await waitForReady(controller);
  controller.exportXlsx();
  await vue.nextTick();
  assert.equal(downloads.length, 1);

  const archive = unzipSync(new Uint8Array(await downloads[0].blob.arrayBuffer()));
  const sharedStrings = strFromU8(archive["xl/sharedStrings.xml"]);
  const styles = strFromU8(archive["xl/styles.xml"]);
  assert.match(sharedStrings, /Cell &quot;value&quot; and &lt;tag&gt;/);
  assert.doesNotMatch(
    sharedStrings,
    /&amp;(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/,
  );
  assert.match(styles, /Calibri &quot;Style&quot;/);
  assert.doesNotMatch(
    styles,
    /&amp;(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/,
  );
});

test("XLSX CSV export uses the active sheet instead of workbook.activeSheet", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
    skipXmlParsing: true,
  });
  t.after(() => app.unmount());

  await waitForReady(controller);
  controller.setActiveTabIndex(1);
  await vue.nextTick();
  assert.equal(controller.activeSheet?.name, "P&L");

  controller.exportCsv();
  await vue.nextTick();
  assert.equal(downloads.length, 1);
  assert.equal(downloads[0].fileName, "financial-model-P&L.csv");

  const csv = await downloads[0].blob.text();
  assert.match(csv, /^Line Item,/);
  assert.doesNotMatch(csv, /^Driver,/);
  assert.match(csv, /\r\n/);

  controller.setActiveTabIndex(0);
  await vue.nextTick();
  assert.ok(controller.workbook);
  const wasmCsv = controller.workbook.saveCsvString();
  controller.exportCsv();
  await vue.nextTick();
  assert.equal(downloads.length, 2);
  assert.equal(
    await downloads[1].blob.text(),
    wasmCsv,
    "the active-sheet writer must match the WASM CSV output for the active sheet",
  );
});

test("XLSX addSheet rejects Excel-invalid worksheet names", async (t) => {
  const { app, controller } = await mountController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
    skipXmlParsing: true,
  });
  t.after(() => app.unmount());

  await waitForReady(controller);
  const initialSheetCount = controller.sheets.length;
  for (const name of ["Bad/Name", "Bad:Name", "Bad*Name", "Bad?Name", "Bad[Name]", "Bad\\Name"]) {
    assert.throws(() => controller.addSheet(name), /Worksheet names/);
  }
  assert.equal(controller.sheets.length, initialSheetCount);
  assert.equal(controller.canUndo, false);

  controller.addSheet("Valid Name");
  await vue.nextTick();
  assert.equal(controller.sheets.length, initialSheetCount + 1);
  assert.equal(controller.sheets.at(-1)?.name, "Valid Name");
});
