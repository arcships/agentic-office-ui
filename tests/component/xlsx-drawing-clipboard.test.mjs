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
  teleportTarget,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
const { strFromU8, strToU8, unzipSync, zipSync } = requireFromXlsxCore("fflate");
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;
globalThis.Document ??= class Document {};
globalThis.ShadowRoot ??= class ShadowRoot {};

const core = await importFromDemo("@arcships/xlsx-core");
const {
  useXlsxViewerController,
  XlsxContextMenu,
} = await importFromDemo("@arcships/vue-xlsx");

const ROOT = path.resolve(import.meta.dirname, "../..");

function configureWasm() {
  const bytes = readFileSync(
    new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
  );
  core.setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}
configureWasm();

/**
 * charts-images has one worksheet plus a chartsheet. Add a plain second
 * worksheet so the clipboard tests can paste across real sheet tabs.
 */
function twoSheetBuffer() {
  const bytes = new Uint8Array(readFileSync(
    new URL("../../apps/demo/public/samples/charts-images.xlsx", import.meta.url),
  ));
  const archive = unzipSync(bytes);
  archive["xl/workbook.xml"] = strToU8(
    strFromU8(archive["xl/workbook.xml"]).replace(
      "</sheets>",
      '<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Target" sheetId="3" r:id="rIdTarget"/></sheets>',
    ),
  );
  archive["xl/_rels/workbook.xml.rels"] = strToU8(
    strFromU8(archive["xl/_rels/workbook.xml.rels"]).replace(
      "</Relationships>",
      '<Relationship Id="rIdTarget" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
    ),
  );
  archive["xl/worksheets/sheet2.xml"] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
  );
  const zipped = zipSync(archive, { level: 6 });
  return zipped.buffer.slice(zipped.byteOffset, zipped.byteOffset + zipped.byteLength);
}

async function mountController() {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: twoSheetBuffer(),
        fileName: "charts-images.xlsx",
        useWorker: false,
      });
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    "XLSX drawing clipboard controller ready",
  );
  assert.ifError(controller.error);
  await waitFor(
    () => controller.images.length === 1
      && controller.shapes.length === 1
      && controller.charts.length === 1,
    "XLSX drawings loaded",
  );
  return { ...mounted, controller };
}

async function onTargetSheet(controller) {
  await goToSheet(controller, 1, { row: 2, col: 3 });
}

async function goToSheet(controller, index, cell) {
  controller.setActiveSheetIndex(index);
  await vue.nextTick();
  if (cell) controller.selectCell(cell);
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

test("an image copied on one sheet pastes on another and undo/redo restores it", async (t) => {
  const { app, controller } = await mountController();
  t.after(() => app.unmount());
  const image = controller.images[0];
  controller.selectImage(image.id);

  const copied = controller.copyDrawing();
  assert.ok(copied, "copyDrawing returns the internal clipboard payload");
  const payload = JSON.parse(copied.structured);
  assert.equal(payload.drawings.length, 1);
  assert.equal(payload.drawings[0].kind, "image");
  assert.equal(payload.drawings[0].sourceId, image.id);
  assert.equal(payload.drawings[0].sourceSheetId, 0);
  assert.ok(payload.drawings[0].size.cx > 0 && payload.drawings[0].size.cy > 0);
  assert.deepEqual(payload.drawings[0].relativeAnchor, { x: 0, y: 0 });

  await onTargetSheet(controller);
  assert.equal(await controller.pasteFromClipboard(), true, "paste lands on the target sheet");
  assert.equal(controller.images.length, 1);
  assert.deepEqual(controller.images[0].anchor.from, { col: 3, colOffsetEmu: 0, row: 2, rowOffsetEmu: 0 });
  assert.notEqual(controller.images[0].id, image.id, "the pasted picture gets a new id");
  assert.equal(controller.selectedImageId, controller.images[0].id);

  controller.undo();
  await waitFor(() => controller.images.length === 0 && controller.canRedo, "undo removes the pasted picture");
  controller.redo();
  await waitFor(() => controller.images.length === 1, "redo restores the pasted picture");
});

test("an image and a chart copied on one sheet paste together and export", async (t) => {
  const { app, controller } = await mountController();
  t.after(() => app.unmount());
  const image = controller.images[0];
  controller.selectImage(image.id);
  assert.ok(controller.copyDrawing());
  await onTargetSheet(controller);
  assert.equal(await controller.pasteDrawing(), true);
  assert.equal(controller.images.length, 1, "the copied picture lands on the target sheet");

  await goToSheet(controller, 0, { row: 0, col: 0 });
  const chart = controller.charts[0];
  controller.selectChart(chart.id);
  assert.ok(controller.copyDrawing());

  await goToSheet(controller, 1, { row: 4, col: 5 });
  assert.equal(await controller.pasteDrawing(), true);
  assert.equal(controller.charts.length, 1);
  const pasted = controller.charts[0];
  assert.notEqual(pasted.id, chart.id);
  assert.notEqual(pasted.chartPath, chart.chartPath);
  assert.match(pasted.chartPath, /chartClipboard\d+\.xml$/);
  assert.equal(pasted.series.length, chart.series.length);
  assert.deepEqual(
    pasted.series.map((series) => series.id),
    pasted.series.map((_, index) => `${pasted.id}-series-${index}`),
  );

  const downloads = installDownloadCapture(t);
  controller.exportXlsx();
  assert.equal(downloads.length, 1, "export produced one workbook");
  const exported = new Uint8Array(await downloads[0].blob.arrayBuffer());
  const archive = unzipSync(exported);
  assert.ok(archive[pasted.chartPath], "the exported workbook includes the cloned chart part");
  const reparsed = core.parseWorkbookImageAssets(exported);
  assert.equal(reparsed.imagesByWorkbookSheetIndex[1].length, 1);
  assert.equal(reparsed.shapesByWorkbookSheetIndex[1].length, 0);
  const exportedWorkbook = core.getSheetsWasmModule
    ? (await core.getSheetsWasmModule()).Workbook.fromBytes(exported)
    : null;
  try {
    assert.ok(exportedWorkbook, "the exported workbook reloads");
    const reparsedCharts = core.loadWorkbookChartAssets(
      exportedWorkbook,
      reparsed,
      new Map([[0, 0], [1, 1]]),
      false,
    );
    assert.equal(reparsedCharts.chartsByWorkbookSheetIndex[1].length, 1, "the pasted chart exports");
    assert.equal(reparsedCharts.chartsByWorkbookSheetIndex[1][0].id.startsWith("chart-1-"), true);
  } finally {
    exportedWorkbook?.free();
  }
});

test("cut-paste moves a shape to another sheet as one undoable edit", async (t) => {
  const { app, controller } = await mountController();
  t.after(() => app.unmount());
  const shape = controller.shapes[0];
  controller.selectShape(shape.id);
  assert.ok(controller.cutDrawing());

  await onTargetSheet(controller);
  assert.equal(await controller.pasteDrawing(), true);
  assert.equal(controller.shapes.length, 1, "the moved shape is on the target sheet");
  assert.notEqual(controller.shapes[0].id, shape.id);

  controller.setActiveSheetIndex(0);
  await waitFor(() => controller.shapes.length === 0, "the cut source sheet no longer has the shape");
  controller.undo();
  await waitFor(() => controller.canRedo, "undo completes");
  controller.setActiveSheetIndex(0);
  await vue.nextTick();
  await waitFor(() => controller.shapes.length === 1, "one undo restores the cut source");
  assert.equal(controller.shapes[0].id, shape.id);
});

test("a drawing pastes onto a worksheet added at runtime", async (t) => {
  const { app, controller } = await mountController();
  t.after(() => app.unmount());
  const image = controller.images[0];
  controller.selectImage(image.id);
  assert.ok(controller.copyDrawing());

  controller.addSheet("Extra");
  await vue.nextTick();
  controller.selectCell({ row: 1, col: 1 });
  assert.equal(await controller.pasteDrawing(), true);
  assert.equal(controller.images.length, 1);
  assert.equal(controller.activeSheet?.name, "Extra");
});

test("duplicate copies the selected drawing in place with a small offset", async (t) => {
  const { app, controller } = await mountController();
  t.after(() => app.unmount());
  const shape = controller.shapes[0];
  controller.selectShape(shape.id);
  controller.selectCell({ row: 0, col: 0 });
  controller.selectShape(shape.id);

  assert.equal(controller.duplicateDrawing(), true);
  assert.equal(controller.shapes.length, 2);
  const duplicate = controller.shapes.find((entry) => entry.id !== shape.id);
  assert.ok(duplicate, "duplicate has a fresh identity");
  assert.notDeepEqual(duplicate.anchor, shape.anchor);
  assert.equal(controller.selectedShapeId, duplicate.id);

  controller.undo();
  await waitFor(() => controller.shapes.length === 1 && controller.canRedo, "undo removes the duplicate");
  assert.equal(controller.shapes[0].id, shape.id);
});

test("drawing layers wire Ctrl+C/X/V/D and the drawing context menu", () => {
  const imageLayer = readFileSync(
    new URL("../../packages/vue-xlsx/src/components/XlsxImageLayer.vue", import.meta.url),
    "utf8",
  );
  const drawingLayer = readFileSync(
    new URL("../../packages/vue-xlsx/src/components/XlsxDrawingLayer.vue", import.meta.url),
    "utf8",
  );
  const chartOverlay = readFileSync(
    new URL("../../packages/vue-xlsx/src/components/XlsxChartOverlay.vue", import.meta.url),
    "utf8",
  );
  const surface = readFileSync(
    new URL("../../packages/vue-xlsx/src/components/XlsxSheetSurface.vue", import.meta.url),
    "utf8",
  );
  for (const source of [imageLayer, drawingLayer, chartOverlay]) {
    assert.match(source, /event\.key\.toLowerCase\(\) === "c"/);
    assert.match(source, /event\.key\.toLowerCase\(\) === "x"/);
    assert.match(source, /event\.key\.toLowerCase\(\) === "v"/);
    assert.match(source, /event\.key\.toLowerCase\(\) === "d"/);
  }
  for (const source of [imageLayer, drawingLayer]) {
    assert.match(source, /@contextmenu\.prevent\.stop="on\w+ContextMenu/);
    assert.match(source, /drawingContextMenu/);
  }
  assert.match(surface, /@drawing-context-menu="onDrawingContextMenu"/);
  assert.match(
    surface,
    /contextMenuRef\.value\?\.openForDrawing\(payload\.kind, payload\.id, payload\.clientX, payload\.clientY\)/,
  );
});
