#!/usr/bin/env node

/**
 * Drives the real XLSX export path — controller edits plus the toolbar button —
 * and writes the download the browser would have saved.
 *
 * `tests/blackbox/xlsx_export_controller_openpyxl.py` reopens this file with a
 * pinned openpyxl and asserts the edits, formats and merges survived.
 */

import { readFileSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import process from "node:process";
import { pathToFileURL } from "node:url";

import { findByTestId, mount, vue, waitFor } from "../../tests/component/vue-test-renderer.mjs";

const root = path.resolve(import.meta.dirname, "../..");
const outputPath = process.argv[2];
if (!outputPath) {
  console.error("Usage: node scripts/ci/export-xlsx-controller.mjs <output.xlsx>");
  process.exit(2);
}

const requireFromXlsxCore = createRequire(
  path.join(root, "packages/xlsx-core/package.json"),
);
const xmldom = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= xmldom.DOMParser;
globalThis.Node ??= xmldom.Node;
globalThis.XMLSerializer ??= xmldom.XMLSerializer;

const core = await import(
  pathToFileURL(path.join(root, "packages/xlsx-core/dist/index.js")).href
);
const wasmBytes = readFileSync(
  path.join(root, "packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm"),
);
core.setWasmSource(
  wasmBytes.buffer.slice(
    wasmBytes.byteOffset,
    wasmBytes.byteOffset + wasmBytes.byteLength,
  ),
);

const { useXlsxViewerController, XlsxToolbar } = await import(
  pathToFileURL(path.join(root, "packages/vue-xlsx/dist/index.js")).href
);

const downloads = [];
const blobs = new Map();
let nextObjectUrl = 0;
const originalDocument = globalThis.document;
const originalCreateObjectUrl = URL.createObjectURL;
const originalRevokeObjectUrl = URL.revokeObjectURL;
globalThis.document = {
  body: { append() {} },
  createElement(tagName) {
    return {
      tagName,
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
  const objectUrl = `blob:controller-export-${(nextObjectUrl += 1)}`;
  blobs.set(objectUrl, blob);
  return objectUrl;
};
URL.revokeObjectURL = () => {};

function sampleBuffer(name) {
  const bytes = readFileSync(path.join(root, "apps/demo/public/samples", name));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

let controller;
const Harness = vue.defineComponent({
  setup() {
    controller = useXlsxViewerController({
      file: sampleBuffer("charts-images.xlsx"),
      fileName: "charts-images.xlsx",
      readOnly: false,
      skipXmlParsing: false,
      useWorker: false,
    });
    return () => vue.h("div", controller.sourceState);
  },
});
const mounted = await mount(Harness);

try {
  await waitFor(
    () => controller.sourceState === "ready" || controller.sourceError,
    "XLSX controller ready state",
  );
  if (controller.sourceError) throw controller.sourceError;
  if (controller.sourceState !== "ready") {
    throw new Error(`XLSX controller did not become ready: ${controller.sourceState}`);
  }

  // Edit a cell, give it a real number format, and merge a fresh region so the
  // reopened file proves all three write paths, not just the ZIP round-trip.
  const edited = { row: 5, col: 4 };
  if (!controller.commitCellInput(edited, "CONTROLLER-EXPORT")) {
    throw new Error("the test cell could not be edited");
  }
  // A1 is a styled header cell; editing it must keep the header fill and font.
  if (!controller.commitCellInput({ row: 0, col: 0 }, "CONTROLLER-EXPORT")) {
    throw new Error("the header cell could not be edited");
  }
  controller.setCellStyle(edited, {
    font: { bold: true },
    numberFormat: { formatType: "custom", formatString: "0.00%" },
  });
  controller.selectRange({
    start: { row: 13, col: 0 },
    end: { row: 14, col: 2 },
  });
  controller.mergeSelection();
  await vue.nextTick();

  if (!controller.canExport) {
    throw new Error("the loaded workbook must expose the export path");
  }
  const toolbar = await mount(XlsxToolbar, { controller });
  try {
    const exportButton = findByTestId(toolbar.root, "xlsx-toolbar-export");
    if (!exportButton) throw new Error("the toolbar export button did not render");
    exportButton.props.onClick();
    await vue.nextTick();
  } finally {
    toolbar.app.unmount();
  }

  if (downloads.length !== 1) {
    throw new Error(`expected one export download, got ${downloads.length}`);
  }
  const bytes = new Uint8Array(await downloads[0].blob.arrayBuffer());
  writeFileSync(outputPath, bytes);
  console.log(
    JSON.stringify({
      output: outputPath,
      bytes: bytes.byteLength,
      fileName: downloads[0].fileName,
      edited: "E6",
      editedHeader: "A1",
      merged: "A14:C15",
      numberFormat: "0.00%",
    }),
  );
} finally {
  mounted.app.unmount();
  if (originalDocument === undefined) delete globalThis.document;
  else globalThis.document = originalDocument;
  if (originalCreateObjectUrl === undefined) delete URL.createObjectURL;
  else URL.createObjectURL = originalCreateObjectUrl;
  if (originalRevokeObjectUrl === undefined) delete URL.revokeObjectURL;
  else URL.revokeObjectURL = originalRevokeObjectUrl;
}
