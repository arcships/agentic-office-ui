import { readFileSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const rootDir = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const requireFromXlsxCore = createRequire(
  resolve(rootDir, "packages/xlsx-core/package.json"),
);
const { DOMParser, Node, XMLSerializer } = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= DOMParser;
globalThis.Node ??= Node;
globalThis.XMLSerializer ??= XMLSerializer;

const outputPath = process.argv[2];
if (!outputPath) {
  console.error("Usage: node scripts/ci/export-xlsx-roundtrip.mjs <output.xlsx>");
  process.exit(2);
}

const core = await import(
  pathToFileURL(resolve(rootDir, "packages/xlsx-core/dist/index.js")).href
);
const wasmBytes = readFileSync(
  resolve(rootDir, "packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm"),
);
core.setWasmSource(
  wasmBytes.buffer.slice(
    wasmBytes.byteOffset,
    wasmBytes.byteOffset + wasmBytes.byteLength,
  ),
);
const wasm = await core.getSheetsWasmModule();

const sourceBytes = new Uint8Array(
  readFileSync(resolve(rootDir, "apps/demo/public/samples/charts-images.xlsx")),
);
const workbook = wasm.Workbook.fromBytes(sourceBytes);
const imageAssets = core.parseWorkbookImageAssets(sourceBytes);
const chartAssets = core.loadWorkbookChartAssets(
  workbook,
  imageAssets,
  new Map([[0, 0]]),
);
const chart = chartAssets.chartsByWorkbookSheetIndex[0]?.[0];
if (!chart?.chartPath) {
  throw new Error("charts-images.xlsx must expose an editable chart");
}
core.updateWorkbookChartDefinition(imageAssets, chartAssets, chart.id, {
  title: "Openpyxl roundtrip chart",
});
workbook.getSheetByName("Dashboard").setCell("A1", "OPENPYXL-ROUNDTRIP");

const exported = core.mergeWorkbookImageAssets(
  workbook.saveXlsxBytes(),
  imageAssets,
  imageAssets.sheetOrigins,
);
workbook.free();
writeFileSync(outputPath, exported);
console.log(JSON.stringify({ outputPath, bytes: exported.byteLength }));
