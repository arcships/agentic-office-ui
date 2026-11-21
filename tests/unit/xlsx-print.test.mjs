import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

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
const core = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);

const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

function sampleBytes(name = "financial-model.xlsx") {
  return new Uint8Array(readFileSync(new URL(name, samples)));
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
      '<printOptions gridLines="1" headings="1" horizontalCentered="1" verticalCentered="0"/>',
      '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="1" scale="85" firstPageNumber="3" useFirstPageNumber="1"/>',
      '<rowBreaks count="2" manualBreakCount="2"><brk id="5" min="0" max="16383" man="1"/><brk id="12" min="0" max="16383" man="1"/></rowBreaks>',
      '<colBreaks count="1" manualBreakCount="1"><brk id="3" min="0" max="1048575" man="1"/></colBreaks>',
      "</worksheet>",
    ].join(""),
  );
  archive["xl/worksheets/sheet1.xml"] = strToU8(sheet);

  let workbook = strFromU8(archive["xl/workbook.xml"]);
  workbook = workbook.replace(
    "</sheets>",
    [
      "</sheets>",
      '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Assumptions!$A$1:$D$20</definedName>',
      '<definedName name="_xlnm.Print_Titles" localSheetId="0">Assumptions!$1:$2,Assumptions!$A:$B</definedName></definedNames>',
    ].join(""),
  );
  archive["xl/workbook.xml"] = strToU8(workbook);
  return { archive, bytes: zipSync(archive) };
}

function makeSheet(overrides = {}) {
  const rowCount = 100;
  const colCount = 20;
  return {
    cachedFormulaValues: {},
    colStyleIds: {},
    colWidthOverridesPx: {},
    colWidths: Array.from({ length: colCount }, () => 80),
    colCount,
    columnWidthCharacterWidthPx: 7,
    comments: [],
    conditionalFormatRules: [],
    dataValidations: [],
    defaultColWidthPx: 80,
    defaultRowHeightPx: 24,
    freezePanes: null,
    hasHorizontalMerges: false,
    hasVerticalMerges: false,
    hiddenCols: [],
    hiddenRows: [],
    hyperlinks: [],
    maxHorizontalMergeEndCol: -1,
    maxUsedCol: colCount - 1,
    maxUsedRow: rowCount - 1,
    maxVerticalMergeEndRow: -1,
    mergedRegions: [],
    minUsedCol: 0,
    minUsedRow: 0,
    name: "Sheet1",
    namedCellStyleByName: {},
    pageMargins: { ...core.DEFAULT_PAGE_MARGINS },
    pageSetup: { ...core.DEFAULT_PAGE_SETUP },
    printAreas: [],
    printOptions: { ...core.DEFAULT_PRINT_OPTIONS },
    printTitles: { columns: [], rows: [] },
    protection: null,
    rowCount,
    rowHeights: Array.from({ length: rowCount }, () => 24),
    rowHeightOverridesPx: {},
    rowStyleIds: {},
    showGridLines: true,
    sparklines: [],
    styleById: {},
    tableStyleByName: {},
    themePalette: { colorsByIndex: {} },
    visibility: "visible",
    visibleCols: Array.from({ length: colCount }, (_, index) => index),
    visibleRows: Array.from({ length: rowCount }, (_, index) => index),
    workbookSheetIndex: 0,
    zoomScale: 100,
    ...overrides,
  };
}

const wasmBytes = readFileSync(
  requireFromXlsxCore.resolve("@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm"),
);
core.setWasmSource(
  wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
);
const wasm = await core.getSheetsWasmModule();

test("worksheet print state is parsed into the sheet model", () => {
  const { bytes } = patchPrintMetadata();
  const assets = core.parseWorkbookStructureAssets(bytes);
  const state = assets.sheetStatesByWorkbookSheetIndex[0];
  assert.ok(state, "the first worksheet state must parse");

  assert.equal(state.pageSetup.paperSize, 9);
  assert.equal(state.pageSetup.orientation, "landscape");
  assert.equal(state.pageSetup.fitToPage, true);
  assert.equal(state.pageSetup.fitToWidth, 1);
  assert.equal(state.pageSetup.fitToHeight, 1);
  assert.equal(state.pageSetup.scale, 85);
  assert.equal(state.pageSetup.firstPageNumber, 3);
  assert.equal(state.pageSetup.useFirstPageNumber, true);

  assert.deepEqual(state.pageMargins, {
    bottom: 0.8,
    footer: 0.3,
    header: 0.2,
    left: 0.5,
    right: 0.6,
    top: 0.7,
  });
  assert.deepEqual(state.printOptions, {
    gridLines: true,
    headings: true,
    horizontalCentered: true,
    verticalCentered: false,
  });
  assert.deepEqual(state.rowBreaks.map((pageBreak) => pageBreak.id), [5, 12]);
  assert.deepEqual(state.colBreaks.map((pageBreak) => pageBreak.id), [3]);
  assert.deepEqual(state.printAreas, [{
    end: { col: 3, row: 19 },
    start: { col: 0, row: 0 },
  }]);
  assert.deepEqual(state.printTitles, {
    columns: [{ end: 1, start: 0 }],
    rows: [{ end: 1, start: 0 }],
  });
});

test("multiple print areas stay separate in the parsed model", () => {
  const { archive } = patchPrintMetadata();
  archive["xl/workbook.xml"] = strToU8(
    strFromU8(archive["xl/workbook.xml"]).replace(
      "Assumptions!$A$1:$D$20",
      "Assumptions!$A$1:$D$20,Assumptions!$F$1:$H$10",
    ),
  );
  const state = core.parseWorkbookStructureAssets(zipSync(archive)).sheetStatesByWorkbookSheetIndex[0];
  assert.deepEqual(state.printAreas, [
    { end: { col: 3, row: 19 }, start: { col: 0, row: 0 } },
    { end: { col: 7, row: 9 }, start: { col: 5, row: 0 } },
  ]);
});

test("fit-to counts of 0 survive parsing as automatic pages", () => {
  // Excel writes `fitToHeight="0"` when the `tall` box is left blank, which
  // means "as many pages as the data needs" — not one page.
  const { archive } = patchPrintMetadata();
  archive["xl/worksheets/sheet1.xml"] = strToU8(
    strFromU8(archive["xl/worksheets/sheet1.xml"]).replace(
      /<pageSetup[^>]*\/>/,
      '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>',
    ),
  );
  const state = core.parseWorkbookStructureAssets(zipSync(archive)).sheetStatesByWorkbookSheetIndex[0];
  assert.equal(state.pageSetup.fitToPage, true);
  assert.equal(state.pageSetup.fitToWidth, 1);
  assert.equal(state.pageSetup.fitToHeight, 0);
});

test("page pagination is computed in core from paper, breaks and print area", () => {
  const layout = core.buildWorksheetPrintLayout(makeSheet());
  assert.equal(layout.paperSize.name, "Letter");
  assert.equal(layout.pages[0].widthPx, 816);
  assert.equal(layout.pages[0].heightPx, 1056);
  assert.equal(layout.pages[0].contentWidthPx, 682);
  assert.equal(layout.pages[0].contentHeightPx, 912);
  assert.equal(layout.scale, 100);
  assert.equal(layout.pages.length, 9, "20 columns and 100 rows make 3 × 3 Letter pages");

  const ranges = layout.pages.map((page) => page.range);
  assert.deepEqual(ranges.slice(0, 3).map((range) => range.start.row), [0, 38, 76]);
  assert.deepEqual([...new Set(ranges.map((range) => range.start.col))].sort((a, b) => a - b), [0, 8, 16]);
  for (const page of layout.pages) {
    assert.ok(page.range.start.row <= page.range.end.row);
    assert.ok(page.range.start.col <= page.range.end.col);
  }

  const manual = core.buildWorksheetPrintLayout(makeSheet({
    colBreaks: [{ id: 3, max: 1048575, min: 0, manual: true }],
    maxUsedCol: 3,
    rowBreaks: [{ id: 5, max: 16383, min: 0, manual: true }],
    visibleCols: [0, 1, 2, 3],
    colWidths: [80, 80, 80, 80],
    printAreas: [{ start: { col: 0, row: 0 }, end: { col: 3, row: 19 } }],
  }));
  assert.equal(manual.pages.length, 4);
  assert.deepEqual(manual.pages.map((page) => [page.rowPage, page.colPage]), [
    [0, 0],
    [1, 0],
    [0, 1],
    [1, 1],
  ]);
  assert.deepEqual(manual.pages.map((page) => page.range.start), [
    { col: 0, row: 0 },
    { col: 0, row: 5 },
    { col: 3, row: 0 },
    { col: 3, row: 5 },
  ]);
});

test("print areas, fit-to-page scaling and landscape paper drive page geometry", () => {
  const limited = core.buildWorksheetPrintLayout(makeSheet({
    printAreas: [{ end: { col: 3, row: 19 }, start: { col: 0, row: 0 } }],
  }));
  assert.equal(limited.pages.length, 1);
  assert.equal(limited.pages[0].range.start.col, 0);
  assert.equal(limited.pages[0].range.end.col, 3);
  assert.equal(limited.pages[0].range.end.row, 19);

  const fitted = core.buildWorksheetPrintLayout(makeSheet({
    pageSetup: {
      ...core.DEFAULT_PAGE_SETUP,
      fitToHeight: 1,
      fitToPage: true,
      fitToWidth: 1,
    },
  }));
  assert.ok(fitted.scale < 100, "fit-to-page shrinks a large sheet");
  assert.equal(fitted.pages.length, 1, "one width and one height page fit the whole sheet");
  assert.ok(fitted.pages[0].widthPx * 0.01 * fitted.scale >= 0);

  const landscape = core.buildWorksheetPrintLayout(makeSheet({
    pageSetup: {
      ...core.DEFAULT_PAGE_SETUP,
      orientation: "landscape",
      paperSize: 9,
    },
  }));
  assert.equal(landscape.paperSize.name, "A4");
  assert.ok(landscape.pages[0].widthPx > landscape.pages[0].heightPx);
  assert.equal(landscape.pages[0].orientation, "landscape");

  const overThenDown = core.buildWorksheetPrintLayout(makeSheet({
    pageSetup: {
      ...core.DEFAULT_PAGE_SETUP,
      pageOrder: "overThenDown",
    },
  }));
  assert.deepEqual(overThenDown.pages.slice(0, 3).map((page) => page.rowPage), [0, 0, 0]);
  assert.deepEqual(overThenDown.pages.slice(0, 3).map((page) => page.colPage), [0, 1, 2]);
});

test("print titles repeat on every page and merged cells keep their spans", () => {
  const sheet = makeSheet({
    maxUsedCol: 3,
    maxUsedRow: 10,
    mergedRegions: [{ end: { col: 1, row: 1 }, start: { col: 0, row: 0 } }],
    minUsedRow: 2,
    printAreas: [{ end: { col: 3, row: 10 }, start: { col: 0, row: 2 } }],
    printTitles: {
      columns: [{ end: 0, start: 0 }],
      rows: [{ end: 1, start: 0 }],
    },
    visibleCols: [0, 1, 2, 3],
    visibleRows: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
  });
  const layout = core.buildWorksheetPrintLayout(sheet);
  assert.equal(layout.pages.length, 1);
  const grid = core.buildPrintPageGrid(sheet, layout.pages[0], layout.printTitles);
  assert.deepEqual(grid.rows.map((row) => row.index), [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
  assert.deepEqual(grid.columns.map((column) => column.index), [0, 1, 2, 3]);
  const anchor = grid.cells.find((cell) => cell.row === 0 && cell.col === 0);
  assert.ok(anchor);
  assert.equal(anchor.rowSpan, 2);
  assert.equal(anchor.colSpan, 2);
  const covered = grid.cells.find((cell) => cell.row === 1 && cell.col === 1);
  assert.ok(covered?.covered);
});

test("export writes page layout and print defined names, and reload parses them", () => {
  const { archive, bytes } = patchPrintMetadata();
  const workbook = wasm.Workbook.fromBytes(bytes);
  let saved;
  try {
    saved = workbook.saveXlsxBytes();
  } finally {
    workbook.free();
  }
  const merged = core.mergeWorkbookSheetAuthoringAssets(saved, archive);
  const reloaded = core.parseWorkbookStructureAssets(merged);
  const state = reloaded.sheetStatesByWorkbookSheetIndex[0];

  assert.equal(state.pageSetup.paperSize, 9);
  assert.equal(state.pageSetup.orientation, "landscape");
  assert.equal(state.pageSetup.fitToPage, true);
  assert.equal(state.pageSetup.scale, 85);
  assert.equal(state.printOptions.gridLines, true);
  assert.deepEqual(state.pageMargins, {
    bottom: 0.8,
    footer: 0.3,
    header: 0.2,
    left: 0.5,
    right: 0.6,
    top: 0.7,
  });
  assert.deepEqual(state.rowBreaks.map((pageBreak) => pageBreak.id), [5, 12]);
  assert.deepEqual(state.colBreaks.map((pageBreak) => pageBreak.id), [3]);
  assert.deepEqual(state.printAreas, [{
    end: { col: 3, row: 19 },
    start: { col: 0, row: 0 },
  }]);
  assert.deepEqual(state.printTitles, {
    columns: [{ end: 1, start: 0 }],
    rows: [{ end: 1, start: 0 }],
  });

  const savedSheet = strFromU8(unzipSync(merged)["xl/worksheets/sheet1.xml"]);
  assert.match(savedSheet, /<pageMargins left="0.5"/);
  assert.match(savedSheet, /<printOptions[^>]*gridLines="1"/);
  assert.match(savedSheet, /<pageSetUpPr fitToPage="1"/);
});

test("engine fallback keeps print metadata when worksheet XML is skipped", () => {
  const { bytes } = patchPrintMetadata();
  const workbook = wasm.Workbook.fromBytes(bytes);
  const worksheet = workbook.getSheet(0);
  try {
    const state = core.resolveWorksheetPrintState(null, worksheet);
    assert.equal(state.pageSetup.paperSize, 9);
    assert.equal(state.pageSetup.orientation, "landscape");
    assert.deepEqual(state.printAreas, [{
      end: { col: 3, row: 19 },
      start: { col: 0, row: 0 },
    }]);
    assert.deepEqual(state.printTitles, {
      columns: [{ end: 1, start: 0 }],
      rows: [{ end: 1, start: 0 }],
    });
  } finally {
    worksheet.free();
    workbook.free();
  }
});

test("unknown paper sizes fall back to Letter instead of failing the page list", () => {
  const paper = core.resolvePaperSize(9999);
  assert.equal(paper.id, 1);
  assert.equal(paper.name, "Letter");
  const layout = core.buildWorksheetPrintLayout(makeSheet({
    pageSetup: { ...core.DEFAULT_PAGE_SETUP, paperSize: 9999 },
  }));
  assert.equal(layout.pages[0].paperSize, 1);
  assert.equal(layout.pages[0].paperSizeName, "Letter");
});
