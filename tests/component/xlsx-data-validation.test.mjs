import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";
import { importFromDemo } from "./vue-test-renderer.mjs";

const ROOT = path.resolve(import.meta.dirname, "../..");
const requireFromVueXlsx = createRequire(new URL("../../packages/vue-xlsx/package.json", import.meta.url));
const { strToU8, zipSync } = requireFromVueXlsx("fflate");
const requireFromXlsxCore = createRequire(new URL("../../packages/xlsx-core/package.json", import.meta.url));
const xmldom = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= xmldom.DOMParser;
globalThis.Node ??= xmldom.Node;
globalThis.XMLSerializer ??= xmldom.XMLSerializer;

function transpile(relativePath) {
  const filePath = path.join(ROOT, relativePath);
  const source = readFileSync(filePath, "utf8");
  return ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ESNext,
      target: ts.ScriptTarget.ES2022,
    },
  }).outputText;
}

function dataUrl(source) {
  return `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;
}

const selectionUrl = dataUrl(transpile("packages/vue-xlsx/src/composables/selection.ts"));
const validationSource = transpile("packages/vue-xlsx/src/composables/validation.ts")
  .replace(/from\s+["']\.\/selection["']/g, `from ${JSON.stringify(selectionUrl)}`);
const validationUrl = dataUrl(validationSource);
const selection = await import(selectionUrl);
const validation = await import(validationUrl);

const GRID_SOURCE = readFileSync(path.join(ROOT, "packages/vue-xlsx/src/components/XlsxGrid.vue"), "utf8");
const VIEWER_SOURCE = readFileSync(path.join(ROOT, "packages/vue-xlsx/src/components/XlsxViewer.vue"), "utf8");
const EDITING_SOURCE = readFileSync(path.join(ROOT, "packages/vue-xlsx/src/composables/editing.ts"), "utf8");
const CLIPBOARD_SOURCE = readFileSync(path.join(ROOT, "packages/vue-xlsx/src/composables/clipboard.ts"), "utf8");
const WORKBOOK_STATE_SOURCE = readFileSync(path.join(ROOT, "packages/vue-xlsx/src/composables/workbook-state.ts"), "utf8");
const CONTROLLER_SOURCE = readFileSync(path.join(ROOT, "packages/vue-xlsx/src/composables/useXlsxViewerController.ts"), "utf8");
const WORKER_SOURCE = readFileSync(path.join(ROOT, "packages/xlsx-core/src/xlsx-worker.ts"), "utf8");

function listValidation(overrides = {}) {
  return {
    validationType: "list",
    ranges: [{ start: { row: 1, col: 0 }, end: { row: 9, col: 0 } }],
    listSource: "Alpha,Beta,Gamma",
    showDropdown: true,
    showErrorAlert: true,
    showInputMessage: true,
    inputMessage: "Choose a value",
    errorMessage: "Pick Alpha, Beta or Gamma",
    errorStyle: "stop",
    ...overrides,
  };
}

function sheetWith(validations) {
  return { dataValidations: validations };
}

function createValidationWorkbookBytes() {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Validation" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
  const worksheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Choice</t></is></c><c r="B1" t="inlineStr"><is><t>Color</t></is></c><c r="C1" t="inlineStr"><is><t>Amount</t></is></c><c r="E1" t="inlineStr"><is><t>Red</t></is></c><c r="E2" t="inlineStr"><is><t>Green</t></is></c></row>
<row r="2"><c r="A2" t="inlineStr"><is><t>Alpha</t></is></c></row>
</sheetData>
<sheetProtection sheet="1" formatCells="0" insertRows="0" selectLockedCells="0"/>
<dataValidations count="3">
<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorStyle="stop" sqref="A2:A10" showDropDown="0" promptTitle="Choose" prompt="Pick one" errorTitle="Bad value" error="Must be Alpha/Beta/Gamma"><formula1>"Alpha,Beta,Gamma"</formula1></dataValidation>
<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="B2:B10" showDropDown="1"><formula1>$E$1:$E$2</formula1></dataValidation>
<dataValidation type="whole" operator="between" allowBlank="1" showErrorMessage="1" errorStyle="warning" sqref="C2:C10"><formula1>1</formula1><formula2>10</formula2></dataValidation>
</dataValidations>
</worksheet>`;

  return zipSync({
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(rootRels),
    "xl/workbook.xml": strToU8(workbook),
    "xl/_rels/workbook.xml.rels": strToU8(workbookRels),
    "xl/worksheets/sheet1.xml": strToU8(worksheet),
  });
}

test("data validation helpers resolve inline, range and named lists", () => {
  assert.deepEqual(
    validation.resolveDataValidationListOptions(listValidation(), {}),
    ["Alpha", "Beta", "Gamma"],
  );
  assert.deepEqual(
    validation.resolveDataValidationListOptions(listValidation({ listSource: "\"One,Two\"" }), {}),
    ["One", "Two"],
  );
  assert.deepEqual(
    validation.resolveDataValidationListOptions(
      listValidation({ listSource: "$E$1:$E$2" }),
      { readRangeValues: (reference) => reference === "$E$1:$E$2" ? ["Red", "Green"] : null },
    ),
    ["Red", "Green"],
  );
  assert.deepEqual(
    validation.resolveDataValidationListOptions(
      listValidation({ listSource: "Colors" }),
      {
        getNamedRange: (name) => name === "Colors" ? "Sheet1!$E$1:$E$2" : null,
        readRangeValues: (reference) => reference === "Sheet1!$E$1:$E$2" ? ["Red", "Green"] : null,
      },
    ),
    ["Red", "Green"],
  );
  assert.equal(validation.isDataValidationListDropdownVisible(listValidation()), true);
  assert.equal(validation.isDataValidationListDropdownVisible(listValidation({ showDropdown: false })), false);
  assert.deepEqual(
    validation.resolveDataValidationListOptions(listValidation({ validationType: "whole" }), {}),
    [],
  );
});

test("data validation finds the covering rule and enforces stop, warning and information styles", () => {
  const sheet = sheetWith([listValidation()]);
  const cell = { row: 2, col: 0 };
  assert.equal(validation.findDataValidationForCell(sheet.dataValidations, cell)?.validationType, "list");

  const valid = validation.checkCellValueAgainstDataValidation(sheet, cell, "Beta");
  assert.deepEqual(valid, { accepted: true, feedback: null });

  const caseInsensitive = validation.checkCellValueAgainstDataValidation(sheet, cell, "beta");
  assert.equal(caseInsensitive.accepted, true);

  const invalid = validation.checkCellValueAgainstDataValidation(sheet, cell, "Zulu");
  assert.equal(invalid.accepted, false);
  assert.equal(invalid.feedback?.kind, "error");
  assert.equal(invalid.feedback?.message, "Pick Alpha, Beta or Gamma");
  assert.equal(invalid.feedback?.cell.col, 0);

  const warning = validation.checkCellValueAgainstDataValidation(
    sheetWith([listValidation({ errorStyle: "warning" })]),
    cell,
    "Zulu",
  );
  assert.equal(warning.accepted, true);
  assert.equal(warning.feedback?.kind, "warning");

  const information = validation.checkCellValueAgainstDataValidation(
    sheetWith([listValidation({ errorStyle: "information" })]),
    cell,
    "Zulu",
  );
  assert.equal(information.accepted, true);
  assert.equal(information.feedback?.kind, "information");

  const silentStop = validation.checkCellValueAgainstDataValidation(
    sheetWith([listValidation({ showErrorAlert: false })]),
    cell,
    "Zulu",
  );
  assert.deepEqual(silentStop, { accepted: false, feedback: null });
});

test("data validation enforces blank, numeric bounds and text length rules", () => {
  const cell = { row: 1, col: 0 };
  const required = listValidation({ allowBlank: false });
  assert.equal(validation.checkCellValueAgainstDataValidation(sheetWith([required]), cell, "").accepted, false);
  assert.equal(validation.checkCellValueAgainstDataValidation(sheetWith([listValidation()]), cell, "").accepted, true);

  const whole = {
    validationType: "whole",
    operator: "between",
    value1: "1",
    value2: "10",
    ranges: [ { start: { row: 1, col: 0 }, end: { row: 1, col: 0 } } ],
  };
  assert.equal(validation.checkCellValueAgainstDataValidation(sheetWith([whole]), cell, "5").accepted, true);
  assert.equal(validation.checkCellValueAgainstDataValidation(sheetWith([whole]), cell, "11").accepted, false);
  assert.equal(validation.checkCellValueAgainstDataValidation(sheetWith([whole]), cell, "5.5").accepted, false);

  const textLength = {
    validationType: "textLength",
    operator: "lessThanOrEqual",
    value1: "3",
    ranges: [ { start: { row: 1, col: 0 }, end: { row: 1, col: 0 } } ],
  };
  assert.equal(validation.checkCellValueAgainstDataValidation(sheetWith([textLength]), cell, "abc").accepted, true);
  assert.equal(validation.checkCellValueAgainstDataValidation(sheetWith([textLength]), cell, "abcd").accepted, false);
});

test("sheet protection parsing exposes enabled flags", () => {
  assert.equal(validation.isSheetEditingLocked(false, true, false), true);
  assert.equal(validation.isSheetEditingLocked(false, true, true), false);
  assert.equal(validation.isSheetEditingLocked(false, false, false), false);
  assert.equal(validation.isSheetEditingLocked(true, false, true), true);

  const protectedSheet = selection.parseWorksheetSheetProtection({
    protection: {
      protected: true,
      formatCells: true,
      insertRows: true,
      sort: false,
    },
  });
  assert.deepEqual(protectedSheet, {
    enabled: true,
    flags: { formatCells: true, insertRows: true, sort: false },
  });
  assert.equal(selection.parseWorksheetSheetProtection({ protection: null }), null);
  assert.equal(selection.parseWorksheetSheetProtection({ protection: { protected: false } }), null);
});

test("WASM worksheet validation and protection records are parsed", async () => {
  const wasmModule = await import(
    pathToFileURL(path.join(ROOT, "packages/xlsx-core/node_modules/@dukelib/sheets-wasm/duke_sheets_wasm.js")).href
  );
  const wasmBytes = readFileSync(
    path.join(ROOT, "packages/xlsx-core/node_modules/@dukelib/sheets-wasm/duke_sheets_wasm_bg.wasm"),
  );
  await wasmModule.default({ module_or_path: new Uint8Array(wasmBytes) });
  const workbook = wasmModule.Workbook.fromBytes(createValidationWorkbookBytes());
  const worksheet = workbook.getSheet(0);
  try {
    const validations = selection.parseWorksheetDataValidations(worksheet);
    assert.equal(validations.length, 3);
    assert.equal(validations[0].validationType, "list");
    assert.equal(validations[0].listSource, "Alpha,Beta,Gamma");
    assert.equal(validations[0].showDropdown, true);
    assert.equal(validations[0].ranges[0].start.col, 0);
    assert.equal(validations[1].listSource, "$E$1:$E$2");
    assert.equal(validations[1].showDropdown, false);
    assert.equal(validations[2].operator, "between");
    assert.equal(validations[2].value1, "1");
    assert.equal(validations[2].value2, "10");
    assert.equal(validations[2].errorStyle, "warning");

    const protection = selection.parseWorksheetSheetProtection(worksheet);
    assert.equal(protection?.enabled, true);
    assert.equal(protection?.flags.formatCells, true);
    assert.equal(protection?.flags.insertRows, true);
  } finally {
    worksheet.free();
    workbook.free();
  }
});

test("XML structure parser exposes validation prompts and error text", async () => {
  const { parseWorkbookStructureAssets } = await importFromDemo("@arcships/xlsx-core");
  const assets = parseWorkbookStructureAssets(createValidationWorkbookBytes());
  const validations = assets.sheetStatesByWorkbookSheetIndex[0]?.dataValidations ?? [];
  assert.equal(validations.length, 3);
  assert.equal(validations[0].listSource, '"Alpha,Beta,Gamma"');
  assert.equal(validations[0].inputTitle, "Choose");
  assert.equal(validations[0].inputMessage, "Pick one");
  assert.equal(validations[0].errorTitle, "Bad value");
  assert.equal(validations[0].errorMessage, "Must be Alpha/Beta/Gamma");
  assert.equal(validations[1].showDropdown, false);
  assert.equal(validations[2].operator, "between");
  assert.equal(validations[2].value1, "1");
  assert.equal(validations[2].value2, "10");
  assert.equal(validations[2].errorStyle, "warning");
});

test("grid renders validation dropdowns, hints and alerts and commits through the controller", () => {
  assert.match(GRID_SOURCE, /data-testid="xlsx-validation-dropdown"/);
  assert.match(GRID_SOURCE, /data-testid="xlsx-validation-dropdown-trigger"/);
  assert.match(GRID_SOURCE, /data-testid="xlsx-validation-hint"/);
  assert.match(GRID_SOURCE, /data-testid="xlsx-validation-alert"/);
  assert.match(GRID_SOURCE, /getCellDataValidation/);
  assert.match(GRID_SOURCE, /getDataValidationListOptions/);
  assert.match(GRID_SOURCE, /isDataValidationListDropdownVisible/);
  assert.match(GRID_SOURCE, /chooseValidationOption/);
  assert.match(GRID_SOURCE, /setCellValue\(cell, option\)/);
});

test("controller wires validation enforcement and sheet protection read-only behaviour", () => {
  assert.match(CONTROLLER_SOURCE, /allowProtectedSheetEditing = false/);
  assert.match(CONTROLLER_SOURCE, /isActiveSheetProtected = computed/);
  assert.match(CONTROLLER_SOURCE, /isEditingLocked = computed/);
  assert.match(CONTROLLER_SOURCE, /checkCellValueAgainstDataValidation/);
  assert.match(CONTROLLER_SOURCE, /getCellDataValidation/);
  assert.match(CONTROLLER_SOURCE, /getDataValidationListOptions/);
  assert.match(CONTROLLER_SOURCE, /dismissDataValidationFeedback/);
  assert.match(EDITING_SOURCE, /ctx\.isEditingLocked\.value/);
  assert.match(EDITING_SOURCE, /ctx\.checkCellValue\(/);
  assert.match(CLIPBOARD_SOURCE, /ctx\.isEditingLocked\.value/);
  assert.match(CLIPBOARD_SOURCE, /ctx\.checkCellValue\(/);
  assert.match(WORKBOOK_STATE_SOURCE, /sheetState\?\.dataValidations \?\? parseWorksheetDataValidations/);
  assert.match(VIEWER_SOURCE, /isEditingLocked/);
  assert.match(WORKER_SOURCE, /parseWorksheetSheetProtection/);
  assert.match(WORKER_SOURCE, /sheetState\?\.dataValidations \?\? parseWorksheetDataValidations/);
});
