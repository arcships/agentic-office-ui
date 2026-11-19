import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { XlsxToolbar, useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

let wasmConfigured = false;

function configureWasm() {
  if (wasmConfigured) return;
  wasmConfigured = true;
  const bytes = readFileSync(
    new URL(
      "../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm",
      import.meta.url,
    ),
  );
  setWasmSource(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
}

function workbookBuffer(name) {
  const bytes = readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function mountController(name, options = {}) {
  configureWasm();
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController({
        file: workbookBuffer(name),
        fileName: name,
        skipXmlParsing: true,
        useWorker: false,
        ...options,
      });
      return () => vue.h("div", { "data-testid": "controller-harness" }, controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.error,
    `workbook ${name} to load`,
  );
  assert.equal(controller.error, null, controller.error?.message);
  return { ...mounted, get controller() { return controller; } };
}

function cellValue(controller, sheetIndex, row, col) {
  const sheet = controller.sheets[sheetIndex];
  const worksheet = controller.workbook.getSheet(sheet.workbookSheetIndex);
  try {
    return worksheet.getFormattedValueAt(row, col);
  } finally {
    worksheet.free();
  }
}

function cellFormula(controller, sheetIndex, row, col) {
  const sheet = controller.sheets[sheetIndex];
  const worksheet = controller.workbook.getSheet(sheet.workbookSheetIndex);
  try {
    return worksheet.getFormulaAt(row, col);
  } finally {
    worksheet.free();
  }
}

test("replaceInSheet rewrites values across the used range and undo restores them", async () => {
  const { app, controller } = await mountController("sales-table.xlsx");
  try {
    const before = cellValue(controller, 0, 4, 2);
    assert.equal(before, "PDF Viewer");

    const result = controller.replaceInSheet("PDF Viewer", "PDF Reader");
    assert.equal(result.changed, true);
    assert.equal(result.sheetIndex, 0);
    assert.equal(result.replacedCells, 46);
    assert.equal(result.replacedOccurrences, 46);
    assert.equal(cellValue(controller, 0, 4, 2), "PDF Reader");
    assert.equal(cellValue(controller, 0, 5, 2), "PDF Reader");

    assert.equal(controller.canUndo, true);
    controller.undo();
    await vue.nextTick();
    assert.equal(cellValue(controller, 0, 4, 2), "PDF Viewer");
    assert.equal(controller.canUndo, false, "one replace is one undo entry");

    controller.redo();
    await vue.nextTick();
    assert.equal(cellValue(controller, 0, 4, 2), "PDF Reader");
  } finally {
    app.unmount();
  }
});

test("replaceInSheet honours matchCase, wholeCell and range scoping", async () => {
  const { app, controller } = await mountController("sales-table.xlsx");
  try {
    // Case-sensitive: the fixture stores "North", so the lowercase query misses.
    assert.equal(controller.replaceInSheet("north", "Northland", { matchCase: true }).changed, false);
    assert.equal(cellValue(controller, 0, 1, 0), "North");

    // Whole cell: "PDF" is only part of "PDF Viewer", so nothing matches.
    assert.equal(
      controller.replaceInSheet("PDF", "Doc", { wholeCell: true }).changed,
      false,
    );
    assert.equal(cellValue(controller, 0, 4, 2), "PDF Viewer");

    // Substring replace in the same column works.
    const substrings = controller.replaceInSheet("PDF Viewer", "Doc Viewer");
    assert.equal(substrings.replacedCells, 46);
    assert.equal(cellValue(controller, 0, 4, 2), "Doc Viewer");
    controller.undo();
    await vue.nextTick();

    // Range scoping limits the scan to the requested cells.
    const scoped = controller.replaceInSheet("North", "Northland", {
      range: { start: { row: 1, col: 0 }, end: { row: 4, col: 0 } },
    });
    assert.equal(scoped.replacedCells, 2);
    assert.equal(cellValue(controller, 0, 1, 0), "Northland");
    assert.equal(cellValue(controller, 0, 3, 0), "APAC");
    assert.equal(cellValue(controller, 0, 4, 0), "Northland");
    controller.undo();
    await vue.nextTick();
    assert.equal(cellValue(controller, 0, 1, 0), "North");

    // Case-insensitive is the default and matches "North" as well.
    const insensitive = controller.replaceInSheet("north", "Northland", {
      range: { start: { row: 1, col: 0 }, end: { row: 4, col: 0 } },
    });
    assert.equal(insensitive.replacedCells, 2);
  } finally {
    app.unmount();
  }
});

test("replaceInSheet replaces formulas and leaves cached formula values alone", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    assert.equal(cellFormula(controller, 1, 4, 1), "=B2*22%");

    const formulas = controller.replaceInSheet("22%", "25%", { sheetIndex: 1 });
    assert.equal(formulas.changed, true);
    assert.equal(formulas.replacedCells, 3);
    assert.equal(cellFormula(controller, 1, 4, 1), "=B2*25%");
    assert.equal(cellFormula(controller, 1, 4, 2), "=C2*25%");

    controller.undo();
    await vue.nextTick();
    assert.equal(cellFormula(controller, 1, 4, 1), "=B2*22%");

    // `formula: false` keeps formulas intact even when the query appears there.
    assert.equal(
      controller.replaceInSheet("22%", "25%", { sheetIndex: 1, formula: false }).changed,
      false,
    );
    assert.equal(cellFormula(controller, 1, 4, 1), "=B2*22%");

    // A formatted result of a formula is not a searchable cell value here: the
    // formula stays authoritative instead of being replaced by text.
    assert.equal(
      controller.replaceInSheet("$11,088", "0", { sheetIndex: 1 }).changed,
      false,
    );
    assert.equal(cellFormula(controller, 1, 4, 1), "=B2*22%");
  } finally {
    app.unmount();
  }
});

test("replaceInSheet activates a non-active sheet before recording history", async () => {
  const { app, controller } = await mountController("financial-model.xlsx");
  try {
    controller.setActiveSheetIndex(1);
    await vue.nextTick();
    assert.equal(controller.activeSheetIndex, 1);

    const result = controller.replaceInSheet("Starting Customers", "Starting Users", {
      sheetIndex: 0,
    });
    assert.equal(result.changed, true);
    assert.equal(controller.activeSheetIndex, 0, "the replaced sheet becomes active");
    assert.equal(cellValue(controller, 0, 1, 0), "Starting Users");

    controller.undo();
    await vue.nextTick();
    assert.equal(cellValue(controller, 0, 1, 0), "Starting Customers");
  } finally {
    app.unmount();
  }
});

test("replaceInSheet replaces a single reported occurrence", async () => {
  const { app, controller } = await mountController("sales-table.xlsx");
  try {
    const result = controller.replaceInSheet("North", "Northland", {
      occurrence: { cell: { row: 1, col: 0 }, start: 0, end: 5 },
    });
    assert.equal(result.changed, true);
    assert.equal(result.replacedCells, 1);
    assert.equal(result.replacedOccurrences, 1);
    assert.equal(cellValue(controller, 0, 1, 0), "Northland");
    assert.equal(cellValue(controller, 0, 4, 0), "North", "other cells keep their text");
    assert.equal(controller.activeCell.row, 1);
    assert.equal(controller.activeCell.col, 0);

    assert.equal(
      controller.replaceInSheet("North", "Northland", {
        occurrence: { cell: { row: 1, col: 0 }, start: 6, end: 11 },
      }).changed,
      false,
      "an occurrence whose offsets no longer hold the query is rejected",
    );
  } finally {
    app.unmount();
  }
});

test("replaceInSheet is a no-op for read-only workbooks", async () => {
  const { app, controller } = await mountController("sales-table.xlsx", { readOnly: true });
  try {
    const result = controller.replaceInSheet("North", "Northland");
    assert.equal(result.changed, false);
    assert.equal(result.replacedCells, 0);
    assert.equal(cellValue(controller, 0, 1, 0), "North");
    assert.equal(controller.canUndo, false);
  } finally {
    app.unmount();
  }
});

// ── Search panel UI ────────────────────────────────────────────────
test("the XLSX search panel exposes replace and rewrites through the controller", async () => {
  const { app, controller } = await mountController("sales-table.xlsx");
  try {
    const mounted = await mount(vue.defineComponent({
      setup() {
        return () => vue.h(XlsxToolbar, { controller });
      },
    }));
    try {
      findByTestId(mounted.root, "xlsx-search-toggle").props.onClick();
      await vue.nextTick();

      const panel = findByTestId(mounted.root, "xlsx-search-panel");
      assert.ok(panel, "the search panel opens from the toolbar");
      assert.ok(findByTestId(mounted.root, "office-find-replace-input"));
      assert.ok(findByTestId(mounted.root, "xlsx-replace-whole-cell"));
      assert.ok(findByTestId(mounted.root, "xlsx-replace-formulas"));

      findByTestId(mounted.root, "xlsx-search-input").props.onInput({
        target: { value: "PDF Viewer" },
      });
      await waitFor(
        () => {
          const node = findByTestId(mounted.root, "office-find-replace-all");
          return node.props.disabled === false;
        },
        "the workbook search reports results to the panel",
      );

      findByTestId(mounted.root, "office-find-replace-input").props.onInput({
        target: { value: "PDF Reader" },
      });
      await vue.nextTick();
      findByTestId(mounted.root, "office-find-replace-all").props.onClick();
      await waitFor(
        () => cellValue(controller, 0, 4, 2) === "PDF Reader",
        "Replace All rewrites the workbook",
      );
      assert.equal(cellValue(controller, 0, 5, 2), "PDF Reader");

      // The result list follows the edit: the query no longer matches.
      await waitFor(
        () => findByTestId(mounted.root, "office-find-replace-all").props.disabled === true,
        "the search results refresh after replace",
      );
      assert.match(
        walk(findByTestId(mounted.root, "office-find-replace-status")).map((node) => node.text ?? "").join(""),
        /46/u,
      );

      controller.undo();
      await vue.nextTick();
      assert.equal(cellValue(controller, 0, 4, 2), "PDF Viewer");
    } finally {
      mounted.app.unmount();
    }
  } finally {
    app.unmount();
  }
});

test("a read-only workbook hides the XLSX replace controls", async () => {
  const { app, controller } = await mountController("sales-table.xlsx", { readOnly: true });
  try {
    const mounted = await mount(vue.defineComponent({
      setup() {
        return () => vue.h(XlsxToolbar, { controller });
      },
    }));
    try {
      findByTestId(mounted.root, "xlsx-search-toggle").props.onClick();
      await vue.nextTick();
      assert.equal(findByTestId(mounted.root, "office-find-replace-input"), undefined);
      assert.equal(findByTestId(mounted.root, "xlsx-replace-whole-cell"), undefined);
    } finally {
      mounted.app.unmount();
    }
  } finally {
    app.unmount();
  }
});
