import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

import {
  findByTestId,
  importFromDemo,
  mount,
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
// `vModelText` feature-detects focus restoration through these globals; the
// custom test renderer hands it plain host nodes, which are neither.
globalThis.Document ??= class Document {};
globalThis.ShadowRoot ??= class ShadowRoot {};
const { strFromU8, unzipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);

const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const { useXlsxViewerController, XlsxRibbon } = await importFromDemo(
  "@arcships/vue-xlsx",
);

function sampleBuffer(name) {
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

async function mountReadyController(options) {
  const mounted = await mountController(options);
  await waitFor(
    () => mounted.controller.sourceState === "ready" || mounted.controller.error,
    "XLSX controller ready state",
  );
  assert.ifError(mounted.controller.error);
  assert.equal(mounted.controller.sourceState, "ready");
  return mounted;
}

/** Captures `controller.exportXlsx()` downloads instead of touching the DOM. */
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
  const download = downloads.at(-1);
  assert.ok(download?.blob, "export produced a workbook download");
  const exported = new Uint8Array(await download.blob.arrayBuffer());
  return { archive: unzipSync(exported), exported };
}

function sampleBytes(exported) {
  return exported.buffer.slice(
    exported.byteOffset,
    exported.byteOffset + exported.byteLength,
  );
}

/** Runs `reader` against a borrowed worksheet handle and frees only that handle. */
function withWorksheet(controller, reader) {
  const worksheet = controller.getActiveWorksheet();
  assert.ok(worksheet, "controller exposes the active worksheet");
  try {
    return reader(worksheet);
  } finally {
    worksheet.free?.();
  }
}

const EDGE_ORDER = ["top", "right", "bottom", "left"];
const BORDER_RANGE = { start: { row: 1, col: 1 }, end: { row: 3, col: 2 } };

function borderCells(range = BORDER_RANGE) {
  const cells = [];
  for (let row = range.start.row; row <= range.end.row; row += 1) {
    for (let col = range.start.col; col <= range.end.col; col += 1) {
      cells.push({ row, col });
    }
  }
  return cells;
}

function drawnEdges(style) {
  const border = style?.border ?? {};
  return EDGE_ORDER.filter((edge) => {
    const value = border[edge];
    return Boolean(value?.style) && value.style !== "none";
  });
}

function readBorder(controller, cell) {
  return withWorksheet(controller, (worksheet) =>
    drawnEdges(worksheet.getCellStyleAt(cell.row, cell.col)));
}

function readAlignment(controller, cell) {
  return withWorksheet(controller, (worksheet) => {
    const alignment = worksheet.getCellStyleAt(cell.row, cell.col)?.alignment;
    return alignment
      ? { horizontal: alignment.horizontal, vertical: alignment.vertical }
      : null;
  });
}

/** Expected perimeter edges, written out independently of the implementation. */
function outsideEdges(cell, range = BORDER_RANGE) {
  const edges = [];
  if (cell.row === range.start.row) edges.push("top");
  if (cell.col === range.end.col) edges.push("right");
  if (cell.row === range.end.row) edges.push("bottom");
  if (cell.col === range.start.col) edges.push("left");
  return EDGE_ORDER.filter((edge) => edges.includes(edge));
}

const UNIFORM_PRESET_EDGES = {
  all: ["top", "right", "bottom", "left"],
  bottom: ["bottom"],
  left: ["left"],
  right: ["right"],
  top: ["top"],
};

function expectedEdges(preset, cell) {
  if (preset === "outside") return outsideEdges(cell);
  if (preset === "none") return [];
  return EDGE_ORDER.filter((edge) => UNIFORM_PRESET_EDGES[preset].includes(edge));
}

function edgeColours(controller, cell) {
  return withWorksheet(controller, (worksheet) => {
    const border = worksheet.getCellStyleAt(cell.row, cell.col)?.border ?? {};
    return EDGE_ORDER
      .filter((edge) => drawnEdges({ border }).includes(edge))
      .map((edge) => border[edge]?.color?.hex ?? null);
  });
}

const PRESET_RUNS = [
  { hex: "FF0000", preset: "all" },
  { hex: "00FF00", preset: "outside" },
  { hex: "0000FF", preset: "left" },
  { hex: "FF8800", preset: "right" },
  { hex: "8800FF", preset: "top" },
  { hex: "00FFFF", preset: "bottom" },
  { hex: "FF0000", preset: "none" },
];

test("every border preset writes its edges to the model and survives export and reload", async (t) => {
  const downloads = installDownloadCapture(t);

  for (const { hex, preset } of PRESET_RUNS) {
    const { app, controller } = await mountReadyController({
      file: sampleBuffer("financial-model.xlsx"),
      fileName: "financial-model.xlsx",
    });
    try {
      controller.selectRange(BORDER_RANGE);
      if (preset === "none") {
        // `No border` clears what a previous preset drew, so start from a
        // fully bordered range and prove the command removes every edge.
        assert.equal(
          controller.setSelectionBorders({ color: { colorType: "rgb", hex }, preset: "all" }),
          true,
        );
        for (const cell of borderCells()) {
          assert.deepEqual(readBorder(controller, cell), expectedEdges("all", cell));
        }
      }

      assert.equal(
        controller.setSelectionBorders({ color: { colorType: "rgb", hex }, preset }),
        true,
        `${preset} borders apply to the selection`,
      );
      for (const cell of borderCells()) {
        assert.deepEqual(
          readBorder(controller, cell),
          expectedEdges(preset, cell),
          `${preset} draws the expected edges on ${cell.row},${cell.col}`,
        );
      }

      controller.exportXlsx();
      await vue.nextTick();
      const { archive, exported } = await exportedArchive(downloads);
      const sheetXml = strFromU8(archive["xl/worksheets/sheet1.xml"]);
      const stylesXml = strFromU8(archive["xl/styles.xml"]);
      if (preset === "none") {
        assert.doesNotMatch(stylesXml, /<border>(?:(?!<\/border>)[\s\S])*style=/);
      } else {
        assert.match(
          stylesXml,
          new RegExp(`<color rgb="FF${hex}"/>`),
          `${preset} writes its border colour into the package`,
        );
        assert.doesNotMatch(stylesXml, /<color[^>]*rgb="#/, "exported colours are ARGB, not CSS hex");
      }
      assert.match(sheetXml, /<worksheet\b/);

      const reloaded = await mountReadyController({
        file: sampleBytes(exported),
        fileName: `borders-${preset}.xlsx`,
      });
      try {
        for (const cell of borderCells()) {
          assert.deepEqual(
            readBorder(reloaded.controller, cell),
            expectedEdges(preset, cell),
            `${preset} keeps its edges on ${cell.row},${cell.col} after reload`,
          );
          if (preset !== "none") {
            const expectedCount = expectedEdges(preset, cell).length;
            assert.deepEqual(
              edgeColours(reloaded.controller, cell),
              Array(expectedCount).fill(`FF${hex}`),
              `${preset} keeps its border colour on ${cell.row},${cell.col} after reload`,
            );
          }
        }
      } finally {
        reloaded.app.unmount();
      }
    } finally {
      app.unmount();
    }
  }
});

test("outside borders leave the inside of the selection untouched", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountReadyController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
  });
  t.after(() => app.unmount());

  // Border a 3x3 block first, so `outside` has an inner cell and inner edges it
  // must keep rather than a blank range it could trivially leave alone.
  const range = { start: { row: 1, col: 1 }, end: { row: 3, col: 3 } };
  controller.selectRange(range);
  controller.setSelectionBorders({ preset: "all" });
  controller.setSelectionBorders({ preset: "outside", color: { colorType: "rgb", hex: "FF0000" } });

  const interior = { row: 2, col: 2 };
  assert.deepEqual(
    readBorder(controller, interior),
    ["top", "right", "bottom", "left"],
    "the inner cell keeps the borders it already had",
  );
  assert.deepEqual(
    edgeColours(controller, interior),
    ["A1A1AA", "A1A1AA", "A1A1AA", "A1A1AA"],
    "the inner cell keeps its own border colour",
  );
  assert.deepEqual(
    edgeColours(controller, { row: 1, col: 1 }),
    ["FF0000", "A1A1AA", "A1A1AA", "FF0000"],
    "only the outside edges of the corner cell take the picked colour",
  );

  controller.exportXlsx();
  await vue.nextTick();
  const { archive } = await exportedArchive(downloads);
  const stylesXml = strFromU8(archive["xl/styles.xml"]);
  assert.match(stylesXml, /<left style="thin"><color rgb="FFFF0000"\/><\/left>/);
  assert.match(stylesXml, /<top style="thin"><color rgb="FFFF0000"\/><\/top>/);
});

test("merge & center merges, centres and undoes in one step", async (t) => {
  const downloads = installDownloadCapture(t);
  const { app, controller } = await mountReadyController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
  });
  t.after(() => app.unmount());

  const range = { start: { row: 0, col: 0 }, end: { row: 0, col: 2 } };
  const originalAlignment = readAlignment(controller, range.start);
  controller.selectRange(range);
  controller.mergeAndCenterSelection();

  for (const cell of borderCells(range)) {
    assert.deepEqual(
      readAlignment(controller, cell),
      { horizontal: "center", vertical: "center" },
      `${cell.row},${cell.col} is centred both ways`,
    );
  }
  assert.deepEqual(
    withWorksheet(controller, (worksheet) => worksheet.mergedRegions.map((region) => region.range)),
    ["A1:C1"],
  );

  // The merge and the alignment share one snapshot, so a single undo is enough.
  controller.undo();
  await waitFor(() => controller.canRedo, "undo records a redo entry");
  assert.deepEqual(
    withWorksheet(controller, (worksheet) => worksheet.mergedRegions.map((region) => region.range)),
    [],
    "undo removes the merge region",
  );
  assert.notEqual(readAlignment(controller, { row: 0, col: 0 })?.vertical, "center");
  assert.deepEqual(
    readAlignment(controller, { row: 0, col: 0 }),
    originalAlignment,
    "undo restores the alignment the range had before merge & center",
  );
  controller.redo();
  await waitFor(
    () => withWorksheet(controller, (worksheet) => worksheet.mergedRegions.length) === 1,
    "redo restores the merge region",
  );

  controller.exportXlsx();
  await vue.nextTick();
  const { archive, exported } = await exportedArchive(downloads);
  assert.match(strFromU8(archive["xl/worksheets/sheet1.xml"]), /<mergeCell ref="A1:C1"\/>/);

  const reloaded = await mountReadyController({
    file: sampleBytes(exported),
    fileName: "merged.xlsx",
  });
  try {
    for (const cell of borderCells(range)) {
      assert.deepEqual(
        readAlignment(reloaded.controller, cell),
        { horizontal: "center", vertical: "center" },
        `${cell.row},${cell.col} stays centred after reload`,
      );
    }
    assert.deepEqual(
      withWorksheet(reloaded.controller, (worksheet) => worksheet.mergedRegions.map((region) => region.range)),
      ["A1:C1"],
    );
  } finally {
    reloaded.app.unmount();
  }
});

test("merge & center only centres a single cell", async (t) => {
  const { app, controller } = await mountReadyController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
  });
  t.after(() => app.unmount());

  const cell = { row: 3, col: 3 };
  controller.selectCell(cell);
  controller.mergeAndCenterSelection();

  assert.deepEqual(readAlignment(controller, cell), { horizontal: "center", vertical: "center" });
  assert.deepEqual(
    withWorksheet(controller, (worksheet) => worksheet.mergedRegions),
    [],
    "Excel never writes a one-cell merge region",
  );
});

test("plain merge stays available for hosts and keeps the alignment", async (t) => {
  const { app, controller } = await mountReadyController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
  });
  t.after(() => app.unmount());

  const range = { start: { row: 3, col: 2 }, end: { row: 4, col: 3 } };
  controller.selectRange(range);
  const before = readAlignment(controller, range.start);
  controller.mergeSelection();

  assert.deepEqual(
    withWorksheet(controller, (worksheet) => worksheet.mergedRegions.map((region) => region.range)),
    ["C4:D5"],
  );
  assert.deepEqual(readAlignment(controller, range.start), before, "plain merge leaves the text alignment alone");
});

test("border and merge commands are no-ops while the sheet is read-only", async (t) => {
  const { app, controller } = await mountReadyController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
    readOnly: true,
  });
  t.after(() => app.unmount());

  controller.selectRange(BORDER_RANGE);
  assert.equal(controller.setSelectionBorders({ preset: "all" }), false);
  for (const cell of borderCells()) {
    assert.deepEqual(readBorder(controller, cell), []);
  }

  controller.mergeAndCenterSelection();
  assert.deepEqual(
    withWorksheet(controller, (worksheet) => worksheet.mergedRegions),
    [],
  );
  assert.notEqual(readAlignment(controller, { row: 1, col: 1 })?.horizontal, "center");
});

test("undo restores an unstyled cell to the workbook default formatting", async (t) => {
  const { app, controller } = await mountReadyController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
  });
  t.after(() => app.unmount());

  // D4 carries no explicit style, so undo restores the parsed workbook default
  // (a theme-coloured font) instead of the engine's own default.
  const cell = { row: 3, col: 3 };
  assert.equal(withWorksheet(controller, (worksheet) => worksheet.getCellStyleAt(cell.row, cell.col)), null);

  controller.selectCell(cell);
  controller.setSelectedCellStyle({ font: { bold: true } });
  assert.equal(
    withWorksheet(controller, (worksheet) => worksheet.getCellStyleAt(cell.row, cell.col)?.font?.bold),
    true,
  );

  controller.undo();
  await vue.nextTick();
  assert.equal(
    withWorksheet(controller, (worksheet) => worksheet.getCellStyleAt(cell.row, cell.col)?.font?.bold),
    false,
    "undo clears the font without tripping over the parsed theme colour",
  );
});

function setModel(node, value) {
  const update = node?.props?.["onUpdate:modelValue"];
  assert.equal(typeof update, "function", "control exposes a model update handler");
  update(value);
}

function fakeRibbonController(overrides = {}) {
  const calls = [];
  const controller = {
    activeCell: null,
    activeSheet: null,
    canRedo: false,
    canUndo: false,
    error: null,
    freezePanes: null,
    isLoading: false,
    readOnly: false,
    revision: 0,
    selection: { start: { row: 1, col: 1 }, end: { row: 2, col: 2 } },
    tabs: [{ id: "sheet-0", kind: "sheet" }],
    zoomScale: 100,
    mergeAndCenterSelection: () => calls.push(["mergeAndCenterSelection"]),
    mergeSelection: () => calls.push(["mergeSelection"]),
    setRangeStyle: (range, style) => calls.push(["setRangeStyle", range, style]),
    setSelectedCellStyle: (style) => calls.push(["setSelectedCellStyle", style]),
    setSelectionBorders: (input) => {
      calls.push(["setSelectionBorders", input]);
      return true;
    },
    unmergeSelection: () => calls.push(["unmergeSelection"]),
    ...overrides,
  };
  return { calls, controller };
}

test("XlsxRibbon drives border presets, the border colour and merge & center", async () => {
  const { calls, controller } = fakeRibbonController();
  const mounted = await mount(XlsxRibbon, { controller });
  try {
    const presets = findByTestId(mounted.root, "xlsx-ribbon-border-preset");
    assert.ok(presets, "the ribbon renders the border preset dropdown");
    presets.props.onChange({ target: { value: "outside" } });
    assert.deepEqual(calls.at(-1), [
      "setSelectionBorders",
      { color: { colorType: "rgb", hex: "A1A1AA" }, preset: "outside" },
    ]);

    const colour = findByTestId(mounted.root, "xlsx-ribbon-border-color");
    assert.ok(colour, "the ribbon renders the border colour picker");
    setModel(colour, "#00ff00");
    await vue.nextTick();
    colour.props.onChange();
    assert.deepEqual(
      calls.at(-1),
      ["setSelectionBorders", { color: { colorType: "rgb", hex: "00FF00" }, preset: "outside" }],
      "the colour picker re-applies the preset it belongs to",
    );

    findByTestId(mounted.root, "xlsx-ribbon-borders-all").props.onClick();
    assert.deepEqual(calls.at(-1), [
      "setSelectionBorders",
      { color: { colorType: "rgb", hex: "00FF00" }, preset: "all" },
    ]);

    findByTestId(mounted.root, "xlsx-ribbon-merge-center").props.onClick();
    assert.deepEqual(calls.at(-1), ["mergeAndCenterSelection"]);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("the all-borders shortcut still works on a controller without the border command", async () => {
  const { calls, controller } = fakeRibbonController({
    activeCell: { row: 1, col: 1 },
    setSelectionBorders: undefined,
  });
  const mounted = await mount(XlsxRibbon, { controller });
  try {
    findByTestId(mounted.root, "xlsx-ribbon-borders-all").props.onClick();
    const [name, range, style] = calls.at(-1);
    assert.equal(name, "setRangeStyle");
    assert.deepEqual(range, controller.selection);
    assert.deepEqual(style, {
      border: {
        bottom: { color: { colorType: "rgb", hex: "A1A1AA" }, style: "thin" },
        left: { color: { colorType: "rgb", hex: "A1A1AA" }, style: "thin" },
        right: { color: { colorType: "rgb", hex: "A1A1AA" }, style: "thin" },
        top: { color: { colorType: "rgb", hex: "A1A1AA" }, style: "thin" },
      },
    });

    // Presets the old controller cannot express are skipped instead of falling
    // back to a different border than the user picked.
    findByTestId(mounted.root, "xlsx-ribbon-border-preset").props.onChange({
      target: { value: "outside" },
    });
    assert.equal(calls.length, 1);
  } finally {
    mounted.app.unmount();
  }
});
