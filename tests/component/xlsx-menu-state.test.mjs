import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";
import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  waitFor,
} from "./vue-test-renderer.mjs";

const ROOT = path.resolve(import.meta.dirname, "../..");
const CONTEXT_MENU_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxContextMenu.vue");
const GRID_METRICS_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/grid-metrics.ts");
const PANE_GEOMETRY_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/grid-pane-geometry.ts");
const CONTEXT_MENU_SOURCE = readFileSync(CONTEXT_MENU_PATH, "utf8");
const SCRIPT_SOURCE = [...CONTEXT_MENU_SOURCE.matchAll(/<script(?:\s+setup)?[^>]*>([\s\S]*?)<\/script>/g)]
  .map((match) => match[1])
  .join("\n");
const AST = ts.createSourceFile(
  CONTEXT_MENU_PATH,
  SCRIPT_SOURCE,
  ts.ScriptTarget.Latest,
  true,
  ts.ScriptKind.TS,
);

function descendants(node, predicate) {
  const matches = [];
  function visit(current) {
    if (predicate(current)) matches.push(current);
    ts.forEachChild(current, visit);
  }
  visit(node);
  return matches;
}

function functionNode(name) {
  return descendants(
    AST,
    (node) => ts.isFunctionDeclaration(node) && node.name?.text === name,
  )[0];
}

const GRID_METRICS_AST = ts.createSourceFile(
  GRID_METRICS_PATH,
  readFileSync(GRID_METRICS_PATH, "utf8"),
  ts.ScriptTarget.Latest,
  true,
  ts.ScriptKind.TS,
);

function gridMetricsFunction(name) {
  return descendants(
    GRID_METRICS_AST,
    (node) => ts.isFunctionDeclaration(node) && node.name?.text === name,
  )[0];
}

function gridMetricsConstant(name) {
  return descendants(
    GRID_METRICS_AST,
    (node) =>
      ts.isVariableStatement(node) &&
      node.declarationList.declarations.some(
        (declaration) => ts.isIdentifier(declaration.name) && declaration.name.text === name,
      ),
  )[0];
}

function constantNode(name) {
  return descendants(
    AST,
    (node) =>
      ts.isVariableStatement(node) &&
      node.declarationList.declarations.some(
        (declaration) => ts.isIdentifier(declaration.name) && declaration.name.text === name,
      ),
  )[0];
}

async function loadMenuGeometry() {
  const constants = [
    "DEFAULT_ROW_HEIGHT",
    "DEFAULT_COL_WIDTH",
    "HEADER_HEIGHT",
    "ROW_HEADER_WIDTH",
    "MIN_DISPLAY_ROWS",
    "MIN_DISPLAY_COLS",
  ];
  const functions = ["resolveDisplayIndices", "findDisplayIndexAtOffset", "resolveGridCell"];
  // resolveGridCell shares the freeze-pane mapping with the grid, so the
  // transpiled module carries the same helpers.
  const paneGeometry = readFileSync(PANE_GEOMETRY_PATH, "utf8").replace(/^import type .*;\n/m, "");
  const sharedFunctions = ["gridZoomFactor", "frozenExtent"].map((name) => {
    const node = gridMetricsFunction(name);
    assert.ok(node, `grid-metrics 缺少 ${name}`);
    return node.getText(GRID_METRICS_AST);
  });
  const sharedConstants = ["GRID_DEFAULT_COL_WIDTH_PX", "GRID_DEFAULT_ROW_HEIGHT_PX"].map((name) => {
    const node = gridMetricsConstant(name);
    assert.ok(node, `grid-metrics 缺少 ${name}`);
    return node.getText(GRID_METRICS_AST);
  });
  const source = [
    paneGeometry,
    ...sharedConstants,
    ...sharedFunctions,
    ...constants.map((name) => constantNode(name)?.getText(AST)),
    ...functions.map((name) => functionNode(name)?.getText(AST)),
  ];
  assert.ok(source.every(Boolean), "菜单命中测试所需函数必须存在");
  const output = ts.transpileModule(
    `${source.join("\n")}\nexport { ${functions.join(", ")} };`,
    {
      compilerOptions: {
        module: ts.ModuleKind.ESNext,
        target: ts.ScriptTarget.ES2022,
      },
    },
  ).outputText;
  return import(`data:text/javascript;base64,${Buffer.from(output).toString("base64")}`);
}

const geometry = await loadMenuGeometry();

// Loading a real workbook needs the WASM runtime and the XML helpers the
// controller installs for workbook parsing.
const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const xmldom = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= xmldom.DOMParser;
globalThis.Node ??= xmldom.Node;
globalThis.XMLSerializer ??= xmldom.XMLSerializer;
const { setWasmSource } = await importFromDemo("@arcships/xlsx-core");
const wasmBytes = readFileSync(
  new URL("../../packages/xlsx-core/dist/assets/duke_sheets_wasm_bg.wasm", import.meta.url),
);
setWasmSource(
  wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
);
const { useXlsxViewerController } = await importFromDemo("@arcships/vue-xlsx");

function sampleBuffer(name) {
  const bytes = readFileSync(
    new URL(`../../apps/demo/public/samples/${name}`, import.meta.url),
  );
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function mountLoadedController(options) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useXlsxViewerController(options);
      return () => vue.h("div", controller.sourceState);
    },
  });
  const mounted = await mount(Harness);
  await waitFor(
    () => controller.sourceState === "ready" || controller.sourceError,
    "XLSX controller ready state",
  );
  assert.ifError(controller.sourceError);
  assert.equal(controller.sourceState, "ready");
  return { ...mounted, controller };
}

test("XlsxViewer 把 Ribbon 的只读请求作为受控事件转发", async () => {
  const { XlsxViewer } = await importFromDemo("@arcships/vue-xlsx");
  const updates = [];
  const controller = {
    activeCell: null,
    activeSheet: null,
    canRedo: false,
    canUndo: false,
    canZoomIn: false,
    canZoomOut: false,
    error: null,
    isLoading: false,
    readOnly: false,
    selection: null,
    tabs: [],
    zoomScale: 100,
  };
  const mounted = await mount(XlsxViewer, {
    controller,
    showDefaultToolbar: false,
    showFormulaBar: false,
    showRibbon: true,
    "onUpdate:readOnly": (value) => updates.push(value),
  });
  const checkbox = findByTestId(mounted.root, "xlsx-ribbon-read-only");
  assert.ok(checkbox, "应渲染公开可定位的只读开关");
  checkbox.props.onChange({ target: { checked: true } });
  await vue.nextTick();
  assert.deepEqual(updates, [true]);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("XLSX 可编辑加载在主线程解析，Worker 只用于只读加载", async (t) => {
  // `typeof Worker` is the controller's capability check; a stub keeps the
  // decision observable on Node without starting a real worker thread.
  const originalWorker = globalThis.Worker;
  globalThis.Worker = class Worker {};
  t.after(() => {
    if (originalWorker === undefined) delete globalThis.Worker;
    else globalThis.Worker = originalWorker;
  });

  const { app, controller } = await mountLoadedController({
    file: sampleBuffer("financial-model.xlsx"),
    fileName: "financial-model.xlsx",
    readOnly: false,
    skipXmlParsing: true,
    useWorker: true,
  });
  t.after(() => app.unmount());

  assert.equal(controller.readOnly, false);
  assert.equal(
    controller.isWorkerBacked,
    false,
    "an editable load must not claim Worker-backed rendering",
  );
  const worksheet = controller.getActiveWorksheet();
  assert.ok(worksheet, "the main thread still parsed a usable workbook");
  worksheet.free();
});
