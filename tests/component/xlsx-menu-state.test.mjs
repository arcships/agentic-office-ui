import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import ts from "../../packages/xlsx-core/node_modules/typescript/lib/typescript.js";
import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
} from "./vue-test-renderer.mjs";

const ROOT = path.resolve(import.meta.dirname, "../..");
const CONTEXT_MENU_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxContextMenu.vue");
const VIEWER_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxViewer.vue");
const RIBBON_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxRibbon.vue");
const TOOLBAR_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxToolbar.vue");
const SHEET_TABS_PATH = path.join(ROOT, "packages/vue-xlsx/src/components/XlsxSheetTabs.vue");
const EDITING_PATH = path.join(ROOT, "packages/vue-xlsx/src/composables/editing.ts");
const DEMO_PATH = path.join(ROOT, "apps/demo/src/pages/XlsxViewerPage.vue");
const CONTEXT_MENU_SOURCE = readFileSync(CONTEXT_MENU_PATH, "utf8");
const VIEWER_SOURCE = readFileSync(VIEWER_PATH, "utf8");
const RIBBON_SOURCE = readFileSync(RIBBON_PATH, "utf8");
const TOOLBAR_SOURCE = readFileSync(TOOLBAR_PATH, "utf8");
const SHEET_TABS_SOURCE = readFileSync(SHEET_TABS_PATH, "utf8");
const EDITING_SOURCE = readFileSync(EDITING_PATH, "utf8");
const DEMO_SOURCE = readFileSync(DEMO_PATH, "utf8");
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
  const source = [
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

test("XLSX 右键菜单只监听当前网格，并在显示前选择命中单元格", () => {
  assert.doesNotMatch(CONTEXT_MENU_SOURCE, /document\.addEventListener\(["']contextmenu/);
  assert.match(CONTEXT_MENU_SOURCE, /targetElement/);
  assert.match(VIEWER_SOURCE, /:target-element="gridElement"/);

  const handler = functionNode("onContextMenu").getText(AST);
  assert.ok(handler.indexOf("props.controller.selectCell(cell)") < handler.indexOf("visible.value = true"));

  const target = {
    scrollLeft: 80,
    scrollTop: 24,
    getBoundingClientRect: () => ({ left: 100, top: 200 }),
  };
  const sheet = {
    colWidths: [40, 80],
    rowHeights: [20, 30],
    defaultColWidthPx: 80,
    defaultRowHeightPx: 24,
    visibleCols: [0, 2],
    visibleRows: [1, 4],
  };
  assert.deepEqual(
    geometry.resolveGridCell(target, sheet, 100, 109, 221),
    { col: 2, row: 4 },
    "滚动后仍按可见行列映射到工作表坐标",
  );
  assert.equal(
    geometry.resolveGridCell(target, sheet, 100, 60, 180),
    null,
    "表头外不应伪造单元格",
  );
});

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

test("XLSX 验证页分别公开 Worker 请求值和实际执行来源", () => {
  assert.match(DEMO_SOURCE, /data-testid="xlsx-worker-requested"/);
  assert.match(DEMO_SOURCE, /data-testid="xlsx-worker-actual"/);
  assert.match(DEMO_SOURCE, /current\.isWorkerBacked \? "Worker" : "主线程"/);
  assert.match(DEMO_SOURCE, /@update:read-only="onReadOnlyChange"/);
});

test("XLSX 格式栏读取活动单元格样式，并将字体效果作为可取消开关", () => {
  assert.match(RIBBON_SOURCE, /const activeCellStyle = computed/);
  assert.match(RIBBON_SOURCE, /getCellStyleAt\(cell\.row, cell\.col\)/);
  assert.match(RIBBON_SOURCE, /:value="fontFamilyValue"/);
  assert.match(RIBBON_SOURCE, /:value="fontSizeValue"/);
  assert.match(RIBBON_SOURCE, /bold: !fontState\.bold/);
  assert.match(RIBBON_SOURCE, /italic: !fontState\.italic/);
  assert.match(RIBBON_SOURCE, /underline: fontState\.underline \? 'none' : 'single'/);
  assert.match(RIBBON_SOURCE, /wrapText: !alignmentState\.wrapText/);
  assert.match(RIBBON_SOURCE, /vertical: 'top'/);
  assert.match(RIBBON_SOURCE, /vertical: 'center'/);
  assert.match(RIBBON_SOURCE, /vertical: 'bottom'/);
});

test("XLSX 只读切换保持同一个控制器和工作簿实例", () => {
  const viewerKey = DEMO_SOURCE.match(/const viewerKey = computed\(\(\) => \{([\s\S]*?)\n\}\)/)?.[1] ?? "";
  assert.doesNotMatch(viewerKey, /readOnly\.value/);
  assert.match(DEMO_SOURCE, /get readOnly\(\) \{ return props\.readOnly \}/);
  assert.match(DEMO_SOURCE, /readOnly: props\.readOnly/);
  assert.match(VIEWER_SOURCE, /const effectiveReadOnly = computed/);
});

test("XLSX 搜索覆盖所有工作表，并在命中后切换到结果工作表", () => {
  assert.match(TOOLBAR_SOURCE, /Search workbook/);
  assert.match(TOOLBAR_SOURCE, /for \(const \[sheetIndex, sheet\] of props\.controller\.sheets\.entries\(\)\)/);
  assert.match(TOOLBAR_SOURCE, /setActiveSheetIndex\(result\.sheetIndex\)/);
  assert.match(TOOLBAR_SOURCE, /sheetName: sheet\.name/);
});

test("XLSX 工作表删除有正式入口，并在 UI 和控制器两层保护最后一张", () => {
  assert.match(SHEET_TABS_SOURCE, /data-testid="xlsx-remove-sheet"/);
  assert.match(SHEET_TABS_SOURCE, /props\.controller\.sheets\.length > 1/);
  assert.match(SHEET_TABS_SOURCE, /props\.controller\.removeActiveSheet\(\)/);
  assert.match(EDITING_SOURCE, /ctx\.sheets\.value\.length <= 1/);
  assert.match(SHEET_TABS_SOURCE, /aria-label="上一个工作表"/);
  assert.match(SHEET_TABS_SOURCE, /aria-label="下一个工作表"/);
  assert.match(SHEET_TABS_SOURCE, /aria-label="添加工作表"/);
  assert.match(SHEET_TABS_SOURCE, /<svg viewBox="0 0 24 24"/);
});

test("XLSX 表格标题右键菜单提供真实升序和降序命令", () => {
  assert.match(CONTEXT_MENU_SOURCE, /data-testid="xlsx-sort-ascending"/);
  assert.match(CONTEXT_MENU_SOURCE, /data-testid="xlsx-sort-descending"/);
  assert.match(CONTEXT_MENU_SOURCE, /candidate\.headerRowCount/);
  assert.match(CONTEXT_MENU_SOURCE, /controller\.sortTable\(target\.tableName, target\.columnIndex, direction\)/);
});

test("XLSX 产品工具栏触发上传，演示页不直接展示原生文件控件", () => {
  assert.match(TOOLBAR_SOURCE, /data-testid="xlsx-toolbar-upload"/);
  assert.match(VIEWER_SOURCE, /@upload="emit\('upload'\)"/);
  assert.match(DEMO_SOURCE, /class="visually-hidden" data-testid="xlsx-file-input"/);
  assert.match(DEMO_SOURCE, /<details class="runtime-details">/);
  assert.doesNotMatch(DEMO_SOURCE, /<details class="runtime-details" open>/);
});
