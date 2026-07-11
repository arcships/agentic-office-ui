#!/usr/bin/env node

import { existsSync, readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const ts = requireFromDocxCore("typescript");
const checks = [];
const violations = [];

function repoPath(file) {
  return relative(root, file).split(sep).join("/");
}

function source(relativePath) {
  return readFileSync(resolve(root, relativePath), "utf8");
}

function scriptFromVue(relativePath) {
  const text = source(relativePath);
  const match = /<script setup(?:\s+lang="ts")?>([\s\S]*?)<\/script>/u.exec(text);
  if (!match) throw new Error(`${relativePath} 缺少 script setup。`);
  return { text, script: match[1] ?? "" };
}

function parse(relativePath, text) {
  return ts.createSourceFile(relativePath, text, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
}

function functionBody(relativePath, text, name) {
  const file = parse(relativePath, text);
  let body;
  function visit(node) {
    if (ts.isFunctionDeclaration(node) && node.name?.text === name) body = node.body;
    if (
      ts.isVariableDeclaration(node) &&
      ts.isIdentifier(node.name) &&
      node.name.text === name &&
      node.initializer &&
      (ts.isArrowFunction(node.initializer) || ts.isFunctionExpression(node.initializer))
    ) {
      body = node.initializer.body;
    }
    ts.forEachChild(node, visit);
  }
  visit(file);
  return body?.getText(file);
}

function addCheck(id, pass, detail) {
  checks.push({ id, status: pass ? "PASS" : "FAIL", detail });
  if (!pass) violations.push({ id, detail });
}

const viewerPath = "packages/vue-docx/src/components/DocxViewer.vue";
const viewer = scriptFromVue(viewerPath);
for (const deadFunction of [
  "pageHeaderNodes",
  "pageFooterNodes",
  "renderNodeRuns",
  "paragraphBlockStyle",
  "resolveTableColumnWidths",
]) {
  addCheck(
    `DOCX_VIEWER_NO_${deadFunction}`,
    functionBody(viewerPath, viewer.script, deadFunction) === undefined,
    `DocxViewer 不得保留旧函数 ${deadFunction}。`,
  );
}
addCheck(
  "DOCX_VIEWER_ONE_SURFACE",
  (viewer.text.match(/<DocxDocumentSurface\b/gu) ?? []).length === 1,
  "DocxViewer 必须只接入一个 DocxDocumentSurface。",
);

const thumbnailsPath = "packages/vue-docx/src/composables/useDocxPageThumbnails.ts";
const thumbnails = source(thumbnailsPath);
for (const [name, requiredCall] of [
  ["prefetchPageThumbnailSurface", "renderPageSurface"],
  ["rerenderAttachedThumbnails", "renderPageThumbnailToCanvas"],
  ["rerender", "paintUnavailableThumbnail"],
]) {
  const body = functionBody(thumbnailsPath, thumbnails, name);
  addCheck(
    `DOCX_THUMBNAIL_${name}`,
    Boolean(body?.includes(requiredCall)),
    `${name} 必须执行 ${requiredCall}，不能是空实现。`,
  );
}
addCheck(
  "DOCX_THUMBNAIL_EXPLICIT_UNAVAILABLE",
  thumbnails.includes('availability: "unavailable"') &&
    thumbnails.includes("unavailableReason: VIEWER_THUMBNAIL_UNAVAILABLE_MESSAGE"),
  "无文档输入的兼容缩略图接口必须公开 unavailable 状态和原因。",
);

const pageSurface = source("packages/vue-docx/src/components/DocxPageSurface.vue");
addCheck(
  "DOCX_PAGE_SURFACE_PUBLISHER",
  pageSurface.includes("registry.pageElements.set(props.pageIndex, element)") &&
    pageSurface.includes("notifyDocxViewerPageSurfaceListeners(controller)"),
  "页面 surface 必须向编辑器实例 registry 发布并通知变更。",
);

const selection = source("packages/vue-xlsx/src/components/XlsxSelectionOverlay.vue");
addCheck(
  "XLSX_SELECTION_OVERLAY_ACCESSIBLE",
  selection.includes('role="status"') &&
    selection.includes("controller.selectedRangeAddress") &&
    !selection.includes('<div class="xlsx-selection-overlay" />'),
  "XlsxSelectionOverlay 必须提供真实的选区无障碍状态，不能挂空 div。",
);

const chartPath = "packages/vue-xlsx/src/components/XlsxChartOverlay.vue";
const chart = scriptFromVue(chartPath);
addCheck(
  "XLSX_CHART_NO_EMPTY_DBLCLICK",
  !chart.text.includes('@dblclick.stop="onChartDoubleClick') &&
    functionBody(chartPath, chart.script, "onChartDoubleClick") === undefined,
  "图表背景不得用空双击处理器拦截事件。",
);

const xlsxRootPath = "packages/xlsx-core/src/index.ts";
const xlsxRoot = source(xlsxRootPath);
for (const [name, requiredCall] of [
  ["columnLabel", "coreColumnLabel"],
  ["rangeToA1", "coreRangeToA1"],
]) {
  const body = functionBody(xlsxRootPath, xlsxRoot, name);
  addCheck(
    `XLSX_ROOT_${name}_DELEGATES_TO_CORE`,
    Boolean(body?.includes(requiredCall)) && !body?.includes("while"),
    `xlsx-core 根入口 ${name} 必须保留兼容签名并委托 core 单一算法。`,
  );
}

for (const barrel of [
  "packages/vue-docx/src/components/index.ts",
  "packages/vue-xlsx/src/components/index.ts",
]) {
  addCheck(
    `NO_UNUSED_BARREL_${barrel}`,
    !existsSync(resolve(root, barrel)),
    `${barrel} 没有公开出口和消费者，应保持删除。`,
  );
}

const viewerBarrel = source("packages/docx-core/src/viewer/index.ts");
addCheck(
  "DOCX_NO_UNSUPPORTED_DEEP_IMPORT_GUIDANCE",
  !viewerBarrel.includes("@arcships/docx-core/viewer/wasm-source") &&
    viewerBarrel.includes("unsupported source-file deep imports"),
  "源码注释不得建议消费端使用导出表拒绝的深层路径。",
);

const report = {
  task: "P2-DEAD-01",
  result: violations.length === 0 ? "PASS" : "FAIL",
  checks,
  violations,
};
process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
if (violations.length > 0) process.exitCode = 1;
