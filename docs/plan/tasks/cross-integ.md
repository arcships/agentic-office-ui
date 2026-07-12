---
id: cross-integ
scope: demo, vue-xlsx, vue-docx
status: pending
depends-on: [xlsx-005, docx-004, docx-005]
---

# 跨线集成验证：viewer → controller, demo → 全链路

## objective

验证 viewer 消费 controller、demo 消费全链路的真实调用路径。覆盖集成点 X4-X7, D6-D9, C1-C2。

## context

- `docs/plan/analysis/migration-split.md` 集成关系枚举 X4-X7, D6-D9, C1-C2
- `docs/upstream-xlsx-feature-alignment.md` 第八节（playground 对齐）、验证清单第十节
- `docs/upstream-docx-feature-alignment.md` 第八节（playground 对齐）、验证清单第九节
- `docs/visual-acceptance-handoff.md` 验收标准

## path

- `packages/vue-xlsx/src/XlsxViewer.vue`（只读验证）
- `packages/vue-docx/src/DocxEditor.vue`（只读验证）
- `packages/vue-docx/src/DocxViewer.vue`（只读验证）
- `apps/demo/src/pages/XlsxViewerPage.vue`（只读验证）
- `apps/demo/src/pages/DocxViewerPage.vue`（只读验证）
- `apps/demo/src/pages/DocxEditorPage.vue`（只读验证）
- `apps/demo/src/main.ts`（只读验证）
- `docs/INDEX.md`（只读验证）

## verification

### X4: XlsxViewer → composables

浏览器 `http://localhost:5173/#/xlsx-viewer`：
- 加载 welcome.xlsx → 单元格渲染（值来自 controller.getSheet）
- 选中单元格 → controller.selectCell
- 双击编辑 → controller.setSelectedCellValue
- 导出 → controller.exportXlsx → 下载 .xlsx
通过标准：viewer 真实消费 controller API，渲染/编辑/导出完整

### X5: XlsxViewer → chart-renderer

浏览器加载 charts-images.xlsx：
- 图表可见（SVG 渲染）
通过标准：至少 bar/line/pie 三种图表渲染

### X6: demo/main → wasm

浏览器控制台无 wasm 加载错误
通过标准：`setWasmSource` 成功，worker 可初始化

### X7: demo/XlsxViewerPage → vue-xlsx

浏览器 `http://localhost:5173/#/xlsx-viewer`：
- fixture 切换正常（至少 3 个 fixture）
- ribbon 功能可用
通过标准：页面渲染 + 交互完整

### D6: DocxEditor → composables

浏览器 `http://localhost:5173/#/docx-editor`：
- 导入 .docx → controller.importDocxFile → 页面渲染
- 编辑文本 → controller.commitParagraphText
- 格式化 → controller.toggleBold/setFontFamily
- 导出 → controller.exportDocx
通过标准：editor 真实消费 controller，渲染/编辑/导出完整

### D7: DocxViewer → layout-engine

浏览器 `http://localhost:5173/#/docx-viewer`：
- 加载 .docx → layoutDocument 分页 → 页面渲染
通过标准：分页正确，段落/表格/图片可见

### D8: demo/main → docx wasm

浏览器控制台无 docx wasm 加载错误

### D9: demo/DocxEditorPage → vue-docx

浏览器 `http://localhost:5173/#/docx-editor`：
- 页面渲染 + toolbar 交互
通过标准：页面渲染 + 交互完整

### C1: 双 wasm 加载

浏览器：xlsx 和 docx 两种 wasm 都加载成功

### C2: 路由切换

浏览器：xlsx-viewer → docx-viewer → docx-editor → xlsx-viewer 切换不崩溃

### 全量 gate

```bash
pnpm typecheck && pnpm build && \
uv run --with python-docx --with openpyxl --with pillow --with pypdf python scripts/verify_test_materials.py && \
git diff --check
```

### 视口矩阵

4 视口逐路由检查：`1440×900` / `1280×720` / `768×1024` / `390×844`
