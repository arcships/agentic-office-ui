# 迁移任务拆分分析

## 拆分原则

1. **按可独立验证的功能层拆**，不按文件粒度拆（太琐碎）
2. **引擎层先行**：先复制零 React 依赖的纯逻辑，可立即 typecheck 验证
3. **controller 次之**：机械改写 hook → Vue reactivity，可 typecheck + 结构测试
4. **viewer 最后**：重写渲染层，需要浏览器验证
5. **XLSX 和 DOCX 两条线并行**，互不依赖
6. **每个任务是最小可独立验证交付物**：如果修复 A 的 blocking findings 必然产生 B 的交付物，则合并为一个任务

## 模块分解

### XLSX 模块链

```
xlsx-core（引擎层，零 React 依赖）
  ├─ wasm.ts              ← @dukelib/sheets-wasm 封装
  ├─ safe-calculate.ts    ← 公式安全计算
  ├─ worker-client.ts     ← Worker RPC 客户端
  ├─ xlsx-worker.ts       ← Worker 主体
  ├─ colors.ts            ← 颜色解析
  ├─ images.ts            ← 图片/形状/表单控件
  ├─ charts.ts            ← 图表数据模型
  ├─ types.ts             ← 数据模型类型（清理 React 标注）
  └─ index.ts             ← 公开导出
        ↓ 被调用
vue-xlsx（Vue 层）
  ├─ composables.ts       ← controller（hook → Vue reactivity）
  ├─ chart-renderer.ts    ← SVG 图表渲染（JSX → Vue render）
  ├─ surface-regl.ts      ← WebGL 3D surface（JSX → Vue render）
  ├─ XlsxViewer.vue       ← canvas 渲染 + 虚拟化 + 交互（重写）
  ├─ types.ts             ← Vue 组件 props 类型
  └─ index.ts             ← 公开导出
        ↓ 被调用
demo/XlsxViewerPage.vue   ← 路由页面 + fixture
```

### DOCX 模块链

```
docx-core（引擎+布局+辅助+helpers，零 React 依赖）
  ├─ wasm.ts              ← docx-wasm 封装
  ├─ ooxml-core.ts        ← OOXML 包操作
  ├─ serializer.ts        ← DocModel → OOXML 序列化
  ├─ types.ts             ← DocModel 类型体系
  ├─ clone.ts             ← 深拷贝
  ├─ normalize.ts         ← JSON 归一化
  ├─ editor-ops.ts        ← 段落文本编辑操作
  ├─ pagination.ts        ← 分页素材收集
  ├─ page-segmentation.ts ← 分页切分引擎
  ├─ layout-engine.ts     ← 几何装箱
  ├─ pretext-layout.ts    ← 文本排版引擎
  ├─ thumbnail-raster.ts  ← 缩略图光栅化
  ├─ layout-snapshot.ts   ← 布局快照
  ├─ docx-import.ts       ← DOCX 加载（worker/主线程）
  ├─ docx-import-worker.ts← Worker 主体
  ├─ section-layout.ts    ← section XML 解析
  ├─ pagination-breaks.ts ← 分页符检测
  ├─ image-render.ts      ← TIFF/EMF 处理
  ├─ page-count-reconciliation.ts ← 页数校准
  ├─ content-signature.ts ← 内容签名
  ├─ canvas/*             ← canvas 诊断类型
  ├─ wasm-source.ts       ← wasm 源配置
  ├─ core/state.ts        ← 编辑器状态机
  ├─ editor-types.ts      ← DocxEditorSelection/DocxTextRange（从 editor.tsx 抽出）
  ├─ editor-helpers.ts    ← editor.tsx 前半部分纯函数（清理 React 类型）
  └─ index.ts             ← 公开导出
        ↓ 被调用
vue-docx（Vue 层）
  ├─ composables.ts       ← useDocxEditor + 11 composables（hook → Vue reactivity）
  ├─ DocxEditor.vue       ← DOM 渲染 + 虚拟化 + contentEditable（重写）
  ├─ DocxViewer.vue       ← 简单只读渲染（重写）
  ├─ render-paragraph-runs.ts ← 段落 run → Vue render（重写）
  ├─ types.ts             ← Vue 组件 props 类型
  └─ index.ts             ← 公开导出
        ↓ 被调用
demo/DocxViewerPage.vue + DocxEditorPage.vue
```

## 集成关系枚举

每个"模块 A 调用/创建/注入模块 B"的关系，必须有集成任务验证真实调用路径（不是 mock/stub）。

### XLSX 集成点

| # | 调用方 → 被调方 | 接口 | 验证方法 |
|---|---|---|---|
| X1 | vue-xlsx/composables → xlsx-core/wasm | `getSheetsWasmModule()` + `Workbook.fromBytes` + `calculate` + `saveXlsxBytes` | Node 脚本：load xlsx → calculate → setFormula → save → openpyxl 验证 |
| X2 | vue-xlsx/composables → xlsx-core/worker-client | `XlsxWorkerClient.loadWorkbook` + `getCellSnapshot` + `getRowsBatch` | Node 脚本：worker 模式加载 → snapshot 查询 |
| X3 | vue-xlsx/composables → xlsx-core/images+charts | `parseWorkbookImageAssets` + `loadWorkbookChartAssets` + `mergeWorkbookImageAssets` | Node 脚本：加载含图片/图表的 xlsx → 验证 assets 非空 |
| X4 | vue-xlsx/XlsxViewer → vue-xlsx/composables | controller API（load/edit/export/undo/redo） | 浏览器：viewer 消费 controller 渲染单元格 + 编辑 + 导出 |
| X5 | vue-xlsx/XlsxViewer → vue-xlsx/chart-renderer | `MemoChartSvg` 渲染图表 | 浏览器：加载 charts-images.xlsx → 图表可见 |
| X6 | demo/main → xlsx-core/wasm | `setWasmSource("/duke_sheets_wasm_bg.wasm")` | 浏览器：wasm 加载成功 |
| X7 | demo/XlsxViewerPage → vue-xlsx | `<XlsxViewer :controller>` + `useXlsxViewerController` | 浏览器：页面渲染 + 交互 |

### DOCX 集成点

| # | 调用方 → 被调方 | 接口 | 验证方法 |
|---|---|---|---|
| D1 | vue-docx/composables → docx-core/wasm | `wasmBuildDocModelFromBytes` + `wasmSerializeDocx` | Node 脚本：load docx → model → serialize → python-docx 验证 |
| D2 | vue-docx/composables → docx-core/docx-import | `importDocxBuffer`（worker/主线程） | Node 脚本：worker 模式加载 → model 非空 |
| D3 | vue-docx/composables → docx-core/editor-helpers | `buildDocumentPageNodeSegments` + `estimateParagraphLineHeightPx` + `buildParagraphNumberingLabels` | Node 脚本：model → 分页 → pages 非空 |
| D4 | vue-docx/composables → docx-core/editor-ops | `updateParagraphText` + `splitParagraphChildrenAtTextOffsets` + `applyRunStyle` | Node 脚本：model → 编辑 → 验证 model 变化 |
| D5 | vue-docx/composables → docx-core/serializer | `serializeDocx(model, basePackage)` | Node 脚本：model → serialize → python-docx 验证段落/样式保留 |
| D6 | vue-docx/DocxEditor → vue-docx/composables | controller API（import/edit/export/undo/redo） | 浏览器：editor 消费 controller 渲染页面 + 编辑 + 导出 |
| D7 | vue-docx/DocxViewer → docx-core/layout-engine | `layoutDocument` | 浏览器：viewer 渲染分页 |
| D8 | demo/main → docx-core/wasm | `setWasmSource("/docx_wasm_bg.wasm")` | 浏览器：wasm 加载成功 |
| D9 | demo/DocxEditorPage → vue-docx | `<DocxEditorViewer :editor>` + `useDocxEditor` | 浏览器：页面渲染 + 交互 |

### 跨线集成点

| # | 调用方 → 被调方 | 接口 | 验证方法 |
|---|---|---|---|
| C1 | demo/main → xlsx-core + docx-core | 两个 `setWasmSource` | 浏览器：两种 wasm 都加载成功 |
| C2 | demo → vue-xlsx + vue-docx | 路由切换 | 浏览器：xlsx/docx 路由切换不崩溃 |

## 任务边界修正

### xlsx-001 + xlsx-002 合并问题

**原方案**：xlsx-001 复制 7 文件 + 创建占位 types.ts；xlsx-002 替换为完整 types.ts。

**问题**：xlsx-001 的 verify 会发现 types 不完整（只有占位），修复必然产生 xlsx-002 的交付物——按规范应合并。

**修正**：**合并为 xlsx-001**。一个任务完成：复制 7 文件 + 复制 types.ts + 清理 React 类型 + 安装依赖。types.ts 清理是 sed 脚本（几秒完成），不构成独立可验证交付物。

### docx-001 + docx-002 不合并

**理由**：docx-001 是 28 文件直接复制（零修改），docx-002 是 24953 行 sed 清理 + 类型抽出。两者工作模式完全不同，docx-001 可独立 typecheck（types.ts 来自 doc-model 包，零 React），docx-002 的清理是增量工作。修复 docx-001 的 blocking 不会产生 docx-002 的内容。

## 集成任务

基于集成关系枚举，需要 3 个集成任务：

| 任务 | 验证的集成点 | depends-on |
|---|---|---|
| xlsx-integ | X1-X3（composables → xlsx-core 真实调用） | xlsx-003 |
| docx-integ | D1-D5（composables → docx-core 真实调用） | docx-003 |
| cross-integ | X4-X7, D6-D9, C1-C2（viewer → controller, demo → 全链路） | xlsx-005, docx-004, docx-005 |

**为什么 xlsx-integ 不等 xlsx-005**：X1-X3 是 composables → xlsx-core 的调用，不需要 viewer。在 controller 改写完成后即可验证引擎调用路径是真实的（不是 stub）。

**为什么 cross-integ 在最后**：X4-X7 和 D6-D9 是 viewer → controller、demo → 全链路，需要 viewer 和 demo 都完成。

## 非任务（不在迁移范围内）

- vue-extend 组件（无上游，保留现有）
- PDF Viewer（无上游，按 shadcn 标准验收）
- Components 页面（无上游，按 shadcn 标准验收）
- Home 页面（不动）
