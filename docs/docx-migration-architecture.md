# DOCX 迁移架构设计

Date: 2026-07-07
Status: 设计阶段,待确认后执行

## 硬约束

- **单文件 ≤ 1000 行**(含注释,不含 generated/wasm 二进制)
- 模块化,按领域分文件
- 功能对齐上游 26 项能力
- 通过第九节验证清单
- 仅限 DOCX,不动 xlsx/pdf

## 一、整体分层

```
packages/docx-core/src/     ← 框架无关引擎+布局+编辑逻辑
├── engine/                 ← 引擎层(已完成 ✅)
├── layout/                 ← 布局/分页层
├── editor/                 ← 编辑操作 + editor-helpers(纯函数)
│   └── helpers/            ← editor.tsx 前 24953 行拆分(细粒度)
├── viewer/                 ← 文本布局/缩略图/导入等 viewer 辅助
└── canvas/                 ← canvas 诊断/类型

packages/vue-docx/src/      ← Vue 组件 + composables
├── composables/            ← useDocxEditor 等 hooks(模块化)
├── components/             ← DocxViewer.vue / DocxEditor.vue 等子组件
└── render/                 ← renderParagraphRuns 等 Vue render 函数
```

## 二、docx-core 模块设计(全部 ≤1000 行)

### 2.1 engine/(已完成 ✅)

| 文件 | 行数 | 状态 |
|---|---:|---|
| types.ts | 438 | ✅ |
| clone.ts | 406 | ✅ |
| normalize.ts | 102 | ✅ |
| wasm.ts | 211 | ✅ |
| ooxml-core.ts | 88 | ✅ |
| serializer.ts | 32 | ✅ |
| doc-model.ts | 35 | ✅ |
| index.ts | barrel | ✅ |
| generated/ | 二进制 | ✅ |

### 2.2 layout/(布局层,拆成 5 文件)

上游 layout-core/page-segmentation.ts(1223行)拆成 2 个:

| 文件 | ~行数 | 内容 |
|---|---:|---|
| layout-engine.ts | 359 | layoutDocument 装箱式布局 |
| pagination.ts | 689 | 硬分页符/section/header-footer 继承 |
| page-segmentation-core.ts | ~650 | buildDocumentPageNodeSegments 核心 + keepNext/widow |
| page-segmentation-table.ts | ~573 | 表格行跨页切分(TableRowRange/Slice) |
| index.ts | 382 | barrel + DEFAULT_LAYOUT_OPTIONS |
| **合计** | **~2653** | |

拆分点:page-segmentation.ts 按段落切分 vs 表格切分分两个文件。

### 2.3 viewer/(辅助模块,拆成 13 文件)

pretext-layout.ts(1389行)和 thumbnail-raster.ts(1239行)超 1000 行,拆分:

| 文件 | ~行数 | 内容 |
|---|---:|---|
| pretext-layout.ts | ~700 | layoutTextWithPretextAroundExclusions 等核心布局 |
| pretext-selection.ts | ~689 | resolveOffsetAtPoint/caretRect/selectionRects 命中测试 |
| thumbnail-raster.ts | ~650 | 快照直绘 + DOM→SVG 光栅化 |
| thumbnail-cache.ts | ~589 | DocxThumbnailSurfaceCache LRU + SerialIdleTaskQueue |
| docx-import.ts | 238 | importDocxBuffer worker/主线程双路径 |
| docx-import-worker.ts | 75 | worker 消息处理 |
| layout-snapshot.ts | 469 | 布局快照 + 选区/光标操作 |
| section-layout.ts | 298 | sectPr 解析 + twipsToPixels |
| pagination-breaks.ts | 225 | 显式分页符检测 |
| page-count-reconciliation.ts | 278 | 页数校准 |
| image-render.ts | 260 | TIFF→PNG + EMF/WMF 占位 |
| content-signature.ts | 155 | FNV-1a 签名 |
| wasm-source.ts | 69 | wasm 源配置 |
| utif.d.ts | 18 | 类型声明 |
| index.ts | barrel | - |
| **合计** | **~4713** | |

拆分点:
- pretext-layout:文本布局算法 vs 命中测试(点击/光标/选区)
- thumbnail-raster:光栅化绘制 vs 缓存队列管理

### 2.4 canvas/(2 文件)

| 文件 | 行数 |
|---|---:|
| types.ts | 44 |
| layout-diagnostics.ts | 156 |

### 2.5 editor/(编辑层,细粒度拆分)

#### editor-ops.ts 拆分(1329行 → 3 文件)

| 文件 | ~行数 | 内容 |
|---|---:|---|
| paragraph-ops.ts | ~480 | splitParagraphChildrenAtTextOffsets/updateParagraphText |
| run-style-ops.ts | ~450 | mutateParagraphTextStyleInRange 等 run 样式操作 |
| table-ops.ts | ~399 | 表格行列操作 |
| index.ts | barrel | - |

#### editor/helpers/(editor.tsx 前 24953 行,细粒度 30+ 文件)

按分析员 24 模块方案,大模块再拆:

| 文件 | ~行数 | 领域 |
|---|---:|---|
| constants.ts | 480 | 全局常量 |
| performance.ts | 40 | markDocxImportPerformance |
| editor-types.ts | ~760 | 公共 API 类型(拆成 types-1 + types-2) |
| editor-types-extra.ts | ~760 | 公共 API 类型(续) |
| cache-utils.ts | 160 | 缓存工具 |
| ooxml-helpers.ts | 205 | XML 属性解析 |
| dom-helpers.ts | 250 | DOM 操作(scheduleDomWrite/placeCaret 等) |
| zoom-utils.ts | 60 | 缩放/滚动工具 |
| header-footer.ts | 625 | 页眉页脚 reserve |
| page-measurement.ts | 740 | 页面高度测量 |
| default-model.ts | 100 | defaultStarterModel |
| letterhead.ts | 215 | 信头布局 |
| paragraph-inspect.ts | ~800 | 段落属性提取(部分) |
| paragraph-geometry.ts | ~800 | 段落几何/wrap 计算 |
| paragraph-tracked.ts | ~750 | 段落修订标记提取 |
| pretext-build.ts | ~900 | buildParagraphPretextLayoutSource |
| pretext-measure.ts | ~900 | pretext 测量集成 |
| drop-cap.ts | 450 | 首字下沉 |
| line-height.ts | ~740 | estimateParagraphLineHeightPx(部分) |
| line-height-table.ts | ~730 | 表格行高估算 |
| table-height.ts | 510 | estimateTableRowHeightsPx |
| pagination-plan-core.ts | ~850 | pageSegmentationPlan 编排核心 |
| pagination-plan-iterate.ts | ~850 | 测量驱动迭代分页 |
| pagination-plan-stabilize.ts | ~800 | 分页稳定化/振荡检测 |
| style-to-css.ts | ~580 | runStyleToCss(部分) |
| style-block-css.ts | ~580 | paragraphBlockStyle 等 |
| xml-parsing.ts | ~560 | XML 解析(部分) |
| xml-parsing-extra.ts | ~560 | XML 解析(续) |
| synthetic-textbox.ts | 750 | 合成文本框 |
| tracked-changes.ts | ~550 | 修订收集(部分) |
| tracked-changes-gutter.ts | ~550 | 修订 gutter 卡片 |
| field-helpers.ts | 700 | 表单域 |
| numbering.ts | 600 | 列表编号 |
| table-utils.ts | ~700 | 表格工具(部分) |
| table-utils-extra.ts | ~705 | 表格工具(续) |
| text-mutation.ts | 760 | 文本变更 |
| selection-helpers.ts | ~605 | 选区/光标(部分) |
| selection-restore.ts | ~605 | DOM 选区恢复 |
| section-manipulation.ts | 310 | section 操作 |
| state.ts | 227 | 编辑器状态机(core/state.ts) |
| index.ts | barrel | - |
| **合计** | **~21200** | |

### 2.6 docx-core/src/index.ts

聚合导出所有层。

## 三、vue-docx 模块设计(全部 ≤1000 行)

### 3.1 composables/(hooks 改写,拆成 ~20 文件)

useDocxEditor(~4600行)是最大块,按职责拆:

| 文件 | ~行数 | 内容 |
|---|---:|---|
| useDocxEditor.ts | ~700 | 核心:model/selection/history state + dispatchEditorTransaction |
| editor-transaction.ts | ~800 | dispatchEditorTransaction 事务分发逻辑 |
| editor-history.ts | ~500 | 快照式历史(undo/redo,上限100) |
| editor-selection.ts | ~700 | 选区管理 + selectionSession + historyRestoreRequest |
| editor-text-input.ts | ~700 | contentEditable + draft 缓存 + commitParagraphText |
| editor-format.ts | ~600 | applySelectedStyleChange + pendingRunStyle |
| editor-table.ts | ~600 | 表格 insert/delete/resize |
| editor-image.ts | ~600 | 图片 insert/resize/move/wrap |
| editor-form-field.ts | ~400 | 表单域 |
| editor-list.ts | ~400 | 列表 toggleList/adjustDepth |
| editor-clipboard.ts | ~300 | 复制/粘贴 |
| editor-import-export.ts | ~500 | importDocxFile/exportDocx |
| page-surface-registry.ts | ~400 | DocxViewerPageSurfaceRegistry |
| useDocxPageThumbnails.ts | ~700 | 缩略图 |
| useDocxDocumentTheme.ts | 100 | 文档主题 |
| useDocxParagraphStyles.ts | 100 | 段落样式 |
| useDocxImageWrapMenu.ts | 200 | 图片 wrap 菜单 |
| useDocxLineSpacing.ts | 100 | 行距 |
| useDocxBorders.ts | 100 | 边框 |
| useDocxFormFields.ts | 200 | 表单域 |
| useDocxTrackChanges.ts | 200 | 修订 |
| useDocxComments.ts | 200 | 批注 |
| useDocxPageLayout.ts | 200 | 页面布局 |
| useDocxPagination.ts | 200 | 分页 |
| useDocxModel.ts | 200 | useDocxModel hook(index.tsx) |
| index.ts | barrel | - |
| **合计** | **~8700** | |

### 3.2 components/(Vue 组件,拆成 ~15 文件)

DocxEditorViewer(24000行)拆成多个子组件 + composables:

| 文件 | ~行数 | 内容 |
|---|---:|---|
| DocxViewer.vue | ~800 | 只读 viewer(index.tsx ReactDocxViewer) |
| DocxEditor.vue | ~800 | 编辑 viewer 容器(DocxEditorViewer 顶层) |
| DocxViewerRoot.vue | ~700 | viewer root + 虚拟化 spacer |
| DocxPageWrapper.vue | ~600 | page wrapper + page surface 容器 |
| DocxPageSurface.vue | ~800 | 单页表面(cover/border/header/body/footer/gutter) |
| DocxPageHeader.vue | ~400 | 页眉渲染 |
| DocxPageFooter.vue | ~400 | 页脚渲染 |
| DocxPageBody.ts | ~800 | body 段落/表格渲染(render function) |
| DocxParagraphHost.vue | ~700 | 段落 host(contentEditable + draft) |
| DocxTableHost.vue | ~700 | 表格渲染 + resize handle |
| DocxImageLayer.vue | ~500 | 浮动图片层 |
| DocxFormFieldLayer.vue | ~400 | 表单域层 |
| DocxTrackedChangeGutter.vue | ~400 | 修订/批注 gutter |
| DocxContextMenu.vue | ~500 | 右键菜单 |
| DocxToolbar.vue | ~800 | 工具栏(26 项功能) |
| DocxThumbnailPanel.vue | ~600 | 缩略图面板 |
| DocxDragOverlay.vue | ~300 | 拖拽覆盖层 |
| index.ts | barrel | - |
| **合计** | **~10900** | |

注:viewer 区 24000 行里有大量是 hooks(useRef/useCallback),已拆到 composables/。组件本身聚焦模板+渲染。

### 3.3 render/(renderParagraphRuns 重写,拆成 4 文件)

| 文件 | ~行数 | 内容 |
|---|---:|---|
| paragraph-runs.ts | ~500 | renderParagraphRuns 主函数(h() 生成 VNode) |
| paragraph-runs-text.ts | ~350 | 文本 run 渲染(含 tab/hyperlink/note) |
| paragraph-runs-image.ts | ~300 | 图片 run 渲染 |
| paragraph-runs-field.ts | ~305 | 表单域 run 渲染 |
| static-html.ts | ~50 | renderStaticHtml |
| index.ts | barrel | - |
| **合计** | **~1505** | |

## 四、文件总数统计

| 层 | 文件数 | ~行数 |
|---|---:|---:|
| docx-core/engine | 8 | 1700 |
| docx-core/layout | 5 | 2653 |
| docx-core/viewer | 15 | 4713 |
| docx-core/canvas | 2 | 200 |
| docx-core/editor/helpers | 39 | 21200 |
| docx-core/editor(state+ops) | 5 | 1556 |
| vue-docx/composables | 26 | 8700 |
| vue-docx/components | 18 | 10900 |
| vue-docx/render | 6 | 1505 |
| **合计** | **~124** | **~53127** |

所有文件 ≤1000 行。

## 五、依赖关系图

```
engine/types ← 基础
    ↑
engine/{clone,normalize,wasm,ooxml-core,serializer,doc-model}
    ↑
layout/{layout-engine,pagination,page-segmentation-*}
    ↑
viewer/{pretext-*,thumbnail-*,docx-import,...}
    ↑
editor/helpers/* (30+ 模块)
    ↑
vue-docx/composables/* (useDocxEditor 等)
    ↑
vue-docx/render/* + vue-docx/components/*
```

## 六、迁移执行顺序

| 步骤 | 内容 | 文件数 | 验证 |
|---|---|---:|---|
| 1 ✅ | engine/ 引擎层 | 8 | wasm smoke + python-docx |
| 2 | layout/ 布局层 | 5 | typecheck + build |
| 3 | canvas/ canvas 层 | 2 | typecheck |
| 4 | viewer/ 辅助模块 | 15 | typecheck + build |
| 5 | editor/helpers/ 30+ 模块 | 39 | typecheck(分批 5-6 个) |
| 6 | editor/state + ops | 5 | typecheck |
| 7 | composables/ hooks | 26 | typecheck + build |
| 8 | render/ renderParagraphRuns | 6 | typecheck |
| 9 | components/ Vue 组件 | 18 | build + 浏览器 |
| 10 | demo + 依赖 + 构建配置 | - | 全量验证 |

## 七、关键技术决策

### 7.1 WASM
- `public/docx_wasm_bg.wasm` + `setWasmSource("/docx_wasm_bg.wasm")`
- tsup onSuccess 复制 wasm(已实现 ✅)

### 7.2 Worker
- docx-import-worker.ts 作为 tsup 额外 entry
- Vite `new URL("./worker.ts", import.meta.url)`

### 7.3 虚拟化
- `@tanstack/vue-virtual`,overscan=2,before/after spacer

### 7.4 contentEditable
- draft 缓存(paragraphDraftsRef 等)→ Vue ref
- nonce 驱动选区恢复
- selectionSession 抢占机制
- v-html + suppress 避免响应式破坏输入

### 7.5 editor-helpers 拆分
- 从 editor.tsx 按行号范围提取
- 每模块独立 typecheck
- 大模块(pagination-plan/paragraph-inspection/pretext-integration)按子领域再拆

### 7.6 renderParagraphRuns
- Vue `h()` 生成 VNode,不用 JSX
- 按 run 类型分文件(text/image/form-field)

## 八、风险

| 风险 | 缓解 |
|---|---|
| 30+ helpers 拆分行号交错 | 分批 typecheck,每批 5-6 模块 |
| composables 机械改写量大 | agent 并行,按职责拆分降低单文件复杂度 |
| DocxEditor 组件拆分 | 先 DocxViewer(只读)再 DocxEditor,子组件独立验证 |
| contentEditable + Vue 响应式 | 严格复刻 draft + suppress + nonce |
| @chenglou/pretext 版本 | 锁定 0.0.8,验证 API |
| 单文件 ≤1000 行限制增加文件数 | 接受 ~124 文件,换取可维护性 |

## 九、与原方案的差异

| 原方案 | 本设计 |
|---|---|
| editor-helpers.ts 24953 行 | 39 个模块,每个 ≤900 行 |
| composables.ts 单文件 | 26 个 composable |
| editor.tsx viewer 区 24000 行 | 18 个 Vue 组件 |
| renderParagraphRuns 单函数 | 4 个文件 |
| ~68000 行 | ~53127 行(去重+清理) |

功能范围、验证清单、License 完全一致。
