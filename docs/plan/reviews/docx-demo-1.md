# DOCX Demo 开发产出 Review #1

Date: 2026-07-08
Reviewer: automated review
Status: **blocked**

## Overview

对 `docx-demo` 任务开发产出进行全面 review，对照架构设计文档 (`docs/docx-migration-architecture.md`) 和上游对齐文档 (`docs/upstream-docx-feature-alignment.md`)。

## 检查结果总览

| 检查项 | 结果 | 说明 |
|---|---|---|
| 文件清单覆盖率 | ✅ pass | 所有架构文档定义的文件均存在，额外 8 个合理拆分文件 |
| Typecheck | ✅ pass | `docx-core` 和 `vue-docx` 均通过 `tsc --noEmit` |
| Build | ✅ pass | 两个包均构建成功 |
| React import 残留 | ✅ pass | 零 React import，所有 `@extend-ai/react-docx` 引用仅限注释 |
| Import 路径 | ✅ pass | 全部使用相对路径，指向正确 |
| Stub/mock/fake | ✅ pass | 无空函数体 stub，所有 composable 有完整实现 |
| 单文件 ≤1000 行 | ✅ pass | 仅 `selection-helpers.ts` 恰好 1000 行，其余均在限额内 |
| 上游 26 项功能对齐 | ⚠️ blocked | 仅 11/26 通过，15 项标记 ⚠️ |

## 一、文件清单对照

### 1.1 docx-core/engine/ (9 文件)

| 架构设计预期 | 实际 | 状态 |
|---|---|---|
| `types.ts` | `engine/types.ts` | ✅ |
| `clone.ts` | `engine/clone.ts` | ✅ |
| `normalize.ts` | `engine/normalize.ts` | ✅ |
| `wasm.ts` | `engine/wasm.ts` | ✅ |
| `ooxml-core.ts` | `engine/ooxml-core.ts` | ✅ |
| `serializer.ts` | `engine/serializer.ts` | ✅ |
| `doc-model.ts` | `engine/doc-model.ts` | ✅ |
| `index.ts` | `engine/index.ts` | ✅ |
| `generated/*.d.ts` | `engine/generated/docx_wasm.d.ts` + `docx_wasm_bg.wasm.d.ts` | ✅ |

### 1.2 docx-core/layout/ (5 文件)

| 架构设计预期 | 实际 | 状态 |
|---|---|---|
| `layout-engine.ts` | ✅ 359 行 | ✅ |
| `pagination.ts` | ✅ 689 行 | ✅ |
| `page-segmentation-core.ts` | ✅ | ✅ |
| `page-segmentation-table.ts` | ✅ | ✅ |
| `index.ts` | ✅ | ✅ |

### 1.3 docx-core/viewer/ (16 文件)

架构设计预期 15 个文件，实际 16 个。额外文件 `pretext-items-layout.ts` 是多 item 布局的合理拆分。

`image-render.ts` - EMF/WMF 占位使用 `imageUsesPlaceholderFallback` 机制，属于上游已有的域逻辑，非 stub。

### 1.4 docx-core/canvas/ (2 文件)

`types.ts` + `layout-diagnostics.ts` — 完全对齐。

### 1.5 docx-core/editor/ (editor-ops 拆分为 4 文件 + helpers 41 文件)

**editor-ops 拆分**：
- `paragraph-ops.ts` ✅
- `run-style-ops.ts` ✅
- `table-ops.ts` ✅
- `index.ts` ✅

**helpers/** (架构预期 ~39 文件，实际 41 文件)：

架构文档列出的 40 个文件全部存在。额外 4 个合理拆分：
- `line-height-table-extra.ts` — 行高估算额外逻辑
- `line-height-wrap.ts` — wrap 相关行高计算
- `paragraph-geometry-image.ts` — 图片几何独立拆分
- `paragraph-toc.ts` — TOC 独立拆分

`selection-helpers.ts` 恰好 1000 行，达到硬约束上限。如需后续迭代，建议优先从此文件继续拆分。

### 1.6 vue-docx/composables/ (28 文件)

架构预期 26 文件，实际 28 文件。额外文件：
- `editor-shared.ts` (91 行) — 共享 `EditorCore` 类型定义，合理的跨模块类型抽象
- `composables.ts` (6 行) — 向后兼容 barrel re-export，指向 `./composables/index`

**关键 composable 实现验证**：
- `editor-history.ts` (60 行) — 完整快照式 undo/redo，含 historyRestoreRequest nonce
- `editor-selection.ts` (54 行) — 完整 selection session 管理
- `editor-list.ts` (69 行) — 完整 list toggle + depth adjustment
- `editor-transaction.ts` (130 行) — 完整 dispatchEditorTransaction 唯一修改入口
- `editor-text-input.ts` (607 行) — 完整 contentEditable + draft 缓存
- `editor-format.ts` (337 行) — 完整 applySelectedStyleChange

无 stub 实现。

### 1.7 vue-docx/components/ (18 文件)

架构预期 18 文件，完全对齐。所有 Vue 组件均有完整模板 + script + style。

### 1.8 vue-docx/render/ (7 文件)

架构预期 6 文件，实际 7 文件。额外文件 `paragraph-runs-field-resolve.ts` 是表单域解析逻辑的合理拆分。

`paragraph-runs.ts` 931 行，在限额内。

### 1.9 Demo 页面

- `apps/demo/src/pages/DocxEditorPage.vue` (234 行) — editor demo，含 26 项 self-assessment table
- `apps/demo/src/pages/DocxViewerPage.vue` (130 行) — viewer demo，含 sample 加载

## 二、上游功能对齐

### 2.1 引擎层 (8/8)

| # | 对齐项 | 来源 | 状态 |
|---|---|---|---|
| 1 | WASM 单例懒加载 + Node/浏览器双路径 | `engine/wasm.ts` | ✅ |
| 2 | `setWasmSource` 覆盖 + SIMD 报错增强 | `engine/wasm.ts` | ✅ |
| 3 | DOCX 引擎无状态函数式 | `engine/wasm.ts` | ✅ |
| 4 | `basePackage` 机制 | `engine/serializer.ts` | ✅ |
| 5 | DocModel 类型体系 | `engine/types.ts` | ✅ |
| 6 | 手工深拷贝 cloneDocModel | `engine/clone.ts` | ✅ |
| 7 | JSON 归一化 Uint8Array 修复 | `engine/normalize.ts` | ✅ |
| 8 | OoxmlPackage Map 表示 + 最小包构造 | `engine/ooxml-core.ts` | ✅ |

### 2.2 布局/分页 (15/15)

| # | 对齐项 | 来源 | 状态 |
|---|---|---|---|
| 9 | 测量与切分解耦 | `layout/page-segmentation-core.ts` | ✅ |
| 10 | 段落跨页按行切分 | `layout/page-segmentation-core.ts` | ✅ |
| 11 | 表格行跨页 TableRowRange/TableRowSlice | `layout/page-segmentation-table.ts` | ✅ |
| 12 | keepNext 链 | `layout/page-segmentation-core.ts` | ✅ |
| 13 | widow/orphan | `layout/page-segmentation-core.ts` | ✅ |
| 14 | margin collapse | `layout/page-segmentation-core.ts` | ✅ |
| 15 | 表格孤行保护 | `layout/page-segmentation-table.ts` | ✅ |
| 16 | 硬分页符识别 + sourceXml 缓存 | `layout/pagination.ts` | ✅ |
| 17 | section header/footer 继承 | `layout/pagination.ts` | ✅ |
| 18 | 页首 spacing 抑制 | `layout/pagination.ts` | ✅ |
| 19 | 表格内分页符 | `layout/pagination.ts` | ✅ |
| 20 | lastRenderedPageBreak 提示 | `layout/pagination.ts` | ✅ |
| 21 | 页数校准 | `viewer/page-count-reconciliation.ts` | ✅ |
| 22 | 测量驱动迭代分页 | `editor/helpers/pagination-plan-iterate.ts` | ✅ |
| 23 | pretext 变宽文本布局 | `viewer/pretext-layout.ts` + `pretext-selection.ts` + `pretext-items-layout.ts` | ✅ |

### 2.3 编辑 (11/11)

| # | 对齐项 | 来源 | 状态 |
|---|---|---|---|
| 24 | dispatchEditorTransaction 唯一修改入口 | `composables/editor-transaction.ts` | ✅ |
| 25 | 不可变模型 cloneDocModel | 全局使用 | ✅ |
| 26 | 快照式历史上限 100 | `composables/editor-history.ts` | ✅ |
| 27 | pendingRunStyle | `composables/useDocxEditor.ts` | ✅ |
| 28 | contentEditable + draft 缓存 | `composables/editor-text-input.ts` | ✅ |
| 29 | applySelectedStyleChange | `composables/editor-format.ts` | ✅ |
| 30 | historyRestoreRequest nonce | `composables/editor-history.ts` | ✅ |
| 31 | selectionSession | `composables/editor-selection.ts` | ✅ |
| 32 | 修订/评论只读派生 | `composables/useDocxTrackChanges.ts` + `useDocxComments.ts` | ✅ |
| 33 | 导入 wasm worker + 主线程 fallback | `composables/editor-import-export.ts` → `viewer/docx-import.ts` | ✅ |
| 34 | 导出 serializeDocx + basePackage | `composables/editor-import-export.ts` | ✅ |

### 2.4 渲染 (9/9)

| # | 对齐项 | 来源 | 状态 |
|---|---|---|---|
| 35 | 纯 DOM 渲染 | `components/DocxPageBody.ts` | ✅ |
| 36 | 虚拟化 | `components/DocxViewerRoot.vue` | ✅ |
| 37 | contain layout style 隔离 | `components/DocxPageBody.ts` | ✅ |
| 38 | v-html 渲染段落 runs | `components/DocxParagraphHost.vue` | ✅ |
| 39 | 浮动图片 absolute 定位 | `components/DocxImageLayer.vue` | ✅ |
| 40 | 表格 resize handle | `components/DocxTableHost.vue` | ✅ |
| 41 | 缩略图三级策略 | `components/DocxThumbnailPanel.vue` + `viewer/thumbnail-raster.ts` + `viewer/thumbnail-cache.ts` | ✅ |
| 42 | page surface registry 跨组件共享 | `composables/page-surface-registry.ts` | ✅ |
| 43 | CSSProperties → Record 类型替换 | 全量 sed 替换，零 React.CSSProperties | ✅ |

### 2.5 加载/导入 (4/4)

| # | 对齐项 | 来源 | 状态 |
|---|---|---|---|
| 44 | importDocxBuffer worker/主线程双路径 + AbortSignal | `viewer/docx-import.ts` | ✅ |
| 45 | worker 消息协议 | `viewer/docx-import-worker.ts` | ✅ |
| 46 | 嵌入字体加载 | `viewer/docx-import.ts` (loadEmbeddedFonts) | ✅ |
| 47 | 默认 starter model | `editor/helpers/default-model.ts` | ✅ |

**底层引擎对齐度**: 47/47 (100%)

### 2.6 Toolbar UI 功能 (11/26)

Demo 页面 `DocxEditorPage.vue` 内嵌 self-assessment table，以下按上游 26 项功能逐项对照：

| # | 功能 | Toolbar 实现 | Demo 自评 | 状态 |
|---|---|---|---|---|
| 1 | Document theme light/dark | 🌙/☀️ toggle button | ✅ | ✅ |
| 2 | App theme light/dark/system | - | ✅ | ✅ |
| 3 | Undo/Redo | ↩/↪ buttons + ⌘Z/⇧⌘Z | ✅ | ✅ |
| 4 | Paragraph styles | `<select>` dropdown | ✅ | ✅ |
| 5 | Font family | **缺失** | ⚠️ | ⚠️ |
| 6 | Font size 8–48pt | **缺失** | ⚠️ | ⚠️ |
| 7 | Line spacing | **缺失** | ⚠️ | ⚠️ |
| 8 | Bold/Italic/Underline/Strikethrough | B/I/U/S buttons | ✅ | ✅ |
| 9 | Superscript/Subscript | x²/x₂ buttons | ✅ | ✅ |
| 10 | Text color + highlight color | **缺失** | ⚠️ | ⚠️ |
| 11 | Hyperlink | **缺失** | ⚠️ | ⚠️ |
| 12 | Alignment L/C/R/J | ⬅/⬌/➡/☰ buttons | ✅ | ✅ |
| 13 | Bullet/Numbered lists | •/1. buttons | ✅ | ✅ |
| 14 | Columns display | **缺失** | ⚠️ | ⚠️ |
| 15 | Page thumbnails | Toggle via demo page checkbox | ⚠️ | ⚠️* |
| 16 | Borders 13 presets | **缺失** | ⚠️ | ⚠️ |
| 17 | Insert image | **缺失** | ⚠️ | ⚠️ |
| 18 | Insert table | **缺失** | ⚠️ | ⚠️ |
| 19 | Zoom 50–200% | **缺失** | ⚠️ | ⚠️ |
| 20 | Import .docx | 📂 button + demo page button | ✅ | ✅ |
| 21 | Export .docx | 💾 button + demo page button + ⌘S | ✅ | ✅ |
| 22 | Track changes toggle | Demo page checkbox | ⚠️ | ⚠️* |
| 23 | Comments toggle | Demo page checkbox | ⚠️ | ⚠️* |
| 24 | Read-only mode toggle | Demo page checkbox | ✅ | ✅ |
| 25 | Context menu | `DocxContextMenu.vue` 存在 | ⚠️ | ⚠️* |
| 26 | Form field config | **缺失** (backend `useDocxFormFields` 存在) | ⚠️ | ⚠️ |

\* 标注 ⚠️\* 的项目后端 composable 已实现但 toolbar 无直接 UI 入口（如 thumbnails/track changes/comments 通过 demo page checkbox 控制；context menu 组件存在但功能可能不完整）。

**核心问题**：Toolbar 缺失 9 个上游 button/selector（font family, font size, line spacing, text color, highlight color, hyperlink, borders, insert image, insert table, zoom, columns）。这些功能对应的 composable 大部分已实现（`useDocxLineSpacing`, `useDocxBorders`, `useDocxImageWrapMenu`, `useDocxFormFields` 等），只是 toolbar 未接线。

## 三、Import 路径检查

全部通过。所有跨模块 import 使用正确相对路径：
- `../engine/types`, `../engine/clone`, `../engine/wasm` 等
- `../../viewer/section-layout`, `../../layout/pagination` 等
- `../composables/useDocxEditor` 等
- `@extend-ai/docx-core` 用于 vue-docx 引用 docx-core（正确的包间依赖）
- 零 `@extend-ai/react-docx-*` import 语句（仅注释中出现）

## 四、Stub/Mock 检查

无 stub/mock/fake 代码。所有 composable 和 helper 函数均有完整实现逻辑。`imageUsesPlaceholderFallback` 是上游已有的 EMF/WMF 占位渲染机制，属于域逻辑。

`composables.ts` (6 行 barrel re-export) 标记为 legacy compatibility file，非 stub。

## 五、问题清单

### Blocking Issues

1. **Toolbar 功能缺失 15/26**: Demo 页自评 15 项 ⚠️。架构硬约束要求"功能对齐上游 26 项能力"，当前 toolbar 仅覆盖核心编辑操作，缺失字体/字号/行距/颜色/高亮/超链接/边框/图文插入/缩放/分栏等格式化能力。虽然后端 composable 已就位，但缺少用户可操作的 UI 入口。

   **修复建议**：在 `DocxToolbar.vue` 中补齐：
   - Font family `<select>` (Calibri/Arial/TNR/Georgia/Helvetica/Courier)
   - Font size `<select>` (8-48pt)
   - Line spacing `<select>` (1/1.15/1.2/1.5/2/2.5/3)
   - Text color `<input type="color">` + highlight color picker
   - Hyperlink button + dialog
   - Borders dropdown (13 presets)
   - Insert image + Insert table buttons
   - Zoom slider (50-200%)
   - Columns display toggle
   - Show edits / Show comments toggles

2. **`selection-helpers.ts` 达到 1000 行硬约束上限**: 恰好 1000 行，无缓冲余量。后续任何增量改动都会突破约束。

### Non-blocking Observations

3. **`composables.ts` legacy barrel**: 6 行 re-export 指向 `./composables/index`，标记为 backward compatibility。建议在确认无外部依赖后删除。

4. **`pageSegmentationPlan` 函数名变更**: 上游 `editor.tsx` 中的 `pageSegmentationPlan` 在迁移后拆分为 `pagination-plan-core.ts` + `pagination-plan-iterate.ts` 中的多个独立函数（`buildRenderColumnSegmentsForPageSection`, `buildDocumentPageNodeSegmentsFromLastRenderedPageBreakHints` 等），无单一同名入口。功能等价但需确认调用方已适配。

5. **额外文件 8 个**: 架构设计预期 ~124 文件，实际 ~137 文件。额外文件均为合理的细粒度拆分，不违反架构原则。

## 六、统计数据

| 维度 | 预期 | 实际 | 状态 |
|---|---|---|---|
| docx-core 文件数 | ~78 | 83 | 合理扩展 |
| vue-docx 文件数 | ~50 | 55 | 合理扩展 |
| Typecheck docx-core | 0 errors | 0 errors | ✅ |
| Typecheck vue-docx | 0 errors | 0 errors | ✅ |
| Build docx-core | 成功 | 646.97 KB ESM | ✅ |
| Build vue-docx | 成功 | 174.73 KB ESM | ✅ |
| 底层引擎对齐 | 47/47 | 47/47 | ✅ |
| Toolbar UI 对齐 | 26/26 | 11/26 | ⚠️ blocked |

## 七、结论

底层架构实现质量高：引擎层/布局层/viewer/editor helpers 共计 47 项底层能力全部对齐上游，typecheck 和 build 通过，文件结构清晰且全部满足 ≤1000 行约束。

阻塞点集中在 Toolbar UI 层：15 项上游功能有后端支撑但缺少 toolbar 入口，导致 demo 页 26 项 feature checklist 仅 11 项通过。补齐 toolbar 按钮/选择器是当前唯一阻塞项。
