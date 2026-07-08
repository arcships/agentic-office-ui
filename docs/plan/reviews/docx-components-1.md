# docx-components 开发产出 Review

Date: 2026-07-08
Reviewer: DimCode
Status: **BLOCKED** — 3 项硬约束违规需修复

## 一、总体评估

开发产出完整覆盖了架构文档定义的全部模块层级（engine → layout → viewer → canvas → editor/helpers → composables → render → components），typecheck 零错误，无 React 残留，无 stub/mock/fake 代码。26 项上游功能均有对应实现文件。但存在 3 项违反"单文件 ≤1000 行"硬约束的违规。

---

## 二、文件清单对照

### 2.1 engine/ ✅ 完全对齐

| 架构文档 | 实际文件 | 行数 | 状态 |
|---|---|---|---|
| types.ts | engine/types.ts | 438 | ✅ |
| clone.ts | engine/clone.ts | 406 | ✅ |
| normalize.ts | engine/normalize.ts | 102 | ✅ |
| wasm.ts | engine/wasm.ts | 211 | ✅ |
| ooxml-core.ts | engine/ooxml-core.ts | 88 | ✅ |
| serializer.ts | engine/serializer.ts | 32 | ✅ |
| doc-model.ts | engine/doc-model.ts | 35 | ✅ |
| index.ts | engine/index.ts | barrel | ✅ |
| generated/ | engine/generated/ (2 文件) | - | ✅ |

### 2.2 layout/ ✅ 完全对齐

| 架构文档 | 实际文件 | 行数 | 状态 |
|---|---|---|---|
| layout-engine.ts | layout/layout-engine.ts | 359 | ✅ |
| pagination.ts | layout/pagination.ts | 690 | ✅ |
| page-segmentation-core.ts | layout/page-segmentation-core.ts | 864 | ✅ |
| page-segmentation-table.ts | layout/page-segmentation-table.ts | 376 | ✅ |
| index.ts | layout/index.ts | 393 | ✅ |

### 2.3 viewer/ ✅ 完全对齐 + 1 额外文件

| 架构文档 | 实际文件 | 行数 | 状态 |
|---|---|---|---|
| pretext-layout.ts | viewer/pretext-layout.ts | 679 | ✅ |
| pretext-selection.ts | viewer/pretext-selection.ts | 237 | ✅ |
| thumbnail-raster.ts | viewer/thumbnail-raster.ts | 941 | ✅ |
| thumbnail-cache.ts | viewer/thumbnail-cache.ts | 295 | ✅ |
| docx-import.ts | viewer/docx-import.ts | 237 | ✅ |
| docx-import-worker.ts | viewer/docx-import-worker.ts | 75 | ✅ |
| layout-snapshot.ts | viewer/layout-snapshot.ts | 468 | ✅ |
| section-layout.ts | viewer/section-layout.ts | 297 | ✅ |
| pagination-breaks.ts | viewer/pagination-breaks.ts | 224 | ✅ |
| page-count-reconciliation.ts | viewer/page-count-reconciliation.ts | 277 | ✅ |
| image-render.ts | viewer/image-render.ts | 266 | ✅ |
| content-signature.ts | viewer/content-signature.ts | 155 | ✅ |
| wasm-source.ts | viewer/wasm-source.ts | 69 | ✅ |
| utif.d.ts | viewer/utif.d.ts | 18 | ✅ |
| index.ts | viewer/index.ts | barrel | ✅ |
| *(无)* | viewer/pretext-items-layout.ts | 471 | ➕ 额外拆分 |

`pretext-items-layout.ts` 是 pretext-layout.ts 的进一步拆分（内联混合字体文本布局），属于合理的子模块细分。

### 2.4 canvas/ ✅ 完全对齐

| 架构文档 | 实际文件 | 行数 | 状态 |
|---|---|---|---|
| types.ts | canvas/types.ts | 44 | ✅ |
| layout-diagnostics.ts | canvas/layout-diagnostics.ts | 156 | ✅ |

### 2.5 editor/ ⚠️ editor-ops 未拆分 + helpers 有局部差异

#### editor-ops.ts ❌ 未按架构文档拆分

| 架构文档（应拆为 4 文件） | 实际 | 行数 | 状态 |
|---|---|---|---|
| paragraph-ops.ts (~480) | *(合并于 editor-ops.ts)* | - | ❌ |
| run-style-ops.ts (~450) | *(合并于 editor-ops.ts)* | - | ❌ |
| table-ops.ts (~399) | *(合并于 editor-ops.ts)* | - | ❌ |
| index.ts (barrel) | *(合并于 editor-ops.ts)* | - | ❌ |
| *(单一文件)* | editor/editor-ops.ts | 1329 | ❌ 超 1000 行 |

架构文档明确要求将 editor-ops.ts 按段落操作/run 样式操作/表格操作拆分为 3 个功能文件 + 1 个 barrel。实际实现保持为单一 1329 行文件。虽然上游对齐文档将此文件标为"直接复制"（不拆分），但架构文档（2026-07-07，晚于上游对齐文档的 2026-07-06）作为设计阶段产物明确提出了拆分要求。

#### editor/helpers/ ⚠️ 2 个文件超 1000 行

| 架构文档 | 实际文件 | 架构预期行数 | 实际行数 | 状态 |
|---|---|---|---|---|
| paragraph-geometry.ts | paragraph-geometry.ts | ~800 | **1573** | ❌ 超限 |
| line-height-table.ts | line-height-table.ts | ~730 | **1128** | ❌ 超限 |
| selection-helpers.ts | selection-helpers.ts | ~605 | **1000** | ⚠️ 临界 |

其余 35 个 helpers 文件全部 ≤990 行，符合约束。

额外的 helpers 文件（不在架构文档清单中）：
- `line-height-wrap.ts` (411 行) — 从 line-height.ts 提取的 wrap 计算逻辑
- `paragraph-toc.ts` (34 行) — TOC 叶子节点，用于打破循环依赖
- 两者均为合理的进一步拆分，帮助控制单文件行数

### 2.6 composables/ ✅ 完全对齐

| 架构文档 | 实际文件 | 行数 | 状态 |
|---|---|---|---|
| useDocxEditor.ts | useDocxEditor.ts | 639 | ✅ |
| editor-transaction.ts | editor-transaction.ts | 130 | ✅ |
| editor-history.ts | editor-history.ts | 60 | ✅ |
| editor-selection.ts | editor-selection.ts | 54 | ✅ |
| editor-text-input.ts | editor-text-input.ts | 607 | ✅ |
| editor-format.ts | editor-format.ts | 337 | ✅ |
| editor-table.ts | editor-table.ts | 186 | ✅ |
| editor-image.ts | editor-image.ts | 161 | ✅ |
| editor-form-field.ts | editor-form-field.ts | 66 | ✅ |
| editor-list.ts | editor-list.ts | 69 | ✅ |
| editor-clipboard.ts | editor-clipboard.ts | 141 | ✅ |
| editor-import-export.ts | editor-import-export.ts | 127 | ✅ |
| page-surface-registry.ts | page-surface-registry.ts | 141 | ✅ |
| useDocxPageThumbnails.ts | useDocxPageThumbnails.ts | 194 | ✅ |
| useDocxDocumentTheme.ts | useDocxDocumentTheme.ts | 33 | ✅ |
| useDocxParagraphStyles.ts | useDocxParagraphStyles.ts | 29 | ✅ |
| useDocxImageWrapMenu.ts | useDocxImageWrapMenu.ts | 90 | ✅ |
| useDocxLineSpacing.ts | useDocxLineSpacing.ts | 34 | ✅ |
| useDocxBorders.ts | useDocxBorders.ts | 34 | ✅ |
| useDocxFormFields.ts | useDocxFormFields.ts | 56 | ✅ |
| useDocxTrackChanges.ts | useDocxTrackChanges.ts | 48 | ✅ |
| useDocxComments.ts | useDocxComments.ts | 48 | ✅ |
| useDocxPageLayout.ts | useDocxPageLayout.ts | 64 | ✅ |
| useDocxPagination.ts | useDocxPagination.ts | 29 | ✅ |
| useDocxModel.ts | useDocxModel.ts | 68 | ✅ |
| index.ts | index.ts | 68 | ✅ |

额外文件：
- `composables.ts` — 向后兼容 barrel，仅 re-export `./composables/index`
- `editor-shared.ts` (91 行) — EditorCore 接口定义，供各子模块共享类型

### 2.7 components/ ✅ 完全对齐

17 个 Vue 组件全部存在，全部 ≤550 行：

| 架构文档 | 实际文件 | 行数 | 状态 |
|---|---|---|---|
| DocxViewer.vue | DocxViewer.vue | 543 | ✅ |
| DocxEditor.vue | DocxEditor.vue | 223 | ✅ |
| DocxViewerRoot.vue | DocxViewerRoot.vue | 339 | ✅ |
| DocxPageWrapper.vue | DocxPageWrapper.vue | 94 | ✅ |
| DocxPageSurface.vue | DocxPageSurface.vue | 249 | ✅ |
| DocxPageHeader.vue | DocxPageHeader.vue | 87 | ✅ |
| DocxPageFooter.vue | DocxPageFooter.vue | 88 | ✅ |
| DocxPageBody.ts | DocxPageBody.ts | 177 | ✅ |
| DocxParagraphHost.vue | DocxParagraphHost.vue | 276 | ✅ |
| DocxTableHost.vue | DocxTableHost.vue | 430 | ✅ |
| DocxImageLayer.vue | DocxImageLayer.vue | 256 | ✅ |
| DocxFormFieldLayer.vue | DocxFormFieldLayer.vue | 202 | ✅ |
| DocxTrackedChangeGutter.vue | DocxTrackedChangeGutter.vue | 197 | ✅ |
| DocxContextMenu.vue | DocxContextMenu.vue | 431 | ✅ |
| DocxToolbar.vue | DocxToolbar.vue | 260 | ✅ |
| DocxThumbnailPanel.vue | DocxThumbnailPanel.vue | 320 | ✅ |
| DocxDragOverlay.vue | DocxDragOverlay.vue | 68 | ✅ |

### 2.8 render/ ✅ 完全对齐 + 1 额外文件

| 架构文档 | 实际文件 | 行数 | 状态 |
|---|---|---|---|
| paragraph-runs.ts | paragraph-runs.ts | 931 | ✅ |
| paragraph-runs-text.ts | paragraph-runs-text.ts | 331 | ✅ |
| paragraph-runs-image.ts | paragraph-runs-image.ts | 509 | ✅ |
| paragraph-runs-field.ts | paragraph-runs-field.ts | 170 | ✅ |
| static-html.ts | static-html.ts | 33 | ✅ |
| index.ts | index.ts | barrel | ✅ |
| *(无)* | paragraph-runs-field-resolve.ts | 158 | ➕ 额外拆分 |

`paragraph-runs-field-resolve.ts` 是表单域解析逻辑的进一步提取，合理。

---

## 三、硬约束违规详情

### ❌ 违规 1: `paragraph-geometry.ts` — 1573 行（上限 1000）

**文件**: `packages/docx-core/src/editor/helpers/paragraph-geometry.ts`
**架构预期**: ~800 行
**实际**: 1573 行（超限 57%）

该文件合并了多个上游行号范围的浮动图片检测函数（5012-5811, 6368-6377, 14008-14052, 14054-14098, 14720-14785, 2045-2054），累计约 976 行上游内容 + TypeScript 类型标注和注释。

**建议修复**：按上游行号范围拆分为 2 个文件：
- `paragraph-geometry.ts`: 浮动图片检测核心（上游 5012-5811，~900 行）
- `paragraph-geometry-image.ts`: 浮动图片渲染模式判定（上游 14008-14785 + 2045-2054 + 6368-6377，~700 行）

### ❌ 违规 2: `line-height-table.ts` — 1128 行（上限 1000）

**文件**: `packages/docx-core/src/editor/helpers/line-height-table.ts`
**架构预期**: ~730 行
**实际**: 1128 行（超限 13%）

**建议修复**：将表格行高估算逻辑拆分为 2 文件：
- `line-height-table.ts`: 单元格/行高核心计算（~700 行）
- `line-height-table-extra.ts`: 段落间距 + 表格边框缓冲（~430 行）

### ❌ 违规 3: `editor-ops.ts` — 1329 行（上限 1000）

**文件**: `packages/docx-core/src/editor/editor-ops.ts`
**架构预期**: 拆分为 4 文件
**实际**: 1329 行单文件（超限 33%）

架构文档明确要求拆分为 `paragraph-ops.ts`（段落操作）、`run-style-ops.ts`（run 样式操作）、`table-ops.ts`（表格操作）和 `index.ts`。

**建议修复**：按架构文档 2.5 节执行拆分。

### ⚠️ 临界: `selection-helpers.ts` — 1000 行（= 上限）

**文件**: `packages/docx-core/src/editor/helpers/selection-helpers.ts`
**行数**: 1000 行（恰好达到上限，零余量）

不阻塞交付，但任何后续修改都会导致超限。

---

## 四、Import 路径检查 ✅

- **docx-core** 内部：全部使用相对路径（`../../viewer/`、`../engine/`、`./`），无跨包裸导入
- **vue-docx** 内部：跨包使用 `@extend-ai/docx-core`，本地使用 `./` 相对路径
- `editor/helpers/section-manipulation.ts` 正确引用 `../editor-ops`
- 无循环依赖检测到（TOC 相关已通过 `paragraph-toc.ts` 叶子节点打破）

---

## 五、typecheck ✅

```bash
pnpm typecheck  # 全量通过
```

- `@extend-ai/docx-core`: 零错误
- `@extend-ai/vue-docx`: 零错误

---

## 六、Stub/Mock/Fake 检查 ✅

- 无 `stub`、`mock`、`fake`、`TODO implement`、`not implemented` 模式
- EMF/WMF 图片渲染占位（`image-placeholder` kind）是显式标注的已知限制，符合架构文档"EMF/WMF 占位"描述
- 所有 `placeholder` 引用为合法的 `FormFieldRunNode.placeholder` 字段或 DOM input placeholder 属性

---

## 七、React 残留检查 ✅

- `docx-core` 中零 `import React`、零 `useState/useRef/useCallback/useEffect/useMemo`
- `React.CSSProperties` 仅在注释中出现（说明"已替换为 `Record<string, string | number | undefined>`"）
- `vue-docx` 使用 Vue `ref/computed/watch/shallowRef/h/defineComponent`

---

## 八、上游功能对齐 ✅

26 项功能均有对应实现：

| 功能 | 实现文件 | 状态 |
|---|---|---|
| 文档主题 | useDocxDocumentTheme.ts | ✅ |
| 应用主题 | useDocxDocumentTheme.ts | ✅ |
| 撤销/重做 | editor-history.ts | ✅ |
| 段落样式 | useDocxParagraphStyles.ts + DocxToolbar.vue | ✅ |
| 字体族 | editor-format.ts + DocxToolbar.vue | ✅ |
| 字号 | editor-format.ts + DocxToolbar.vue | ✅ |
| 行距 | useDocxLineSpacing.ts | ✅ |
| 字符格式 | editor-format.ts + style-to-css.ts | ✅ |
| 颜色/高亮 | editor-format.ts | ✅ |
| 超链接 | editor-format.ts | ✅ |
| 段落对齐 | editor-ops.ts (setParagraphAlignment) | ✅ |
| 列表 | editor-list.ts | ✅ |
| 分栏 | useDocxPageLayout.ts | ✅ |
| 缩略图 | useDocxPageThumbnails.ts + DocxThumbnailPanel.vue | ✅ |
| 边框 | useDocxBorders.ts | ✅ |
| 插入图片 | editor-image.ts | ✅ |
| 插入表格 | editor-table.ts | ✅ |
| 缩放 | zoom-utils.ts | ✅ |
| 导入 .docx | editor-import-export.ts + docx-import.ts | ✅ |
| 导出 .docx | editor-import-export.ts + serializer.ts | ✅ |
| 修订开关 | useDocxTrackChanges.ts | ✅ |
| 批注开关 | useDocxComments.ts | ✅ |
| 只读模式 | DocxViewer.vue / DocxEditor.vue | ✅ |
| 右键菜单 | DocxContextMenu.vue | ✅ |
| 表单域配置 | editor-form-field.ts + DocxFormFieldLayer.vue | ✅ |
| DEV 测试钩子 | *(未在组件层暴露，属于构建配置)* | ⚠️ |

---

## 九、额外发现

### 9.1 "react-docx" 字符串残留

以下文件仍使用上游项目名称 `react-docx`：

- `engine/wasm.ts:25` — 错误消息前缀 `"react-docx: setWasmSource must be called..."`
- `engine/wasm.ts:36` — 错误消息 `"react-docx: failed to load the bundled WebAssembly..."`
- `engine/wasm.ts:48` — 错误消息 `"react-docx: the bundled WebAssembly binary failed..."`
- `editor/helpers/performance.ts:6` — `DOCX_IMPORT_PERFORMANCE_PREFIX = "react-docx.import"`

不影响功能，但应统一更新为 `@extend-ai/docx-core`。

### 9.2 无 public wasm 目录

架构文档 7.1 节提到 `public/docx_wasm_bg.wasm`，但 `packages/vue-docx/public/` 目录不存在。WASM 文件可能由 tsup/vite 构建配置管理，需在集成测试时验证。

---

## 十、判定

**Status: BLOCKED**

阻塞原因：3 项违反"单文件 ≤1000 行"硬约束：

1. `paragraph-geometry.ts` — 1573 行（超限 573 行）
2. `line-height-table.ts` — 1128 行（超限 128 行）
3. `editor-ops.ts` — 1329 行（超限 329 行，且未按架构文档拆分）

修复后即可通过。其余所有检查项（typecheck、import 路径、React 残留、stub/mock、功能对齐、文件清单）均通过。
