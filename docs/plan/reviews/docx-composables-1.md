# docx-composables 开发产出 Review

Date: 2026-07-08
Reviewer: dim
Status: **blocked**

## 一、检查范围

- 架构设计: `docs/docx-migration-architecture.md` 第三节 (vue-docx composables)
- 上游对齐: `docs/upstream-docx-feature-alignment.md` (commit `6f70b92`)
- 实现文件: `packages/vue-docx/src/composables/*.ts` (15 文件, 2268 行)

## 二、文件清单对照

| # | 架构文档指定文件 | 状态 | 实际行数 | 目标行数 |
|---|---:|---|---:|---|
| 1 | useDocxEditor.ts | ✅ | **1365** | ≤700 |
| 2 | editor-transaction.ts | ❌ 缺失 | — | ~800 |
| 3 | editor-history.ts | ❌ 缺失 | — | ~500 |
| 4 | editor-selection.ts | ❌ 缺失 | — | ~700 |
| 5 | editor-text-input.ts | ❌ 缺失 | — | ~700 |
| 6 | editor-format.ts | ❌ 缺失 | — | ~600 |
| 7 | editor-table.ts | ❌ 缺失 | — | ~600 |
| 8 | editor-image.ts | ❌ 缺失 | — | ~600 |
| 9 | editor-form-field.ts | ❌ 缺失 | — | ~400 |
| 10 | editor-list.ts | ❌ 缺失 | — | ~400 |
| 11 | editor-clipboard.ts | ❌ 缺失 | — | ~300 |
| 12 | editor-import-export.ts | ❌ 缺失 | — | ~500 |
| 13 | page-surface-registry.ts | ✅ | 141 | ~400 |
| 14 | useDocxPageThumbnails.ts | ✅ | 194 | ~700 |
| 15 | useDocxDocumentTheme.ts | ✅ | 33 | 100 |
| 16 | useDocxParagraphStyles.ts | ✅ | 29 | 100 |
| 17 | useDocxImageWrapMenu.ts | ✅ | 90 | 200 |
| 18 | useDocxLineSpacing.ts | ✅ | 24 | 100 |
| 19 | useDocxBorders.ts | ✅ | 34 | 100 |
| 20 | useDocxFormFields.ts | ✅ | 56 | 200 |
| 21 | useDocxTrackChanges.ts | ✅ | 48 | 200 |
| 22 | useDocxComments.ts | ✅ | 48 | 200 |
| 23 | useDocxPageLayout.ts | ✅ | 64 | 200 |
| 24 | useDocxPagination.ts | ✅ | 23 | 200 |
| 25 | useDocxModel.ts | ✅ | 68 | 200 |
| 26 | index.ts | ✅ | 51 | barrel |

**交付 15/26 文件 (58%)。**

11 个子编辑 composable 被折叠进 useDocxEditor.ts，导致该文件 **1365 行，超出 1000 行硬约束 36.5%**。架构文档明确要求每个 composable 独立文件 (≤1000 行)。

## 三、功能对齐 (上游 26 项能力)

### 3.1 useDocxEditor.ts 已实现

| 上游能力 | 状态 | 说明 |
|---|---|:---:|---|
| dispatchEditorTransaction 唯一修改入口 | ✅ | 核心事务分发器已实现 |
| 快照式历史 (undo/redo, 上限 100) | ✅ | 内联在 useDocxEditor 中 |
| pendingRunStyle (光标处待应用样式) | ✅ | 已实现 |
| applySelectedStyleChange (展开/折叠处理) | ✅ | 已实现 |
| 字符格式 B/I/U/S/上下标 | ✅ | 完整 |
| 字体族/字号/颜色/高亮 | ✅ | 完整 |
| 超链接设置/移除 | ✅ | 完整 |
| 段落对齐/行距/标题级别 | ✅ | 完整 |
| 段落样式切换 | ✅ | 完整 |
| 列表 toggle/unordered/ordered | ✅ | 基础功能 |
| 表格 insert/delete row/column/table/move | ✅ | 完整 |
| 表格单元格文本编辑 | ✅ | commitTableCellText + recursive |
| 表单域 select/toggle/value/widget | ✅ | 完整 |
| 页眉页脚文本编辑 | ✅ | commitSectionParagraphText |
| 图片 resize | ✅ | resizeImage |
| 导入 .docx (wasm worker) | ✅ | importDocxFile |
| 选区管理 (paragraph/table-cell) | ✅ | 基础 |
| newDocument | ✅ | 重置到 starter model |

### 3.2 useDocxEditor.ts 存根/未实现

| 上游能力 | 状态 | 位置 | 说明 |
|---|---|:---:|---|---|
| **contentEditable + draft 缓存** | ⚠️ 部分 | L712-768 | commitParagraphText 存在，但 suppress + nonce + selectionSession 机制缺失 |
| **selectionSession** | ❌ 缺失 | — | 架构文档明确要求保留的机制 |
| **clipboard (复制/粘贴)** | ❌ 缺失 | — | 无任何实现 |
| **图片 insert** | ❌ 存根 | L1105 | insertImageFile 仅注释，无实际逻辑 |
| **图片 wrap mode** | ❌ 存根 | L1134 | setImageWrapMode 空函数体 |
| **图片浮动定位** | ❌ 存根 | L1135-1136 | moveFloatingImage / moveSectionFloatingImage 空函数体 |
| **图片拖动/移动** | ❌ 存根 | L1140 | moveImage 空函数体 |
| **首字下沉 (drop cap)** | ❌ 存根 | L1137-1139 | 3 个方法全部空函数体 |
| **合成文本框** | ❌ 存根 | L1133 | setSyntheticTextBoxText 空函数体 |
| **边框预设** | ❌ 存根 | L902-905 | applyBorderPreset 仅注释 "Schema stub" |
| **splitParagraphAtSelection** | ❌ 存根 | L940-948 | return undefined |
| **insertListItemAfterSelection** | ❌ 存根 | L950-958 | return undefined |
| **replaceExpandedSelection** | ❌ 存根 | L975-978 | return undefined |
| **deleteExpandedSelection** | ❌ 存根 | L980-983 | return undefined |
| **adjustSelectedListDepth** | ❌ 存根 | L934-937 | 始终 return true，无实际操作 |
| **moveEmbeddedTableToBody** | ❌ 存根 | L1099-1101 | 空函数体 + 注释 |
| **exportDocx** | ❌ 存根 | L1255-1257 | 仅设置 status 字符串，无文件写入 |
| **表格 resize handle** | ❌ 缺失 | — | 无 pointer 事件处理 |

**计数: 17 个存根/缺失方法，覆盖图片操作、文本编辑核心、导出、边框、拖拽等关键领域。**

### 3.3 useDocxPageThumbnails.ts

| 能力 | 状态 | 说明 |
|---|---|:---:|
| canvas attach/detach | ✅ | 已实现 |
| 分辨率计算 | ✅ | 已实现 |
| 页面表面注册订阅 | ✅ | 已实现 |
| **光栅化渲染** | ⚠️ 延期 | `renderPageThumbnailToCanvas` 设 status="unavailable" 但不绘制 |
| **prefetch** | ⚠️ 延期 | 空实现 |
| **rerender** | ⚠️ 延期 | 空实现 (useDocxViewerThumbnails) |

架构文档标注缩略图光栅化 pipeline 在 phase 9 组件集成阶段完成，此处可接受。

### 3.4 薄包装 composables (已对齐)

以下 9 个 composable 是对 editor controller 属性的薄包装，功能与上游完全对齐:

- useDocxDocumentTheme ✅
- useDocxParagraphStyles ✅
- useDocxImageWrapMenu ✅
- useDocxLineSpacing ✅
- useDocxBorders ✅
- useDocxFormFields ✅
- useDocxTrackChanges ✅
- useDocxComments ✅
- useDocxPagination ✅
- useDocxPageLayout ✅
- useDocxModel ✅

## 四、Import 路径

| 检查项 | 结果 |
|---|---|
| composable → docx-core 使用 `@extend-ai/docx-core` (workspace 包名) | ✅ 正确 |
| composable 间引用使用相对路径 (`./page-surface-registry`) | ✅ 正确 |
| 无绝对路径引用 | ✅ |

## 五、Typecheck

| 包 | 命令 | 结果 |
|---|---|---|
| `@extend-ai/docx-core` | `pnpm --filter @extend-ai/docx-core typecheck` | ✅ 通过 |
| `@extend-ai/vue-docx` | `pnpm --filter @extend-ai/vue-docx typecheck` | ✅ 通过 |

## 六、存根/延期代码汇总

### 需修复 (非架构文档标注的延期项)

`[useDocxEditor.ts](packages/vue-docx/src/composables/useDocxEditor.ts)`:

- `:903` `applyBorderPreset` — "Schema stub" 注释，空实现
- `:934-937` `adjustSelectedListDepth` — 始终 return true
- `:940-948` `splitParagraphAtSelection` — 标 "Stub"，return undefined
- `:950-958` `insertListItemAfterSelection` — 标 "Stub"，return undefined
- `:975-978` `replaceExpandedSelection` — 标 "Stub"，return undefined
- `:980-983` `deleteExpandedSelection` — 标 "Stub"，return undefined
- `:1099-1101` `moveEmbeddedTableToBody` — 空函数体
- `:1105-1108` `insertImageFile` — 注释占位，无实现
- `:1133` `setSyntheticTextBoxText` — 空函数体
- `:1134` `setImageWrapMode` — 空函数体
- `:1135` `moveFloatingImage` — 空函数体
- `:1136` `moveSectionFloatingImage` — 空函数体
- `:1137` `moveParagraphDropCap` — 空函数体
- `:1138` `setParagraphDropCapFontSizePt` — 空函数体
- `:1139` `setParagraphDropCapText` — 空函数体
- `:1140` `moveImage` — 空函数体
- `:1255-1257` `exportDocx` — 仅 status 字符串，无序列化/文件写入

### 可接受延期 (架构文档明确标注)

`[useDocxPageThumbnails.ts](packages/vue-docx/src/composables/useDocxPageThumbnails.ts)`:

- `:120-121` 光栅化 pipeline — "deferred to component integration phase"
- `:124-126` prefetch — "deferred to component integration"
- `:189-191` rerender (useDocxViewerThumbnails) — "Full raster pipeline deferred"

## 七、架构硬约束合规

| 约束 | 要求 | 实际 | 合规 |
|---|---|---|---|
| 单文件 ≤ 1000 行 | 全部文件 | useDocxEditor.ts 1365 行 | ❌ |
| 模块化按领域分文件 | 26 个 composable 文件 | 15 个文件 (11 个被折叠) | ❌ |
| 功能对齐上游 26 项能力 | 完整实现 | 17 个存根/缺失 | ❌ |
| 仅限 DOCX | DOCX only | 符合 | ✅ |

## 八、判定

**blocked** — 三项硬约束均未通过:

1. **文件结构不符合设计**: 11 个独立的子编辑 composable 被折叠进 useDocxEditor.ts，违反了架构文档的模块化设计要求
2. **单文件行数超标**: useDocxEditor.ts (1365 行) 超出 1000 行限制
3. **功能存根过多**: 17 个方法为存根/空实现，覆盖图片操作、文本编辑核心、导出、边框等非延期领域

### 修复建议

按架构文档 3.1 节文件清单，从 useDocxEditor.ts 拆出以下 11 个文件:

1. `editor-transaction.ts` — dispatchEditorTransaction + applyModelChange
2. `editor-history.ts` — undo/redo + 快照栈管理
3. `editor-selection.ts` — selection + selectionSession + historyRestoreRequest
4. `editor-text-input.ts` — commitParagraphText + contentEditable draft 缓存 + splitParagraphAtSelection + replaceExpandedSelection + deleteExpandedSelection
5. `editor-format.ts` — applySelectedStyleChange + 所有格式化方法 (B/I/U/S/字族/字号/颜色/高亮/链接/对齐/行距/段落样式/边框)
6. `editor-table.ts` — 表格 insert/delete/move/resize + clearTableCellContents
7. `editor-image.ts` — insertImageFile + resizeImage + wrap mode + 浮动定位 + 拖拽移动 + drop cap
8. `editor-form-field.ts` — selectFormField + toggleFormCheckbox + setFormFieldValue + updateFormFieldWidget
9. `editor-list.ts` — toggleList + adjustSelectedListDepth + insertListItemAfterSelection
10. `editor-clipboard.ts` — copy/paste (navigator.clipboard)
11. `editor-import-export.ts` — importDocxFile (含 abort logic) + exportDocx + newDocument + registerPendingExportModelTransformer

拆出后 useDocxEditor.ts 仅保留 ~200 行的 state 初始化 + controller 组装和返回。
