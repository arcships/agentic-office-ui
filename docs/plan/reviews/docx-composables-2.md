# docx-composables 开发产出 Review #2

Date: 2026-07-08
Reviewer: dim
Status: **blocked**

## 一、检查范围

- 架构设计: `docs/docx-migration-architecture.md` 第三节 (vue-docx composables + components + render)
- 上游对齐: `docs/upstream-docx-feature-alignment.md` (commit `6f70b92`)
- 实现文件:
  - `packages/vue-docx/src/composables/*.ts` (27 文件, 2741 行)
  - `packages/vue-docx/src/components/` (0 文件)
  - `packages/vue-docx/src/render/` (0 文件)
  - `packages/vue-docx/src/index.ts` (115 行)

## 二、Review #1 修复项验证

Review #1 判定 blocked 的三项硬约束修复情况:

| #1 问题 | 状态 | 说明 |
|---|---|---|
| 11 个子编辑 composable 折叠进 useDocxEditor.ts | ✅ 已修复 | 全部拆分为独立文件 |
| useDocxEditor.ts 1365 行超出 1000 行限制 | ✅ 已修复 | 现 637 行 |
| 17 个存根方法 | ❌ 仍部分存在 | 拆出后存根移至子模块，13 个仍为存根 |

## 三、文件清单对照

### 3.1 composables/ (架构 3.1 节: 26 文件)

| # | 架构文档指定文件 | 状态 | 实际行数 | 目标行数 |
|---|---:|---|---:|---|
| 1 | useDocxEditor.ts | ✅ | 637 | ~700 |
| 2 | editor-transaction.ts | ✅ | 130 | ~800 |
| 3 | editor-history.ts | ✅ | 60 | ~500 |
| 4 | editor-selection.ts | ✅ | 54 | ~700 |
| 5 | editor-text-input.ts | ✅ | 150 | ~700 |
| 6 | editor-format.ts | ✅ | 233 | ~600 |
| 7 | editor-table.ts | ✅ | 152 | ~600 |
| 8 | editor-image.ts | ✅ | 97 | ~600 |
| 9 | editor-form-field.ts | ✅ | 66 | ~400 |
| 10 | editor-list.ts | ✅ | 45 | ~400 |
| 11 | editor-clipboard.ts | ✅ | 18 | ~300 |
| 12 | editor-import-export.ts | ✅ | 88 | ~500 |
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
| 26 | index.ts | ✅ | 68 | barrel |
| + | editor-shared.ts | ✅ | 91 | 内部类型 |

**交付 27/26 文件 (104%)。所有文件 ≤1000 行。**

### 3.2 components/ (架构 3.2 节: 18 文件)

| # | 架构文档指定文件 | 状态 |
|---|---:|---|
| 1 | DocxViewer.vue | ❌ 缺失 |
| 2 | DocxEditor.vue | ❌ 缺失 |
| 3 | DocxViewerRoot.vue | ❌ 缺失 |
| 4 | DocxPageWrapper.vue | ❌ 缺失 |
| 5 | DocxPageSurface.vue | ❌ 缺失 |
| 6 | DocxPageHeader.vue | ❌ 缺失 |
| 7 | DocxPageFooter.vue | ❌ 缺失 |
| 8 | DocxPageBody.ts | ❌ 缺失 |
| 9 | DocxParagraphHost.vue | ❌ 缺失 |
| 10 | DocxTableHost.vue | ❌ 缺失 |
| 11 | DocxImageLayer.vue | ❌ 缺失 |
| 12 | DocxFormFieldLayer.vue | ❌ 缺失 |
| 13 | DocxTrackedChangeGutter.vue | ❌ 缺失 |
| 14 | DocxContextMenu.vue | ❌ 缺失 |
| 15 | DocxToolbar.vue | ❌ 缺失 |
| 16 | DocxThumbnailPanel.vue | ❌ 缺失 |
| 17 | DocxDragOverlay.vue | ❌ 缺失 |
| 18 | index.ts | ❌ 缺失 |

**交付 0/18 文件。目录 `packages/vue-docx/src/components/` 为空。**

### 3.3 render/ (架构 3.3 节: 6 文件)

| # | 架构文档指定文件 | 状态 |
|---|---:|---|
| 1 | paragraph-runs.ts | ❌ 缺失 |
| 2 | paragraph-runs-text.ts | ❌ 缺失 |
| 3 | paragraph-runs-image.ts | ❌ 缺失 |
| 4 | paragraph-runs-field.ts | ❌ 缺失 |
| 5 | static-html.ts | ❌ 缺失 |
| 6 | index.ts | ❌ 缺失 |

**交付 0/6 文件。目录 `packages/vue-docx/src/render/` 为空。**

### 3.4 vue-docx/src/index.ts 存根组件

```typescript
// Line 5-10
export const DocxViewer = defineComponent({
  name: "DocxViewer",
  setup() {
    return () => h("div", { class: "docx-viewer-stub" }, "DOCX Viewer (pending)")
  },
})

export const DocxEditorViewer = defineComponent({
  name: "DocxEditorViewer",
  setup() {
    return () => h("div", { class: "docx-editor-stub" }, "DOCX Editor (pending)")
  },
})
```

这两个是正式组件文件 (`components/DocxViewer.vue`, `components/DocxEditor.vue`) 实现前的占位符。

## 四、存根/延期代码汇总

### 4.1 需修复 (非架构文档标注的延期项, 共 13 方法)

`[editor-text-input.ts](packages/vue-docx/src/composables/editor-text-input.ts)`:

- `:101-108` `splitParagraphAtSelection` — 标 "Stub", return undefined
- `:111-119` `insertListItemAfterSelection` — 标 "Stub", return undefined
- `:122-128` `replaceExpandedSelection` — 标 "Stub", return undefined
- `:130-136` `deleteExpandedSelection` — 标 "Stub", return undefined

`[editor-image.ts](packages/vue-docx/src/composables/editor-image.ts)`:

- `:19-22` `insertImageFile` — 注释占位, 无实现
- `:48-50` `setSyntheticTextBoxText` — 空函数体
- `:52-56` `setImageWrapMode` — 空函数体
- `:58-61` `moveFloatingImage` — 空函数体
- `:63-66` `moveSectionFloatingImage` — 空函数体
- `:68-71` `moveParagraphDropCap` — 空函数体
- `:73-75` `setParagraphDropCapFontSizePt` — 空函数体
- `:77-79` `setParagraphDropCapText` — 空函数体
- `:81-83` `moveImage` — 空函数体

`[editor-clipboard.ts](packages/vue-docx/src/composables/editor-clipboard.ts)`:

- `:8-9` `copy` — 标 "Stub", 无实现
- `:12-13` `paste` — 标 "Stub", 无实现

`[editor-list.ts](packages/vue-docx/src/composables/editor-list.ts)`:

- `:39-41` `adjustSelectedListDepth` — 始终 return true, 无实际操作

`[editor-table.ts](packages/vue-docx/src/composables/editor-table.ts)`:

- `:135-138` `moveEmbeddedTableToBody` — 空函数体 + 注释

`[editor-format.ts](packages/vue-docx/src/composables/editor-format.ts)`:

- `:209-212` `applyBorderPreset` — 标 "Schema stub", 空实现

`[editor-import-export.ts](packages/vue-docx/src/composables/editor-import-export.ts)`:

- `:78-80` `exportDocx` — 仅设置 status 字符串, 无 `serializeDocx` 调用或文件写入逻辑

### 4.2 可接受延期 (架构文档 Phase 9 组件集成阶段)

`[useDocxPageThumbnails.ts](packages/vue-docx/src/composables/useDocxPageThumbnails.ts)`:

- `:120-121` 光栅化渲染 pipeline — "deferred to component integration phase"
- `:124-126` prefetch — "deferred to component integration"
- `:189-191` rerender (useDocxViewerThumbnails) — "Full raster pipeline deferred"

### 4.3 行数偏低项

以下子编辑 composable 实际行数显著低于架构文档目标, 反映功能存根导致的代码量不足:

| 文件 | 实际 | 目标 | 缺口原因 |
|---|---:|---|
| editor-transaction.ts | 130 | ~800 | dispatchEditorTransaction 已实现, 但缺少复杂事务类型 (如表格分割/合并等) |
| editor-history.ts | 60 | ~500 | 缺少 history 栈边界控制/快照压缩/历史导航辅助 |
| editor-selection.ts | 54 | ~700 | selectionSession 已实现, 但缺少非矩形选区/hit-test 绑定等 |
| editor-text-input.ts | 150 | ~700 | 4 个核心方法为存根 |
| editor-image.ts | 97 | ~600 | 9 个方法为存根 |
| editor-list.ts | 45 | ~400 | adjustSelectedListDepth 存根 + 缺少多级列表缩进锁 |
| editor-clipboard.ts | 18 | ~300 | copy/paste 均为存根 |

## 五、Import 路径

| 检查项 | 结果 |
|---|---|
| composable → docx-core 使用 `@extend-ai/docx-core` (workspace 包名) | ✅ |
| composable 间引用使用相对路径 (`./editor-shared`, `./page-surface-registry`) | ✅ |
| `packages/docx-core/src/` 内部使用相对路径 | ✅ |
| 无绝对路径引用 | ✅ |

## 六、Typecheck

| 包 | 命令 | 结果 |
|---|---|---|
| `@extend-ai/docx-core` | `pnpm --filter @extend-ai/docx-core typecheck` | ✅ 通过 |
| `@extend-ai/vue-docx` | `pnpm --filter @extend-ai/vue-docx typecheck` | ✅ 通过 |

## 七、架构硬约束合规

| 约束 | 要求 | 实际 | 合规 |
|---|---|---|---|
| 单文件 ≤ 1000 行 | 全部文件 | composables/ 全部 ≤637 行 | ✅ |
| 模块化按领域分文件 | 26 个 composable 文件 | 27 个文件 (含 editor-shared.ts) | ✅ |
| 功能对齐上游 26 项能力 | 完整实现 | 13 个存根方法 + components/render 完全缺失 | ❌ |
| 仅限 DOCX | DOCX only | 符合 | ✅ |
| components/ 18 文件 | 架构 3.2 节 | 0 文件 | ❌ |
| render/ 6 文件 | 架构 3.3 节 | 0 文件 | ❌ |

## 八、与 Review #1 对比

| 维度 | #1 | #2 | 变化 |
|---|---|---|---|
| composable 文件数 | 15 | 27 | +12 (11 拆出 + 1 editor-shared.ts) |
| useDocxEditor.ts 行数 | 1365 | 637 | -728 (53%) |
| 文件结构合规 | ❌ | ✅ | 子编辑 composable 已拆出 |
| 存根方法数 | 17 | 13 | 减少 4 (transaction/history/selection/form-field 现在有实现) |
| 行数超标文件 | 1 | 0 | ✅ |
| Typecheck | ✅ | ✅ | 保持 |
| components/ 交付 | — | 0/18 | 新增检查维度 |
| render/ 交付 | — | 0/6 | 新增检查维度 |

## 九、判定

**blocked** — 三项硬约束未通过:

1. **存根过多 (13 个方法)**: 图片操作 (9/10 方法存根)、文本编辑核心 (4/4 复杂操作存根)、clipboard、export、list depth、table move、border preset 均为空实现或占位符
2. **components/ 完全缺失 (0/18)**: 架构文档 3.2 节要求的全部 18 个 Vue 组件未实现, 目录为空
3. **render/ 完全缺失 (0/6)**: `renderParagraphRuns` 及其子模块、`renderStaticHtml` 均未实现

### 修复优先级

**P0 — 阻塞组件集成的存根:**
1. `editor-text-input.ts` 4 个存根 (`splitParagraphAtSelection`, `insertListItemAfterSelection`, `replaceExpandedSelection`, `deleteExpandedSelection`) — 文本编辑核心路径
2. `editor-import-export.ts` `exportDocx` — 需调用 `serializeDocx(model, basePackage)` 并触发文件下载
3. `editor-clipboard.ts` `copy`/`paste` — 需要实现 `navigator.clipboard` 读写

**P1 — 工具栏依赖的存根:**
4. `editor-image.ts` `insertImageFile` — FileReader + Uint8Array 转换 + ImageRunNode 插入
5. `editor-format.ts` `applyBorderPreset` — 段落/表格边框样式变更

**P2 — 高级编辑功能:**
6. `editor-image.ts` 剩余 8 个存根 (wrap mode, 浮动定位, drop cap, 移动)
7. `editor-list.ts` `adjustSelectedListDepth`
8. `editor-table.ts` `moveEmbeddedTableToBody`

**P3 — 组件和渲染 (Phase 9):**
9. `components/` 18 个 Vue 组件
10. `render/` 6 个 render 函数文件
