# DOCX Composables — 开发产出 Review #3

Date: 2026-07-08
Reviewer: DimCode
Scope: `packages/vue-docx/src/composables/`（27 文件）
Baseline: `docs/docx-migration-architecture.md` §3.1 + `docs/upstream-docx-feature-alignment.md`

## 总评：PASS（带条件）

核心编辑器骨架（事务分发、撤销/重做、选区管理、文本输入、格式操作）实现扎实，typecheck 零错误，import 路径正确。但存在 **12 处 stub/deferred 实现**，需在下一次 composables 迭代或组件集成阶段补齐。

---

## 1. 文件清单覆盖

| 架构文档预期（§3.1） | 实现 | 状态 |
|---|---|---|
| `useDocxEditor.ts` | ✅ 638 行 | 完成 |
| `editor-transaction.ts` | ✅ 130 行 | 完成 |
| `editor-history.ts` | ✅ 60 行 | 完成 |
| `editor-selection.ts` | ✅ 54 行 | 完成 |
| `editor-text-input.ts` | ✅ 607 行 | 完成 |
| `editor-format.ts` | ✅ 337 行 | 完成 |
| `editor-table.ts` | ✅ 186 行 | 完成 |
| `editor-image.ts` | ⚠️ 161 行 | stub 多 |
| `editor-form-field.ts` | ✅ 66 行 | 完成 |
| `editor-list.ts` | ✅ 69 行 | 完成 |
| `editor-clipboard.ts` | ⚠️ 141 行 | paste 部分实现 |
| `editor-import-export.ts` | ✅ 127 行 | 完成 |
| `page-surface-registry.ts` | ✅ 141 行 | 完成 |
| `useDocxPageThumbnails.ts` | ⚠️ 194 行 | raster pipeline deferred |
| `useDocxDocumentTheme.ts` | ✅ 33 行 | 完成 |
| `useDocxParagraphStyles.ts` | ✅ 29 行 | 完成 |
| `useDocxImageWrapMenu.ts` | ✅ 90 行 | 完成 |
| `useDocxLineSpacing.ts` | ✅ 24 行 | 完成 |
| `useDocxBorders.ts` | ✅ 34 行 | 完成 |
| `useDocxFormFields.ts` | ✅ 56 行 | 完成 |
| `useDocxTrackChanges.ts` | ✅ 48 行 | 完成 |
| `useDocxComments.ts` | ✅ 48 行 | 完成 |
| `useDocxPageLayout.ts` | ✅ 64 行 | 完成 |
| `useDocxPagination.ts` | ✅ 23 行 | 完成 |
| `useDocxModel.ts` | ✅ 68 行 | 完成 |
| `index.ts` | ✅ barrel | 完成 |
| `editor-shared.ts` | ➕ 额外 | 共享 context 类型 |

**结论**：26 个架构文档文件全部存在，覆盖率 100%。额外增加 1 个 shared types 文件（`editor-shared.ts`），属于合理工程决策。

**行数差距**：实际 ~2,700 行 vs 预期 ~8,700 行。差距主要来自两层 wrapper composables（useDocx*）天生轻薄（30-90 行），以及 editor-image、useDocxPageThumbnails 的 stub 实现。

---

## 2. 上游功能对齐（@extend-ai/react-docx commit 6f70b92）

### 2.1 已实现（完整对齐）

| 能力 | 实现位置 | 对齐度 |
|---|---|---|
| `dispatchEditorTransaction` 唯一入口 | `editor-transaction.ts` | ✅ 完整复刻 resolver→normalize→pushHistory→apply 流程 |
| `applyModelChange` 便捷包装 | `editor-transaction.ts:121` | ✅ |
| 快照式撤销/重做（上限 100） | `editor-history.ts` | ✅ past/future 栈 + nonce 选区恢复 |
| Selection session（pointer/keyboard/composition） | `editor-selection.ts` | ✅ begin/clear/settle + timeout |
| contentEditable draft 缓存 + commit | `editor-text-input.ts` | ✅ commitParagraphText / commitTableCellText 等 7 个入口 |
| 跨段落选区替换 | `editor-text-input.ts:replaceExpandedSelection` | ✅ 同段/跨段/表格/整文档 4 场景 |
| 段落拆分 + 列表项插入 | `editor-text-input.ts` | ✅ splitParagraphAtSelection + insertListItemAfterSelection |
| 字符格式化（B/I/U/S/上下标/字号/颜色/高亮） | `editor-format.ts` | ✅ applySelectedStyleChange + toggle* 7 项 |
| pendingRunStyle（光标处待应用样式） | `useDocxEditor.ts:110` | ✅ |
| 段落格式化（heading/styleId/lineSpacing/align/border） | `editor-format.ts` | ✅ |
| 超链接设置/移除 | `editor-format.ts:setLink` | ✅ 展开选区范围内所有 text run 设置 link |
| 表格 CRUD + 行列增删 + 清空 + 移动 | `editor-table.ts` | ✅ insertTable/row/column/delete/move/clear 9 个操作 |
| 嵌套表格提取 | `editor-table.ts:moveEmbeddedTableToBody` | ✅ runtimeKey 解析 + 深度提取 |
| 表单域（checkbox toggle/value/widget） | `editor-form-field.ts` | ✅ select/toggle/setValue/updateWidget |
| 列表 toggle + depth 调整 | `editor-list.ts` | ✅ unordered/ordered toggle + ilvl ±1 (0-8) |
| 剪贴板 copy | `editor-clipboard.ts:copy` | ✅ 同段/跨段/表格 三种选区文本提取 |
| 导入 DOCX（wasm worker/主线程双路径） | `editor-import-export.ts:importDocxFile` | ✅ AbortSignal + AbortController + dynamic import |
| 导出 DOCX（basePackage 保留 styles） | `editor-import-export.ts:exportDocx` | ✅ transformer + blob download |
| 新建文档 | `editor-import-export.ts:newDocument` | ✅ starterTemplate 重置 |
| 文档主题（light/dark） | `useDocxDocumentTheme.ts` | ✅ toggle + isDarkDocument |
| 段落样式列表 | `useDocxParagraphStyles.ts` | ✅ |
| 行距 | `useDocxLineSpacing.ts` | ✅ |
| 边框 | `useDocxBorders.ts` | ✅ |
| 修订/批注 | `useDocxTrackChanges.ts` + `useDocxComments.ts` | ✅ location-indexed lookup |
| 页面布局 | `useDocxPageLayout.ts` | ✅ sectPr 解析 + columns + margins |
| 分页信息 | `useDocxPagination.ts` | ✅ thin wrapper |
| 图片 wrap 菜单 | `useDocxImageWrapMenu.ts` | ✅ mode + positioning options |
| 表单域集合 | `useDocxFormFields.ts` | ✅ collectFormFieldsFromModel + widget patch |
| 页面 surface registry | `page-surface-registry.ts` | ✅ WeakMap 单例 + 订阅 |
| useDocxModel（独立 parser） | `useDocxModel.ts` | ✅ AbortController + watchEffect |

### 2.2 部分实现

| 能力 | 位置 | 差距 |
|---|---|---|
| 图片 resize | `editor-image.ts:resizeImage` | ✅ 已实现，修改 widthPx/heightPx |
| 图片插入 | `editor-image.ts:insertImageFile` | ✅ 已实现，base64 data URI + 段落/单元格双路径 |
| 剪贴板 paste | `editor-clipboard.ts:paste` | ⚠️ readText 成功但**未将文本实际插入模型**，仅设置 status 并把文本暂存到 `_pendingPasteText` |
| 表格对角线边框 | `editor-format.ts:239` | ⚠️ `preset === "diagonal-down"/"diagonal-up"` 直接 return current（未修改模型），注释标记为 deferred |

### 2.3 Stub 实现（8 处，editor-image.ts）

以下函数有签名和空白实现体，上游 react-docx 中有对应完整逻辑：

| 函数 | 预期行为（上游） |
|---|---|
| `setSyntheticTextBoxText` | 修改合成文本框图片的文本内容 |
| `setImageWrapMode` | 设置图片文字环绕模式（inline/square/tight/through/top-bottom/behind/in-front） |
| `moveFloatingImage` | 拖动浮动图片时更新位置 offset |
| `moveSectionFloatingImage` | 页眉/页脚中的浮动图片移动 |
| `moveParagraphDropCap` | 拖动首字下沉位置/大小 |
| `setParagraphDropCapFontSizePt` | 设置首字下沉字号 |
| `setParagraphDropCapText` | 设置首字下沉文本 |
| `moveImage` | 跨位置移动图片（drag & drop） |

### 2.4 缩略图 Raster Pipeline（3 处，useDocxPageThumbnails.ts）

| 函数 | 状态 |
|---|---|
| `renderPageThumbnailToCanvas` | ⚠️ 仅设置 status="unavailable"，不执行实际光栅化 |
| `prefetchPageThumbnailSurface` | ⚠️ 空实现 |
| `rerender`（useDocxViewerThumbnails） | ⚠️ 空实现 |

注释标记为 "deferred to component integration phase"。上游支持三级策略（快照直绘 → DOM→SVG 栅格化 → 离屏渲染）。

---

## 3. Import 路径检查

| 类型 | 路径 | 状态 |
|---|---|---|
| 跨 package | `@arcships/docx-core` | ✅ monorepo workspace 解析正确，typecheck 通过 |
| Intra-composables | `./editor-shared`、`./editor-transaction` 等 | ✅ 相对路径，moduleResolution bundler 兼容 |

所有 import 均正确解析，无 `../../` 越级引用、无循环依赖（composables 之间有向无环：子模块 → editor-shared ← useDocxEditor）。

---

## 4. Typecheck 结果

```
$ pnpm --filter @arcships/vue-docx typecheck
> tsc --noEmit
# exit code 0，零错误
```

✅ 通过。27 个 composable 文件全部通过 strict 模式类型检查。

---

## 5. Stub/Mock/Fake 残留

以下为确认的未完成实现（架构文档 §3.1 未标注为"待实现"）：

| 文件 | 数量 | 影响 |
|---|---|---|
| `editor-image.ts` | **8 个 stub** | 图片 move/wrap/dropCap 能力缺失。调用这些方法不会报错，但无实际效果。 |
| `useDocxPageThumbnails.ts` | **3 个 deferred** | 缩略图面板无法正常渲染。attachCanvas 后端不执行光栅化。 |
| `editor-clipboard.ts` | **1 个 partial** | paste 无法将剪贴板内容写入文档模型。 |
| `editor-format.ts` | **1 个 deferred** | 表格对角线边框 preset 为 no-op。 |
| **合计** | **13 处** | |

架构文档 §3.1 未标记上述任何一处为"待实现"。建议在下一轮迭代中补齐。

---

## 6. 架构合规性

| 约束 | 状态 |
|---|---|
| 单文件 ≤ 1000 行 | ✅ 最大文件 useDocxEditor.ts 638 行 |
| 模块化，按领域分文件 | ✅ 26+1 文件，职责单一 |
| 功能对齐上游 26 项能力 | ⚠️ 核心流程对齐，图片/缩略图/paste 有缺口 |
| 仅限 DOCX | ✅ |
| composables 按职责拆分 | ✅ 事务/历史/选区/文本/格式/表格/图片/域/列表/剪贴板/导入导出 11 个子模块 |

---

## 7. 建议行动

1. **P0**：补齐 `editor-image.ts` 8 个 stub——图片 move/wrap/dropCap（影响图片操作完整性）
2. **P1**：实现 `useDocxPageThumbnails` raster pipeline——至少 direct-draw fast path（影响缩略图面板）
3. **P2**：修复 `editor-clipboard.ts:paste` 实际调用 `replaceExpandedSelection` 插入文本
4. **P2**：实现 `editor-format.ts` 对角线边框（影响表格边框 preset 完整性）

推荐在 composables 迭代 #4 中集中处理上述 13 处缺口，然后在组件集成阶段补齐 raster pipeline 的 live-DOM 路径。
