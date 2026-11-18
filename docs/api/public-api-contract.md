# 九个公开包的接口合同

本文登记 `@arcships/office-interaction`、`@arcships/docx-core`、`@arcships/xlsx-core`、`@arcships/pptx-core`、`@arcships/vue-docx`、`@arcships/vue-xlsx`、`@arcships/vue-pptx`、`@arcships/vue-pdf` 和 `@arcships/vue-ui` 的公开入口、推荐用法、事件、错误与兼容期限。

当前九个公开包的源码候选版本统一为 `0.7.0`，npm 稳定版仍为 `0.5.4`。`@arcships/office-interaction` 提供跨格式引用、选择事件、运行时校验和纯函数；DOCX、XLSX、PPTX 与 PDF 已接入阶段一格式适配器和统一 Surface 事件。已有公开搜索、原始选择、对象点击和缩放合同保持向后兼容。源码中的 `@deprecated Since 0.2.0` 继续生效；整个 `0.x` 保留旧入口，最早只能在 `1.0.0` 删除。

## 1. 适用规则

- 只有各包 `package.json` 的 `exports` 列出的路径是受支持的导入路径。`dist` 内其他文件和源码路径不属于公开入口。
- 根入口中已经导出的名称属于公开接口，即使名称看起来像内部组件，也不能在 `0.x` 中直接删除。
- 新代码优先使用高层组件、实例 Runtime 和 `./core` 纯数据入口。旧的全局 WASM 配置只用于兼容。
- 事件名、事件参数、错误码、属性默认值、Worker/WASM 地址和样式入口都属于合同。
- 已公开但准备收口的接口从 `0.2.0` 标记弃用，整个 `0.x` 保留，最早在 `1.0.0` 删除。删除前必须提供迁移示例，并用真实发布压缩包同时验证新旧入口。
- 从未导出的临时桥接不受公开接口版本周期保护，但必须登记负责人和删除日期。清理内部桥接不得导致公开兼容入口提前消失。

## 2. 精确导出路径

下表与九个包当前 `package.json` 的 `exports` 一一对应。未列出的深层路径必须由 Node 返回 `ERR_PACKAGE_PATH_NOT_EXPORTED`。

| 包 | 公开路径 | 类型文件 | 运行文件或资源 |
|---|---|---|---|
| `@arcships/office-interaction` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./package.json` | 不适用 | `./package.json` |
| `@arcships/docx-core` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./core` | `./dist/core.d.ts` | `./dist/core.js` |
|  | `./runtime` | `./dist/runtime.d.ts` | `./dist/runtime.js` |
|  | `./wasm-url` | `./dist/wasm-url.d.ts` | `./dist/wasm-url.js` |
|  | `./assets/docx_wasm_bg.wasm` | 不适用 | `./dist/assets/docx_wasm_bg.wasm` |
|  | `./worker` | 不适用 | `./dist/docx-import-worker.js` |
|  | `./export-worker` | 不适用 | `./dist/docx-export-worker.js` |
|  | `./package.json` | 不适用 | `./package.json` |
| `@arcships/xlsx-core` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./core` | `./dist/core.d.ts` | `./dist/core.js` |
|  | `./runtime` | `./dist/runtime.d.ts` | `./dist/runtime.js` |
|  | `./wasm-url` | `./dist/wasm-url.d.ts` | `./dist/wasm-url.js` |
|  | `./assets/duke_sheets_wasm_bg.wasm` | 不适用 | `./dist/assets/duke_sheets_wasm_bg.wasm` |
|  | `./worker` | 不适用 | `./dist/xlsx-worker.js` |
|  | `./package.json` | 不适用 | `./package.json` |
| `@arcships/vue-docx` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./style.css` | 不适用 | `./dist/style.css` |
| `@arcships/vue-xlsx` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./chart` | `./dist/optional/chart.d.ts` | `./dist/chart.js` |
|  | `./map` | `./dist/optional/map.d.ts` | `./dist/map.js` |
|  | `./webgl` | `./dist/optional/webgl.d.ts` | `./dist/webgl.js` |
|  | `./style.css` | 不适用 | `./dist/style.css` |
| `@arcships/vue-pdf` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./assets/pdfium.wasm` | 不适用 | `./dist/pdfium.wasm` |
|  | `./style.css` | 不适用 | `./dist/style.css` |
| `@arcships/vue-ui` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./style.css` | 不适用 | `./dist/style.css` |
| `@arcships/pptx-core` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./browser` | `./dist/browser.d.ts` | `./dist/browser.js` |
|  | `./package.json` | 不适用 | `./package.json` |
| `@arcships/vue-pptx` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./style.css` | 不适用 | `./dist/style.css` |

Worker 和 WASM 必须来自安装后的真实包。消费项目不得从 demo 公共目录复制这些文件，也不得通过工作区源码别名绕过 `exports`。

## 3. 推荐的稳定入口

### 3.1 `@arcships/docx-core`

- `@arcships/docx-core/core`：平台无关的文档模型、标准化、克隆、布局及不可变段落、表格和文字样式命令。
- `@arcships/docx-core/runtime`：`createDocxRuntime`、`DocxRuntime`、`DocxRuntimeConfig`、来源与限制类型，以及 DOCX 导入的公开诊断和错误类型。
- `@arcships/docx-core/wasm-url`：包自带的 `bundledDocxWasmUrl`。
- 根入口的 `exportDocxBuffer(model, basePackage, options)` 在 UI 环境中使用专用导出 Worker，支持 `onProgress`、`AbortSignal` 取消和 `maxExportMs` 超时；`createBundledDocxExportWorker` 与 `@arcships/docx-core/export-worker` 供自定义 Worker 配置。`DocxExportError` 返回结构化 `ABORTED`、`TIMEOUT`、`WORKER_UNAVAILABLE`、`EXPORT_FAILED` 或 `INVALID_MODEL`。
- 根入口的 `prepareModelForSerialization(model)` 是导出前的模型准备步骤：递归正文、页眉、页脚、分节页眉/页脚、脚注和尾注后，把带 `collapsedGridColumns` 标记的表格投影到物理列（`projectCollapsedTableGrid`、`hasCollapsedTableGrid`）。只有重建的表格（编辑清除了 `sourceXml`）会写物理网格，未编辑表格继续使用原始标记。违反样式/边框间距不变式或网格自相矛盾的模型会抛出 `DocxSerializationError`（`COLLAPSED_TABLE_GRID_STYLE_CHANGED`、`COLLAPSED_TABLE_GRID_BORDER_SPACING`、`INVALID_COLLAPSED_TABLE_GRID` 及 `reason`）；`exportDocxBuffer` 把它映射为 `DocxExportError` 的 `INVALID_MODEL` 码。（issue #134）
- `DocxEditorController.getDocxBytes()` 与 `getDocxBlob()` 返回与 `exportDocx()` 相同的序列化结果，但不创建下载链接，宿主可以直接上传、存储或比较字节。两者复用导出准备流程（提交文字草稿、pending export transformer、`basePackage` 与必填表单校验）和 `exportDocxBuffer` 的 Worker/取消/超时语义：文档被替换或编辑器销毁时，进行中的读取会以 `DocxExportError`（`ABORTED`）拒绝；`exportDocx()` 的待处理请求还会被更新的导出取消。`exportDocx()` 内部调用 `getDocxBytes()`，只在拿到字节后创建 `<a download>`。
- 根入口：保留现有模型、布局、序列化、解析和编辑命令，供 `0.x` 兼容使用。新代码如只需要纯数据能力，应改用 `./core`；需要 Worker/WASM 时应改用 `./runtime`。
- 根入口同时提供 `createDocxTextReferenceDraft`、`createDocxPageReferenceDraft`、`createDocxRegionReferenceDraft`、`resolveDocxReference` 和 `describeDocxReference`。它们只把已有 DOCX 模型选区转换/解析为统一引用，不持有宿主的引用集合。
- 根入口的 `TextStyle` 继续兼容原有布尔字段，同时新增 underline/strike 变体、`caps`/`smallCaps`、字符间距与缩放、分脚本字体、run 边框、文字阴影/描边和 `direction`；`ParagraphStyle` 新增 `direction`。`paragraphTypographyStyle` 与 `isJustifiedAlignment` 供渲染层统一处理扩展对齐和 RTL。
- 根入口提供 DOCX 字体子系统（issue #123）：`fontMetrics`、`invalidateFontMetrics()`、`getFontMetricsRevision()`、`registerFontMetricCache()`、`subscribeFontMetrics()`；`fontAdvanceWidthPx` / `measureTextWidthPx` 按字体族、字重、字号、kerning 与 OpenType 特性缓存 advance，缺失 canvas 时回退到隐藏 DOM span，再回退到确定性字宽。`resolveDocxFontKerning`、`resolveDocxFontFeatureSettings`、`segmentTextByDocxScriptFont`、`resolveDocxScriptFontFamily`、`fontTableEntries`、`documentFontRequests` 与 `awaitDocumentFonts` 可供自定义渲染和宿主加载文档字体（含 `word/fontTable.xml` 内嵌字体）。`runStyleToCss` 输出 `font-kerning` / `font-feature-settings`，渲染层按 ascii/hAnsi/eastAsia/cs 分槽输出 span。`prepareMeasuredCached(text, font, wordBreak, settings)` 返回一份把 pretext 分段宽度替换为 DOCX 度量的 prepared text，`resolveUniformPretextSourceTypography()` 只返回字体、kerning 与特性全部一致的单字体快速路径参数；断行引擎和 DOM 因此使用同一组 advance。
- 根入口提供修订显示模型（issue #149）：`DocxRevisionDisplay` 取 `inline`、`balloons` 或 `none`；`reviewDisplayModel(model, display)` 返回只读显示模型，`balloons`/`none` 会把删除/移出内容移出正文，但源模型、接受/拒绝和导出仍保留全部修订。`reviewDisplayParagraph`、`hiddenParagraphChildIndexes`、`hiddenParagraphTextRanges`、`displayTextOffsetForSourceOffset` 与 `sourceTextOffsetForDisplayOffset` 供自定义渲染复用同一套显示文本与光标偏移规则。`@arcships/vue-docx` 还导出 `paragraphDisplayOffsetToSourceOffset`、`paragraphSourceOffsetToDisplayOffset`、`displayTextRangeToSourceTextRange`、`sourceTextRangeToDisplayTextRange` 与 `effectiveControllerRevisionDisplay`，供正文、表格、页眉页脚和脚注/尾注表面复用同一套 DOM 选区 ↔ 源模型偏移映射。
- 根入口同时公开段落排版输入：`ParagraphSpacing.beforeAutoSpacing` / `afterAutoSpacing`、`DocumentCompatibilitySettings.compatibilityMode` 与 `metadata.defaultTabStopTwips` 由 `word/settings.xml` 与段落属性导入；`resolveParagraphAutoSpacingPx`、`resolveParagraphSpacingPx`、`resolveDocumentLayoutDefaults`、`resolveDefaultTabStopPx`、`documentLayoutCacheKey` 与 `allowsWrappedLeftTabRefit` 供自定义布局复用同一套规则（自动间距仅在其兼容模式 12/15 生效、同编号列表项之间折叠、默认制表位可换行重排）。（issue #129）
- 根入口公开测量驱动分页迭代（issue #99）：`PaginationMeasurementLoop` 以页面身份、渲染高度和反馈判定 `repaginate` / `converged` / `oscillating` / `exhausted`，`iteratePaginationWithMeasurements` 为同步驱动器；`PaginationMeasurementFeedback` 会把 HTML 测得的段落高度与表格行高反馈给分页引擎。`LayoutSnapshot.pageCountTarget`、`resolvePaginationPageCountTarget` 和 `LayoutSnapshotOptions.pageCountTarget` 把 `documentPageCount` 表达为 `enforced: false` 的目标提示，不会缩放页面或合并尾页；`reuseUnchangedPaginationItems` 与 `resolvePaginationScrollAnchor` / `resolvePaginationScrollTop` 供渲染层只重建分段变化的页面并保持滚动锚点。旧的 `reconcilePageCountCandidateToTargetCountByScalingHeight` 与 `mergeTrailingPagesToTargetCount` 在 `0.x` 继续导出但已标记 `@deprecated`，默认不再生成 1.02–1.3 的放大候选。

根入口公开裁剪后的页面测量（issue #135）：`createClippedBodyRectReader` 沿祖先交集 `overflow: hidden/clip/auto/scroll` 的裁剪框、按元素缓存结果，并把屏幕坐标换算回缩放前的布局像素；`collectPageContentMeasurementDescendants` 收集参与测量的后代并跳过 `data-docx-pagination-ignore` 子树；`resolveClippedBodyRenderedBottomPx` 把它交给 `resolveMeasuredBodyRenderedBottomPx`，`resolveClippedMeasuredHeightPx` 决定用可见底部还是 `offsetHeight` 回退。页面高度与页面内容高度的反馈只统计可见部分，被 `overflow` 裁掉的内容不再撑高页高。（issue #135）

`DocxRuntimeConfig.limits` 使用实例自己的 `DocxRuntimeLimits`。公开限制覆盖输入字节、压缩包条目与解压字节、XML、关系、图片、模型节点、布局页数、解析时间与导出时间；默认值由 `DEFAULT_DOCX_RUNTIME_LIMITS` 给出，`resolveDocxRuntimeLimits` 返回已约束的实例快照。超限会在高成本解析、图片解码或布局之前停止，并通过结构化错误返回实际值与允许值。UI 环境不会回退到主线程同步 WASM 解析。

根入口同时公开低层校验辅助项 `assertDocxModelBudget`、`assertDocxParseTime`、`validateDocxImageAssets` 和 `DocxImageBudgetSnapshot`，供自定义导入流程复用与内置 Runtime 相同的模型、耗时和图片规则。

根入口公开分节页码跟踪（issue #122）：`createSectionPageNumberTracker({ sections, pages, nodeCount })` 把“每页所属分节 + 各分节 `w:pgNumType`”解析成每页的 `PAGE` / `SECTIONPAGES` 取值与标签（`pageNumberForPage`、`pageNumberFormatForPage`、`pageLabelForPage`、`sectionPagesForPage`、`fieldsForPage`），`NUMPAGES` 始终是文档物理页数。`w:pgNumType/@w:start` 重新开始计数、`@w:fmt` 决定格式（`decimal`、`decimalZero`、`lowerRoman`、`upperRoman`、`lowerLetter`、`upperLetter`、`none` 等），无 `@w:start` 的分节接续上一分节编号；`normalizeDocxPageNumberFormat` 返回 `w:fmt` 的规范拼写。`DocxSectionPageRange` 给出每个分节的页区间、起止编号、格式、`requiredParity`、`parityBlankPageIndexes` 以及覆及的 `DocumentPageRange` 节点区间。`w:type="oddPage"/"evenPage"` 插入的空白页由 `createSectionParityBlankPage` / `isSectionParityBlankPage` 描述，并属于结束于该处的分节（沿用其页眉页脚、继续其编号、计入其 `SECTIONPAGES`），`layoutDocument` 与分页引擎对此一致。

单节点编辑采用写时复制：`cloneDocModelShallow(model)` 只复制模型对象、`nodes` 数组和 `metadata` 对象，`cloneDocNode(node)` 深拷贝被编辑的顶层节点，`cloneDocModelNodes(model, indexes)` 深拷贝命令涉及的一组顶层节点，`cloneDocModelSectionLists(model, region)` 复制页眉/页脚 section 列表外壳。它们与 `model` 共享未修改的节点、页眉页脚节点和图片 `data` 缓冲区。`cloneDocModel(model)` 继续完整拷贝文档图（包括图片字节），用于导入、标准化和长生命周期检查点。编辑路径共享的图片 `data` 只能读，调用方不得原地写入。（issue #37）

根入口公开 DOCX 批注写入（issue #96）：`addDocumentComment(model, { range, text, author, initials, date })`、`replyToDocumentComment(model, parentId, text, options)`、`resolveDocumentComment(model, commentId, resolved)` 与 `deleteDocumentComment(model, commentId)` 以不可变模型返回新修订；`resolveDocumentCommentAnchors` 把跨段落/单元格选区拆成每个段落的 `commentRanges` 偏移，`commentThreadRootId`、`commentThreadMemberIds`、`commentSubtreeIds`、`findDocumentComment` 与 `nextDocumentCommentId` 供宿主检查线程。序列化会写出 `word/comments.xml` 与 `word/commentsExtended.xml`（回复与 `w15:done`），同步 `[Content_Types].xml` 和 `word/_rels/document.xml.rels`；删除最后一条批注时移除这些部件，并清除已删除批注在 `document.xml`、页眉页脚中残留的 `w:commentRangeStart/End` 与 `w:commentReference`。`DocxEditorController` 对应命令为 `addComment(range, text)`、`replyToComment(id, text)`、`resolveComment(id, resolved)`、`deleteComment(id)`，以及 `beginComment` / `beginCommentReply` / `submitCommentComposer` / `cancelCommentComposer` 驱动的内联编辑器；每条命令占一个撤销历史条目，撤销会同时还原批注定义与锚点。新文档的默认作者来自 `docx.editor.commentAuthor`，宿主可用 `commentAuthor` / `commentInitials` / `commentClock` 覆盖。

文本提交合并：`DocxEditorTransactionPatch.textCommitKey` 标记一次文本提交的目标，相同键且无其他事务介入的连续提交只写一条撤销历史；`DocxEditorController.endTextCommitBurst()` 显式结束当前 burst（段落/单元格宿主在 blur 时已自动调用）。选区变化、格式/结构命令、撤销重做、导入/新建均会结束 burst。约定与测试见 [`docs/testing/docx-edit-history-performance.md`](../testing/docx-edit-history-performance.md) 第 4 节。

根入口公开 DOCX 查找替换（issue #113）：`replaceTextRangeInParagraph(paragraph, start, end, replacement)` 在单段内按 UTF-16 偏移替换，只在匹配边界拆分文本 run，保留图片/表单域/链接，替换文字继承被替换位置的 run 样式，平移 `commentRanges` 并使段落 `sourceXml` 失效；`replaceTextInParagraphLocation(model, { location, start, end, replacement })` 接受正文、表格单元格和嵌套表格路径（`kind: "nested-table-cell"` 加 `path`），以写时复制返回新模型；`replaceTextInDocxParagraphs(model, edits)` 按段落从后往前应用一批编辑。`DocxEditorController` 对应命令为 `replaceMatch(target, text)`（一条撤销记录）与 `replaceAll(query, text, options?)`（正文全部命中，一条撤销记录，返回 `{ replaced, skipped }`）；`replaceMatch` 的目标由 `docxReplaceTargetFromMatch(match)` 从 `findDocxSearchMatches()` 的结果转换，页眉、页脚和注释结果不可替换（返回 `undefined`）。

根入口公开 DOCX 图片图层顺序（issue #145）：`DocxEditorController` 新增 `bringToFront(location)`、`bringForward(location)`、`sendBackward(location)` 与 `sendToBack(location)`，每条命令占一个撤销条目并清除图片/段落/表格的 `sourceXml`，导出时重新写入 `wp:anchor/@relativeHeight`。`resolveDocxImageLayerOrderState(floating)` 返回当前截断后的 `zIndex` 与四个 `can*` 可用标志，`resolveDocxImageLayerOrderMenuActions(floating)` 把 `DEFAULT_CONTEXT_MENU_IMAGE_LAYER_ACTIONS` 展开为带禁用状态的四个菜单项，`resolveDocxImageLayerOrderPatch(command, floating)` 给出单个命令的模型补丁；非浮动图片没有图层顺序。`resolveFloatingImageStackingZIndex(floating)` 供自定义渲染层复用与 `DocxImageLayer` 相同的 CSS 叠放规则。

根入口公开 DOCX 书签导航（issue #148）：`createDocxBookmarkIndex(model)` 从 `w:bookmarkStart` 建立书签名到正文节点的索引，`get(name)` / `has(name)` 接受 `#Name`、任意大小写和 URL 转义，`targets` 按文档顺序给出 `{ name, nodeIndex, text, tableCell? }`；`normalizeDocxBookmarkName(name)` 是同一套归一化规则。`_GoBack` 等 Word 自动书签不参与索引；重复书签名以第一个目标为准。

根入口公开 DOCX 表格矩形选区（issue #146）：`DocxTableCellSelectionRange` 增加可选 `gridRange: { startColumnIndex, endColumnIndex }`，表示合并单元格展开后的物理网格列范围。`DocxEditorController` 增加 `selectedTableCellRange`、`selectWholeTable(tableIndex?)`、`collapseTableCellSelection()` 与 `clearTableSelection()`；`clearTableSelection()` 复用 `clearTableCellContents(tableIndex, cells[])`，整个矩形一条撤销记录。`@arcships/docx-core` 同时导出 `isCellWithinTableSelectionRange`、`tableCellSelectionRangeCells`、`tableCellSelectionRangeParagraphEntries`、`tableCellSelectionRangeHasMultipleCells`、`tableCellSelectionAnchor` 与 `tableSelectionGridRectangle`，供自定义表面复用高亮、格式、复制和清除的同一套网格判定。`isCellSelected(selection, tableIndex, rowIndex, cellIndex, table?)` 增加可选 `table` 参数，按物理网格判断合并单元格是否在选区内。

推荐实例写法：

```ts
import { createDocxRuntime } from "@arcships/docx-core/runtime"
import { bundledDocxWasmUrl } from "@arcships/docx-core/wasm-url"

const runtime = createDocxRuntime({
  wasmUrl: bundledDocxWasmUrl,
  createWorker: () => new Worker(workerUrl, { type: "module" }),
})

try {
  const result = await runtime.loadSource({ kind: "bytes", bytes })
  // 使用 result.model
} finally {
  runtime.dispose()
}
```

### 3.2 `@arcships/xlsx-core`

- `@arcships/xlsx-core/core`：平台无关的工作簿数据类型、颜色解析、A1 地址、图表公式和图片锚点计算。
- `@arcships/xlsx-core/runtime`：`createXlsxRuntime`、实例配置、Worker 客户端、来源检查、加载错误和诊断类型。
- `@arcships/xlsx-core/wasm-url`：包自带的 `bundledXlsxWasmUrl`。
- 根入口：保留现有数据、图片、图表、Worker 和控制器类型供 `0.x` 兼容使用。新代码应按纯数据或运行能力选择 `./core`、`./runtime`。
- 图表：`refreshChartDataFromWorkbook(workbook, chartsByWorkbookSheetIndex, chartsheets)` 从编辑后的工作簿重新解析图表分类和值，保留图表定义与锚点；Worker 客户端也提供 `refreshCharts()`，不会重新解析源文件字节。
- 范围高亮（issue #132）：控制器新增 `highlightRanges(ranges, { color, sheetIndex, role, id, label })`、`clearHighlightedRanges()` 和响应式 `highlightedRanges`。`XlsxRangeHighlight`（`range`、`color`、`sheetIndex`、`role`、`seriesIndex`、`seriesName`、`id`、`label`）与 `XlsxRangeHighlightRole`（`category`、`value`、`bubbleSize`、`seriesName`、`custom`）随根入口导出；`highlightRanges()` 替换宿主区域并默认落在活动工作表/琥珀色，图表选择派生区域由选择状态驱动并在取消选择时清除，两者合并到 `highlightedRanges`。该 API 只读，不写工作簿。`XlsxViewer`/`XlsxSheetSurface` 挂载 `XlsxRangeHighlightLayer`，按冻结窗格拆段、跳过隐藏行列，并跟随滚动与缩放。
- 更改图表类型：`changeWorkbookChartType(assets, chartAssets, chartId, type)` 重写图表部件（经典 `c:plotArea` 或 chartEx `cx:chartSpace`），保留系列名称、颜色、数据标签和坐标轴格式；经典与 chartEx 互转时同步 drawing 关系、`mc:AlternateContent` 包裹和内容类型。`chartTypeOptions`、`chartTypeFamilies()`、`resolveChartTypeOption()`、`matchesChartType()`、`chartTypeUnavailable()` 提供菜单目录和可用性判断；`XlsxViewerController.changeChartType(id, type)` 是其控制器入口。
- 根入口同时提供工作表、单元格、A1 范围、整行、整列、图表和人工 sheet 区域的 `createXlsx*ReferenceDraft`、`resolveXlsxReference` 与 `describeXlsxReference`。适配器优先使用 sheet id 和 drawing id；名称/类型证据只有唯一时才用于迁移。
- 富文本单元格（issue #138）：`XlsxCellTextRun` / `XlsxCellTextRunFont` 描述单元格内每一段文字的文本与字体覆盖（粗体、斜体、下划线、删除线、颜色、字号、字体族、`vertAlign` 上标/下标），`XlsxSheetData.richTextCells` 以 A1 地址保存多段单元格，`controller.getCellRuns(cell)` 返回当前单元格的段（值变化后自动不再返回）。`parseRichTextRuns`、`parseInlineRichTextCell`、`parseSharedStringRichText`、`richTextRunsToPlainText` 与 `resolveRichTextRuns` 供自定义渲染复用同一套解析与"文本仍匹配"校验；单元格的纯文本 `value` 始终是各段拼接，公式、搜索和剪贴板继续使用它。

`XlsxRuntimeConfig.limits` 使用实例自己的 `XlsxRuntimeLimits`。公开限制除通用输入、压缩包、XML、关系、图片和解析时间外，还覆盖工作表 XML、共享字符串、行列、工作表数量和公式数量。`DEFAULT_XLSX_RUNTIME_LIMITS` 是默认值，`resolveXlsxRuntimeLimits` 返回实例快照；`validateXlsxArchive` 可在创建工作簿前执行同一套检查。

根入口还公开 `XlsxArchiveEntry`、`XlsxArchiveValidationResult`、`XlsxImageBudgetSnapshot`、`validateXlsxImageAssets`、`trackXlsxWorkbookLifetime`、`anchorToAbsoluteRect`、`XlsxWorkerError` 和 `XlsxWorkerErrorCode`。自定义控制器使用前两组校验结果观察资源边界；直接创建 WASM `Workbook` 时可用生命周期辅助函数确保子对象先于父对象释放；图片和图表浮层可用锚点转换函数按工作表行列尺寸还原位置。

`0.2.0` 还以向后兼容的方式新增 `WorkbookTableMetadata` 和 `normalizeWorkbookTableMetadata`，用于把 XLSX 压缩包中的真实表定义转换为公开 `XlsxTable`；`XlsxChartsheet.charts` 是可选的只读图表列表，用于呈现图表工作表。旧调用方不需要提供或读取这些字段，现有签名不变。

推荐实例写法：

```ts
import { createXlsxRuntime } from "@arcships/xlsx-core/runtime"
import { bundledXlsxWasmUrl } from "@arcships/xlsx-core/wasm-url"

const runtime = createXlsxRuntime({
  wasmSource: bundledXlsxWasmUrl,
  createWorker: () => new Worker(workerUrl, { type: "module" }),
})

try {
  const worker = runtime.createWorkerClient()
  // 通过实例拥有的 Worker 执行解析
} finally {
  runtime.dispose()
}
```

### 3.3 `@arcships/vue-docx`

DOCX 受控只读文档面、分页切片、缩放测量、纸面主题与批注布局的目标合同见 [DOCX Viewer 受控文档面与分页修复设计](../docx-viewer-controlled-surface-and-pagination-fix.md)。兼容性判断以本页当前 `0.7.0` 候选导出清单和对应实现状态为准。

稳定高层入口为：

- 组件：`DocxViewer`、`DocxEditor`。
- 状态与命令：`useDocxEditor`、`useDocxModel`。
- 公共能力组合函数：文档主题、段落样式、图片环绕、行距、边框、表单域、修订、批注、页面布局、分页和 `useDocxPageThumbnails`。
- 编辑辅助界面（issue #101）：`DocxFormFieldDialog`、`DocxLinkHoverCard`、`DocxStylePreviewPopover` 及其纯函数 `docxStylePreviewCss`、`resolveFormFieldAt`、`parseDocxFormFieldLocationKey`；拖拽排序的 `useDocxNodeReorder` 与 `resolveDocxReorderInsertionIndex`、`resolveDocxReorderTarget`、`docxReorderNodeIndexFromElement`。它们只通过 `DocxEditorController` 读写文档，`DocxEditor` 用 `show-form-field-dialog`、`show-link-card`、`show-style-preview`、`show-node-reorder` 四个属性控制是否渲染。
- 书签导航（issue #148）：`DocxViewer`、`DocxEditor` 和 `DocxDocumentSurface` 暴露 `revealBookmark(name)`，返回 `DocxBookmarkRevealResult`（`{ status: "revealed", name, nodeIndex, pageIndex?, text }` 或 `{ status: "missing", name }`）；找不到书签时只显示非阻塞提示，不抛错。查看器需要挂载新页时发出 `requestPageReveal(pageIndex)`，找不到书签时发出 `bookmarkMissing(name)`。链接悬浮卡片的“打开”对内部链接走同一导航，并保留 `link-open`。
- 浮动表格（issue #150）：`DocxEditorController.setTableFloating(tableIndex, update)` 把一个正文表格放到落下位置（写 `TableStyle.floating`，导出为 `w:tblpPr`），传 `null` 把它放回文字流，一次调用一条撤销记录。`DocxTableHost` 的表格移动手柄（`data-docx-table-move-handle`，拖拽时显示幽灵框 `data-docx-table-drag-ghost`）可用 `:show-table-move-handle="false"` 关闭；`DocxTableFloatingUpdate`、`DocxTableFloatingDrop`、`resolveTableFloatingDropPatch()` 和 `mergeTableFloatingStyle()` 是可复用的规则。
- 图形对象与公式（issue #33）：`DrawingRunNode`（`type: "drawing"`，`kind` 为 `textBox` / `shape` / `wordArt` / `group` / `smartArt` / `chart` / `oleObject` / `model3d`）与 `MathRunNode`（`type: "math"`）是新的段落子节点类型；`ImageRunNode` 新增 `textBox` / `textBoxSourceXml`。导入侧的 `paragraphDrawingObjects()`、`classifyDrawingXml()`、`parseChartData()`、`linearizeOmml()`、`parseDrawingAnchor()` 和导出侧的 `applyPreservedDrawingObjectsToDocumentXml()` / `...ToPartContent()` / `...ToPackage()`、`collectPreservedDrawingObjects()` 都是公开的纯函数，与内置序列化补丁共用。`@arcships/docx-core` 导出 `drawingObjectPlaceholderStyle()`、`drawingObjectFrameStyle()`、`drawingObjectSizePx()`、`chartPlaceholderText()`、`chartPreviewDataUri()` 供自定义绘图层使用；文案在 `messages.docx.drawing` 下。`renderParagraphRuns()` 新增 `renderDropCaps`（默认 `true`）和 `drawingLabels` 选项。
- 类型：`DocModel`、`DocxEditorController`、`DocxFormFieldUpdate`、编辑选区、图片、表格、分页、修订、批注及上述组合函数的公开类型。
- 样式：`@arcships/vue-docx/style.css`。
- 历史记录：`UseDocxEditorOptions.historyMaxEntries` 和 `historyMaxBytes` 同时生效；Vue 编辑器默认保留最多 100 条、估算 32 MiB，超过后优先淘汰最旧记录。

页面、段落、表格、图片、工具栏和浮层的单独组件虽然在旧版本已经导出，但不再作为新功能的接入点。它们的兼容期限见第 6 节。

### 3.4 `@arcships/vue-xlsx`

稳定高层入口为：

- `XlsxViewer`。
- `useXlsxViewerController`、`useXlsxViewerThumbnails`。
- `XlsxFileSizeLimitExceededError`。
- `XlsxDiagnostic`、`XlsxLoadError`、`XlsxLoadErrorCode`、`XlsxSourceKind`、`XlsxSourceState`、`XlsxUrlPolicy`。
- 样式：`@arcships/vue-xlsx/style.css`。
- 可选功能：`@arcships/vue-xlsx/chart`、`@arcships/vue-xlsx/map`、`@arcships/vue-xlsx/webgl`。根入口保留旧渲染名称的异步兼容外观；新代码可显式导入对应子入口。未使用功能时不会请求其实现分包。
- 工作表缩略图：`useXlsxViewerThumbnails` 只在需要时读取单元格（首次悬停或显式预热），`prefetchThumbnails` 在空闲时预热指定工作表、按 `revision` 去重、单次最多 3 张，`paintThumbnail` 随后只绘制。
- 历史记录：控制器选项 `historyMaxEntries` 和 `historyMaxBytes` 同时生效，默认最多 100 条、估算 32 MiB。
- 图表数据：`refreshChartData()` 在主线程工作簿或 Worker 当前工作簿上重新解析图表分类和值，用于源单元格编辑后刷新绘图。
- 编写入口：`XlsxViewer` 的功能区与右键菜单打开 `XlsxAuthoringDialog`（图表、批注、超链接、表格、条件格式、数据验证）。宿主也可以直接挂载该组件，或只调用控制器上的 `insertChart`、`setCellComment`、`setCellHyperlink`、`addConditionalFormatRule`、`updateConditionalFormatRule`、`deleteConditionalFormatRule`、`createTable`、`removeTable`、`setSelectionDataValidation`、`setFreezePanes`、`toggleFreezePanesAtActiveCell`。每个操作都写入工作簿模型、记录撤销条目，并随导出的 XLSX 一起保存；`XlsxViewerAuthoring` 汇总这些方法的说明。直接读取/清空用 `getCellComment`、`getCellHyperlink`、`getConditionalFormatRules`、`getTableAtCell`、`freezePanes`。
- 编写限制：这些命令需要可修改的主线程工作簿（Worker 只用于强制只读的大文件），在 `readOnly`、受保护工作表或 Worker 只读模式下返回 `false`/`null`。
- 查找替换（issue #113）：`replaceInSheet(query, replacement, { matchCase, wholeCell, formula, sheetIndex, range, occurrence })` 在同一工作表内替换值和公式文本，`formula` 默认开启，非活动工作表会先被激活，整批写入只占一条撤销记录，返回 `XlsxReplaceResult`（`replacedCells`、`replacedOccurrences`、`changed`）。只读/受保护/Worker 只读模式不改工作簿，返回 `changed: false`。搜索与替换面板、`useXlsxSearch()` 与 `OfficeFindBar` 的替换行复用同一个控制器命令。
- 单元格格式：功能区提供边框预设下拉（全部/外侧/左/右/上/下/无）与边框颜色，并保留“所有边框”快捷按钮；宿主也可以直接调用 `setSelectionBorders({ preset, style, color })`，`XlsxBorderPreset`、`XlsxBorderPresetInput` 描述这组入参。`all` 给选区内每格画边框，`outside` 只画选区外围（内部格子保持原样），`none` 清除四边；颜色同时写入每一面，`XlsxCellBorderStyleInput` 仍支持宿主用 `setRangeStyle` 逐边设置不同颜色。
- 合并：`mergeAndCenterSelection()` 是功能区的默认合并动作，合并后同时写入 `horizontal: center` 与 `vertical: center`，两步共用一条撤销记录；单格选区只居中，不写一个格的合并区。`mergeSelection()` 保留给只需合并、不改对齐的宿主，`unmergeSelection()` 取消合并。
- 范围与图表数据源高亮（issue #132）：`XlsxRangeHighlightLayer` 是根入口新增的浮层出口，`XlsxViewer`/`XlsxSheetSurface` 已自动挂载；`controller.highlightRanges()`/`clearHighlightedRanges()`/`highlightedRanges` 读写范围，选中图表/系列/数据点/图例时按 category/value/bubbleSize 角色自动高亮源区域。根入口同时导出 `CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE`、`XLSX_RANGE_HIGHLIGHT_DEFAULT_COLOR`、`collectChartRangeHighlights`、`parseChartSourceRange`、`parseChartSourceA1Range`、`chartRangeCellAt`、`normalizeHighlightRange`、`createRangeHighlightLayout`、`rangeHighlightStyle` 及对应类型，供自建网格复用同一套解析、颜色和冻结窗格几何。
- 公式引用高亮（issue #126）：编辑 `=SUM(A1:A3)+B2` 时，`XlsxFormulaBar` 与 `XlsxGrid` 的内联编辑器都会把每个引用按 Excel 调色板着色；光标所在的引用在公式栏加下划线，网格用同色边框描出当前工作表上的目标区域（活动引用更粗）。支持 A1、`$A$1`、区域、整行/整列、`Sheet2!A1` 与名称引用；字符串、函数名、外部工作簿和无法解析的名称不着色。回车/取消/失焦后清除，且不写工作簿。根入口导出 `collectFormulaReferences`、`parseFormulaReferences`、`formulaReferenceSegments`、`FORMULA_REFERENCE_COLORS`、`resolveNamedRangeTarget`、`useFormulaReferenceEditing`、`createNamedRangeResolver`、`isReferenceOnActiveSheet` 及对应类型；自定义公式输入框用这些 API 驱动同一个网格高亮会话。网格根节点暴露 `data-formula-reference-ranges`、`data-formula-reference-active` 与 `data-formula-reference-count` 供宿主和测试读取。

网格、工具栏、功能区、公式栏、工作表标签、图表/图片/选区浮层、菜单及单独渲染函数是旧的低层出口，兼容期限见第 6 节。`XlsxAuthoringDialog` 是新增的独立对话框出口，`XlsxViewer` 会自动渲染它。

### 3.5 `@arcships/vue-pdf`

稳定入口为 `PdfViewer`，以及根入口导出的属性、PDF 来源、诊断和错误类型。PDF 查看器默认创建自己专用的 `PdfRenderRuntime`，也可由宿主通过 `runtime` 属性注入实例；`pdfiumWasmUrl` 可覆盖包内默认资源地址。`createPdfRenderRuntime`、`bundledPdfiumWasmUrl`、`PdfRenderRuntimeConfig` 和页面、搜索、渲染类型是公开的实例配置入口。`DEFAULT_PDF_MAX_FILE_SIZE`、`PdfLoadOptions.maxFileSize` 和 `PdfViewer.maxFileSize` 是公开的整份文件体积配置入口；`DEFAULT_PDF_MAX_RENDER_PIXELS`、`PdfRenderRuntimeConfig.maxRenderPixels` 和 `PdfRenderError.code` 约束单页渲染预算。`password` 与 `onPassword` 为加密 PDF 提供密码通道。新代码使用 `@arcships/vue-pdf/style.css`。

根入口还提供 `createPdfTextReferenceDraft`、`createPdfPageReferenceDraft`、`createPdfRegionReferenceDraft`、`resolvePdfReference`、`describePdfReference` 和 `normalizePdfReferenceRect`。解析器通过调用方传入的 Runtime 与当前 `PdfRenderDocument` 验证字符范围或重新搜索，不上传正文，也不管理确认后的产品状态。

PDF Runtime 只接收已经取得并验证的 `ArrayBuffer`，不会再次请求原始 PDF 地址。URL 来源必须显式提供 `urlPolicy`；加载会跟随重定向并在读取正文前用同一协议与来源规则重新校验最终地址，跨来源或协议降级重定向返回 `SOURCE_NOT_ALLOWED`。`PdfVerifiedDocument.resolvedUrl`、`PdfLoadError` 与诊断中的 URL 都去除查询串、片段和内嵌凭据。默认使用包内 PDFium Worker，关闭外部字体回退；Worker 或 WASM 初始化失败会返回错误，不会把验证字节状态当成渲染成功。宿主自带严格内容安全策略时，需要允许 `worker-src blob:`，并允许读取 `@arcships/vue-pdf/assets/pdfium.wasm` 对应的同源资源。

PDF 唯一的资源拒绝配置是整份文件体积 `maxFileSize`：默认 `50 MiB`，宿主可公开调整，不叠加第二层隐藏硬上限。本地 Blob 在读取前检查声明大小；URL 有可信 `Content-Length` 时先检查响应头，否则在取得正文后、创建 PDF 引擎前检查实际字节。超过当前值返回 `PDF_TOO_LARGE`，包含 `actual` 和 `allowed`，并且不显示旧页面或创建新页面。页数、单页像素、总内存和并发页不是公开拒绝条件。

文本选择可以跨页：`PdfSelectionState` 的 `text` 是完整选区文字（多页用 `\n` 连接），`range`/`rects` 仍描述第一页，`segments` 按文档顺序给出每页的 `PdfSelectionSegment`。多页选区只用于选择和复制；PDF 引用仍是单页文本范围。`PdfRenderRuntime.getPageAnnotations?` 是可选能力，返回 `PdfRenderAnnotation`（链接、表单控件和其他批注的规范化矩形、链接目标、字段名和批注文字）。`PdfSurface` 据此渲染不可见的链接/控件命中层，并在点击时发出 `linkActivate`（`url`、`targetPageIndex`、`blocked`）或 `annotationActivate`（`kind`、`fieldName`、`fieldValue`、`contents`）。链接只接受绝对 `http`、`https`、`mailto` 地址并去掉内嵌凭据；提供了 `urlPolicy` 时还会按该策略再次校验，未通过的链接以 `blocked: true` 上报而不会自动跳转。

AcroForm 写入通过 `PdfRenderRuntime.setFormFieldValue?()` 完成：`PdfSurface` 为文本、复选框、单选框、下拉框和列表框渲染原生控件，提交后调用 Runtime、刷新页面位图并发出 `formFieldChange`（失败为 `formFieldError`）。`resolvePdfFormFieldValue()` 把 `PdfFormFieldValue`（字符串或布尔值）转换成 PDFium 表单值：文本字段接受字符串，复选框接受布尔值，单选与选择字段接受选项标签或索引。只读字段仍显示控件但被禁用；无法写入的 Runtime 不渲染控件，避免出现看似可编辑但实际无效的输入框。

批注外观编辑通过 `PdfRenderRuntime.updateAnnotationAppearance?()` 完成：`PdfAnnotationAppearancePatch` 支持 `contents`、`color`、`opacity`、`strokeColor` 和 `strokeWidth`，Runtime 重写批注并重新生成外观流，`PdfSurface.updateAnnotationAppearance(pageIndex, annotationId, patch)` 随后刷新页面位图并发出 `annotationChange`（失败为 `annotationError`）。`PdfAnnotationAppearancePatch`、`PdfFormFieldType`、`PdfFormFieldValue`、`PdfFormFieldOption` 和 `PdfFormFieldDescriptor` 是根入口类型。

无障碍文本层通过 `PdfRenderRuntime.getPageTextLayer?()` 提供：`PdfPageTextLayer` 给出 canonical 页宽高和按 PDFium 阅读顺序排列的 `PdfPageTextLayerGlyph`（`charIndex`、`text`、glyph `rect`、`unmapped`）。`PdfSurface` 与 `PdfViewer` 为每个 glyph 渲染透明 `<span>`，位置来自 glyph box，并只用一个 CSS matrix 随缩放与 0/90/180/270 旋转变换；文本层不是 `aria-hidden`，屏幕阅读器、浏览器查找和 CJK 提取均可读取，`unmapped` glyph 标记 `aria-hidden` 且不可选。文本层 `pointer-events: none`，指针选择仍由自定义 interaction layer 处理，因此跨页选择、链接/表单命中与区域选择不变。`PdfPageTextLayer` 和 `PdfPageTextLayerGlyph` 是根入口类型。

### 3.6 `@arcships/vue-ui`

稳定入口为 `OfficeObjectOutlineLayer`、`OfficeRegionSelector`、`SignaturePad`、`FileUpload`、`FileThumbnail`、`BoundingBoxCitations`、`LayoutBlocks`、`Spinner`、`Tooltip`、`OfficeFindBar`、`OfficeLiveStatus`、`useOfficeMenuKeyboard`、`resolveOfficeMenuKeyCommand`、`useViewerKeys`、`resolveViewerKeyCommand`、`isEditableViewerTarget`、`isInteractiveViewerTarget` 和 `VIEWER_KEYBOARD_SHORTCUTS`，以及根入口导出的属性和事件类型。新代码使用 `@arcships/vue-ui/style.css`。

`OfficeLiveStatus` 是所有 Office 视图共享的单一 `role=status` / `aria-live` 区域：视图只改文本，不改节点，因此加载、错误、页码和幻灯片切换都能被辅助技术重新播报。`useOfficeMenuKeyboard` 统一 DOCX 与 XLSX 右键菜单的键盘语义：打开时把焦点放到第一个 `menuitem`，上下方向键、`Home`、`End` 在菜单项之间循环，`Escape` / `Tab` 关闭菜单并把焦点还给打开它的表面；`resolveOfficeMenuKeyCommand` 是可单测的纯函数版本。

`useViewerKeys` 统一 DOCX、PDF 和 PPTX 的键盘映射：`PageUp`/`PageDown`、`Home`/`End`、`Ctrl/Cmd` + `+`/`=`/`-`/`0`、`Ctrl/Cmd + F` 和 `Escape`；`mode: "present"` 时才额外接管方向键和 `Space`。组合函数在可编辑目标（`input`、`textarea`、`select`、`contenteditable`）内不拦截按键，只有宿主处理了命令才调用 `preventDefault()`。`resolveViewerKeyCommand` 是可单测的纯函数版本，`VIEWER_KEYBOARD_SHORTCUTS` 是导出的映射清单。

两个 Office 选择原语依赖 `@arcships/office-interaction` 的公开类型。它们只消费规范化几何和受控临时状态：不解析 Office 文件，不从 DOM 猜测格式对象，也不保存引用集合。轮廓层负责候选可视化和键盘事件；区域框负责 `0..1` 坐标中的拖选与键盘调整。确认后的工具栏、引用列表和 Agent 工作流由宿主实现。

### 3.7 `@arcships/pptx-core`

根入口提供平台无关的播放类型、对象身份、能力报告、时间安排和属性轨道。浏览器中的文档会话、静态预览、播放解析和控制器从 `@arcships/pptx-core/browser` 导入。预览会话默认 `renderMode: "list"`，纵向连续展示页面；播放文档会话默认 `renderMode: "slide"`，只挂载当前页。调用方可通过 `PptxPreviewSessionOptions.renderMode` 和 `listOptions` 显式配置。补丁后的渲染器代码已经包含在浏览器入口中，消费项目不安装 `@aiden0z/pptx-renderer`。

根入口还提供 `createPptxSlideReferenceDraft`、`createPptxTextReferenceDraft`、`createPptxObjectReferenceDraft`、`createPptxRegionReferenceDraft`、`resolvePptxReference` 和 `describePptxReference`。对象适配直接复用现有 `objectKey`、shape id、source 与 group path；精确文字定位使用对象内段落和字符边界。它只生成/解析引用，不实现对象菜单、图层面板或确认后的业务动作。

无法精确执行的 PPTX 内容通过能力报告标记为近似、静态或未解析；公开说明不得宣称完整兼容 PowerPoint 全部动画。

浏览器入口还提供 `loadPptxArchive`、`PPTX_ARCHIVE_READER` 和 `resolvePptxExternalMediaUrl`。`PptxPreviewLimits` 覆盖完整的归档预算（压缩比、路径长度、XML 总量与单部件、嵌套层数、标签与文本节点、关系数）。每次加载只解压一次：`inspectOfficeArchive` 只读中央目录，渲染器完成唯一一次解压，共享校验器直接检查已解码的部件。外部媒体默认禁用；`externalMedia: "allowed"` 必须同时配合 `externalMediaUrlPolicy`，目标 URL 经协议和来源校验后才加载。

### 3.8 `@arcships/vue-pptx`

稳定入口为 `PptxViewer`、`PptxThumbnail`、`PptxStage`、`usePptxDocument`、`usePptxPlayback`、`PptxStageSelection`、`PptxStageObjectClick`、`PptxStageContextMenu` 和其他公开类型。`PptxViewer` 默认纵向连续浏览；`mode="present"` 使用单页舞台并提供下一步、上一步、暂停、继续、重播、跳页、媒体恢复和全屏。自定义静态 Surface 使用 `PptxStage` + `usePptxDocument(renderMode: "list")`，自定义播放器增加 `usePptxPlayback` 并使用 `renderMode: "slide"`。样式从 `@arcships/vue-pptx/style.css` 导入。

### 3.9 `@arcships/office-interaction`

根入口提供框架无关、零运行时依赖的 `OfficeObjectReference`、`OfficeReferenceConfirmEvent`、四种格式定位器和四维可靠度类型。运行时入口包括引用与确认事件的 `parse` / `safeParse`、规范化稳定序列化、候选排序、几何计算、短生命周期选择 reducer、候选导航状态和键盘解释器。

引用只保存文档修订、语义、格式定位器、内容指纹和规范化区域，不保存 DOM、Vue 实例、Blob URL 或截图位图。`OfficeInteractionValidationError` 的 `code` 固定为 `INVALID_OFFICE_INTERACTION`，详细原因从结构化 `issues` 读取。该包不解析 Office 文件、不渲染界面、不保存引用集合、不调用 Agent，也不修改文件；格式包负责生成定位器，宿主决定如何组织和使用确认后的引用。

根入口同时提供共享消息目录：`officeMessagesEn`、`officeMessagesZhCN`、`OfficeLocale`、`OfficeMessagesOverride`、`resolveOfficeLocale`、`resolveOfficeMessages`、`formatOfficeMessage`、`formatOfficeDate` 和 `formatOfficeNumber`。`DocxViewer`、`DocxEditor`、`DocxDocumentSurface`、`XlsxViewer`、`XlsxSheetSurface`、`PptxViewer`、`PdfViewer` 与 `PdfSurface` 均接受 `locale`（默认 `"en"`，内置 `"zh-CN"`）和 `messages` 深合并覆盖；`XlsxViewer`/`XlsxSheetSurface` 的 `locale` 同时作为数字和日期格式的默认区域（可用 `documentLocale` 单独覆盖），键入 `1.234,5` 在 `de-DE` 下解析为 1234.5 并显示为 `1.234,50`，`[$-407]` 等格式内置区域始终优先，显式 `@` 文本格式保持原样；用户可见文案、aria-label 和数字/日期显示不再与查看器语言脱节。

### 3.10 统一 Surface 引用选择

`DocxDocumentSurface`、`XlsxSheetSurface`、`PdfSurface` 和 `PptxStage` 统一接受 `documentId?: string`、`selectionMode?: "content" | "object" | "region"` 和 `emitReferenceCandidates?: boolean`。`selectionMode` 是受控输入，只解释下一次选择手势；Surface 不保存确认后的引用。

四个 Surface 共同 expose `getDocumentRevision()`、`hitTest()`、`describeReference()`、`resolveReference()`、`scrollToReference()` 和 `captureReferencePreview()`。当前预览位图采集由宿主渲染管线负责，内置实现会发出 `CAPTURE_UNSUPPORTED` 并拒绝 Promise，不会返回伪造截图。

统一事件为 `documentRevisionChange`、`referenceCandidateChange`、`referenceConfirm`、`regionDraftChange`、`selectionCancel`、`referenceResolve` 和 `referenceError`。既有 `selectionChange`、`objectClick`、`contextMenu` 等格式事件在 `0.x` 中继续保留。每个格式包从根入口重导出共享 Surface 类型，消费端无需深层导入。

### 3.11 统一 Surface 缩放

`DocxDocumentSurface`、`XlsxSheetSurface`、`PdfSurface` 和 `PptxStage` 统一接受倍率 `zoom?: number`、`enableGestureZoom?: boolean`，并发出 `update:zoom`。传入 `zoom` 后，Surface 在 50%–200% 范围内处理 `Ctrl+wheel`、trackpad pinch、两指 pointer pinch、双击缩放和 WebKit gesture；普通 wheel 不被拦截。触屏长按打开上下文菜单，原生重复 `contextmenu` 在长按后被抑制。未传 `zoom` 时保持旧缩放接口和行为。完整契约见 `docs/touch-interaction-spec.md`。

`zoom` 与旧缩放属性或 `fitWidth` 同时存在时以受控 `zoom` 为准。宿主应使用 `v-model:zoom` 持有产品级百分比、toolbar 和切文件重置；页面、幻灯片或 XLSX 单元格锚点由各 Surface 自己恢复。

## 4. 公开事件和诊断

### 4.0 四种 Office Surface

| 事件 | 参数或语义 |
|---|---|
| `documentRevisionChange` | 当前 `OfficeDocumentRevision`；加载新模型、工作簿 revision、PPTX 渲染树或 PDF 文档变化时发出 |
| `referenceCandidateChange` | 短生命周期候选和当前候选 id；只有启用 `emitReferenceCandidates` 时发出 |
| `referenceConfirm` | 已分配 `referenceId` 的 `OfficeReferenceConfirmEvent`；宿主可直接保存或交给 Agent |
| `regionDraftChange` | 页面、幻灯片或工作表区域的 `start` / `change` 草稿 |
| `selectionCancel` | 当前模式和 Escape、pointer cancel 或 programmatic 原因 |
| `referenceResolve` | 引用 id 与 `exact`、`relocated`、`ambiguous`、`not-found` 或 `unsupported` 结果 |
| `referenceError` | `OfficeReferenceError`，包含操作、格式、稳定错误码和是否可恢复 |

这些事件不定义确认后的业务动作。引用托盘、多目标角色、语言指令、提示词拼装、Agent 调用和修改后文件重载均由宿主负责。

### 4.1 DOCX

| 来源 | 事件或诊断 | 参数或关键字段 |
|---|---|---|
| `DocxViewer` | `load-start` | 无参数 |
| `DocxViewer` | `load-success` | `DocxImportResult` |
| `DocxViewer` | `load-error` | `Error`；若为导入错误，可读取公开 `code` |
| `DocxRuntimeConfig.onDiagnostic`、单次导入的 `onDiagnostic` | `load-start`、`worker-created`、`worker-success`、`worker-error`、`main-thread-start`、`main-thread-success`、`main-thread-fallback`、`aborted` | `requestId` 必有；可含 `runtimeId`、执行来源、脱敏后的资源地址、错误码、消息和耗时 |

诊断只提供执行信息，不得包含文档正文、压缩包内容或未脱敏的输入地址。诊断回调抛错不得改变加载结果。

DOCX 剪贴板公开合同：`copy()` / `paste()` 以及 `writeSelectionToClipboardData()` / `pasteClipboardData()` 使用四种 clipboard flavour——内部 MIME `application/x-arcships-docx-fragment+json`（payload 版本 1，携带来源 `DocxTextRange` 与选中的 runs）、带内联 run 样式的 `text/html`、面向只接受 RTF 的编辑器的 `text/rtf`，以及 `text/plain`。粘贴顺序固定为内部 payload → `text/html` → `text/rtf` → `text/plain`；RTF 侧只解析段落、字符格式、颜色表、`HYPERLINK` 域和 Word 列表标记，无法解析时回退纯文本；外部 HTML 只保留白名单内联声明与 `b/strong/i/em/u/s/strike/del/a/font/span`、`h1`–`h6`、`ul/ol/li`，链接必须通过 `sanitizeOfficeLink`，`script/style/iframe/object/embed/svg` 内容整段丢弃。单次 payload、HTML、RTF、段落和 run 数量都有上限（512 KiB × 3、2000、20000），超限的 flavour 会被忽略而不是解析；组合写入被拒绝时仍回退到 `writeText`。

### 4.2 XLSX

| 来源 | 事件或诊断 | 参数或关键字段 |
|---|---|---|
| `XlsxViewer` | `cellDoubleClick` | `XlsxCellAddress` |
| `UseXlsxViewerControllerOptions.onDiagnostic` | `load-start`、`load-deferred`、`load-resumed`、`load-success`、`load-error`、`load-cancelled`、`download`、`worker-error`、`image-decode-error` | `requestId` 必有；可含 `taskId`、来源、脱敏地址、字节数、图片标识和 `XlsxLoadError`。Worker 或图片失败不得静默伪装成成功 |
| `XlsxRuntimeConfig.onDiagnostic` | `worker-client-created`、`worker-client-disposed`、`runtime-disposed` | `runtimeId` 必有；可含 Worker/WASM 地址 |

XLSX 剪切合同：`cutSelection()` 把选区写入内部/系统剪贴板并进入剪切状态（`cutState = { range, sheetIndex }`），源单元格保持不动；下一次 `pasteFromClipboard()` 把值、公式和样式搬到目标，并在同一个撤销条目里清空源区域，随后 `cutState` 变为 `null`。`cancelCut()`（`Escape`）只取消剪切状态，不改工作簿；`copySelectionToClipboard()`、仅格式的 `pasteSpecialFromClipboard({ mode: "formats" })` 以及工作簿重建同样结束剪切状态。`Ctrl/Cmd+X`、右键菜单和功能区的“剪切”共用 `cutSelection()`，网格用虚线 marquee 标出 `cutState` 的源区域（仅源工作表），`XlsxSelectionOverlay` 以 `selection.cutPending`（英文/中文）向辅助技术播报。

### 4.3 `vue-pdf`

| 组件 | 事件 | 参数 |
|---|---|---|
| `PdfViewer` | `document-load-success` | 页数 |
| `PdfViewer` | `document-load-error` | `PdfLoadError` |
| `PdfViewer` | `active-page-change` | 当前页码，从 1 开始 |
| `PdfViewer` | `selectionChange` | `PdfSelectionState`；当前页拖选、双击选词、复制和 `Escape` 清除共用同一状态 |
| `PdfViewer` | `diagnostic` | `PdfDiagnostic`，加载类型为 `load-start`、`load-success`、`load-error`、`load-cancelled`；页面类型为 `render-start`、`render-success`、`render-error`；搜索类型为 `search-start`、`search-success`、`search-error`；另有 `object-url-revoked` 和 `download`。可读取 `runtimeId`、`taskId`、页码、页数、缩放、旋转、匹配数和结构化错误 |
| `PdfSurface` | `selectionChange` | `PdfSelectionState`；多页选区通过 `segments` 给出每页范围 |
| `PdfSurface` | `selectionCancel` | `{ mode, reason }`；区域层按一次 Escape 只发出一次 `escape` |
| `PdfSurface` | `linkActivate` | 页码、`annotationId`、容器坐标，以及 `url`、`targetPageIndex`、`blocked` 中的可用项 |
| `PdfSurface` | `annotationActivate` | 页码、`annotationId`、`kind`（`widget`/`annotation`）、容器坐标，以及 `fieldName`、`fieldValue`、`contents` |
| `PdfSurface` | `formFieldChange`、`formFieldError` | 写入成功的 widget（含 `fieldName`、`fieldType`、`value`、完整 annotation）或失败原因 |
| `PdfSurface` | `annotationChange`、`annotationError` | 更新后的批注与所用 `PdfAnnotationAppearancePatch`，或失败原因 |

### 4.4 `vue-ui`

| 组件 | 事件 | 参数 |
|---|---|---|
| `FileUpload` | `files-accepted` | 通过检查的 `File[]` |
| `FileUpload` | `files-rejected` | `FileUploadRejection[]`；每项包含 `code`、被拒绝的 `file` 和可显示的 `message` |
| `SignaturePad` | `update:signature` | 数据地址；空签名为 `null` |
| `BoundingBoxCitations` | `field-click` | `BoundingBoxField` |
| `LayoutBlocks` | `block-click` | `LayoutBlock` |
| `OfficeObjectOutlineLayer` | `activate`、`confirm`、`dismiss`、`navigate-hierarchy` | 候选轮廓；`confirm` 另含 `additiveRequested` 与 `penetrateRequested` |
| `OfficeRegionSelector` | `update:modelValue`、`selection-start`、`selection-change`、`selection-commit`、`selection-cancel` | 规范化点或矩形；取消可返回原区域或 `null` |

`FileThumbnail`、`Spinner` 和 `Tooltip` 当前没有业务事件。文件拒绝使用 `files-rejected`；宿主不能解析界面文字推断原因。选择控件只表达候选、确认或取消，宿主可以使用 `@arcships/office-interaction` 的临时 reducer，也可以自行管理状态。

### 4.5 `vue-pptx`

`PptxViewer` 提供 `load-start`、`load-success`、`load-error`、`slide-change`、`playback-ready`、`playback-state-change`、`step-change`、`playback-warning`、`capability`、`media-request`、`action` 和 `playback-error`。程序判断使用公开错误码和能力报告，不解析界面文案。

`PptxStage` 提供：

| 事件 | 参数 |
|---|---|
| `selection-change` | `PptxStageSelection`：`{ kind: "slide", slideIndex }` |
| `object-click` | `PptxStageObjectClick`：`{ kind: "object", slideIndex, objectKey }` |
| `context-menu` | `PptxStageContextMenu`：`kind`、`slideIndex`、可选 `objectKey`、`clientX/Y` 和 `containerX/Y` |

`PptxStageContextMenu` 是判别联合：`kind: "slide"` 时 `objectKey` 不存在，`kind: "object"` 时 `objectKey` 必填。页面编号从零开始。普通对象点击先选择所在幻灯片，再发出对象事件；右键只发出上下文菜单事件。`objectKey` 是不透明稳定标识，调用方只能传回公开控制器，不依赖其内部编码。

## 5. 公开错误

| 包 | 类型 | 错误码 | 含义 |
|---|---|---|---|
| `office-interaction` | `OfficeInteractionValidationError` | `INVALID_OFFICE_INTERACTION` | 引用或确认事件不满足 schema；具体问题由 `issues[].code`、`path` 和 `message` 表达 |
| `docx-core` | `DocxImportError` / `DocxImportErrorCode` | `ABORTED` | 调用方取消或旧任务失效 |
|  |  | `STALE_RESULT` | 低层导入结果已过期；Runtime 对外通常归一为 `ABORTED` |
|  |  | `SOURCE_NOT_ALLOWED` | 地址或来源不符合公开规则 |
|  |  | `FETCH_FAILED` | 受控获取失败 |
|  |  | `LIMIT_EXCEEDED` | 超过当前输入限制 |
|  |  | `IMAGE_LIMIT_EXCEEDED` | 图片字节、尺寸、像素总量或并发解码超过实例限制 |
|  |  | `INVALID_IMAGE` | 图片文件头或格式无效 |
|  |  | `IMAGE_DECODE_FAILED` | 浏览器或解析器无法解码已验证图片 |
|  |  | `WORKER_UNAVAILABLE` | Worker 无法创建或不可用 |
|  |  | `WASM_LOAD_FAILED` | WASM 地址、字节或初始化失败 |
|  |  | `PARSE_FAILED` | 文件无法解析为受支持的 DOCX |
|  |  | `TIMEOUT` | 解析超过实例配置的时间上限 |
|  |  | `RUNTIME_DISPOSED` | 已销毁的实例被再次使用 |
| `xlsx-core` / `vue-xlsx` | `XlsxLoadError` / `XlsxLoadErrorCode` | `SOURCE_NOT_ALLOWED`、`FETCH_FAILED`、`INVALID_WORKBOOK`、`LIMIT_EXCEEDED`、`IMAGE_LIMIT_EXCEEDED`、`INVALID_IMAGE`、`IMAGE_DECODE_FAILED`、`TIMEOUT`、`WORKER_UNAVAILABLE`、`ABORTED` | 分别表示来源、获取、格式、资源、图片、超时、Worker 和取消错误；资源错误可含 `phase`、`limit`、`actual`、`allowed` |
| `xlsx-core` | `XlsxWorkerError` / `XlsxWorkerErrorCode` | `LIMIT_EXCEEDED`、`TIMEOUT`、`INVALID_SOURCE`、`WORKER_FAILED`、`SUPERSEDED` | Worker 内资源超限、超时、输入无效或执行失败，或加载被更新的一次加载取代；资源、输入和串行队列导致的失败会终止该 Worker、不静默转主线程，`SUPERSEDED` 只拒绝被取代的加载 |
| `xlsx-core/runtime` | `XlsxRuntimeError` / `XlsxRuntimeErrorCode` | `RUNTIME_DISPOSED` | 已销毁的 XLSX Runtime 被再次使用；错误同时保留 `format: "xlsx"` 和 `runtimeId` |
| `vue-xlsx` | `XlsxFileSizeLimitExceededError` | 类的公开字段 | 本地输入超过控制器允许的文件大小 |
| `vue-pdf` | `PdfSourceError` / `PdfLoadErrorCode` | `SOURCE_NOT_ALLOWED`、`FETCH_FAILED`、`INVALID_PDF`、`PDF_TOO_LARGE`、`ENCRYPTED`、`ABORTED` | 分别表示来源被拒绝、获取失败、PDF 无效、整份文件超过当前 `maxFileSize`、文档需要密码和取消；体积错误同时提供 `actual`、`allowed` |
| `vue-ui` | `FileUploadRejection` / `FileUploadRejectionCode` | `FILE_TYPE_NOT_ACCEPTED`、`FILE_TOO_LARGE` | 文件类型不符合 `accept`，或文件超过 `maxSize` |
| `pptx-core` / `vue-pptx` | `PptxPreviewError`、`PptxPlaybackError` | `ABORTED`、`STALE_RESULT`、`LIMIT_EXCEEDED`、`INVALID_SOURCE`、`PARSE_FAILED` 及播放错误码 | 加载、资源限制、解析、对象匹配、媒体和播放控制错误 |

公开错误必须保留稳定的 `code`，界面文案可以改进但不能替代错误码。地址写入错误和诊断前必须脱敏。Worker、WASM 或解析失败不得静默转成成功；只有调用方明确开启且文件小于已配置上限时，DOCX 才能进行主线程回退，并必须发出 `main-thread-fallback` 诊断。

## 6. 兼容登记

登记字段中的 `kind` 只有 `public-api` 和 `internal-bridge` 两种。负责人是代码所有者，不表示某次会话的独立验收人。

| kind | 入口或实现 | 引入版本/日期 | 负责人 | 替代入口 | 删除期限 |
|---|---|---|---|---|---|
| `public-api` | `@arcships/docx-core` 根入口的 `setWasmSource` | `<=0.1.0` | `docx-core` Runtime | `createDocxRuntime({ wasmSource })` 或 `{ wasmUrl }` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `public-api` | `@arcships/xlsx-core` 根入口的 `setWasmSource`、`canUseConfiguredWasmSourceInWorker`、`getConfiguredWorkerWasmSource`、`getSheetsWasmModule` | `<=0.1.0` | `xlsx-core` Runtime | `createXlsxRuntime({ wasmSource, createWorker })` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `public-api` | `DocxEditorViewer` | `<=0.1.0` | `vue-docx` | `DocxEditor` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `public-api` | `vue-docx` 已导出的页面、段落、表格、图片、字段、修订、菜单、工具栏、缩略图和拖放组件 | `<=0.1.0` | `vue-docx` | `DocxViewer`、`DocxEditor` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `public-api` | `vue-docx` 的静态渲染函数和页面 surface registry | `<=0.1.0` | `vue-docx` | `DocxViewer`、`DocxEditor` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `public-api` | `vue-docx` 的 `useDocxViewerThumbnails`、`UseDocxViewerThumbnailsOptions`、`DocxViewerThumbnails` | `<=0.1.0` | `vue-docx` | `useDocxPageThumbnails`、`UseDocxPageThumbnailsOptions`、`UseDocxPageThumbnailsResult` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `public-api` | `vue-xlsx` 已导出的低层组件、`MemoChartSvg`、`MemoSurfaceChartComposite` 及四个对应渲染类型 | `<=0.1.0` | `vue-xlsx` | `XlsxViewer`、`useXlsxViewerController` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `internal-bridge` | `docx-core/src/viewer/wasm-source.ts` 的 `createLegacyWorkerWasmBridge` / `legacyWorkerWasmBridge` | `2026-07-10` | `docx-core` Runtime | 由正式默认 `DocxRuntime` 适配器接管；公开旧入口继续保留 | `removeByDate: 2026-08-09` |
| `internal-bridge` | `docx-core/src/engine/wasm.ts` 的 `createLegacyDefaultWasmRuntime` / `legacyDefaultWasmRuntime` | `2026-07-10` | `docx-core` Runtime | 新代码使用 `createDocxRuntime`；公开 `initWasm` / `setWasmSource` 在整个 `0.x` 继续保留 | `removeByDate: 2026-08-09` |
| `internal-bridge` | `xlsx-core/src/wasm.ts` 的 `createLegacyDefaultWasmRuntime` / `legacyDefaultWasmRuntime` | `2026-07-10` | `xlsx-core` Runtime | 由正式默认 `XlsxRuntime` 适配器接管；公开旧入口继续保留 | `removeByDate: 2026-08-09` |

如果内部桥接到期时仍有必要保留，必须在到期前把原因、负责人和新日期写回本表；不能用内部桥接的期限提前删除公开入口。

## 7. 发布门禁

每次改变公开接口或发布资源时至少检查：

1. 从当前源码构建，并对九个公开包执行真实 `npm pack`。
2. 在工作区外只安装生成的九个压缩包，不使用源码别名，不复制 demo 中的 Worker/WASM。
3. 检查本页列出的所有 `exports` 都能解析，未列出的代表性深层路径都被拒绝。
4. 检查根入口、`./core`、`./runtime`、`./wasm-url`、四个 `style.css`、两个独立 Worker 资源、三个 WASM，以及 PDF 包内生成的 Worker 代码。
5. 检查生成的 `.d.ts` 保留本页登记的 `@deprecated`，并且公开声明不引用私有 `@arcships/office-runtime` 或不受支持的深层路径。
6. 用正式 preview 验证 Worker/WASM 请求来自压缩包、无 404、WASM 类型和字节签名正确。
7. 黑盒消费页只通过本文公开入口、公开事件和公开错误观察结果，不读取模块私有变量。

本会话的开发后自测只能作为同会话复测；候选发布仍需另一个全新会话执行 `BB-RELEASE` 独立复核。
