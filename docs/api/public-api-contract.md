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
|  | `./tokens.css` | 不适用 | `./dist/tokens.css` |
|  | `./theme.css` | 不适用 | `./dist/theme.css` |
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
- 根入口公开分节分栏（issue #98）：`readSectionColumns(sectionPropertiesXml)` 返回 `{ count, spaceTwips, widthsTwips? }`（缺省值为 Word 的单栏、720 twips 间距），`writeSectionColumns(sectionPropertiesXml, { count, spaceTwips })` 改写 `w:cols` 并保留其他 `w:sectPr` 子元素（只改间距时保留不等宽 `<w:col>` 宽度，改栏数时清除旧宽度，栏数限制在 1–45），`setSectionColumns(model, sectionIndex, update)` 同时写入 `metadata.sections`、中间分节段落与正文级 `sectPr`。`resolvePageSetup()` 的每个 `DocxSectionPageSetup` 新增 `columns`。

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

OpenDocument 文本只读导入（issue #144）：`@arcships/docx-core` 的运行时入口在字节层识别 ODF 文本（`.odt`、`.ott`、`.odm`、`.otm`、`.oth`、扁平 `.fodt`），读取 `content.xml`/`styles.xml` 和包内图片，返回不带 Word 包的 `DocModel`。`buildDocModelFromBytes()`、`importDocxBuffer()` 和 `DocxRuntimeLoader.load()` 都走同一分支，`parseDocx()` 仍只接受 OOXML 并对 OpenDocument 给出明确错误。模型在 `metadata.sourcePartNames.document` 记录 `content.xml`；根入口导出 `isOpenDocumentModel(model)` 和 `OPEN_DOCUMENT_VIEW_ONLY_REASON` 供宿主判断与展示。该类模型只能渲染，`exportDocxBuffer()` 会以 `INVALID_MODEL` 拒绝，不能由 DOCX 写入器序列化。

根入口公开 DOCX 源文档图（issue #131）：`inspectDocumentSource(bytes)` 接受 `ArrayBuffer | Uint8Array`，返回 `DocumentSourceGraph`（`schemaVersion`、`documentId`、`parts`、`assets`、`warnings`、`preservedFeatures`）。每个 `parts[]` 条目带 `name`、`byteLength`、`root`（`document`/`element`/`text`/`comment`/`processing_instruction` 节点，含 UTF-8 字节范围 `byteStart`/`byteEnd`、命名空间、属性和子节点）以及解析失败时的 `error`；二进制资源列在 `assets[]` 且 `root` 为 `null`。`preservedFeatures[]` 按部件统计 `altChunk`、`subDoc`、`object`、`control`、`movie`、`sound`、`diagram` 的 `preserved_only` 以及 `oMath` 的 `approximate_rendering` 出现次数，排序稳定。该实现不依赖新版 WASM 导出：即使包内 `docx_wasm_bg.wasm` 没有 `inspect_document_source`，图也由 TypeScript 的 ZIP/XML 索引生成；格式错误的 XML 只写入对应部件的 `error`，损坏的 ZIP 或 OpenDocument 输入会以明确错误拒绝。

导入源文档图（issue #131）：`DocxImportOptions.onSourceGraph` 与 `DocxRuntimeConfig.onSourceGraph` 在成功导入后接收 `DocumentSourceGraph`。只有配置了回调才会构建图，并由 Worker 在成功响应中带回（主线程路径直接使用已解析的包），不会改变导入结果；回调抛出的异常被隔离。`@arcships/docx-core/runtime` 同时导出 `DocumentSourceGraph`、`DocumentSourcePart`、`DocumentSourceAsset`、`SourceElement`、`SourceAttribute` 和 `PreservedFeature` 类型。demo 的 DOCX 运行配置面板在“源文档图：保留 vs 近似”区域显示部件字节范围、`preserved_only` 与 `approximate_rendering` 分组。

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
- 旧版索引色（issue #105）：`parseWorkbookStyles()` 读取 `xl/styles.xml` 的 `<colors><indexedColors>`，`buildThemePalette()` 将其与 ECMA-376 18.8.27 默认索引调色板（0–63 及系统前景/背景 64/65）合并到新的 `XlsxThemePalette.indexedColors`；`resolveWorkbookColor()` 现在解析 `{ indexed }`（数字或数字字符串），并按与主题色相同的 HSL 规则应用 `tint`；缺失、越界或非整数索引返回 `null`。填充、字体、边框、条件格式、迷你图、缩略图和打印预览共用该解析器。
- 图表：`refreshChartDataFromWorkbook(workbook, chartsByWorkbookSheetIndex, chartsheets)` 从编辑后的工作簿重新解析图表分类和值，保留图表定义与锚点；Worker 客户端也提供 `refreshCharts()`，不会重新解析源文件字节。
- 图表内嵌图片与 `chartUserShapes`（issue #140）：`XlsxChart` 新增 `images`（关系 id → data URL）、`drawings`（关系 id → `{ xml, images, shapes }`）、`chartAreaFillImage`、`plotAreaFillImage`，`XlsxChartSeries.fillImage` 与 `XlsxChartPointStyle.fillImage` 保存系列/数据点的 `a:blipFill` 图片填充。`readChartEmbeddedMedia(archive, chartPath, options)` 读取图表部件的 `/image` 与 `/chartUserShapes` 关系（含 shapes 部件自己的图片），只接受按签名识别出的 PNG/JPEG/GIF/WebP/BMP/TIFF 字节并生成 data URL：外部关系（`TargetMode="External"`）、带 scheme 的目标、SVG/非图片字节一律丢弃；单张图片上限 `CHART_EMBEDDED_IMAGE_MAX_BYTES`（72 MB）、单图表图片总量 `CHART_EMBEDDED_IMAGE_TOTAL_MAX_BYTES`（100 MB）、`chartUserShapes` 部件 `CHART_USER_SHAPES_MAX_BYTES`（8 MB），三者均可用 `options` 收紧。`parseChartUserShapes(xml, { extentEmu, images, themePalette })` 把形状锚点归一化为相对图表区域的 0–1 矩形（`XlsxChartUserShape.rect`），并解析预设几何（`geometry`、`geometryAdjustments`）、填充/描边与文本段落（复用 `XlsxShapeParagraph` / `XlsxShapeTextRun`）。SVG 渲染器把图片填充画成 `<pattern>` + `<image>`，并按预设几何把文字框和图片叠加在绘图区之上；导出/重新打开时图表部件、`chartUserShapes` 部件与其媒体通过绘图关系图完整保留。
- 内置图表样式（issue #142）：经典 `<c:style>` 1–48 / 101–148 会在没有自定义 `chartColorStyle` 时按 Excel 调色板填充 `chartColorPalette` 和系列颜色（灰度、6 级灰、accent 单色阶梯与彩色变体，含线性 RGB 的 tint/shade）；显式 `chartColorStyle` 部件和系列 `spPr` 填充优先，带 `fillImage` 的系列/数据点仍保留图片填充。三维曲面图保留专用调色板，并从同一套内置样式取墙面/底面填充和主网格线色；`XlsxChartAxis` 新增只读 `majorGridlineColor` 供自定义渲染层使用。
- 范围高亮（issue #132）：控制器新增 `highlightRanges(ranges, { color, sheetIndex, role, id, label })`、`clearHighlightedRanges()` 和响应式 `highlightedRanges`。`XlsxRangeHighlight`（`range`、`color`、`sheetIndex`、`role`、`seriesIndex`、`seriesName`、`id`、`label`）与 `XlsxRangeHighlightRole`（`category`、`value`、`bubbleSize`、`seriesName`、`custom`）随根入口导出；`highlightRanges()` 替换宿主区域并默认落在活动工作表/琥珀色，图表选择派生区域由选择状态驱动并在取消选择时清除，两者合并到 `highlightedRanges`。该 API 只读，不写工作簿。`XlsxViewer`/`XlsxSheetSurface` 挂载 `XlsxRangeHighlightLayer`，按冻结窗格拆段、跳过隐藏行列，并跟随滚动与缩放。
- 更改图表类型：`changeWorkbookChartType(assets, chartAssets, chartId, type)` 重写图表部件（经典 `c:plotArea` 或 chartEx `cx:chartSpace`），保留系列名称、颜色、数据标签和坐标轴格式；经典与 chartEx 互转时同步 drawing 关系、`mc:AlternateContent` 包裹和内容类型。`chartTypeOptions`、`chartTypeFamilies()`、`resolveChartTypeOption()`、`matchesChartType()`、`chartTypeUnavailable()` 提供菜单目录和可用性判断；`XlsxViewerController.changeChartType(id, type)` 是其控制器入口。
- 根入口同时提供工作表、单元格、A1 范围、整行、整列、图表和人工 sheet 区域的 `createXlsx*ReferenceDraft`、`resolveXlsxReference` 与 `describeXlsxReference`。适配器优先使用 sheet id 和 drawing id；名称/类型证据只有唯一时才用于迁移。
- 单元格文字旋转与缩小适应（issue #141）：`XlsxSheetData.cellTextLayouts` 以 A1 地址保存显式样式里带 `textRotation`（0–180，91–180 表示顺时针）或 `shrinkToFit` 的单元格，主线程与 Worker 快照都会带回；网格、缩略图和打印预览按 Excel 角度语义旋转、按旋转后的包围盒裁剪，并在 `shrinkToFit` 时二分缩小到基础字号的 0.4 倍。
- 富文本单元格（issue #138）：`XlsxCellTextRun` / `XlsxCellTextRunFont` 描述单元格内每一段文字的文本与字体覆盖（粗体、斜体、下划线、删除线、颜色、字号、字体族、`vertAlign` 上标/下标），`XlsxSheetData.richTextCells` 以 A1 地址保存多段单元格，`controller.getCellRuns(cell)` 返回当前单元格的段（值变化后自动不再返回）。`parseRichTextRuns`、`parseInlineRichTextCell`、`parseSharedStringRichText`、`richTextRunsToPlainText` 与 `resolveRichTextRuns` 供自定义渲染复用同一套解析与"文本仍匹配"校验；单元格的纯文本 `value` 始终是各段拼接，公式、搜索和剪贴板继续使用它。
- RTL/bidi（issue #112）：`XlsxSheetData.rightToLeft` 与 `WorkbookSheetState.rightToLeft` 保存 `<sheetView rightToLeft="1">`，主线程和 Worker 工作表都用它镜像列、列头和冻结列；纯几何函数 `axisLayoutOffset()`、`axisLayoutBoundary()`、`axisLayoutScroll()`、`mirrorBodyOffset()` 和 RTL 感知的 `resolvePaneOverlayRect()` 位于 `vue-xlsx` 的 `grid-pane-geometry.ts`。DOCX 的 `ParagraphStyle.direction`（`w:bidi` / `w:rtl`）继续通过 `dir` 和 `paragraphTypographyStyle()` 输出；PPTX 播放能力报告新增 `text:rtl-paragraph`（`strict`）与 `text:rtl-columns`（`static`），渲染后的段落带 `dir`。三种格式均为渲染语义新增，不改变单元格/段落文字模型或导出字节。
- 预设形状几何与连接线（issue #143）：`buildPresetShapePath(geometry, adjustments, { width, height })` 返回 `{ path, viewBox }`，覆盖 ECMA-376 的 `a:prstGeom` 预设（箭头、V 形、星形、标注、流程图、括号/花括号、月牙、齿轮、云、弧形等）；`adjustments` 接受 `Record<string, number>` 或 `PresetShapeAdjustments`（`a:avLst` 的 1/100000 单位）。`XlsxShape.svgPath`/`svgViewBox` 现在对预设几何也会写入，连接线的 `stroke.dash`、`headEndType`、`tailEndType` 继续保留；`isStrokeOnlyShapeGeometry(geometry)` 标记不填充的开放路径（连接线、弧线、括号/花括号）。`XlsxDrawingLayer` 用同一函数按当前尺寸重建路径，并渲染 `prstDash` 虚线与 `headEnd`/`tailEnd` 端点箭头。
- 选区汇总（issue #111）：`summarizeSelectionValues(values, { totalCells, scannedCells, numberFormat })` 把主线程和 Worker 读取的非空单元格汇总为 `XlsxSelectionSummary`（`cellCount`、`count`、`numericCount`、`sum`、`average`、`min`、`max`、`numberFormat`、`truncated`），使用补偿求和；`XLSX_SELECTION_STATS_CELL_LIMIT`（10,000）是统一扫描上限。`XlsxSelectionStatsQuery` 描述 Worker 读取窗口（含 `hiddenRows` 与 `cellCount`），`XlsxWorkerClient.getSelectionStats()` / `controller.getSelectionStatsAsync()` 只返回聚合结果，不传输单元格。

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
- OpenDocument 文本只读界面（issue #144）：`DocxViewer`、`DocxEditor` 和 `DocxViewerRoot` 的文件选择器/拖放接受 `.odt`、`.ott`、`.odm`、`.otm`、`.oth`、`.fodt`，并在打开后显示 `docx-view-only-reason` / `docx-editor-view-only-reason`。`DocxEditorController.viewOnlyReason` 非空时所有改模型的事务被拒绝，`exportDocx()`、`getDocxBytes()` 和 `getDocxBlob()` 也拒绝；查看器下载按钮禁用。根入口导出 `OPEN_DOCUMENT_VIEW_ONLY_REASON`、`OPEN_DOCUMENT_FILE_EXTENSIONS`、`OPEN_DOCUMENT_FILE_MIME_TYPES`、`SUPPORTED_DOCUMENT_FILE_EXTENSIONS`、`SUPPORTED_DOCUMENT_FILE_MIME_TYPES`、`isWordFileName()`、`isOpenDocumentTextFileName()`、`isSupportedDocumentFileName()`、`isSupportedDocxFileName()` 和 `DOCX_FILE_ACCEPT`。
- 浮动表格（issue #150）：`DocxEditorController.setTableFloating(tableIndex, update)` 把一个正文表格放到落下位置（写 `TableStyle.floating`，导出为 `w:tblpPr`），传 `null` 把它放回文字流，一次调用一条撤销记录。`DocxTableHost` 的表格移动手柄（`data-docx-table-move-handle`，拖拽时显示幽灵框 `data-docx-table-drag-ghost`）可用 `:show-table-move-handle="false"` 关闭；`DocxTableFloatingUpdate`、`DocxTableFloatingDrop`、`resolveTableFloatingDropPatch()` 和 `mergeTableFloatingStyle()` 是可复用的规则。
- 图形对象与公式（issue #33）：`DrawingRunNode`（`type: "drawing"`，`kind` 为 `textBox` / `shape` / `wordArt` / `group` / `smartArt` / `chart` / `oleObject` / `model3d`）与 `MathRunNode`（`type: "math"`）是新的段落子节点类型；`ImageRunNode` 新增 `textBox` / `textBoxSourceXml`。导入侧的 `paragraphDrawingObjects()`、`classifyDrawingXml()`、`parseChartData()`、`linearizeOmml()`、`parseDrawingAnchor()` 和导出侧的 `applyPreservedDrawingObjectsToDocumentXml()` / `...ToPartContent()` / `...ToPackage()`、`collectPreservedDrawingObjects()` 都是公开的纯函数，与内置序列化补丁共用。`@arcships/docx-core` 导出 `drawingObjectPlaceholderStyle()`、`drawingObjectFrameStyle()`、`drawingObjectSizePx()`、`chartPlaceholderText()`、`chartPreviewDataUri()` 供自定义绘图层使用；文案在 `messages.docx.drawing` 下。`renderParagraphRuns()` 新增 `renderDropCaps`（默认 `true`）和 `drawingLabels` 选项。
- 类型：`DocModel`、`DocxEditorController`、`DocxFormFieldUpdate`、编辑选区、图片、表格、分页、修订、批注及上述组合函数的公开类型。
- 样式：`@arcships/vue-docx/style.css`。
- 页面布局（issue #110）：`DocxViewer`、`DocxDocumentSurface` 和 `DocxViewerRoot` 接受 `viewMode`（`continuous` / `single` / `spread`）与 `fitMode`（`width` / `page`）；`DocxViewer` 另有 `defaultViewMode`、`defaultFitMode`，并抛出 `update:viewMode`、`update:fitMode`。`fitWidth` 保留为兼容入口。根入口导出几何纯函数 `resolveDocxViewLayout`、`docxPageRows`、`docxRowIndexAtOffset`、`docxRowIndexForPage`、`docxRowsInViewport`、`docxSnapTopForRow`、`docxPageAtPoint`、`docxSpreadStepPage`、`isDocxSpreadMode`、`normalizeDocxViewMode`、`normalizeDocxFitMode` 与常量 `DOCX_VIEW_MODES`，类型 `DocxViewMode`、`DocxFitMode`、`DocxViewLayout`、`DocxViewLayoutOptions`、`DocxPageRowLayout`、`DocxPageDimension`。三种模式共用同一套滚动坐标，页面偏移、分页锚点、书签和搜索定位不变。
- 历史记录：`UseDocxEditorOptions.historyMaxEntries` 和 `historyMaxBytes` 同时生效；Vue 编辑器默认保留最多 100 条、估算 32 MiB，超过后优先淘汰最旧记录。
- 分节分栏（issue #98）：`DocxEditorController.setSectionColumns(sectionIndex, { count, spaceTwips? }, options?)` 是光标所在或指定分节的栏数命令（一次撤销记录）；`useDocxPageSetup()` 新增 `columns` / `applyColumns(count)`，工具栏选择器 `data-testid="editor-columns"` 提供 1/2/3 栏并跟随光标分节。查看器与编辑器都按快照的逐栏分段渲染，`w:space`、显式分栏符与末页栏平衡与 Word 一致；跨栏段落的每个行段仍可编辑，但只显示该栏的行。

页面、段落、表格、图片、工具栏和浮层的单独组件虽然在旧版本已经导出，但不再作为新功能的接入点。它们的兼容期限见第 6 节。

### 3.4 `@arcships/vue-xlsx`

稳定高层入口为：

- `XlsxViewer`。
- `useXlsxViewerController`、`useXlsxViewerThumbnails`。
- `XLSX_FILE_EXTENSIONS`、`XLSX_FILE_MIME_TYPES`、`XLSX_FILE_ACCEPT`、`xlsxFileExtension`、`isSupportedXlsxFileName`（上传入口和拖放共用的类型策略）。

- `XlsxStatusBar`（选择状态栏；`XlsxViewer` 默认挂载，`showStatusBar="false"` 关闭）。
- `useXlsxViewerController`、`useXlsxViewerThumbnails`、`useXlsxSelectionStats`、`formatSelectionStatValue`。
- `XlsxFileSizeLimitExceededError`。
- `XlsxDiagnostic`、`XlsxLoadError`、`XlsxLoadErrorCode`、`XlsxSourceKind`、`XlsxSourceState`、`XlsxUrlPolicy`。
- 样式：`@arcships/vue-xlsx/style.css`。
- 可选功能：`@arcships/vue-xlsx/chart`、`@arcships/vue-xlsx/map`、`@arcships/vue-xlsx/webgl`。根入口保留旧渲染名称的异步兼容外观；新代码可显式导入对应子入口。未使用功能时不会请求其实现分包。
- 工作表缩略图：`useXlsxViewerThumbnails` 只在需要时读取单元格（首次悬停或显式预热），`prefetchThumbnails` 在空闲时预热指定工作表、按 `revision` 去重、单次最多 3 张，`paintThumbnail` 随后只绘制。
- 历史记录：控制器选项 `historyMaxEntries` 和 `historyMaxBytes` 同时生效，默认最多 100 条、估算 32 MiB。
- 选择状态栏（issue #111）：`XlsxViewer` 在网格下方渲染 `XlsxStatusBar`，显示单元格/区域地址、非空计数、数值计数、求和、平均值、最小值/最大值，并用第一个数值单元格的数字格式渲染合计；自动筛选隐藏的行不计入合计并显示 `Filtered`。扫描上限为 `XLSX_SELECTION_STATS_CELL_LIMIT`（10,000），更大的选区显示 `…`。自建界面可挂载 `XlsxStatusBar`，或用 `useXlsxSelectionStats(controller)` 读取 `summary`、`filtered`、`pending`，并用 `formatSelectionStatValue(value, numberFormat, locale)` 复用同一套数字格式渲染。Worker 只读视图通过 `controller.getSelectionStatsAsync()` 在 Worker 内聚合，不传输单元格。
- 网格渲染模式（issue #107）：`XlsxViewer`、`XlsxSheetSurface` 和 `XlsxGrid` 接受 `renderMode: "canvas" | "dom"`，默认 `"canvas"`。`"dom"` 只把可见窗口（含冻结窗格与行列表头）渲染为 `role="row"`/`role="gridcell"` 元素，文字是真实文本节点，可被浏览器查找、原生选择与屏幕阅读器读取，且不创建 canvas；两种模式共用 `resolveCanvasCellStyle` 的字体/填充/对齐/边框、条件格式、合并区、超链接、批注、迷你图、格式化显示值与 `getCellStyle` 覆盖，切换模式保留滚动位置和控制器选区。DOM 模式按单元格裁剪文本，按住 Alt 拖动选择文字。`XlsxGridRenderMode` 从根入口导出。
- 图表数据：`refreshChartData()` 在主线程工作簿或 Worker 当前工作簿上重新解析图表分类和值，用于源单元格编辑后刷新绘图。
- 重新加载本地工作簿：`XlsxViewerController.loadFile(source: XlsxLoadFileSource, fileName?)` 接受 `ArrayBuffer` 或任何带 `arrayBuffer()` 的 `File`/`Blob`。`XlsxViewer` 的拖放导入调用它，与 `file` 选项共用大小上限、URL 规则、延迟解析阈值和 Worker 路径；原来的 URL 来源被替换，加载结束前 `displayFileName`、`sourceState` 和 `sourceError` 与普通加载一致地更新。
- 编写入口：`XlsxViewer` 的功能区与右键菜单打开 `XlsxAuthoringDialog`（图表、批注、超链接、表格、条件格式、数据验证）。宿主也可以直接挂载该组件，或只调用控制器上的 `insertChart`、`setCellComment`、`setCellHyperlink`、`addConditionalFormatRule`、`updateConditionalFormatRule`、`deleteConditionalFormatRule`、`createTable`、`removeTable`、`setSelectionDataValidation`、`setFreezePanes`、`toggleFreezePanesAtActiveCell`。每个操作都写入工作簿模型、记录撤销条目，并随导出的 XLSX 一起保存；`XlsxViewerAuthoring` 汇总这些方法的说明。直接读取/清空用 `getCellComment`、`getCellHyperlink`、`getConditionalFormatRules`、`getTableAtCell`、`freezePanes`。
- 编写限制：这些命令需要可修改的主线程工作簿（Worker 只用于强制只读的大文件），在 `readOnly`、受保护工作表或 Worker 只读模式下返回 `false`/`null`。
- 查找替换（issue #113）：`replaceInSheet(query, replacement, { matchCase, wholeCell, formula, sheetIndex, range, occurrence })` 在同一工作表内替换值和公式文本，`formula` 默认开启，非活动工作表会先被激活，整批写入只占一条撤销记录，返回 `XlsxReplaceResult`（`replacedCells`、`replacedOccurrences`、`changed`）。只读/受保护/Worker 只读模式不改工作簿，返回 `changed: false`。搜索与替换面板、`useXlsxSearch()` 与 `OfficeFindBar` 的替换行复用同一个控制器命令。
- 单元格格式：功能区提供边框预设下拉（全部/外侧/左/右/上/下/无）与边框颜色，并保留“所有边框”快捷按钮；宿主也可以直接调用 `setSelectionBorders({ preset, style, color })`，`XlsxBorderPreset`、`XlsxBorderPresetInput` 描述这组入参。`all` 给选区内每格画边框，`outside` 只画选区外围（内部格子保持原样），`none` 清除四边；颜色同时写入每一面，`XlsxCellBorderStyleInput` 仍支持宿主用 `setRangeStyle` 逐边设置不同颜色。
- 合并：`mergeAndCenterSelection()` 是功能区的默认合并动作，合并后同时写入 `horizontal: center` 与 `vertical: center`，两步共用一条撤销记录；单格选区只居中，不写一个格的合并区。`mergeSelection()` 保留给只需合并、不改对齐的宿主，`unmergeSelection()` 取消合并。
- 范围与图表数据源高亮（issue #132）：`XlsxRangeHighlightLayer` 是根入口新增的浮层出口，`XlsxViewer`/`XlsxSheetSurface` 已自动挂载；`controller.highlightRanges()`/`clearHighlightedRanges()`/`highlightedRanges` 读写范围，选中图表/系列/数据点/图例时按 category/value/bubbleSize 角色自动高亮源区域。根入口同时导出 `CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE`、`XLSX_RANGE_HIGHLIGHT_DEFAULT_COLOR`、`collectChartRangeHighlights`、`parseChartSourceRange`、`parseChartSourceA1Range`、`chartRangeCellAt`、`normalizeHighlightRange`、`createRangeHighlightLayout`、`rangeHighlightStyle` 及对应类型，供自建网格复用同一套解析、颜色和冻结窗格几何。
- 公式引用高亮（issue #126）：编辑 `=SUM(A1:A3)+B2` 时，`XlsxFormulaBar` 与 `XlsxGrid` 的内联编辑器都会把每个引用按 Excel 调色板着色；光标所在的引用在公式栏加下划线，网格用同色边框描出当前工作表上的目标区域（活动引用更粗）。支持 A1、`$A$1`、区域、整行/整列、`Sheet2!A1` 与名称引用；字符串、函数名、外部工作簿和无法解析的名称不着色。回车/取消/失焦后清除，且不写工作簿。根入口导出 `collectFormulaReferences`、`parseFormulaReferences`、`formulaReferenceSegments`、`FORMULA_REFERENCE_COLORS`、`resolveNamedRangeTarget`、`useFormulaReferenceEditing`、`createNamedRangeResolver`、`isReferenceOnActiveSheet` 及对应类型；自定义公式输入框用这些 API 驱动同一个网格高亮会话。网格根节点暴露 `data-formula-reference-ranges`、`data-formula-reference-active` 与 `data-formula-reference-count` 供宿主和测试读取。

网格、工具栏、功能区、公式栏、工作表标签、图表/图片/选区浮层、菜单及单独渲染函数是旧的低层出口，兼容期限见第 6 节。`XlsxAuthoringDialog` 是独立对话框出口，`XlsxViewer` 会自动渲染它；0.8 起它渲染在共享 `OfficeDialog`（Reka）上，表单状态、校验、Esc/backdrop/关闭后焦点归还保持不变。`XlsxHeaderMenu` 与 `XlsxAutoFilterMenu`（issue #160）是新增的独立菜单出口，分别承载行/列表头菜单与 AutoFilter 值面板，均渲染在共享 `OfficeDropdownMenu`（Reka）上并保留原有 `data-testid`；`XlsxGrid` 会自动挂载它们，只嵌入网格的宿主也可以直接使用。

### 3.5 `@arcships/vue-pdf`

稳定入口为 `PdfViewer`，以及根入口导出的属性、PDF 来源、诊断和错误类型。PDF 查看器默认创建自己专用的 `PdfRenderRuntime`，也可由宿主通过 `runtime` 属性注入实例；`pdfiumWasmUrl` 可覆盖包内默认资源地址。`createPdfRenderRuntime`、`bundledPdfiumWasmUrl`、`PdfRenderRuntimeConfig` 和页面、搜索、渲染类型是公开的实例配置入口。`DEFAULT_PDF_MAX_FILE_SIZE`、`PdfLoadOptions.maxFileSize` 和 `PdfViewer.maxFileSize` 是公开的整份文件体积配置入口；`DEFAULT_PDF_MAX_RENDER_PIXELS`、`PdfRenderRuntimeConfig.maxRenderPixels` 和 `PdfRenderError.code` 约束单页渲染预算。`password` 与 `onPassword` 为加密 PDF 提供密码通道。新代码使用 `@arcships/vue-pdf/style.css`。

根入口还提供 `createPdfTextReferenceDraft`、`createPdfPageReferenceDraft`、`createPdfRegionReferenceDraft`、`resolvePdfReference`、`describePdfReference` 和 `normalizePdfReferenceRect`。解析器通过调用方传入的 Runtime 与当前 `PdfRenderDocument` 验证字符范围或重新搜索，不上传正文，也不管理确认后的产品状态。

PDF Runtime 只接收已经取得并验证的 `ArrayBuffer`，不会再次请求原始 PDF 地址。URL 来源必须显式提供 `urlPolicy`；加载会跟随重定向并在读取正文前用同一协议与来源规则重新校验最终地址，跨来源或协议降级重定向返回 `SOURCE_NOT_ALLOWED`。`PdfVerifiedDocument.resolvedUrl`、`PdfLoadError` 与诊断中的 URL 都去除查询串、片段和内嵌凭据。默认使用包内 PDFium Worker，关闭外部字体回退；Worker 或 WASM 初始化失败会返回错误，不会把验证字节状态当成渲染成功。宿主自带严格内容安全策略时，需要允许 `worker-src blob:`，并允许读取 `@arcships/vue-pdf/assets/pdfium.wasm` 对应的同源资源。

Chrome（issue #162）：`PdfViewer` 的默认行按 Extend `PDFViewerToolbar` 逐像素复刻（sidebar toggle、Page/页码/of N、rotate、zoom、search、`⋯`），额外命令（上一页/下一页/首页/末页、布局、文件名、下载、打印）收进 `⋯` actions 菜单；其余 chrome 使用 `DocumentViewerSidebar`、`OfficeEmptyState`/`OfficeLoadingState`/`OfficeErrorState`、`OfficeStatusBar`、`OfficePopover`/`OfficeTooltip`/`OfficeDialog`。没有新增或移除 PDF 根入口导出；只增加向后兼容的可选 props：`PdfViewer.showPrint`（默认 `true`）、`PdfViewer.showStatusBar`（默认 `false`）、`PdfSurface.showStatusBar`（默认 `false`）、`PdfSurface.showAnnotationDialog`（默认 `false`）。既有事件、`data-testid`、`expose` 方法和 `PdfViewerProps` 必填合同不变（`pdf-next-page` 等命令仍在 actions 菜单中保留原 test id）；详细映射与 Extend 像素对照见 `docs/components/pdf-chrome.md`。

PDF 唯一的资源拒绝配置是整份文件体积 `maxFileSize`：默认 `50 MiB`，宿主可公开调整，不叠加第二层隐藏硬上限。本地 Blob 在读取前检查声明大小；URL 有可信 `Content-Length` 时先检查响应头，否则在取得正文后、创建 PDF 引擎前检查实际字节。超过当前值返回 `PDF_TOO_LARGE`，包含 `actual` 和 `allowed`，并且不显示旧页面或创建新页面。页数、单页像素、总内存和并发页不是公开拒绝条件。

页面布局（issue #110）：`PdfViewer` 和 `PdfSurface` 共用 `pageLayout`（`continuous` / `single` / `spread`）属性；`PdfViewer` 另有 `defaultPageLayout`，并抛出 `update:pageLayout`（工具栏下拉框驱动未受控状态），默认仍为 `single`。两种模式切换都保留当前页；`single` 只保留当前页位图，`continuous`/`spread` 按可视行懒加载，`spread` 把对页作为独立位图渲染并保留各自页码。根入口导出 `resolvePdfPageLayoutMetrics`、`pdfPageLayoutRows`、`pdfRowIndexAtOffset`、`pdfPageIndexAtOffset`、`pdfSnapTopForRow`、`pdfPageAtPoint`、`pdfRowsInViewport`、`pdfSpreadPagePair`、`pdfSpreadStepPage`、`normalizePdfPageLayout`、`isPdfSpreadLayout`、`PDF_PAGE_LAYOUTS` 及类型 `PdfPageLayout`、`PdfPageLayoutMetrics`、`PdfPageLayoutRow`、`PdfPageLayoutOptions`、`PdfPageDisplaySize`。

导航辅助（issue #103）：`PdfRenderRuntime.getOutline?()` 返回 `PdfOutlineItem[]`（`title`、`children`、可选 0 基 `pageIndex` 与 URI `url`），`PdfRenderRuntime.getPageLabels?()` 按页返回显示标签。两者都是可选能力；`PdfSurface` 在文档打开后读取，运行时未实现时省略对应界面。`PdfSurface.showOutline` 默认 `true`，文档带书签时渲染 ARIA tree（`data-testid="pdf-surface-outline"`、`pdf-outline-item`），支持 `↑`/`↓`/`→`/`←`/`Home`/`End`/`Enter`/`Space`，激活页内条目滚动到目标页，URI 条目抛出 `outlineActivate`（清洗后的 `url`，失败时 `blocked: true`）。`showThumbnails` 默认 `true`；`PdfSurface` 与 `PdfViewer` 共用 `PdfThumbnailStrip` 和同一缩略图管线（`data-testid="pdf-surface-thumbnails"`、`pdf-surface-thumbnail`），缩略图文字使用 `getPageLabels()` 的页标签并随当前可见页高亮；可用 `showThumbnails="false"` 关闭。根入口新增类型 `PdfOutlineItem`。

文本选择可以跨页：`PdfSelectionState` 的 `text` 是完整选区文字（多页用 `\n` 连接），`range`/`rects` 仍描述第一页，`segments` 按文档顺序给出每页的 `PdfSelectionSegment`。多页选区只用于选择和复制；PDF 引用仍是单页文本范围。`PdfRenderRuntime.getPageAnnotations?` 是可选能力，返回 `PdfRenderAnnotation`（链接、表单控件和其他批注的规范化矩形、链接目标、字段名和批注文字）。`PdfSurface` 据此渲染不可见的链接/控件命中层，并在点击时发出 `linkActivate`（`url`、`targetPageIndex`、`blocked`）或 `annotationActivate`（`kind`、`fieldName`、`fieldValue`、`contents`）。链接只接受绝对 `http`、`https`、`mailto` 地址并去掉内嵌凭据；提供了 `urlPolicy` 时还会按该策略再次校验，未通过的链接以 `blocked: true` 上报而不会自动跳转。

AcroForm 写入通过 `PdfRenderRuntime.setFormFieldValue?()` 完成：`PdfSurface` 为文本、复选框、单选框、下拉框和列表框渲染原生控件，提交后调用 Runtime、刷新页面位图并发出 `formFieldChange`（失败为 `formFieldError`）。`resolvePdfFormFieldValue()` 把 `PdfFormFieldValue`（字符串或布尔值）转换成 PDFium 表单值：文本字段接受字符串，复选框接受布尔值，单选与选择字段接受选项标签或索引。只读字段仍显示控件但被禁用；无法写入的 Runtime 不渲染控件，避免出现看似可编辑但实际无效的输入框。

批注外观编辑通过 `PdfRenderRuntime.updateAnnotationAppearance?()` 完成：`PdfAnnotationAppearancePatch` 支持 `contents`、`color`、`opacity`、`strokeColor` 和 `strokeWidth`，Runtime 重写批注并重新生成外观流，`PdfSurface.updateAnnotationAppearance(pageIndex, annotationId, patch)` 随后刷新页面位图并发出 `annotationChange`（失败为 `annotationError`）。`PdfAnnotationAppearancePatch`、`PdfFormFieldType`、`PdfFormFieldValue`、`PdfFormFieldOption` 和 `PdfFormFieldDescriptor` 是根入口类型。

无障碍文本层通过 `PdfRenderRuntime.getPageTextLayer?()` 提供：`PdfPageTextLayer` 给出 canonical 页宽高和按 PDFium 阅读顺序排列的 `PdfPageTextLayerGlyph`（`charIndex`、`text`、glyph `rect`、`unmapped`）。`PdfSurface` 与 `PdfViewer` 为每个 glyph 渲染透明 `<span>`，位置来自 glyph box，并只用一个 CSS matrix 随缩放与 0/90/180/270 旋转变换；文本层不是 `aria-hidden`，屏幕阅读器、浏览器查找和 CJK 提取均可读取，`unmapped` glyph 标记 `aria-hidden` 且不可选。文本层 `pointer-events: none`，指针选择仍由自定义 interaction layer 处理，因此跨页选择、链接/表单命中与区域选择不变。`PdfPageTextLayer` 和 `PdfPageTextLayerGlyph` 是根入口类型。

### 3.6 `@arcships/vue-ui`

稳定入口为 `OfficeObjectOutlineLayer`、`OfficeRegionSelector`、`OfficeRegionAnnotator`、`captureRegion`、`releaseRegionCapture`、`computeRegionCaptureGeometry`、`computeRegionRasterCrop`、`createOfficeRegionAnnotationDocument`、`parseOfficeRegionAnnotationDocument`、`stringifyOfficeRegionAnnotationDocument`、`officeRegionAnnotationBounds`、`SignaturePad`、`FileUpload`、`FileThumbnail`、`BoundingBoxCitations`、`LayoutBlocks`、`Spinner`、`Tooltip`、`OfficeFindBar`、`OfficeLiveStatus`、`OfficeThemeProvider`、`OfficeIcon`、`HugeIcon`、`HugeIconData`、`useOfficeMenuKeyboard`、`resolveOfficeMenuKeyCommand`、`useViewerKeys`、`resolveViewerKeyCommand`、`isEditableViewerTarget`、`isInteractiveViewerTarget`、`VIEWER_KEYBOARD_SHORTCUTS`、`useFileDrop`、`isSupportedDropFile`、`fileDropExtension`、`FILE_DROP_OUTLINE`、`OfficeButton`、`OfficeIconButton`、`OfficeInput`、`OfficeTextarea`、`OfficeSelect`、`OfficeDropdownMenu`、`OfficeContextMenu`、`OfficeDialog`、`OfficeSheet`、`OfficePopover`、`OfficeTooltip`、`OfficeTabs`、`OfficeToggle`、`OfficeBadge`、`OfficeCard`、`OfficeSeparator`、`OfficeSkeleton`、`OfficeSpinner`、`OfficeScrollArea`、`OfficeToastProvider`、`OfficeToolbar`、`OfficeToolbarGroup`、`OfficeToolbarSeparator`、`OfficePageControl`、`OfficeZoomControl`、`OfficeSearchPopover`、`OfficeStatusBar`、`OfficeEmptyState`、`OfficeErrorState`、`OfficeLoadingState`、`DocumentViewerSidebar`、`DocumentViewerSidebarSkeleton`、`useElementWidth`、`useOfficeToast` 及对应的 `Office*Props`/`OfficeMenuEntry`/`OfficeSelectEntry`/`OfficeZoomFitMode`/`OfficeSearchStatus` 等类型、`resolveOfficeTheme`、`officeSystemPrefersDark`、`provideOfficeTheme`、`useOfficeTheme`、`OFFICE_ICONS`、`OFFICE_ICON_NAMES`、`normalizeOfficeIconName`、`resolveOfficeIcon`、`provideOfficeIcons`、`useOfficeIcons` 和主题/图标相关常量，以及根入口导出的属性和事件类型。`HugeIcon` 是 `@hugeicons/core-free-icons` 数据的 Vue 渲染器（16–18px，跟随 token 与 `currentColor`），供需要与 Extend 图标逐像素对齐的组件复用。`DocumentViewerSidebar` 新增可选 `flush`（默认 `false`，为 `true` 时 body 去掉 padding 并交给子内容滚动）；`DocumentViewerSidebarSkeleton` 新增可选 `variant`（`list` 默认；`thumbnail` 对齐上游 PDF 侧栏骨架）；均为向后兼容新增。新代码使用 `@arcships/vue-ui/style.css` 和 `@arcships/vue-ui/theme.css`；只使用浅色 token 契约时可以改用 `@arcships/vue-ui/tokens.css`。

FileUpload / FileThumbnail（0.8 起，#163）在保留原有 props/事件的前提下新增可选视觉参数。`FileUpload` 新增 `title`、`description`、`browseLabel`、`draggingLabel`、`acceptedFileTypes`、`showFileList`（默认 `false`）、`showBorderBeam`（默认 `false`）、`borderBeamTheme`；`files-accepted` 与 `files-rejected` 的载荷不变，拒绝项的 `code`/`file`/`message` 保持稳定。`FileThumbnail` 新增 `previewImageUrl`、`previewAspectRatio`、`previewClassName`、`isLoading`、`hasError`、`showTypeGlyph`、`showMeta`、`alt`；`size` 在原 `sm`/`md`/`lg` 之外接受数字（像素）或 CSS 长度，旧的 `file.url` 图片回退仍可用。视觉只引用 `--office-*`/`--office-ui-*` token；`showFileList=false` 与 `showTypeGlyph=false` 是兼容默认值。完整合同见 [components/file-components.md](../components/file-components.md)。


主题与 token 层（0.8 起）是所有 Office 组件唯一的视觉来源。`theme.css` 把 `@arcships/vue-ui/tokens.css` 内联进发布产物，并在 `[data-office-theme="light"]`、`[data-office-theme="dark"]`、`[data-office-theme="system"]` 下提供浅色、深色和 `prefers-color-scheme` 切换；所有声明都限定在 `data-office-theme` 子树内，不修改 `:root`、`html`、`body`，也不依赖全局 `.dark` 类。`OfficeThemeProvider` 写入 `data-office-theme`、`data-office-theme-resolved`、`data-office-density` 和可选的 `lang`，提供 `theme`（默认 `system`）、`locale`、`density`（`comfortable`/`compact`）、`messages`、`as` 五个可选 props，并通过 `useOfficeTheme()` 暴露只读上下文；`system` 在客户端监听 `matchMedia`，无 `window` 的 SSR 环境回退为浅色且不崩溃。字体 token `--office-font-sans` / `--office-font-mono` 默认 Geist / Geist Mono，但包内不打包字体（无 `@font-face`、无 `url()`），宿主自行加载并可通过 `[data-office-theme]` 覆盖。

`OfficeIcon` 以 `@lucide/vue` 为运行时依赖（构建时 external），props 为 `name`、`icon`、`size`、`strokeWidth`、`label`；默认尺寸来自 `--office-icon-size`（16px），默认线宽为 `2`，省略 `label` 时输出 `aria-hidden="true"`。内置注册表只命名导入常用图标，未使用的 Lucide 图标不进入消费者 bundle；`provideOfficeIcons()` 可以在子树内扩展或覆盖，`resolveOfficeIcon()` 支持 `chevron-down`、`ChevronDown` 等写法。token 数值与 Extend 参考 commit `70fe70c6f1e1eee53954c6923f3d9f583557a57e` 对齐：`--office-radius: 0.625rem`、8% alpha 边框、4% alpha muted/secondary/accent、中性 primary、2px 焦点环、`shadow-xs`/`shadow-sm`。

Primitive 层（0.8 起）以 `reka-ui@^2.11` 为运行时依赖，构建时 external，不引入 Tailwind/preflight。共享约定：`class`、`variant`（`default`/`secondary`/`ghost`/`outline`/`destructive`/`link`）、`size`（`xs`/`sm`/`default`/`lg`，按钮另有 `icon`/`icon-sm`/`icon-lg`）、`disabled`、`loading`、`data-slot` 和 `asChild`。视觉只引用 `--office-*` token；`.office-ui` 上只补充共享 token 未覆盖的间距、overlay、z-index 和动画时长。Dialog/Sheet 的焦点陷阱、Esc、backdrop、关闭后焦点归还，Select/Menu/Tooltip/Popover 的 portal 与碰撞定位，Tabs 的 roving tabindex，Toast 的 status live region 均由 reka-ui 提供。`Spinner` 和 `Tooltip` 的公开 props 保持兼容，后者内部改为委托 `OfficeTooltip`。完整合同见 `docs/components/ui-primitives.md`；包体积预算为新增 gzip ≤35 KB（不含图标），由 `tests/unit/vue-ui-primitives-contract.test.mjs` 检查。

共享 chrome 层（issue #158）在 primitive 之上提供 `OfficeToolbar`/`OfficeToolbarGroup`/`OfficeToolbarSeparator`、`OfficePageControl`、`OfficeZoomControl`、`OfficeSearchPopover`、`OfficeStatusBar`、`OfficeEmptyState`/`OfficeErrorState`/`OfficeLoadingState`、`DocumentViewerSidebar`/`DocumentViewerSidebarSkeleton` 和 `useElementWidth`。所有组件保持受控/非受控双模式，视觉只读 `--office-*` token；toolbar 最小高度 48px、`px-3 py-2`、4/8px gap、底边和 `bg-background`，sidebar 固定 160px 并在容器宽度低于 768px 时自动 overlay。新增 i18n 键位于 `common.chrome.*` 与 `common.zoom.fitWidth`/`fitPage`/`actualSize`/`percent`，`OfficeChromeMessages` 已导出。完整 API、键盘/ARIA 和响应式截图证据见 `docs/components/office-chrome.md`。四个 viewer 的迁移留给 #159–#162，现有公开 props 不变。


`OfficeLiveStatus` 是所有 Office 视图共享的单一 `role=status` / `aria-live` 区域：视图只改文本，不改节点，因此加载、错误、页码和幻灯片切换都能被辅助技术重新播报。`useOfficeMenuKeyboard` 统一 DOCX 与 XLSX 右键菜单的键盘语义：打开时把焦点放到第一个 `menuitem`，上下方向键、`Home`、`End` 在菜单项之间循环，`Escape` / `Tab` 关闭菜单并把焦点还给打开它的表面；`resolveOfficeMenuKeyCommand` 是可单测的纯函数版本。

`useViewerKeys` 统一 DOCX、PDF 和 PPTX 的键盘映射：`PageUp`/`PageDown`、`Home`/`End`、`Ctrl/Cmd` + `+`/`=`/`-`/`0`、`Ctrl/Cmd + F` 和 `Escape`；`mode: "present"` 时才额外接管方向键和 `Space`。组合函数在可编辑目标（`input`、`textarea`、`select`、`contenteditable`）内不拦截按键，只有宿主处理了命令才调用 `preventDefault()`。`resolveViewerKeyCommand` 是可单测的纯函数版本，`VIEWER_KEYBOARD_SHORTCUTS` 是导出的映射清单。

`useFileDrop` 是四个查看器共用的拖放导入合同：`accept`（`extensions` + `mimeTypes`，忽略大小写）决定可接受的文件，`onFile` 只收到第一个匹配文件，`onReject` 收到全部未匹配文件，`enabled` 支持布尔值或 getter。组合函数返回 `onDragEnter`/`onDragOver`/`onDragLeave`/`onDrop`/`reset` 和 `isDragOver`；深度只由 `dragenter`/`dragleave` 的嵌套层级维护，`dragover` 只设置 `dropEffect`，因此不会复现 #27 的轮廓卡滞。`FILE_DROP_OUTLINE` 是约定的虚线轮廓（`3px dashed #2563eb`），`isSupportedDropFile` 与 `fileDropExtension` 是可单独复用的纯函数。

各查看器在根元素绑定该组合函数，并在命中时抛出 `fileDrop(file)` 事件；拖入时根元素带 `data-file-drop-active="true"`。不支持的类型不会进入加载管线：DOCX 通过 `controller.setStatus()` 提示，XLSX/PDF/PPTX 在查看器上方渲染 `role="alert"` 条（文案来自 `xlsx.viewer.unsupportedDrop`、`pdf.viewer.unsupportedDrop`、`pptx.viewer.unsupportedDrop`）。`@arcships/vue-xlsx` 还导出 `XLSX_FILE_EXTENSIONS`、`XLSX_FILE_MIME_TYPES`、`XLSX_FILE_ACCEPT` 和 `isSupportedXlsxFileName`，`@arcships/vue-pdf`/`@arcships/vue-pptx` 提供同名的 `PDF_*`/`PPTX_*` 常量与判断函数。

两个 Office 选择原语依赖 `@arcships/office-interaction` 的公开类型。它们只消费规范化几何和受控临时状态：不解析 Office 文件，不从 DOM 猜测格式对象，也不保存引用集合。轮廓层负责候选可视化和键盘事件；区域框负责 `0..1` 坐标中的拖选与键盘调整。确认后的工具栏、引用列表和 Agent 工作流由宿主实现。

区域截图与标注是可选能力，headless 宿主不引入任何 DOM 依赖即可继续只使用引用事件。`captureRegion({ surface, region, scale })` 接受元素、宿主自建的 capture provider 或内置 Surface 组件实例（四个 Surface 都实现可选的 `getCaptureSurface()`）；`region` 可以是规范化矩形，也可以是 `ManualRegionLocator`，后者由有能力解析的 Surface 返回 `region` 覆盖值。默认 `scale` 为 `devicePixelRatio`，输出像素尺寸先经过 `maxPixels`、`maxDimension`、`maxBytes` 与 `MAX_REGION_CAPTURE_SCALE` 检查；返回对象持有 `blob`、`url`、`source` 元数据和 `release()`，宿主必须在替换或卸载时调用 `release()`/`releaseRegionCapture()` 以撤销 object URL。Canvas/图片 Surface 应提供 intrinsic `raster`，通过 `drawImage` 精确裁剪；纯 DOM Surface 使用 SVG `foreignObject` 尽力栅格化。

`OfficeRegionAnnotator` 是同一能力的可选 Vue 组件：它在截图上提供箭头、矩形、文字和高亮工具，坐标始终相对截图自身归一化到 `[0,1]`，样式中的 `strokeWidth`/`fontSize` 相对短边归一化，因此 1x 与 2x 位图渲染一致。`OfficeRegionAnnotationDocument` 可经 `stringifyOfficeRegionAnnotationDocument()`/`parseOfficeRegionAnnotationDocument()` 往返 JSON；文档不属于 `OfficeObjectReference`，也不会上传或写入引用。组件只在 `image` 是 `Blob` 时创建 object URL，并在 prop 变化和卸载时撤销。

完整字段、限制与示例见[区域截图与标注](region-capture-and-annotation.md)。

### 3.7 `@arcships/pptx-core`

根入口提供平台无关的播放类型、对象身份、能力报告、时间安排和属性轨道。浏览器中的文档会话、静态预览、播放解析和控制器从 `@arcships/pptx-core/browser` 导入。预览会话默认 `renderMode: "list"`，纵向连续展示页面；播放文档会话默认 `renderMode: "slide"`，只挂载当前页。调用方可通过 `PptxPreviewSessionOptions.renderMode` 和 `listOptions` 显式配置。补丁后的渲染器代码已经包含在浏览器入口中，消费项目不安装 `@aiden0z/pptx-renderer`。

根入口还提供 `createPptxSlideReferenceDraft`、`createPptxTextReferenceDraft`、`createPptxObjectReferenceDraft`、`createPptxRegionReferenceDraft`、`resolvePptxReference` 和 `describePptxReference`。对象适配直接复用现有 `objectKey`、shape id、source 与 group path；精确文字定位使用对象内段落和字符边界。它只生成/解析引用，不实现对象菜单、图层面板或确认后的业务动作。

根入口同时公开页面切换能力：`pptxTransitionCapability()`、`isPptxTransitionExecutable()`，以及 `pptxTransitionDirection()`、`pptxTransitionWipeInsets()`、`pptxTransitionWipeClip()`、`pptxTransitionSplitMask()`、`pptxTransitionProgress()` 和类型 `PptxTransitionCapability`、`PptxTransitionDirection`。`cut`、`fade`、`push`、`wipe`（含四角方向）、`cover`、`uncover`、`split` 为 `strict`；`dissolve`、`morph` 为 `approximate`（默认 `static`）；三维切换固定为 `static`（`transition:3d:<kind>`），不会退化为二维淡入。`isPptxTransitionExecutable(kind, approximation)` 是播放控制器决定是否挂载切换层的唯一入口。

`PptxSlideInfo.notes` 是演讲者备注文本（一篇一段，用换行分隔）；没有备注部件的页面不带该字段。备注部件不是渲染器的解析目标，浏览器入口因此另提供 `readPptxSpeakerNotes()`、`parsePptxNotesXml()`、`pptxNotesPartPath()`、`PPTX_NOTES_ARCHIVE_READER` 以及 `PptxNotesArchiveReader`、`PptxSpeakerNoteSlide`、`ReadPptxSpeakerNotesOptions`：备注在页 XML 解析完成后从同一字节只解压备注部件，损坏的备注只丢该页备注，预算超限和取消仍然使打开失败。

无法精确执行的 PPTX 内容通过能力报告标记为近似、静态或未解析；公开说明不得宣称完整兼容 PowerPoint 全部动画。

播放 v2 的 `PptxPlaybackEffect` 新增 `characterRange`/`characterGranularity`（按词/按字）、`tableCell`、`chartElement` 和 `graphicElement`；`PptxSlideTransition` 新增 `sound`。对应能力码固定为 `text:by-word`、`text:by-character`、`text:character-range-unavailable`、`table:cell`、`table-build:<mode>`、`chart:element`、`chart:series`/`chart:category`/`chart:point`、`smartart:node`/`smartart:level`、`group:sub-shape`、`transition:sound` 和 `transition:3d:<kind>`。`blockedMediaIds` 同时包含被浏览器阻止的媒体和切换声音，`resumeBlockedMedia()` 统一重试。

浏览器入口还提供 `loadPptxArchive`、`PPTX_ARCHIVE_READER` 和 `resolvePptxExternalMediaUrl`。`PptxPreviewLimits` 覆盖完整的归档预算（压缩比、路径长度、XML 总量与单部件、嵌套层数、标签与文本节点、关系数）。每次加载只解压一次：`inspectOfficeArchive` 只读中央目录，渲染器完成唯一一次解压，共享校验器直接检查已解码的部件。外部媒体默认禁用；`externalMedia: "allowed"` 必须同时配合 `externalMediaUrlPolicy`，目标 URL 经协议和来源校验后才加载。

### 3.8 `@arcships/vue-pptx`

稳定入口为 `PptxViewer`、`PptxThumbnail`、`PptxStage`、`PptxLinkMenu`、`usePptxDocument`、`usePptxPlayback`、`PptxStageSelection`、`PptxStageObjectClick`、`PptxStageContextMenu`、`PptxStageLinkActivate`、`PptxObjectHyperlink` 和其他公开类型。`PptxLinkMenu` 是链接右键菜单的低层组件（共享 `OfficeDropdownMenu`），供自建舞台复用；`PptxViewer` 内部已使用它。`PptxViewer` 默认纵向连续浏览；`mode="present"` 使用单页舞台并提供下一步、上一步、暂停、继续、重播、跳页、媒体恢复和全屏。`showNotes`（默认 `false`）挂载演讲者面板：浏览模式显示当前页备注，演示模式显示演讲者视图（当前页备注 + 下一页缩略图 + 计时），面板根节点为 `data-testid="pptx-presenter"` 并带 `data-presenter-mode`，备注可用 `notes` 插槽替换。浏览布局由 `viewMode`（`single` / `sorter`，未受控时由工具栏控制）选择，另有 `defaultViewMode` 和 `update:viewMode`（issue #110）；`mode`/`showNotes` 同样支持受控与非受控（`update:mode`、`update:showNotes`），工具栏的 `showPresentToggle`、`showNotesToggle`、`showFullscreenControl` 默认开启（#161）。空/加载/错误状态由共享 `Office*` state 渲染，OLE/CFB 密码保护或旧版二进制文件返回 `UNSUPPORTED_CONTENT` 并显示受保护状态（`data-encrypted="true"`，`data-state` 仍为 `error`）。`sorter` 是懒加载的缩略图网格，只挂载可视行附近的格子，点击格子切页但保持网格打开，两种布局共享同一个活动页。超链接（issue #114）在两种模式下都可点击：`a:hlinkClick` 经共享 URL 策略只允许 `http:`/`https:`/`mailto:` 和文档内跳页，外部链接在新标签页打开、内部链接走同一个 `goTo()`，`javascript:` 等目标被拒绝；`action` 事件仍按原样抛出，`linkHandling="host"` 关闭内置打开/复制行为。`PptxStage` 额外发出 `link-activate`（键盘激活链接形状），右键带链接的形状由查看器渲染共享 `OfficeDropdownMenu` 内置菜单（内容节点 `data-testid="pptx-context-menu"`，`showContextMenu` 可关闭，`context-menu` 事件始终先抛出），悬停链接只作为 tooltip/`static` 能力报告，不自动执行。链接元数据写入 `PptxObjectIdentity.hyperlinks`（`PptxObjectHyperlink`）和 `PptxSlideAction.tooltip`，并在渲染元素上以 `data-pptx-hyperlink*`、`role="link"`、`tabindex` 和 `title` 暴露。自定义静态 Surface 使用 `PptxStage` + `usePptxDocument(renderMode: "list")`，自定义播放器增加 `usePptxPlayback` 并使用 `renderMode: "slide"`。样式从 `@arcships/vue-pptx/style.css` 导入。

### 3.9 `@arcships/office-interaction`

根入口提供框架无关、零运行时依赖的 `OfficeObjectReference`、`OfficeReferenceConfirmEvent`、四种格式定位器和四维可靠度类型。运行时入口包括引用与确认事件的 `parse` / `safeParse`、规范化稳定序列化、候选排序、几何计算、短生命周期选择 reducer、候选导航状态和键盘解释器。

引用只保存文档修订、语义、格式定位器、内容指纹和规范化区域，不保存 DOM、Vue 实例、Blob URL 或截图位图。`OfficeInteractionValidationError` 的 `code` 固定为 `INVALID_OFFICE_INTERACTION`，详细原因从结构化 `issues` 读取。该包不解析 Office 文件、不渲染界面、不保存引用集合、不调用 Agent，也不修改文件；格式包负责生成定位器，宿主决定如何组织和使用确认后的引用。

根入口还提供对象注册层：`OfficeObjectProvider<TContext>`、`createOfficeObjectRegistry<TContext>()`、`createOfficeObjectDescriptor()`、`createOfficeObjectFingerprint()`、`isOfficeObjectVisible()`、`officeObjectAncestors()` / `officeObjectChildren()`、`rankOfficeObjectCandidates()` 和 `officeObjectCandidatePreview()`。provider 只返回当前修订的 descriptor（含 draft、fingerprint、四维可靠度，可选 `visual` / `facets`），不提交引用；阶段四行为没有 `visual`，因此不会进入画布轮廓。注册层是四个格式包内置对象族的公共扩展点。

根入口同时提供共享消息目录：`officeMessagesEn`、`officeMessagesZhCN`、`OfficeLocale`、`OfficeMessagesOverride`、`resolveOfficeLocale`、`resolveOfficeMessages`、`formatOfficeMessage`、`formatOfficeDate` 和 `formatOfficeNumber`。`DocxViewer`、`DocxEditor`、`DocxDocumentSurface`、`XlsxViewer`、`XlsxSheetSurface`、`PptxViewer`、`PdfViewer` 与 `PdfSurface` 均接受 `locale`（默认 `"en"`，内置 `"zh-CN"`）和 `messages` 深合并覆盖；`XlsxViewer`/`XlsxSheetSurface` 的 `locale` 同时作为数字和日期格式的默认区域（可用 `documentLocale` 单独覆盖），键入 `1.234,5` 在 `de-DE` 下解析为 1234.5 并显示为 `1.234,50`，`[$-407]` 等格式内置区域始终优先，显式 `@` 文本格式保持原样；用户可见文案、aria-label 和数字/日期显示不再与查看器语言脱节。

### 3.10 统一 Surface 引用选择

`DocxDocumentSurface`、`XlsxSheetSurface`、`PdfSurface` 和 `PptxStage` 统一接受 `documentId?: string`、`selectionMode?: "content" | "object" | "region"` 和 `emitReferenceCandidates?: boolean`。`selectionMode` 是受控输入，只解释下一次选择手势；Surface 不保存确认后的引用。

四个 Surface 共同 expose `getDocumentRevision()`、`hitTest()`、`describeReference()`、`resolveReference()`、`scrollToReference()` 和 `captureReferencePreview()`；同时提供可选的 `getCaptureSurface()` 供 `@arcships/vue-ui` 的 `captureRegion()` 使用。`captureReferencePreview()` 的位图采集仍由宿主渲染管线负责，内置实现会发出 `CAPTURE_UNSUPPORTED` 并拒绝 Promise，不会返回伪造截图；需要内建区域位图时使用新的 `captureRegion()` 合同。

统一事件为 `documentRevisionChange`、`referenceCandidateChange`、`referenceConfirm`、`regionDraftChange`、`selectionCancel`、`referenceResolve` 和 `referenceError`。既有 `selectionChange`、`objectClick`、`contextMenu` 等格式事件在 `0.x` 中继续保留。每个格式包从根入口重导出共享 Surface 类型，消费端无需深层导入。

阶段二–四的首批对象在 `0.8` 批次接入：DOCX/XLSX/PPTX/PDF 的对象索引与行为 adapter 见各格式包根入口（例如 `createDocxObjectRegistry`、`createXlsxObjectRegistry`、`createPptxObjectRegistry`、`createPdfObjectRegistry`）。`selectionMode="object"` 时的悬停轮廓只依赖候选 `preview.label` 和 DOM/页面几何，宿主无需读取私有对象；PDF 的 `PdfVisualObjectProvider` 由宿主注入，库不内置 OCR。表格单元格、图表元素等内部对象需要宿主提供 `getObjectTable` / `getObjectChart` 证据时才能创建引用，缺失证据返回结构化 `unsupported`，不会猜测。

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

XLSX 绘图剪贴板合同：图表、图片和形状共用同一套内部剪贴板载荷。选中绘图后 `Ctrl/Cmd+C`/`X`/`V`/`D` 以及绘图右键菜单的复制、剪切、粘贴、创建副本、删除分别调用 `copySelectionToClipboard()`/`cutSelection()`/`pasteFromClipboard()`/`duplicateDrawing()`/`deleteSelectedDrawing()`。载荷的 `drawings` 数组为每个绘图携带 `kind`、`model`、`anchor`、`relativeAnchor`、`size`、`sourceSheetId` 和 `sourceId`；粘贴在同工作簿的目标工作表上克隆 DrawingML 锚点，生成新的模型 id、关系 id、图表部件（`xl/charts/chartClipboardN.xml`）和必要的绘图部件，因此支持跨工作表粘贴、一个撤销条目内完成的剪切移动以及导出/重新加载。`copyDrawing()`、`cutDrawing()`、`pasteDrawing()` 和 `duplicateDrawing()` 是同一流程的可编程入口。

XLSX 自动筛选合同：`controller.autoFilterRanges` 汇总工作表 `<autoFilter ref>` 和表部件的 `<autoFilter>`（`{ reference, start, end, headerRowCount, dataEndRow, columns, source, tableName? }`，同一区域以表定义优先）；`autoFilterStates` 给出当前列条件及 `hiddenRows`，`hiddenAutoFilterRows` 是所有条件隐藏的绝对行号。`applyAutoFilterValues(reference, columnIndex, values, { includeBlank })` 应用或移除单列筛选，`clearAutoFilter(reference?)` 清除一个/全部范围，二者只改视图、不改单元格、合并区域或冻结窗格；`getAutoFilterColumnValues(range, columnIndex)` 返回去重后的显示值（含空白）。`sortAutoFilter(reference, columnIndex, direction)` 是一次可撤销的范围编辑，只读/受保护工作表或数据区有合并时返回 `false`。筛选条件在导出时写回工作表或表部件的 `autoFilter`，重开工作簿后仍生效。`XlsxTableHeaderMenuRenderProps` 增加 `autoFilter`、`applyFilter`、`clearFilter`、`filterValues`、`getValueOptions`、`hasFilter`、`hasValueFilter`，`table` 在纯工作表筛选时为 `null`；`XlsxViewer`/`XlsxSheetSurface`/`XlsxGrid` 通过 `#table-header-menu` 插槽把参数交给宿主替换内置 `▾` 触发器。

XLSX 网格键盘与多格填充合同：`XlsxGrid` 聚焦时实现 Excel 的 `PageUp`/`PageDown`（正文视口高度的一屏并保留一行重叠）、`Home`/`End`（行首/行内最后一个非空单元格）、`Ctrl/Cmd+Home`/`Ctrl/Cmd+End`（A1/已用区域右下角，空表为 A1）、`Ctrl/Cmd+A`（第一次已用区域，第二次整张显示网格；空表一次全选）、`Ctrl/Cmd+Arrow`（数据区边缘，RTL 下与普通方向键同向镜像）和 `Escape`（收拢区域选区）；带 `Shift` 时扩展选区。翻页步长、跳转落点、已用区域收敛和选区比较集中在 `packages/vue-xlsx/src/components/grid-keyboard.ts`（`resolvePageRowStep()`、`resolveHorizontalJumpDirection()`、`resolveJumpTarget()`、`resolveLastUsedDisplayIndex()`、`isSameRange()`），组件只做事件接线。多格选区上按 `Enter` 会打开编辑器，在编辑器里再按 `Enter` 时调用 `XlsxViewerController.commitSelectionInput(text: string) => number`：`Tab`、失焦和上下方向键仍只写活动单元格；填充本身与 `commitCellInput()` 使用相同的 `=` 公式/值判定，公式原样复制不位移，跳过合并延续单元格和验证失败的单元格，整个填充只写一条撤销记录并保持选区，返回写入的单元格数（工作表锁定、条目被拒或无写入时为 `0`）。

XLSX 表格样式与透视/切片合同：`BUILTIN_TABLE_STYLES`、`resolveBuiltinTableStyle()`、`resolveTableCellStyle()`、`withTableCellStyle()` 和 `findTableAtCell()` 从根入口导出；后两者把表样式（整表 → 条纹 → 首末列 → 表头 → 总计 → 角单元格）合并到解析后的单元格 `XlsxResolvedCellStyle` 上；`headerRowCount="0"` 的首行按数据行处理并取首行条纹，`declaredTableStyleOverrides` 只允许 `xf` 声明的 font/fill/border 覆盖表样式，因此 `XlsxResolvedCellStyle` 新增可选 `tableStyleOverrides`，`XlsxSheetData` 新增 `cellStyleIds`。`XlsxPivotTable`（`cacheId`、`range`、`source`、`rowFields`/`columnFields`/`pageFields`、`dataFields`、`fields`、`layout`、`styleInfo`、`refreshOnLoad`）和 `XlsxSlicer`（`name`、`caption`、`cacheName`、`sourceName`、`pivotTables`、`style`、`columnCount`、`items`、`anchor`）通过 `parseWorksheetPivotTables()` / `parseWorksheetSlicers()` 解析，并出现在 `XlsxSheetData.pivotTables`/`slicers`/`hasPivotTable`、Worker 快照和 `XlsxViewerController.pivotTables`/`slicers`/`hasPivotTable` 上。`XlsxDefinedName`（`name`、`scope`、`sheetName?`、`range`、`comment?`、`hidden?`）由 `parseWorkbookDefinedNames()` / `definedNamesForWorksheet()` 解析，出现在 `XlsxSheetData.definedNames` 和 `controller.definedNames`（过滤 `_xlnm.*`）；Worker 只读加载没有主线程 workbook 时，`controller.definedNames` 回退到活动工作表快照的 `definedNames`。公式栏名称框据此列出并跳转；跨工作表名称先切换标签并在 `nextTick()` 后选区，名称列表依赖 `revision` 实时更新；聚焦时以当前 A1 地址作为草稿不会过滤名称（地址/区域草稿显示全部名称），非地址文本才按名称筛选。导出时 `mergeWorkbookSheetAuthoringAssets()` 保留 `xl/pivotTables`、`xl/pivotCache`、`xl/slicers`、`xl/slicerCaches` 部件及工作表/工作簿关系，并把源 `xl/styles.xml` 的 `<dxfs>`/自定义 `<tableStyles>`（含 `dxfId` 索引）写回引擎保存的样式表，重载后自定义表格样式与条件格式 dxf 一致。

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
| `OfficeRegionAnnotator` | `update:modelValue`、`update:mode`、`update:color`、`annotation-add`、`annotation-remove`、`annotation-select`、`annotation-change`、`annotation-error`、`image-load` | 标注文档、单条标注、模式/颜色或错误；坐标归一化到截图 |
| `OfficeSelect` | `update:modelValue`、`update:open` | 字符串选项值或打开状态 |
| `OfficeDropdownMenu` / `OfficeContextMenu` | `select`、`update:open` | `select` 参数为被选中的 `OfficeMenuEntry` 和原始事件；`update:open` 为布尔值 |
| `OfficeDialog` / `OfficeSheet` / `OfficePopover` | `update:open` | 布尔值 |
| `OfficeInput` / `OfficeTextarea` / `OfficeToggle` / `OfficeTabs` | `update:modelValue` | 字符串、布尔值或 tab value |
| `OfficeToastProvider` / `useOfficeToast` | 通过组合函数触发 | `success`/`error`/`warning`/`info`/`loading`/`add`/`update`/`close`/`promise` |


`FileThumbnail`、`Spinner` 和 `Tooltip` 当前没有业务事件。`OfficeThemeProvider` 只写主题属性并暴露上下文，`OfficeIcon` 只渲染图标，两者不发出业务事件。文件拒绝使用 `files-rejected`；宿主不能解析界面文字推断原因。选择控件只表达候选、确认或取消，宿主可以使用 `@arcships/office-interaction` 的临时 reducer，也可以自行管理状态。

### 4.5 `vue-pptx`

`PptxViewer` 提供 `load-start`、`load-success`、`load-error`、`slide-change`、`playback-ready`、`playback-state-change`、`step-change`、`playback-warning`、`capability`、`media-request`、`action`、`playback-error`、`context-menu`、`update:viewMode`、`update:mode` 和 `update:showNotes`。程序判断使用公开错误码和能力报告，不解析界面文案。

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
