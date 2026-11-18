# DOCX 使用指南

`DocxViewer` 支持 `.docx`、`.docm`、`.dotx` 和 `.dotm`。宏文档和宏模板只用于安全预览，VBA 不会执行；模板按普通文档打开。`DocxEditor` 的原格式编辑和宏保留不在本次支持范围内。

## 查看文档

```vue
<script setup lang="ts">
import { DocxViewer } from "@arcships/vue-docx"
import "@arcships/vue-docx/style.css"
</script>

<template>
  <DocxViewer
    :file="bytes"
    file-name="contract.docx"
    :show-toolbar="true"
    @load-success="onLoaded"
    @load-error="onError"
  />
</template>
```

`DocxViewer` 可以接收 `file: ArrayBuffer`，也可以直接接收已经解析的 `model: DocModel`。两者同时提供时优先使用 `model`。

常用属性：

| 属性 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `file` | `ArrayBuffer` | 无 | DOCX、DOCM、DOTX 或 DOTM 字节 |
| `model` | `DocModel` | 无 | 已解析文档模型 |
| `runtime` | `DocxRuntime` | 组件创建 | 调用方持有的 Runtime |
| `showToolbar` | `boolean` | `true` | 显示查看工具栏 |
| `showUpload` | `boolean` | `true` | 上传入口，含工具栏按钮与空态、错误态的选择文件按钮 |
| `showDownload` | `boolean` | `true` | 工具栏下载按钮 |
| `defaultThumbnailsOpen` | `boolean` | `false` | 默认展开缩略图 |
| `defaultZoom` | `number` | `100` | 初始缩放百分比 |
| `isDark` | `boolean` | `false` | 深色界面 |
| `showTrackedChanges` | `boolean \| null` | `null` | 受控显示修订 |
| `revisionDisplay` | `"inline" \| "balloons" \| "none" \| null` | `null` | 受控选择修订显示方式；`null` 跟随内部状态 |
| `defaultRevisionDisplay` | `"inline" \| "balloons" \| "none"` | `"balloons"` | 未受控时的初始修订显示方式 |
| `showComments` | `boolean \| null` | `null` | 受控显示批注 |

主要事件为 `load-start`、`load-success`、`load-error`、`file-selected`、`update:isDark`、`update:showTrackedChanges`、`update:revisionDisplay` 和 `update:showComments`。

## 修订显示方式

`revisionDisplay` 只改变页面如何绘制修订，不修改模型：导出、接受/拒绝和重新打开仍保留全部修订。

| 模式 | 正文 | 右侧 gutter |
|---|---|---|
| `inline` | 插入和删除都留在正文（绿色下划线 / 红色删除线） | 隐藏修订卡片；批注卡片仍显示 |
| `balloons` | 删除内容离开正文，插入内容仍留在正文 | 显示删除、移动和格式修订卡片 |
| `none` | 只显示最终文本，没有内联修订 | 隐藏修订卡片；批注卡片仍显示 |

默认值是 `balloons`：原来的“正文内联删除 + 同一删除的卡片”重复显示会被消除。`reviewDisplayModel(model, display)` 生成只读显示模型，`reviewDisplayParagraph()`、`hiddenParagraphChildIndexes()`、`hiddenParagraphTextRanges()`、`displayTextOffsetForSourceOffset()` 和 `sourceTextOffsetForDisplayOffset()` 可供自定义表面复用；`renderParagraphRuns()` 的 `revisionDisplay` 默认 `inline`，保证直接调用渲染器时行为不变。编辑器和查看器工具栏都提供“审阅显示方式”选择器；编辑器也可以通过 `controller.setRevisionDisplay()` 切换。

DOM 选区使用显示文本偏移，控制器区间使用源模型偏移：写入选区前会把 DOM 的显示偏移映射回源偏移，恢复 DOM 选区和右键菜单定位时再把控制器/源区间映射回显示偏移。正文、表格单元格、页眉页脚和脚注/尾注段落都可以复用 `paragraphDisplayOffsetToSourceOffset()` / `paragraphSourceOffsetToDisplayOffset()`；因此删除内容隐藏后，光标仍落在 Word 最终文本对应的源位置，撤销/重做也能恢复到正确的可见位置。

## SVG 图片

DOCX 中的 `image/svg+xml` 图片按矢量渲染，支持直接引用 SVG 的 `a:blip` 和 Word 常见的 `asvg:svgBlip` + PNG 回退结构（此时取 SVG）。图片保持 `wp:extent` 的锚定尺寸，缩放或 200% 预览时仍按显示分辨率重新栅格化，不会被替换成位图。

渲染前会对 SVG 标记做安全清洗：移除 `<script>`、`on*` 事件属性、`javascript:` 与所有外部引用；保留 `#id` 内部引用和 `data:image/` 内嵌图片，因此渐变、裁剪路径和 `<use>` 仍正常绘制。使用 `<foreignObject>` 或滤镜（`<filter>` / `filter=`）的图片无法用 `<img>` 管线绘制，会回退到占位符并在图片框内显示 `SVG`（小图标显示 `s`）。标记无法解析或超过 `MAX_SVG_PICTURE_MARKUP_CHARS`（4 MiB）时同样回退。

```ts
import {
  inspectSvgPicture,
  resolveRenderableImageSource,
  svgPictureUnsupportedDiagnostic,
} from "@arcships/docx-core"

const inspection = inspectSvgPicture(imageRun)
if (inspection && !inspection.supported) {
  // { code: "IMAGE_UNSUPPORTED_FEATURE", phase: "image-render", reason, partName }
  console.warn(svgPictureUnsupportedDiagnostic(imageRun))
}
```

`renderParagraphRuns` 的 `onImageUnsupported` 选项会为每个回退的 SVG 回调一次，诊断带上 `imageKey` 和本地化消息；宿主可用 `imageUnsupportedText` 覆盖消息模板（支持 `{feature}` 占位符）。原始字节保留在模型和基础包中：编辑后导出并重新打开，SVG 媒体部件逐字节不变；缩略图不会降采样 SVG，而是让最终整页栅格化负责缩放。

## 图片图层顺序

浮动图片（`wp:anchor`）的叠放顺序来自 `wp:anchor/@relativeHeight`，模型里对应 `ImageRunNode.floating.zIndex`。编辑器右键图片菜单的“图层”分组提供四个命令：`bringToFront()`、`bringForward()`、`sendBackward()`、`sendToBack()`（均接收 `DocxImageLocation`）。前两个把 `zIndex` 提升到 `WORD_IMAGE_Z_INDEX_MAX`（251658240）或增加一个 `WORD_IMAGE_Z_INDEX_STEP`（65536），后两个降低到 `WORD_IMAGE_Z_INDEX_MIN`（0）或减少一个步长；超出范围时按 Word 的边界截断。内联图片没有叠放顺序，四个命令禁用。

每个命令只写入一条撤销记录，同时清除图片/段落/表格的 `sourceXml`，因此重新导出会写出新的 `relativeHeight`，与屏幕上 `DocxImageLayer` 使用的 CSS `z-index` 保持一致。宿主也可以通过 `resolveDocxImageLayerOrderState()` 和 `resolveDocxImageLayerOrderMenuActions()` 复用同一套禁用规则。

## 分节页码与奇偶页空白页

页码按“物理页序号 + 分节的 `w:pgNumType`”解析：`w:pgNumType/@w:start` 重新开始计数，`@w:fmt` 决定格式（`decimal`、`decimalZero`、`lowerRoman`、`upperRoman`、`lowerLetter`、`upperLetter`、`none` 等 Word 写入的 `ST_NumberFormat` 值），没有 `@w:start` 的分节接续上一分节的编号。同一份数据同时驱动页脚里的 `PAGE`/`NUMPAGES`/`SECTIONPAGES` 和缩略图的页码标签，所以页脚页码和缩略图不会互相矛盾。

```ts
import {
  createSectionPageNumberTracker,
  resolveDocumentSectionsFromMetadata,
} from "@arcships/docx-core"

const tracker = createSectionPageNumberTracker({
  sections: resolveDocumentSectionsFromMetadata(model.metadata),
  // 每页所属分节：有正文的页面给出首个节点下标，空白页省略（沿用上一页的分节）
  pages: pageSegments.map((segments) => ({ startNodeIndex: segments[0]?.nodeIndex })),
  nodeCount: model.nodes.length,
})

tracker.pageLabelForPage(2) // "i"：按本分节的格式渲染
const fields = tracker.fieldsForPage(2)
// { pageNumber: 1, pageNumberFormat: "lowerRoman", sectionPages: 3, totalPages: 9, pageLabel: "i" }
```

`w:type` 为 `oddPage`/`evenPage` 的分节必须落在奇数页/偶数页上，Word 会在需要时插入一张空白页。空白页属于**结束于该处的分节**：它沿用该分节的页眉页脚、继续该分节的编号，也计入该分节的 `SECTIONPAGES`；新分节的第一张正文页才用 `w:start` 重新编号并满足奇偶要求。`tracker.isParityBlankPage(pageIndex)` 判断某页是否是这种空白页，`tracker.parityBlankPages` 列出全部空白页（`createSectionParityBlankPage()` / `isSectionParityBlankPage()` 是同一结构的构造器和类型守卫）。

视图层的三处可见信息都来自同一个 tracker：

| 位置 | 可观测值 |
|---|---|
| `DocxPageSurface` | `data-docx-page-label`（如 `i`）、`data-docx-page-parity-blank` |
| `DocxThumbnailPanel` | `data-thumbnail-page-label`、`data-thumbnail-parity-blank` |
| 页眉/页脚字段 | `PAGE`、`NUMPAGES`（文档物理页数）、`SECTIONPAGES`（本分节页数） |

例如“正文 `1…n`，附录重新从罗马数字 `i` 开始且要求奇数页起”的文档，标签依次是 `1, 2, i, ii, …`；若正文最后停在奇数页，中间会插入一张沿用正文分节页脚、编号接续的空白页。

## 键盘导航

`DocxViewer` 根节点接收键盘事件：`PageDown`/`PageUp` 翻一页，`Home`/`End` 跳到第一页/最后一页，`Ctrl/Cmd` + `+`/`=`、`-`、`0` 按缩放档位放大、缩小或重置到 100%，`Ctrl/Cmd + F` 聚焦工具栏搜索框，`Escape` 关闭搜索。焦点在页码或搜索输入框内时，查看器不会抢这些按键；快捷键也不会从同级查看器串入。根节点带 `:focus-visible` 焦点环。完整映射见[查看器键盘导航](viewer-keyboard.md)。

## 编辑文档

```vue
<DocxEditor
  :file="bytes"
  :editable="true"
  :show-toolbar="true"
  :show-thumbnails="true"
  @load-error="onError"
/>
```

`DocxEditor` 可以自己创建编辑控制器，也可以通过 `editor` 传入 `useDocxEditor()` 返回的控制器。传入控制器适合把保存、撤销重做和业务按钮放到组件外部。

编辑器和查看器一样接受 `isDark` 和 `minHeight`。传入 `isDark` 时外壳、工具栏、缩略图和纸张一起切换为深色；不传时跟随 `controller.documentTheme`。`minHeight` 可选，数字按 px；不传时组件不设置最小高度。打印时工具栏、缩略图和状态栏自动隐藏，为每页插入分页，并按文档页面尺寸设置 `@page`。

## 导出前的模型准备

导出（`exportDocx()`、`getDocxBytes()`、`serializeDocx()`、`serializeDocModel()`、`modelToDocumentXml()`）会先调用 `prepareModelForSerialization(model)`：递归正文、页眉、页脚、分节页眉/页脚、脚注和尾注，再把折叠表格网格投影到物理列。

合并过的表格可能带一个逻辑网格：被合并掉的那一列在 `TableStyle.columnWidthsTwips` 中保留宽度 `0`，列号记录在 `TableStyle.collapsedGridColumns`，单元格的 `gridSpan` 仍按逻辑列计数。导入时会从 `w:tblGrid` 恢复这个逻辑网格、折叠列标记和 `styleId` / `sourceStyleId`。编辑清除表格的 `sourceXml` 后，导出会按物理网格重写 `w:tblGrid` 与 `w:gridSpan`（`projectCollapsedTableGrid` / `hasCollapsedTableGrid` 可直接复用同一套投影）；未编辑的表格继续写原始标记。垂直合并同样保留：编辑后的表格会重新写回 `w:vMerge w:val="restart"` 与 `w:vMerge`。插入/删除表格列会同步平移或删除 `collapsedGridColumns`，手动拖动列宽则会把网格实体化并清除该标记。

模型不合法时不会生成文件：样式在导入后被改动（`sourceStyleId` 与 `styleId` 不一致）、折叠网格带可见边框，或标记、列宽与单元格跨度自相矛盾，都会抛出 `DocxSerializationError`（`code` 为 `COLLAPSED_TABLE_GRID_STYLE_CHANGED`、`COLLAPSED_TABLE_GRID_BORDER_SPACING` 或 `INVALID_COLLAPSED_TABLE_GRID`，后者带 `reason`）。`exportDocxBuffer` 会把该错误映射为 `DocxExportError`，`code` 为 `INVALID_MODEL`。

## 查找与替换

编辑器工具栏右侧的放大镜（或 `Ctrl/Cmd + F`）打开查找栏。查找语义与查看器一致：默认忽略大小写、按字面量匹配，可分别打开区分大小写、全字匹配和正则表达式。输入即搜索，`Enter`/`Shift+Enter` 或 `‹`/`›` 切换结果，`Esc` 关闭。

查找栏第二行是替换框与“替换”“全部替换”按钮：

- “替换”只改写当前结果。匹配跨多个文本 run 时按一段文字处理，未命中的文字与图片、表单域保持原位；替换文字继承被替换文字的字号、粗斜体、颜色和超链接等 run 样式。
- “全部替换”对正文（含普通表格与嵌套表格）中当前查询的每一处命中各替换一次，整个操作只写一条撤销记录。撤销一次即可还原全部替换，重做重新应用。
- 页眉、页脚、脚注和尾注中的结果仍会出现在查找列表里，但不参与替换：页眉页脚通过页面上的可编辑带（见页边距/纸张设置一节）修改，脚注尾注只读。正文之外的结果会被跳过并保持原文不变。

同样的命令可直接在控制器上调用，适合宿主自己的查找界面：

```ts
import { docxReplaceTargetFromMatch, findDocxSearchMatches } from "@arcships/vue-docx"

const matches = findDocxSearchMatches(editor.model, "needle")
const target = docxReplaceTargetFromMatch(matches[0]) // 非正文结果返回 undefined
if (target) editor.replaceMatch(target, "pin")        // 一条撤销记录

editor.replaceAll("needle", "pin", { wholeWord: true }) // { replaced, skipped }
```

`replaceMatch(target, text)` 返回是否真的改动了模型；`replaceAll(query, text, options?)` 返回 `{ replaced, skipped }`，`skipped` 是无法替换的页眉/页脚/注释结果数量。替换后段落与所属表格的 `sourceXml` 失效，导出会按模型重新写入文字；批注范围的字符偏移（`commentRanges`）也会随替换一起平移。

## 内部书签导航

Word 的目录、交叉引用和 `w:hyperlink/@w:anchor` 会导入为 `#书签` 内部链接。查看器从 `w:bookmarkStart` 建立“书签名 → 正文节点”索引：点击目录项会滚动到书签所在页和段落，短暂高亮目标段落，并在可编辑表面上把光标放到目标段落。页眉、页脚、脚注、尾注和表格单元格中的内部链接走同一条路径。

书签查找不区分大小写，接受 `#Name`、`Name` 和 URL 转义（`#My%20Bookmark`）。`_GoBack` 等 Word 自动书签不是导航目标。

宿主也可以直接调用导航方法：

```ts
const surface = surfaceRef.value
const result = await surface.revealBookmark("#Section_1")
// { status: "revealed", name: "Section_1", nodeIndex: 12, pageIndex: 3, text: "..." }

const missing = await surface.revealBookmark("No_Such_Bookmark")
// { status: "missing", name: "No_Such_Bookmark" }
```

`DocxViewer`、`DocxEditor` 和 `DocxDocumentSurface` 都暴露 `revealBookmark(name)`。查看器需要挂载另一页时发出 `requestPageReveal(pageIndex)`；找不到书签时不抛错，而是在文档上显示一条非阻塞提示，并发出 `bookmarkMissing(name)`。

需要自己实现导航时，`@arcships/docx-core` 的 `createDocxBookmarkIndex(model)` 返回 `get(name)` / `has(name)` / `targets`，`normalizeDocxBookmarkName()` 提供与查看器一致的大小写、`#` 和 URL 转义处理。

## 字体度量与分脚本字体

查看器在首次分页前调用 `awaitDocumentFonts(model, { extraFamilies, embeddedFonts })`：它收集正文、页眉页脚、脚注/尾注、批注和编号标签实际使用的字体族，安装 `word/fontTable.xml` 里的内嵌 `FontFace`，并等待 `document.fonts.ready`。字体加载完成后 `invalidateFontMetrics()` 递增度量版本，`DocxViewerRoot` 会据此重新分页，因此网页字体晚到不会让页数变化。

`measureTextWidthPx()` 与 `resolveMeasureFontForText()` 按 OOXML `w:rFonts` 的 ascii / hAnsi / eastAsia / cs 分槽度量；`resolveDocxFontKerning()` 解析 `w:kern` 阈值，`resolveDocxFontFeatureSettings()` 把 `w14:ligatures` / `w14:cntxtAlts` 映射为 `font-feature-settings`。渲染层用同一组规则把 run 拆成多个 span，并输出 `font-kerning` / `font-feature-settings`，所以估算和 DOM 的换行位置一致。

没有 Canvas 时（SSR、Worker）度量优先使用隐藏 DOM span，再退到确定性字宽；`fontMetrics`、`getFontMetricsRevision()`、`registerFontMetricCache()` 与 `subscribeFontMetrics()` 供自定义渲染层复用同一套缓存和失效事件。

## 编辑辅助操作

`DocxEditor` 默认启用下列编辑辅助操作。不需要时用属性关闭；关闭后不会渲染对应界面，也不改变文档或控制器行为。

| 能力 | 交互 | 关闭方式 | 对应控制器 API |
|---|---|---|---|
| 表单域属性对话框 | 双击表单域打开，编辑值、类型、名称、占位提示、必填、最大长度和下拉选项；保存为一次撤销 | `:show-form-field-dialog="false"` | `updateFormField()` |
| 链接悬浮卡片 | 鼠标悬停链接显示打开/编辑/删除；编辑和删除会先选中链接文字 | `:show-link-card="false"` | `setLink()`；打开时发出 `link-open` 事件 |
| 表格行高拖拽 | 悬停或选中表格后拖动行边界；提交 `w:trHeight`（`atLeast`） | `:show-row-resize-handles="false"`（`DocxTableHost`） | `setTableRowHeight()` |
| 表格移动手柄 | 表格左上角的手柄：点击选中整张表格，拖到新位置后转为浮动表格（拖动时显示虚线幽灵框） | `:show-table-move-handle="false"`（`DocxTableHost`） | `setTableFloating()` |
| 段落样式悬浮预览 | 悬停工具栏样式选择器时按各样式自身的字体/字号/缩进渲染示例，点击即应用 | `:show-style-preview="false"` | `setParagraphStyle()` / `setHeading()` |
| 段落/表格拖拽排序 | 悬停正文段落或表格显示拖拽手柄，拖动到目标位置后插入；手柄支持 `↑`/`↓`，`Esc` 取消 | `:show-node-reorder="false"` | `moveBodyNode()` |

拖拽排序一次移动只产生一条撤销记录：落在节点自身位置（前或后）视为无操作，不写入历史；跨分节拖动时，分节起始位置跟随对应节点，而不是由被移动的节点重新定义分节。

表格移动手柄只出现在正文表格的第一页片段上：嵌套表格位于单元格内，页眉页脚表格没有可寻址的正文节点，两者都不会浮动。拖动超过 3px 才算拖拽，否则只是一次“选中整张表格”的点击（选区是一个覆盖全部单元格的 `cellRange`）。落下时调用 `setTableFloating(tableIndex, update)`，`update` 由 `resolveTableFloatingDropPatch()` 从指针位移（px → twips）算出：`xTwips` / `yTwips` 是落下位置，`horizontalAnchor` 为 `margin`、`verticalAnchor` 为 `text`（Word 把表格拖出文字流时写的锚点对）。再拖一次只改变位置，`leftFromText` / `rightFromText` 等环绕距离沿用上一次的值；`setTableFloating(tableIndex, null)` 把表格放回文字流。每次调用只写一条撤销记录，只读表格不渲染手柄，页码、筛选、嵌套表格等宿主可以传 `:show-table-move-handle="false"` 关闭。

浮动位置存在 `TableStyle.floating`，导出时由序列化补丁写成 `w:tblPr/w:tblpPr`（`w:tblpX` / `w:tblpY`、`w:vertAnchor` / `w:horzAnchor` 与环绕距离），并保证 `tblpPr` 是 `tblPr` 的第一个子元素。引擎会按模型重建被编辑过的表格，所以单元格编辑不再让浮动表格退回内联；重新打开导出的文件时，`w:tblpPr` 又解析回同一个 `floating`。渲染继续复用 `tableWrapperStyle()`：表格按环绕侧浮动，`tblpX` / `tblpY` 折算成相对锚点的外边距，文字绕在另一侧。

表单域的值只有在符合字段约束时才会写入模型：下拉值必须来自选项列表，日期必须是有效日期，文本遵守最大长度；被拒绝的保存不会修改文档，撤销栈也不增加。

链接卡片对外部地址调用 `window.open(..., "noopener,noreferrer")`；文档内部书签（`#name`）不会被当作 URL 打开，而是由查看器滚动并高亮目标段落，同时仍发出 `link-open` 事件供宿主记录或自行处理。

## 表格区域选择

可在表格内拖拽或按住 `Shift` 单击建立矩形选区；选区跨合并单元格时按物理网格高亮整个单元格（`gridRange` 记录 `startColumnIndex` / `endColumnIndex`）。工具栏和右键菜单提供“选择整个表格”和“清除内容”；字符/段落格式、复制和清除都作用于矩形内的全部单元格。按 `Delete`/`Backspace` 清空选区，按 `Esc` 折叠到锚点单元格。

```ts
editor.selectTableCellRange(0, {
  tableIndex: 0,
  anchorRowIndex: 0,
  anchorCellIndex: 0,
  focusRowIndex: 1,
  focusCellIndex: 2,
})
editor.selectedTableCellRange // { …, gridRange: { startColumnIndex: 0, endColumnIndex: 2 } }
editor.selectWholeTable(0)     // 整个表格一条矩形选区
editor.clearTableSelection()   // 清空矩形内所有单元格，一条撤销记录
editor.collapseTableCellSelection() // Esc 的控制器等价命令
```

`clearTableSelection()` 复用 `clearTableCellContents(tableIndex, cells[])`，整个矩形只写一条撤销记录；复制时按行输出制表符分隔的文本，并附带一个 `<table>` HTML 片段。嵌套表格的选区通过 `tableRuntimeKey` 区分，按钮状态只使用普通正文表格。

## 图形对象、图表与公式

模型不再丢弃文本框以外的图形内容。导入时，`@arcships/docx-core` 把每个段落里的非图片绘图对象补成新的段落子节点：

- `DrawingRunNode`（`type: "drawing"`）覆盖 `textBox`、`shape`、`wordArt`、`group`、`smartArt`、`chart`、`oleObject` 和 `model3d`。节点带 `widthPx` / `heightPx`、`floating`（`wp:anchor` 的锚点、环绕与 `relativeHeight` 派生的 `zIndex`）、`sourceXml`，以及能解析出来的内容：`chart` 缓存数据（标题、图表类型、分类、每个系列的数值）、`shape` 填充/描边、`ole.progId`、3D 模型的海报图 `src`。
- `MathRunNode`（`type: "math"`）表示 `m:oMath` / `m:oMathPara`，`text` 是线性化后的公式（`a/b+x^2`），`sourceXml` 是原始 OMML。
- 文本框继续沿用已有的合成图片表示（`ImageRunNode.syntheticTextBox`），但内容现在也存在模型里：`ImageRunNode.textBox` 是解析后的段落/文字块，`textBoxSourceXml` 保留原始绘图 XML，因此编辑文本框文字后导出仍能写回文本框。

渲染器对无法直接绘制的对象画占位框而不是留白：`data-docx-drawing`、`data-docx-drawing-kind`、`data-docx-drawing-placeholder`，背后文字的对象额外带 `data-docx-drawing-behind` 和负 `z-index`。图表用缓存数据画一个 SVG 预览（`data-docx-drawing-preview`），3D 模型画海报图，文本框绘制模型里的文字段落（`data-docx-drawing-textbox`），公式按线性文本渲染（`data-docx-math`）。

导出时，仍带 `sourceXml` 的段落原样写出；被编辑重建的段落由序列化补丁把每个对象的原始标记插回它所在的文字位置，所以编辑同一段里别的文字不会丢掉图形/图表/公式。段落数量与部件不一致时补丁放弃该部件，而不是插到错误位置。`w:drawing`、`w:object`、`w:pict` 会包回 `w:r`，`m:oMath` 作为段落内容直接插入；文本框文字被编辑过时，补丁把新文字写进 `w:txbxContent`（`mc:Choice` 与 `mc:Fallback` 两处都写）。

```ts
import { parseChartData, linearizeOmml, drawingObjectPlaceholderStyle } from "@arcships/docx-core"

linearizeOmml("<m:oMath>…</m:oMath>") // "a/b+x^2"
```

段落带 `w:framePr/@w:dropCap` 时，首字母按 `resolveDropCapFontSizePx()` / `resolveDropCapVisualHeightPx()` 解析出的字号和行高浮动显示（`data-docx-drop-cap`、`data-docx-drop-cap-lines`），其余文字照常排版；直接调用 `renderParagraphRuns()` 的宿主可以传 `renderDropCaps: false` 关闭。

## 文档块、词与高亮

`@arcships/docx-core` 把已解析的文档模型拆成可寻址的文档块和词：表格、单元格、页眉页脚、脚注尾注、图片也都包含在块树里。每个块带一个稳定的 source id，查看器把同一个 id 渲染为 `data-docx-source-id`，因此几何解析不需要按文字猜位置。

```ts
import {
  extractDocumentBlocks,
  flattenDocumentBlocks,
  resolveDocxBlockGeometries,
  resolveDocxWordGeometries,
} from "@arcships/docx-core"

const document = await extractDocumentBlocks(model, {
  documentId: "contract.docx",
  includeWords: true,
})
const blocks = flattenDocumentBlocks(document.blocks)
const words = blocks.flatMap((block) => block.words ?? [])

const root = viewerRoot.value // 包含 [data-docx-page-surface] 的查看器根元素
const blockGeometry = resolveDocxBlockGeometries(root, blocks)
const wordGeometry = resolveDocxWordGeometries(root, words)
```

块类型沿用 Extend parser 的 15 个名称（`DOCUMENT_BLOCK_TYPES`）：段落 `text`、`heading`、`section_heading`、`legend`、列表段落（`details.list`）、`figure`、`table`/`table_head`/`table_cell`、表单域 `key_value`、页码字段 `page_number`，以及 `header`/`footer` 容器。空段落和空单元格按上游约定省略，空段落里的图片会提升为同级 figure，source id 仍保留原段落路径。

词分割对空白分隔的文本保留标点，对中日韩文字按字符切分；每个词用 UTF-16 的 `sourceRefs[].start/end` 指回段落，id 形如 `<blockId>/word/{start}-{end}`。几何解析把 DOM 文本范围换算成页面 CSS 像素：页面表面的 client rect 除以其 CSS 尺寸，因此 100%、200% 缩放以及横版/旋转纸张都使用同一套页面坐标；跨页切分的段落每个挂载页各返回一个 fragment。

`@arcships/vue-docx` 提供两个覆盖层组件，几何盒子通过 Teleport 放进对应页面表面，和页面一起缩放：

```vue
<DocxBlockHighlights
  :root="viewerRoot"
  :blocks="blocks"
  :block-ids="selectedBlockIds"
  :page-index="0"
  color="#2563eb"
  @geometry="onBlockGeometry"
/>
<DocxWordHighlights
  :root="viewerRoot"
  :word-ranges="[{ blockId: paragraphId, start: 6, end: 10 }]"
  color="#f59e0b"
/>
```

`blocks`/`block` 传入提取结果；`blockIds` 可以只按 id 选择（没有块列表时 id 直接当 source id）；`pageIndex` 只画某一页；`wordRanges` 不需要先运行词提取。盒子带 `data-docx-block-highlight`、`data-docx-word-highlight`、`data-docx-block-type` 和 `data-docx-highlight-active`，便于宿主样式和测试定位。

`DocxDocumentSurface` 暴露 `getDocumentBlocks()`、`confirmBlockReference()` 和 `confirmWordReference()`：块走结构化定位（`structure`），词走 `text-range` 定位，两者都通过统一引用契约发出 `reference-confirm`，`resolveReference()` 会按当前修订返回 `exact`、`relocated` 或明确失败。

## 直接使用核心能力

纯数据操作从 `@arcships/docx-core/core` 导入；需要加载文件、Worker 和 WASM 时使用 `@arcships/docx-core/runtime`。

```ts
import { createDocxRuntime } from "@arcships/docx-core/runtime"

const runtime = createDocxRuntime()
try {
  const result = await runtime.loadSource({ kind: "bytes", bytes })
  console.log(result.model)
} finally {
  runtime.dispose()
}
```

不要直接组合已经标记弃用的页面、段落、表格和工具栏组件。它们只为 `0.x` 兼容保留，新代码使用 `DocxViewer` 或 `DocxEditor`。
