# 自定义 DOCX 编辑界面

最稳定的方式是让 `DocxEditor` 负责文档表面，自定义组件只操作同一个 `DocxEditorController`。

```vue
<script setup lang="ts">
import { DocxEditor, useDocxEditor } from "@arcships/vue-docx"
import "@arcships/vue-docx/style.css"

const editor = useDocxEditor({ initialFileName: "document.docx" })
</script>

<template>
  <nav class="my-toolbar">
    <button :disabled="!editor.canUndo" @click="editor.undo()">撤销</button>
    <button :disabled="!editor.canRedo" @click="editor.redo()">重做</button>
    <button @click="editor.toggleBold()">粗体</button>
    <button @click="editor.toggleItalic()">斜体</button>
    <button @click="editor.exportDocx()">导出</button>
  </nav>

  <DocxEditor :editor="editor" :show-toolbar="false" />
</template>
```

不要自己复制 `DocxPageSurface`、`DocxParagraphHost` 或 `DocxTableHost` 组成另一套编辑表面。这些低层组件已弃用，而且选区、分页、图片和历史记录必须共享同一个控制器。

只读自定义界面优先使用 `DocxViewer` 的属性和事件。需要从字节得到模型但不显示界面时使用 `useDocxModel()`。

需要保存到服务器或比较两次编辑结果时，不要让宿主拦截浏览器下载：

```ts
const bytes = await editor.getDocxBytes()          // ArrayBuffer，内容与导出文件相同
const blob = await editor.getDocxBlob()            // 带当前容器 MIME 类型的 Blob
await fetch(`/api/documents/${id}`, { method: "PUT", body: bytes })
```

两者复用 `exportDocx()` 的序列化、`basePackage` 和必填表单校验，文档在序列化期间被替换时会以 `DocxExportError`（`ABORTED`）拒绝；`exportDocx()` 本身也走同一条路径，只是最后多创建一次下载链接，并在更新的导出开始时取消上一次待处理请求。

需要自建查找替换界面时，控制器的两个命令足以完成全部操作：

| 命令 | 说明 |
|---|---|
| `replaceMatch(target, text)` | 替换一条搜索结果；`target` 由 `findDocxSearchMatches()` 的命中项经 `docxReplaceTargetFromMatch()` 转换得到（页眉/页脚/注释结果返回 `undefined`）。一条撤销记录。 |
| `replaceAll(query, text, options?)` | 替换正文中全部命中，`options` 与搜索选项相同（`caseSensitive` / `wholeWord` / `regex`）；返回 `{ replaced, skipped }`，整批一条撤销记录。 |

内置查找栏本身也是普通 Vue 组合（`useDocxFindReplace`）加 `DocxToolbar` 的 `find` 插槽；宿主可以复用 `OfficeFindBar`，也可以只调控制器命令自己绘制界面。

需要自建图片图层菜单时，`bringToFront(location)`、`bringForward(location)`、`sendBackward(location)`、`sendToBack(location)` 分别对应 Word 的四个叠放命令；`resolveDocxImageLayerOrderState(floating)` 和 `resolveDocxImageLayerOrderMenuActions(floating)` 给出与内置菜单相同的可用/禁用状态。每个命令是一条撤销记录，导出时写成 `wp:anchor/@relativeHeight`。

需要自建表格选区界面时，控制器暴露同一套矩形选区 API：

| 命令/属性 | 说明 |
|---|---|
| `selectedTableCellRange` | 当前多单元格矩形；`cellRange.gridRange` 给出物理网格列范围，跨 `gridSpan` / `vMerge` 时已展开。 |
| `selectTableCellRange(tableIndex, range, tableRuntimeKey?)` | 设置锚点/焦点矩形；普通正文表格会按模型网格归一化，嵌套表格用 `tableRuntimeKey` 区分。 |
| `selectWholeTable(tableIndex?)` | 选中整张表。 |
| `clearTableSelection()` | 通过 `clearTableCellContents(tableIndex, cells[])` 清空矩形内所有单元格，返回是否执行；整个矩形一条撤销记录。 |
| `collapseTableCellSelection()` | 把矩形折叠到锚点单元格（内置编辑器里是 `Esc`），返回是否执行。 |
| `isCellSelected(selection, tableIndex, rowIndex, cellIndex, table?)` | 判断单元格是否在选区中；传入 `table` 时按物理网格处理合并单元格，否则退回模型行列范围。 |

## 复用内置编辑辅助界面

五个编辑辅助能力也单独导出，自定义表面可以只挂其中几个，全部通过同一个 `DocxEditorController` 工作：

```vue
<script setup lang="ts">
import {
  DocxEditor,
  DocxFormFieldDialog,
  DocxLinkHoverCard,
  DocxStylePreviewPopover,
  resolveFormFieldAt,
  useDocxEditor,
  useDocxNodeReorder,
} from "@arcships/vue-docx"

const editor = useDocxEditor({ initialFileName: "document.docx" })
</script>
```

| 组件 / 组合式 | 用途 |
|---|---|
| `DocxFormFieldDialog` | 传入 `field`、`location` 和 `controller`，保存时调用 `controller.updateFormField()` |
| `DocxLinkHoverCard` | 传入链接地址与屏幕坐标；`open`/`edit`/`remove` 事件由宿主转给 `controller.setLink()` |
| `DocxStylePreviewPopover` | 传入段落样式列表和悬停样式 id，配合 `docxStylePreviewCss()` 渲染 |
| `useDocxNodeReorder(controller)` | 悬停、拖拽、插入位置和 `moveBodyNode()` 提交；纯函数 `resolveDocxReorderInsertionIndex()`、`resolveDocxReorderTarget()` 可用于自定义手柄 |

`DocxFormFieldDialog`、`DocxLinkHoverCard` 和 `DocxStylePreviewPopover` 与 `DocxEditor` 使用同一份文案库（`provideDocxMessages`），在 `DocxEditor` 内部挂载时自动跟随 `locale`/`messages`。

只要宿主自己渲染手柄，就可以关闭 `DocxEditor` 的对应属性（`show-node-reorder`、`show-style-preview` 等），不会出现两套控件。

## 内部书签导航

`DocxViewer`、`DocxEditor` 和 `DocxDocumentSurface` 都用 `revealBookmark(name)` 暴露内部链接导航：滚动到书签所在页和段落、短暂高亮目标，并在可编辑表面上放置光标。`name` 可写成 `Name` 或 `#Name`，查找不区分大小写并接受 URL 转义；返回 `{ status: "revealed", nodeIndex, pageIndex, text }` 或 `{ status: "missing", name }`。

查看器需要挂载新页时向上发出 `requestPageReveal(pageIndex)`；书签不存在时不抛错，而是显示非阻塞提示并发出 `bookmarkMissing(name)`。链接悬浮卡片的“打开”动作对内部链接也调用同一导航，并保留 `link-open` 事件。

自定义菜单或目录可以直接使用 `@arcships/docx-core` 的 `createDocxBookmarkIndex(model)`（`get` / `has` / `targets`）和 `normalizeDocxBookmarkName()`，无需自己解析 `sourceXml`。

## 文档块/词高亮覆盖层

`DocxBlockHighlights` 和 `DocxWordHighlights` 是独立覆盖层，可以挂在 `DocxViewer`、`DocxEditor` 或 `DocxDocumentSurface` 的根元素上：

```vue
<DocxBlockHighlights
  :root="viewerRoot"
  :blocks="blocks"
  :block-ids="blockIds"
  :page-index="pageIndex"
  :color-for-block="(block) => block.type === 'table' ? '#0ea5e9' : '#2563eb'"
  @geometry="(map) => (geometry = map)"
/>
<DocxWordHighlights
  :root="viewerRoot"
  :words="words"
  :word="activeWord"
  :word-ranges="wordRanges"
/>
```

| 属性 | 组件 | 说明 |
|---|---|---|
| `root` | 两者 | 含 `[data-docx-page-surface]` 的查看器根元素 |
| `block` / `blocks` / `blockIds` | `DocxBlockHighlights` | 强调块、待绘制块，或只按 source id 选择 |
| `word` / `words` / `wordRanges` | `DocxWordHighlights` | 强调词、待绘制词，或 `{ blockId, start, end }` 偏移 |
| `pageIndex` | 两者 | 只绘制指定 0 基页 |
| `color` / `colorForBlock` | 两者 | 边框与底色；底色用 8 位十六进制透明度 |
| `onGeometry` | 两者 | 每次测量后回调页内几何 |

盒子通过 Teleport 进入页面表面，因此自动复用页面缩放，不需要宿主换算 zoom。几何在页面挂载、重分页、字体加载、滚动和窗口缩放后自动刷新；`onGeometry` 可以用来保存引用预览或调试。

`DocxDocumentSurface` 的 `getDocumentBlocks()` 返回带 source id 的块树，`confirmBlockReference(block)` 和 `confirmWordReference(block, word)` 会以统一引用契约发出 `reference-confirm`，可直接用于引用条或引用列表。
