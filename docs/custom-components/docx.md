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
