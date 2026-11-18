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
