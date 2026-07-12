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
