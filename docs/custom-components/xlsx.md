# 自定义 XLSX 界面

`XlsxViewer` 已把网格、图表、图片和选区绑定到控制器。自定义工具栏时保留查看器，只关闭内置区域。

```vue
<script setup lang="ts">
import { XlsxViewer, useXlsxViewerController } from "@arcships/vue-xlsx"
import "@arcships/vue-xlsx/style.css"

const props = defineProps<{ bytes: ArrayBuffer }>()
const controller = useXlsxViewerController({ file: props.bytes })
</script>

<template>
  <nav>
    <button :disabled="!controller.canUndo" @click="controller.undo()">撤销</button>
    <button :disabled="!controller.canRedo" @click="controller.redo()">重做</button>
    <button :disabled="!controller.canExport" @click="controller.exportXlsx()">导出</button>
    <span>{{ controller.selectedRangeAddress ?? "未选择" }}</span>
  </nav>

  <XlsxViewer
    :controller="controller"
    :show-default-toolbar="false"
    :show-ribbon="false"
    height="720px"
  />
</template>
```

上例的控制器以创建时的文件为来源。更换 `bytes` 时应重新创建承载该控制器的子组件，不要直接修改控制器内部状态冒充换文件。

不要用已经弃用的 `XlsxGrid`、`XlsxFormulaBar` 和浮层组件重新拼装查看器。它们依赖大量内部协调状态。需要缩略图时，从同一个控制器调用 `useXlsxViewerThumbnails()`。
