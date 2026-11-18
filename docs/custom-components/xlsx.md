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

需要宿主自己的标签栏时，关闭内置标签条并受控地持有活动标签：

```vue
<XlsxViewer
  :controller="controller"
  :show-sheet-tabs="false"
  v-model:active-tab-index="activeTab"
  @tab-change="onTabChange"
/>
```

`showSheetTabs=false` 只隐藏底部标签条，网格、功能区、公式栏和控制器状态都还在；活动标签由 `activeTab` 控制（`undefined` 时由控制器持有），`tabChange` 在每次提交的切换后返回完整标签（含 `kind`）。`activeTabIndex` 是 `controller.tabs` 的下标，不是 `controller.sheets` 的下标。

不要用已经弃用的 `XlsxGrid`、`XlsxFormulaBar` 和浮层组件重新拼装查看器。它们依赖大量内部协调状态。需要缩略图时，从同一个控制器调用 `useXlsxViewerThumbnails()`。

自建查找替换界面时，用 `useXlsxSearch()` 维护结果（`search` / `searchNext` / `activateSearchMatch` / `clearSearch`），用 `controller.replaceInSheet(query, replacement, options)` 写入：`options` 支持 `matchCase`、`wholeCell`、`formula`、`sheetIndex`、`range`，以及精确到单处命中的 `occurrence: { cell, start, end }`。一次调用一条撤销记录；返回 `{ replacedCells, replacedOccurrences, changed }` 供宿主回报结果。改变工作簿后 `revision` 会变化，`useXlsxSearch` 会清空旧结果，所以替换后重新 `search()` 即可刷新列表。

剪切是控制器上的状态，不是选区状态：`cutSelection()` 进入剪切、`cutState` 给出 `{ range, sheetIndex }`（没有待处理的移动时为 `null`）、`cancelCut()` 取消。自建网格可以据 `cutState` 画自己的 marquee 和“取消剪切”按钮；`pasteFromClipboard()` 搬移内容并在同一个撤销条目里清空源区域，`copySelectionToClipboard()` 与仅格式的 `pasteSpecialFromClipboard({ mode: "formats" })` 会结束剪切状态。内置网格已经把 `Ctrl/Cmd+X`、`Escape` 和虚线标记接到这套状态上。
