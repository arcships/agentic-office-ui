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

自建界面同样可以控制数字/日期区域：在 `useXlsxViewerController({ locale: "de-DE" })` 中设置，运行时用 `controller.setDocumentLocale(tag)` 切换；`controller.locale` 报告当前值。`XlsxSheetSurface` 也接受 `locale`/`documentLocale`。显式 `[$-407]` 格式段优先于查看器区域，`@` 文本格式保持键入内容。

公式引用高亮可以接到自建公式输入框上：用 `useFormulaReferenceEditing(controller)` 的 `begin(owner, text, caret)` / `update(...)` / `end(owner)` 维护瞬态编辑会话，内置网格会按会话内容描出当前工作表的引用；`collectFormulaReferences()`、`formulaReferenceSegments()`、`FORMULA_REFERENCE_COLORS` 和 `createNamedRangeResolver(controller)` 负责解析、着色和名称解析。会话只保存草稿文本与光标，不改工作簿；提交或取消时调用 `end()` 即可清除边框。

图表数据源和宿主搜索结果可以走公开的范围高亮：`controller.highlightRanges(ranges, { color, sheetIndex, role, id, label })` 替换宿主区域，`controller.highlightedRanges` 同时给出图表选择派生区域和宿主区域，`controller.clearHighlightedRanges()` 清除当前所有高亮。`XlsxViewer`/`XlsxSheetSurface` 已挂载 `XlsxRangeHighlightLayer`，它会按冻结窗格拆段并跟随滚动、缩放；自建网格可以导出 `createRangeHighlightLayout(sheet, zoom, scrollLeft, scrollTop).resolve(range, viewport)` 复用同一套几何，或用 `createRangeHighlightLayout` + `rangeHighlightStyle()` 自己画。`collectChartRangeHighlights(chart, { selection, sheets, fallbackSheetIndex })` 与 `CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE` 提供与内置图例一致的角色/颜色规则。

图表内嵌图片和 `chartUserShapes` 叠加层随 `XlsxChart.images` / `XlsxChart.drawings` 提供给视图层：`MemoChartSvg` 已把 `a:blipFill` 画成 `<pattern>` + `<image>`，并用 `chartUserShapes` 的 0–1 锚点矩形按预设几何绘制文字框和图片。自建图表渲染器可以直接使用 `XlsxChartUserShape`（`geometry`、`geometryAdjustments`、`rect`、`fill`、`stroke`、`paragraphs`、`textBox`）和 `XlsxChartDrawing`；需要重新解析时调用 `readChartEmbeddedMedia(archive, chartPath, { extentEmu, themePalette })`，它已处理外部关系过滤、图片签名识别和 `CHART_EMBEDDED_IMAGE_MAX_BYTES` / `CHART_EMBEDDED_IMAGE_TOTAL_MAX_BYTES` / `CHART_USER_SHAPES_MAX_BYTES` 上限。
