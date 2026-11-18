# XLSX 使用指南

`XlsxViewer` 接收一个 `XlsxViewerController`。控制器负责文件、工作表、选区、编辑命令和历史记录；组件负责界面。

支持 `.xlsx`、`.xls`、`.xlsb`、`.xlsm`、`.xltx`、`.xltm` 和 `.csv`。XLSB 会按需转换为内存 XLSX 后进入现有解析链路；宏工作簿和宏模板只读取表格内容，不执行 VBA；模板按普通工作簿打开。

```vue
<script setup lang="ts">
import { useXlsxViewerController, XlsxViewer } from "@arcships/vue-xlsx"
import "@arcships/vue-xlsx/style.css"

const props = defineProps<{ bytes: ArrayBuffer }>()
const controller = useXlsxViewerController({
  file: props.bytes,
  fileName: "workbook.xlsx",
})
</script>

<template>
  <XlsxViewer
    :controller="controller"
    height="720px"
    :read-only="false"
    @cell-double-click="onCellDoubleClick"
  />
</template>
```

常用属性：

| 属性 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `controller` | `XlsxViewerController` | 必填 | 查看器控制器 |
| `height` | `string` | `100%` | 查看器高度 |
| `minHeight` | `string \| number` | 无 | 可选最小高度（数字按 px）；不传时组件不设置最小高度 |
| `readOnly` | `boolean` | `false` | 禁止编辑 |
| `showDefaultToolbar` | `boolean` | `true` | 默认工具栏 |
| `showRibbon` | `boolean` | `true` | 功能区 |
| `showFormulaBar` | `boolean` | `true` | 公式栏 |
| `showImages` | `boolean` | `true` | 图片和绘图对象 |
| `showSheetTabs` | `boolean` | `true` | 底部工作表标签条 |
| `activeTabIndex` | `number` | 无 | 受控的活动标签下标；不传时由控制器持有 |
| `showUpload` | `boolean` | `false` | 上传入口 |
| `getCellStyle` | `function` | 无 | 按单元格补充样式 |

事件包括 `cellDoubleClick`、`upload`、`update:readOnly`、`tabChange`、`update:activeTabIndex`、`load-deferred`、`deferred-resumed` 和 `deferred-cancel`。打印时 `@page` 尺寸取自活动工作表的 `pageSetup` 纸张与方向；图表工作表使用 16:9 的自适应高度。

## 单元格编辑

双击单元格、按 `Enter` 或 `F2` 打开内联编辑器。编辑器是 `textarea`，可以输入多行：

- `Enter` 提交；`Alt+Enter`（macOS 为 `Option+Enter`）在光标处插入硬换行，编辑器保持打开。
- `Escape` 放弃草稿，单元格保持原值；`Tab` 提交并移到右侧单元格继续编辑。
- 草稿没有硬换行时，`↑`/`↓` 提交并上下移动选区；有硬换行时方向键在草稿的行间移动光标。
- 粘贴多行文本会保留换行（`CRLF` 统一为 `LF`）；单行粘贴仍走浏览器原生路径。
- 输入法组字期间的 `Enter`/`Escape` 由输入法处理，组字结束后只提交一次。

编辑器用一个隐藏的镜像层按当前宽度测量草稿的换行高度：硬换行或自动换行后编辑器向下增高（裁剪在网格可视区内），不会把后续行藏在单行滚动里。提交后的换行会写进单元格值；要让换行在单元格里显示为多行，请启用“自动换行”并给行足够的高度——未启用时与 Excel 一致，只显示第一行。

## 工作表标签

`showSheetTabs` 设为 `false` 只移除底部标签条：网格、功能区、公式栏和控制器状态都不变，`Ctrl/Cmd+PageUp/PageDown` 仍然可以切换工作表。此时切换只能通过宿主自己的界面调用 `controller.setActiveTabIndex(index)`。需要完全自定义的导航界面时，直接用 `XlsxSheetSurface` + `useXlsxViewerController`（见[自定义 XLSX 界面](../custom-components/xlsx.md)）。

需要宿主记录或恢复活动工作表时，用 `v-model:active-tab-index` 加 `tabChange`：

```vue
<script setup lang="ts">
import { ref } from "vue"
import { XlsxViewer, type XlsxWorkbookTab } from "@arcships/vue-xlsx"

const activeTab = ref(0)

function onTabChange(tab: XlsxWorkbookTab) {
  // kind 区分工作表与图表工作表；index 是 tabs 下标
  console.log(tab.kind, tab.name, tab.index, tab.sheetIndex, tab.chartsheetIndex)
}
</script>

<template>
  <XlsxViewer
    v-model:active-tab-index="activeTab"
    :controller="controller"
    @tab-change="onTabChange"
  />
</template>
```

- `activeTabIndex` 是 `controller.tabs` 的下标。工作表和图表工作表共用同一条标签条，所以它**不是** `controller.sheets` 的下标；越界值会被忽略，与 `setActiveTabIndex()` 一致。
- 不传（或传 `undefined`）时活动标签由控制器持有，行为与旧版本完全一致。传数字时，值变化会同步到控制器；控制器内部切换后发出 `update:activeTabIndex` 供 `v-model` 更新。受控值可以在工作簿解析完成前传入，首个标签出现时会应用一次。
- `tabChange(tab)` 在每次真正提交的切换后触发一次，参数是完整的 `XlsxWorkbookTab`（含 `kind`、`id`、`index`、`sheetIndex`/`chartsheetIndex`）。打开工作簿时的首个标签、重命名/排序/隐藏导致的标签重建、普通重渲染都不会触发。
- 宿主自己写入 `activeTabIndex` 引起的切换不会回发 `tabChange`，避免路由或日志出现回环。
- 拖拽排序活动标签时它的 `index` 会变化但没有切换工作表：只发 `update:activeTabIndex`，不发 `tabChange`。

## 查找与替换

工具栏的放大镜按钮打开查找面板。搜索覆盖工作簿所有工作表（值和公式），命中后切到对应工作表并选中单元格；大小写/全字/正则开关与结果列表都在同一个面板里。

面板同时提供替换：输入替换文字后可点“替换”（仅改写当前选中结果）或“全部替换”（改写当前工作表范围内的所有命中）。下方两个复选项控制替换范围：

- “匹配整个单元格”：只有单元格文本与查询完全相等才替换。
- “包含公式”：默认开启，公式单元格按公式文本替换（例如 `=B2*22%` → `=B2*25%`）；公式单元格的缓存值不会被当成普通文字替换。关闭后只替换普通值。

替换命令在控制器上，宿主可以绕开内置面板直接调用：

```ts
const result = controller.replaceInSheet("PDF Viewer", "PDF Reader", {
  matchCase: false,   // 默认 false，与 Excel 一致
  wholeCell: false,   // 默认 false：替换单元格内的子串
  formula: true,      // 默认 true：同时替换公式文本
  sheetIndex: 0,      // 默认活动工作表；非活动表会先被激活
  range: { start: { row: 1, col: 0 }, end: { row: 40, col: 7 } }, // 默认整张已用区域
})
// result: { sheetIndex, workbookSheetIndex, replacedCells, replacedOccurrences, changed }
```

每次调用只写一条撤销记录：一次“全部替换”无论改了多少单元格，撤销一次即可全部还原，重做重新应用。替换后的值按键入规则转换（数字写回数字），并随 `exportXlsx()` 写入导出文件；`readOnly`、受保护工作表和 Worker 只读大文件模式不改动工作簿，直接返回 `changed: false`。`occurrence` 选项可精确到搜索结果报告的那一处：

```ts
controller.replaceInSheet("PDF Viewer", "PDF Reader", {
  occurrence: { cell: { row: 4, col: 2 }, start: 0, end: 10 },
})
```

## 打印与分页

工作表的 `pageSetup`、`printOptions`、`pageMargins`、`rowBreaks`/`colBreaks` 以及 `_xlnm.Print_Area`/`_xlnm.Print_Titles` 会解析进 `XlsxSheetData`。分页在 `@arcships/xlsx-core` 内完成，Vue 层只消费页面结果，不从 DOM 高度推导分页：

```ts
import { buildWorksheetPrintLayout } from "@arcships/xlsx-core"

const layout = buildWorksheetPrintLayout(controller.activeSheet)
// layout.pages: { range, paperSize, widthPx, heightPx, scale, rowPage, colPage, ... }[]
```

控制器提供同一份结果和打印动作：

| 命令 | 说明 |
|---|---|
| `printLayout` / `printPages` | 活动工作表的扁平页面列表；每页带单元格范围和纸张尺寸 |
| `getPrintLayout()` | 同 `printLayout`，适合在命令式代码里取一次快照 |
| `setPageBreakPreview(enabled)` | 在网格中显示 Excel 风格的蓝色虚线分页线和页码 |
| `openPrintPreview()` / `closePrintPreview()` | 打开/关闭分页打印预览；`XlsxPrintPreview` 也可由宿主单独挂载 |
| `print()` | 打开打印预览并调用 `window.print()`；无工作表或无 `window.print` 时返回 `false` |

`XlsxViewer` 的 `@media print` 只输出预览页面，每页一个区块；纸张方向、页边距、手动分页线、重复标题行/列和合并单元格跨页都由核心分页结果决定。导出 XLSX 时会保留 `pageMargins`、`printOptions`、`pageSetUpPr`、分页线和打印区域定义，重新打开后分页结果一致。

图表、地图和 WebGL 是按需入口：

```ts
import "@arcships/vue-xlsx/chart"
import "@arcships/vue-xlsx/map"
import "@arcships/vue-xlsx/webgl"
```

不使用这些功能时，不需要加载相应的大型实现。旧的网格、工具栏和浮层组件只为兼容保留，新代码从 `XlsxViewer` 和控制器开始。

## 作者入口

功能区新增“图表 / 批注 / 超链接 / 表格 / 条件格式 / 数据验证 / 冻结窗格”一组按钮，右键菜单提供同名入口。对话框由 `XlsxViewer`（以及最小化的 `XlsxSheetSurface` 右键菜单）自动渲染，宿主也可以直接挂载 `XlsxAuthoringDialog`：

```vue
<script setup lang="ts">
import { XlsxAuthoringDialog, XlsxViewer } from "@arcships/vue-xlsx"
</script>

<template>
  <XlsxViewer :controller="controller" />
  <!-- 或自行控制 dialog：kind = "chart" | "comment" | "hyperlink" | "table" | "conditional" | "validation" -->
  <XlsxAuthoringDialog :controller="controller" kind="comment" @close="close" />
</template>
```

控制器 API 与对话框一一对应，宿主可以完全不用内置 UI：

| 命令 | 说明 |
|---|---|
| `insertChart({ family, range, title })` | 从区域插入柱形/折线/饼图，返回新图表 |
| `getCellComment(cell?)` / `setCellComment(cell, input \| null)` | 读取、写入或删除单元格批注 |
| `getCellHyperlink(cell?)` / `setCellHyperlink(cell, input \| null)` | 读取、写入或移除超链接 |
| `getConditionalFormatRules()` / `addConditionalFormatRule(draft)` / `updateConditionalFormatRule(index, draft)` / `deleteConditionalFormatRule(index)` | 条件格式（色阶、数据条、图标集，含 cfvo 阈值） |
| `createTable({ range, name, hasHeaders })` / `removeTable(name)` / `getTableAtCell(cell?)` | 从区域创建/删除 ListObject |
| `setSelectionDataValidation(input \| null)` | 为选区写入或清除数据验证 |
| `freezePanes` / `setFreezePanes(freeze)` / `toggleFreezePanesAtActiveCell()` | 冻结窗格状态与切换 |

每个命令都会改动工作簿模型、刷新视图、写一条撤销记录，并随导出的 XLSX 保存；导出后再打开仍然保留。命令只在可写的主线程工作簿上生效：`readOnly`、受保护工作表或 Worker 只读大文件模式下返回 `false` / `null`，不做任何修改。详细实现见 [XLSX 作者接口设计](../xlsx-authoring-surface.md)。

## 剪切与移动单元格

选区按 `Ctrl/Cmd+X`、右键菜单“剪切”或功能区“剪切”进入剪切状态：源单元格暂时保留内容，网格在源区域上画虚线“marching ants”，辅助技术会播报“已剪切 K1:K2，按 Esc 取消”。真正的移动发生在下一次粘贴：

| 操作 | 行为 |
|---|---|
| `Ctrl/Cmd+V` / `pasteFromClipboard()` | 值、公式和样式写到目标位置，公式按位移换算；源区域在**同一个撤销条目**里清空 |
| 一次 `Ctrl/Cmd+Z` | 同时恢复源区域并撤掉目标内容；一次重做又把整段移动放回去 |
| `Ctrl/Cmd+C` 或再次剪切 | 结束旧的剪切状态，按普通复制或新的剪切处理 |
| 仅格式粘贴（`pasteSpecialFromClipboard({ mode: "formats" })`） | 只写样式，源内容保留，剪切状态结束 |
| `Escape` / `cancelCut()` | 取消剪切状态，工作簿不变；之后粘贴按系统剪贴板内容当普通复制处理 |

控制器接口：`cutSelection()` 返回是否进入剪切状态，`cutState` 给出 `{ range, sheetIndex }`（源区域和它所在的工作簿工作表），`cancelCut()` 返回是否真的取消了剪切。跨工作表剪切是允许的：虚线只画在源工作表上，粘贴到其他工作表后仍然清空源区域。重新加载或重建工作簿（撤销快照、重命名、换文件）会自动清除剪切状态，不会把旧文件的数据粘到新文件里。

## 单元格格式

功能区的合并与边框组对应 Excel 的同一组操作：

| 控件 | 命令 | 说明 |
|---|---|---|
| 边框预设下拉 | `setSelectionBorders({ preset, style, color })` | `preset` 为 `all` / `outside` / `left` / `right` / `top` / `bottom` / `none`；`outside` 只画选区外围，内部格子保持原样 |
| 边框颜色 | 同上 | 颜色写入预设涉及的每一面；默认细线灰色 `A1A1AA` |
| 所有边框 | `setSelectionBorders({ preset: "all" })` | 保留的快捷按钮，同样是上面那条命令 |
| 合并并居中 | `mergeAndCenterSelection()` | 合并选区并写入 `horizontal: center` + `vertical: center`，一条撤销记录 |
| 取消合并 | `unmergeSelection()` | 只拆合并区 |

只想合并、不改对齐的宿主用 `mergeSelection()`。边框命令都走现有的样式命令，因此跨越导出/重新打开后保留：预设会写成 `styles.xml` 里对应边的 `<border>`，颜色是 8 位 ARGB。`XlsxCellBorderStyleInput` 仍然允许逐边不同的样式与颜色，需要时直接调 `setRangeStyle(range, { border })`。

## 延迟解析

控制器配置了 `deferLoadingAboveBytes` 且文件超过阈值时，不会立即解析工作簿。`XlsxViewer` 会直接渲染一个内置提示卡片，显示文件名和大小，并提供“继续解析”和“取消”两个动作；宿主不需要再额外实现加载入口。

行为：

- “继续解析”调用 `controller.continueDeferredLoad()`，事件 `deferred-resumed` 在调用前触发；
- “取消”调用 `controller.cancelDeferredLoad()`，释放缓存的源文件字节并回到空查看器，事件 `deferred-cancel` 在调用前触发；
- 进入延迟状态时组件触发 `load-deferred`，事件参数为 `{ fileName, bytes }`；
- 根元素暴露 `data-state="deferred"`，便于宿主和端到端测试识别。

如需替换默认卡片，使用 `deferredLoad` 插槽：

```vue
<XlsxViewer :controller="controller">
  <template #deferredLoad="{ fileName, bytes, sizeLabel, resume, cancel }">
    <div>
      <p>{{ fileName }}（{{ sizeLabel }}）较大，是否继续解析？</p>
      <button type="button" @click="resume">继续解析</button>
      <button type="button" @click="cancel">取消</button>
    </div>
  </template>
</XlsxViewer>
```

插槽属性：`fileName`、`bytes`、`sizeLabel`，以及 `resume`、`cancel` 两个动作函数。

## CSV 输入

控制器可以直接打开 CSV。接入方仍然传入 `file`/`src`；本地 `ArrayBuffer` 需要同时传入以 `.csv` 结尾的 `fileName`：

```ts
const controller = useXlsxViewerController({
  file: await csvFile.arrayBuffer(),
  fileName: csvFile.name,
})
```

远程 CSV 根据最终 URL 的 `.csv` 扩展名或响应的 `Content-Type: text/csv` 识别。CSV 会作为单工作表打开，支持 UTF-8（含 BOM）以及带 BOM 的 UTF-16；格式、图表和多工作表等 XLSX 专属信息不属于 CSV 文件内容。
