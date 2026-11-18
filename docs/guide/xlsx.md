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
| `showUpload` | `boolean` | `false` | 上传入口 |
| `getCellStyle` | `function` | 无 | 按单元格补充样式 |

事件包括 `cellDoubleClick`、`upload`、`update:readOnly`、`load-deferred`、`deferred-resumed` 和 `deferred-cancel`。打印时 `@page` 尺寸取自活动工作表已用区域，图表工作表使用 16:9 的自适应高度。

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
