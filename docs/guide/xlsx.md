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
