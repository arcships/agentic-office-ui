# @arcships/vue-xlsx

Vue 3 Excel（XLSX、XLS、XLSB、XLSM、XLTX、XLTM）和 CSV 电子表格查看器，含图表、地图、WebGL 和控制器组合函数。XLSB 转换器仅在对应输入时按需加载；宏内容不会执行。

## 安装

```bash
pnpm add @arcships/vue-xlsx @arcships/xlsx-core
```

## 使用

```vue
<script setup lang="ts">
import { XlsxViewer } from "@arcships/vue-xlsx"
import "@arcships/vue-xlsx/style.css"
</script>

<template>
  <XlsxViewer :src="xlsxUrl" />
</template>
```

## 本地化

`XlsxViewer` 与 `XlsxSheetSurface` 都接受 `locale`（默认 `"en"`，内置 `"zh-CN"`）与 `messages` 覆盖。用户可见文案、表头和 aria-label 来自 `@arcships/office-interaction` 的共享消息目录：

```vue
<XlsxViewer
  :controller="controller"
  locale="zh-CN"
  :messages="{ xlsx: { toolbar: { search: '查找单元格' } } }"
/>
```

## 可选入口

```ts
// 按需导入，不使用时不会进入主包
import "@arcships/vue-xlsx/chart"
import "@arcships/vue-xlsx/map"
import "@arcships/vue-xlsx/webgl"
```

## 组件与组合函数

- `XlsxSheetSurface` — 最小嵌入渲染面（支持受控单元格/范围、行列/工作表对象与区域选择；宿主持有 zoom，并保持指针下的单元格和冻结轴）
- `XlsxViewer` — 电子表格查看器（工具栏含分页预览和打印预览）
- `XlsxPrintPreview` — 按 `pageSetup`、打印区域和手动分页线拆分的分页预览，可直接挂载
- `useXlsxViewerController` — 控制器组合函数（`printLayout`、`printPages`、`print()`）
- `useXlsxViewerThumbnails` — 缩略图

## 事件

| 事件 | 参数 |
|---|---|
| `cellDoubleClick` | `XlsxCellAddress` |
| `tabChange` | `XlsxWorkbookTab`（每次提交的工作表切换一次，含 `kind`） |
| `update:activeTabIndex` | `number`（配合受控 `activeTabIndex` / `v-model:active-tab-index`） |

`tabChange` 与 `update:activeTabIndex` 让宿主可以记录或恢复活动工作表。`activeTabIndex` 是 `controller.tabs` 的下标，工作表和图表工作表共用一条标签条；不传时活动标签仍由控制器拥有。`showSheetTabs="false"` 只隐藏底部标签条，网格、功能区、公式栏保持可用，此时用 `controller.setActiveTabIndex()` 或宿主的标签栏切换。详细说明见 [XLSX 使用指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/guide/xlsx.md)。

`XlsxSheetSurface` 统一支持 `documentId`、`selectionMode`、`emitReferenceCandidates`，并发出 `referenceConfirm`、`regionDraftChange`、`selectionCancel`、`referenceResolve` 和 `referenceError`。组件 expose 可描述、解析和滚动到引用；确认后的引用集合和 Agent 工作流由宿主负责。

## 文档

- [XLSX 使用指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/guide/xlsx.md)
- [组件手册](https://github.com/arcships/agentic-office-ui/blob/master/docs/components/README.md)
- [XLSX 组合函数](https://github.com/arcships/agentic-office-ui/blob/master/docs/composables/xlsx.md)
- [自定义 XLSX 界面](https://github.com/arcships/agentic-office-ui/blob/master/docs/custom-components/xlsx.md)
- [公开接口合同](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/public-api-contract.md)
- [迁移指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/migration-0.2.md)

## License

Apache-2.0
