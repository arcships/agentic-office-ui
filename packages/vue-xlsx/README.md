# @arcships/vue-xlsx

Vue 3 电子表格查看器，含图表、地图、WebGL 和控制器组合函数。

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

## 可选入口

```ts
// 按需导入，不使用时不会进入主包
import "@arcships/vue-xlsx/chart"
import "@arcships/vue-xlsx/map"
import "@arcships/vue-xlsx/webgl"
```

## 组件与组合函数

- `XlsxViewer` — 电子表格查看器
- `useXlsxViewerController` — 控制器组合函数
- `useXlsxViewerThumbnails` — 缩略图

## 事件

| 事件 | 参数 |
|---|---|
| `cellDoubleClick` | `XlsxCellAddress` |

## 文档

- [公开接口合同](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/public-api-contract.md)
- [迁移指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/migration-0.2.md)

## License

MIT
