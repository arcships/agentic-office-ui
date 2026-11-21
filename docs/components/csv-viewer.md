# CSV Viewer

`CsvViewer` 是 `@arcships/vue-ui` 的只读分隔文本查看器：CSV/TSV 解析、虚拟化表格、单元格搜索、缩放、上传和下载。它复用共享 Office chrome，不依赖 XLSX controller，也不加载 Glide canvas。

决策与视觉差异见 [CSV Viewer ADR](../design/csv-viewer-adr.md)；`XlsxViewer` 的 CSV 兼容路径见 [XLSX guide](../guide/xlsx.md#csv-输入)。

## 最小用法

```vue
<script setup lang="ts">
import { CsvViewer } from "@arcships/vue-ui"
import "@arcships/vue-ui/style.css"
import "@arcships/vue-ui/theme.css"

const csv = `Name,Team,Amount
Alice,Alpha,"1,200"
Bob,Beta,900`
</script>

<template>
  <CsvViewer :data="csv" file-name="pipeline.csv" search />
</template>
```

默认高度为 560px。宿主可以用 `height`（数字按 px，字符串按 CSS length）或 class 覆盖；class 也可以设置 `--csv-viewer-height`（例如 `--csv-viewer-height: 70vh`）：

```vue
<CsvViewer :data="csv" :height="700" class="my-csv" />
```

## Props

| Prop | 类型 | 默认 | 说明 |
|---|---|---|---|
| `data` | `string` | — | 分隔文本；优先于 `file` / `src` |
| `file` | `Blob \| null` | — | 初始文件；UTF-8 BOM、UTF-16LE/BE BOM 自动解码 |
| `src` | `string` | — | 远程 URL；`fetch` 失败进入 error 状态 |
| `fileName` | `string` | — | 下载名与 `getState().fileName`；`.tsv` 自动按 tab 解析 |
| `delimiter` | `string` | `"auto"` | `auto` 或单字符；支持 `tab`、`\\t` |
| `search` | `boolean` | `true` | 搜索浮层；上游 React block 默认 `false` |
| `searchDebounce` | `number` | `300` | 输入防抖（毫秒） |
| `height` | `number \| string` | CSS `560px` | 根高度 |
| `maxFileSize` | `number` | 50 MiB | 上传上限 |
| `loading` | `boolean` | `false` | 宿主异步加载时强制 loading |
| `locale` / `messages` | `OfficeLocale` / `OfficeMessagesOverride` | 继承 | 共享消息目录 |
| `class` | `string` | — | 根 class |

## 事件

| 事件 | 参数 | 说明 |
|---|---|---|
| `load` | `{ fileName, headers, rows, delimiter }` | 每次成功解析后发出 |
| `error` | `CsvViewerError` | `PARSE_ERROR` / `UNTERMINATED_QUOTE` / `INVALID_DELIMITER` / `READ_ERROR` / `FETCH_ERROR` / `FILE_TOO_LARGE` |
| `download` | `{ fileName, text }` | 浏览器保存前发出，宿主可接管下载 |
| `update:data` | `string` | 上传或 URL 读取后的原始文本 |

## 实例方法

```ts
const viewer = ref<InstanceType<typeof CsvViewer> | null>(null)
viewer.value?.runSearch("alice")
viewer.value?.nextResult()
viewer.value?.setZoom(150)
viewer.value?.download()
viewer.value?.getState()
```

`getState()` 返回 `{ status, fileName, headers, rows, delimiter, zoom, searchQuery, searchResults, activeResultIndex, error }`，供宿主同步状态或写验收断言。

## 状态

| 状态 | 触发 | 渲染 |
|---|---|---|
| `empty` | 没有数据，或只有表头/空行 | `OfficeEmptyState` + 上传按钮 |
| `loading` | `loading` prop、上传读取中、`src` 拉取中 | `OfficeLoadingState` |
| `error` | 解析失败、读取失败、HTTP 失败、超限 | `OfficeErrorState` + 重试/上传 |
| `ready` | 至少一行数据 | 工具栏 + 虚拟化表格 |

## 解析合同

- 第一条非空记录是表头；空行按 greedy 方式跳过。
- RFC 4180：`""` 转义、引号内逗号/换行、CRLF/CR/LF。
- 列数取表头与数据行的最大值；短行补空，超出表头的列标为 `Extra N`。
- `auto` 在前 8 条记录内统计 `,`、`\t`、`;`、`|`，平手时优先逗号；`.tsv` 文件名优先 tab。
- `decodeCsvBuffer()` 识别 UTF-8 BOM、UTF-16LE/BE BOM；没有 BOM 时按 UTF-8。
- `serializeCsv()` 用 RFC 4180 引号规则，默认 `\r\n`。

## 搜索与虚拟化

- 搜索是大小写不敏感的子串匹配，扫描所有数据行；大表按 500 行分批让出事件循环。
- `nextResult()` / `previousResult()` 循环切换，活动单元格滚动到行窗口内。
- 表格只渲染视口附近的行（默认 overscan 6 行）；上下 spacer row 保持滚动条总高度。
- 窄视口或 200% 缩放时仍使用固定行高，列宽按 zoom 缩放。

## 可访问性

- `<table role="grid">`，列头 `role="columnheader"`，行号 `role="rowheader"`，单元格 `role="gridcell"`；选中单元格 `aria-selected`。
- 工具栏、缩放、搜索、动作菜单全部来自共享 chrome，标签走 `vueUi.csvViewer.*` 消息目录（en / zh-CN）。
- 上传 input 有可访问名称；错误状态是 `role="alert"`，加载状态是 `role="status"`。
