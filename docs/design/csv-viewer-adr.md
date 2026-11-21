# ADR: CSV Viewer block（#178）

- 状态：**Accepted**（0.8.1 实现，0.9 保持）
- 日期：2026-11-21
- 关联：[#178](https://github.com/arcships/agentic-office-ui/issues/178)、[#160](https://github.com/arcships/agentic-office-ui/issues/160)、[#166](https://github.com/arcships/agentic-office-ui/issues/166)
- 上游基准：[`extend-hq/ui@70fe70c6f1e1eee53954c6923f3d9f583557a57e`](https://github.com/extend-hq/ui/tree/70fe70c6f1e1eee53954c6923f3d9f583557a57e)（MIT）`apps/v4/components/extend/csv-viewer.tsx`（975 行）

## 1. 背景

#166 的组件映射把 CSV 写成“XLSX 的 CSV 路径（单独 block 待评估）”，但一直没有评估结论。#160 交付的 `XlsxViewer` 可以打开 CSV：`Workbook.fromCsvString` 将文本转成单工作表，格式、公式、图表、多工作表等 XLSX 专属信息在转换时丢失。它解决的是“把 CSV 当工作表看”，不是上游 CSV Viewer 的独立组件合同：

- `data?: string` 直接接收分隔文本；上传文件走 `file.text()`，不创建 Workbook。
- toolbar 是 zoom + search + upload/download actions，不经过 XLSX ribbon / sheet tabs / formula bar。
- 表格是只读 canvas，不提供单元格编辑、range、公式、sheet 切换。
- 空态、解析错误、文件菜单都是 CSV 自己的状态。

因此“复用 XlsxViewer”在功能上可用，但在 API、视觉、状态和测试面上不是同一个 block。0.8.1 的架构审计（[`viewer-editor-architecture.md`](viewer-editor-architecture.md)）将其归类为 `product-gap`，而不是 `not-used` 或私有 runtime 限制。

## 2. 选项

| 选项 | 内容 | API | 视觉 | 编辑能力 | 依赖 / 维护成本 | 结论 |
|---|---|---|---|---|---|---|
| A. 只保留 `XlsxViewer` CSV 路径 | 继续 `Workbook.fromCsvString` | `XlsxViewer` 的 file/src/controller API；CSV 没有独立 entry | XLSX chrome（ribbon、sheet tabs、状态栏），与上游 CSV 不同 | 意外获得整套 XLSX 编辑能力，但 CSV 的格式信息已丢失 | 0 新依赖；限制只能写文档 | 保留为兼容路径，不作为 CSV block |
| B. 独立 `CsvViewer`（采用） | Vue 组件 + 共享 Office chrome + 语义表格 | `data` / `file` / `src` / `fileName` / `delimiter` / `search` / `height` | 560px surface、48px toolbar、语义 grid；不加载 Glide canvas | **只读**：选择单元格、搜索、缩放、上传/下载；没有编辑事务 | 无新运行时依赖；解析器与表格 ~1k 行，复用 `OfficeToolbar`/`OfficeZoomControl`/`OfficeSearchPopover`/states | **接受** |
| C. 不提供独立 block | 明确 `not planned` + 产品理由 | — | — | — | 最低 | 拒绝：上游公开可移植，宿主需要“看 CSV”而不创建 Workbook |

## 3. 决定

采用 **B：独立 `@arcships/vue-ui` 的 `CsvViewer`**，同时保留 A 作为 `XlsxViewer` 的兼容 CSV 路径。

为什么不是 vendor Glide：

1. **第二套 canvas 网格运行时**与现有 XLSX canvas/DOM 网格、DOCX 分页、PDF 光栅化并列，增加主题、缩放、搜索、可访问性和虚拟滚动的重复实现。
2. Glide 6.0.4-alpha24 是 alpha 版本；把它固定为新的公开依赖，会把 alpha 的 breaking 风险带进 0.8 公共包。
3. CSV 的表格是**只读、固定行高、无合并单元格**，语义 `<table>` + 行窗口足以覆盖；搜索命中可以直接用 `<mark>` 表达，屏幕阅读器也能读到真实单元格。
4. 上游 `papaparse` 的解析语义（header、greedy empty lines、`__parsed_extra`）已经在本地解析器中复现；没有新增第三方许可证义务。若未来需要 Papa 的完整错误恢复，可在不改变组件 API 的前提下替换 `parseCsvText()`。

**结论有效期**：0.8.1 实现；0.9 继续维护。只有出现“CSV 需要冻结列/行、列宽拖拽、单元格编辑、百万行级别性能”这类明确需求时，才重新评估 Glide 或虚拟表格依赖。

## 4. 公共 API

```ts
import { CsvViewer } from "@arcships/vue-ui"
import type {
  CsvViewerProps,
  CsvViewerState,
  CsvViewerLoadEvent,
  CsvViewerDownloadEvent,
  CsvViewerError,
} from "@arcships/vue-ui"
```

| Prop | 类型 | 默认 | 说明 |
|---|---|---|---|
| `data` | `string` | — | 分隔文本；优先级高于 `file`/`src` |
| `file` | `Blob \| null` | — | 初始文件；按 BOM 解码 UTF-8 / UTF-16LE / UTF-16BE |
| `src` | `string` | — | 远程 URL；`fetch` + `arrayBuffer()`，失败进入 error 状态 |
| `fileName` | `string` | — | 下载名与 `getState().fileName`；`.tsv` 缺省分隔符为 tab |
| `delimiter` | `string` | `"auto"` | `auto` 或单字符；`tab` / `\\t` 是 tab 别名 |
| `search` | `boolean` | `true` | 是否显示搜索浮层；上游默认 `false`，本 block 默认开启 |
| `searchDebounce` | `number` | `300` | 输入防抖毫秒数 |
| `height` | `number \| string` | CSS `560px` | 数字按 px；字符串按 CSS length；不传时宿主 class 可覆盖 |
| `maxFileSize` | `number` | 50 MiB | 上传上限；超过时 `FILE_TOO_LARGE` |
| `loading` | `boolean` | `false` | 宿主异步加载时强制 loading 状态 |
| `locale` / `messages` | `OfficeLocale` / `OfficeMessagesOverride` | 继承 | 共享 i18n catalog |
| `class` | `string` | — | 合并到根节点 |

事件：`load`（解析成功后一次）、`error`（`CsvViewerError`）、`download`（`{ fileName, text }`，在浏览器保存前发出）、`update:data`（上传/URL 读取后的原始文本）。

实例方法：`reload()`、`retry()`、`download()`、`openFilePicker()`、`runSearch(query)`、`clearSearch()`、`nextResult()`、`previousResult()`、`setZoom(percent)`、`zoomIn()`、`zoomOut()`、`selectCell(row, col)`、`getState()`。

纯函数（根入口）：`parseCsvText`、`decodeCsvBuffer`、`serializeCsv`、`detectCsvDelimiter`、`normalizeCsvDelimiter`、`findCsvMatches`、`resolveCsvVisibleRange`、`csvCellLetters`、`csvCellAddress`。

## 5. 与上游的精确视觉差异

上游是 Glide canvas，本 block 是 DOM table。**没有像素对照，也不声称 pixel parity**。差异清单：

| 区域 | 上游 | 本 block | 影响 |
|---|---|---|---|
| 网格 | `@glideapps/glide-data-grid` canvas；`rowMarkers="number"`、`rowHeight=34`、`headerHeight=36`、首列 180/其余 160 | 语义 `<table>`；同样的行高/列宽/缩放公式；行号列宽 44 | 单元格文字、焦点环、选择高亮由 CSS 绘制；抗锯齿和 canvas 基线无法逐像素比较 |
| 选择 | Glide `GridSelection`，多行选择 | 单单元格选择 + 搜索 active cell | 没有多行/范围选择、没有 Glide 的拖拽选区 |
| 搜索高亮 | 只选中命中单元格 | 选中 + `<mark>` 高亮所有可见命中 | 更易读，但视觉不同 |
| 空/错误/加载 | 自定义 `max-w-md` 卡片 | `OfficeEmptyState` / `OfficeErrorState` / `OfficeLoadingState` | 共享 chrome 的尺寸/间距，不是 Glide 卡片 |
| 工具栏 | `min-h-12`、右对齐；zoom + search + ⋯ | `OfficeToolbar` + 左侧行/列统计；zoom + search + ⋯ | 多一个 leading 统计文本 |
| 主题 | 手写 Glide `Theme` 十六进制值 | `--office-*` token（light/dark/system） | token 值来自 0.8 设计系统；不是 Glide theme 的逐值复制 |
| 滚动 | Glide 内部虚拟化 | DOM 行窗口 + spacer row | 行窗口大小/overscan 不同，滚动条总高度一致 |
| 下载 | Papa.unparse | `serializeCsv`（RFC 4180） | 分隔符/换行可配置；默认 `\r\n` |

结构/行为测试覆盖这些合同：`tests/component/csv-viewer.test.mjs`。没有 `tools/parity-fixtures` 的 `csv-viewer` surface，因为 canvas 与 DOM 的像素比较没有意义；L1 仍由共享 chrome 的既有 surface 覆盖。

## 6. 后果

- **正面**：CSV 有独立、可发现、可测试的 API；不引入 Glide/papaparse；上传、搜索、缩放、下载与错误状态都有真实契约；`XlsxViewer` 的 CSV 兼容路径不受影响。
- **负面**：没有 Glide 的多行选择、列宽拖拽、冻结列、单元格编辑；超宽 CSV 只做行虚拟化，不做列虚拟化。
- **后续**：若需要大文件分块解析，可在 `parseCsvText` 之外增加 Worker；若需要像素对照，需要先决定是否接受 Glide alpha 依赖。
- **兼容性**：全部为新增导出；`XlsxViewer` 的 CSV 行为不变，`docs/guide/xlsx.md` 增加只读/格式丢失/编码限制。
