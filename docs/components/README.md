# Vue 组件手册

本页列出推荐给新项目使用的高层组件。已经标记弃用的低层组件不在这里展开。

如果现成组件布局不满足需求，请查看[自定义组件指南](../custom-components/README.md)和[组合函数手册](../composables/README.md)。

## Office 组件

| 组件 | 包 | 用途 | 详细指南 |
|---|---|---|---|
| `DocxViewer` | `@arcships/vue-docx` | DOCX 查看、搜索、缩略图、批注和修订显示 | [DOCX](../guide/docx.md) |
| `DocxEditor` | `@arcships/vue-docx` | DOCX 编辑、撤销重做和导入导出 | [DOCX](../guide/docx.md) |
| `XlsxViewer` | `@arcships/vue-xlsx` | 工作表、图表、图片、公式和编辑界面 | [XLSX](../guide/xlsx.md) |
| `PdfViewer` | `@arcships/vue-pdf` | PDFium 查看、搜索、缩略图、旋转和下载 | [PDF](../guide/pdf.md) |
| `PptxViewer` | `@arcships/vue-pptx` | 纵向连续浏览 PPTX；演示模式使用单页播放 | [PPTX](../guide/pptx.md) |
| `PptxStage` | `@arcships/vue-pptx` | 无 chrome 的纵向 Surface 或单页播放舞台；提供选中、对象点击和右键事件 | [PPTX](../guide/pptx.md) |
| `PptxThumbnail` | `@arcships/vue-pptx` | 使用现有会话渲染单页缩略图 | [PPTX](../guide/pptx.md) |

## 基础组件

### `FileUpload`

属性：`accept`、`multiple`、`maxSize`、`disabled`、`className`。事件：`files-accepted` 和 `files-rejected`。拒绝结果包含稳定的 `code`、原始 `file` 和可显示的 `message`。

### `SignaturePad`

属性：`width`、`height`、`penColor`、`backgroundColor`、`className`。通过 `update:signature` 返回数据地址或 `null`。模板引用公开 `clear()`、`isEmpty()` 和 `toDataURL()`。

### `FileThumbnail`

接收 `file: { name, type, url? }` 和 `size: "sm" | "md" | "lg"`，用于显示文件类型和预览信息。

### `BoundingBoxCitations`

接收原始文件地址和 `BoundingBoxField[]`，通过 `field-click` 返回被点击字段。矩形格式为 `[x, y, width, height]`，页码由字段的 `page` 给出。

### `LayoutBlocks`

接收文件地址和 OCR 版面结果，通过 `block-click` 返回被点击的 `LayoutBlock`。

### `Spinner`

`size` 支持 `sm`、`md` 和 `lg`，没有业务事件。

### `Tooltip`

属性包括 `content`、`side`、`align` 和 `delayMs`。默认插槽放置触发元素。

## 通用约定

- Vue 属性在模板中使用短横线名称，例如 `showToolbar` 写作 `show-toolbar`。
- 必须引入对应包的 `style.css`。
- 错误判断使用公开 `code`，不要依赖中文或英文错误文案。
- 标记为 `@deprecated` 的低层组件在整个 `0.x` 保留，但不建议新项目继续使用。

## PPTX Surface 事件

`PptxStage` 的事件均使用 Vue 模板短横线名称：

| 事件 | 参数 |
|---|---|
| `selection-change` | `{ kind: "slide", slideIndex }` |
| `object-click` | `{ kind: "object", slideIndex, objectKey }` |
| `context-menu` | `kind: "slide"` 时禁止 `objectKey`；`kind: "object"` 时必须包含 `objectKey`；两者都含 `slideIndex`、`clientX/Y`、`containerX/Y` |

页面编号从零开始。`objectKey` 是渲染会话生成的稳定对象标识；宿主不应自行解析其字符串结构。
