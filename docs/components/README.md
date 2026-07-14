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

## 最小 Surface

| 组件 | 包 | 无界面能力 |
|---|---|---|
| `DocxDocumentSurface` | `@arcships/vue-docx` | DOCX 分页文档面、批注 gutter、搜索、选择与引用事件 |
| `XlsxSheetSurface` | `@arcships/vue-xlsx` | XLSX 网格、对象层、搜索、选择与引用事件 |
| `PdfSurface` | `@arcships/vue-pdf` | PDF 连续页面、搜索、glyph 文字选择与引用事件 |
| `PptxStage` | `@arcships/vue-pptx` | PPTX 纵向或单页舞台、搜索、对象交互与引用事件 |

四个 Surface 统一接受 `documentId`、受控 `selectionMode` 和候选事件开关。具体接入见[引用选择接入指南](../guide/reference-selection.md)。

## 基础组件

### Office 对象选择控件

`@arcships/vue-ui` 提供两个与格式无关的受控选择原语：

| 组件 | 输入 | 输出 |
|---|---|---|
| `OfficeObjectOutlineLayer` | 当前容器内的 `OfficeObjectOutline[]` 和活动候选 id | `activate`、`confirm`、`dismiss`、`navigate-hierarchy` |
| `OfficeRegionSelector` | 规范化 `modelValue` 区域 | `update:modelValue`、`selection-start/change/commit/cancel` |

轮廓和区域坐标均为当前页面、幻灯片或工作表容器内的 `0..1` 规范坐标。组件不读取 Office 文件、不生成格式定位器、不保存引用集合，也不决定确认后的界面或动作；宿主可以使用 `@arcships/office-interaction` 的临时 reducer，也可以自行处理事件。四种格式 Surface 已直接提供同合同的候选、确认和区域事件；这两个原语用于宿主需要自定义轮廓或区域交互时按需组合，不要求 Surface 固定挂载某种产品 UI。

轮廓层支持 `Tab`、`Shift+Tab`、上下方向键、`Enter`、`Escape` 和左右层级导航事件；点击时 `Shift` 会在 `confirm` 参数中设置 `additiveRequested`，`Alt/Option` 设置 `penetrateRequested`。区域框支持拖选、方向键移动、`Shift + 方向键` 调整大小、`Enter` 确认和 `Escape` 取消。

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
