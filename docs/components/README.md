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
| `DocxBlockHighlights` / `DocxWordHighlights` | `@arcships/vue-docx` | 在已渲染页面上叠加文档块/词高亮，支持 `pageIndex`、`blockIds` 和 `wordRanges` | [DOCX](../guide/docx.md) |

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

`@arcships/vue-ui` 提供与格式无关的受控选择与标注原语：

| 组件 | 输入 | 输出 |
|---|---|---|
| `OfficeObjectOutlineLayer` | 当前容器内的 `OfficeObjectOutline[]` 和活动候选 id | `activate`、`confirm`、`dismiss`、`navigate-hierarchy` |
| `OfficeRegionSelector` | 规范化 `modelValue` 区域 | `update:modelValue`、`selection-start/change/commit/cancel` |
| `OfficeRegionAnnotator` | 截图 URL/`Blob` 与 `OfficeRegionAnnotationDocument` | `update:modelValue`、`annotation-add/remove/select/change/error` |

轮廓和区域坐标均为当前页面、幻灯片或工作表容器内的 `0..1` 规范坐标。组件不读取 Office 文件、不生成格式定位器、不保存引用集合，也不决定确认后的界面或动作；宿主可以使用 `@arcships/office-interaction` 的临时 reducer，也可以自行处理事件。四种格式 Surface 已直接提供同合同的候选、确认和区域事件；这两个原语用于宿主需要自定义轮廓或区域交互时按需组合，不要求 Surface 固定挂载某种产品 UI。

轮廓层支持 `Tab`、`Shift+Tab`、上下方向键、`Enter`、`Escape` 和左右层级导航事件；点击时 `Shift` 会在 `confirm` 参数中设置 `additiveRequested`，`Alt/Option` 设置 `penetrateRequested`。区域框支持拖选、方向键移动、`Shift + 方向键` 调整大小、`Enter` 确认和 `Escape` 取消。

### `OfficeRegionAnnotator`

在 `captureRegion()` 返回的位图上提供箭头、矩形、文字和高亮工具。`modelValue` 是 `OfficeRegionAnnotationDocument`，坐标相对截图归一化；`toJSON()`/`loadJSON()` 可与 `stringifyOfficeRegionAnnotationDocument()`/`parseOfficeRegionAnnotationDocument()` 往返。`image` 传 `Blob` 时组件会创建并在替换/卸载时撤销 object URL。详细属性、限制和示例见[区域截图与标注](../api/region-capture-and-annotation.md)。

### 区域截图

`captureRegion({ surface, region, scale })` 接受元素、capture provider 或内置 Surface 组件实例。Canvas/图片 Surface 提供 intrinsic `raster` 时按像素精确裁剪；纯 DOM Surface 使用 `foreignObject` 尽力栅格化。返回对象包含 `blob`、`url`、`source` 元数据和幂等 `release()`，宿主必须在替换或卸载时撤销 URL。完整合同见[区域截图与标注](../api/region-capture-and-annotation.md)。

### `FileUpload`

属性：`accept`、`multiple`、`maxSize`、`disabled`、`className`，以及 0.8 新增的可选视觉/交互属性：`title`、`description`、`browseLabel`、`draggingLabel`、`acceptedFileTypes`、`showFileList`、`showBorderBeam`、`borderBeamTheme`。事件保持 `files-accepted` 和 `files-rejected`；拒绝结果包含稳定的 `code`、原始 `file` 和可显示的 `message`。默认 `showFileList=false` 以避免与宿主自己的文件列表重复；需要 Extend 内置列表时显式开启。完整合同、像素对照和状态说明见[File components](file-components.md)。

### `SignaturePad`

属性：`width`、`height`、`penColor`、`backgroundColor`、`className`。通过 `update:signature` 返回数据地址或 `null`。模板引用公开 `clear()`、`isEmpty()` 和 `toDataURL()`。

### `FileThumbnail`

接收 `file: { name, type, url?, size? }`、`size`（`sm`/`md`/`lg`，或像素/CSS 长度）、`previewImageUrl`、`previewAspectRatio`、`isLoading`、`hasError`、`showTypeGlyph`、`showMeta`、`alt` 和 `className`。方形/自定义比例预览卡包含图片 blur-in、loading overlay、error/empty fallback；需要旧版文件名/类型行时设置 `showMeta`。完整合同见[File components](file-components.md)。

### `BoundingBoxCitations`

接收原始文件地址和 `BoundingBoxField[]`，通过 `field-click` 返回被点击字段。矩形格式为 `[x, y, width, height]`，页码由字段的 `page` 给出。

### `LayoutBlocks`

接收文件地址和 OCR 版面结果，通过 `block-click` 返回被点击的 `LayoutBlock`。

### `Spinner`

`size` 支持 `sm`、`md` 和 `lg`（`md` 等价于 primitive 的 `default`），没有业务事件。内部委托 `OfficeSpinner`；新代码可以直接使用后者，并传 `label` 控制可访问名称。

### `Tooltip`

属性包括 `content`、`side`、`align`、`delayMs` 和 `className`。默认插槽放置触发元素。内部委托 `OfficeTooltip`（reka-ui），因此具备碰撞定位、焦点触发的 ARIA 关联和 `delayDuration` 行为；样式和公开 props 保持向后兼容。

## Token 驱动 primitive（0.8）

`@arcships/vue-ui` 在 0.8 提供基于 reka-ui 的完整 primitive 层，视觉全部读取 `--office-*` token，不引入 Tailwind/preflight：

| Primitive | 用途 |
|---|---|
| `OfficeButton` / `OfficeIconButton` | 六种 variant、七种 size、disabled/loading/asChild；图标按钮强制 `label` |
| `OfficeInput` / `OfficeTextarea` | 受控输入、invalid、disabled、readonly、三种尺寸 |
| `OfficeSelect` | 选项/分组/分隔符数组、portal、键盘选择和碰撞定位 |
| `OfficeDropdownMenu` / `OfficeContextMenu` | 共用 `OfficeMenuEntry` 描述符；item/checkbox/separator/label/group/submenu |
| `OfficeDialog` / `OfficeSheet` | v-model:open、焦点陷阱、Esc/backdrop、标题与描述 ARIA、footer、四向 Sheet |
| `OfficePopover` / `OfficeTooltip` | 可配置 side/align/offset 的 portal 浮层；Tooltip 支持延迟与 disabled |
| `OfficeTabs` | default/underline、横向/纵向、roving tabindex 和动态面板 slot |
| `OfficeToggle` | 受控按下态、`aria-pressed`、outline variant |
| `OfficeBadge` / `OfficeCard` | variant/size；卡片标题、描述、header/body/footer |
| `OfficeSeparator` / `OfficeSkeleton` / `OfficeSpinner` | 装饰分隔符、reduced-motion 骨架、状态 spinner |
| `OfficeScrollArea` | Reka 自定义滚动条，viewport 可聚焦 |
| `OfficeToastProvider` + `useOfficeToast()` | success/error/warning/info/loading、promise、更新与关闭 |

完整 props、事件、slot、键盘/ARIA 合同、token 规则和包体积预算见
[Vue primitive 层](ui-primitives.md)。示例页在 demo 的 `/primitives` 路由，展示每个 primitive 的
default/hover/focus/disabled/loading 状态。

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
| `link-activate` | `{ slideIndex, objectKey }`；链接形状获得焦点后按 `Enter`/`Space` |

链接形状会获得 `title`（`tooltip`，缺省时回退到 URL）、`data-pptx-hyperlink*` 标记和（点击链接的）`role="link"` + `tabindex="0"`；`PptxViewer` 默认使用这些标记完成点击/键盘打开，自定义界面可以只监听 `link-activate` 并自行处理。

页面编号从零开始。`objectKey` 是渲染会话生成的稳定对象标识；宿主不应自行解析其字符串结构。
