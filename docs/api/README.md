# API 导航

API 分为三层：

| 层级 | 适用情况 | 入口 |
|---|---|---|
| 高层组件 | 直接嵌入业务页面 | [组件手册](../components/README.md) |
| 组合函数与 Runtime | 自定义界面、状态和生命周期 | 各格式指南 |
| 纯数据核心 | 不依赖 Vue 的模型、布局和命令 | 对应 `*-core` 包 |

## 对象引用与选择事件

- `@arcships/office-interaction`：跨格式引用与确认/取消事件、运行时校验、稳定序列化、候选导航、键盘解释器和临时选择 reducer。
- `@arcships/vue-ui`：`OfficeObjectOutlineLayer` 和 `OfficeRegionSelector` 两个受控选择原语。
- 四个格式包均已提供首批引用适配器：DOCX/PDF 覆盖文字、页面与区域；XLSX 覆盖工作表、单元格、范围、行列、图表与区域；PPTX 覆盖幻灯片、精确文字、可见对象、组合层级与区域。
- `DocxDocumentSurface`、`XlsxSheetSurface`、`PptxStage` 与 `PdfSurface` 统一接受 `documentId`、受控 `selectionMode` 和候选事件开关，并发出确认、取消、区域草稿、revision、解析与错误事件。
- Surface 不保存确认后的引用集合，不提供提示词、工具栏或 Agent 调用；宿主只需消费同一种引用和事件合同即可自行实现产品交互。

需要直接使用组合函数时，从[组合函数手册](../composables/README.md)开始；需要自己搭建界面时查看[自定义组件指南](../custom-components/README.md)。

## DOCX

- `@arcships/docx-core/core`：文档模型、布局和纯数据编辑命令。
- `@arcships/docx-core/runtime`：`createDocxRuntime`、来源、限制、诊断和加载错误。
- `@arcships/docx-core`：`createDocxTextReferenceDraft`、页面/区域草稿与 `resolveDocxReference`。
- `@arcships/vue-docx`：`useDocxEditor`、`useDocxModel`、页面缩略图、分页、批注、修订和样式组合函数。
- [DOCX 使用指南](../guide/docx.md)

## XLSX

- `@arcships/xlsx-core/core`：地址、颜色、图表公式、图片锚点和工作簿类型。
- `@arcships/xlsx-core/runtime`：`createXlsxRuntime`、Worker、来源、限制和诊断。
- `@arcships/xlsx-core`：工作表、单元格、范围、图表、人工区域引用草稿，以及 `resolveXlsxReference`。
- `@arcships/vue-xlsx`：`useXlsxViewerController` 和 `useXlsxViewerThumbnails`。
- [XLSX 使用指南](../guide/xlsx.md)

## PDF

- `@arcships/vue-pdf`：`PdfViewer`、`createPdfRenderRuntime`、来源校验、错误和诊断类型，以及 `PageTextSlice` 基础引用适配器。
- [PDF 使用指南](../guide/pdf.md)

## PPTX

- `@arcships/pptx-core`：平台无关类型、对象身份、引用适配器、能力报告、时间安排和属性轨道。
- `@arcships/pptx-core/browser`：纵向列表/单页预览会话、文档会话和播放控制器。
- `@arcships/vue-pptx`：`PptxViewer`、`PptxStage`、`usePptxDocument`、`usePptxPlayback` 以及 Surface 交互事件类型。
- [PPTX 使用指南](../guide/pptx.md)

## 接口稳定性

[公开接口合同](public-api-contract.md)登记精确导出路径、错误码、弃用期限和发布检查。它用于兼容性审查，不代替使用指南。
