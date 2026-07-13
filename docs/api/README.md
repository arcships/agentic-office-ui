# API 导航

API 分为三层：

| 层级 | 适用情况 | 入口 |
|---|---|---|
| 高层组件 | 直接嵌入业务页面 | [组件手册](../components/README.md) |
| 组合函数与 Runtime | 自定义界面、状态和生命周期 | 各格式指南 |
| 纯数据核心 | 不依赖 Vue 的模型、布局和命令 | 对应 `*-core` 包 |

需要直接使用组合函数时，从[组合函数手册](../composables/README.md)开始；需要自己搭建界面时查看[自定义组件指南](../custom-components/README.md)。

## DOCX

- `@arcships/docx-core/core`：文档模型、布局和纯数据编辑命令。
- `@arcships/docx-core/runtime`：`createDocxRuntime`、来源、限制、诊断和加载错误。
- `@arcships/vue-docx`：`useDocxEditor`、`useDocxModel`、页面缩略图、分页、批注、修订和样式组合函数。
- [DOCX 使用指南](../guide/docx.md)

## XLSX

- `@arcships/xlsx-core/core`：地址、颜色、图表公式、图片锚点和工作簿类型。
- `@arcships/xlsx-core/runtime`：`createXlsxRuntime`、Worker、来源、限制和诊断。
- `@arcships/vue-xlsx`：`useXlsxViewerController` 和 `useXlsxViewerThumbnails`。
- [XLSX 使用指南](../guide/xlsx.md)

## PDF

- `@arcships/vue-pdf`：`PdfViewer`、`createPdfRenderRuntime`、来源校验、错误和诊断类型。
- [PDF 使用指南](../guide/pdf.md)

## PPTX

- `@arcships/pptx-core`：平台无关类型、能力报告、时间安排和属性轨道。
- `@arcships/pptx-core/browser`：纵向列表/单页预览会话、文档会话和播放控制器。
- `@arcships/vue-pptx`：`PptxViewer`、`PptxStage`、`usePptxDocument`、`usePptxPlayback` 以及 Surface 交互事件类型。
- [PPTX 使用指南](../guide/pptx.md)

## 接口稳定性

[公开接口合同](public-api-contract.md)登记精确导出路径、错误码、弃用期限和发布检查。它用于兼容性审查，不代替使用指南。
