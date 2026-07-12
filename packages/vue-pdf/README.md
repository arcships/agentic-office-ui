# @arcships/vue-pdf

Vue 3 PDF 查看器组件，基于 PDFium 渲染。

## 安装

```bash
pnpm add @arcships/vue-pdf
```

## 使用

```vue
<script setup lang="ts">
import { PdfViewer } from "@arcships/vue-pdf"
import "@arcships/vue-pdf/style.css"
</script>

<template>
  <PdfViewer :src="pdfUrl" />
</template>
```

正常 PDF 支持查看、翻页、缩放、旋转、缩略图、搜索和下载。PDF 唯一的公开文件拒绝上限是整份文件体积 `maxFileSize`，默认 `50 MiB`，宿主可调整。

## 事件

| 事件 | 参数 |
|---|---|
| `document-load-success` | 页数 |
| `document-load-error` | `PdfLoadError` |
| `active-page-change` | 当前页码，从 1 开始 |
| `diagnostic` | `PdfDiagnostic` |

## Worker 和 WASM

```ts
import { createPdfRenderRuntime, bundledPdfiumWasmUrl } from "@arcships/vue-pdf"

const runtime = createPdfRenderRuntime({ pdfiumWasmUrl: bundledPdfiumWasmUrl })
```

## 文档

- [PDF 使用指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/guide/pdf.md)
- [组件手册](https://github.com/arcships/agentic-office-ui/blob/master/docs/components/README.md)
- [自定义 PDF 界面](https://github.com/arcships/agentic-office-ui/blob/master/docs/custom-components/pdf.md)
- [公开接口合同](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/public-api-contract.md)

## License

Apache-2.0
