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

`PdfSurface` 提供最小嵌入渲染：全部页面垂直堆叠滚动，无工具栏/缩略图/搜索栏。宿主可通过 `v-model:zoom` 接入 50%–200% 手势缩放，也可继续使用 expose zoom 和 scrollToPage。

## 文字、页面与区域引用

根入口可把 `PdfSurface` 已有的 `PageTextSlice + text + rects` 转换为统一引用草稿，并复用 PDF Runtime 的文字切片和搜索能力解析旧引用：

```ts
import {
  createPdfTextReferenceDraft,
  resolvePdfReference,
} from "@arcships/vue-pdf"

const context = { revision, document: renderDocument, runtime }
const draft = createPdfTextReferenceDraft(context, selection)
const result = await resolvePdfReference(context, { ...draft, referenceId })
```

字符范围在同修订内精确验证；修订变化后，唯一文字匹配返回 `relocated`，重复文字返回 `ambiguous`，删除返回 `not-found`。页面和人工区域只有在页序与尺寸/旋转签名未变时才迁移。

`PdfSurface` 接受受控的 `selectionMode="content | object | region"`：文字选择、页面点击和区域拖选都会发出统一 `referenceConfirm`，同时保留原始 `selectionChange`。宿主可通过 expose 的 `resolveReference()` 在文件更新后重新定位；确认后的引用集合和 Agent 工作流不由组件保存。

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
