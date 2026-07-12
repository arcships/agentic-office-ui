# PDF 使用指南

```vue
<script setup lang="ts">
import { PdfViewer } from "@arcships/vue-pdf"
import "@arcships/vue-pdf/style.css"
</script>

<template>
  <PdfViewer
    :source="file"
    file-name="report.pdf"
    :show-toolbar="true"
    @document-load-success="onLoaded"
    @document-load-error="onError"
  />
</template>
```

`source` 是推荐输入；同时提供 `source` 和旧的 `src` 时，`source` 优先。URL 来源必须显式提供 `urlPolicy`，本地 `Blob`、`File` 和字节输入不需要网络规则。

| 属性 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `source` | `PdfSource` | 无 | 推荐的 PDF 来源 |
| `src` | `string` | 无 | 兼容 URL 输入 |
| `urlPolicy` | `PdfUrlPolicy` | 无 | URL 访问规则 |
| `runtime` | `PdfRenderRuntime` | 组件创建 | 调用方持有的 Runtime |
| `pdfiumWasmUrl` | `string` | 包内资源 | 覆盖 PDFium WASM 地址 |
| `maxFileSize` | `number` | 50 MiB | 最大 PDF 字节数 |
| `defaultZoom` | `number` | `1` | 初始缩放倍率，范围由组件约束 |
| `showToolbar` | `boolean` | `true` | 工具栏 |
| `showDownload` | `boolean` | `true` | 下载按钮 |
| `showRotateControls` | `boolean` | `true` | 旋转按钮 |

事件包括 `document-load-success`、`document-load-error`、`active-page-change` 和 `diagnostic`。错误处理使用 `PdfLoadError.code`，不要通过错误文案判断类型。

需要完全控制 Runtime 时，可以使用 `createPdfRenderRuntime()`。调用方传入的 Runtime 必须由调用方销毁。
