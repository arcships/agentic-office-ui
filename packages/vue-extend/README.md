# @arcships/vue-extend

Vue 3 PDF 查看器和通用组件（签名、上传、缩略图、引用框、版面）。

## 安装

```bash
pnpm add @arcships/vue-extend
```

## PDF 查看器

```vue
<script setup lang="ts">
import { PdfViewer } from "@arcships/vue-extend"
import "@arcships/vue-extend/style.css"
</script>

<template>
  <PdfViewer :src="pdfUrl" />
</template>
```

正常 PDF 支持查看、翻页、缩放、旋转、缩略图、搜索和下载。PDF 唯一的公开文件拒绝上限是整份文件体积 `maxFileSize`，默认 `50 MiB`，宿主可调整。

## 组件

| 组件 | 用途 |
|---|---|
| `PdfViewer` | PDF 查看 |
| `SignaturePad` | 签名板 |
| `FileUpload` | 文件上传（含校验） |
| `FileThumbnail` | 文件缩略图 |
| `BoundingBoxCitations` | 引用框标注 |
| `LayoutBlocks` | 版面块 |
| `Spinner` | 加载指示 |
| `Tooltip` | 提示 |

## 事件

`PdfViewer` 发出 `document-load-success`、`document-load-error`、`active-page-change` 和 `diagnostic`；`FileUpload` 发出 `files-accepted` 和 `files-rejected`。

## 文档

- [公开接口合同](https://github.com/eric8810/agentic-office-ui/blob/master/docs/api/public-api-contract.md)

## License

MIT
