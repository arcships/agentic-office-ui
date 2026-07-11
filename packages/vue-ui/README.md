# @arcships/vue-ui

Vue 3 UI 组件库：签名板、文件上传、缩略图、引用框、版面块和基础 UI。

## 安装

```bash
pnpm add @arcships/vue-ui
```

## 使用

```vue
<script setup lang="ts">
import { SignaturePad, FileUpload, Spinner } from "@arcships/vue-ui"
import "@arcships/vue-ui/style.css"
</script>

<template>
  <FileUpload accept=".pdf,.docx" :max-size="10485760" />
  <SignaturePad :width="400" :height="200" />
  <Spinner size="md" />
</template>
```

## 组件

| 组件 | 用途 |
|---|---|
| `SignaturePad` | 签名板 |
| `FileUpload` | 文件上传（含校验） |
| `FileThumbnail` | 文件缩略图 |
| `BoundingBoxCitations` | 引用框标注 |
| `LayoutBlocks` | 版面块 |
| `Spinner` | 加载指示 |
| `Tooltip` | 提示 |

## 事件

`FileUpload` 发出 `files-accepted`（通过校验的 `File[]`）和 `files-rejected`（`FileUploadRejection[]`，含 `code`、`file`、`message`）。`SignaturePad` 发出 `update:signature`。

## 文档

- [公开接口合同](https://github.com/eric8810/agentic-office-ui/blob/master/docs/api/public-api-contract.md)

## License

MIT
