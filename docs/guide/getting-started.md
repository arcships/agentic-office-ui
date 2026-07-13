# 开始使用 Agentic Office UI

Agentic Office UI 是一组 Vue 3 Office 文档组件。可以直接使用完整查看器，也可以使用核心包和组合函数搭建自己的界面。

## 选择需要的包

| 需求 | 安装包 | 推荐入口 |
|---|---|---|
| 查看 DOCX | `@arcships/vue-docx`、`@arcships/docx-core` | `DocxViewer` |
| 编辑 DOCX | `@arcships/vue-docx`、`@arcships/docx-core` | `DocxEditor` |
| 查看和操作 XLSX | `@arcships/vue-xlsx`、`@arcships/xlsx-core` | `XlsxViewer`、`useXlsxViewerController` |
| 查看 PDF | `@arcships/vue-pdf` | `PdfViewer` |
| 查看和播放 PPTX | `@arcships/vue-pptx`、`@arcships/pptx-core` | `PptxViewer` |
| 自定义 PPTX 纵向浏览 | `@arcships/vue-pptx`、`@arcships/pptx-core` | `PptxStage` + `usePptxDocument` |
| 自定义 PPTX 播放界面 | `@arcships/vue-pptx`、`@arcships/pptx-core` | `PptxStage` 和两个组合函数 |
| 文件上传、签名等基础界面 | `@arcships/vue-ui` | 对应组件 |

## 安装

按需安装，不要求一次安装全部包：

```bash
pnpm add @arcships/vue-docx @arcships/docx-core
pnpm add @arcships/vue-xlsx @arcships/xlsx-core
pnpm add @arcships/vue-pdf
pnpm add @arcships/vue-pptx @arcships/pptx-core
pnpm add @arcships/vue-ui
```

所有 Vue 包都需要 Vue 3。当前支持范围为 `>=3.2.25 <4`。

## 引入样式

使用哪个 Vue 包，就引入哪个包的样式：

```ts
import "@arcships/vue-docx/style.css"
import "@arcships/vue-xlsx/style.css"
import "@arcships/vue-pdf/style.css"
import "@arcships/vue-pptx/style.css"
import "@arcships/vue-ui/style.css"
```

## 最小示例

```vue
<script setup lang="ts">
import { ref } from "vue"
import { DocxViewer } from "@arcships/vue-docx"
import "@arcships/vue-docx/style.css"

const bytes = ref<ArrayBuffer>()

async function onFileChange(event: Event) {
  const file = (event.target as HTMLInputElement).files?.[0]
  bytes.value = file ? await file.arrayBuffer() : undefined
}
</script>

<template>
  <input
    type="file"
    accept=".docx"
    @change="onFileChange"
  />
  <DocxViewer :file="bytes" />
</template>
```

## 输入、资源和生命周期

- DOCX、XLSX、PDF 和 PPTX 都支持本地文件或字节输入；各组件的具体来源类型见对应指南。
- Worker、WASM 和 PPTX 渲染代码都从已安装的 npm 包加载，不要从仓库 demo 目录复制资源。
- 高层组件会清理自己创建的 Worker、媒体和对象地址。调用方注入的 Runtime 仍由调用方销毁。
- 快速切换文件时只保留最新结果。业务代码应显示公开错误，不要解析界面文字判断失败原因。

## 下一步

- [DOCX 使用指南](docx.md)
- [XLSX 使用指南](xlsx.md)
- [PDF 使用指南](pdf.md)
- [PPTX 使用指南](pptx.md)
- [组件手册](../components/README.md)
- [API 导航](../api/README.md)
- [组合函数手册](../composables/README.md)
- [自定义组件指南](../custom-components/README.md)
