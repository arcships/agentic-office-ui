<div align="center">

<img src="docs/assets/banner.png" alt="Agentic Office UI：Vue 3 Office 文档组件库" width="100%">

# Agentic Office UI

在浏览器中查看和编辑 DOCX、XLSX、PDF 与 PPTX

[![npm version](https://img.shields.io/badge/npm-v0.5.2-cb3837)](https://www.npmjs.com/org/arcships)
[![CI](https://github.com/arcships/agentic-office-ui/actions/workflows/ci.yml/badge.svg)](https://github.com/arcships/agentic-office-ui/actions/workflows/ci.yml)
[![Release](https://github.com/arcships/agentic-office-ui/actions/workflows/release.yml/badge.svg)](https://github.com/arcships/agentic-office-ui/actions/workflows/release.yml)
[![Vue](https://img.shields.io/badge/Vue-%3E%3D3.2.25%20%3C4-42b883)](https://vuejs.org/)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](LICENSE)

[开始使用](docs/guide/getting-started.md) · [组件手册](docs/components/README.md) · [组合函数](docs/composables/README.md) · [自定义组件](docs/custom-components/README.md) · [完整文档](docs/INDEX.md)

</div>

## 能做什么

| 格式 | 开箱即用 | 可组合能力 |
|---|---|---|
| DOCX | 查看、编辑、分页、缩略图、搜索、批注、修订、导入导出 | 文档模型、布局、编辑命令、Runtime、Worker/WASM |
| XLSX | 工作表、公式、图表、图片、选区、编辑、撤销重做 | 控制器、Runtime、按需图表/地图/WebGL |
| PDF | PDFium 渲染、翻页、缩放、旋转、缩略图、搜索、下载 | 来源规则、渲染 Runtime、诊断和错误 |
| PPTX | 纵向连续预览、缩略图、搜索、逐步动画、切换、媒体、全屏 | 列表/单页文档会话、Surface 事件、播放控制器、无界面组合函数 |

这个项目既提供完整 Vue 组件，也保留不依赖 Vue 的核心包。你可以从一个查看器开始，也可以接管状态、工具栏和播放界面。

## 30 秒开始

安装需要的格式：

```bash
pnpm add @arcships/vue-docx @arcships/docx-core
pnpm add @arcships/vue-xlsx @arcships/xlsx-core
pnpm add @arcships/vue-pdf
pnpm add @arcships/vue-pptx @arcships/pptx-core
```

使用组件并引入对应样式：

```vue
<script setup lang="ts">
import { PptxViewer } from "@arcships/vue-pptx"
import "@arcships/vue-pptx/style.css"

defineProps<{ file: File | null }>()
</script>

<template>
  <PptxViewer :source="file" mode="present" height="720px" />
</template>
```

其他格式的使用方式：

```vue
<DocxViewer :file="docxBytes" />
<XlsxViewer :controller="xlsxController" />
<PdfViewer :source="pdfFile" />
```

完整安装、来源类型和资源说明见[开始使用](docs/guide/getting-started.md)。

## 选择正确的入口

| 你想做什么 | 使用什么 |
|---|---|
| 尽快显示一个文档 | `DocxViewer`、`XlsxViewer`、`PdfViewer`、`PptxViewer` |
| 自建工具栏嵌入渲染区 | `DocxDocumentSurface`（仅分页渲染 + 批注 gutter） |
| 编辑 DOCX | `DocxEditor` 或 `useDocxEditor` |
| 自定义 XLSX 工具栏 | `useXlsxViewerController` + `XlsxSheetSurface`（无内置 chrome） |
| 自定义 XLSX 工具栏（全套） | `useXlsxViewerController` + `XlsxViewer` |
| 自建 PDF 控制栏嵌入渲染 | `PdfSurface`（垂直滚动全部页面） |
| 自建 PPTX 浏览 Surface | `PptxStage` + `usePptxDocument({ session: { renderMode: "list" } })`（垂直滚动全部幻灯片） |
| 自定义 PPTX 播放器 | `usePptxDocument` + `usePptxPlayback` + `PptxStage` |
| 在非 Vue 代码中处理文档 | `docx-core`、`xlsx-core`、`pptx-core` |
| 管理 Worker、WASM 和资源限制 | 对应格式的 Runtime |

## 公开包

| 包 | 说明 |
|---|---|
| [`@arcships/docx-core`](packages/docx-core/README.md) | DOCX 模型、布局、命令和 Runtime |
| [`@arcships/vue-docx`](packages/vue-docx/README.md) | `DocxDocumentSurface`、`DocxViewer`、`DocxEditor` 和组合函数 |
| [`@arcships/xlsx-core`](packages/xlsx-core/README.md) | XLSX 模型、公式、图表和 Runtime |
| [`@arcships/vue-xlsx`](packages/vue-xlsx/README.md) | `XlsxSheetSurface`、`XlsxViewer` 和查看器控制器 |
| [`@arcships/vue-pdf`](packages/vue-pdf/README.md) | PDFium 查看器与渲染 Runtime |
| [`@arcships/pptx-core`](packages/pptx-core/README.md) | PPTX 预览、播放模型和浏览器控制器 |
| [`@arcships/vue-pptx`](packages/vue-pptx/README.md) | 纵向 `PptxStage` Surface、`PptxViewer`、统一交互事件和最小播放组合能力 |
| [`@arcships/vue-ui`](packages/vue-ui/README.md) | 上传、签名、缩略图、引用和基础组件 |

八个包使用统一版本发布。相关 Vue 包和核心包建议保持相同版本。

PPTX 两包从 `0.3.0` 开始公开，`0.4.0` 增加最小组合接口，当前八个包的版本为 `0.5.2`。

## 文档

### 按格式

- [DOCX 使用指南](docs/guide/docx.md)
- [XLSX 使用指南](docs/guide/xlsx.md)
- [PDF 使用指南](docs/guide/pdf.md)
- [PPTX 使用指南](docs/guide/pptx.md)

### 按问题

- [组件属性、事件和选择建议](docs/components/README.md)
- [组合函数和状态所有权](docs/composables/README.md)
- [自定义工具栏与播放器](docs/custom-components/README.md)
- [组合函数、Runtime 和核心入口](docs/api/README.md)
- [公开入口、错误码和兼容期限](docs/api/public-api-contract.md)
- [版本升级说明](docs/migration-0.3.md)
- [浏览器兼容矩阵](docs/testing/compatibility-matrix.md)

## 设计原则

- 高层组件解决常见使用场景，核心包不依赖 Vue。
- Worker、WASM 和渲染资源随 npm 包发布，不要求复制仓库文件。
- 文件快速切换时只接受最新结果，组件卸载时释放自己持有的资源。
- 错误和能力报告提供稳定代码，业务不需要解析界面文案。
- 已公开的旧接口在整个 `0.x` 保持兼容，弃用项最早在 `1.0.0` 删除。

## 本地开发

```bash
pnpm install --frozen-lockfile
pnpm typecheck
pnpm build
pnpm test
```

PPTX 专项验证：

```bash
pnpm test:pptx
pnpm test:pptx:headless
```

## 浏览器要求

正式验证覆盖 Chromium、Firefox 和 WebKit。PDF、DOCX 和 XLSX 的部分能力依赖 Worker 与 WASM；严格内容安全策略需要允许对应的同源资源和 Worker。详细要求见各格式指南与[兼容矩阵](docs/testing/compatibility-matrix.md)。

PDF 支持翻页、缩放、搜索和下载。默认最大文件大小为 `50 MiB`，超过后返回稳定错误码 `PDF_TOO_LARGE`。

## License

[Apache 2.0](LICENSE)
