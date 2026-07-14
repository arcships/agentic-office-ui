<div align="center">

<img src="docs/assets/banner-agent-collaboration.png" alt="Agentic Office UI：让人看得见、指得准，让 Agent 找得到" width="100%">

# Agentic Office UI

**面向 Office Agent 人机协作的 Vue 3 组件库。**

让人完整看见文档、准确指出对象，并把 Agent 真正能定位的引用交给它。

[English](README.md) · 简体中文

[![npm stable](https://img.shields.io/npm/v/@arcships/vue-docx?label=npm%20stable&color=cb3837)](https://www.npmjs.com/org/arcships)
[![CI](https://github.com/arcships/agentic-office-ui/actions/workflows/ci.yml/badge.svg)](https://github.com/arcships/agentic-office-ui/actions/workflows/ci.yml)
[![Release](https://github.com/arcships/agentic-office-ui/actions/workflows/release.yml/badge.svg)](https://github.com/arcships/agentic-office-ui/actions/workflows/release.yml)
[![Vue](https://img.shields.io/badge/Vue-%3E%3D3.2.25%20%3C4-42b883)](https://vuejs.org/)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](LICENSE)
[![GitHub stars](https://img.shields.io/github/stars/arcships/agentic-office-ui?style=flat)](https://github.com/arcships/agentic-office-ui/stargazers)

[开始使用](docs/guide/getting-started.md) · [引用选择](docs/guide/reference-selection.md) · [组件手册](docs/components/README.md) · [API](docs/api/README.md) · [完整文档](docs/INDEX.md)

</div>

Agentic Office UI 是一个开源组件库：在浏览器中展示 DOCX、XLSX、PPTX 和 PDF，并把用户在画面上的选择转换成精确的 Office 引用。

它位于人与 Agent 之间。组件库负责预览和选择；宿主应用负责对话、Agent、文档修改工具和具体产品流程。

> **版本状态：** npm 当前稳定版为 `0.5.4`。仓库正在准备 `0.6.0`，新增跨格式引用合同、四格式阶段一定位器、统一 Surface 事件和区域选择原语。截图采集、标注编辑器与自动修改审阅尚未内置。

## 为什么需要这个库

Agent 通常通过代码修改 Office 文件；人类面对的却是段落、单元格、图表、幻灯片对象和页面区域。

这会反复产生两个问题：

1. Agent 工作时，人无法完整看见正在处理的文件，也难以核对具体改动。
2. “第二页右下角那张表”对人很自然，却不是 Agent 可以稳定执行的目标。

Agentic Office UI 用一条清晰的协作链路连接两者：

```mermaid
flowchart LR
  A["预览 Office 文件"] --> B["选择文字、单元格、对象或区域"]
  B --> C["生成精确 Office 引用"]
  C --> D["宿主把指令和引用交给 Agent"]
  D --> E["重新加载结果并再次解析引用"]
  E --> B
```

引用说明**用户指的是谁**，自然语言说明**用户想做什么**。

## 这个库提供什么

- **真实 Office Surface**：DOCX、XLSX、PPTX、PDF 的浏览器查看器，以及可嵌入的最小渲染面。
- **精确引用**：包含文档版本、语义定位器、内容证据、规范化区域和可靠度信息，不需要把整份 DOM 或 UI 状态塞给 Agent。
- **符合用户习惯的选择**：文字、表格范围、可见对象、行列、页面、幻灯片，以及无法识别对象时的区域兜底。
- **修改后的重新定位**：宿主重新加载 Agent 返回的文件后，可以描述、解析并滚动回原引用。
- **受控选择原语**：可选的对象轮廓和区域框；工具栏、引用列表、聊天框和命令界面仍由宿主自由设计。
- **与框架无关的核心**：共享合同和纯函数位于 `@arcships/office-interaction`，格式定位器和对应文档模型放在一起。

## 支持的格式

| 格式 | 查看器能力 | `0.6.0` 阶段一引用 |
|---|---|---|
| DOCX | 查看、编辑、分页、搜索、缩略图、批注、修订、导入导出 | 精确文字、页面、页面区域 |
| XLSX | 工作表、公式、图表、图片、范围选择、编辑、撤销重做 | 工作表、单元格、范围、行、列、图表、工作表区域 |
| PPTX | 连续预览、缩略图、搜索、对象事件、动画播放、切换、媒体、全屏 | 幻灯片、对象内精确文字、可见对象、组合层级、幻灯片区域 |
| PDF | PDFium 渲染、翻页、缩放、旋转、缩略图、搜索、文字选择、下载 | 精确字符范围、页面、页面区域 |

阶段一只选择那些可以稳定定位和校验的高频目标。段落对象、图表内部元素、更完整的文档图层，以及动画等不可见行为属于后续选择层，不是当前版本已经支持的能力。

## 快速开始

只安装需要的格式：

```bash
pnpm add @arcships/vue-docx @arcships/docx-core
pnpm add @arcships/vue-xlsx @arcships/xlsx-core
pnpm add @arcships/vue-pdf
pnpm add @arcships/vue-pptx @arcships/pptx-core
```

使用完整查看器，并引入公开样式：

```vue
<script setup lang="ts">
import { PptxViewer } from "@arcships/vue-pptx"
import "@arcships/vue-pptx/style.css"

defineProps<{ file: File | null }>()
</script>

<template>
  <PptxViewer :source="file" mode="browse" height="720px" />
</template>
```

其他格式对应的高层入口为 `DocxViewer`、`XlsxViewer` 和 `PdfViewer`。

## 把用户选择变成引用

下面的统一引用 API 属于 `0.6.0` 源码候选。正式发布后，直接导入共享类型或纯函数的应用，应把 `@arcships/office-interaction` 声明为直接依赖。

```vue
<script setup lang="ts">
import { shallowRef } from "vue"
import type { OfficeReferenceConfirmEvent } from "@arcships/office-interaction"
import { XlsxSheetSurface, useXlsxViewerController } from "@arcships/vue-xlsx"

const props = defineProps<{ file: ArrayBuffer }>()
const controller = useXlsxViewerController({
  file: props.file,
  fileName: "budget-2026.xlsx",
})
const selectedReference = shallowRef<OfficeReferenceConfirmEvent["reference"]>()

function onReferenceConfirm(event: OfficeReferenceConfirmEvent) {
  selectedReference.value = event.reference
}
</script>

<template>
  <XlsxSheetSurface
    :controller="controller"
    document-id="budget-2026.xlsx"
    selection-mode="content"
    @reference-confirm="onReferenceConfirm"
  />
</template>
```

Surface 只发出引用。把它放进聊天输入框、命令面板、多引用集合，还是结构化工具调用，由宿主应用自己决定。

[引用选择接入指南](docs/guide/reference-selection.md)完整说明了 `content`、`object`、`region` 三种模式、统一事件、expose 方法和状态归属。

## 选择合适的接入层级

| 目标 | 推荐入口 |
|---|---|
| 尽快显示一个文件 | `DocxViewer`、`XlsxViewer`、`PdfViewer`、`PptxViewer` |
| 围绕文档自建工具栏 | `DocxDocumentSurface`、`XlsxSheetSurface`、`PdfSurface`、`PptxStage` |
| 自建 DOCX 编辑器 | `DocxEditor` 或 `useDocxEditor` |
| 自建 PPTX 播放器 | `usePptxDocument` + `usePptxPlayback` + `PptxStage` |
| 在 Vue 之外处理文档模型 | `@arcships/docx-core`、`@arcships/xlsx-core`、`@arcships/pptx-core` |
| 使用共享引用类型和纯选择逻辑 | `@arcships/office-interaction` |
| 组合自己的选择界面 | `@arcships/vue-ui` 的 `OfficeObjectOutlineLayer`、`OfficeRegionSelector` |

## 包列表

| 包 | 用途 |
|---|---|
| [`@arcships/office-interaction`](packages/office-interaction/README.md) | 跨格式引用、运行时校验、候选导航、几何与临时选择状态 |
| [`@arcships/docx-core`](packages/docx-core/README.md) | DOCX 模型、布局、编辑命令、引用适配器和 Runtime |
| [`@arcships/vue-docx`](packages/vue-docx/README.md) | DOCX Surface、Viewer、Editor 和组合函数 |
| [`@arcships/xlsx-core`](packages/xlsx-core/README.md) | XLSX 模型、公式、图表、引用适配器和 Runtime |
| [`@arcships/vue-xlsx`](packages/vue-xlsx/README.md) | XLSX Surface、Viewer 和查看器控制器 |
| [`@arcships/vue-pdf`](packages/vue-pdf/README.md) | PDFium Surface、Viewer、引用适配器和渲染 Runtime |
| [`@arcships/pptx-core`](packages/pptx-core/README.md) | PPTX 预览/播放模型、对象身份、引用适配器和浏览器控制器 |
| [`@arcships/vue-pptx`](packages/vue-pptx/README.md) | PPTX Stage、Viewer、缩略图、Surface 事件和播放组合函数 |
| [`@arcships/vue-ui`](packages/vue-ui/README.md) | Office 选择原语，以及上传、签名、缩略图、引用框和版面组件 |

九个公开包使用同一发布版本线。npm 稳定版为 `0.5.4`，当前源码候选为 `0.6.0`。PPTX 两包从 `0.3.0` 开始公开，`0.4.0` 增加最小组合接口。

## 能力边界

这个项目**负责**：

- 用 Vue 组件在浏览器中预览 Office 文件；
- 发出能够定位到具体格式对象的选择与引用事件；
- 提供解析、导航、能力报告和结构化错误；
- 把需要的 Worker 和 WASM 资源随 npm 包一起发布。

这个项目**不负责**：

- 调用模型或指定 Agent 框架；
- 规定提示词、工具 schema、MCP resource 或 RPC 协议；
- 替宿主选择 `python-docx`、`openpyxl`、LibreOffice 或其他修改工具；
- 保存宿主确认后的引用集合或意图界面状态；
- 承诺完整复刻 Microsoft Office 的全部行为或像素效果；
- 在无法截图时伪造结果——`captureReferencePreview()` 会用 `CAPTURE_UNSUPPORTED` 明确拒绝。

PDF 对外公开的文件体积上限为 `50 MiB`；超限时返回结构化错误 `PDF_TOO_LARGE`。页数、总像素和内存不会被包装成隐藏的公共拒绝条件。

## 文档

- [开始使用](docs/guide/getting-started.md)
- [引用选择接入指南](docs/guide/reference-selection.md)
- [DOCX 指南](docs/guide/docx.md)
- [XLSX 指南](docs/guide/xlsx.md)
- [PDF 指南](docs/guide/pdf.md)
- [PPTX 指南](docs/guide/pptx.md)
- [组件手册](docs/components/README.md)
- [公开 API 合同](docs/api/public-api-contract.md)
- [对象语义与选择设计](docs/product/object-semantics-and-selection.md)
- [对象引用与选择技术设计](docs/product/object-reference-and-selection-technical-design.md)
- [完整文档索引](docs/INDEX.md)

## 质量与发布纪律

每个候选版本都必须从公开包边界验证，而不只是确认 monorepo 内部能运行。发布门禁包括：

- TypeScript 与 Vue 类型检查；
- 全工作区正式构建；
- 单元与组件行为测试；
- 真实浏览器黑盒与压力流程；
- 两份隔离源码的可复现构建；
- 九个真实 npm 压缩包的工作区外安装；
- Vue 与浏览器兼容矩阵；
- 文档合同和发布就绪检查。

验证方法见 [VERIFICATION.md](VERIFICATION.md)，已发布变更见 [RELEASE_NOTES.md](RELEASE_NOTES.md)。

## 本地开发

仓库当前验证基线为 Node.js `22.13.0` 和 pnpm `9.0.6`；公开包尚未声明更宽泛的 Node `engines` 范围。浏览器流程还需要安装 `requirements-ci.txt` 中的 Python 依赖。

```bash
pnpm install --frozen-lockfile
pnpm dev
pnpm typecheck
pnpm build
pnpm test
```

执行完整发布门禁：

```bash
pnpm test:release
```

## 参与贡献

欢迎提交 Issue 和范围清晰的 Pull Request。发起 PR 前请：

1. 说明用户可见的 Office 场景和涉及的文件格式；
2. 除非是可复用的选择器能力，否则把产品工作流状态留在宿主；
3. 在最低有效层补测试，公共 API 变化还要覆盖真实包消费边界；
4. 运行 `pnpm check`，并同步对应指南或 API 合同。

Bug、兼容缺口和功能提案请使用 [GitHub Issues](https://github.com/arcships/agentic-office-ui/issues)。请勿上传真实机密 Office 文件，测试材料应尽量使用最小、合成的夹具。

## 致谢

DOCX 与 XLSX 的实现参考了公开的 Extend UI / Extend AI React 包。上游归属、固定对比提交和署名信息记录在 [docs/upstream-extend-ui.md](docs/upstream-extend-ui.md)。本项目是独立的 Vue 3 实现，并在各包中保留适用的许可证声明。

## 许可证

[Apache License 2.0](LICENSE)
