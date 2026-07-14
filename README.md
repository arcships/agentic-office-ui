<div align="center">

<img src="docs/assets/banner-agent-collaboration.png" alt="Agentic Office UI：面向 Office Agent 的 Vue 3 人机协同组件库" width="100%">

# Agentic Office UI

面向 Office Agent 的 Vue 3 人机协同组件库

**让人看得见、指得准，让 Agent 找得到，让结果验得了。**

[![npm version](https://img.shields.io/badge/npm-v0.5.4-cb3837)](https://www.npmjs.com/org/arcships)
[![CI](https://github.com/arcships/agentic-office-ui/actions/workflows/ci.yml/badge.svg)](https://github.com/arcships/agentic-office-ui/actions/workflows/ci.yml)
[![Release](https://github.com/arcships/agentic-office-ui/actions/workflows/release.yml/badge.svg)](https://github.com/arcships/agentic-office-ui/actions/workflows/release.yml)
[![Vue](https://img.shields.io/badge/Vue-%3E%3D3.2.25%20%3C4-42b883)](https://vuejs.org/)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](LICENSE)

[开始使用](docs/guide/getting-started.md) · [组件手册](docs/components/README.md) · [组合函数](docs/composables/README.md) · [自定义组件](docs/custom-components/README.md) · [完整文档](docs/INDEX.md)

</div>

> npm 稳定版为 `0.5.4`。当前源码正在准备 `0.6.0`：已加入跨格式引用合同、四格式阶段一定位器、统一 Surface 选择事件和区域选择原语。截图位图采集、标注编辑器与自动修改审阅尚未内置。本项目不承诺与 Microsoft Office 的全部高级能力或像素级效果完全一致。

## 为什么需要这个库

Agent 通常通过代码操作 DOCX、XLSX、PDF 和 PPTX：它读取文件结构、调用文档工具、生成新文件。人类面对的却是段落、表格、单元格、幻灯片对象和页面区域。

这里存在两个缺口：

1. Agent 修改文件时，人类往往无法在同一界面完整预览正在处理的内容和修改结果。
2. 人类很难仅靠语言准确描述“要修改哪里”。“第二页右下角那张表”对人直观，对 Agent 却不是稳定定位信息。

Agentic Office UI 位于人和 Agent 之间。项目的目标工作流是：

- 用浏览器内的 Office Surface 展示 Agent 正在处理的输入文件；
- 把文字选择、单元格范围、对象点击、区域框选和截图标注转换成精确引用；
- 让宿主把“用户指令 + 精确引用”交给任意 Agent 或文档工具；
- 在 Agent 返回新文件后重新加载，通过同一引用重新定位并验证结果。

```mermaid
flowchart LR
  A["Agent 通过代码处理 Office 文件"] --> B["Office Surface 浏览器内预览"]
  B --> C["用户选择文字、单元格、对象或区域"]
  C --> D["生成精确 Office Reference"]
  D --> E["宿主把引用和指令交给 Agent"]
  E --> F["重新加载并验证修改结果"]
  F --> C
```

上图描述的是组件库与宿主共同完成的闭环。当前 `0.6.0` 源码候选已经覆盖预览、统一引用、区域选择、重新解析和滚动定位；截图位图采集、标注编辑器及自动 diff 仍由宿主实现或留待后续可插拔能力。

这个库不绑定模型厂商，不调用模型，不编排 Agent，也不要求 Agent 使用某一种 Office 修改工具。它负责的是人机协同中的交互层：**预览、引用、表达和验证**。

`DocxEditor` 和 `XlsxViewer` 仍提供人工编辑与本地文档命令；在 Agent 工作流中，自动修改可以由宿主连接的 `python-docx`、`openpyxl`、LibreOffice 或其他工具完成。组件库负责交互和结果呈现，不替宿主决定修改工具。

## 精确引用，而不是庞大的 Context

Agent 通常可以自行读取原文件，因此交互层不应默认传递 DOM、整份文档模型、全部视口状态或长篇拼装文本。用户意图只需要两部分：

- **指令**：要做什么；
- **引用**：对哪些内容做。

一个可靠的引用需要包含足够的定位与校验信息，但不需要暴露整套 UI 状态：

| 信息 | 作用 |
|---|---|
| 文件和版本 | 确认引用属于哪一份输入，避免修改过期文件 |
| 语义位置 | 让 Agent 能通过段落、单元格、幻灯片对象或 PDF 文字范围定位 |
| 内容证据 | 用精确引文、前后文、公式或对象类型排除相似目标 |
| 视觉证据 | 在结构信息不足时，用页面区域或标注截图补充说明 |

语义位置用于执行，内容和视觉证据用于消歧。具体字段、坐标和序列化格式见[对象引用与选择技术设计](docs/product/object-reference-and-selection-technical-design.md)；README 只解释产品概念。

## 统一引用模型覆盖什么

阶段一先覆盖高频且可稳定定位的文字、单元格、页面、幻灯片对象和区域；下表同时列出后续对象层级的扩展方向。

| 格式 | 可以引用的元素 | Agent 的主要定位特征 | 用户的快速表达方式 |
|---|---|---|---|
| DOCX | 文字范围、段落、标题、列表项、表格、单元格、图片、批注 | 文档节点路径、文字偏移、精确引文、前后文、页码区域 | 拖选文字、点击块、框选区域、多选引用 |
| XLSX | 单元格、范围、行列、工作表、表格、公式、图表、图片 | 工作表名、A1 范围、表格名、对象 id、公式 | 拖选范围、点行列头、点击图表或图片 |
| PPTX | 幻灯片、形状、文本框、图片、表格、图表、组合对象 | 幻灯片索引、稳定 `objectKey`、对象类型、文字范围 | 点击对象、拖选文字、框选多个对象 |
| PDF | 文字范围、页面、矩形区域、版面块、图片区域 | 页面索引、字符范围、精确引文、归一化矩形 | 拖选文字、区域框选、截图标注 |

单一特征并不总是可靠：坐标会随缩放变化，文本可能重复，文档节点也可能在修改后重排。引用因此需要组合语义位置、内容证据和必要的视觉证据，并绑定当前文件版本。

## 除了语言，还可以怎样表达意图

在目标交互中，语言负责说明“做什么”，交互引用负责说明“对什么做”。宿主可以组合以下方式：

- **直接选择**：拖选 DOCX/PDF 文字或 XLSX 单元格范围；
- **对象引用**：点击图片、图表、形状、表格或幻灯片对象；
- **区域框选**：框出无法用单一文档节点描述的视觉区域；
- **多目标引用**：同时附加多个目标，例如“修改这三段，但保留这张表”；
- **角色引用**：分别标记目标、素材、样式参考和插入位置，例如“把目标改成参考对象的样式”；
- **相对关系**：用多个引用表达“插到这段后面”“用这张表替换那张表”；
- **截图标注**：截取绑定到页面或幻灯片区域的局部画面，并附加矩形、箭头、画笔和文字标记；
- **快捷动作**：由宿主提供替换、删除、移动、插入、格式化等动作按钮，减少自然语言中的动作歧义。

截图不是脱离文件的普通图片。它必须和原文件中的页面、幻灯片、工作表或对象保持关联，并保留用户画出的矩形、箭头、画笔和文字说明。这样 Agent 既能理解视觉意图，也能回到原文件定位对应内容。具体截图格式和存储方式由后续设计确定。

## 能力边界

| 环节 | 当前状态 | 责任方 |
|---|---|---|
| Office 文件预览 | 四格式 Viewer/Surface 已在 `0.5.4` 发布 | 组件库 |
| 搜索、文字或范围选择、对象事件 | 已在 `0.5.4` 按格式提供 | 组件库 |
| 跨格式引用、解析和统一 Surface 事件 | 已进入 `0.6.0` 源码候选 | 组件库 |
| 对象轮廓与区域选择原语 | 已进入 `0.6.0` 源码候选 | 组件库 |
| DOCX/XLSX 人工编辑 | 已提供部分编辑和导出能力 | 组件库与用户 |
| 保存多目标引用、组合语言指令 | 由宿主按产品需要组织 | 宿主应用 |
| 调用模型、编排 Agent、选择文件修改工具 | 不属于组件库 | 宿主应用 |
| 截图位图采集和标注编辑器 | 尚未内置，Surface 会明确报告不支持采集 | 宿主应用或后续插件 |
| Agent 修改后的自动 diff、接受或拒绝 | 尚未内置 | 宿主应用或后续能力 |

## 当前能做什么

| 格式 | 开箱即用 | 可组合能力 |
|---|---|---|
| DOCX | 查看、编辑、分页、缩略图、搜索、文字选择、批注、修订、导入导出 | 精确文字、页面和区域引用，解析与重新定位，Runtime、Worker/WASM |
| XLSX | 工作表、公式、图表、图片、范围选择、编辑、撤销重做 | 单元格、范围、行列、工作表、图表和区域引用，Runtime、按需图表/地图/WebGL |
| PDF | PDFium 渲染、翻页、缩放、旋转、缩略图、搜索、文字选择、下载 | 精确字符范围、页面和区域引用，来源规则和渲染 Runtime |
| PPTX | 纵向预览、缩略图、搜索、对象事件、逐步动画、切换、媒体、全屏 | 幻灯片、精确文字、稳定对象和区域引用，播放控制器与无界面组合函数 |

项目同时提供完整 Vue 组件和不依赖 Vue 的核心包。你可以从一个查看器开始，也可以接管状态、工具栏和交互意图界面。

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

统一引用 API 属于当前 `0.6.0` 源码候选。发布后，直接消费共享类型或纯函数的宿主再显式安装 `@arcships/office-interaction`；四种格式包会自行声明它们的运行依赖。

`0.6.0` 源码候选中的四种 Surface 会把普通选择转换成同一种确认事件。宿主只需保存引用，不必再自行拼接工作表名、A1 范围和文档 revision：

```vue
<script setup lang="ts">
import { shallowRef } from "vue"
import { XlsxSheetSurface, useXlsxViewerController } from "@arcships/vue-xlsx"
import type { OfficeReferenceConfirmEvent } from "@arcships/office-interaction"

const props = defineProps<{ file: ArrayBuffer }>()
const controller = useXlsxViewerController({
  file: props.file,
  fileName: "workbook.xlsx",
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

## 选择正确的入口

| 你想做什么 | 使用什么 |
|---|---|
| 尽快显示一个文档 | `DocxViewer`、`XlsxViewer`、`PdfViewer`、`PptxViewer` |
| 自建工具栏嵌入渲染区 | `DocxDocumentSurface`（仅分页渲染 + 批注 gutter） |
| 编辑 DOCX | `DocxEditor` 或 `useDocxEditor` |
| 自定义 XLSX 工具栏 | `useXlsxViewerController` + `XlsxSheetSurface` |
| 自建 PDF 控制栏 | `PdfSurface`（垂直滚动全部页面） |
| 自建 PPTX 浏览 Surface | `PptxStage` + `usePptxDocument({ session: { renderMode: "list" } })` |
| 自定义 PPTX 播放器 | `usePptxDocument` + `usePptxPlayback` + `PptxStage` |
| 在非 Vue 代码中处理文档 | `docx-core`、`xlsx-core`、`pptx-core` |
| 管理 Worker、WASM 和资源限制 | 对应格式的 Runtime |

## 公开包

| 包 | 说明 |
|---|---|
| [`@arcships/office-interaction`](packages/office-interaction/README.md) | 跨格式对象引用、运行时校验、选择状态与纯函数 |
| [`@arcships/docx-core`](packages/docx-core/README.md) | DOCX 模型、布局、命令和 Runtime |
| [`@arcships/vue-docx`](packages/vue-docx/README.md) | `DocxDocumentSurface`、`DocxViewer`、`DocxEditor` 和组合函数 |
| [`@arcships/xlsx-core`](packages/xlsx-core/README.md) | XLSX 模型、公式、图表和 Runtime |
| [`@arcships/vue-xlsx`](packages/vue-xlsx/README.md) | `XlsxSheetSurface`、`XlsxViewer` 和查看器控制器 |
| [`@arcships/vue-pdf`](packages/vue-pdf/README.md) | PDFium 查看器与渲染 Runtime |
| [`@arcships/pptx-core`](packages/pptx-core/README.md) | PPTX 预览、播放模型和浏览器控制器 |
| [`@arcships/vue-pptx`](packages/vue-pptx/README.md) | `PptxStage`、`PptxViewer`、Surface 事件和最小播放组合能力 |
| [`@arcships/vue-ui`](packages/vue-ui/README.md) | 上传、签名、缩略图、引用框、Find Bar 和版面组件 |

九个公开包使用统一版本发布。npm 稳定版为 `0.5.4`；当前源码候选为 `0.6.0`。相关 Vue 包、核心包和 `office-interaction` 应保持相同 minor 版本。

PPTX 两包从 `0.3.0` 开始公开，`0.4.0` 增加最小组合接口。

## 文档

### 按格式

- [DOCX 使用指南](docs/guide/docx.md)
- [XLSX 使用指南](docs/guide/xlsx.md)
- [PDF 使用指南](docs/guide/pdf.md)
- [PPTX 使用指南](docs/guide/pptx.md)
- [引用选择接入指南](docs/guide/reference-selection.md)

### 按问题

- [组件属性、事件和选择建议](docs/components/README.md)
- [组合函数和状态所有权](docs/composables/README.md)
- [自定义工具栏与播放器](docs/custom-components/README.md)
- [组合函数、Runtime 和核心入口](docs/api/README.md)
- [公开入口、错误码和兼容期限](docs/api/public-api-contract.md)
- [版本升级说明](docs/migration-0.3.md)
- [浏览器兼容矩阵](docs/testing/compatibility-matrix.md)

## 设计原则

- Office Surface 负责预览与交互，不绑定 Agent、模型厂商或业务工作流。
- 用户意图由精确引用和简短指令组成，不默认传递庞大的 UI 状态。
- 语义定位是执行依据，引文和视觉区域负责消歧与校验。
- 高层组件解决常见场景，核心包和控制器允许宿主接管交互界面。
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

PDF 默认最大文件大小为 `50 MiB`，超过后返回稳定错误码 `PDF_TOO_LARGE`。

## License

[Apache 2.0](LICENSE)
