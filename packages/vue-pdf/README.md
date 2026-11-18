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

## 本地化

`PdfViewer` 与 `PdfSurface` 都接受 `locale`（默认 `"en"`，内置 `"zh-CN"`）与 `messages` 覆盖。用户可见文案和 aria-label 来自 `@arcships/office-interaction` 的共享消息目录：

```vue
<PdfViewer
  :src="pdfUrl"
  locale="zh-CN"
  :messages="{ pdf: { viewer: { download: '保存文件' } } }"
/>
```

正常 PDF 支持查看、翻页、缩放、旋转、缩略图、搜索和下载。PDF 唯一的公开文件拒绝上限是整份文件体积 `maxFileSize`，默认 `50 MiB`，宿主可调整。加密 PDF 可通过 `password` 或 `onPassword` 提供密码；无法取得密码时错误码为 `ENCRYPTED`。URL 来源必须提供 `urlPolicy`；同源重定向会被跟随并重新校验最终地址，跨来源或协议降级重定向在读取前拒绝，诊断和错误地址会去除查询串、片段与凭据。

`PdfSurface` 提供最小嵌入渲染：全部页面垂直堆叠滚动，无工具栏/缩略图/搜索栏。仅视口附近的页面会栅格化（默认上下各预留 1 页），离开视口的位图进入字节预算内的 LRU 缓存，滚动回去直接复用，超出预算才按最久未用逐张释放；缩放、旋转和容器尺寸变化只重渲染当前视口，不再重渲染整份文档。每页栅格化像素数另有上限（默认 8 MP），超出时降低渲染倍率并由 CSS 缩小显示。滚动与右键命中复用缓存的页面偏移，滚动事件按帧合并并做 100 ms 防抖，容器 resize 做 120 ms 防抖。宿主可通过 `v-model:zoom` 接入 50%–200% 手势缩放，也可继续使用 expose zoom 和 scrollToPage。

内存相关的两个属性默认适用于笔记本与移动端，宿主可按设备下调：

| 属性 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `maxBitmapBytes` | `number` | `64 MiB` | 视口外页面位图的字节预算；设为 `0` 时退化为离开视口立即释放 |
| `maxPagePixels` | `number` | `8_000_000` | 单页栅格化像素上限（设备像素），超出则降低渲染倍率并缩小显示 |

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

字符范围在同修订内精确验证；修订变化后，唯一文字匹配返回 `relocated`，重复文字返回 `ambiguous`，删除返回 `not-found`。页面和人工区域只有在页序、尺寸/旋转以及可用的页面文字摘要均未变时才迁移；新引用会带上页面文字摘要，避免同尺寸换内容被误认为同一页。

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

单页渲染在进入 PDFium 前校验 `scale`、`dpr` 和旋转，并按 `maxRenderPixels`（默认 32 兆像素）拒绝超预算请求。Runtime 不会修改全局 `URL.createObjectURL`；若宿主有严格 CSP，需要允许 Worker 使用 `blob:` 脚本。

## 文档

- [PDF 使用指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/guide/pdf.md)
- [组件手册](https://github.com/arcships/agentic-office-ui/blob/master/docs/components/README.md)
- [自定义 PDF 界面](https://github.com/arcships/agentic-office-ui/blob/master/docs/custom-components/pdf.md)
- [公开接口合同](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/public-api-contract.md)

## License

Apache-2.0
