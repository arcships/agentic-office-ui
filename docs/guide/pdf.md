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

`source` 是推荐输入；同时提供 `source` 和旧的 `src` 时，`source` 优先。URL 来源必须显式提供 `urlPolicy`，本地 `Blob`、`File` 和字节输入不需要网络规则。URL 加载会跟随重定向，但最终地址会按同一协议与来源规则重新校验；跨来源或协议降级重定向会在读取正文前拒绝。诊断和错误只暴露去掉查询串、片段与内嵌凭据后的地址。

| 属性 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `source` | `PdfSource` | 无 | 推荐的 PDF 来源 |
| `src` | `string` | 无 | 兼容 URL 输入 |
| `urlPolicy` | `PdfUrlPolicy` | 无 | URL 访问规则 |
| `runtime` | `PdfRenderRuntime` | 组件创建 | 调用方持有的 Runtime |
| `pdfiumWasmUrl` | `string` | 包内资源 | 覆盖 PDFium WASM 地址 |
| `maxFileSize` | `number` | 50 MiB | 最大 PDF 字节数 |
| `password` | `string` | 无 | 加密 PDF 的初始密码 |
| `onPassword` | `(request) => string \| Promise<string \| undefined>` | 无 | 需要密码时调用；返回 `undefined` 表示取消 |
| `defaultZoom` | `number` | `1` | 初始缩放倍率，范围由组件约束 |
| `pageLayout` | `"continuous" \| "single" \| "spread" \| null` | `null` | 页面布局；`null` 跟随内部工具栏状态 |
| `defaultPageLayout` | `"continuous" \| "single" \| "spread"` | `"single"` | 未受控时的初始页面布局 |
| `showToolbar` | `boolean` | `true` | 工具栏 |
| `showDownload` | `boolean` | `true` | 下载按钮 |
| `showRotateControls` | `boolean` | `true` | 旋转按钮 |
| `minHeight` | `string \| number` | 无 | 可选最小高度（数字按 px）；不传时组件不设置最小高度 |
| `isDark` | `boolean` | `false` | 深色外壳与页面背景；页面位图默认反色，可用 `--pdf-dark-page-filter: none` 关闭 |

`PdfSurface` 也接受 `isDark` 和 `minHeight`。打印时工具栏、缩略图和状态层会自动隐藏，PDF 会在打印前一次性栅格化全部页面（而不是只打印当前页或占位符），并按文档首页尺寸设置 `@page`。

### 页面布局

`PdfViewer` 和 `PdfSurface` 共用同一个 `pageLayout` 属性：

- `single`（查看器默认）：一次只显示一页，适合逐页阅读。
- `continuous`：整份文档纵向堆叠、按可视窗口懒加载。`PdfSurface` 一直是这个模式。
- `spread`：两页并排，按 (1,2)、(3,4) 分组；行高取两页中较高的一页，落单的最后一页居中占满整个跨页宽度。两部分都把对页栅格化成独立位图，每页保留自己的页码（`data-page-number`、`alt` 文案）。

切换布局不会丢失当前页：新布局会把滚动位置对齐到当前页所在的行。`PdfViewer` 工具栏的下拉框（`data-testid="pdf-page-layout"`）驱动未受控的布局，并抛出 `update:pageLayout`；`PdfSurface` 只接受受控属性。渲染窗口按模式收敛：`single` 只保留当前页的位图，`continuous`/`spread` 按可视行懒加载。

需要自建 PDF 表面时使用根入口导出的纯函数 `resolvePdfPageLayoutMetrics()`、`pdfPageLayoutRows()`、`pdfRowIndexAtOffset()`、`pdfPageIndexAtOffset()`、`pdfSnapTopForRow()`、`pdfPageAtPoint()`、`pdfRowsInViewport()`、`pdfSpreadPagePair()`、`pdfSpreadStepPage()`、`normalizePdfPageLayout()`、`isPdfSpreadLayout()` 和 `PDF_PAGE_LAYOUTS`，类型为 `PdfPageLayout`、`PdfPageLayoutMetrics`、`PdfPageLayoutRow`、`PdfPageLayoutOptions`、`PdfPageDisplaySize`。

事件包括 `document-load-success`、`document-load-error`、`active-page-change`、`update:pageLayout`、`selectionChange`、`fileDrop` 和 `diagnostic`。当前页支持拖选文字、双击选词、`Ctrl/Cmd+C` 复制和 `Escape` 清除；`clearSelection()` 与 `getSelectionState()` 通过组件实例暴露。位图上方还渲染真实 DOM 文本层（每个 glyph 一个透明 span，随缩放/旋转应用 CSS matrix），可由屏幕阅读器、浏览器查找和 CJK 提取读取；文本层不拦截指针，拖选、链接/表单命中与区域选择行为不变。错误处理使用 `PdfLoadError.code`，不要通过错误文案判断类型。加密文档在无密码可用时返回 `ENCRYPTED`；`load-start` 诊断中的 URL 已去除查询串、片段和凭据。

### 拖放导入

把 `.pdf` 拖到查看器上会作为新的有效来源，走与 `source` 属性相同的 `load()` 流程，因此大小上限、URL 规则、密码回调和诊断事件都保持一致；再次拖入会替换当前文档。宿主更新 `source`/`src` 后，属性来源重新优先。`PdfViewer` 同时抛出 `fileDrop(file)` 事件。拖入时根元素出现虚线轮廓，拖离、放下或拖拽在窗口其他位置结束后清除；不是 `.pdf` 的文件不会加载，并在查看器上方显示本地化的拒绝提示（`pdf.viewer.unsupportedDrop`）。

## 键盘导航

`PdfViewer` 的快捷键与 DOCX 共用同一套映射：`PageDown`/`PageUp` 翻一页，`Home`/`End` 跳到第一页/最后一页，`Ctrl/Cmd` + `+`/`=`、`-`、`0` 缩放或重置到 100%，`Ctrl/Cmd + F` 聚焦搜索框，`Escape` 先取消搜索再清除选择；`Ctrl/Cmd + P` 保留为打印。`PdfSurface` 的 `PageDown`/`PageUp`、`Home`/`End` 按页滚动连续页面，`Ctrl/Cmd` 缩放通过受控 `update:zoom` 通知宿主，`Escape` 取消选择或区域。焦点在输入框、文本域、下拉框或 `contenteditable` 内时，查看器不会拦截这些按键。完整映射见[查看器键盘导航](viewer-keyboard.md)。

需要完全控制 Runtime 时，可以使用 `createPdfRenderRuntime()`。调用方传入的 Runtime 必须由调用方销毁。单页渲染会在进入 PDFium 前校验 `scale`、`dpr` 和旋转，并按 `maxRenderPixels`（默认 32 兆像素）拒绝超出预算的请求，避免一次渲染耗尽内存。
