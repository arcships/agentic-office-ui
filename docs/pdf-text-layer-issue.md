# PDF 文字选择层（历史方案）

> 状态：已回退并废止（`0.5.0` 不含此功能）
> 回退原因：`@embedpdf` 引擎 `getPageTextRects` 对 CID 字体 + 缺失 ToUnicode CMap 的中文 PDF 输出可见乱码，无法通过字符清洗修复。
> 适用范围：`packages/vue-pdf`
> 关联：PdfSurface、PdfViewer
> 替代设计：[`surface-search-and-pdf-selection-design.md`](./surface-search-and-pdf-selection-design.md)

## 结论更新（2026-07-13）

本文以下“透明文字块 `<span>` + 浏览器原生 Selection”只保留为历史记录，不得继续实现。复盘确认旧尝试有三类根因：

1. `@embedpdf` 2.14.4 已在 `getPageTextRects()` / `getPageGeometry()` 内调用 `FPDF_PageToDevice()`，结果为左上角原点的 device-space 坐标；旧代码再次按 PDF 左下角坐标翻转 Y，造成二次转换。
2. 浏览器会用本机/替代字体重新排版透明文本，字宽、kerning、ligature 与 PDF glyph advance 不一致，长行会累计漂移并选到相邻字符。
3. CID 字体缺失或错误 ToUnicode 时没有可靠 Unicode 可复制，字体 fallback 和 CSS 无法恢复原文。

新方案使用 PDFium 字符索引、glyph 几何和矩形 overlay 完成 hit-test、选择、高亮与复制，并对坏 Unicode 映射显式降级。完整公共 API、交互、OCR 边界和验收矩阵见替代设计。

## 问题

当前 PDF 渲染输出为 PNG 图片（`PdfRenderRuntime.renderPage()` → Blob → `<img>`），页面上的文字无法被选中或复制。

## 技术基础

`@embedpdf` 引擎已暴露以下 API：

- `getPageTextRects(doc, page)` → `PdfTextRectObject[]`：每页所有文字块的 content + bounding rect
- `getPageTextRuns(doc, page)` → `PdfPageTextRuns`：带字体/颜色/字符索引的富文本 run
- `searchAllPages(doc, query)` → `PdfSearchHit[]`：已使用

## 已废止的实现计划

### 1. PdfRenderRuntime 新增公开方法

```ts
interface PdfTextRectItem {
  content: string
  x: number
  y: number
  width: number
  height: number
}

getPageTextRectItems(doc: PdfRenderDocument, pageIndex: number): Promise<PdfTextRectItem[]>
```

透传 `engine.getPageTextRects()`，做坐标归一化。

### 2. PdfSurface 添加 TextLayer 子组件

- 每页 `<img>` 上方叠加 `position: absolute` 的透明容器
- 每个 `TextRectItem` → 一个 `<span>`，`position: absolute`，映射 PDF 坐标到 CSS 坐标
- 样式：`color: transparent`、`user-select: text`、`pointer-events: none`（不拦截滚动）
- 选择用浏览器原生 `Selection API`，无需自定义选中逻辑

### 3. PdfViewer 同步支持

单页模式同样叠加 text layer（按当前页渲染）。

## 历史工作量评估（不再适用）

半天到一天。核心是坐标映射（PDF 坐标系 → CSS 像素），其余为薄封装。
