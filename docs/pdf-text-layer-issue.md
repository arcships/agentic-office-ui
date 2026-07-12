# PDF 文字选择层

> 状态：已回退（`0.5.0` 不含此功能）
> 回退原因：`@embedpdf` 引擎 `getPageTextRects` 对 CID 字体 + 缺失 ToUnicode CMap 的中文 PDF 输出可见乱码，无法通过字符清洗修复。
> 适用范围：`packages/vue-pdf`
> 关联：PdfSurface、PdfViewer

## 问题

当前 PDF 渲染输出为 PNG 图片（`PdfRenderRuntime.renderPage()` → Blob → `<img>`），页面上的文字无法被选中或复制。

## 技术基础

`@embedpdf` 引擎已暴露以下 API：

- `getPageTextRects(doc, page)` → `PdfTextRectObject[]`：每页所有文字块的 content + bounding rect
- `getPageTextRuns(doc, page)` → `PdfPageTextRuns`：带字体/颜色/字符索引的富文本 run
- `searchAllPages(doc, query)` → `PdfSearchHit[]`：已使用

## 实现计划

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

## 工作量评估

半天到一天。核心是坐标映射（PDF 坐标系 → CSS 像素），其余为薄封装。
