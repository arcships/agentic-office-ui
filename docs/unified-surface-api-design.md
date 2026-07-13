# Unified Surface API 设计

> 状态：已实现；PPTX 纵向列表与完整交互事件在 `0.5.2` 收口
> 适用范围：`DocxDocumentSurface`、`XlsxSheetSurface`、`PptxStage`、`PdfSurface`
> 目标：四个 surface 组件对外统一暴露渲染边界、样式定制、自适应缩放、鼠标事件和选中状态
> 搜索与 PDF 文字选择：见 [`surface-search-and-pdf-selection-design.md`](./surface-search-and-pdf-selection-design.md)

> 2026-07-13 更新：本文继续负责 Surface 边界与通用交互命名。搜索状态机、格式命中锚点、Find Bar、精确定位和 PDF glyph 选择以新设计为准；PDF 不再等待透明 text layer，而改用字符几何 + overlay。

## 1. 设计原则

1. **Surface 只做渲染面**：分页/网格/画布本身 + 内置必需交互（cell 编辑、滚动、批注 gutter）。Toolbar、菜单、属性面板全部由宿主构建。
2. **统一的 props / events / expose 命名**：能对齐的尽量对齐，格式特有逻辑走差异化字段。
3. **同层能力不要缺位**：一个 surface 有 background 定制，其他三个也要有；一个有 context menu 事件，其他三个也要能触发。

## 2. 统一边界 & 样式

### 2.1 组件边界

Surface 的根 DOM 元素 = 内容画布本身，无外层 chrome、无 border-radius、无 padding。宿主负责给它套容器和样式。

### 2.2 背景色

每个 surface 暴露 `--xxx-surface-bg` CSS 变量，默认值保持当前外观：

| 组件 | CSS 变量 | 默认值 |
|---|---|---|
| `DocxDocumentSurface` | `--docx-surface-bg` | `#f4f4f5` / `#111827`（已有） |
| `XlsxSheetSurface` | `--xlsx-surface-bg` | `#ffffff` / `#18181b` |
| `PptxStage` | `--pptx-surface-bg` | 透明（继承容器） |
| `PdfSurface` | `--pdf-surface-bg` | `#525659` |

### 2.3 自适应缩放（fit-width）

新增 prop `fitWidth?: boolean`（默认 `false`）。启用后 surface 监听容器 ResizeObserver，自动计算 zoom = containerWidth / contentWidth，保持宽高比。

适用范围：
- **DOCX**：页面宽度 = `pageWidthPx`，zoom 映射到 `DocxViewerRoot.zoomScale`
- **PDF**：最大页宽，zoom 映射到 `PdfSurface.zoom`
- **XLSX**：不需要（网格本身就是滚动视口）
- **PPTX**：没有独立 `fitWidth` prop；列表/单页渲染都使用会话 `fitMode: "contain"` 随容器宽度计算页面尺寸

### 2.4 滚动

横向 + 纵向滚动由 surface 内置处理。DOCX、PPTX 和 PdfSurface 纵向滚动全部内容，XLSX 横纵二维滚动。PPTX 只有 `renderMode: "slide"` 的播放舞台不滚动页面列表。

## 3. 统一事件模型

### 3.1 Context Menu

每个 surface 发出 `context-menu` 事件，携带格式相关的选中上下文，**由宿主决定显示什么菜单**：

```ts
// DOCX
emit("contextMenu", {
  kind: "paragraph" | "table" | "image" | "text-selection",
  nodeIndex: number,
  pageIndex: number,
  clientX: number,
  clientY: number,
})

// XLSX
emit("contextMenu", {
  kind: "cell" | "column-header" | "row-header" | "chart" | "image",
  selection: XlsxCellRange,
  clientX: number,
  clientY: number,
})

// PPTX
type PptxStageContextMenu = {
  slideIndex: number,
  clientX: number,
  clientY: number,
  containerX: number,
  containerY: number,
} & (
  | { kind: "slide"; objectKey?: never }
  | { kind: "object"; objectKey: string }
)

emit("contextMenu", context satisfies PptxStageContextMenu)

// PDF
emit("contextMenu", {
  kind: "page" | "text-selection",
  pageIndex: number,
  selectedText?: string,
  clientX: number,
  clientY: number,
})
```

**注意**：DOCX/XLSX 当前内置了 `ContextMenu` 子组件。需改为：surface 仍可在**内部**渲染默认菜单骨架，但菜单项列表由 `context-menu` 事件 + 宿主回传的 `contextMenuActions` prop 组合决定。或者干脆拆掉内置菜单，由宿主完全自主。

### 3.2 选中内容变化

每个 surface 发出 `selection-change` 事件，携带当前选中的结构化内容：

```ts
// DOCX
emit("selectionChange", {
  kind: "paragraph" | "table-cell" | "image" | "none",
  text?: string,
  nodeIndex?: number,
  pageIndex?: number,
})

// XLSX
emit("selectionChange", {
  kind: "cell" | "range" | "chart" | "image" | "none",
  range?: XlsxCellRange,
  value?: string,
})

// PPTX — 点击页面或对象时选择所在幻灯片
emit("selectionChange", { kind: "slide", slideIndex: number })

// PDF — 实现 glyph geometry selection overlay 后生效
emit("selectionChange", {
  kind: "text" | "none",
  text?: string,
  pageIndex?: number,
})
```

### 3.3 对象点击

点击到图片/图表/形状等嵌入对象时发出 `object-click` 事件：

```ts
// DOCX — 点击图片
emit("objectClick", { kind: "image", nodeIndex: number, imageKey?: string })

// XLSX — 点击图表/图片/形状
emit("objectClick", { kind: "chart" | "image" | "shape", id: string })

// PPTX
emit("objectClick", { kind: "object", slideIndex: number, objectKey: string })

// PDF — 不适用（图片嵌入页面不可独立点击）
```

### 3.4 统一事件清单

| 事件 | DOCX | XLSX | PPTX | PDF | 参数 |
|---|---|---|---|---|---|
| `contextMenu` | ✅ | ✅ | ✅ | ✅ | 格式相关上下文 + clientX/Y |
| `selectionChange` | ✅ | ✅ | ✅ | ✅ | 选中内容结构化描述 |
| `objectClick` | ✅ | ✅ | ✅ | — | 对象 kind + id |
| `pageCountChange` | ✅ | — | — | — | 页数 |
| `visiblePageRangeChange` | ✅ | — | — | — | 可见页范围 |
| `visiblePageChange` | — | — | — | ✅ | 当前可见页 |
| `documentLoadSuccess` | — | — | — | ✅ | 页数 |
| `documentLoadError` | — | — | — | ✅ | PdfLoadError |
| `cellDoubleClick` | — | ✅ | — | — | cell address |
| `searchStateChange` | ✅ | ✅ | ✅ | ✅ | 格式化 search state；精确合同见搜索设计 |

`searchStateChange` 是 2026-07-13 追加的目标合同，尚未包含在本文顶部“已实现”状态中。

## 4. 统一 expose

下表原有 Surface 方法已实现；`scrollToCell` 和 `search*` / `getSearch*` 行是 2026-07-13 追加的目标合同，状态仍为待实现，详细 settle 语义以搜索设计为准。

| 方法 | DOCX | XLSX | PPTX | PDF | 说明 |
|---|---|---|---|---|---|
| `scrollToPage(n)` | ✅ | — | `usePptxDocument.goTo(n)` | ✅ | 滚动到指定页 |
| `scrollToNode(n)` | ✅ | — | — | — | 滚动到指定节点 |
| `scrollToCell(cell)` | — | ✅ | — | — | 等待布局后把单元格滚入视口 |
| `search(query, options)` | ✅ | ✅ | ✅ | ✅ | 执行当前文档会话搜索 |
| `activateSearchMatch(n)` | ✅ | ✅ | ✅ | ✅ | 异步定位并高亮指定结果 |
| `searchNext()` / `searchPrevious()` | ✅ | ✅ | ✅ | ✅ | 循环导航同一搜索会话 |
| `clearSearch()` | ✅ | ✅ | ✅ | ✅ | 取消 pending 并清理高亮 |
| `getSearchState()` / `getSearchCapabilities()` | ✅ | ✅ | ✅ | ✅ | 获取状态快照和格式能力 |
| `zoom` | — | — | — | ✅ | 读写缩放 |
| `rotation` | — | — | — | ✅ | 读写旋转 |
| `scrollContainer` | ✅ | — | — | — | 滚动容器 DOM |

PPTX 通过 `usePptxDocument` composable 暴露 `goTo` / `setZoom`，不走 expose。列表模式下 `goTo` 滚动定位页面，滚动产生的可见页变化同步回 `activeIndex`。

## 5. 实现顺序

| 阶段 | 内容 | 涉及组件 |
|---|---|---|
| 1. 背景色 | `--xxx-surface-bg` CSS 变量 | XLSX、PPTX、PDF |
| 2. fit-width | 新增 prop + ResizeObserver 自适应 | DOCX、PDF |
| 3. context-menu | 拆掉内置菜单，改发 `contextMenu` 事件 | DOCX、XLSX |
| 4. selection-change | 新增 `selectionChange` 事件 | DOCX、XLSX、PDF |
| 5. object-click | 新增 `objectClick` 事件 | DOCX、XLSX |
| 6. demo 更新 | 四个 demo 页展示完整宿主控制 | demo |

阶段 1 改动最小、立即可用。阶段 3-5 涉及拆掉 DOCX/XLSX 内置 ContextMenu 组件，需评估兼容影响——可在 `0.x` 保留旧组件但标记 `@deprecated`。
