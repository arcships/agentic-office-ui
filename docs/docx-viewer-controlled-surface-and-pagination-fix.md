# DOCX Viewer 受控文档面与分页修复设计

> 状态：待实现
>
> 适用范围：`packages/docx-core`、`packages/vue-docx`、`apps/demo`、`tests`
>
> 目标版本：`0.5.0`（新增兼容接口）；现有 `0.x` 接口继续保留
>
> 关联文档：[项目架构审查与目标设计](./architecture-review-and-target-design.md)、[公开 API 合同](./api/public-api-contract.md)、[稳定化整改路线图](./plan/stabilization-roadmap.md)

## 1. 目的

修正 `DocxViewer` 当前的分页、缩放、纸面主题、批注栏布局和外部控制问题，并提供一个可由宿主自定义工具栏直接控制的稳定只读组件。

本设计只处理 DOCX 只读显示和其共享页面渲染基础。编辑命令、保存、输入法和协同编辑不在本轮范围内。

## 2. 当前问题与证据

### 2.1 界面主题和纸面主题绑定

`DocxViewer.isDark` 同时决定 viewer 外壳和页面纸面主题：

```text
DocxViewer.isDark
  ├─ viewer 外壳背景
  └─ DocxDocumentSurface.theme
       └─ DocxPageSurface
            ├─ dark: #1f2937 / #f9fafb
            └─ light: #ffffff / #111827
```

证据：

- `packages/vue-docx/src/components/DocxViewer.vue:100-110`
- `packages/vue-docx/src/components/DocxPageSurface.vue:312-322`

这无法表达常见桌面预览需求：深色 App 外围 + 白色纸张 + 深色正文。

### 2.2 批注和修订默认显示

`DocxViewer` 的默认值为：

```ts
defaultShowTrackedChanges: true
defaultShowComments: true
```

证据：`packages/vue-docx/src/components/DocxViewer.vue:128-157`。

只读预览打开带批注文档时，右侧 gutter 会立即占用页面空间。当前黑盒测试还把默认开启作为通过条件：`tests/blackbox/docx_annotations.py:91-95`。

### 2.3 分页结果与渲染输入不一致

仓库存在两套布局路径：

```text
A. 旧 LayoutSnapshot 路径（DocxViewerRoot 当前使用）
DocModel
  -> layout/index.ts: buildLayoutSnapshot()
  -> layout-engine.ts: layoutDocument()
  -> 整段 / 整表 block 分页

B. 分段分页路径（具备跨页语义）
DocModel
  -> page-segmentation-core.ts: buildDocumentPageNodeSegments()
  -> DocumentPageNodeSegment
       ├─ nodeIndex
       ├─ paragraphLineRange
       └─ tableRowRange
```

`DocxViewerRoot` 当前使用 A 路径：

- `packages/vue-docx/src/components/DocxViewerRoot.vue:153-165`
- `packages/docx-core/src/layout/index.ts:378-392`
- `packages/docx-core/src/layout/layout-engine.ts:317-358`

A 路径只会把完整 block 移到下一页；超过单页高度的长段落或表格不会拆分。

B 路径已经计算段落行范围和表格行范围：

- `packages/docx-core/src/layout/page-segmentation-core.ts:258-266`
- `packages/docx-core/src/layout/page-segmentation-core.ts:452-556`

但 `DocxViewerRoot` 构造 `pageNodeSegments` 时仅保留 `nodeIndex`：

```ts
[{ nodeIndex }]
```

证据：`packages/vue-docx/src/components/DocxViewerRoot.vue:153-165`。

`DocxPageBody` 随后按完整 `ParagraphNode` 渲染，并没有按 `paragraphLineRange` 截取段落：

- `packages/vue-docx/src/components/DocxPageBody.ts:271-304`

结果是长段落溢出页面；下一页的绝对定位纸面覆盖溢出内容，用户看到页尾文字被裁切、跨页内容重叠或重复。

### 2.4 页面高度存在双重缩放

页面内容使用 CSS `transform: scale(zoomFactor)`：

- `packages/vue-docx/src/components/DocxViewerRoot.vue:355-373`

页面通过 `getBoundingClientRect().height` 上报测量值：

- `packages/vue-docx/src/components/DocxPageSurface.vue:333-339`

`getBoundingClientRect()` 已包含 CSS transform。`DocxViewerRoot` 保存这个高度后，在 page offset 和 total height 中又乘一次 `zoomFactor`：

- `packages/vue-docx/src/components/DocxViewerRoot.vue:278-292`
- `packages/vue-docx/src/components/DocxViewerRoot.vue:403-409`

因此低于 100% 时下一页提前出现，高于 100% 时页间距离异常增大。

### 2.5 批注 gutter 宽度没有进入页面布局

批注 gutter 使用：

```css
left: calc(100% + 16px);
width: 240px;
```

证据：`packages/vue-docx/src/components/DocxTrackedChangeGutter.vue:147-156`。

页面容器只额外增加 `48px`：

- `packages/vue-docx/src/components/DocxViewerRoot.vue:345-353`

实际 gutter 至少需要 256px。打开批注后，水平滚动范围和页面居中都不准确。

### 2.6 高层 Viewer 无法由外部工具栏控制

`DocxViewer` 内部拥有：

- `zoomPercent`
- `searchQuery`
- `searchResultIndex`
- `currentPage`
- `totalPages`

证据：`packages/vue-docx/src/components/DocxViewer.vue:170-212`。

公开 props 只提供 `defaultZoom`，没有受控 `zoom`、`searchQuery`、当前搜索结果和页码；公开 expose 也没有 `searchNext()`、`searchPrevious()`、`scrollToPage()`。宿主关闭默认 toolbar 后，无法实现自己的搜索和分页 UI。

## 3. 设计决定

1. 新增稳定高层组件 `DocxSurface`，用于无默认工具栏的受控只读显示。
2. `DocxViewer` 保留并改为组合 `DocxSurface + DocxViewerToolbar`；现有 props 和事件在整个 `0.x` 保持兼容。
3. `DocxEditor` 与 `DocxSurface` 共用同一个 `DocxDocumentSurface` 页面渲染树。
4. 统一使用 `buildDocumentPageNodeSegments()` 产出的 `DocumentPageNodeSegment[][]` 作为分页事实来源。
5. 页面渲染必须消费 `paragraphLineRange`、`tableRowRange` 和后续可能增加的分段字段。
6. 外壳主题与纸面主题分开。默认纸面始终为白色；深色纸面必须显式请求。
7. `zoom` 使用倍率：`1 = 100%`。公开组件内部不得混用百分比和倍率。
8. 页面测量只使用未缩放高度；缩放只在布局投影时应用一次。
9. 批注 gutter 不改变纸张宽度；它属于页面旁栏，必须进入每页横向占用计算。
10. 搜索、缩放、当前页和批注显示状态可由宿主受控。

## 4. 目标结构

```text
DocxViewer
  ├─ DocxViewerToolbar               默认完整工具栏
  └─ DocxSurface                     稳定受控文档面
       └─ DocxDocumentSurface
            └─ DocxViewerRoot
                 ├─ useDocxPaginationPlan
                 │    └─ buildDocumentPageNodeSegments
                 ├─ DocxPageSurface
                 │    ├─ DocxPageBody
                 │    │    ├─ ParagraphSegmentHost
                 │    │    └─ DocxTableHost(rowRange)
                 │    └─ DocxTrackedChangeGutter
                 └─ page measurement / virtualization

DocxEditor
  ├─ 自定义或默认编辑工具栏
  └─ 同一个 DocxDocumentSurface
       └─ 额外注入 controller + editable interactions
```

### 状态归属

| 状态 | Owner | 规则 |
|---|---|---|
| DOCX bytes / model 加载 | `DocxViewer` 或宿主 | `DocxSurface` 只接收已解析 `model` |
| 分页 plan | `DocxDocumentSurface` | model、layout options、测量变化时重算 |
| zoom | 宿主或 `DocxViewer` | `DocxSurface` 只消费受控倍率 |
| searchQuery / active result | 宿主或 `DocxViewer` | `DocxSurface` 不维护第二份查询状态 |
| current page / total pages | `DocxSurface` 计算并发事件 | 宿主只展示，不反向猜测 DOM |
| viewer shell theme | 宿主 | 控制外围背景和控件颜色 |
| page theme | `DocxSurface` | 默认 `paper`，显式选择深色纸面 |
| comments / tracked changes | 宿主或 `DocxViewer` | 受控 boolean，无隐式 `null` 分支 |
| Worker / WASM | `DocxRuntime` | 与 Surface 生命周期分离 |

## 5. `DocxSurface` 公共合同

### 5.1 Props

```ts
export type DocxPageTheme = "paper" | "dark"

export interface DocxSurfaceProps {
  /** Required parsed document model. */
  model: DocModel

  /** Optional editor controller. Omit for pure read-only rendering. */
  controller?: DocxEditorController

  /** Enables editor listeners and editable hosts. Defaults to false. */
  editable?: boolean

  /** Layout override. Missing values come from document section metadata. */
  layoutOptions?: LayoutOptions

  /** Viewer shell theme. Does not change paper colors. */
  viewerTheme?: "light" | "dark"

  /** Page surface theme. Defaults to "paper". */
  pageTheme?: DocxPageTheme

  /** Controlled zoom ratio. 1 = 100%. Defaults to 1. */
  zoom?: number

  /** Controlled full-document search query. Defaults to undefined. */
  searchQuery?: string

  /** Zero-based active match index. undefined means no active match. */
  activeSearchMatchIndex?: number

  /** Controlled annotation visibility. Defaults to false. */
  showTrackedChanges?: boolean
  showComments?: boolean

  /** Optional class applied to the outer scroll container. */
  className?: string
}
```

Contract：

- `model` 必填；缺失由 Vue 类型层阻止。
- `zoom` 必须是有限数值且 `0.25 <= zoom <= 4`。无效值 fail fast，抛出 `RangeError`，不静默归一。
- `searchQuery` 缺失使用 `undefined`；空字符串表示清除搜索。
- `viewerTheme="dark"` 只改变页面外围背景，不改变纸面和正文颜色。
- `pageTheme="paper"` 固定白色纸面和深色正文。
- `showTrackedChanges`、`showComments` 默认 `false`。

### 5.2 事件

```ts
export interface DocxVisiblePageRange {
  startPageIndex: number
  endPageIndex: number
}

export interface DocxSearchState {
  query: string
  matchCount: number
  activeMatchIndex?: number
}

update:activeSearchMatchIndex(index: number | undefined): void
pageCountChange(count: number): void
visiblePageRangeChange(range: DocxVisiblePageRange): void
searchStateChange(state: DocxSearchState): void
layoutError(error: DocxLayoutError): void
```

事件规则：

- 页码从 UI 展示时使用 1-based；公开 range 与 scroll API 使用 0-based。
- `pageCountChange` 只在值变化时触发。
- 空查询返回 `matchCount: 0`、`activeMatchIndex: undefined`。
- `searchNext()` / `searchPrevious()` 计算目标后发 `update:activeSearchMatchIndex`；宿主回传 prop 后才改变 active 高亮。未受控时 `DocxViewer` 自己接收该事件并更新内部状态。
- 未知节点、无法分页的超高对象或布局不一致必须通过 `layoutError` 返回稳定 `code`，不能仅写 console。

### 5.3 Expose

```ts
export interface DocxSurfaceExpose {
  scrollToPage(pageIndex: number): void
  scrollToNode(nodeIndex: number): void
  scrollToSearchMatch(matchIndex: number): void
  searchNext(): void
  searchPrevious(): void
  getScrollContainer(): HTMLElement | undefined
}
```

错误规则：

- page / node / match index 非整数或越界时抛出 `RangeError`。
- scroll container 尚未挂载时返回 `undefined`，不返回占位 DOM。
- 空搜索状态调用 next / previous 时不抛错，保持当前滚动位置。

## 6. `DocxViewer` 兼容合同

现有 `DocxViewer` 继续接受：

- `file` / `model`
- `runtime`
- `showToolbar`
- `defaultThumbnailsOpen`
- `defaultZoom`
- `isDark`
- `defaultShowTrackedChanges`
- `defaultShowComments`
- `showTrackedChanges`
- `showComments`

兼容映射：

| 旧属性 | 内部映射 |
|---|---|
| `defaultZoom: 100` | `DocxSurface.zoom = 1` |
| `isDark: true` | `viewerTheme="dark"`；`pageTheme` 仍为 `paper` |
| `defaultShowTrackedChanges` | Viewer 内部初始状态；整个 `0.x` 保持当前默认值 `true` |
| `showTrackedChanges` | 受控传给 Surface |
| `defaultShowComments` | Viewer 内部初始状态；整个 `0.x` 保持当前默认值 `true` |
| `showComments` | 受控传给 Surface |

新增受控 props：

```ts
zoom?: number
searchQuery?: string
activeSearchMatchIndex?: number
pageTheme?: DocxPageTheme
```

优先级：受控属性优先于 `default*`。受控属性传入后，内部 toolbar 只能发出 `update:*`，不能自行覆盖受控值。

新增事件：

```ts
update:zoom
update:searchQuery
update:activeSearchMatchIndex
page-count-change
visible-page-range-change
search-state-change
```

### 6.1 受控状态矩阵

`DocxSurface` 本身不拥有 toolbar，也不改变受控 props。所有受控值遵循单向数据流：宿主传值 → Surface 渲染；Surface 只发状态事件，不自行写回宿主。

| 属性 | Surface 内部是否修改 | 发出的事件 | controller 关系 |
|---|---|---|---|
| `zoom` | 否 | 无；宿主自己的按钮直接更新 prop | 与 controller 无关 |
| `searchQuery` | 否 | `searchStateChange` | 与 controller 无关 |
| `activeSearchMatchIndex` | 否 | `update:activeSearchMatchIndex` + `searchStateChange` | 与 controller 无关 |
| `showTrackedChanges` | 否 | 无 | controller 有同名状态时，显式 prop 优先 |
| `showComments` | 否 | 无 | controller 有同名状态时，显式 prop 优先 |
| `editable` | 否 | 无 | `true` 时 controller 必填；缺失抛 `TypeError` |
| `currentPage` | Surface 计算 | `visiblePageRangeChange` | controller 可镜像显示，不能覆盖 Surface 计算 |
| `pageCount` | Surface 计算 | `pageCountChange` | controller 可镜像显示，不能提供 fallback |

`DocxViewer` 可以继续拥有内部 toolbar 状态。它将内部状态映射为 `DocxSurface` 的受控 props；新增受控 prop 存在时，Viewer 内部 toolbar 只发 `update:*`，等待调用方回传新值。

兼容约束：

- 现有 `DocxViewer` 在整个 `0.x` 继续默认显示批注和修订；新 `DocxSurface` 默认隐藏。
- 现有 `DocxViewer.defaultZoom` 继续使用百分比；新 `DocxSurface.zoom` 使用倍率。
- `DocxViewer` 内部负责百分比与倍率的唯一转换，消费端不做二次转换。

## 7. 分页实现

### 7.1 单一分页入口

新增 `useDocxPaginationPlan()`，由 `DocxDocumentSurface` 使用：

```ts
export interface DocxPaginationPlan {
  pages: DocumentPageNodeSegment[][]
  pageLayouts: ResolvedPageLayout[]
  version: number
}
```

输入：

```ts
model
layoutOptions
section metrics
measured paragraph line metrics
measured table row heights
measured page content heights
```

内部调用 `buildDocumentPageNodeSegments()`；`layout/index.ts:buildLayoutSnapshot()` 不再承担 Vue 页面分页事实来源，只保留兼容与静态布局用途。

### 7.2 段落跨页渲染

扩展现有 `DocxParagraphHost`，不新建第二套段落渲染器：

```ts
paragraphLineRange?: ParagraphLineRange
```

接线边界：

```text
DocxViewerRoot.pageNodeSegments
  -> DocxPageSurface.pageNodeSegments
  -> DocxPageBody segment
  -> DocxParagraphHost.paragraphLineRange
  -> renderParagraphRuns（同一 run 渲染链）
```

实现使用 `ParagraphLineRange` 对应的文本 offset 范围裁剪现有 run tree；offset 由和分页测量相同的 pretext line fragments 生成。禁止在渲染时用 CSS `line-clamp`、固定高度 clip 或重新调用不同换行算法，这些方式会造成搜索、超链接、修订和字符边界漂移。

必须满足：

- 第一段保留段前距；续页段不重复段前距。
- 最后一段保留段后距；非末段不重复段后距。
- 行号区间为 `[startLineIndex, endLineIndex)`。
- 上下页不能重复或遗漏字符。
- 同一文本 offset 分片同时用于显示、搜索命中、修订和点击定位。
- 字体 fallback、中文换行、列表编号、超链接和修订标记保持一致。
- `widowControl`、`keepLines`、`keepNext` 使用分页器已计算的 segment，不在渲染层重新决定分页。
- `paragraphLineRange` 越界或对应不到 line fragment 时抛 `INVALID_SEGMENT`，禁止回退渲染完整段落。

### 7.3 表格跨页渲染

扩展现有 `DocxTableHost`，不建立新的表格 surface：

```ts
rowRange?: TableRowRange
tableRowSlice?: TableRowSlice
```

接线边界：

```text
DocxViewerRoot.pageNodeSegments
  -> DocxPageSurface.pageNodeSegments
  -> DocxPageBody segment
  -> DocxTableHost.rowRange + tableRowSlice
```

规则：

- 只渲染 `[startRowIndex, endRowIndex)`。
- 重复表头按 DOCX 表属性决定。
- 跨页行切片使用 `tableRowSlice.startOffsetPx` 和 `sliceHeightPx` 投影同一行内容，不复制整行。
- 合并单元格跨页时保持列宽和边框连续。
- `tableRowRange` / `tableRowSlice` 无效时抛 `INVALID_SEGMENT`，禁止回退整表。
- Viewer 和 Editor 共享这一实现；编辑模式的选区位置继续使用模型 row/cell/paragraph identity。

### 7.4 页面测量与缩放

统一约束：

```text
unscaledPageHeight = page layout / content measurement result
visualPageHeight   = unscaledPageHeight * zoom
pageOffset         = previous visual heights + fixed visual gaps
```

实现规则：

- `DocxPageSurface.measure` 上报未缩放布局高度，优先读取 `scrollHeight` / untransformed box。
- page height cache 只存未缩放值。
- `zoom` 只在 `visualPageHeight` 和 CSS transform 中各表达同一投影，不把 transform 后测量值写回 cache。
- `DOC_PAGE_BREAK_GAP` 是视觉像素，默认不随 zoom 放大。
- zoom 改变时清除依赖视觉尺寸的 range cache，但不重新解析 DOCX。

### 7.5 批注 gutter 布局

```text
                    16px     240px
┌──────── paper ────────┐  ┌─ gutter ─┐
│                       │  │ comments │
└───────────────────────┘  └──────────┘
```

规则：

- gutter 在每页外部，不算进纸张宽度。
- 页面横向 footprint：`paperWidth * zoom + gutterGap + gutterWidth`。
- 没有可见批注/修订时 footprint 不预留 gutter。
- 页面居中以 footprint 为整体，避免 gutter 把纸张挤偏到右侧。
- 窄视口 `< 1100px` 时 gutter 进入纸张下方；此时页面高度测量必须包含下方 gutter，但纸张高度仍独立。

## 8. 主题与样式合同

```css
.docx-surface {
  --docx-viewport-bg: var(--app-background, #f4f4f5);
  --docx-paper-bg: #ffffff;
  --docx-paper-fg: #111827;
  --docx-paper-border: #d4d4d4;
  --docx-paper-shadow: 0 8px 24px rgba(0, 0, 0, 0.08);
  --docx-page-gap: 28px;
  --docx-gutter-width: 240px;
  --docx-gutter-gap: 16px;
}
```

`viewerTheme="dark"` 只重映射 `--docx-viewport-bg` 和外围控件 token。`pageTheme="paper"` 保持纸张 token。宿主可通过公开 CSS 变量定制外围和纸张，但无需覆盖组件内部 scoped class。

## 9. 错误模型

```ts
export type DocxLayoutErrorCode =
  | "INVALID_ZOOM"
  | "INVALID_SEGMENT"
  | "CONTROLLER_REQUIRED"
  | "PAGINATION_FAILED"
  | "PAGE_LIMIT_EXCEEDED"
  | "UNSPLITTABLE_CONTENT"
  | "MEASUREMENT_FAILED"

export interface DocxLayoutError extends Error {
  code: DocxLayoutErrorCode
  pageIndex?: number
  nodeIndex?: number
  actual?: number
  allowed?: number
}
```

- 缺失 required model、unknown segment variant、invalid line/row range 直接失败。
- `editable=true` 且 controller 缺失时抛 `TypeError`，同时映射一次 `layoutError({ code: "CONTROLLER_REQUIRED" })`；组件不挂载编辑监听器。
- `activeSearchMatchIndex` 越界属于宿主调用错误：渲染时发 `searchStateChange`，其中 `activeMatchIndex: undefined`；Expose 的 `scrollToSearchMatch()` 直接抛 `RangeError`。
- `scrollToPage()`、`scrollToNode()` 和 `scrollToSearchMatch()` 的参数错误由同步 `RangeError` 表达，不重复发 `layoutError`。
- `layoutError` 只报告异步分页、segment、资源限制和测量失败；同一失败每个 pagination version 只发一次。
- 不允许用空页面、空 segment 或完整节点 fallback 掩盖分页 contract mismatch。
- `UNSPLITTABLE_CONTENT` 可显示受控降级 UI，但必须带 node index 和明确提示。

## 10. 实现顺序

| 阶段 | 修改范围 | 产出 |
|---|---|---|
| 1. 回归测试 | `tests/unit`、`tests/component`、`tests/blackbox` | 先复现深色纸面、默认 gutter、50%/150% 页面偏移、长段落跨页裁切 |
| 2. 分页事实来源 | `docx-core` + `DocxViewerRoot` | `buildDocumentPageNodeSegments` 接入，segment 字段完整传递 |
| 3. 段落/表格切片 | `DocxPageBody`、Paragraph/Table host | 跨页无重复、无遗漏、无覆盖 |
| 4. 测量/缩放 | `DocxViewerRoot`、`DocxPageSurface` | unscaled cache + 单次视觉缩放 |
| 5. 主题/gutter | 页面样式与 gutter footprint | 深色外围 + 白纸；批注布局正确 |
| 6. `DocxSurface` | 新组件 + public export | 受控 zoom/search/page/annotations |
| 7. `DocxViewer` 迁移 | Viewer 组合 Surface | 旧 API 保持兼容 |
| 8. Demo/文档/发布 | demo、API docs、release notes | 公开用法和发布包消费验证 |

## 11. 验收标准

### 11.1 视觉

固定环境：Chromium、1440×900、device scale factor 1。

| 用例 | 预期 |
|---|---|
| 深色宿主 + `viewerTheme="dark"` + `pageTheme="paper"` | 外围深色，纸张白色，正文深色 |
| `showTrackedChanges=false`、`showComments=false` | 无 gutter，纸张居中 |
| annotations 开启 | gutter 完整可见，不覆盖纸张，不产生错误居中 |
| 多页文档 | 相邻页面保持 28px 间距，无贴合和重叠 |

### 11.2 分页

至少包含以下真实 DOCX fixture：

- 单段超过一页；
- 中英文混排跨页；
- 长表格跨页；
- 合并单元格跨页；
- 页眉、页脚、脚注；
- 批注和修订；
- 50%、100%、150%、200% zoom。

断言：

- 每个正文字符在页面集合中出现且只出现一次（字段动态值另设预期）。
- 页尾与下一页首没有字符重复或遗漏。
- 每个 page wrapper 的视觉区间互不相交。
- `next.top - current.bottom === pageGap`，允许 1px 取整误差。
- 改变 zoom 不改变 page count，不重新请求 Worker/WASM。

### 11.3 受控交互

- 外部修改 `zoom` 立即更新视图和 `update:zoom` contract。
- 外部修改 `searchQuery` 更新命中数和高亮。
- `searchNext()` / `searchPrevious()` 循环导航并滚动到命中页。
- `scrollToPage()` 能定位首、中、末页；越界抛 `RangeError`。
- 隐藏默认 toolbar 后，全部功能仍可由宿主 toolbar 操作。

### 11.4 错误与事件

- 无效 zoom 触发 `INVALID_ZOOM`，视图不写入默认 zoom。
- invalid segment 触发一次 `layoutError(INVALID_SEGMENT)`，并且不渲染完整节点 fallback。
- `editable=true` 且 controller 缺失时同步失败并发一次 `CONTROLLER_REQUIRED`。
- `pageCountChange` 只在 page count 变化时触发；滚动和 zoom 不产生重复同值事件。
- `visiblePageRangeChange` 只在 start / end 任一变化时触发。
- 空搜索发 `matchCount=0`、`activeMatchIndex=undefined`；next / previous 保持滚动位置。
- Expose 越界使用 `RangeError`，不重复发 `layoutError`。

### 11.5 兼容与发布

- 现有 `DocxViewer` 示例无需修改即可加载。
- `DocxViewer` 在整个 `0.x` 保持现有批注、修订和 `defaultZoom` 默认行为。
- 现有公开低层组件在整个 `0.x` 保留 `@deprecated`。
- `@arcships/vue-docx` 真实压缩包导出 `DocxSurface` 类型、实现和 `style.css`。
- 工作区外 Vue + Vite 项目使用发布压缩包，Worker/WASM、分页、搜索和缩放全部通过。
- 控制台无 Vue warning、未处理 error 和关键资源 404。

## 12. 必须新增的测试

| 测试 | 证明内容 |
|---|---|
| `docx-surface-theme.test.mjs` | shell theme 与 page theme 分离 |
| `docx-surface-controlled-state.test.mjs` | zoom/search/page/annotations 单向受控合同；editable/controller 组合 |
| `docx-surface-events-errors.test.mjs` | 事件去重、空搜索、RangeError 与 layoutError 分界 |
| `docx-pagination-segments.test.mjs` | `paragraphLineRange` / `tableRowRange` / `tableRowSlice` 不丢失；invalid segment fail fast |
| `docx-pagination-visual-height.test.mjs` | transform 后高度不二次缩放 |
| `docx-long-paragraph.blackbox.py` | 长段落跨页无覆盖、重复、遗漏 |
| `docx-annotation-footprint.blackbox.py` | gutter 宽度、居中和窄视口回流 |
| 外部压缩包消费测试 | `DocxSurface`、Worker、WASM、CSS 公共出口可用 |

这些测试必须读取用户可见结果和公开 API。禁止仅搜索源码字段或断言 DOM 节点存在。