# 文档 blocks：#164 Extend 对齐说明

> 适用批次：0.8（#164）。参考固定提交：`extend-hq/ui@70fe70c6f1e1eee53954c6923f3d9f583557a57e`。
> 像素门禁：`pnpm ui:parity -- --only bounding-box-citations,layout-blocks,e-signature,document-splits`。

## 交付范围

| Issue 条目 | Vue 入口 | 状态 | 像素对照 |
|---|---|---|---|
| `BoundingBoxCitations` | `packages/vue-ui/src/components/BoundingBoxCitations.vue` | 完成（Form 标量字段） | `bounding-box-citations`，8 个状态 |
| `LayoutBlocks` | `packages/vue-ui/src/components/LayoutBlocks.vue` | 完成（OCR 面板、叠加、markdown） | `layout-blocks`，7 个状态 |
| `SignaturePad` / E-Signature | `SignaturePad.vue` + `ESignatureFields.vue` | 完成（手写/输入、清除、确认、预览） | `e-signature`，6 个状态 |
| `DocumentSplits` | `packages/vue-ui/src/components/DocumentSplits.vue` | 完成（区间、缩略图、增删、拖放/键盘） | `document-splits`，8 个状态 |
| `FileSystem` | — | **延期**（见下文证据） | — |
| `SchemaBuilder` | — | **延期**（见下文证据） | — |

四个已完成 block 的参考实现分别在 `tools/extend-reference/src/vendor/extend/`：
`bounding-box-citations.tsx`、`layout-blocks-panel.tsx`、`signature-fields.tsx`、
`document-splits.tsx`。Vue 对照页在 `tools/vue-reference/src/surfaces/`，共享数据在
`tools/parity-fixtures/fixtures.json` 的 `blocks` 节点。

## 像素结果（2026-11-20）

```bash
PYTHON=/tmp/bb-venv/bin/python pnpm ui:parity \
  -- --only bounding-box-citations,layout-blocks,e-signature,document-splits \
  --skip-tokens --out output/ui-parity-164-final
```

- 232/232 case 通过；`diffPixels = 0`，几何 delta `0px`，无 `pass-with-AA`。
- 4 个 viewport（1440×900、1280×720、768×1024、390×844）× light/dark。
- 产物：`output/ui-parity-164-final/report.{json,html,md}`、`SUMMARY.md`、`cases/**`。

复现时的两个 harness 约定：

1. `tools/vue-reference/vite.config.ts` 设置 `cssMinify: false`。压缩器会把 `oklch()` /
   `color-mix()` 改成 8-bit hex，实测在浅色大面积填充上产生 1–2 通道差；保留原色彩函数才能与参考页
   的计算色一致。
2. `DocumentSplitsSurface` 两侧都会预热 `/samples/contract-page.png`。缩略图使用 `loading="lazy"`，
   预热避免离屏 case 的图片解码与截图竞态。

## 各 block 的 API 与状态

### BoundingBoxCitations

保留 0.8 之前的 `file` + `fields` 页卡模式（`field-click` 事件不变）。新增可选 review 模式：

```ts
interface BoundingBoxCitationsProps {
  file?: string
  fields?: BoundingBoxField[]
  reviewFields?: BoundingBoxReviewField[]
  pageImage?: string
  pageWidth?: number
  pageHeight?: number
  activeFieldKey?: string
  expectedValues?: Record<string, OfficeJsonValue>
  showExpected?: boolean
  readOnly?: boolean
  view?: "split" | "panel"
  loading?: boolean
  error?: string
  emptyMessage?: string
  className?: string
}
```

- 事件：`field-focus`、`update:activeFieldKey`、`update:expected`；legacy `field-click` 仅在后向模式使用。
- 键盘：卡片 `Tab` 可聚焦，`ArrowUp/ArrowDown` 切换活动字段，`Escape` 清除本地活动状态；
  `Form` / `JSON` 标签页可切换。
- 状态：空（`role=status`）、加载（`role=status`）、错误（`role=alert`）、已修改（琥珀色圆点 + 撤销）、
  `Expected` 只读/禁用。
- 页面高亮：`location.area` 为百分比，按 `left/top/width/height` 定位；`view="split"` 时左侧显示页图，
  也可以由宿主通过页图/PDF viewer 自行组合。
- 明确未移植：上游 JSON 标签页（`@pierre/diffs`）和数组/对象 `glide-data-grid` 编辑器。二者都不是
  文档 block 的可复用交互，而且会引入两套重量级运行时；Form 标量字段覆盖了 0.8 的引用审核场景。

### LayoutBlocks

```ts
interface LayoutBlocksProps {
  file?: string
  output?: ParsedOcrOutput
  blocks?: LayoutBlock[]
  pageImage?: string
  pageWidth?: number
  pageHeight?: number
  activeBlockId?: string
  view?: "split" | "panel"
  heading?: string
  emptyMessage?: string
  renderMarkdown?: boolean
  loading?: boolean
  error?: string
  ariaLabel?: string
  className?: string
}
```

- 旧 API（`file` + `output`）保持；新 `blocks` 支持 Extend 的 `heading/paragraph/list/table/figure/header/footer/page_number`。
- 面板：类型徽章、置信度、页码 chip、活动态；`ArrowUp/ArrowDown/Home/End` 移动焦点，鼠标进入/聚焦也会激活 block。
- `view="split"` 渲染页图/`#viewer` slot + 百分比叠加框；`pageWidth/pageHeight` 用于没有 `output` 尺寸时的几何。
- `renderMarkdown` 使用**先转义再格式化**的安全子集（粗体、斜体、行内代码、标题、列表、http(s) 链接），
  不执行宿主 HTML。
- 空/加载/错误状态均为 `role=status` / `role=alert`；`PDF overlay` 由 `#viewer` slot 接入任意 PDF/DOCX surface。

### SignaturePad / ESignatureFields

`SignaturePad`：

- 旧 API（`width`、`height`、`penColor`、`backgroundColor`、`update:signature`、`clear()`/`isEmpty()`/`toDataURL()`）保持。
- 新增 `mode`（`draw` / `typed`）、`modelValue`、`defaultValue`、`placeholder`、`showControls`、
  `showClear/showCancel/showConfirm`、`confirmDisabled`、`disabled`、`guide`、`loading`、`error`、`ariaLabel`。
- 事件：`update:signature`、`update:modelValue`、`clear`、`cancel`、`confirm`；确认按钮在空签名/禁用/错误时不可用。
- 移动端：触摸 `touchAction: none`、`pointercapture`、32px 触控目标；输入模式提供 28px 签名预览。

`ESignatureFields`（新增，对应上游 `SignatureFieldsPanel`）：

- `fields: ESignatureField[]`、`canExport`、`exporting`、`showPreview`、`title`、`description`、
  `loading`、`error`、`className`。
- 事件：`sign`、`edit`、`clear`、`export`；Signed/Unsigned 徽章、尺寸/页码、导出按钮禁用态。
- `showPreview` 默认 `false`，打开后在字段卡内显示签名图；默认关闭以保持与上游面板一致。

签名像素门禁覆盖字段面板（未签名/已签名/导出中/预览/加载/错误）。手写画布的笔迹是运行时输入，
不适合作为确定性截图基线，因此由 `tests/component/document-blocks.test.mjs` 覆盖指针笔迹、清除、
确认禁用和 typed 模式。

### DocumentSplits（新组件）

```ts
interface DocumentSplitsProps {
  splits: DocumentSplit[]
  thumbnailImages?: Record<string, string>
  withFrameDivider?: boolean
  hideAddSplit?: boolean
  disableRemove?: boolean
  loading?: boolean
  error?: string
  emptyMessage?: string
  className?: string
}
```

- 事件：`select-page`、`update:splits`、`splits-change`、`add-split`、`remove-split`。
- 拆分卡片：标题、页码范围、页数、删除（仅剩一个时禁用）、拖拽手柄、页缩略图（`FileThumbnail`）。
- 拖放：页可在组内/组间移动，拆分卡可重排；同时提供键盘替代：
  - 页按钮 `ArrowLeft/ArrowRight` 组内重排，`Ctrl/Cmd + ArrowLeft/ArrowRight` 跨组移动；
  - 拖拽手柄 `ArrowUp/ArrowDown` 重排拆分卡。
- 状态：无拆分（`role=status`）、空页组（“Drop pages here”）、加载、错误（`role=alert`）；
  缩略图缺失时 `FileThumbnail` 显示加载覆盖层。
- 默认不自动滚动宿主 PDF；`select-page` 交由宿主定位页面。

## FileSystem 暂缓

上游：`apps/v4/components/extend/file-system.tsx`，**5,381 行**。评估证据：

- 运行时依赖 `@pierre/trees`（`FileTree`、`useFileTree`、sprite sheet、file-tree 输入归一化）以及
  `Command`、`DropdownMenu`、`Popover`、`Dialog`、`ContextMenu` 等一整套 shadcn primitive；
  本仓库 0.8 依赖中没有 `@pierre/trees`，也没有 `command`/`context-menu` 的 Vue parity 层。
- 交互面包括 Finder 风格列表/网格切换、面包屑、排序、重命名/新建/删除、拖放移动、缩略图、
  空状态、DocumentPreviewCard 悬浮预览；这些都需要先有文件系统数据模型与宿主持久化契约。
- 0.8 的宿主（demo、viewer 组合）没有文件浏览入口；`FileUpload` / `FileThumbnail` 已覆盖上传与缩略图需求。
- 若后续启动，需要新增 `FileSystem.vue`、`DocumentPreviewCard.vue`、文件树数据模型与至少 2 个
  parity surface；预计工作量与 #163 同级，不适合在 #164 内静默塞入。

结论：**延期**。`docs/upstream/attribution.json` 与 `docs/upstream/extend-ui.md` 保留条目和本页链接；
不创建空壳组件，避免宿主误以为已支持。

## SchemaBuilder 暂缓

上游：`apps/v4/components/extend/schema-builder.tsx`，**2,703 行**。评估证据：

- 依赖 `@dnd-kit/core` + `@dnd-kit/sortable`（拖拽字段树）、`@pierre/diffs`（JSON 预览）、
  `next-themes`，以及完整的 shadcn 对话框/下拉/命令面板 primitive；本仓库没有这些运行时依赖。
- 功能面是抽取 schema 编辑器：字段/类型/嵌套编辑、枚举、必填、校验、排序、JSON 预览；
  目前产品没有 schema 编辑入口（没有宿主消费方）。
- 与 `LayoutBlocks`/`BoundingBoxCitations` 的 review 数据模型只有弱耦合；强行实现会引入
  尚未定义的 schema 持久化契约。
- 若后续启动，建议先定义 `ExtractionSchema` 数据模型与宿主保存/发布流程，再拆成独立的
  `SchemaBuilder.vue` + primitive 层；不在 #164 内做半成品。

结论：**延期**，登记为 `not-used`，并保留固定 commit 的上游路径以便后续复刻。

## 测试

- `tests/component/document-blocks.test.mjs`（21 个用例）：键盘焦点/激活、Tab 切换、Expected 编辑/撤销、
  空/加载/错误、markdown 转义、split 高亮、legacy API、typed/drawn 签名、字段面板事件、
  拆分增删/跨组移动/拖放/键盘重排。
- 运行：`node --test --test-concurrency=1 tests/component/document-blocks.test.mjs`。
