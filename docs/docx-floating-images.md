# DOCX 浮动图片：定位、层级、绕排与限制

本文记录 `@arcships/vue-docx` 中 DOCX 浮动图片（`wp:anchor`）的渲染模型，以及当前
实现有意保留的近似行为。

## 定位基准（page origin）

`.docx-page-surface` 是页面内所有绝对定位元素的定位祖先，它的 padding box 从纸张
边缘开始（页边距用 padding 表达）。因此：

- `horizontalRelativeTo/verticalRelativeTo = "page"` 直接以纸张左上角为原点；
- `"margin"` / `"column"` 以页边距内边缘为原点（`marginLeft/marginTop`）；
- `"paragraph"` / `"line"` 以段落原点为原点。

基准对象由 `resolveFloatingPageOriginPx(pageWidthPx, marginsPx)` 生成，经
`DocxPageSurface → DocxPageBody/DocxPageHeader/DocxPageFooter → DocxParagraphHost /
DocxTableHost → renderParagraphRuns` 传递。缺少基准时 `absoluteFloatingImageStyle`
会把 `margin|page|column` 锚点退化到纸张 (0, 0)，图片整体偏移一个页边距。

锚点基准只有一处实现：`resolveFloatingAnchorOriginPx(floating, origin)` 按
`wp:anchor` 的 `relativeFrom` 把模型偏移换算成页面坐标，`absoluteFloatingImageStyle`
（只读视图）与编辑模式的 `DocxImageLayer` 都调用它。编辑模式的图片层是
`.docx-page-surface` 的绝对定位子元素，坐标原点同样是纸张边缘，因此
`DocxPageSurface` 通过 `floatingPageOriginPx` 把同一个基准对象下传；`margin` /
`column` 锚点不再少加一个页边距（issue #194）。`DocxImageLayer` 单独挂载且未传基准
时退化到纸张 (0, 0)，与旧行为一致。

## 层级（z-order）

- 只读视图：浮动图片随段落渲染，`behindDocument` 使用负 z-index（`-100000`），
  位于文字之下；其余按 `zIndex` 排序，位于文字之上。
- 编辑模式：段落宿主跳过浮动图片，由 `DocxImageLayer` 分成两层绘制：
  - `data-docx-image-layer="behind"`（`behindDocument`，z-index `-1`）：在页面文字的
    下方、页面背景的上方。`.docx-page-surface` 建立独立的 stacking context
    （`isolation: isolate`），负层级不会跑到纸张背景后面。这一层不接收指针事件，
    点击穿透给正文文字。
  - `data-docx-image-layer="front"`（其余浮动图片，z-index `20`）：在文字上方，带
    拖动/缩放热区。

## 旋转

`a:xfrm/@rot` 为 90°/270° 时，图片的视觉占位宽高互换。`resolveImageFrameSizePx`
返回互换后的外框尺寸，`resolveImageRenderTransformStyle(..., { fillFrame: true })`
返回未旋转的内容盒加旋转位移，图片内容旋转后正好填满外框。未旋转或任意角度旋转
仍按中心旋转，不改占位。

## 绕排（wrap）

- `wrapText = left|right`：Word 语义是「文字只放在图片的哪一侧」，因此图片浮到另一
  侧。`resolveWrappedFloatingSide` 与 `resolveWrappedFloatingTextBands` 是渲染和
  文本测量的唯一判定点，两侧共用同一结论，避免行数/分页漂移。
- `wrapText = bothSides|largest`：CSS float 只能让文字绕单侧，退化为
  `resolveWrappedFloatingSide` 判定的那一侧。
- `topAndBottom`：块级居中/靠边并 `clear: both`，独占整行。

### `tight` / `through` 目前按 `square` 渲染

`a:wrapTight` / `a:wrapThrough` 的 `wrapPolygon` 未建模：`shouldRenderWrappedFloatingImage`
把它们当作 `square`，使用矩形排他区。文字因此按矩形绕排，而不是沿多边形轮廓。

### 跨段落排他区

页面正文是「每段一个 flex item」的模型，CSS float 无法缩短后续段落的行盒。渲染层
用 `resolveWrappedImageFlowClearancePx` 计算「浮动图片超出锚点段落」的高度，并把该
高度作为下一段的上外边距（`wrapClearanceTopPx`），使后续段落从图片下方开始。

已知限制：

- 只做垂直避让：后续段落落在图片下方，而不是贴着图片旁边的空带（宽度未按绕排带
  收窄）。
- 只对整段分页的页面生效；该页存在被分页切分的段落或表格时不启用。
- `behindDocument` 图片与 `topAndBottom` 图片不参与避让（前者本就在文字之下，后者
  已随段落增高）。

相关代码：`packages/docx-core/src/editor/helpers/page-flow-wrap-clearance.ts`。
