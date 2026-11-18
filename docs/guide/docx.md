# DOCX 使用指南

`DocxViewer` 支持 `.docx`、`.docm`、`.dotx` 和 `.dotm`。宏文档和宏模板只用于安全预览，VBA 不会执行；模板按普通文档打开。`DocxEditor` 的原格式编辑和宏保留不在本次支持范围内。

## 查看文档

```vue
<script setup lang="ts">
import { DocxViewer } from "@arcships/vue-docx"
import "@arcships/vue-docx/style.css"
</script>

<template>
  <DocxViewer
    :file="bytes"
    file-name="contract.docx"
    :show-toolbar="true"
    @load-success="onLoaded"
    @load-error="onError"
  />
</template>
```

`DocxViewer` 可以接收 `file: ArrayBuffer`，也可以直接接收已经解析的 `model: DocModel`。两者同时提供时优先使用 `model`。

常用属性：

| 属性 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `file` | `ArrayBuffer` | 无 | DOCX、DOCM、DOTX 或 DOTM 字节 |
| `model` | `DocModel` | 无 | 已解析文档模型 |
| `runtime` | `DocxRuntime` | 组件创建 | 调用方持有的 Runtime |
| `showToolbar` | `boolean` | `true` | 显示查看工具栏 |
| `showUpload` | `boolean` | `true` | 上传入口，含工具栏按钮与空态、错误态的选择文件按钮 |
| `showDownload` | `boolean` | `true` | 工具栏下载按钮 |
| `defaultThumbnailsOpen` | `boolean` | `false` | 默认展开缩略图 |
| `defaultZoom` | `number` | `100` | 初始缩放百分比 |
| `isDark` | `boolean` | `false` | 深色界面 |
| `showTrackedChanges` | `boolean \| null` | `null` | 受控显示修订 |
| `showComments` | `boolean \| null` | `null` | 受控显示批注 |

主要事件为 `load-start`、`load-success`、`load-error`、`file-selected`、`update:isDark`、`update:showTrackedChanges` 和 `update:showComments`。

## SVG 图片

DOCX 中的 `image/svg+xml` 图片按矢量渲染，支持直接引用 SVG 的 `a:blip` 和 Word 常见的 `asvg:svgBlip` + PNG 回退结构（此时取 SVG）。图片保持 `wp:extent` 的锚定尺寸，缩放或 200% 预览时仍按显示分辨率重新栅格化，不会被替换成位图。

渲染前会对 SVG 标记做安全清洗：移除 `<script>`、`on*` 事件属性、`javascript:` 与所有外部引用；保留 `#id` 内部引用和 `data:image/` 内嵌图片，因此渐变、裁剪路径和 `<use>` 仍正常绘制。使用 `<foreignObject>` 或滤镜（`<filter>` / `filter=`）的图片无法用 `<img>` 管线绘制，会回退到占位符并在图片框内显示 `SVG`（小图标显示 `s`）。标记无法解析或超过 `MAX_SVG_PICTURE_MARKUP_CHARS`（4 MiB）时同样回退。

```ts
import {
  inspectSvgPicture,
  resolveRenderableImageSource,
  svgPictureUnsupportedDiagnostic,
} from "@arcships/docx-core"

const inspection = inspectSvgPicture(imageRun)
if (inspection && !inspection.supported) {
  // { code: "IMAGE_UNSUPPORTED_FEATURE", phase: "image-render", reason, partName }
  console.warn(svgPictureUnsupportedDiagnostic(imageRun))
}
```

`renderParagraphRuns` 的 `onImageUnsupported` 选项会为每个回退的 SVG 回调一次，诊断带上 `imageKey` 和本地化消息；宿主可用 `imageUnsupportedText` 覆盖消息模板（支持 `{feature}` 占位符）。原始字节保留在模型和基础包中：编辑后导出并重新打开，SVG 媒体部件逐字节不变；缩略图不会降采样 SVG，而是让最终整页栅格化负责缩放。

## 键盘导航

`DocxViewer` 根节点接收键盘事件：`PageDown`/`PageUp` 翻一页，`Home`/`End` 跳到第一页/最后一页，`Ctrl/Cmd` + `+`/`=`、`-`、`0` 按缩放档位放大、缩小或重置到 100%，`Ctrl/Cmd + F` 聚焦工具栏搜索框，`Escape` 关闭搜索。焦点在页码或搜索输入框内时，查看器不会抢这些按键；快捷键也不会从同级查看器串入。根节点带 `:focus-visible` 焦点环。完整映射见[查看器键盘导航](viewer-keyboard.md)。

## 编辑文档

```vue
<DocxEditor
  :file="bytes"
  :editable="true"
  :show-toolbar="true"
  :show-thumbnails="true"
  @load-error="onError"
/>
```

`DocxEditor` 可以自己创建编辑控制器，也可以通过 `editor` 传入 `useDocxEditor()` 返回的控制器。传入控制器适合把保存、撤销重做和业务按钮放到组件外部。

编辑器和查看器一样接受 `isDark` 和 `minHeight`。传入 `isDark` 时外壳、工具栏、缩略图和纸张一起切换为深色；不传时跟随 `controller.documentTheme`。`minHeight` 可选，数字按 px；不传时组件不设置最小高度。打印时工具栏、缩略图和状态栏自动隐藏，为每页插入分页，并按文档页面尺寸设置 `@page`。

## 编辑辅助操作

`DocxEditor` 默认启用下列编辑辅助操作。不需要时用属性关闭；关闭后不会渲染对应界面，也不改变文档或控制器行为。

| 能力 | 交互 | 关闭方式 | 对应控制器 API |
|---|---|---|---|
| 表单域属性对话框 | 双击表单域打开，编辑值、类型、名称、占位提示、必填、最大长度和下拉选项；保存为一次撤销 | `:show-form-field-dialog="false"` | `updateFormField()` |
| 链接悬浮卡片 | 鼠标悬停链接显示打开/编辑/删除；编辑和删除会先选中链接文字 | `:show-link-card="false"` | `setLink()`；打开时发出 `link-open` 事件 |
| 表格行高拖拽 | 悬停或选中表格后拖动行边界；提交 `w:trHeight`（`atLeast`） | `:show-row-resize-handles="false"`（`DocxTableHost`） | `setTableRowHeight()` |
| 段落样式悬浮预览 | 悬停工具栏样式选择器时按各样式自身的字体/字号/缩进渲染示例，点击即应用 | `:show-style-preview="false"` | `setParagraphStyle()` / `setHeading()` |
| 段落/表格拖拽排序 | 悬停正文段落或表格显示拖拽手柄，拖动到目标位置后插入；手柄支持 `↑`/`↓`，`Esc` 取消 | `:show-node-reorder="false"` | `moveBodyNode()` |

拖拽排序一次移动只产生一条撤销记录：落在节点自身位置（前或后）视为无操作，不写入历史；跨分节拖动时，分节起始位置跟随对应节点，而不是由被移动的节点重新定义分节。

表单域的值只有在符合字段约束时才会写入模型：下拉值必须来自选项列表，日期必须是有效日期，文本遵守最大长度；被拒绝的保存不会修改文档，撤销栈也不增加。

链接卡片对外部地址调用 `window.open(..., "noopener,noreferrer")`；文档内部书签（`#name`）不会被当作 URL 打开，而是通过 `link-open` 事件交给宿主处理，便于宿主滚动到自己的书签位置。

## 文档块、词与高亮

`@arcships/docx-core` 把已解析的文档模型拆成可寻址的文档块和词：表格、单元格、页眉页脚、脚注尾注、图片也都包含在块树里。每个块带一个稳定的 source id，查看器把同一个 id 渲染为 `data-docx-source-id`，因此几何解析不需要按文字猜位置。

```ts
import {
  extractDocumentBlocks,
  flattenDocumentBlocks,
  resolveDocxBlockGeometries,
  resolveDocxWordGeometries,
} from "@arcships/docx-core"

const document = await extractDocumentBlocks(model, {
  documentId: "contract.docx",
  includeWords: true,
})
const blocks = flattenDocumentBlocks(document.blocks)
const words = blocks.flatMap((block) => block.words ?? [])

const root = viewerRoot.value // 包含 [data-docx-page-surface] 的查看器根元素
const blockGeometry = resolveDocxBlockGeometries(root, blocks)
const wordGeometry = resolveDocxWordGeometries(root, words)
```

块类型沿用 Extend parser 的 15 个名称（`DOCUMENT_BLOCK_TYPES`）：段落 `text`、`heading`、`section_heading`、`legend`、列表段落（`details.list`）、`figure`、`table`/`table_head`/`table_cell`、表单域 `key_value`、页码字段 `page_number`，以及 `header`/`footer` 容器。空段落和空单元格按上游约定省略，空段落里的图片会提升为同级 figure，source id 仍保留原段落路径。

词分割对空白分隔的文本保留标点，对中日韩文字按字符切分；每个词用 UTF-16 的 `sourceRefs[].start/end` 指回段落，id 形如 `<blockId>/word/{start}-{end}`。几何解析把 DOM 文本范围换算成页面 CSS 像素：页面表面的 client rect 除以其 CSS 尺寸，因此 100%、200% 缩放以及横版/旋转纸张都使用同一套页面坐标；跨页切分的段落每个挂载页各返回一个 fragment。

`@arcships/vue-docx` 提供两个覆盖层组件，几何盒子通过 Teleport 放进对应页面表面，和页面一起缩放：

```vue
<DocxBlockHighlights
  :root="viewerRoot"
  :blocks="blocks"
  :block-ids="selectedBlockIds"
  :page-index="0"
  color="#2563eb"
  @geometry="onBlockGeometry"
/>
<DocxWordHighlights
  :root="viewerRoot"
  :word-ranges="[{ blockId: paragraphId, start: 6, end: 10 }]"
  color="#f59e0b"
/>
```

`blocks`/`block` 传入提取结果；`blockIds` 可以只按 id 选择（没有块列表时 id 直接当 source id）；`pageIndex` 只画某一页；`wordRanges` 不需要先运行词提取。盒子带 `data-docx-block-highlight`、`data-docx-word-highlight`、`data-docx-block-type` 和 `data-docx-highlight-active`，便于宿主样式和测试定位。

`DocxDocumentSurface` 暴露 `getDocumentBlocks()`、`confirmBlockReference()` 和 `confirmWordReference()`：块走结构化定位（`structure`），词走 `text-range` 定位，两者都通过统一引用契约发出 `reference-confirm`，`resolveReference()` 会按当前修订返回 `exact`、`relocated` 或明确失败。

## 直接使用核心能力

纯数据操作从 `@arcships/docx-core/core` 导入；需要加载文件、Worker 和 WASM 时使用 `@arcships/docx-core/runtime`。

```ts
import { createDocxRuntime } from "@arcships/docx-core/runtime"

const runtime = createDocxRuntime()
try {
  const result = await runtime.loadSource({ kind: "bytes", bytes })
  console.log(result.model)
} finally {
  runtime.dispose()
}
```

不要直接组合已经标记弃用的页面、段落、表格和工具栏组件。它们只为 `0.x` 兼容保留，新代码使用 `DocxViewer` 或 `DocxEditor`。
