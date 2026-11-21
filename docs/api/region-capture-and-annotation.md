# 区域截图与标注

0.8.0 在 `@arcships/vue-ui` 中提供两个可选能力：

- `captureRegion({ surface, region, scale })`：把页面、幻灯片或工作表上的可见区域裁成位图，并返回采集来源元数据。
- `OfficeRegionAnnotator`：在截图上添加箭头、矩形、文字和高亮，并输出可序列化的 `OfficeRegionAnnotationDocument`。

两者都不属于 `OfficeObjectReference`。引用只保存区域；截图与标注由宿主决定是否保存、上传、脱敏或丢弃。没有 DOM 的 headless 宿主不需要引入组件，原有 `regionConfirm` / `regionDraftChange` 事件保持不变。

## captureRegion

```ts
import { captureRegion, releaseRegionCapture } from "@arcships/vue-ui"

const capture = await captureRegion({
  // 内置 Surface 都实现可选的 getCaptureSurface()
  surface: pdfSurfaceRef.value,
  // OfficeRegionDraftChange 的 region.rect，或 ManualRegionLocator
  region: { x: 0.1, y: 0.2, width: 0.5, height: 0.25 },
  target: { pageIndex: 2 },
  scale: window.devicePixelRatio,
  maxPixels: 24_000_000,
})

image.src = capture.url ?? URL.createObjectURL(capture.blob)
console.log(capture.source)
// { kind: "canvas", format: "pdf", space: "page", pageIndex: 2,
//   region: { ... }, scale: 2, width: 600, height: 300,
//   sourceWidth: 1200, sourceHeight: 600, capturedAt: "...", ... }

// 替换截图或卸载宿主时释放 object URL
releaseRegionCapture(capture)
```

`surface` 可以是：

- 普通 `Element`（按元素边界把 `region` 解释为归一化矩形）；
- 实现了 `getCaptureSurface(request)` 的组件实例或适配器；
- 直接提供 `element` / `raster` / `region` 的 provider。

`getCaptureSurface()` 返回：

```ts
interface OfficeRegionCaptureSource {
  /** region 的归一化坐标空间。 */
  element: Element
  /** 可选的 intrinsic 位图；提供时直接从它按像素裁剪。 */
  raster?: { source: CanvasImageSource; width: number; height: number }
  /** ManualRegionLocator 解析后的归一化矩形。 */
  region?: NormalizedRect
  format?: "docx" | "xlsx" | "pptx" | "pdf"
  documentId?: string
  revision?: string
  space?: "page" | "slide" | "sheet"
  pageIndex?: number
  slideIndex?: number
  sheetId?: string
}
```

四个内置 Surface 的可选支持：

| Surface | `getCaptureSurface()` | 说明 |
|---|---|---|
| `PdfSurface` | 页面 interaction 元素 + 当前 `<img>` 位图 | 像素级裁剪；页面尚未渲染时退回 DOM 栅格化 |
| `XlsxSheetSurface` | body canvas（canvas 模式）或 grid 元素（DOM 模式） | 提供 `ManualRegionLocator` 时会映射到当前可见视口；跨出视口的部分会被裁掉 |
| `DocxDocumentSurface` | 当前 `[data-docx-page-wrapper]` | 纯 DOM，通过 SVG `foreignObject` 栅格化 |
| `PptxStage` | 当前 `[data-slide-index]` 幻灯片 | 纯 DOM，通过 SVG `foreignObject` 栅格化 |

### 输出与限制

- `width` / `height`：请求的 `scale` 乘 CSS 区域尺寸，四舍五入到整数像素。
- `region`：实际被裁剪的归一化区域；`source.region` 与它一致。
- `source.elementWidth` / `elementHeight`：元素 CSS 像素尺寸；`source.sourceWidth` / `sourceHeight`：提供 raster 时的 intrinsic 像素。
- 默认 `scale = devicePixelRatio`，最大 `8`。
- `maxPixels` 默认 `24_000_000`，`maxDimension` 默认 `16384`，`maxBytes` 默认 `32 MiB`。
- `scale`、像素数、边长或编码后字节数超限时抛出 `OfficeRegionCaptureError`，`code` 为 `CAPTURE_LIMIT_EXCEEDED`。
- 无 DOM canvas 且未提供 `createCanvas` 时抛出 `CAPTURE_UNSUPPORTED`，不会访问 `document`。
- `signal` 取消返回 `code: "ABORTED"`。
- 默认返回 `objectUrl`；传 `objectUrl: false` 可只拿 `blob`。`release()` 幂等。

### DOM 栅格化限制

纯 DOM 路径会内联计算样式、把 canvas 子节点转成 data URL，并尽力内联同源图片。以下内容不会写入位图：

- 跨域图片；
- 外部 CSS `background-image`；
- 未安装的 web font（会使用可用字体回退）；
- `::before` / `::after` 伪元素；
- 元素内部滚动容器的不可见部分。

需要像素级结果时，优先让 Surface 提供 `raster`，或由宿主绘制自己的 canvas 后调用 `captureRegion`。

## OfficeRegionAnnotator

```vue
<script setup lang="ts">
import { ref } from "vue"
import {
  OfficeRegionAnnotator,
  createOfficeRegionAnnotationDocument,
  parseOfficeRegionAnnotationDocument,
  stringifyOfficeRegionAnnotationDocument,
  type OfficeRegionAnnotationDocument,
} from "@arcships/vue-ui"

const document = ref<OfficeRegionAnnotationDocument>(
  createOfficeRegionAnnotationDocument(),
)
const json = ref("")

function save() {
  json.value = stringifyOfficeRegionAnnotationDocument(document.value)
}

function load(value: string) {
  document.value = parseOfficeRegionAnnotationDocument(value)
}
</script>

<template>
  <OfficeRegionAnnotator
    v-model="document"
    :image="capture.url"
    :width="capture.width"
    :height="capture.height"
    @annotation-change="save"
  />
  <pre>{{ json }}</pre>
</template>
```

属性：

| 属性 | 默认值 | 说明 |
|---|---|---|
| `modelValue` | `null` | 受控标注文档；组件只发出新对象，不原地修改 |
| `image` | `null` | 截图 URL 或 `Blob`；`Blob` 的 object URL 由组件创建并在变化/卸载时撤销 |
| `width` / `height` | 解码尺寸或 `1000` | intrinsic 位图尺寸，用于 SVG viewBox |
| `mode` | `"rectangle"` | `arrow` / `rectangle` / `text` / `highlight` |
| `color` | `"#ef4444"` | 新标注颜色 |
| `strokeWidth` | `0.006` | 相对短边的归一化线宽 |
| `fontSize` | `0.05` | 相对短边的归一化文字大小 |
| `disabled` / `readOnly` | `false` | 禁用交互；`readOnly` 仍渲染已有标注 |
| `maxAnnotations` | `500` | 超过后不再添加，并发出 `annotation-error` |

组件实例 expose：`addAnnotation()`、`removeAnnotation()`、`removeSelected()`、`clearAnnotations()`、`getDocument()`、`toJSON()`、`loadJSON()`、`selectAnnotation()`、`getSelectedId()` 和 `getImageSize()`。

键盘：`Escape` 取消当前草稿/选择；选中标注后按 `Delete` 或 `Backspace` 删除。文字工具点击后在位图上显示内联输入框，回车或失焦提交。

## 标注文档

```ts
interface OfficeRegionAnnotationDocument {
  schemaVersion: 1
  annotations: OfficeRegionAnnotation[]
  source?: OfficeRegionCaptureMetadata
}

type OfficeRegionAnnotation =
  | { id: string; kind: "rectangle"; rect: NormalizedRect; style?: Style }
  | { id: string; kind: "highlight"; rect: NormalizedRect; style?: Style }
  | { id: string; kind: "arrow"; start: NormalizedPoint; end: NormalizedPoint; style?: Style }
  | { id: string; kind: "text"; point: NormalizedPoint; text: string; style?: Style }

interface Style {
  color?: string
  opacity?: number
  /** 相对短边的归一化值，(0, 0.25]。 */
  strokeWidth?: number
  /** 相对短边的归一化值，(0, 0.5]。 */
  fontSize?: number
}
```

- 所有几何坐标都相对截图自身左上角与宽高归一化为 `[0, 1]`，与 capture scale、DPR、页面滚动和缩放无关。
- 样式中的线宽和字号相对位图短边归一化，因此同一文档在 1x/2x 位图上视觉一致。
- `parseOfficeRegionAnnotationDocument()` 校验 schema、坐标、样式、文字长度、id 唯一性和数量上限；未知字段会被丢弃，返回值就是规范化结果。
- `stringifyOfficeRegionAnnotationDocument()` 输出 JSON；再次解析得到深度相等对象。
- `officeRegionAnnotationFromRegionAnnotation()` 可把 `@arcships/office-interaction` 的 `RegionAnnotation`（rectangle / arrow / text）转换为本文档类型；`ellipse`、`freehand`、`pin` 返回 `null`，因为标注编辑器不提供这些工具。

## Demo

`apps/demo` 的 DOCX Surface 页面演示完整链路：选择“区域”后拖出区域 → 选择 1x/2x → 截图 → 在 `OfficeRegionAnnotator` 中绘制 → 右侧 JSON 随编辑更新。demo 在替换截图、切换文件和卸载页面时都会 `release()` 上一张位图。

## 迁移说明

- 现有 `OfficeRegionSelector`、`regionDraftChange` 和 `regionConfirm` 不变。
- `captureReferencePreview()` 仍是宿主渲染管线责任；没有新增伪造截图行为。需要内建截图时使用 `captureRegion()`。
- 新组件和函数都是增量导出，`0.x` 内不删除既有入口。
