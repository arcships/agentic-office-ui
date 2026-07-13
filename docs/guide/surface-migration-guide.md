# 接入 0.5.3 最小 Surface 组件指南

`@arcships/vue-docx@0.5.3` / `vue-xlsx@0.5.3` / `vue-pptx@0.5.3` / `vue-pdf@0.5.3`

## 升级

```bash
pnpm add @arcships/vue-docx@0.5.3 @arcships/vue-xlsx@0.5.3 @arcships/vue-pptx@0.5.3 @arcships/vue-pdf@0.5.3
```

版本号推到 `^0.5.3`。公开入口与 `0.5.2` 兼容；新增格式与 Surface 手势缩放不要求迁移旧调用。

## 四种 Surface 组件

原来的 Viewer 组件自带 toolbar / 缩略图 / 搜索栏 / 播放控制等全套 UI。新版本提供**只渲染内容画布**的 Surface 组件，宿主自己搭 toolbar 和控制栏。

| 场景 | 之前（0.4.x） | 现在（0.5.x） |
|---|---|---|
| DOCX 查看 | `<DocxViewer :file="buf" />` | `<DocxDocumentSurface :model="model" />` |
| XLSX 查看 | `<XlsxViewer :controller="ctrl" />` | `<XlsxSheetSurface :controller="ctrl" />` |
| PPTX 查看 | `<PptxViewer :source="file" />` | `<PptxStage />` + `usePptxDocument`，纵向连续页面 |
| PDF 查看 | `<PdfViewer :src="url" />` | `<PdfSurface :source="source" />` |

旧的 Viewer 组件全部保留。PPTX 的 `mode="browse"` 从 `0.5.2` 起改为纵向连续页面；`mode="present"` 仍保持单页播放。

## 统一缩放

四个 Surface 使用相同的受控倍率，宿主保留 toolbar 和产品状态：

```vue
<FormatSurface v-model:zoom="zoom" />
```

```ts
const zoom = ref(1) // 100%，范围 0.5–2
```

传入 `zoom` 后，Surface 处理 `Ctrl+wheel`、trackpad pinch 和 WebKit gesture，并保持指针下的页面、幻灯片或单元格。普通 wheel 继续滚动。`enableGestureZoom` 默认为 `true`；未传 `zoom` 时不接管手势。受控 `zoom` 与 `fitWidth` 同时存在时以 `zoom` 为准。

## DOCX

```vue
<script setup>
import { DocxDocumentSurface } from '@arcships/vue-docx'
import { createDocxRuntime } from '@arcships/docx-core'

const runtime = createDocxRuntime()
const model = ref(null)
const surfaceRef = ref()

async function open(file) {
  const result = await runtime.createLoader().load(file)
  model.value = result.model
}
</script>

<template>
  <div style="height: 80vh">
    <DocxDocumentSurface
      ref="surfaceRef"
      :model="model"
      :show-comments="true"
      :show-tracked-changes="false"
      :fit-width="true"
      @page-count-change="total = $event"
      @selection-change="onSelection"
      @context-menu="onContextMenu"
    />
  </div>

  <!-- 宿主自己的工具栏 -->
  <button @click="surfaceRef.scrollToPage(0)">首页</button>
</template>
```

**Props**：`model`（必填）、`zoom`、`enableGestureZoom`、兼容属性 `zoomScale`、`fitWidth`、`showComments`、`showTrackedChanges`、`searchQuery`、`theme`

**事件**：`pageCountChange`、`visiblePageRange`、`contextMenu`（带 `pageIndex` + `containerX/Y`）、`selectionChange`（带选中文字 + 页码）

**Expose**：`scrollToPage(n)`、`scrollToNode(n)`、`scrollContainer`

**CSS 变量**：`--docx-surface-bg`（滚动区背景色）

## XLSX

```vue
<script setup>
import { XlsxSheetSurface, useXlsxViewerController } from '@arcships/vue-xlsx'

const controller = useXlsxViewerController({ file: buffer, fileName: 'a.xlsx' })
</script>

<template>
  <div style="height: 80vh">
    <XlsxSheetSurface
      :controller="controller"
      :read-only="false"
      @context-menu="onContextMenu"
      @selection-change="onSelection"
    />
  </div>

  <!-- 宿主自己的 sheet tabs -->
  <div v-for="(t, i) in controller.tabs" @click="controller.setActiveTabIndex(i)">
    {{ t.label }}
  </div>
</template>
```

**事件**：`contextMenu`（带 `selection` range + `activeCell` + `sheetName` + `containerX/Y`）、`selectionChange`（带 cell range）、`cellDoubleClick`

**CSS 变量**：`--xlsx-surface-bg`

**缩放**：`zoom`、`enableGestureZoom`、`update:zoom`；冻结行列和行列标题按轴保持单元格锚点。

## PPTX

```vue
<script setup lang="ts">
import { ref } from 'vue'
import {
  PptxStage,
  usePptxDocument,
  type PptxPreviewSource,
  type PptxStageContextMenu,
  type PptxStageObjectClick,
  type PptxStageExpose,
  type PptxStageSelection,
} from '@arcships/vue-pptx'
import '@arcships/vue-pptx/style.css'

const props = defineProps<{ source: PptxPreviewSource | null }>()
const stage = ref<PptxStageExpose | null>(null)
const doc = usePptxDocument(() => stage.value?.element, {
  source: () => props.source,
  session: {
    renderMode: 'list',
    listOptions: { windowed: true, initialSlides: 4 },
  },
})
const { activeIndex, document } = doc

function onSelection(selection: PptxStageSelection) {}
function onObjectClick(object: PptxStageObjectClick) {}
function onContextMenu(context: PptxStageContextMenu) {}
</script>

<template>
  <PptxStage
    ref="stage"
    @selection-change="onSelection"
    @object-click="onObjectClick"
    @context-menu="onContextMenu"
  />
  <button @click="doc.previousSlide()">上一页</button>
  <span>{{ activeIndex + 1 }} / {{ document?.slides.length ?? 0 }}</span>
  <button @click="doc.nextSlide()">下一页</button>
</template>
```

**事件（Stage）**：`selectionChange`（带 `slideIndex`）、`objectClick`（带 `slideIndex` + `objectKey`）、`contextMenu`（`kind: 'slide'` 时无 `objectKey`，`kind: 'object'` 时 `objectKey` 必填；两者都带 `slideIndex` 和 `clientX/Y` + `containerX/Y`）

**Composable**：`usePptxDocument` 返回 `state`、`document`、`activeIndex`、`goTo(n)`、`nextSlide()`、`previousSlide()`、`setZoom(n)`、`getSession()`。列表模式下跳页方法滚动到目标页，滚动也同步 `activeIndex`。

**播放模式**：自定义播放器改用 `session: { renderMode: 'slide' }` 并增加 `usePptxPlayback`。不要在列表 Surface 上执行播放时间轴。

**CSS 变量**：`--pptx-surface-bg`

**缩放**：`zoom`、`enableGestureZoom`、`update:zoom`；列表模式需要把同一个显式滚动容器传给 `PptxStage` 和 session。

## PDF

```vue
<script setup>
import { PdfSurface } from '@arcships/vue-pdf'

const surfaceRef = ref()
const source = { kind: 'url', url: '/doc.pdf' }
</script>

<template>
  <div style="height: 80vh">
    <PdfSurface
      ref="surfaceRef"
      :source="source"
      :fit-width="true"
      @document-load-success="pages = $event"
      @visible-page-change="page = $event"
      @context-menu="onContextMenu"
    />
  </div>

  <button @click="surfaceRef.scrollToPage(0)">首页</button>
</template>
```

**Props**：`source`、`src`、`zoom`、`enableGestureZoom`、`fitWidth`、`defaultZoom`

**事件**：`documentLoadSuccess`、`documentLoadError`、`visiblePageChange`、`contextMenu`（带 `pageIndex` + `containerX/Y`）

**Expose**：`scrollToPage(n)`、`zoom`、`rotation`

**CSS 变量**：`--pdf-surface-bg`

## contextMenu 事件通用结构

所有 surface 的右键事件统一携带：

```ts
{
  clientX: number      // 视口坐标
  clientY: number
  containerX: number   // 相对 surface 容器左上角
  containerY: number
  // 格式特有：
  pageIndex?: number           // DOCX / PDF
  selection?: { start, end }   // XLSX cell range
  sheetName?: string           // XLSX
  activeCell?: { row, col }    // XLSX
  // PPTX 另使用判别联合：
  // { kind: 'slide', slideIndex, ...coords }
  // { kind: 'object', slideIndex, objectKey, ...coords }
}
```

宿主在这个坐标弹出自己的菜单即可。

## 搜索

Surface 组件不内置搜索 UI。搜索能力分布：

| 格式 | 搜索入口 | 说明 |
|---|---|---|
| DOCX | Surface prop `searchQuery` + `activeSearchNodeIndex` | 传 prop 即可，surface 内部高亮命中节点 |
| XLSX | `controller.findCells(query)` → `XlsxCellAddress[]` | 遍历 used range，返回匹配 cell 地址列表 |
| PPTX | `session.searchText(query)` → `PptxSearchResult[]` | 宿主调 API，根据 `slideIndex` 调用 `goTo()` 滚动定位 |
| PDF | `surfaceRef.search(query)` → `PdfSearchHit[]` | expose 方法，返回 pageIndex + rects |

XLSX、PPTX 和 PDF 的搜索为 controller/composable/runtime 级 API，不通过 surface prop 透传。

DOCX 之外的格式，宿主需要自己实现搜索输入框 + 结果列表 + 导航逻辑。
