# 接入 0.5.0 最小 Surface 组件指南

`@arcships/vue-docx@0.5.0` / `vue-xlsx@0.5.0` / `vue-pptx@0.5.0` / `vue-pdf@0.5.0`

## 升级

```bash
pnpm add @arcships/vue-docx@0.5.0 @arcships/vue-xlsx@0.5.0 @arcships/vue-pptx@0.5.0 @arcships/vue-pdf@0.5.0
```

版本号推到 `^0.5.0`，API 与 `0.4.0` 兼容。

## 四种 Surface 组件

原来的 Viewer 组件自带 toolbar / 缩略图 / 搜索栏 / 播放控制等全套 UI。新版本提供**只渲染内容画布**的 Surface 组件，宿主自己搭 toolbar 和控制栏。

| 场景 | 之前（0.4.x） | 现在（0.5.x） |
|---|---|---|
| DOCX 查看 | `<DocxViewer :file="buf" />` | `<DocxDocumentSurface :model="model" />` |
| XLSX 查看 | `<XlsxViewer :controller="ctrl" />` | `<XlsxSheetSurface :controller="ctrl" />` |
| PPTX 查看 | `<PptxViewer :source="file" />` | `<PptxStage />` + `usePptxDocument` |
| PDF 查看 | `<PdfViewer :src="url" />` | `<PdfSurface :source="source" />` |

旧的 Viewer 组件全部保留，行为不变。

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

**Props**：`model`（必填）、`zoomScale`、`fitWidth`、`showComments`、`showTrackedChanges`、`searchQuery`、`theme`

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

## PPTX

```vue
<script setup>
import { PptxStage, usePptxDocument, usePptxPlayback } from '@arcships/vue-pptx'

const stageRef = useTemplateRef('stage')
const doc = usePptxDocument(() => stageRef.value?.element, { source })
const playback = usePptxPlayback(doc, { autoplay: false })
</script>

<template>
  <PptxStage ref="stage" />
  <button @click="doc.previousSlide()">上一页</button>
  <span>{{ doc.activeIndex + 1 }} / {{ doc.document?.slides.length }}</span>
  <button @click="doc.nextSlide()">下一页</button>
</template>
```

**事件（Stage）**：`contextMenu`（带 `containerX/Y`）、`objectClick`

**Composable**：`usePptxDocument` 返回 `state`、`document`、`activeIndex`、`goTo(n)`、`nextSlide()`、`previousSlide()`、`setZoom(n)`、`getSession()`

**CSS 变量**：`--pptx-surface-bg`

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

**Props**：`source`、`src`、`fitWidth`、`defaultZoom`

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
}
```

宿主在这个坐标弹出自己的菜单即可。

## 搜索

Surface 组件不内置搜索 UI。搜索能力分布：

| 格式 | 搜索入口 | 说明 |
|---|---|---|
| DOCX | Surface prop `searchQuery` + `activeSearchNodeIndex` | 传 prop 即可，surface 内部高亮命中节点 |
| XLSX | 无公开搜索 API | 当前版本不提供搜索，宿主需自行实现（如遍历 cell 值匹配） |
| PPTX | `session.searchText(query)` → `PptxSearchResult[]` | 宿主调 API，根据 `slideIndex` 自行 goTo 翻页 |
| PDF | `runtime.search(doc, query)` → `PdfSearchHit[]`（API 存在，Surface 未接入） | 宿主需自行调用 `createPdfRenderRuntime().search()`，根据 `pageIndex` 自行 scrollToPage |

XLSX 搜索暂未提供公开 API；PPTX 和 PDF 的搜索为 composable/runtime 级 API，不通过 surface prop 透传。

DOCX 之外的格式，宿主需要自己实现搜索输入框 + 结果列表 + 导航逻辑。
