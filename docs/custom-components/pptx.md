# 自定义 PPTX Surface 与播放器

## 纵向浏览 Surface

普通浏览只需要 `PptxStage` 和 `usePptxDocument`：

```vue
<script setup lang="ts">
import { computed, ref } from "vue"
import {
  PptxStage,
  usePptxDocument,
  type PptxPreviewSource,
  type PptxStageContextMenu,
  type PptxStageExpose,
  type PptxStageObjectClick,
  type PptxStageSelection,
} from "@arcships/vue-pptx"
import "@arcships/vue-pptx/style.css"

const props = defineProps<{ source: PptxPreviewSource | null }>()
const stage = ref<PptxStageExpose | null>(null)
const element = computed(() => stage.value?.element ?? null)
const document = usePptxDocument(element, {
  source: () => props.source,
  session: {
    renderMode: "list",
    listOptions: { windowed: true, initialSlides: 4 },
  },
})

function onSelection(selection: PptxStageSelection) {}
function onObjectClick(object: PptxStageObjectClick) {}
function onContextMenu(context: PptxStageContextMenu) {}
</script>

<template>
  <PptxStage
    ref="stage"
    class="my-surface"
    @selection-change="onSelection"
    @object-click="onObjectClick"
    @context-menu="onContextMenu"
  />
</template>
```

列表模式纵向连续展示全部幻灯片，并按可视区域挂载页面。`document.goTo()` 和上一页/下一页方法执行滚动定位；滚动也会同步 `activeIndex`。

## 单页播放器

PPTX 提供完整的无界面组合方式：

```vue
<script setup lang="ts">
import { computed, ref } from "vue"
import {
  PptxStage,
  usePptxDocument,
  usePptxPlayback,
  type PptxPreviewSource,
  type PptxStageExpose,
} from "@arcships/vue-pptx"
import "@arcships/vue-pptx/style.css"

const props = defineProps<{ source: PptxPreviewSource | null }>()
const stage = ref<PptxStageExpose | null>(null)
const element = computed(() => stage.value?.element ?? null)
const document = usePptxDocument(element, {
  source: () => props.source,
  session: { renderMode: "slide" },
})
const playback = usePptxPlayback(document, { autoplay: false })
const { activeIndex } = document

async function onStageClick(event: MouseEvent) {
  const object = (event.target as Element | null)
    ?.closest<HTMLElement>("[data-pptx-object-key]")
  const handled = object?.dataset.pptxObjectKey
    ? await playback.activateObject(object.dataset.pptxObjectKey)
    : false
  if (!handled) await playback.next()
}
</script>

<template>
  <nav>
    <button @click="playback.previous()">上一步</button>
    <button @click="playback.next()">下一步</button>
    <button @click="playback.pause()">暂停</button>
    <button @click="playback.resume()">继续</button>
    <span>第 {{ activeIndex + 1 }} 页</span>
  </nav>
  <PptxStage ref="stage" class="my-stage" @click="onStageClick" />
</template>
```

每个播放器各自调用两个组合函数。不要让多个舞台共享同一个文档组合函数，也不要在播放期间用 `document.nextSlide()` 代替 `playback.next()`。

左右方向键和页面切换属于播放模式。普通 Surface 使用纵向滚动，以及上/下方向键或 PageUp/PageDown。

全屏属于外层业务界面，可以对包含舞台和工具栏的元素调用浏览器 Fullscreen API。
