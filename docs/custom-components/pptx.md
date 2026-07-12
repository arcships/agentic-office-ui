# 自定义 PPTX 播放器

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
const document = usePptxDocument(element, { source: () => props.source })
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

全屏属于外层业务界面，可以对包含舞台和工具栏的元素调用浏览器 Fullscreen API。
