# @arcships/vue-pptx

## 当前状态

本包支持 PPTX、PPTM、PPSX、PPSM、POTX、POTM 的静态浏览和 `mode="present"` 演示模式。默认静态浏览会像 PDF/DOCX 一样纵向连续展示所有幻灯片；演示模式使用单页舞台，左右键执行上/下一播放步骤，跨页由播放控制器完成。宏内容不会执行。

```bash
pnpm add @arcships/pptx-core@0.5.4 @arcships/vue-pptx@0.5.4
```

```ts
import { PptxViewer } from "@arcships/vue-pptx"
import "@arcships/vue-pptx/style.css"
```

## 组件边界

Vue 包负责：

- 浏览和演示界面；
- 控制条、键盘、鼠标、触摸和全屏；
- 加载、媒体恢复、错误和能力提示；
- 将公开控制器暴露给组件调用方。

Vue 包不负责：

- 解包 PPTX；
- 解析动画 XML；
- 计算时间节点；
- 合并动画属性；
- 保存独立于控制器的第二份播放状态。

`PptxViewer` 默认以纵向连续页面浏览。传入 `mode="present"` 后切换为单页播放，并可使用下一步、上一步、暂停、继续、重播、跳页、媒体恢复和全屏；模板引用也暴露同一组公开方法。

## 最小组合方式

不需要完整工具栏时，可以只使用舞台和两个组合能力：

```vue
<script setup lang="ts">
import { computed, ref } from "vue"
import {
  PptxStage,
  usePptxDocument,
  usePptxPlayback,
  type PptxStageExpose,
  type PptxPreviewSource,
} from "@arcships/vue-pptx"
import "@arcships/vue-pptx/style.css"

const props = defineProps<{ source: PptxPreviewSource | null }>()
const stage = ref<PptxStageExpose | null>(null)
const element = computed(() => stage.value?.element ?? null)
const document = usePptxDocument(element, {
  source: () => props.source,
  session: { renderMode: "slide" },
})
const playback = usePptxPlayback(document)

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
  <PptxStage ref="stage" class="my-pptx-stage" @click="onStageClick" />
</template>
```

- `usePptxDocument` 负责打开文件、滚动跳页、缩放和销毁；`renderMode: "list"` 用于纵向 surface，`renderMode: "slide"` 用于播放舞台；
- `usePptxPlayback` 负责动画步骤、播放状态和播放事件；
- `PptxStage` 不持有文档或播放状态，但会从渲染标记派生统一 Surface 事件；也可以换成普通 `<div>` 并自行处理这些事件。

`PptxStage` 还提供完整的 surface 交互事件：

- `selection-change`：点击幻灯片时返回 `{ kind: "slide", slideIndex }`；
- `object-click`：点击形状、图片或图表时返回 `{ kind: "object", slideIndex, objectKey }`；
- `context-menu`：右键幻灯片或对象时返回判别联合；`kind: "slide"` 时没有 `objectKey`，`kind: "object"` 时必须包含 `objectKey`，两者都有页码以及视口/`PptxStage` 根元素坐标。

传入 `v-model:zoom` 和显式 `scroll-container` 后，`PptxStage` 支持 50%–200% 的 `Ctrl+wheel` / pinch 缩放，并保持指针下的幻灯片位置。播放模式建议关闭 `enableGestureZoom`。

## 文档

- [PPTX 使用指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/guide/pptx.md)
- [组件手册](https://github.com/arcships/agentic-office-ui/blob/master/docs/components/README.md)
- [PPTX 组合函数](https://github.com/arcships/agentic-office-ui/blob/master/docs/composables/pptx.md)
- [自定义 PPTX 播放器](https://github.com/arcships/agentic-office-ui/blob/master/docs/custom-components/pptx.md)
- [API 导航](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/README.md)
