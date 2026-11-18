# @arcships/vue-pptx

## 当前状态

本包支持 PPTX、PPTM、PPSX、PPSM、POTX、POTM 的静态浏览和 `mode="present"` 演示模式。默认静态浏览会像 PDF/DOCX 一样纵向连续展示所有幻灯片；演示模式使用单页舞台，左右键执行上/下一播放步骤，跨页由播放控制器完成。宏内容不会执行。

```bash
pnpm add @arcships/pptx-core @arcships/vue-pptx
```

```ts
import { PptxViewer } from "@arcships/vue-pptx"
import "@arcships/vue-pptx/style.css"
```

## 本地化

`PptxViewer` 接受 `locale`（默认 `"en"`，内置 `"zh-CN"`）与 `messages` 覆盖。用户可见文案和 aria-label 来自 `@arcships/office-interaction` 的共享消息目录：

```vue
<PptxViewer
  :source="source"
  locale="zh-CN"
  :messages="{ pptx: { viewer: { search: '查找' } } }"
/>
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

`PptxViewer` 默认以纵向连续页面浏览。传入 `mode="present"` 后切换为单页播放，并可使用下一步、上一步、暂停、继续、重播、跳页、媒体恢复和全屏；模板引用也暴露同一组公开方法。切换浏览/演示模式会保留当前页，全屏时自动占满屏幕，演示模式按幻灯片比例等比缩放并留黑边。

界面状态通过 `data-state` 暴露：`empty`、`waiting`、`loading`、`ready`、`cancelled` 和 `error`。单页渲染失败不再锁死查看器：上一页保持可见，顶部提示失败页并提供“重试/跳过”；加载超时或取消会进入可重试状态。文档 `warnings`（如不支持的演讲者备注）显示在提示条中，可用 `warnings` 插槽替换，或用 `showWarnings=false` 隐藏。

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
- `context-menu`：右键幻灯片或对象时返回判别联合；`kind: "slide"` 时没有 `objectKey`，`kind: "object"` 时必须包含 `objectKey`，两者都有页码以及视口/`PptxStage` 根元素坐标；
- `link-activate`：链接形状获得焦点后按 `Enter`/`Space`，返回 `{ slideIndex, objectKey }`。

`PptxViewer` 还会先抛出同一个 `context-menu` 事件，并（默认）为带链接的形状显示内置菜单；`PptxViewer` 在浏览和演示模式下都会把 `a:hlinkClick` 交给共享 URL 策略：`http:`/`https:` 和 `mailto:` 在新标签页打开，文档内跳页走 `goTo()`，`javascript:` 等目标被拒绝，`action` 事件照旧抛出。`showContextMenu="false"` 关闭内置菜单，`linkHandling="host"` 关闭内置打开/复制行为。悬停链接（`a:hlinkHover`）只提供 tooltip 并记为 `static` 能力。

传入受控的 `selection-mode="content | object | region"` 后，`PptxStage` 会把对象内文字、幻灯片/可见对象点击和区域拖选转换为统一 `reference-confirm`，并提供候选、取消、revision、解析和错误事件。Stage 只读取渲染标记并发送事件，不持有引用集合、工具栏或 Agent 状态。

传入 `v-model:zoom` 和显式 `scroll-container` 后，`PptxStage` 支持 50%–200% 的 `Ctrl+wheel` / pinch 缩放，并保持指针下的幻灯片位置。播放模式建议关闭 `enableGestureZoom`。

## 文档

- [PPTX 使用指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/guide/pptx.md)
- [组件手册](https://github.com/arcships/agentic-office-ui/blob/master/docs/components/README.md)
- [PPTX 组合函数](https://github.com/arcships/agentic-office-ui/blob/master/docs/composables/pptx.md)
- [自定义 PPTX 播放器](https://github.com/arcships/agentic-office-ui/blob/master/docs/custom-components/pptx.md)
- [API 导航](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/README.md)
