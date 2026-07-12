# PPTX 最小组合能力与舞台组件设计

## 1. 范围

- 状态：已实现并通过组件、外部消费和浏览器验证
- 日期：2026-07-12
- 适用包：`@arcships/vue-pptx`
- 底层接口：`@arcships/pptx-core/browser`
- 关联文档：[播放接口](pptx-playback-api-design.md)、[正式实现设计](pptx-playback-implementation-design.md)

第一版只增加：

- `usePptxDocument`：打开文档和浏览页面；
- `usePptxPlayback`：控制动画与播放；
- `PptxStage`：提供渲染所需的 DOM 元素。

现有 `PptxThumbnail` 和 `PptxViewer` 保留。第一版不增加搜索、全屏、Provider、控制按钮组件、业务封装，也不增加新的解析和动画能力。

## 2. 基本结构

```text
usePptxDocument
  └── 唯一 PptxDocumentSession
        └── usePptxPlayback
              └── 唯一 PptxPlaybackController

PptxStage
  └── 只提供 DOM，不保存文档和播放状态
```

`PptxViewer` 改用这三个公开能力，不再维护另一套会话、控制器和播放状态。

## 3. 公开导出

```ts
export { usePptxDocument } from "./composables/usePptxDocument"
export { usePptxPlayback } from "./composables/usePptxPlayback"
export { default as PptxStage } from "./PptxStage.vue"

export type {
  PptxDocumentState,
  UsePptxDocumentOptions,
  UsePptxDocumentReturn,
  UsePptxPlaybackOptions,
  UsePptxPlaybackReturn,
  PptxStageExpose,
} from "./headless-types"
```

统一从 `@arcships/vue-pptx` 导入，样式仍使用 `@arcships/vue-pptx/style.css`。

## 4. `usePptxDocument`

### 4.1 接口

```ts
type PptxStageTarget =
  | Readonly<Ref<HTMLElement | null>>
  | (() => HTMLElement | null)

interface UsePptxDocumentOptions {
  source?: MaybeRefOrGetter<PptxPreviewSource | null | undefined>
  initialSlide?: MaybeRefOrGetter<number | undefined>
  session?: MaybeRefOrGetter<PptxDocumentSessionOptions | undefined>
}

type PptxDocumentState =
  | "idle"
  | "waiting-for-stage"
  | "loading"
  | "ready"
  | "error"
  | "disposed"

interface UsePptxDocumentReturn {
  readonly state: Readonly<Ref<PptxDocumentState>>
  readonly error: Readonly<ShallowRef<PptxPreviewError | null>>
  readonly document: Readonly<ShallowRef<PptxPreviewDocument | null>>
  readonly capability: Readonly<ShallowRef<PptxCapabilityReport | null>>
  readonly activeIndex: Readonly<Ref<number>>
  readonly zoomPercent: Readonly<Ref<number>>

  open(source: PptxPreviewSource): Promise<PptxPreviewDocument>
  close(): void
  goTo(index: number): Promise<void>
  nextSlide(): Promise<void>
  previousSlide(): Promise<void>
  setZoom(percent: number): Promise<void>
  getSession(): PptxDocumentSession | null
  dispose(): void
}

function usePptxDocument(
  target: PptxStageTarget,
  options?: UsePptxDocumentOptions,
): UsePptxDocumentReturn
```

### 4.2 规则

- 舞台和来源都可用时自动打开文档；舞台尚未挂载时进入 `waiting-for-stage`；
- 新来源或新舞台出现时，旧打开任务失效，旧会话销毁，只保留最新结果；
- `goTo()`、`nextSlide()` 和 `previousSlide()` 只做普通翻页，不执行动画；
- 页面编号从零开始，页面渲染成功后才更新 `activeIndex`；
- `close()` 关闭当前文档，之后仍可再次打开；
- Vue 作用域销毁时自动调用 `dispose()`；
- `getSession()` 只给高级功能使用，普通调用方使用上面的公开方法。

## 5. `usePptxPlayback`

### 5.1 接口

```ts
interface UsePptxPlaybackOptions {
  enabled?: MaybeRefOrGetter<boolean>
  autoplay?: boolean
  skipHiddenSlides?: boolean
  approximation?: PptxApproximationPolicy
  onEvent?: (event: PptxPlaybackEvent) => void
}

interface UsePptxPlaybackReturn {
  readonly controller: Readonly<ShallowRef<PptxPlaybackController | null>>
  readonly snapshot: Readonly<ShallowRef<PptxPlaybackSnapshot | null>>
  readonly status: ComputedRef<PptxPlaybackStatus | "unavailable">
  readonly capability: Readonly<ShallowRef<PptxCapabilityReport | null>>
  readonly lastWarning: Readonly<ShallowRef<PptxPlaybackWarning | null>>
  readonly lastError: Readonly<ShallowRef<PptxPlaybackError | null>>

  next(): Promise<void>
  activateObject(objectKey: string): Promise<boolean>
  previous(): Promise<void>
  play(): Promise<void>
  pause(): void
  resume(): Promise<void>
  reset(): Promise<void>
  goToSlide(index: number, options?: { includeHidden?: boolean }): Promise<void>
  resumeBlockedMedia(mediaId?: string): Promise<void>
  dispose(): void
}

function usePptxPlayback(
  document: UsePptxDocumentReturn,
  options?: UsePptxPlaybackOptions,
): UsePptxPlaybackReturn
```

### 5.2 规则

- 文档就绪后创建控制器；文档关闭、更换或禁用播放时销毁控制器；
- `next()` 先执行下一动画步骤，本页动画完成后才换页；
- `activateObject()` 返回 `false` 时，由调用方决定是否继续调用 `next()`；
- 播放控制器换页后直接同步 `document.activeIndex`，不能再次渲染同一页；
- 演示模式使用 `goToSlide()`，不混用 `document.goTo()`；
- `dispose()` 只销毁播放控制器，不销毁文档会话。

## 6. `PptxStage`

`PptxStage` 只提供渲染容器：

```ts
interface PptxStageExpose {
  readonly element: HTMLElement | null
}
```

根元素包含基础类名和测试标记：

```html
<div class="pptx-stage" data-testid="pptx-stage" />
```

组件把 `class`、`style`、无障碍属性和 DOM 事件透传到根元素，不自行定义加载状态、尺寸、点击行为或焦点行为。使用者也可以不用该组件，直接把普通 `<div>` 传给 `usePptxDocument`。

## 7. 最小用法

```vue
<script setup lang="ts">
import { computed, useTemplateRef } from "vue"
import {
  PptxStage,
  usePptxDocument,
  usePptxPlayback,
  type PptxStageExpose,
} from "@arcships/vue-pptx"

const props = defineProps<{ file: File | ArrayBuffer | null }>()
const stage = useTemplateRef<PptxStageExpose>("stage")
const element = computed(() => stage.value?.element ?? null)
const document = usePptxDocument(element, { source: () => props.file })
const playback = usePptxPlayback(document)

async function onClick(event: MouseEvent) {
  const target = (event.target as HTMLElement)
    .closest<HTMLElement>("[data-pptx-object-key]")
  const handled = target?.dataset.pptxObjectKey
    ? await playback.activateObject(target.dataset.pptxObjectKey)
    : false

  if (!handled) await playback.next()
}
</script>

<template>
  <PptxStage ref="stage" @click="onClick" />
</template>
```

## 8. `PptxViewer` 兼容要求

`PptxViewer` 内部改为使用 `PptxStage`、`usePptxDocument` 和 `usePptxPlayback`，但以下内容保持不变：

- 现有属性、事件和公开方法；
- 默认浏览模式；
- 工具栏、缩略图、键盘、全屏和状态界面；
- `pptx-viewer__stage` 类名；
- `PptxThumbnail` 的行为。

工具栏和界面状态仍由 `PptxViewer` 负责。文档会话和播放控制器只能由两个组合能力持有。

## 9. 文件结构与实现顺序

```text
packages/vue-pptx/src/
├── PptxStage.vue
├── PptxViewer.vue
├── composables/
│   ├── usePptxDocument.ts
│   └── usePptxPlayback.ts
├── headless-types.ts
└── index.ts
```

实现顺序：

1. 增加公开类型和 `PptxStage`；
2. 从 `PptxViewer` 提取 `usePptxDocument`；
3. 提取 `usePptxPlayback`；
4. 让 `PptxViewer` 改用新能力；
5. 增加 README 示例和 npm 压缩包消费验证。

每一步完成后，现有 `PptxViewer` 都必须可以运行。

## 10. 验收条件

- 普通 `<div>` 和 `PptxStage` 都能打开并显示 PPTX；
- 浏览翻页、缩放、动画播放和对象点击可用；
- 快速换文件或换舞台只保留最新结果；
- 卸载后没有旧控制器、媒体和对象地址残留；
- `PptxViewer` 的公开接口和现有行为不变；
- npm 压缩包安装后，完整组件和最小组合方式都能运行；
- 使用“装修付款节奏与现金流规划.pptx”确认第一次点击执行动画，不直接换页。

满足以上条件后停止第一版开发，不继续拆分更多组合能力和组件。
