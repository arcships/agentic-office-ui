<template>
  <button
    ref="rootRef"
    class="pptx-thumbnail"
    :class="{ 'pptx-thumbnail--active': active }"
    type="button"
    :aria-label="`第 ${slide.number} 页${slide.hidden ? '，隐藏页' : ''}`"
    :aria-current="active ? 'page' : undefined"
    :data-pptx-thumbnail-index="slide.index"
    @click="emit('select', slide.index)"
  >
    <span ref="canvasRef" class="pptx-thumbnail__canvas" />
    <span class="pptx-thumbnail__label">
      <span>{{ slide.number }}</span>
      <span v-if="slide.hidden" class="pptx-thumbnail__hidden">隐藏</span>
    </span>
  </button>
</template>

<script setup lang="ts">
import { inject, onBeforeUnmount, onMounted, ref, watch } from "vue"
import type {
  PptxDisposableHandle,
  PptxPreviewSession,
} from "@arcships/pptx-core/browser"
import type { PptxSlideInfo } from "@arcships/pptx-core"
import {
  PPTX_THUMBNAIL_OBSERVER,
  type PptxThumbnailObserver,
} from "./composables/thumbnail-observer"

const props = withDefaults(defineProps<{
  session: PptxPreviewSession
  slide: PptxSlideInfo
  width?: number
  active?: boolean
}>(), {
  width: 148,
  active: false,
})

const emit = defineEmits<{ select: [index: number] }>()
const rootRef = ref<HTMLElement | null>(null)
const canvasRef = ref<HTMLElement | null>(null)
/** Shared gate provided by `PptxViewer`; standalone thumbnails observe themselves. */
const sharedObserver = inject(PPTX_THUMBNAIL_OBSERVER, null) as PptxThumbnailObserver | null
let observer: IntersectionObserver | null = null
let observedRoot: HTMLElement | null = null
let handle: PptxDisposableHandle | null = null
let renderId = 0

function clear(): void {
  renderId += 1
  handle?.dispose()
  handle = null
  const canvas = canvasRef.value
  canvas?.replaceChildren?.()
  if (canvas?.dataset) delete canvas.dataset.error
}

async function render(): Promise<void> {
  const target = canvasRef.value
  if (!target || handle) return
  const id = ++renderId
  const nextHandle = props.session.renderThumbnail(props.slide.index, target, props.width)
  if (!nextHandle) return
  handle = nextHandle
  try {
    await nextHandle.ready
  } catch {
    if (id === renderId) target.dataset.error = "true"
  }
}

function stopObserving(): void {
  if (observedRoot) {
    sharedObserver?.unobserve(observedRoot)
    observedRoot = null
  }
  observer?.disconnect()
  observer = null
}

function observe(): void {
  stopObserving()
  const root = rootRef.value
  if (!root) {
    void render()
    return
  }
  if (sharedObserver) {
    observedRoot = root
    sharedObserver.observe(root, () => void render())
    return
  }
  if (typeof IntersectionObserver === "undefined") {
    void render()
    return
  }
  observer = new IntersectionObserver((entries) => {
    if (!entries.some((entry) => entry.isIntersecting)) return
    observer?.disconnect()
    observer = null
    void render()
  }, { rootMargin: "240px" })
  observer.observe(root)
}

watch(() => [props.session, props.slide.index, props.width] as const, () => {
  clear()
  observe()
})

onMounted(observe)
onBeforeUnmount(() => {
  stopObserving()
  clear()
})
</script>
