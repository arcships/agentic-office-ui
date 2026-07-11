<template>
  <button
    ref="rootRef"
    class="pptx-thumbnail"
    :class="{ 'pptx-thumbnail--active': active }"
    type="button"
    :aria-label="`第 ${slide.number} 页${slide.hidden ? '，隐藏页' : ''}`"
    :aria-current="active ? 'page' : undefined"
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
import { onBeforeUnmount, onMounted, ref, watch } from "vue"
import type {
  PptxDisposableHandle,
  PptxPreviewSession,
} from "@arcships/pptx-core/browser"
import type { PptxSlideInfo } from "@arcships/pptx-core"

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
let handle: PptxDisposableHandle | null = null
let observer: IntersectionObserver | null = null
let renderId = 0

function clear(): void {
  renderId += 1
  handle?.dispose()
  handle = null
  const canvas = canvasRef.value
  canvas?.replaceChildren()
  if (canvas) delete canvas.dataset.error
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

function observe(): void {
  observer?.disconnect()
  observer = null
  const root = rootRef.value
  if (!root || typeof IntersectionObserver === "undefined") {
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
  observer?.disconnect()
  clear()
})
</script>
