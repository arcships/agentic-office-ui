<template>
  <button
    ref="rootRef"
    class="pptx-thumbnail"
    :class="{
      'pptx-thumbnail--active': active,
      'pptx-thumbnail--loading': status === 'loading',
      'pptx-thumbnail--error': status === 'error',
    }"
    type="button"
    :aria-label="thumbnailLabel"
    :aria-current="active ? 'page' : undefined"
    :aria-busy="status === 'loading' ? 'true' : undefined"
    :data-pptx-thumbnail-index="slide.index"
    :data-thumbnail-state="status"
    @click="emit('select', slide.index)"
  >
    <span class="pptx-thumbnail__frame">
      <span ref="canvasRef" class="pptx-thumbnail__canvas" />
      <OfficeSkeleton
        v-if="status === 'loading'"
        class="pptx-thumbnail__skeleton"
        data-testid="pptx-thumbnail-skeleton"
      />
      <span
        v-else-if="status === 'error'"
        class="pptx-thumbnail__error"
        data-testid="pptx-thumbnail-error"
        role="img"
        :aria-label="m.pptx.thumbnail.failed"
      >
        <OfficeIcon name="triangle-alert" />
        <span class="pptx-thumbnail__error-label">{{ m.pptx.thumbnail.failed }}</span>
      </span>
    </span>
    <span v-if="showLabel" class="pptx-thumbnail__label">
      <span>{{ slide.number }}</span>
      <OfficeBadge
        v-if="slide.hidden"
        variant="secondary"
        class="pptx-thumbnail__hidden"
        data-testid="pptx-thumbnail-hidden"
      >{{ m.pptx.thumbnail.hidden }}</OfficeBadge>
    </span>
  </button>
</template>

<script setup lang="ts">
import { computed, inject, onBeforeUnmount, onMounted, ref, watch } from "vue"
import type {
  PptxDisposableHandle,
  PptxPreviewSession,
} from "@arcships/pptx-core/browser"
import type { PptxSlideInfo } from "@arcships/pptx-core"
import { OfficeBadge, OfficeIcon, OfficeSkeleton } from "@arcships/vue-ui"
import {
  PPTX_THUMBNAIL_OBSERVER,
  type PptxThumbnailObserver,
} from "./composables/thumbnail-observer"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { usePptxMessages } from "./i18n"

const props = withDefaults(defineProps<{
  session: PptxPreviewSession
  slide: PptxSlideInfo
  width?: number
  active?: boolean
  /** Render the slide-number label under the canvas (the sorter card owns it). */
  showLabel?: boolean
}>(), {
  width: 148,
  active: false,
  showLabel: true,
})

const emit = defineEmits<{ select: [index: number] }>()
const m = usePptxMessages()
const thumbnailLabel = computed(() => formatOfficeMessage(
  props.slide.hidden ? m.value.pptx.thumbnail.hiddenLabel : m.value.pptx.thumbnail.label,
  { number: props.slide.number },
))
const rootRef = ref<HTMLElement | null>(null)
const canvasRef = ref<HTMLElement | null>(null)
/**
 * Thumbnail lifecycle for the shared skeleton/error chrome:
 * `idle` never rendered (the session returned no handle), `loading` awaits the
 * renderer, `ready` shows the painted slide and `error` shows the fallback.
 */
const status = ref<"idle" | "loading" | "ready" | "error">("idle")
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
  status.value = "idle"
  canvasRef.value?.replaceChildren?.()
}

async function render(): Promise<void> {
  const target = canvasRef.value
  if (!target || handle) return
  const id = ++renderId
  const nextHandle = props.session.renderThumbnail(props.slide.index, target, props.width)
  if (!nextHandle) return
  handle = nextHandle
  status.value = "loading"
  try {
    await nextHandle.ready
    if (id === renderId) status.value = "ready"
  } catch {
    if (id === renderId) status.value = "error"
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
