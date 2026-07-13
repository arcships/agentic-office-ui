<template>
  <div
    ref="elementRef"
    class="pptx-stage"
    data-testid="pptx-stage"
    @click="onClick"
    @contextmenu.prevent="onContextMenu"
  />
</template>

<script setup lang="ts">
import { nextTick, onBeforeUnmount, onMounted, ref, watch } from "vue"
import { clampSurfaceZoom, nextSurfaceZoom } from "@arcships/office-runtime/gesture-zoom"
import type {
  PptxStageContextMenu,
  PptxStageObjectClick,
  PptxStageSelection,
} from "./headless-types"

const props = withDefaults(defineProps<{
  /** Controlled zoom factor. 1 = 100%. */
  zoom?: number
  enableGestureZoom?: boolean
  /** Explicit owner of list scrolling. */
  scrollContainer?: HTMLElement | null
}>(), {
  enableGestureZoom: true,
  scrollContainer: null,
})

const elementRef = ref<HTMLElement | null>(null)

const emit = defineEmits<{
  contextMenu: [ctx: PptxStageContextMenu]
  selectionChange: [selection: PptxStageSelection]
  objectClick: [object: PptxStageObjectClick]
  "update:zoom": [zoom: number]
}>()

interface StageHit {
  slideIndex: number
  objectKey?: string
}

type SlideAnchor = {
  slideIndex: number
  xRatio: number
  yRatio: number
  clientX: number
  clientY: number
}
type WebKitGestureEvent = Event & { clientX?: number; clientY?: number; scale?: number }

let pendingAnchor: (SlideAnchor & { requestedZoom: number; token: number }) | undefined
let pendingZoom: number | undefined
let gestureStartZoom = 1
let gestureToken = 0
let webkitGestureActive = false

function controlledZoom(): number | undefined {
  return typeof props.zoom === "number" && Number.isFinite(props.zoom)
    ? clampSurfaceZoom(props.zoom)
    : undefined
}

function resolveHit(event: MouseEvent): StageHit | null {
  const target = event.target as Element | null
  if (!target || typeof target.closest !== "function") return null
  const slideElement = target.closest<HTMLElement>("[data-slide-index]")
  const slideIndex = Number(slideElement?.dataset.slideIndex)
  if (!Number.isInteger(slideIndex) || slideIndex < 0) return null
  const objectElement = target.closest<HTMLElement>("[data-pptx-object-key]")
  return {
    slideIndex,
    objectKey: objectElement?.dataset.pptxObjectKey || undefined,
  }
}

function captureZoomAnchor(clientX: number, clientY: number): SlideAnchor | undefined {
  if (typeof document === "undefined") return undefined
  const slide = document.elementFromPoint?.(clientX, clientY)
    ?.closest<HTMLElement>("[data-slide-index]")
  const slideIndex = Number(slide?.dataset.slideIndex)
  if (!slide || !Number.isInteger(slideIndex)) return undefined
  const rect = slide.getBoundingClientRect()
  if (rect.width <= 0 || rect.height <= 0) return undefined
  return {
    slideIndex,
    xRatio: (clientX - rect.left) / rect.width,
    yRatio: (clientY - rect.top) / rect.height,
    clientX,
    clientY,
  }
}

async function restoreZoomAnchor(anchor: SlideAnchor & { token: number }): Promise<void> {
  await nextTick()
  requestAnimationFrame(() => {
    if (anchor.token !== gestureToken) return
    const container = props.scrollContainer
    const slide = elementRef.value?.querySelector<HTMLElement>(`[data-slide-index="${anchor.slideIndex}"]`)
    if (!container || !slide) return
    const rect = slide.getBoundingClientRect()
    container.scrollLeft += rect.left + rect.width * anchor.xRatio - anchor.clientX
    container.scrollTop += rect.top + rect.height * anchor.yRatio - anchor.clientY
  })
}

function requestGestureZoom(nextZoom: number, clientX: number, clientY: number): void {
  const current = controlledZoom()
  if (current === undefined || nextZoom === (pendingZoom ?? current)) return
  const token = ++gestureToken
  const anchor = captureZoomAnchor(clientX, clientY)
  pendingZoom = nextZoom
  pendingAnchor = anchor ? { ...anchor, requestedZoom: nextZoom, token } : undefined
  emit("update:zoom", nextZoom)
}

function onWheel(event: WheelEvent): void {
  const current = controlledZoom()
  if (props.enableGestureZoom === false || current === undefined || !event.ctrlKey) return
  event.preventDefault()
  if (webkitGestureActive) return
  requestGestureZoom(
    nextSurfaceZoom(pendingZoom ?? current, event.deltaY, event.deltaMode),
    event.clientX,
    event.clientY,
  )
}

function onGestureStart(event: Event): void {
  const current = controlledZoom()
  if (props.enableGestureZoom === false || current === undefined) return
  event.preventDefault()
  webkitGestureActive = true
  gestureStartZoom = current
  pendingZoom = current
}

function onGestureChange(event: WebKitGestureEvent): void {
  if (!webkitGestureActive) return
  event.preventDefault()
  const rect = elementRef.value?.getBoundingClientRect()
  requestGestureZoom(
    clampSurfaceZoom(gestureStartZoom * (event.scale ?? 1)),
    event.clientX ?? (rect ? rect.left + rect.width / 2 : 0),
    event.clientY ?? (rect ? rect.top + rect.height / 2 : 0),
  )
}

function onGestureEnd(): void {
  webkitGestureActive = false
}

function onClick(event: MouseEvent): void {
  const hit = resolveHit(event)
  if (!hit) return
  emit("selectionChange", { kind: "slide", slideIndex: hit.slideIndex })
  if (hit.objectKey) {
    emit("objectClick", {
      kind: "object",
      slideIndex: hit.slideIndex,
      objectKey: hit.objectKey,
    })
  }
}

function onContextMenu(event: MouseEvent): void {
  const hit = resolveHit(event)
  if (!hit) return
  const rect = elementRef.value?.getBoundingClientRect?.()
  const position = {
    slideIndex: hit.slideIndex,
    clientX: event.clientX,
    clientY: event.clientY,
    containerX: rect ? event.clientX - rect.left : 0,
    containerY: rect ? event.clientY - rect.top : 0,
  }
  emit("contextMenu", hit.objectKey
    ? { ...position, kind: "object", objectKey: hit.objectKey }
    : { ...position, kind: "slide" })
}

watch(() => props.zoom, (next) => {
  if (pendingAnchor && typeof next === "number" && clampSurfaceZoom(next) === pendingAnchor.requestedZoom) {
    void restoreZoomAnchor(pendingAnchor)
  } else if (pendingZoom !== undefined) {
    gestureToken += 1
  }
  pendingAnchor = undefined
  pendingZoom = undefined
})

onMounted(() => {
  const element = elementRef.value
  if (!element) return
  element.addEventListener("wheel", onWheel, { passive: false })
  element.addEventListener("gesturestart", onGestureStart, { passive: false })
  element.addEventListener("gesturechange", onGestureChange as EventListener, { passive: false })
  element.addEventListener("gestureend", onGestureEnd)
})

onBeforeUnmount(() => {
  gestureToken += 1
  const element = elementRef.value
  element?.removeEventListener("wheel", onWheel)
  element?.removeEventListener("gesturestart", onGestureStart)
  element?.removeEventListener("gesturechange", onGestureChange as EventListener)
  element?.removeEventListener("gestureend", onGestureEnd)
})

defineExpose({
  get element(): HTMLElement | null {
    return elementRef.value
  },
  get scrollContainer(): HTMLElement | null {
    return props.scrollContainer
  },
})
</script>

<style scoped>
.pptx-stage {
  background: var(--pptx-surface-bg, transparent);
}
</style>
