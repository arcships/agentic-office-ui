<template>
  <DocxViewerRoot
    ref="viewerRootRef"
    :model="model"
    :controller="controller"
    :editable="editable"
    :layout-options="layoutOptions"
    :theme="theme"
    :zoom-scale="effectiveZoomScale"
    :fit-width="zoom === undefined ? fitWidth : false"
    :show-tracked-changes="showTrackedChanges"
    :show-comments="showComments"
    :search-query="searchQuery"
    :active-search-node-index="activeSearchNodeIndex"
    @page-count-change="emit('pageCountChange', $event)"
    @visible-page-range="emit('visiblePageRange', $event)"
    @context-menu="emit('contextMenu', $event)"
    @selection-change="emit('selectionChange', $event)"
  />
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from "vue"
import { clampSurfaceZoom, nextSurfaceZoom } from "@arcships/office-runtime/gesture-zoom"
import type {
  DocModel,
  DocxDocumentTheme,
  DocxEditorController,
  LayoutOptions,
} from "@arcships/docx-core"
import DocxViewerRoot from "./DocxViewerRoot.vue"

const props = withDefaults(
  defineProps<{
    model: DocModel
    controller?: DocxEditorController
    editable?: boolean
    layoutOptions?: LayoutOptions
    theme?: DocxDocumentTheme
    /** Controlled zoom factor. 1 = 100%. */
    zoom?: number
    /** Legacy percentage zoom. Ignored when zoom is provided. */
    zoomScale?: number
    enableGestureZoom?: boolean
    fitWidth?: boolean
    showTrackedChanges?: boolean
    showComments?: boolean
    searchQuery?: string
    activeSearchNodeIndex?: number
  }>(),
  { editable: false, zoomScale: 100, enableGestureZoom: true }
)

const emit = defineEmits<{
  pageCountChange: [count: number]
  visiblePageRange: [range: { startPageIndex: number; endPageIndex: number }]
  contextMenu: [ctx: { pageIndex: number; clientX: number; clientY: number; containerX: number; containerY: number }]
  selectionChange: [sel: { kind: string; text?: string; nodeIndex?: number }]
  objectClick: [obj: { kind: "image" | "table"; nodeIndex?: number }]
  "update:zoom": [zoom: number]
}>()

const viewerRootRef = ref<InstanceType<typeof DocxViewerRoot>>()
const effectiveZoomScale = computed(() =>
  props.zoom === undefined ? props.zoomScale : clampSurfaceZoom(props.zoom) * 100
)

type PageAnchor = {
  pageIndex: number
  xRatio: number
  yRatio: number
  clientX: number
  clientY: number
}
type WebKitGestureEvent = Event & { clientX?: number; clientY?: number; scale?: number }

let cleanupGestureListeners: (() => void) | undefined
let pendingAnchor: (PageAnchor & { requestedZoom: number; token: number }) | undefined
let pendingZoom: number | undefined
let gestureStartZoom = 1
let gestureToken = 0
let webkitGestureActive = false

function controlledZoom(): number | undefined {
  return typeof props.zoom === "number" && Number.isFinite(props.zoom)
    ? clampSurfaceZoom(props.zoom)
    : undefined
}

function captureAnchor(clientX: number, clientY: number): PageAnchor | undefined {
  const container = viewerRootRef.value?.scrollContainer
  if (!container || typeof document === "undefined") return undefined
  const page = document.elementFromPoint?.(clientX, clientY)
    ?.closest<HTMLElement>("[data-docx-page-index]")
  const pageIndex = Number(page?.dataset.docxPageIndex)
  if (!page || !Number.isInteger(pageIndex)) return undefined
  const rect = page.getBoundingClientRect()
  if (rect.width <= 0 || rect.height <= 0) return undefined
  return {
    pageIndex,
    xRatio: (clientX - rect.left) / rect.width,
    yRatio: (clientY - rect.top) / rect.height,
    clientX,
    clientY,
  }
}

async function restoreAnchor(anchor: PageAnchor & { token: number }): Promise<void> {
  await nextTick()
  requestAnimationFrame(() => {
    if (anchor.token !== gestureToken) return
    const container = viewerRootRef.value?.scrollContainer
    const page = container?.querySelector<HTMLElement>(`[data-docx-page-index="${anchor.pageIndex}"]`)
    if (!container || !page) return
    const rect = page.getBoundingClientRect()
    container.scrollLeft += rect.left + rect.width * anchor.xRatio - anchor.clientX
    container.scrollTop += rect.top + rect.height * anchor.yRatio - anchor.clientY
  })
}

function requestGestureZoom(nextZoom: number, clientX: number, clientY: number): void {
  const current = controlledZoom()
  if (current === undefined || nextZoom === (pendingZoom ?? current)) return
  const token = ++gestureToken
  const anchor = captureAnchor(clientX, clientY)
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
  const rect = viewerRootRef.value?.scrollContainer?.getBoundingClientRect()
  requestGestureZoom(
    clampSurfaceZoom(gestureStartZoom * (event.scale ?? 1)),
    event.clientX ?? (rect ? rect.left + rect.width / 2 : 0),
    event.clientY ?? (rect ? rect.top + rect.height / 2 : 0),
  )
}

function attachGestureListeners(): void {
  cleanupGestureListeners?.()
  const element = viewerRootRef.value?.scrollContainer
  if (!element) return
  const onGestureEnd = () => { webkitGestureActive = false }
  element.addEventListener("wheel", onWheel, { passive: false })
  element.addEventListener("gesturestart", onGestureStart, { passive: false })
  element.addEventListener("gesturechange", onGestureChange as EventListener, { passive: false })
  element.addEventListener("gestureend", onGestureEnd)
  cleanupGestureListeners = () => {
    element.removeEventListener("wheel", onWheel)
    element.removeEventListener("gesturestart", onGestureStart)
    element.removeEventListener("gesturechange", onGestureChange as EventListener)
    element.removeEventListener("gestureend", onGestureEnd)
  }
}

watch(() => props.zoom, (next) => {
  if (pendingAnchor && typeof next === "number" && clampSurfaceZoom(next) === pendingAnchor.requestedZoom) {
    void restoreAnchor(pendingAnchor)
  } else if (pendingZoom !== undefined) {
    gestureToken += 1
  }
  pendingAnchor = undefined
  pendingZoom = undefined
})

watch(() => props.model, () => {
  gestureToken += 1
  pendingAnchor = undefined
  pendingZoom = undefined
})

onMounted(() => { void nextTick(attachGestureListeners) })
onBeforeUnmount(() => {
  gestureToken += 1
  cleanupGestureListeners?.()
})

defineExpose({
  scrollToPage(pageIndex: number): void {
    viewerRootRef.value?.scrollToPage(pageIndex)
  },
  scrollToNode(nodeIndex: number): void {
    viewerRootRef.value?.scrollToNode(nodeIndex)
  },
  get scrollContainer() {
    return viewerRootRef.value?.scrollContainer
  },
})
</script>
