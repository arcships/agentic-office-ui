<template>
  <div ref="surfaceRef" class="docx-document-surface">
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
      @page-count-change="emit('pageCountChange', $event)"
      @visible-page-range="onVisiblePageRange"
      @context-menu="emit('contextMenu', $event)"
      @selection-change="emit('selectionChange', $event)"
    />
    <div class="docx-document-surface__search-layer" aria-hidden="true">
      <div
        v-for="highlight in searchHighlights"
        :key="highlight.key"
        class="docx-document-surface__search-highlight"
        :class="{ 'docx-document-surface__search-highlight--active': highlight.active }"
        :style="{
          left: `${highlight.left}px`,
          top: `${highlight.top}px`,
          width: `${highlight.width}px`,
          height: `${highlight.height}px`,
        }"
      />
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, watch } from "vue"
import { clampSurfaceZoom, nextSurfaceZoom } from "@arcships/office-runtime/gesture-zoom"
import {
  createSurfaceSearchSession,
} from "@arcships/office-runtime"
import type {
  DocModel,
  DocxDocumentTheme,
  DocxEditorController,
  DocxTextRangeLocation,
  LayoutOptions,
} from "@arcships/docx-core"
import {
  findDocxSearchMatches,
  type DocxSearchMatch,
  type DocxSearchState,
} from "../composables/useDocxSearch"
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
  }>(),
  { editable: false, zoomScale: 100, enableGestureZoom: true }
)

const emit = defineEmits<{
  pageCountChange: [count: number]
  visiblePageRange: [range: { startPageIndex: number; endPageIndex: number }]
  contextMenu: [ctx: { pageIndex: number; clientX: number; clientY: number; containerX: number; containerY: number }]
  selectionChange: [sel: { kind: string; text?: string; nodeIndex?: number }]
  objectClick: [obj: { kind: "image" | "table"; nodeIndex?: number }]
  searchStateChange: [state: DocxSearchState]
  "update:zoom": [zoom: number]
}>()

const surfaceRef = ref<HTMLElement>()
const viewerRootRef = ref<InstanceType<typeof DocxViewerRoot>>()
const effectiveZoomScale = computed(() =>
  props.zoom === undefined ? props.zoomScale : clampSurfaceZoom(props.zoom) * 100
)

interface SearchHighlight {
  key: string
  left: number
  top: number
  width: number
  height: number
  active: boolean
}

interface DomPoint {
  node: Node
  offset: number
}

interface DomTextSegment {
  start: number
  end: number
  startPoint: DomPoint
  endPoint: DomPoint
}

const searchHighlights = shallowRef<SearchHighlight[]>([])
let highlightFrame = 0
let cleanupSearchSubscription: (() => void) | undefined
let cleanupSearchScrollListener: (() => void) | undefined

function throwIfAborted(signal: AbortSignal): void {
  if (!signal.aborted) return
  if (typeof DOMException !== "undefined") throw new DOMException("DOCX search aborted.", "AbortError")
  const error = new Error("DOCX search aborted.")
  error.name = "AbortError"
  throw error
}

function nextFrame(): Promise<void> {
  if (typeof requestAnimationFrame !== "function") return Promise.resolve()
  return new Promise((resolve) => requestAnimationFrame(() => resolve()))
}

function selectorForLocation(location: DocxTextRangeLocation): string {
  if (location.kind === "paragraph") {
    return `[data-docx-paragraph-node-index="${location.nodeIndex}"]`
  }
  return [
    `[data-docx-table-cell-paragraph-host="true"]`,
    `[data-docx-table-index="${location.tableIndex}"]`,
    `[data-docx-table-row-index="${location.rowIndex}"]`,
    `[data-docx-table-cell-index="${location.cellIndex}"]`,
    `[data-docx-paragraph-index="${location.paragraphIndex}"]`,
  ].join("")
}

function isExcludedTextElement(element: Element): boolean {
  return element.matches(
    "[data-docx-numbering-label='true'], [data-docx-image-child-index], [data-docx-form-field='true'], [data-docx-tracked-change='deletion']",
  )
}

function collectDomTextSegments(host: HTMLElement): DomTextSegment[] {
  const segments: DomTextSegment[] = []
  let cursor = 0

  function visit(node: Node): void {
    if (node.nodeType === Node.TEXT_NODE) {
      const length = node.textContent?.length ?? 0
      if (!length) return
      segments.push({
        start: cursor,
        end: cursor + length,
        startPoint: { node, offset: 0 },
        endPoint: { node, offset: length },
      })
      cursor += length
      return
    }
    if (!(node instanceof Element) || isExcludedTextElement(node)) return
    if (node.tagName.toLowerCase() === "br") {
      const parent = node.parentNode
      if (!parent) return
      const index = Array.prototype.indexOf.call(parent.childNodes, node) as number
      segments.push({
        start: cursor,
        end: cursor + 1,
        startPoint: { node: parent, offset: index },
        endPoint: { node: parent, offset: index + 1 },
      })
      cursor += 1
      return
    }
    for (const child of node.childNodes) visit(child)
  }

  for (const child of host.childNodes) visit(child)
  return segments
}

function pointAtOffset(
  host: HTMLElement,
  segments: readonly DomTextSegment[],
  offset: number,
): DomPoint | undefined {
  if (!segments.length) return offset === 0 ? { node: host, offset: 0 } : undefined
  for (const segment of segments) {
    if (offset < segment.start || offset > segment.end) continue
    if (segment.startPoint.node === segment.endPoint.node && segment.startPoint.node.nodeType === Node.TEXT_NODE) {
      return { node: segment.startPoint.node, offset: offset - segment.start }
    }
    if (offset === segment.start) return segment.startPoint
    if (offset === segment.end) return segment.endPoint
  }
  const last = segments[segments.length - 1]
  if (offset === last.end) return last.endPoint
  return undefined
}

function rangeForHost(host: HTMLElement, match: DocxSearchMatch): Range | undefined {
  const segments = collectDomTextSegments(host)
  const start = pointAtOffset(host, segments, match.range.start.offset)
  const end = pointAtOffset(host, segments, match.range.end.offset)
  if (!start || !end || typeof document === "undefined") return undefined
  try {
    const range = document.createRange()
    range.setStart(start.node, start.offset)
    range.setEnd(end.node, end.offset)
    return range
  } catch {
    return undefined
  }
}

function rangesForMatch(match: DocxSearchMatch): Array<{ host: HTMLElement; range: Range }> {
  const root = surfaceRef.value
  if (!root) return []
  const selector = selectorForLocation(match.range.start.location)
  const ranges: Array<{ host: HTMLElement; range: Range }> = []
  for (const host of root.querySelectorAll<HTMLElement>(selector)) {
    const range = rangeForHost(host, match)
    if (range) ranges.push({ host, range })
  }
  return ranges
}

function intersectionRect(
  rect: DOMRect,
  hostRect: DOMRect,
  viewportRect: DOMRect,
): { left: number; top: number; right: number; bottom: number } | undefined {
  const left = Math.max(rect.left, hostRect.left, viewportRect.left)
  const top = Math.max(rect.top, hostRect.top, viewportRect.top)
  const right = Math.min(rect.right, hostRect.right, viewportRect.right)
  const bottom = Math.min(rect.bottom, hostRect.bottom, viewportRect.bottom)
  if (right - left < 0.5 || bottom - top < 0.5) return undefined
  return { left, top, right, bottom }
}

function renderSearchHighlights(): void {
  const root = surfaceRef.value
  const container = viewerRootRef.value?.scrollContainer
  const state = surfaceSearch.getSearchState()
  if (!root || !container || !state.matches.length) {
    searchHighlights.value = []
    return
  }

  const rootRect = root.getBoundingClientRect()
  const viewportRect = container.getBoundingClientRect()
  const seen = new Set<string>()
  const next: SearchHighlight[] = []
  state.matches.forEach((match, matchIndex) => {
    for (const { host, range } of rangesForMatch(match)) {
      const hostRect = host.getBoundingClientRect()
      for (const rect of Array.from(range.getClientRects())) {
        const clipped = intersectionRect(rect, hostRect, viewportRect)
        if (!clipped) continue
        const left = clipped.left - rootRect.left
        const top = clipped.top - rootRect.top
        const width = clipped.right - clipped.left
        const height = clipped.bottom - clipped.top
        const rectKey = `${matchIndex}:${Math.round(left * 10)}:${Math.round(top * 10)}:${Math.round(width * 10)}:${Math.round(height * 10)}`
        if (seen.has(rectKey)) continue
        seen.add(rectKey)
        next.push({
          key: rectKey,
          left,
          top,
          width,
          height,
          active: matchIndex === state.activeIndex,
        })
      }
    }
  })
  searchHighlights.value = next
}

function scheduleSearchHighlights(): void {
  if (highlightFrame && typeof cancelAnimationFrame === "function") {
    cancelAnimationFrame(highlightFrame)
  }
  if (typeof requestAnimationFrame !== "function") {
    renderSearchHighlights()
    return
  }
  highlightFrame = requestAnimationFrame(() => {
    highlightFrame = 0
    renderSearchHighlights()
  })
}

async function activateSearchMatchInView(match: DocxSearchMatch, signal: AbortSignal): Promise<void> {
  const container = viewerRootRef.value?.scrollContainer
  let candidates = rangesForMatch(match)
  let candidate = container
    ? candidates.find(({ range }) => {
        const rect = range.getClientRects()[0]
        const viewport = container.getBoundingClientRect()
        return rect && rect.top >= viewport.top && rect.bottom <= viewport.bottom
      })
    : undefined
  if (candidate) {
    scheduleSearchHighlights()
    return
  }

  if (!candidates.length) {
    viewerRootRef.value?.scrollToNode(match.nodeIndex, { behavior: "auto" })
    await nextTick()
    await nextFrame()
    throwIfAborted(signal)
    candidates = rangesForMatch(match)
  }

  if (container && candidates.length) {
    const viewport = container.getBoundingClientRect()
    const center = viewport.top + viewport.height / 2
    candidate = [...candidates].sort((left, right) => {
      const leftRect = left.range.getClientRects()[0]
      const rightRect = right.range.getClientRects()[0]
      return Math.abs((leftRect?.top ?? Infinity) - center) - Math.abs((rightRect?.top ?? Infinity) - center)
    })[0]
  } else {
    candidate = candidates[0]
  }
  if (!candidate) {
    await nextFrame()
    throwIfAborted(signal)
    candidate = rangesForMatch(match)[0]
  }

  if (!candidate) {
    throw new Error("DOCX search match could not be projected into the mounted document.")
  }

  const rect = candidate.range.getClientRects()[0]
  if (container && rect) {
    const viewport = container.getBoundingClientRect()
    if (rect.top < viewport.top || rect.bottom > viewport.bottom) {
      container.scrollBy({
        top: rect.top - viewport.top - (viewport.height - rect.height) / 2,
        behavior: "auto",
      })
      await nextFrame()
      throwIfAborted(signal)
    }
  }
  scheduleSearchHighlights()
}

const surfaceSearch = createSurfaceSearchSession<DocxSearchMatch>({
  search(query, { signal }) {
    return findDocxSearchMatches(props.model, query, signal)
  },
  activate(match, _index, { signal }) {
    return activateSearchMatchInView(match, signal)
  },
  clear() {
    searchHighlights.value = []
  },
})
const searchState = shallowRef<DocxSearchState>(surfaceSearch.getSearchState())
cleanupSearchSubscription = surfaceSearch.subscribe((next) => {
  searchState.value = next
  emit("searchStateChange", next)
  scheduleSearchHighlights()
})

function onVisiblePageRange(range: { startPageIndex: number; endPageIndex: number }): void {
  emit("visiblePageRange", range)
  scheduleSearchHighlights()
}

function attachSearchScrollListener(): void {
  cleanupSearchScrollListener?.()
  const container = viewerRootRef.value?.scrollContainer
  if (!container) return
  const onScroll = () => scheduleSearchHighlights()
  container.addEventListener("scroll", onScroll, { passive: true })
  cleanupSearchScrollListener = () => container.removeEventListener("scroll", onScroll)
}

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
  surfaceSearch.clearSearch()
})

watch(
  () => [effectiveZoomScale.value, props.fitWidth, props.showTrackedChanges, props.showComments],
  () => { void nextTick(scheduleSearchHighlights) },
)

onMounted(() => {
  void nextTick(() => {
    attachGestureListeners()
    attachSearchScrollListener()
    scheduleSearchHighlights()
  })
})
onBeforeUnmount(() => {
  gestureToken += 1
  cleanupGestureListeners?.()
  cleanupSearchScrollListener?.()
  cleanupSearchSubscription?.()
  surfaceSearch.dispose()
  if (highlightFrame && typeof cancelAnimationFrame === "function") {
    cancelAnimationFrame(highlightFrame)
  }
})

defineExpose({
  scrollToPage(pageIndex: number): void {
    viewerRootRef.value?.scrollToPage(pageIndex)
  },
  scrollToNode(nodeIndex: number, options?: ScrollToOptions): void {
    viewerRootRef.value?.scrollToNode(nodeIndex, options)
  },
  search: (query: string): Promise<DocxSearchState> => surfaceSearch.search(query),
  activateSearchMatch: (index: number) => surfaceSearch.activateSearchMatch(index),
  searchNext: () => surfaceSearch.searchNext(),
  searchPrevious: () => surfaceSearch.searchPrevious(),
  clearSearch: () => surfaceSearch.clearSearch(),
  getSearchState: (): DocxSearchState => surfaceSearch.getSearchState(),
  get searchState() {
    return searchState.value
  },
  get scrollContainer() {
    return viewerRootRef.value?.scrollContainer
  },
})
</script>

<style scoped>
.docx-document-surface {
  display: flex;
  flex: 1;
  min-height: 0;
  min-width: 0;
  overflow: hidden;
  position: relative;
}

.docx-document-surface__search-layer {
  inset: 0;
  overflow: hidden;
  pointer-events: none;
  position: absolute;
  z-index: 8;
}

.docx-document-surface__search-highlight {
  background: rgb(250 204 21 / 42%);
  border-radius: 2px;
  position: absolute;
}

.docx-document-surface__search-highlight--active {
  background: rgb(245 158 11 / 58%);
  box-shadow: inset 0 0 0 1px rgb(180 83 9 / 76%);
}
</style>
