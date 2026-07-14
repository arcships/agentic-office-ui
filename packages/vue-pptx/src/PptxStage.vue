<template>
  <div
    ref="hostRef"
    class="pptx-stage"
    data-testid="pptx-stage"
    @click="onClick"
    @pointermove="onReferencePointerMove"
    @mouseup="onTextSelectionMouseUp"
    @keydown.esc="onReferenceEscape"
    @contextmenu.prevent="onContextMenu"
  >
    <div ref="elementRef" class="pptx-stage__content" />
    <div
      v-if="selectionMode === 'region'"
      class="pptx-stage__region-layer"
      aria-label="选择幻灯片区域"
      tabindex="0"
      @pointerdown="onRegionPointerDown"
      @pointermove="onRegionPointerMove"
      @pointerup="onRegionPointerUp"
      @pointercancel="onRegionPointerCancel"
      @keydown.esc="onRegionKeyboardCancel"
    >
      <div v-if="regionFrame" class="pptx-stage__region-frame" :style="regionFrame" />
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from "vue"
import { clampSurfaceZoom, nextSurfaceZoom } from "@arcships/office-runtime/gesture-zoom"
import type {
  PptxStageContextMenu,
  PptxStageObjectClick,
  PptxStageSelection,
} from "./headless-types"
import {
  createPptxObjectReferenceDraft,
  createPptxRegionReferenceDraft,
  createPptxSlideReferenceDraft,
  createPptxTextReferenceDraft,
  describePptxReference,
  resolvePptxReference,
  type PptxObjectIdentity,
  type PptxOfficeReference,
  type PptxPlaybackDocument,
  type PptxPreviewDocument,
  type PptxReferenceContext,
} from "@arcships/pptx-core"
import {
  confirmOfficeReferenceDraft,
  createOfficeReferenceId,
  type NormalizedRect,
  type OfficeDocumentRevision,
  type OfficeObjectReference,
  type OfficeReferenceCandidatePreview,
  type OfficeReferenceError,
  type OfficeSelectionMode,
  type ResolveReferenceResult,
} from "@arcships/office-interaction"

const props = withDefaults(defineProps<{
  /** Controlled zoom factor. 1 = 100%. */
  zoom?: number
  enableGestureZoom?: boolean
  /** Explicit owner of list scrolling. */
  scrollContainer?: HTMLElement | null
  documentId?: string
  selectionMode?: OfficeSelectionMode
  emitReferenceCandidates?: boolean
}>(), {
  enableGestureZoom: true,
  scrollContainer: null,
  selectionMode: "content",
  emitReferenceCandidates: false,
})

const elementRef = ref<HTMLElement | null>(null)
const hostRef = ref<HTMLElement | null>(null)
const selectionMode = computed(() => props.selectionMode)
const revisionCounter = ref(0)
const generatedDocumentId = `pptx-stage-${typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2)}`
const documentRevision = computed<OfficeDocumentRevision & { format: "pptx" }>(() => ({
  format: "pptx",
  documentId: props.documentId?.trim() || generatedDocumentId,
  revision: `${generatedDocumentId}:${revisionCounter.value}`,
}))

function slideElements(): HTMLElement[] {
  return [...(elementRef.value?.querySelectorAll<HTMLElement>("[data-slide-index]") ?? [])]
    .filter((element, index, all) => all.findIndex((candidate) => candidate.dataset.slideIndex === element.dataset.slideIndex) === index)
    .sort((left, right) => Number(left.dataset.slideIndex) - Number(right.dataset.slideIndex))
}

function slidePathFor(element: HTMLElement, index: number): string {
  return element.dataset.pptxSlidePath
    || element.querySelector<HTMLElement>("[data-pptx-slide-path]")?.dataset.pptxSlidePath
    || `rendered-slide:${index}`
}

function identitiesForSlide(slide: HTMLElement, slidePath: string): PptxObjectIdentity[] {
  const seen = new Set<string>()
  const identities: PptxObjectIdentity[] = []
  for (const element of slide.querySelectorAll<HTMLElement>("[data-pptx-object-key]")) {
    const key = element.dataset.pptxObjectKey
    if (!key || seen.has(key)) continue
    seen.add(key)
    identities.push({
      key,
      slidePath: element.dataset.pptxPartPath || slidePath,
      source: element.dataset.pptxSource === "layout" || element.dataset.pptxSource === "master" ? element.dataset.pptxSource : "slide",
      shapeId: element.dataset.pptxNodeId || key,
      groupPath: element.dataset.pptxGroupPath?.split("/").filter(Boolean) ?? [],
      ...(element.getAttribute("aria-label") || element.getAttribute("title") ? { name: element.getAttribute("aria-label") || element.getAttribute("title") || undefined } : {}),
      nodeType: element.dataset.pptxNodeType || "shape",
    })
  }
  return identities
}

function paragraphsForObject(slideIndex: number, objectKey: string): string[] | undefined {
  const slide = slideElements().find((element) => Number(element.dataset.slideIndex) === slideIndex)
  const object = [...(slide?.querySelectorAll<HTMLElement>("[data-pptx-object-key]") ?? [])]
    .find((element) => element.dataset.pptxObjectKey === objectKey)
  if (!object) return undefined
  const paragraphs = [...object.querySelectorAll<HTMLElement>("[data-pptx-paragraph-index]")]
    .sort((left, right) => Number(left.dataset.pptxParagraphIndex) - Number(right.dataset.pptxParagraphIndex))
    .map((paragraph) => paragraph.textContent ?? "")
  return paragraphs.length ? paragraphs : [object.textContent ?? ""]
}

let cachedReferenceContext: { revision: string; context: PptxReferenceContext } | undefined

function referenceContext(): PptxReferenceContext {
  const revision = documentRevision.value.revision
  if (cachedReferenceContext?.revision === revision) return cachedReferenceContext.context
  const renderedSlides = slideElements()
  const byIndex = new Map(renderedSlides.map((slide) => [Number(slide.dataset.slideIndex), slide]))
  const count = Math.max(0, ...byIndex.keys()) + (renderedSlides.length ? 1 : 0)
  const slideIndices = Array.from({ length: count }, (_, index) => index)
  const firstRect = renderedSlides[0]?.getBoundingClientRect()
  const preview: PptxPreviewDocument = {
    width: firstRect?.width && firstRect.width > 0 ? firstRect.width : 1280,
    height: firstRect?.height && firstRect.height > 0 ? firstRect.height : 720,
    slides: slideIndices.map((index) => ({ index, number: index + 1, hidden: false })),
    warnings: [],
  }
  const playbackDocument: PptxPlaybackDocument = {
    slides: slideIndices.map((index) => {
      const slide = byIndex.get(index)
      const slidePath = slide ? slidePathFor(slide, index) : `rendered-slide:${index}`
      return {
        index,
        hidden: false,
        objects: slide ? identitiesForSlide(slide, slidePath) : [],
        morphFromPrevious: [], nodes: {}, media: [], actions: [],
        capability: { discovered: 0, strict: 0, approximate: 0, static: 0, unparsed: 0, features: [] },
      }
    }),
    capability: { discovered: 0, strict: 0, approximate: 0, static: 0, unparsed: 0, features: [] },
  }
  const context: PptxReferenceContext = {
    revision: documentRevision.value,
    document: preview,
    playbackDocument,
    slideIds: slideIndices.map((index) => {
      const slide = byIndex.get(index)
      return slide ? slidePathFor(slide, index) : `rendered-slide:${index}`
    }),
    getObjectParagraphs: paragraphsForObject,
  }
  cachedReferenceContext = { revision, context }
  return context
}

function emitReferenceError(operation: OfficeReferenceError["operation"], code: OfficeReferenceError["code"], message: string, referenceId?: string): void {
  emit("referenceError", { code, operation, format: "pptx", recoverable: true, ...(referenceId ? { referenceId } : {}), message })
}

function confirmDraft(
  draft: Parameters<typeof confirmOfficeReferenceDraft>[0],
  trigger: "pointer" | "keyboard" | "touch" | "programmatic",
  snapshot?: Parameters<typeof confirmOfficeReferenceDraft>[1]["snapshot"],
  additiveRequested = false,
): void {
  try {
    emit("referenceConfirm", confirmOfficeReferenceDraft(draft, { referenceId: createOfficeReferenceId(), trigger, additiveRequested, ...(snapshot ? { snapshot } : {}) }))
  } catch (reason) {
    emitReferenceError("describe", "INVALID_REFERENCE", reason instanceof Error ? reason.message : "Unable to confirm PPTX reference.")
  }
}

const emit = defineEmits<{
  contextMenu: [ctx: PptxStageContextMenu]
  selectionChange: [selection: PptxStageSelection]
  objectClick: [object: PptxStageObjectClick]
  "update:zoom": [zoom: number]
  documentRevisionChange: [revision: OfficeDocumentRevision]
  referenceCandidateChange: [change: { candidates: readonly OfficeReferenceCandidatePreview[]; activeCandidateId?: string }]
  referenceConfirm: [event: ReturnType<typeof confirmOfficeReferenceDraft>]
  regionDraftChange: [event: { phase: "start" | "change"; region: { space: "slide"; slideIndex: number; rect: NormalizedRect } }]
  selectionCancel: [event: { mode: OfficeSelectionMode; reason: "escape" | "pointer-cancel" | "programmatic" }]
  referenceResolve: [event: { referenceId: string; result: ResolveReferenceResult }]
  referenceError: [error: OfficeReferenceError]
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
let contentObserver: MutationObserver | undefined
let revisionScheduled = false

function scheduleRevisionChange(): void {
  cachedReferenceContext = undefined
  if (revisionScheduled) return
  revisionScheduled = true
  queueMicrotask(() => {
    revisionScheduled = false
    revisionCounter.value += 1
    emit("documentRevisionChange", documentRevision.value)
  })
}

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

function candidateForHit(hit: StageHit): OfficeReferenceCandidatePreview | undefined {
  const context = referenceContext()
  if (!context.document.slides[hit.slideIndex]) return undefined
  if (hit.objectKey) {
    const draft = createPptxObjectReferenceDraft(context, hit.slideIndex, hit.objectKey)
    const identity = context.playbackDocument.slides[hit.slideIndex]?.objects.find((object) => object.key === hit.objectKey)
    const label = identity?.name || `${draft.kind} ${identity?.shapeId ?? ""}`.trim()
    return {
      candidateId: `pptx:object:${hit.objectKey}`,
      draft,
      preview: {
        label,
        path: [{ kind: "document", label: "Presentation" }, { kind: "slide", label: `Slide ${hit.slideIndex + 1}` }, { kind: draft.kind, label }],
      },
      hit: "direct",
      depth: 2,
    }
  }
  const draft = createPptxSlideReferenceDraft(context, hit.slideIndex)
  return {
    candidateId: `pptx:slide:${hit.slideIndex}`,
    draft,
    preview: { label: `Slide ${hit.slideIndex + 1}`, path: [{ kind: "document", label: "Presentation" }, { kind: "slide", label: `Slide ${hit.slideIndex + 1}` }] },
    hit: "inside",
    depth: 1,
  }
}

function hitTest(point: { clientX: number; clientY: number }): readonly OfficeReferenceCandidatePreview[] {
  try {
    const target = elementRef.value?.ownerDocument?.elementFromPoint?.(point.clientX, point.clientY)
    if (!target) return []
    const slide = target.closest<HTMLElement>("[data-slide-index]")
    const slideIndex = Number(slide?.dataset.slideIndex)
    if (!Number.isSafeInteger(slideIndex) || slideIndex < 0) return []
    const objectKey = target.closest<HTMLElement>("[data-pptx-object-key]")?.dataset.pptxObjectKey
    const candidate = candidateForHit({ slideIndex, ...(objectKey ? { objectKey } : {}) })
    if (!candidate) return []
    if (!objectKey) return [candidate]
    const slideCandidate = candidateForHit({ slideIndex })
    return slideCandidate ? [candidate, slideCandidate] : [candidate]
  } catch (reason) {
    emitReferenceError("hit-test", "HIT_TEST_FAILED", reason instanceof Error ? reason.message : "PPTX hit test failed.")
    return []
  }
}

function onReferencePointerMove(event: PointerEvent): void {
  if (selectionMode.value !== "object" || !props.emitReferenceCandidates) return
  const candidates = hitTest(event)
  emit("referenceCandidateChange", { candidates, ...(candidates[0] ? { activeCandidateId: candidates[0].candidateId } : {}) })
}

function elementForNode(node: Node): HTMLElement | null {
  return node.nodeType === Node.ELEMENT_NODE ? node as HTMLElement : node.parentElement
}

function offsetWithin(host: HTMLElement, node: Node, offset: number): number | undefined {
  try {
    const range = host.ownerDocument.createRange()
    range.setStart(host, 0)
    range.setEnd(node, offset)
    return range.toString().length
  } catch { return undefined }
}

function paragraphForNode(node: Node, object: HTMLElement): { element: HTMLElement; index: number } | undefined {
  const marked = elementForNode(node)?.closest<HTMLElement>("[data-pptx-paragraph-index]")
  if (marked && object.contains(marked)) {
    const index = Number(marked.dataset.pptxParagraphIndex)
    if (Number.isSafeInteger(index) && index >= 0) return { element: marked, index }
  }
  return object.contains(node) ? { element: object, index: 0 } : undefined
}

let ignoreNextObjectClick = false

function onTextSelectionMouseUp(): void {
  if (selectionMode.value !== "content") return
  const selection = elementRef.value?.ownerDocument?.getSelection?.()
  if (!selection || selection.rangeCount === 0 || selection.isCollapsed || !selection.toString().trim()) return
  const range = selection.getRangeAt(0)
  const startObject = elementForNode(range.startContainer)?.closest<HTMLElement>("[data-pptx-object-key]")
  const endObject = elementForNode(range.endContainer)?.closest<HTMLElement>("[data-pptx-object-key]")
  const slide = startObject?.closest<HTMLElement>("[data-slide-index]")
  const slideIndex = Number(slide?.dataset.slideIndex)
  if (!startObject || startObject !== endObject || !Number.isSafeInteger(slideIndex)) {
    emitReferenceError("hit-test", "HIT_TEST_FAILED", "PPTX text selection must stay inside one text object.")
    return
  }
  const start = paragraphForNode(range.startContainer, startObject)
  const end = paragraphForNode(range.endContainer, startObject)
  const startOffset = start && offsetWithin(start.element, range.startContainer, range.startOffset)
  const endOffset = end && offsetWithin(end.element, range.endContainer, range.endOffset)
  const objectKey = startObject.dataset.pptxObjectKey
  if (!start || !end || startOffset === undefined || endOffset === undefined || !objectKey) return
  try {
    const draft = createPptxTextReferenceDraft(
      referenceContext(), slideIndex, objectKey,
      { paragraphIndex: start.index, offset: startOffset },
      { paragraphIndex: end.index, offset: endOffset },
    )
    const text = selection.toString()
    confirmDraft(draft, "pointer", {
      label: text.length > 80 ? `${text.slice(0, 77)}…` : text,
      path: [{ kind: "document", label: "Presentation" }, { kind: "slide", label: `Slide ${slideIndex + 1}` }, { kind: "text-range", label: "Selected text" }],
      content: { text },
    })
    ignoreNextObjectClick = true
  } catch (reason) {
    emitReferenceError("describe", "INVALID_REFERENCE", reason instanceof Error ? reason.message : "Unable to create PPTX text reference.")
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
  if (ignoreNextObjectClick) {
    ignoreNextObjectClick = false
    return
  }
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
  if (selectionMode.value === "object") {
    const candidate = candidateForHit(hit)
    if (candidate) confirmDraft(candidate.draft, event.detail === 0 ? "keyboard" : "pointer", candidate.preview, event.shiftKey)
  }
}

function onReferenceEscape(): void {
  if (selectionMode.value !== "content") emit("selectionCancel", { mode: selectionMode.value, reason: "escape" })
}

function slideAtPoint(clientX: number, clientY: number): HTMLElement | undefined {
  const elements = elementRef.value?.ownerDocument?.elementsFromPoint?.(clientX, clientY) ?? []
  for (const element of elements) {
    const slide = element.closest<HTMLElement>("[data-slide-index]")
    if (slide) return slide
  }
  return undefined
}

type RegionGesture = { pointerId: number; slideIndex: number; bounds: DOMRect; startX: number; startY: number }
const regionGesture = ref<RegionGesture | null>(null)
const regionRect = ref<NormalizedRect | null>(null)
const regionFrame = computed(() => {
  const rect = regionRect.value
  const gesture = regionGesture.value
  const root = hostRef.value?.getBoundingClientRect()
  return rect && gesture && root ? {
    left: `${gesture.bounds.left - root.left + rect.x * gesture.bounds.width}px`,
    top: `${gesture.bounds.top - root.top + rect.y * gesture.bounds.height}px`,
    width: `${rect.width * gesture.bounds.width}px`,
    height: `${rect.height * gesture.bounds.height}px`,
  } : undefined
})

function clampUnit(value: number): number { return Math.max(0, Math.min(1, value)) }
function regionRectFor(gesture: RegionGesture, clientX: number, clientY: number): NormalizedRect {
  const x1 = clampUnit((Math.min(gesture.startX, clientX) - gesture.bounds.left) / gesture.bounds.width)
  const y1 = clampUnit((Math.min(gesture.startY, clientY) - gesture.bounds.top) / gesture.bounds.height)
  const x2 = clampUnit((Math.max(gesture.startX, clientX) - gesture.bounds.left) / gesture.bounds.width)
  const y2 = clampUnit((Math.max(gesture.startY, clientY) - gesture.bounds.top) / gesture.bounds.height)
  return { x: x1, y: y1, width: x2 - x1, height: y2 - y1 }
}

function onRegionPointerDown(event: PointerEvent): void {
  if (event.button !== 0) return
  const slide = slideAtPoint(event.clientX, event.clientY)
  const slideIndex = Number(slide?.dataset.slideIndex)
  const bounds = slide?.getBoundingClientRect()
  if (!bounds || bounds.width <= 0 || bounds.height <= 0 || !Number.isSafeInteger(slideIndex)) return
  regionGesture.value = { pointerId: event.pointerId, slideIndex, bounds, startX: event.clientX, startY: event.clientY }
  regionRect.value = { x: clampUnit((event.clientX - bounds.left) / bounds.width), y: clampUnit((event.clientY - bounds.top) / bounds.height), width: 0, height: 0 }
  ;(event.currentTarget as HTMLElement).setPointerCapture?.(event.pointerId)
  emit("regionDraftChange", { phase: "start", region: { space: "slide", slideIndex, rect: regionRect.value } })
  event.preventDefault()
}

function onRegionPointerMove(event: PointerEvent): void {
  const gesture = regionGesture.value
  if (!gesture || gesture.pointerId !== event.pointerId) return
  regionRect.value = regionRectFor(gesture, event.clientX, event.clientY)
  emit("regionDraftChange", { phase: "change", region: { space: "slide", slideIndex: gesture.slideIndex, rect: regionRect.value } })
}

function finishRegion(event: PointerEvent, cancelled: boolean): void {
  const gesture = regionGesture.value
  if (!gesture || gesture.pointerId !== event.pointerId) return
  const target = event.currentTarget as HTMLElement
  if (target.hasPointerCapture?.(event.pointerId)) target.releasePointerCapture?.(event.pointerId)
  const rect = regionRect.value
  regionGesture.value = null
  regionRect.value = null
  if (cancelled || !rect || rect.width < 0.005 || rect.height < 0.005) {
    emit("selectionCancel", { mode: "region", reason: cancelled ? "pointer-cancel" : "programmatic" })
    return
  }
  confirmDraft(createPptxRegionReferenceDraft(referenceContext(), { space: "slide", slideIndex: gesture.slideIndex, rect }), event.pointerType === "touch" ? "touch" : "pointer")
}

function onRegionPointerUp(event: PointerEvent): void { finishRegion(event, false) }
function onRegionPointerCancel(event: PointerEvent): void { finishRegion(event, true) }
function onRegionKeyboardCancel(): void {
  regionGesture.value = null
  regionRect.value = null
  emit("selectionCancel", { mode: "region", reason: "escape" })
}

function onContextMenu(event: MouseEvent): void {
  const hit = resolveHit(event)
  if (!hit) return
  const rect = hostRef.value?.getBoundingClientRect?.()
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
  const host = hostRef.value
  const element = elementRef.value
  if (!host || !element) return
  host.addEventListener("wheel", onWheel, { passive: false })
  host.addEventListener("gesturestart", onGestureStart, { passive: false })
  host.addEventListener("gesturechange", onGestureChange as EventListener, { passive: false })
  host.addEventListener("gestureend", onGestureEnd)
  if (typeof MutationObserver !== "undefined") {
    contentObserver = new MutationObserver(scheduleRevisionChange)
    contentObserver.observe(element, { childList: true, subtree: true, attributes: true, attributeFilter: ["data-slide-index", "data-pptx-object-key", "data-pptx-paragraph-index"] })
  }
  scheduleRevisionChange()
})

watch(() => props.documentId, scheduleRevisionChange)

onBeforeUnmount(() => {
  gestureToken += 1
  const host = hostRef.value
  host?.removeEventListener("wheel", onWheel)
  host?.removeEventListener("gesturestart", onGestureStart)
  host?.removeEventListener("gesturechange", onGestureChange as EventListener)
  host?.removeEventListener("gestureend", onGestureEnd)
  contentObserver?.disconnect()
  contentObserver = undefined
})

defineExpose({
  get element(): HTMLElement | null {
    return elementRef.value
  },
  get scrollContainer(): HTMLElement | null {
    return props.scrollContainer
  },
  getDocumentRevision: (): OfficeDocumentRevision => documentRevision.value,
  hitTest,
  async describeReference(reference: OfficeObjectReference, signal?: AbortSignal) {
    if (signal?.aborted) throw new DOMException("PPTX reference description aborted.", "AbortError")
    const descriptor = describePptxReference(referenceContext(), reference)
    if (!descriptor) throw new Error("PPTX reference is not resolvable in the current revision.")
    return descriptor
  },
  async resolveReference(reference: OfficeObjectReference): Promise<ResolveReferenceResult> {
    const result = resolvePptxReference(referenceContext(), reference)
    emit("referenceResolve", { referenceId: reference.referenceId, result })
    return result
  },
  async scrollToReference(reference: OfficeObjectReference): Promise<void> {
    const result = resolvePptxReference(referenceContext(), reference)
    if (result.status !== "exact" && result.status !== "relocated") return
    const locator = (result.reference as PptxOfficeReference).locator
    const slideIndex = locator.type === "manual-region" ? locator.value.slideIndex : locator.value.slide.index
    const slide = elementRef.value?.querySelector<HTMLElement>(`[data-slide-index="${slideIndex}"]`)
    slide?.scrollIntoView({ behavior: "smooth", block: "center" })
  },
  async captureReferencePreview(reference: OfficeObjectReference): Promise<Blob> {
    emitReferenceError("capture", "CAPTURE_UNSUPPORTED", "PPTX preview capture is not provided by the Stage; use the host's capture pipeline.", reference.referenceId)
    throw new Error("PPTX reference preview capture is unsupported")
  },
})
</script>

<style scoped>
.pptx-stage {
  background: var(--pptx-surface-bg, transparent);
  position: relative;
}

.pptx-stage__content {
  min-height: inherit;
  min-width: inherit;
}

.pptx-stage__region-layer {
  cursor: crosshair;
  inset: 0;
  position: absolute;
  touch-action: none;
  z-index: 30;
}

.pptx-stage__region-frame {
  background: rgb(37 99 235 / 0.08);
  border: 2px solid #2563eb;
  box-sizing: border-box;
  pointer-events: none;
  position: absolute;
}
</style>
