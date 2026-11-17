<template>
  <div
    ref="surfaceRef"
    class="docx-document-surface"
    :data-selection-mode="selectionMode"
    @pointermove="onReferencePointerMove"
    @pointerleave="onReferencePointerLeave"
    @click="onReferenceClick"
    @keydown.esc="onReferenceEscape"
  >
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
      @page-count-change="onPageCountChange"
      @visible-page-range="onVisiblePageRange"
      @context-menu="emit('contextMenu', $event)"
      @selection-change="onRawSelectionChange"
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
    <div
      v-if="selectionMode === 'region'"
      class="docx-document-surface__region-layer"
      aria-label="选择文档区域"
      tabindex="0"
      @pointerdown="onRegionPointerDown"
      @pointermove="onRegionPointerMove"
      @pointerup="onRegionPointerUp"
      @pointercancel="onRegionPointerCancel"
      @keydown.esc.stop="onRegionKeyboardCancel"
    >
      <div v-if="regionFrame" class="docx-document-surface__region-frame" :style="regionFrame" />
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
  createDocxCommentReferenceDraft,
  createDocxImageReferenceDraft,
  createDocxPageReferenceDraft,
  createDocxRegionReferenceDraft,
  createDocxTableReferenceDraft,
  createDocxTextReferenceDraft,
  docNodeContentSignature,
  resolveDocxReference,
  type DocxOfficeReference,
  type DocxReferenceContext,
} from "@arcships/docx-core"
import {
  confirmOfficeReferenceDraft,
  createOfficeReferenceId,
  MAX_OFFICE_VISIBLE_CANDIDATES,
  officeReferenceErrorForResult,
  rankOfficeHitCandidates,
  toOfficeReferenceCandidatePreview,
  type NormalizedRect,
  type OfficeDocumentRevision,
  type OfficeHitCandidate,
  type OfficeObjectReference,
  type OfficeReferenceCandidatePreview,
  type OfficeReferenceError,
  type OfficeReferencePreviewOptions,
  type OfficeRegionDraftChange,
  type OfficeSelectionMode,
  type ResolveReferenceResult,
} from "@arcships/office-interaction"
import {
  findDocxSearchMatches,
  type DocxSearchMatch,
  type DocxSearchState,
} from "../composables/useDocxSearch"
import DocxViewerRoot from "./DocxViewerRoot.vue"
import { docxTextRangeFromDomSelection } from "../composables/editor-selection"

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
    /** `null` follows the controller flag; `false` always hides. */
    showTrackedChanges?: boolean | null
    showComments?: boolean | null
    documentId?: string
    selectionMode?: OfficeSelectionMode
    emitReferenceCandidates?: boolean
  }>(),
  {
    editable: false,
    zoomScale: 100,
    enableGestureZoom: true,
    selectionMode: "content",
    emitReferenceCandidates: false,
    showTrackedChanges: null,
    showComments: null,
  }
)

const emit = defineEmits<{
  pageCountChange: [count: number]
  visiblePageRange: [range: { startPageIndex: number; endPageIndex: number }]
  contextMenu: [ctx: { pageIndex: number; clientX: number; clientY: number; containerX: number; containerY: number; target?: EventTarget | null }]
  selectionChange: [sel: { kind: string; text?: string; nodeIndex?: number }]
  objectClick: [obj: { kind: "image" | "table"; nodeIndex?: number }]
  searchStateChange: [state: DocxSearchState]
  "update:zoom": [zoom: number]
  documentRevisionChange: [revision: OfficeDocumentRevision]
  referenceCandidateChange: [change: { candidates: readonly OfficeReferenceCandidatePreview[]; activeCandidateId?: string }]
  referenceConfirm: [event: ReturnType<typeof confirmOfficeReferenceDraft>]
  regionDraftChange: [event: OfficeRegionDraftChange]
  selectionCancel: [event: { mode: OfficeSelectionMode; reason: "escape" | "pointer-cancel" | "programmatic" }]
  referenceResolve: [event: { referenceId: string; result: ResolveReferenceResult }]
  referenceError: [error: OfficeReferenceError]
}>()

const surfaceRef = ref<HTMLElement>()
const viewerRootRef = ref<InstanceType<typeof DocxViewerRoot>>()
const effectiveZoomScale = computed(() =>
  props.zoom === undefined ? props.zoomScale : clampSurfaceZoom(props.zoom) * 100
)
const selectionMode = computed(() => props.selectionMode)
const pageCount = ref(0)
const revisionCounter = ref(0)
const generatedDocumentId = `docx-surface-${typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2)}`
const documentRevision = computed<OfficeDocumentRevision & { format: "docx" }>(() => {
  const hostDocumentId = props.documentId?.trim()
  if (hostDocumentId) {
    // A host-supplied identity must survive remounts. The fingerprint keeps
    // page/region references honest while never embedding this instance's id.
    return {
      format: "docx",
      documentId: hostDocumentId,
      revision: `${hostDocumentId}:${docNodeContentSignature(props.model)}:${pageCount.value}`,
    }
  }
  return {
    format: "docx",
    documentId: generatedDocumentId,
    revision: `${generatedDocumentId}:${revisionCounter.value}`,
  }
})

function referenceContext(): DocxReferenceContext {
  return { revision: documentRevision.value, model: props.model, pageCount: pageCount.value || undefined }
}

function emitReferenceError(operation: OfficeReferenceError["operation"], code: OfficeReferenceError["code"], message: string, referenceId?: string): void {
  emit("referenceError", { code, operation, format: "docx", recoverable: true, ...(referenceId ? { referenceId } : {}), message })
}

function emitResolveFailure(
  reference: OfficeObjectReference,
  result: ResolveReferenceResult,
  operation: "resolve" | "describe" | "scroll",
): void {
  const failure = officeReferenceErrorForResult("docx", result, {
    operation,
    referenceId: reference.referenceId,
  })
  if (failure) emit("referenceError", failure)
}

function confirmDraft(
  draft: Parameters<typeof confirmOfficeReferenceDraft>[0],
  trigger: "pointer" | "keyboard" | "touch" | "programmatic",
  snapshot?: Parameters<typeof confirmOfficeReferenceDraft>[1]["snapshot"],
  additiveRequested = false,
): void {
  try {
    emit("referenceConfirm", confirmOfficeReferenceDraft(draft, {
      referenceId: createOfficeReferenceId(),
      trigger,
      additiveRequested,
      ...(snapshot ? { snapshot } : {}),
    }))
  } catch (reason) {
    emitReferenceError("describe", "INVALID_REFERENCE", reason instanceof Error ? reason.message : "Unable to confirm DOCX reference.")
  }
}

function onPageCountChange(count: number): void {
  const changed = count !== pageCount.value
  pageCount.value = count
  if (changed && props.documentId?.trim()) emit("documentRevisionChange", documentRevision.value)
  emit("pageCountChange", count)
}

function onRawSelectionChange(selection: { kind: string; text?: string; pageIndex?: number }): void {
  emit("selectionChange", selection)
  if (selectionMode.value !== "content" || selection.kind !== "text" || !selection.text) return
  const root = surfaceRef.value
  const nativeSelection = root?.ownerDocument?.getSelection?.()
  if (!root || !nativeSelection) return
  const range = docxTextRangeFromDomSelection(nativeSelection, root)
  if (!range) {
    emitReferenceError("hit-test", "HIT_TEST_FAILED", "The DOCX text selection could not be mapped to the document model.")
    return
  }
  try {
    const draft = createDocxTextReferenceDraft(referenceContext(), range)
    confirmDraft(draft, "pointer", {
      label: selection.text.length > 80 ? `${selection.text.slice(0, 77)}…` : selection.text,
      path: [
        { kind: "document", label: "Document" },
        ...(selection.pageIndex === undefined ? [] : [{ kind: "page" as const, label: `Page ${selection.pageIndex + 1}` }]),
        { kind: "text-range", label: "Selected text" },
      ],
      content: { text: selection.text },
    })
  } catch (reason) {
    emitReferenceError("describe", "INVALID_REFERENCE", reason instanceof Error ? reason.message : "Unable to create DOCX text reference.")
  }
}

function dataAttribute(element: Element | null | undefined, name: string): string | undefined {
  if (!element) return undefined
  const dataset = (element as HTMLElement).dataset
  if (dataset) {
    const key = name
      .replace(/^data-/u, "")
      .replace(/-([a-z])/gu, (_match, letter: string) => letter.toUpperCase())
    const value = dataset[key]
    if (value !== undefined) return String(value)
  }
  return typeof element.getAttribute === "function" ? element.getAttribute(name) ?? undefined : undefined
}

function elementsAtPoint(clientX: number, clientY: number): readonly Element[] {
  const document = surfaceRef.value?.ownerDocument
  const all = document?.elementsFromPoint?.(clientX, clientY)
  if (all?.length) return all
  const single = document?.elementFromPoint?.(clientX, clientY)
  return single ? [single] : []
}

function pageAtPoint(clientX: number, clientY: number): HTMLElement | undefined {
  for (const element of elementsAtPoint(clientX, clientY)) {
    const page = element.closest<HTMLElement>("[data-docx-page-index]")
    if (page) return page
  }
  return undefined
}

function paragraphLocationForElement(element: Element): DocxTextRangeLocation | undefined {
  const cell = element.closest<HTMLElement>("[data-docx-table-cell-paragraph-host='true']")
  if (cell) {
    const tableIndex = Number(dataAttribute(cell, "data-docx-table-index"))
    const rowIndex = Number(dataAttribute(cell, "data-docx-table-row-index"))
    const cellIndex = Number(dataAttribute(cell, "data-docx-table-cell-index"))
    const paragraphIndex = Number(dataAttribute(cell, "data-docx-paragraph-index"))
    if ([tableIndex, rowIndex, cellIndex, paragraphIndex].every((value) => Number.isSafeInteger(value) && value >= 0)) {
      return { kind: "table-cell", tableIndex, rowIndex, cellIndex, paragraphIndex }
    }
  }
  const paragraph = element.closest<HTMLElement>("[data-docx-paragraph-node-index]")
  if (paragraph) {
    const nodeIndex = Number(dataAttribute(paragraph, "data-docx-paragraph-node-index"))
    if (Number.isSafeInteger(nodeIndex) && nodeIndex >= 0) return { kind: "paragraph", nodeIndex }
  }
  const floatingImage = element.closest<HTMLElement>("[data-image-node-index]")
  if (floatingImage) {
    const nodeIndex = Number(dataAttribute(floatingImage, "data-image-node-index"))
    if (Number.isSafeInteger(nodeIndex) && nodeIndex >= 0) return { kind: "paragraph", nodeIndex }
  }
  return undefined
}

function pageHitCandidate(pageIndex: number): OfficeHitCandidate {
  return {
    candidateId: `docx:page:${pageIndex}`,
    draft: createDocxPageReferenceDraft(referenceContext(), pageIndex),
    preview: {
      label: `Page ${pageIndex + 1}`,
      path: [{ kind: "document", label: "Document" }, { kind: "page", label: `Page ${pageIndex + 1}` }],
    },
    hit: "inside",
    depth: 1,
  }
}

function tableHitCandidate(element: Element): OfficeHitCandidate | undefined {
  const table = element.closest<HTMLElement>("[data-docx-table-host='true']")
    ?? element.closest<HTMLElement>("[data-docx-table-cell='true']")
  if (!table) return undefined
  const tableIndex = Number(dataAttribute(table, "data-docx-table-index"))
  if (!Number.isSafeInteger(tableIndex) || tableIndex < 0) return undefined
  return {
    candidateId: `docx:table:${tableIndex}`,
    draft: createDocxTableReferenceDraft(referenceContext(), tableIndex),
    preview: {
      label: `Table ${tableIndex + 1}`,
      path: [{ kind: "document", label: "Document" }, { kind: "table", label: `Table ${tableIndex + 1}` }],
    },
    hit: "direct",
    depth: 2,
  }
}

function imageHitCandidate(element: Element): OfficeHitCandidate | undefined {
  const image = element.closest<HTMLElement>("[data-docx-image-child-index]")
    ?? element.closest<HTMLElement>("[data-image-child-index]")
  if (!image) return undefined
  const childIndex = Number(dataAttribute(image, "data-docx-image-child-index") ?? dataAttribute(image, "data-image-child-index"))
  const paragraph = paragraphLocationForElement(image)
  if (!paragraph || !Number.isSafeInteger(childIndex) || childIndex < 0) return undefined
  const key = paragraph.kind === "paragraph"
    ? `p${paragraph.nodeIndex}`
    : `t${paragraph.tableIndex}.${paragraph.rowIndex}.${paragraph.cellIndex}.${paragraph.paragraphIndex}`
  return {
    candidateId: `docx:image:${key}:${childIndex}`,
    draft: createDocxImageReferenceDraft(referenceContext(), { paragraph, childIndex }),
    preview: {
      label: "Image",
      path: [
        { kind: "document", label: "Document" },
        { kind: "image", label: "Image" },
      ],
    },
    hit: "direct",
    depth: 3,
  }
}

function commentHitCandidates(element: Element): readonly OfficeHitCandidate[] {
  const marked = element.closest<HTMLElement>("[data-docx-comment-ids], [data-docx-comment-id]")
  if (!marked) return []
  const ids = new Set(
    (dataAttribute(marked, "data-docx-comment-ids") ?? dataAttribute(marked, "data-docx-comment-id") ?? "")
      .split(/\s+/u)
      .filter(Boolean),
  )
  if (ids.size === 0) return []
  const anchor = paragraphLocationForElement(marked)
  const comments = props.model.metadata.comments ?? []
  const candidates: OfficeHitCandidate[] = []
  for (const commentId of ids) {
    const comment = comments.find((entry) => String(entry.id) === commentId)
    const label = comment?.text
      ? `Comment: ${comment.text.replace(/\s+/gu, " ").trim().slice(0, 80)}`
      : `Comment ${commentId}`
    candidates.push({
      candidateId: `docx:comment:${commentId}`,
      draft: createDocxCommentReferenceDraft(referenceContext(), { commentId, ...(anchor ? { anchor } : {}) }),
      preview: {
        label,
        path: [{ kind: "document", label: "Document" }, { kind: "comment", label }],
      },
      hit: "direct",
      depth: 3,
    })
  }
  return candidates
}

function hitTest(point: { clientX: number; clientY: number }): readonly OfficeReferenceCandidatePreview[] {
  try {
    const elements = elementsAtPoint(point.clientX, point.clientY)
    const pageElement = elements
      .map((element) => element.closest<HTMLElement>("[data-docx-page-index]"))
      .find((element) => element !== null)
    const pageIndex = Number(dataAttribute(pageElement, "data-docx-page-index"))
    const candidates: OfficeHitCandidate[] = []
    const seen = new Set<string>()
    const push = (candidate: OfficeHitCandidate | undefined): void => {
      if (!candidate || seen.has(candidate.candidateId) || candidates.length >= MAX_OFFICE_VISIBLE_CANDIDATES) return
      seen.add(candidate.candidateId)
      candidates.push(candidate)
    }

    for (const element of elements) {
      try {
        push(tableHitCandidate(element))
        push(imageHitCandidate(element))
        for (const comment of commentHitCandidates(element)) push(comment)
      } catch {
        // Stale DOM nodes (for example a virtualized page being replaced) must
        // not discard candidates that still resolve against the model.
      }
    }
    if (Number.isSafeInteger(pageIndex) && pageIndex >= 0) {
      try {
        push(pageHitCandidate(pageIndex))
      } catch {
        // Page candidates are unavailable until pagination produces a page.
      }
    }
    return rankOfficeHitCandidates(candidates).map(toOfficeReferenceCandidatePreview)
  } catch (reason) {
    emitReferenceError("hit-test", "HIT_TEST_FAILED", reason instanceof Error ? reason.message : "DOCX hit test failed.")
    return []
  }
}

let candidateFrame = 0
let pendingCandidatePoint: { clientX: number; clientY: number } | undefined

function flushCandidatePointer(): void {
  const point = pendingCandidatePoint
  pendingCandidatePoint = undefined
  if (!point) return
  const candidates = hitTest(point)
  emit("referenceCandidateChange", { candidates, ...(candidates[0] ? { activeCandidateId: candidates[0].candidateId } : {}) })
}

function cancelCandidateFrame(): void {
  pendingCandidatePoint = undefined
  if (!candidateFrame) return
  if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(candidateFrame)
  candidateFrame = 0
}

function onReferencePointerMove(event: PointerEvent): void {
  if (selectionMode.value !== "object" || !props.emitReferenceCandidates) return
  pendingCandidatePoint = { clientX: event.clientX, clientY: event.clientY }
  if (typeof requestAnimationFrame !== "function") {
    flushCandidatePointer()
    return
  }
  if (candidateFrame) return
  candidateFrame = requestAnimationFrame(() => {
    candidateFrame = 0
    flushCandidatePointer()
  })
}

function onReferencePointerLeave(): void {
  cancelCandidateFrame()
  if (selectionMode.value !== "object" || !props.emitReferenceCandidates) return
  emit("referenceCandidateChange", { candidates: [] })
}

function onReferenceClick(event: MouseEvent): void {
  if (selectionMode.value !== "object") return
  const candidate = hitTest(event)[0]
  if (candidate) confirmDraft(candidate.draft, event.detail === 0 ? "keyboard" : "pointer", candidate.preview, event.shiftKey)
}

function onReferenceEscape(): void {
  if (selectionMode.value === "content") return
  emit("selectionCancel", { mode: selectionMode.value, reason: "escape" })
}

type RegionGesture = {
  pageIndex: number
  page: HTMLElement
  startX: number
  startY: number
  lastClientX: number
  lastClientY: number
  pointerId: number
}
const regionGesture = ref<RegionGesture | null>(null)
const regionRect = ref<NormalizedRect | null>(null)
const regionBounds = shallowRef<DOMRect | null>(null)
const regionFrame = computed(() => {
  const rect = regionRect.value
  const bounds = regionBounds.value
  const root = rect && bounds ? surfaceRef.value?.getBoundingClientRect() : undefined
  return rect && bounds && root ? {
    left: `${bounds.left - root.left + rect.x * bounds.width}px`,
    top: `${bounds.top - root.top + rect.y * bounds.height}px`,
    width: `${rect.width * bounds.width}px`,
    height: `${rect.height * bounds.height}px`
  } : undefined
})

function clampUnit(value: number): number { return Math.max(0, Math.min(1, value)) }

function normalizedRegionRect(
  gesture: RegionGesture,
  bounds: DOMRect,
  clientX: number,
  clientY: number,
): NormalizedRect {
  const x1 = clampUnit((Math.min(gesture.startX, clientX) - bounds.left) / bounds.width)
  const y1 = clampUnit((Math.min(gesture.startY, clientY) - bounds.top) / bounds.height)
  const x2 = clampUnit((Math.max(gesture.startX, clientX) - bounds.left) / bounds.width)
  const y2 = clampUnit((Math.max(gesture.startY, clientY) - bounds.top) / bounds.height)
  return { x: x1, y: y1, width: x2 - x1, height: y2 - y1 }
}

function liveRegionBounds(gesture: RegionGesture): DOMRect | undefined {
  const bounds = gesture.page.getBoundingClientRect()
  return bounds && bounds.width > 0 && bounds.height > 0 ? bounds : undefined
}

/**
 * Re-read the page rect on every gesture sample (and on scroll/zoom while a
 * gesture is active) so edge auto-scroll and Ctrl-wheel zoom cannot leave the
 * denominators or the preview frame stale.
 */
function updateRegionGeometry(clientX: number, clientY: number, emitChange: boolean): void {
  const gesture = regionGesture.value
  if (!gesture) return
  const bounds = liveRegionBounds(gesture)
  if (!bounds) return
  const rect = normalizedRegionRect(gesture, bounds, clientX, clientY)
  regionBounds.value = bounds
  regionRect.value = rect
  if (emitChange && rect.width > 0 && rect.height > 0) {
    emit("regionDraftChange", {
      phase: "change",
      region: { space: "page", pageIndex: gesture.pageIndex, rect },
    })
  }
}

function refreshRegionGeometry(): void {
  const gesture = regionGesture.value
  if (!gesture) return
  updateRegionGeometry(gesture.lastClientX, gesture.lastClientY, true)
}

function onRegionPointerDown(event: PointerEvent): void {
  if (event.button !== 0) return
  const page = pageAtPoint(event.clientX, event.clientY)
  const pageIndex = Number(dataAttribute(page, "data-docx-page-index"))
  const bounds = page?.getBoundingClientRect()
  if (!page || !bounds || bounds.width <= 0 || bounds.height <= 0 || !Number.isSafeInteger(pageIndex)) return
  const x = clampUnit((event.clientX - bounds.left) / bounds.width)
  const y = clampUnit((event.clientY - bounds.top) / bounds.height)
  regionGesture.value = {
    pageIndex,
    page,
    startX: event.clientX,
    startY: event.clientY,
    lastClientX: event.clientX,
    lastClientY: event.clientY,
    pointerId: event.pointerId,
  }
  regionBounds.value = bounds
  regionRect.value = { x, y, width: 0, height: 0 }
  ;(event.currentTarget as HTMLElement).setPointerCapture?.(event.pointerId)
  emit("regionDraftChange", { phase: "start", point: { space: "page", pageIndex, x, y } })
  event.preventDefault()
}

function onRegionPointerMove(event: PointerEvent): void {
  const gesture = regionGesture.value
  if (!gesture || gesture.pointerId !== event.pointerId) return
  gesture.lastClientX = event.clientX
  gesture.lastClientY = event.clientY
  updateRegionGeometry(event.clientX, event.clientY, true)
}

function finishRegion(event: PointerEvent, cancelled: boolean): void {
  const gesture = regionGesture.value
  if (!gesture || gesture.pointerId !== event.pointerId) return
  const target = event.currentTarget as HTMLElement
  if (target.hasPointerCapture?.(event.pointerId)) target.releasePointerCapture?.(event.pointerId)
  gesture.lastClientX = event.clientX
  gesture.lastClientY = event.clientY
  updateRegionGeometry(event.clientX, event.clientY, false)
  const rect = regionRect.value
  regionGesture.value = null
  regionRect.value = null
  regionBounds.value = null
  if (cancelled || !rect || rect.width < 0.005 || rect.height < 0.005) {
    emit("selectionCancel", { mode: "region", reason: cancelled ? "pointer-cancel" : "programmatic" })
    return
  }
  confirmDraft(createDocxRegionReferenceDraft(referenceContext(), { space: "page", pageIndex: gesture.pageIndex, rect }), event.pointerType === "touch" ? "touch" : "pointer")
}

function onRegionPointerUp(event: PointerEvent): void { finishRegion(event, false) }
function onRegionPointerCancel(event: PointerEvent): void { finishRegion(event, true) }
function onRegionKeyboardCancel(): void {
  regionGesture.value = null
  regionRect.value = null
  regionBounds.value = null
  emit("selectionCancel", { mode: "region", reason: "escape" })
}

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

interface HighlightRect {
  left: number
  top: number
  right: number
  bottom: number
  width: number
  height: number
}

interface CachedMatchGeometry {
  hostRect: HighlightRect
  rects: HighlightRect[]
}

interface SearchGeometryCache {
  matches: readonly DocxSearchMatch[]
  container: HTMLElement
  scrollTop: number
  scrollLeft: number
  rootRect: HighlightRect
  viewportRect: HighlightRect
  entries: CachedMatchGeometry[][]
}

const searchHighlights = shallowRef<SearchHighlight[]>([])
let highlightFrame = 0
let pendingReuseGeometry = false
let searchGeometryCache: SearchGeometryCache | undefined
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

const SEARCH_ACTIVATION_RETRY_LIMIT = 30
const visiblePageWaiters = new Set<() => void>()

function waitForActivationRetry(signal: AbortSignal): Promise<void> {
  return new Promise((resolve) => {
    let settled = false
    const finish = (): void => {
      if (settled) return
      settled = true
      visiblePageWaiters.delete(finish)
      signal.removeEventListener?.("abort", finish)
      resolve()
    }
    if (signal.aborted) {
      finish()
      return
    }
    visiblePageWaiters.add(finish)
    signal.addEventListener?.("abort", finish, { once: true })
    void nextFrame().then(finish)
  })
}

function preferredSearchCandidate(
  candidates: readonly { host: HTMLElement; range: Range }[],
  container: HTMLElement | undefined,
): { host: HTMLElement; range: Range } | undefined {
  if (!container) return candidates[0]
  const viewport = container.getBoundingClientRect()
  const visible = candidates.find(({ range }) => {
    const rect = range.getClientRects()[0]
    return rect && rect.top >= viewport.top && rect.bottom <= viewport.bottom
  })
  if (visible) return visible
  const center = viewport.top + viewport.height / 2
  return [...candidates].sort((left, right) => {
    const leftRect = left.range.getClientRects()[0]
    const rightRect = right.range.getClientRects()[0]
    return Math.abs((leftRect?.top ?? Infinity) - center) - Math.abs((rightRect?.top ?? Infinity) - center)
  })[0]
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
    "[data-docx-numbering-label='true'], [data-docx-note-marker='true'], [data-docx-image-child-index], [data-docx-form-field='true'], [data-docx-tracked-change='deletion']",
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
  rect: HighlightRect,
  hostRect: HighlightRect,
  viewportRect: HighlightRect,
): { left: number; top: number; right: number; bottom: number } | undefined {
  const left = Math.max(rect.left, hostRect.left, viewportRect.left)
  const top = Math.max(rect.top, hostRect.top, viewportRect.top)
  const right = Math.min(rect.right, hostRect.right, viewportRect.right)
  const bottom = Math.min(rect.bottom, hostRect.bottom, viewportRect.bottom)
  if (right - left < 0.5 || bottom - top < 0.5) return undefined
  return { left, top, right, bottom }
}

function highlightRect(rect: DOMRect): HighlightRect {
  const width = Number.isFinite(rect.width) ? rect.width : rect.right - rect.left
  const height = Number.isFinite(rect.height) ? rect.height : rect.bottom - rect.top
  return { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom, width, height }
}

function cachedSearchGeometry(
  root: HTMLElement,
  container: HTMLElement,
  matches: readonly DocxSearchMatch[],
): SearchGeometryCache {
  return {
    matches,
    container,
    scrollTop: container.scrollTop,
    scrollLeft: container.scrollLeft,
    rootRect: highlightRect(root.getBoundingClientRect()),
    viewportRect: highlightRect(container.getBoundingClientRect()),
    entries: matches.map((match) =>
      rangesForMatch(match).map(({ host, range }) => ({
        hostRect: highlightRect(host.getBoundingClientRect()),
        rects: Array.from(range.getClientRects()).map(highlightRect),
      }))
    ),
  }
}

/**
 * Highlight geometry is expensive to collect (selector walk + text segment
 * collection + getClientRects per range). Cache it and translate the cached
 * rectangles by the scroll delta while scrolling; only rebuild when the
 * document, zoom, search state or mounted page window changes.
 */
function renderSearchHighlights(reuseGeometry = false): void {
  const root = surfaceRef.value
  const container = viewerRootRef.value?.scrollContainer
  const state = surfaceSearch.getSearchState()
  if (!root || !container || !state.matches.length) {
    searchGeometryCache = undefined
    searchHighlights.value = []
    return
  }

  let cache = reuseGeometry
    && searchGeometryCache?.matches === state.matches
    && searchGeometryCache.container === container
    ? searchGeometryCache
    : undefined
  if (!cache) {
    cache = cachedSearchGeometry(root, container, state.matches)
    searchGeometryCache = cache
  }

  const deltaX = container.scrollLeft - cache.scrollLeft
  const deltaY = container.scrollTop - cache.scrollTop
  const { rootRect, viewportRect } = cache
  const seen = new Set<string>()
  const next: SearchHighlight[] = []
  cache.entries.forEach((hosts, matchIndex) => {
    for (const { hostRect, rects } of hosts) {
      const translatedHost: HighlightRect = {
        ...hostRect,
        left: hostRect.left - deltaX,
        right: hostRect.right - deltaX,
        top: hostRect.top - deltaY,
        bottom: hostRect.bottom - deltaY,
      }
      for (const cachedRect of rects) {
        const rect: HighlightRect = {
          ...cachedRect,
          left: cachedRect.left - deltaX,
          right: cachedRect.right - deltaX,
          top: cachedRect.top - deltaY,
          bottom: cachedRect.bottom - deltaY,
        }
        const clipped = intersectionRect(rect, translatedHost, viewportRect)
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

function scheduleSearchHighlights(reuseGeometry = false): void {
  if (typeof requestAnimationFrame !== "function") {
    renderSearchHighlights(reuseGeometry)
    return
  }
  if (highlightFrame) {
    cancelAnimationFrame(highlightFrame)
    // A forced rebuild scheduled in the same frame must win over a cheap
    // scroll-only refresh.
    pendingReuseGeometry = pendingReuseGeometry && reuseGeometry
  } else {
    pendingReuseGeometry = reuseGeometry
  }
  highlightFrame = requestAnimationFrame(() => {
    highlightFrame = 0
    const reuse = pendingReuseGeometry
    pendingReuseGeometry = false
    renderSearchHighlights(reuse)
  })
}

async function activateSearchMatchInView(match: DocxSearchMatch, signal: AbortSignal): Promise<void> {
  for (let attempt = 0; attempt <= SEARCH_ACTIVATION_RETRY_LIMIT; attempt += 1) {
    throwIfAborted(signal)
    const container = viewerRootRef.value?.scrollContainer
    const candidate = preferredSearchCandidate(rangesForMatch(match), container)
    if (candidate) {
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
      return
    }
    if (attempt === SEARCH_ACTIVATION_RETRY_LIMIT) break
    // The match lives on a virtualized page that is not mounted yet. Ask the
    // viewer to scroll it into range and retry on the next frame or as soon as
    // the visible page range changes.
    viewerRootRef.value?.scrollToNode(match.nodeIndex, { behavior: "auto" })
    await nextTick()
    await waitForActivationRetry(signal)
  }
  throw new Error(
    `DOCX search match could not be projected into the mounted document after ${SEARCH_ACTIVATION_RETRY_LIMIT + 1} attempts.`,
  )
}

const surfaceSearch = createSurfaceSearchSession<DocxSearchMatch>({
  search(query, { signal }) {
    return findDocxSearchMatches(props.model, query, signal)
  },
  activate(match, _index, { signal }) {
    return activateSearchMatchInView(match, signal)
  },
  clear() {
    searchGeometryCache = undefined
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
  refreshRegionGeometry()
  for (const waiter of [...visiblePageWaiters]) waiter()
}

function attachSearchScrollListener(): void {
  cleanupSearchScrollListener?.()
  const container = viewerRootRef.value?.scrollContainer
  if (!container) return
  const onScroll = () => {
    scheduleSearchHighlights(true)
    refreshRegionGeometry()
  }
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
  revisionCounter.value += 1
  emit("documentRevisionChange", documentRevision.value)
  gestureToken += 1
  pendingAnchor = undefined
  pendingZoom = undefined
  surfaceSearch.clearSearch()
}, { immediate: true })

watch(() => props.documentId, () => emit("documentRevisionChange", documentRevision.value))

watch(
  () => [effectiveZoomScale.value, props.fitWidth, props.showTrackedChanges, props.showComments],
  () => {
    void nextTick(() => {
      scheduleSearchHighlights()
      refreshRegionGeometry()
    })
  },
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
  cancelCandidateFrame()
  visiblePageWaiters.clear()
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
  getDocumentRevision: (): OfficeDocumentRevision => documentRevision.value,
  hitTest,
  async describeReference(reference: OfficeObjectReference, signal?: AbortSignal) {
    if (signal?.aborted) {
      emitReferenceError("describe", "ABORTED", "DOCX reference description aborted.", reference.referenceId)
      throw new DOMException("DOCX reference description aborted.", "AbortError")
    }
    const result = resolveDocxReference(referenceContext(), reference)
    if (result.status !== "exact" && result.status !== "relocated") {
      emitResolveFailure(reference, result, "describe")
      const reason = result.status === "ambiguous" ? result.reasonCodes.join(", ") : result.reasonCode
      throw new Error(`DOCX reference is not resolvable in the current revision (${reason}).`)
    }
    return result.descriptor
  },
  async resolveReference(reference: OfficeObjectReference): Promise<ResolveReferenceResult> {
    const result = resolveDocxReference(referenceContext(), reference)
    emitResolveFailure(reference, result, "resolve")
    emit("referenceResolve", { referenceId: reference.referenceId, result })
    return result
  },
  async scrollToReference(reference: OfficeObjectReference): Promise<void> {
    const result = resolveDocxReference(referenceContext(), reference)
    if (result.status !== "exact" && result.status !== "relocated") {
      emitResolveFailure(reference, result, "scroll")
      return
    }
    const locator = (result.reference as DocxOfficeReference).locator
    if (locator.type === "manual-region") viewerRootRef.value?.scrollToPage(locator.value.pageIndex)
    else if (locator.value.kind === "page") viewerRootRef.value?.scrollToPage(locator.value.pageIndex)
    else if (locator.value.kind === "text-range") {
      const first = locator.value.start.path.find((segment) => segment.kind === "paragraph" || segment.kind === "node")
      if (first) viewerRootRef.value?.scrollToNode(first.index, { behavior: "smooth" })
    } else if (locator.value.kind === "image") {
      const first = locator.value.paragraphPath.find((segment) => segment.kind === "paragraph" || segment.kind === "node")
      if (first) viewerRootRef.value?.scrollToNode(first.index, { behavior: "smooth" })
    } else if (locator.value.kind === "structure") {
      const first = locator.value.path.find((segment) => segment.kind === "table" || segment.kind === "node")
      if (first) viewerRootRef.value?.scrollToNode(first.index, { behavior: "smooth" })
    } else if (locator.value.kind === "comment" && locator.value.anchorPath) {
      const first = locator.value.anchorPath.find((segment) => segment.kind === "paragraph" || segment.kind === "node")
      if (first) viewerRootRef.value?.scrollToNode(first.index, { behavior: "smooth" })
    }
  },
  async captureReferencePreview(reference: OfficeObjectReference, options?: OfficeReferencePreviewOptions): Promise<Blob> {
    const maxDimension = 16_384
    const requested = [options?.maxWidth, options?.maxHeight]
    if (requested.some((value) => value !== undefined && (!Number.isFinite(value) || value <= 0 || value > maxDimension))) {
      emitReferenceError(
        "capture",
        "CAPTURE_LIMIT_EXCEEDED",
        `DOCX preview capture limits must be between 1 and ${maxDimension} pixels.`,
        reference.referenceId,
      )
      throw new RangeError("DOCX reference preview capture limit exceeded")
    }
    emitReferenceError("capture", "CAPTURE_UNSUPPORTED", "DOCX preview capture is not provided by the Surface; use the host's capture pipeline.", reference.referenceId)
    throw new Error("DOCX reference preview capture is unsupported")
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

.docx-document-surface__region-layer {
  cursor: crosshair;
  inset: 0;
  position: absolute;
  touch-action: none;
  z-index: 30;
}

.docx-document-surface__region-frame {
  background: rgb(37 99 235 / 0.08);
  border: 2px solid #2563eb;
  box-sizing: border-box;
  pointer-events: none;
  position: absolute;
}
</style>
