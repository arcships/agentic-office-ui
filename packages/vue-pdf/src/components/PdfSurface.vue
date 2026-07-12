<template>
  <div ref="scrollRef" class="pdf-surface" data-testid="pdf-surface" @scroll.passive="onScroll" @contextmenu.prevent="onContextMenu">
    <div v-if="!effectiveSource" class="pdf-surface__empty">No PDF loaded</div>
    <div v-else-if="loading" class="pdf-surface__empty">Opening PDF…</div>
    <div v-else-if="loadError" class="pdf-surface__error" :data-error-code="loadError.code">
      {{ loadError.code }}: {{ loadError.message }}
    </div>
    <div v-else-if="renderDocument" class="pdf-surface__pages">
      <div
        v-for="pageIndex in Array.from({ length: renderDocument.pageCount }, (_, i) => i)"
        :key="`page-${pageIndex}`"
        class="pdf-surface__page-slot"
        :data-page-index="pageIndex"
      >
        <div v-if="pageSlots[pageIndex]?.state === 'loading'" class="pdf-surface__page-loading">
          Rendering page {{ pageIndex + 1 }}…
        </div>
        <div v-else-if="pageSlots[pageIndex]?.state === 'error'" class="pdf-surface__page-error">
          Failed to render page {{ pageIndex + 1 }}
        </div>
        <img
          v-else-if="pageSlots[pageIndex]?.url"
          :src="pageSlots[pageIndex].url"
          :alt="`Page ${pageIndex + 1}`"
          :style="pageImgStyle(pageIndex)"
          class="pdf-surface__page-img"
          draggable="false"
        />
        <div v-else class="pdf-surface__page-placeholder" />
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, reactive, ref, watch } from "vue"
import {
  createLatestTaskCoordinator,
  createOfficeTaskSequence,
  type OfficeSource,
} from "@arcships/office-runtime"
import {
  DEFAULT_PDF_MAX_FILE_SIZE,
  loadVerifiedPdfSource,
  toPdfLoadError,
  type PdfDiagnostic,
  type PdfLoadError,
  type PdfSource,
  type PdfUrlPolicy,
} from "../pdf-url-policy"
import {
  createPdfRenderRuntime,
  type PdfRenderDocument,
  type PdfRenderRuntime,
  type PdfRotation,
} from "../pdf/pdf-render-runtime"

// ── Types ────────────────────────────────────────────────────────────
interface PageSlot {
  url: string | null
  state: "pending" | "loading" | "ready" | "error"
}

// ── Props ────────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    src?: string
    source?: PdfSource
    urlPolicy?: PdfUrlPolicy
    runtime?: PdfRenderRuntime
    pdfiumWasmUrl?: string
    maxFileSize?: number
    className?: string
    defaultZoom?: number
    /** When true, zoom auto-fits the widest page to container width. */
    fitWidth?: boolean
  }>(),
  {
    maxFileSize: DEFAULT_PDF_MAX_FILE_SIZE,
    defaultZoom: 1,
  },
)

const emit = defineEmits<{
  "document-load-success": [numPages: number]
  "document-load-error": [error: PdfLoadError]
  "visible-page-change": [pageIndex: number]
  contextMenu: [ctx: { pageIndex: number; clientX: number; clientY: number; containerX: number; containerY: number }]
  selectionChange: [sel: { kind: string; text?: string; pageIndex?: number }]
  diagnostic: [event: PdfDiagnostic]
}>()

// ── Runtime & loading ────────────────────────────────────────────────
const scrollRef = ref<HTMLElement | null>(null)
const zoom = ref(Math.min(2, Math.max(0.5, props.defaultZoom)))
const rotation = ref<PdfRotation>(0)
const loading = ref(false)
const loadError = ref<PdfLoadError | null>(null)
const renderDocument = ref<PdfRenderDocument | null>(null)
const verifiedBlob = ref<Blob | null>(null)

const effectiveSource = computed<PdfSource | undefined>(() => {
  if (props.source) return props.source
  if (props.src) return { kind: "url", url: props.src }
  return undefined
})

let ownedRuntime: PdfRenderRuntime | undefined
let ownedRuntimeWasmUrl: string | undefined
let runtimeForDocument: PdfRenderRuntime | undefined
let latestRequestId = 0
let renderGeneration = 0
let renderAbort: AbortController | undefined
let unmounted = false

const pdfRuntimeId = `pdf-surface-${
  typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`
}`
const pdfTaskSequence = createOfficeTaskSequence(pdfRuntimeId)
const pdfLoadCoordinator = createLatestTaskCoordinator(pdfTaskSequence)

// ── Page slots ───────────────────────────────────────────────────────
const pageSlots = reactive<Record<number, PageSlot>>({})

function initPageSlots(): void {
  for (const key of Object.keys(pageSlots)) delete pageSlots[Number(key)]
  if (!renderDocument.value) return
  for (let i = 0; i < renderDocument.value.pageCount; i++) {
    pageSlots[i] = { url: null, state: "pending" }
  }
}

// ── Runtime helpers ──────────────────────────────────────────────────
async function selectRuntime(): Promise<PdfRenderRuntime> {
  if (props.runtime) {
    if (ownedRuntime) {
      const rt = ownedRuntime
      ownedRuntime = undefined
      ownedRuntimeWasmUrl = undefined
      await rt.dispose()
    }
    return props.runtime
  }
  const wasmUrl = props.pdfiumWasmUrl?.trim() || undefined
  if (!ownedRuntime || ownedRuntimeWasmUrl !== wasmUrl) {
    if (ownedRuntime) await ownedRuntime.dispose()
    ownedRuntime = createPdfRenderRuntime({ wasmUrl })
    ownedRuntimeWasmUrl = wasmUrl
  }
  return ownedRuntime
}

function cancelRender(): void {
  renderGeneration++
  renderAbort?.abort()
  renderAbort = undefined
}

async function releaseDocument(): Promise<void> {
  cancelRender()
  for (const slot of Object.values(pageSlots)) {
    if (slot.url) URL.revokeObjectURL(slot.url)
  }
  for (const key of Object.keys(pageSlots)) delete pageSlots[Number(key)]
  const doc = renderDocument.value
  const rt = runtimeForDocument
  renderDocument.value = null
  runtimeForDocument = undefined
  if (doc && rt) await rt.closeDocument(doc)
}

function createLinkedAbort(parent?: AbortSignal) {
  const controller = new AbortController()
  const abort = () => controller.abort()
  if (parent?.aborted) controller.abort()
  else parent?.addEventListener("abort", abort, { once: true })
  return { controller, unlink: () => parent?.removeEventListener("abort", abort) }
}

function toOfficeSource(source: PdfSource): OfficeSource {
  if (source.kind === "url") return { kind: "url", url: source.url }
  return {
    kind: "file",
    name: (source.blob as File).name ?? "document.pdf",
    file: {
      name: (source.blob as File).name ?? "document.pdf",
      size: source.blob.size,
      type: source.blob.type,
      arrayBuffer: () => source.blob.arrayBuffer(),
    },
  }
}

// ── Render all pages sequentially ────────────────────────────────────
async function renderAllPages(requestId: number, parentSignal?: AbortSignal): Promise<void> {
  const doc = renderDocument.value
  const rt = runtimeForDocument
  if (!doc || !rt) return
  const gen = ++renderGeneration
  renderAbort?.abort()
  const { controller, unlink } = createLinkedAbort(parentSignal)
  renderAbort = controller

  for (let i = 0; i < doc.pageCount; i++) {
    if (controller.signal.aborted || gen !== renderGeneration || requestId !== latestRequestId) return
    const slot = pageSlots[i]
    if (!slot || slot.state === "ready") continue
    slot.state = "loading"
    try {
      const blob = await rt.renderPage(doc, i, {
        scale: zoom.value,
        rotation: rotation.value,
        dpr: typeof devicePixelRatio === "number" ? Math.min(2, Math.max(1, devicePixelRatio)) : 1,
        signal: controller.signal,
      })
      if (controller.signal.aborted || gen !== renderGeneration || requestId !== latestRequestId) return
      const url = URL.createObjectURL(blob)
      if (slot.url) URL.revokeObjectURL(slot.url)
      slot.url = url
      slot.state = "ready"
    } catch (err) {
      if ((err as Error)?.name === "AbortError" || controller.signal.aborted) return
      slot.state = "error"
    }
  }
}

// ── Main load flow ───────────────────────────────────────────────────
async function load(): Promise<void> {
  const source = effectiveSource.value
  const requestId = ++latestRequestId
  pdfLoadCoordinator.cancel()
  loading.value = !!source
  loadError.value = null
  await releaseDocument()
  if (requestId !== latestRequestId || unmounted) return
  if (!source) { loading.value = false; return }

  const task = pdfLoadCoordinator.start(toOfficeSource(source), {
    urlPolicy: props.urlPolicy,
  })
  let openedDocument: PdfRenderDocument | undefined
  let selectedRuntime: PdfRenderRuntime | undefined
  try {
    const verified = await loadVerifiedPdfSource(source, props.urlPolicy, {
      signal: task.signal,
      maxFileSize: props.maxFileSize,
    })
    task.assertCurrent()
    selectedRuntime = await selectRuntime()
    task.assertCurrent()
    const content = verified.bytes.buffer.slice(
      verified.bytes.byteOffset,
      verified.bytes.byteOffset + verified.bytes.byteLength,
    ) as ArrayBuffer
    openedDocument = await selectedRuntime.openDocument(content, { signal: task.signal })
    task.assertCurrent()
    if (openedDocument.pageCount <= 0) throw new Error("PDF document contains no pages.")
    runtimeForDocument = selectedRuntime
    renderDocument.value = openedDocument
    verifiedBlob.value = verified.blob
    initPageSlots()
    emit("document-load-success", openedDocument.pageCount)
    loading.value = false
    void renderAllPages(requestId, task.signal)
  } catch (cause) {
    if (openedDocument && selectedRuntime) {
      await selectedRuntime.closeDocument(openedDocument)
    }
    if (!task.isCurrent() || requestId !== latestRequestId || (cause as Error)?.name === "AbortError") return
    const sourceKind = source.kind
    const normalized = toPdfLoadError(cause, sourceKind, sourceKind === "url" ? source.url : undefined)
    loadError.value = normalized
    loading.value = false
    emit("document-load-error", normalized)
  }
}

// ── Watchers ─────────────────────────────────────────────────────────
watch([zoom, rotation], () => {
  if (!renderDocument.value) return
  cancelRender()
  initPageSlots()
  void renderAllPages(latestRequestId)
})

watch(effectiveSource, () => {
  void load()
}, { immediate: true })

// ── Page image style ─────────────────────────────────────────────────
function pageImgStyle(pageIndex: number): Record<string, string> | undefined {
  const doc = renderDocument.value
  if (!doc) return undefined
  const page = doc.pages[pageIndex]
  if (!page) return undefined
  const quarterTurns = ((page.rotation + rotation.value) / 90) % 4
  const width = (quarterTurns % 2 ? page.height : page.width) * zoom.value
  return {
    width: `${Math.max(1, Math.round(width))}px`,
    maxWidth: "100%",
    display: "block",
  }
}

// ── Scroll tracking ──────────────────────────────────────────────────
function onScroll(): void {
  const el = scrollRef.value
  if (!el || !renderDocument.value) return
  const center = el.scrollTop + el.clientHeight / 2
  let current = 0
  for (let i = 0; i < renderDocument.value.pageCount; i++) {
    const slot = el.querySelector(`[data-page-index="${i}"]`) as HTMLElement | null
    if (!slot) continue
    if (slot.offsetTop <= center) current = i
  }
  emit("visible-page-change", current)
}

function onContextMenu(event: MouseEvent): void {
  const el = scrollRef.value
  if (!el || !renderDocument.value) return
  const rect = el.getBoundingClientRect()
  const clickY = event.clientY - rect.top + el.scrollTop
  let pageIndex = 0
  for (let i = 0; i < renderDocument.value.pageCount; i++) {
    const slot = el.querySelector(`[data-page-index="${i}"]`) as HTMLElement | null
    if (!slot) continue
    if (clickY >= slot.offsetTop && clickY < slot.offsetTop + slot.offsetHeight) {
      pageIndex = i; break
    }
  }
  emit("contextMenu", {
    pageIndex,
    clientX: event.clientX,
    clientY: event.clientY,
    containerX: event.clientX - rect.left,
    containerY: event.clientY - rect.top,
  })
}

// ── fit-width zoom ───────────────────────────────────────────────────
watch([scrollRef, () => props.fitWidth, renderDocument], ([el, fit, doc]) => {
  if (!fit || !el || !doc) return
  const observer = new ResizeObserver(() => {
    const maxPageWidth = Math.max(
      ...doc.pages.map((p) => (p.rotation / 90) % 2 ? p.height : p.width),
      200,
    )
    const containerW = el.clientWidth - 32
    zoom.value = Math.min(2, Math.max(0.25, containerW / maxPageWidth))
  })
  observer.observe(el)
  return () => observer.disconnect()
})

// ── Expose ───────────────────────────────────────────────────────────
function scrollToPage(pageIndex: number): void {
  const el = scrollRef.value
  if (!el) return
  const slot = el.querySelector(`[data-page-index="${pageIndex}"]`) as HTMLElement | null
  slot?.scrollIntoView({ behavior: "smooth", block: "start" })
}

defineExpose({ scrollToPage, zoom, rotation })

onBeforeUnmount(async () => {
  unmounted = true
  cancelRender()
  await releaseDocument()
  if (ownedRuntime) await ownedRuntime.dispose()
})
</script>

<style scoped>
.pdf-surface {
  flex: 1;
  overflow-y: auto;
  overflow-x: hidden;
  min-height: 0;
  min-width: 0;
  background: var(--pdf-surface-bg, #525659);
}
.pdf-surface__empty {
  display: flex; align-items: center; justify-content: center;
  height: 200px; color: #9ca3af; font-size: 14px;
}
.pdf-surface__error {
  padding: 16px; color: #ef4444; font-size: 13px;
}
.pdf-surface__pages {
  display: flex; flex-direction: column; align-items: center;
  gap: 8px; padding: 16px 8px;
}
.pdf-surface__page-slot {
  position: relative;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3);
}
.pdf-surface__page-img { display: block; }
.pdf-surface__page-loading,
.pdf-surface__page-error,
.pdf-surface__page-placeholder {
  display: flex; align-items: center; justify-content: center;
  background: #fff; min-height: 200px; min-width: 200px;
  font-size: 12px; color: #6b7280;
}
.pdf-surface__page-error { color: #ef4444; }
</style>
