<script setup lang="ts">
import { computed, onBeforeUnmount, ref, watch } from "vue"
import {
  createLatestTaskCoordinator,
  createOfficeTaskSequence,
  type OfficeSource,
  type OfficeUrlPolicy,
} from "@arcships/office-runtime"
import {
  DEFAULT_PDF_MAX_FILE_SIZE,
  loadVerifiedPdfSource,
  toPdfLoadError,
  type PdfDiagnostic,
  type PdfLoadError,
  type PdfSource,
} from "../pdf-url-policy"
import {
  createPdfRenderRuntime,
  type PdfRenderDocument,
  type PdfRenderRuntime,
  type PdfRotation,
  type PdfSearchHit,
} from "../pdf/pdf-render-runtime"
import type { PdfViewerProps } from "../types"

const props = withDefaults(defineProps<PdfViewerProps>(), {
  src: undefined,
  source: undefined,
  urlPolicy: undefined,
  runtime: undefined,
  pdfiumWasmUrl: undefined,
  maxFileSize: DEFAULT_PDF_MAX_FILE_SIZE,
  fileName: undefined,
  defaultZoom: 1,
  className: "",
  showToolbar: true,
  showDownload: true,
  showRotateControls: true,
})

const emit = defineEmits<{
  "document-load-success": [numPages: number]
  "document-load-error": [error: PdfLoadError]
  "active-page-change": [pageNumber: number]
  diagnostic: [event: PdfDiagnostic]
}>()

const zoomOptions = [0.5, 0.75, 1, 1.25, 1.5, 2]
const zoom = ref(Math.min(2, Math.max(0.5, props.defaultZoom)))
const rotation = ref<PdfRotation>(0)
const activePage = ref(1)
const pageInput = ref("1")
const showThumbnails = ref(true)
const searchDraft = ref("")
const viewerRef = ref<HTMLElement | null>(null)
const searchInputRef = ref<HTMLInputElement | null>(null)
const searchResult = ref("")
const searchHits = ref<readonly PdfSearchHit[]>([])
const searchHitIndex = ref(-1)
const loadError = ref<PdfLoadError | null>(null)
const pageRenderError = ref("")
const downloadError = ref("")
const isLoading = ref(false)
const numPages = ref(0)
const verifiedBlob = ref<Blob | null>(null)
const loadedFileName = ref("")
const renderDocument = ref<PdfRenderDocument | null>(null)
const pageImageUrl = ref("")
const pageRenderState = ref<"idle" | "rendering" | "rendered" | "error">("idle")
const renderedPageNumber = ref(0)
const renderedZoom = ref(1)
const renderedRotation = ref<PdfRotation>(0)
const thumbnailUrls = ref<Record<number, string>>({})

let latestRequestId = 0
let latestTaskId: string | undefined
let renderGeneration = 0
let searchGeneration = 0
let thumbnailGeneration = 0
let renderAbort: AbortController | undefined
let searchAbort: AbortController | undefined
let thumbnailAbort: AbortController | undefined
let runtimeForDocument: PdfRenderRuntime | undefined
let ownedRuntime: PdfRenderRuntime | undefined
let ownedRuntimeWasmUrl: string | undefined
let unmounted = false

const pdfRuntimeId = `pdf-controller-${
  typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`
}`
const pdfTaskSequence = createOfficeTaskSequence(pdfRuntimeId)
const pdfLoadCoordinator = createLatestTaskCoordinator(pdfTaskSequence)

const effectiveSource = computed<PdfSource | undefined>(() => {
  if (props.source) return props.source
  if (props.src) return { kind: "url", url: props.src }
  return undefined
})
const viewerState = computed(() => {
  if (!effectiveSource.value) return "idle"
  if (isLoading.value || pageRenderState.value === "rendering") return "loading"
  if (loadError.value || pageRenderState.value === "error") return "error"
  return renderDocument.value && pageImageUrl.value ? "ready" : "idle"
})
const controlsDisabled = computed(() => !renderDocument.value || !!loadError.value)
const pageLabels = computed(() => Array.from({ length: numPages.value }, (_, index) => index + 1))
const currentSearchHit = computed(() => searchHits.value[searchHitIndex.value])
const pageImageStyle = computed(() => {
  const document = renderDocument.value
  const page = document?.pages[renderedPageNumber.value - 1]
  if (!page) return undefined
  const quarterTurns = ((page.rotation + renderedRotation.value) / 90) % 4
  const width = (quarterTurns % 2 ? page.height : page.width) * renderedZoom.value
  return { width: `${Math.max(1, Math.round(width))}px` }
})

function cn(...classes: (string | undefined | null | false)[]): string {
  return classes.filter(Boolean).join(" ")
}

function isAbortError(error: unknown): boolean {
  return error instanceof Error && (error.name === "AbortError" || /aborted|cancelled/i.test(error.message))
}

function getDownloadFileName(): string {
  if (props.fileName?.trim()) return props.fileName.toLowerCase().endsWith(".pdf") ? props.fileName : `${props.fileName}.pdf`
  if (loadedFileName.value) return loadedFileName.value
  if (effectiveSource.value?.kind === "blob") {
    const file = effectiveSource.value.blob as File
    if (typeof file.name === "string" && file.name.trim()) {
      return file.name.toLowerCase().endsWith(".pdf") ? file.name : `${file.name}.pdf`
    }
  }
  const raw =
    (effectiveSource.value?.kind === "url" ? effectiveSource.value.url.split(/[?#]/)[0]?.split("/").pop() : undefined) ||
    "document.pdf"
  return raw.toLowerCase().endsWith(".pdf") ? raw : `${raw}.pdf`
}

function emitDiagnostic(event: Omit<PdfDiagnostic, "requestId" | "taskId">): void {
  emit("diagnostic", {
    ...event,
    requestId: latestRequestId,
    taskId: latestTaskId,
    runtimeId: runtimeForDocument?.runtimeId,
  })
}

function revokeObjectUrl(url: string): void {
  if (!url) return
  try {
    URL.revokeObjectURL(url)
    emitDiagnostic({ type: "object-url-revoked" })
  } catch {
    // The URL is already unusable; local state is still cleared below.
  }
}

function replacePageImage(nextUrl = ""): void {
  const previous = pageImageUrl.value
  pageImageUrl.value = nextUrl
  if (previous && previous !== nextUrl) revokeObjectUrl(previous)
}

function clearThumbnails(): void {
  for (const url of Object.values(thumbnailUrls.value)) revokeObjectUrl(url)
  thumbnailUrls.value = {}
}

function cancelRenderWork(): void {
  renderGeneration++
  renderAbort?.abort()
  renderAbort = undefined
  searchGeneration++
  searchAbort?.abort()
  searchAbort = undefined
  thumbnailGeneration++
  thumbnailAbort?.abort()
  thumbnailAbort = undefined
}

function resetSearch(): void {
  searchResult.value = ""
  searchHits.value = []
  searchHitIndex.value = -1
}

function clearSearchAndRestoreFocus(): void {
  searchDraft.value = ""
  searchGeneration++
  searchAbort?.abort()
  searchAbort = undefined
  resetSearch()
  viewerRef.value?.focus()
}

function onViewerKeydown(event: KeyboardEvent): void {
  if (
    !props.showToolbar
    || !(event.ctrlKey || event.metaKey)
    || event.key.toLocaleLowerCase() !== "f"
  ) return
  event.preventDefault()
  searchInputRef.value?.focus()
}

async function releaseCurrentDocument(options: { closeDocument?: boolean } = {}): Promise<void> {
  cancelRenderWork()
  replacePageImage()
  clearThumbnails()
  pageRenderState.value = "idle"
  pageRenderError.value = ""
  renderedPageNumber.value = 0
  const document = renderDocument.value
  const runtime = runtimeForDocument
  renderDocument.value = null
  runtimeForDocument = undefined
  numPages.value = 0
  activePage.value = 1
  pageInput.value = "1"
  verifiedBlob.value = null
  loadedFileName.value = ""
  resetSearch()
  if (document && runtime && options.closeDocument !== false) await runtime.closeDocument(document)
}

async function selectRuntime(): Promise<PdfRenderRuntime> {
  if (props.runtime) {
    if (ownedRuntime) {
      const runtime = ownedRuntime
      ownedRuntime = undefined
      ownedRuntimeWasmUrl = undefined
      await runtime.dispose()
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

function createLinkedAbort(parent?: AbortSignal): { controller: AbortController; unlink: () => void } {
  const controller = new AbortController()
  const abort = () => controller.abort()
  if (parent?.aborted) controller.abort()
  else parent?.addEventListener("abort", abort, { once: true })
  return { controller, unlink: () => parent?.removeEventListener("abort", abort) }
}

async function confirmImageDecodes(url: string): Promise<void> {
  if (typeof Image === "undefined") return
  const image = new Image()
  image.src = url
  if (typeof image.decode === "function") await image.decode()
  else {
    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve()
      image.onerror = () => reject(new Error("Rendered PDF page is not a decodable image."))
    })
  }
  if (image.naturalWidth <= 0 || image.naturalHeight <= 0) {
    throw new Error("Rendered PDF page has no visible pixels.")
  }
}

function zoomIn(): void {
  const next = zoomOptions.find((value) => value > zoom.value)
  if (next) zoom.value = next
}

function zoomOut(): void {
  const next = [...zoomOptions].reverse().find((value) => value < zoom.value)
  if (next) zoom.value = next
}

function rotateClockwise(): void {
  rotation.value = ((rotation.value + 90) % 360) as PdfRotation
}

function setActivePage(page: number): void {
  const next = Math.max(1, Math.min(numPages.value || 1, Math.trunc(page) || 1))
  activePage.value = next
  pageInput.value = String(next)
  emit("active-page-change", next)
}

function commitPageInput(): void {
  setActivePage(Number(pageInput.value))
}

async function handleDownload(): Promise<void> {
  if (!verifiedBlob.value) return
  downloadError.value = ""
  try {
    const url = URL.createObjectURL(verifiedBlob.value)
    const anchor = document.createElement("a")
    anchor.href = url
    anchor.download = getDownloadFileName()
    anchor.rel = "noopener"
    document.body.append(anchor)
    anchor.click()
    anchor.remove()
    setTimeout(() => revokeObjectUrl(url), 0)
    emitDiagnostic({
      type: "download",
      sourceKind: effectiveSource.value?.kind,
      bytes: verifiedBlob.value.size,
    })
  } catch (error) {
    downloadError.value = error instanceof Error ? error.message : String(error)
  }
}

async function renderCurrentPage(
  requestId: number,
  parentSignal?: AbortSignal,
  failDocumentOnError = false,
): Promise<void> {
  const document = renderDocument.value
  const runtime = runtimeForDocument
  if (!document || !runtime) return
  const pageNumber = activePage.value
  const renderZoom = zoom.value
  const renderRotation = rotation.value
  const generation = ++renderGeneration
  renderAbort?.abort()
  const { controller, unlink } = createLinkedAbort(parentSignal)
  renderAbort = controller
  pageRenderError.value = ""
  pageRenderState.value = "rendering"
  replacePageImage()
  emitDiagnostic({
    type: "render-start",
    pageNumber,
    pageCount: document.pageCount,
    zoom: renderZoom,
    rotation: renderRotation,
  })
  let candidateUrl = ""
  try {
    const blob = await runtime.renderPage(document, pageNumber - 1, {
      scale: renderZoom,
      rotation: renderRotation,
      dpr: typeof devicePixelRatio === "number" ? Math.min(2, Math.max(1, devicePixelRatio)) : 1,
      signal: controller.signal,
    })
    if (
      controller.signal.aborted ||
      generation !== renderGeneration ||
      requestId !== latestRequestId ||
      document.id !== renderDocument.value?.id
    ) return
    candidateUrl = URL.createObjectURL(blob)
    await confirmImageDecodes(candidateUrl)
    if (
      controller.signal.aborted ||
      generation !== renderGeneration ||
      requestId !== latestRequestId ||
      document.id !== renderDocument.value?.id
    ) return
    replacePageImage(candidateUrl)
    candidateUrl = ""
    renderedPageNumber.value = pageNumber
    renderedZoom.value = renderZoom
    renderedRotation.value = renderRotation
    pageRenderState.value = "rendered"
    emitDiagnostic({
      type: "render-success",
      pageNumber,
      pageCount: document.pageCount,
      zoom: renderZoom,
      rotation: renderRotation,
      bytes: blob.size,
    })
  } catch (error) {
    if (controller.signal.aborted || generation !== renderGeneration || requestId !== latestRequestId || isAbortError(error)) return
    pageRenderState.value = "error"
    pageRenderError.value = "PDF 页面渲染失败。"
    const source = effectiveSource.value
    const normalized = toPdfLoadError(error, source?.kind ?? "blob", source?.kind === "url" ? source.url : undefined)
    emitDiagnostic({ type: "render-error", pageNumber, pageCount: document.pageCount, error: normalized })
    if (failDocumentOnError) throw error
  } finally {
    if (candidateUrl) revokeObjectUrl(candidateUrl)
    unlink()
    if (renderAbort === controller) renderAbort = undefined
  }
}

async function renderThumbnails(requestId: number): Promise<void> {
  const document = renderDocument.value
  const runtime = runtimeForDocument
  if (!document || !runtime) return
  const generation = ++thumbnailGeneration
  thumbnailAbort?.abort()
  const controller = new AbortController()
  thumbnailAbort = controller
  let nextIndex = 0
  const worker = async () => {
    while (nextIndex < document.pageCount) {
      const pageIndex = nextIndex++
      let url = ""
      try {
        const blob = await runtime.renderThumbnail(document, pageIndex, { signal: controller.signal })
        if (
          controller.signal.aborted ||
          generation !== thumbnailGeneration ||
          requestId !== latestRequestId ||
          document.id !== renderDocument.value?.id
        ) return
        url = URL.createObjectURL(blob)
        thumbnailUrls.value = { ...thumbnailUrls.value, [pageIndex + 1]: url }
        url = ""
      } catch (error) {
        if (!controller.signal.aborted && !isAbortError(error)) {
          // A failed thumbnail must not hide an otherwise usable page.
        }
      } finally {
        if (url) revokeObjectUrl(url)
      }
    }
  }
  await Promise.all([worker(), worker()])
  if (thumbnailAbort === controller) thumbnailAbort = undefined
}

async function submitSearch(): Promise<void> {
  const query = searchDraft.value.trim()
  const document = renderDocument.value
  const runtime = runtimeForDocument
  searchAbort?.abort()
  const generation = ++searchGeneration
  searchHits.value = []
  searchHitIndex.value = -1
  if (!query) {
    searchResult.value = ""
    return
  }
  if (!document || !runtime) return
  const controller = new AbortController()
  searchAbort = controller
  searchResult.value = "正在搜索…"
  emitDiagnostic({ type: "search-start" })
  try {
    const hits = await runtime.search(document, query, { signal: controller.signal })
    if (
      controller.signal.aborted ||
      generation !== searchGeneration ||
      document.id !== renderDocument.value?.id
    ) return
    searchHits.value = hits
    if (hits.length) {
      searchHitIndex.value = 0
      setActivePage(hits[0].pageIndex + 1)
      searchResult.value = `第 1 / ${hits.length} 项，第 ${hits[0].pageIndex + 1} 页`
    } else {
      searchResult.value = "没有找到"
    }
    emitDiagnostic({ type: "search-success", matches: hits.length })
  } catch (error) {
    if (controller.signal.aborted || generation !== searchGeneration || isAbortError(error)) return
    searchResult.value = "搜索失败"
    const source = effectiveSource.value
    emitDiagnostic({
      type: "search-error",
      error: toPdfLoadError(error, source?.kind ?? "blob", source?.kind === "url" ? source.url : undefined),
    })
  } finally {
    if (searchAbort === controller) searchAbort = undefined
  }
}

function moveSearch(delta: number): void {
  if (!searchHits.value.length) return
  const next = (searchHitIndex.value + delta + searchHits.value.length) % searchHits.value.length
  searchHitIndex.value = next
  const hit = searchHits.value[next]
  setActivePage(hit.pageIndex + 1)
  searchResult.value = `第 ${next + 1} / ${searchHits.value.length} 项，第 ${hit.pageIndex + 1} 页`
}

function toOfficeSource(source: PdfSource): OfficeSource {
  return source.kind === "url"
    ? { kind: "url", url: source.url }
    : {
        kind: "file",
        name: source.fileName,
        file: {
          name: source.fileName,
          size: source.blob.size,
          type: source.blob.type,
          arrayBuffer: () => source.blob.arrayBuffer(),
        },
      }
}

async function load(): Promise<void> {
  const source = effectiveSource.value
  const requestId = ++latestRequestId
  pdfLoadCoordinator.cancel()
  latestTaskId = undefined
  isLoading.value = !!source
  loadError.value = null
  downloadError.value = ""
  rotation.value = 0
  zoom.value = Math.min(2, Math.max(0.5, props.defaultZoom))
  await releaseCurrentDocument()
  if (requestId !== latestRequestId || unmounted) return
  if (!source) {
    isLoading.value = false
    return
  }

  const task = pdfLoadCoordinator.start(toOfficeSource(source), {
    urlPolicy: props.urlPolicy as OfficeUrlPolicy | undefined,
  })
  latestTaskId = task.context.taskId
  emitDiagnostic({
    type: "load-start",
    sourceKind: source.kind,
    url: source.kind === "url" ? source.url : undefined,
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
    numPages.value = openedDocument.pageCount
    activePage.value = 1
    pageInput.value = "1"
    verifiedBlob.value = verified.blob
    loadedFileName.value = verified.fileName
    await renderCurrentPage(requestId, task.signal, true)
    task.assertCurrent()
    emit("document-load-success", openedDocument.pageCount)
    emit("active-page-change", 1)
    emitDiagnostic({
      type: "load-success",
      sourceKind: verified.sourceKind,
      url: verified.resolvedUrl,
      bytes: verified.bytes.byteLength,
      pageNumber: 1,
      pageCount: openedDocument.pageCount,
    })
    void renderThumbnails(requestId)
  } catch (error) {
    if (openedDocument && selectedRuntime) {
      await selectedRuntime.closeDocument(openedDocument)
    }
    if (!task.isCurrent() || requestId !== latestRequestId || isAbortError(error)) return
    const normalized = toPdfLoadError(error, source.kind, source.kind === "url" ? source.url : undefined)
    const pdfError: PdfLoadError = normalized.code === "INVALID_PDF" && !(error instanceof Error && /来源|文件不是|application\/pdf/.test(error.message))
      ? { ...normalized, message: "PDF 文件无法解析或渲染。" }
      : normalized
    loadError.value = pdfError
    pageRenderState.value = "error"
    replacePageImage()
    clearThumbnails()
    verifiedBlob.value = null
    numPages.value = 0
    renderDocument.value = null
    runtimeForDocument = undefined
    emit("document-load-error", pdfError)
    emitDiagnostic({ type: "load-error", sourceKind: source.kind, error: pdfError })
  } finally {
    if (task.isCurrent() && requestId === latestRequestId) isLoading.value = false
    task.finish()
  }
}

watch(
  () => [props.src, props.source, props.urlPolicy, props.runtime, props.pdfiumWasmUrl, props.maxFileSize],
  () => void load(),
  { immediate: true },
)
watch(activePage, (page) => { pageInput.value = String(page) })
watch([activePage, zoom, rotation], () => {
  if (renderDocument.value && !isLoading.value) void renderCurrentPage(latestRequestId)
})

onBeforeUnmount(() => {
  unmounted = true
  latestRequestId++
  latestTaskId = undefined
  pdfLoadCoordinator.dispose()
  pdfTaskSequence.dispose()
  const runtime = ownedRuntime
  ownedRuntime = undefined
  const closeDocument = !runtime || runtimeForDocument !== runtime
  void releaseCurrentDocument({ closeDocument }).finally(() => runtime?.dispose())
})
</script>

<template>
  <div
    ref="viewerRef"
    :class="cn('pdf-viewer', props.className)"
    data-testid="pdf-viewer"
    tabindex="0"
    @keydown="onViewerKeydown"
  >
    <div v-if="props.showToolbar" data-testid="pdf-toolbar" class="pdf-toolbar">
      <button
        data-testid="pdf-sidebar-toggle"
        type="button"
        :disabled="controlsDisabled"
        aria-label="显示或隐藏缩略图"
        :aria-pressed="showThumbnails"
        @click="showThumbnails = !showThumbnails"
      >☰</button>
      <span class="file-name">{{ getDownloadFileName() }}</span>
      <button data-testid="pdf-first-page" type="button" :disabled="controlsDisabled || activePage <= 1" aria-label="第一页" @click="setActivePage(1)">«</button>
      <button data-testid="pdf-previous-page" type="button" :disabled="controlsDisabled || activePage <= 1" aria-label="上一页" @click="setActivePage(activePage - 1)">‹</button>
      <input
        v-model="pageInput"
        data-testid="pdf-page-input"
        class="page-input"
        type="number"
        min="1"
        :max="numPages || 1"
        :disabled="controlsDisabled"
        aria-label="页码"
        @change="commitPageInput"
        @keydown.enter.prevent="commitPageInput"
      />
      <span class="page-indicator">/ <span data-testid="pdf-page-count">{{ numPages || '—' }}</span></span>
      <button data-testid="pdf-next-page" type="button" :disabled="controlsDisabled || activePage >= numPages" aria-label="下一页" @click="setActivePage(activePage + 1)">›</button>
      <button data-testid="pdf-last-page" type="button" :disabled="controlsDisabled || activePage >= numPages" aria-label="最后一页" @click="setActivePage(numPages)">»</button>
      <button data-testid="pdf-zoom-out" type="button" :disabled="controlsDisabled || zoom <= zoomOptions[0]" aria-label="缩小" @click="zoomOut">−</button>
      <select v-model.number="zoom" data-testid="pdf-zoom-select" :disabled="controlsDisabled" aria-label="缩放比例">
        <option v-for="value in zoomOptions" :key="value" :value="value">{{ Math.round(value * 100) }}%</option>
      </select>
      <button data-testid="pdf-zoom-in" type="button" :disabled="controlsDisabled || zoom >= zoomOptions[zoomOptions.length - 1]" aria-label="放大" @click="zoomIn">+</button>
      <button v-if="props.showRotateControls" data-testid="pdf-rotate" type="button" :disabled="controlsDisabled" aria-label="顺时针旋转" @click="rotateClockwise">⟳</button>
      <form class="pdf-search" role="search" @submit.prevent="submitSearch">
        <input
          ref="searchInputRef"
          v-model="searchDraft"
          data-testid="pdf-search-input"
          :disabled="controlsDisabled"
          aria-label="搜索 PDF 文本"
          placeholder="搜索文本"
          @keydown.esc.prevent="clearSearchAndRestoreFocus"
        />
        <button data-testid="pdf-search-submit" type="submit" :disabled="controlsDisabled || !searchDraft.trim()">搜索</button>
        <button data-testid="pdf-search-prev" type="button" :disabled="!searchHits.length" aria-label="上一个搜索结果" @click="moveSearch(-1)">↑</button>
        <button data-testid="pdf-search-next" type="button" :disabled="!searchHits.length" aria-label="下一个搜索结果" @click="moveSearch(1)">↓</button>
      </form>
      <span v-if="searchResult" data-testid="pdf-search-result" class="search-result" role="status">{{ searchResult }}</span>
      <button v-if="props.showDownload" data-testid="pdf-download" type="button" :disabled="!verifiedBlob || !!loadError" aria-label="下载 PDF" @click="handleDownload">⬇</button>
    </div>

    <div data-testid="pdf-status" class="pdf-status" :data-state="viewerState" aria-live="polite">{{ viewerState }}</div>
    <div v-if="!effectiveSource" class="pdf-empty">No PDF loaded</div>
    <div v-else-if="isLoading && !renderDocument" class="pdf-empty">正在打开 PDF…</div>
    <div v-else-if="loadError" data-testid="load-error" class="pdf-error" :data-error-code="loadError.code">
      {{ loadError.code }}: {{ loadError.message }}
    </div>
    <div
      v-else-if="renderDocument"
      data-testid="pdf-document"
      class="pdf-body"
      :data-page-count="numPages"
      :data-active-page="activePage"
      :data-zoom="zoom"
      :data-rotation="rotation"
    >
      <aside v-if="showThumbnails" data-testid="pdf-thumbnails" class="pdf-thumbnails" aria-label="页面缩略图">
        <button
          v-for="page in pageLabels"
          :key="page"
          data-testid="pdf-thumbnail"
          type="button"
          :data-page-number="page"
          :aria-current="page === activePage ? 'page' : undefined"
          :class="['thumbnail', { active: page === activePage }]"
          @click="setActivePage(page)"
        >
          <img
            v-if="thumbnailUrls[page]"
            data-testid="pdf-thumbnail-image"
            :src="thumbnailUrls[page]"
            :alt="`第 ${page} 页缩略图`"
          />
          <span v-else class="thumbnail-placeholder">{{ page }}</span>
          <span>第 {{ page }} 页</span>
        </button>
      </aside>
      <main class="pdf-frame-wrap">
        <div v-if="currentSearchHit" class="pdf-search-context" role="status">
          <span>{{ currentSearchHit.before }}</span><mark>{{ currentSearchHit.match }}</mark><span>{{ currentSearchHit.after }}</span>
        </div>
        <div
          data-testid="pdf-page"
          class="pdf-page-stage"
          :data-page-number="activePage"
          :data-render-state="pageRenderState"
        >
          <div v-if="pageRenderState === 'rendering'" class="page-loading">正在渲染第 {{ activePage }} 页…</div>
          <div v-if="pageRenderError" class="pdf-error page-error">{{ pageRenderError }}</div>
          <div v-if="pageImageUrl" data-testid="pdf-page-render" class="pdf-page-render">
            <img
              data-testid="pdf-page-image"
              class="pdf-page-image"
              :src="pageImageUrl"
              :style="pageImageStyle"
              :alt="`PDF 第 ${renderedPageNumber} 页`"
            />
          </div>
        </div>
      </main>
    </div>
    <div v-if="downloadError" class="download-error">{{ downloadError }}</div>
  </div>
</template>

<style scoped>
.pdf-viewer { display: flex; flex-direction: column; height: 100%; min-height: 620px; background: var(--background, #fff); color: var(--foreground, #111); }
.pdf-toolbar { display: flex; align-items: center; gap: 7px; min-height: 48px; padding: 8px; border-bottom: 1px solid var(--border, #e5e5e5); flex-wrap: wrap; }
.pdf-toolbar button, .pdf-toolbar select, .pdf-toolbar input { height: 32px; border: 1px solid var(--border, #d4d4d4); border-radius: 6px; background: white; padding: 0 8px; }
.pdf-toolbar button { cursor: pointer; }
.pdf-toolbar button:focus-visible, .pdf-toolbar select:focus-visible, .pdf-toolbar input:focus-visible, .thumbnail:focus-visible { outline: 2px solid #2563eb; outline-offset: 2px; }
.pdf-toolbar button:disabled, .pdf-toolbar input:disabled, .pdf-toolbar select:disabled { opacity: .45; cursor: not-allowed; }
.file-name { max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 13px; font-weight: 600; margin-right: auto; }
.page-input { width: 58px; text-align: center; }
.page-indicator { font-size: 12px; white-space: nowrap; }
.pdf-search { display: flex; align-items: center; gap: 4px; }
.pdf-search input { width: 150px; }
.search-result { font-size: 12px; color: #1d4ed8; white-space: nowrap; }
.pdf-status { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); clip-path: inset(50%); white-space: nowrap; }
.pdf-empty, .pdf-error { display: grid; place-items: center; flex: 1; min-height: 220px; color: var(--muted-foreground, #737373); }
.pdf-error, .download-error { color: #b91c1c; }
.download-error { padding: 8px 12px; font-size: 12px; border-top: 1px solid var(--border, #e5e5e5); }
.pdf-body { display: flex; min-height: 540px; flex: 1; min-width: 0; background: #e5e7eb; }
.pdf-thumbnails { width: 150px; flex: 0 0 150px; overflow: auto; border-right: 1px solid var(--border, #d4d4d4); padding: 10px; background: #fafafa; }
.thumbnail { display: flex; width: 100%; min-height: 118px; margin-bottom: 10px; padding: 6px; align-items: center; justify-content: center; flex-direction: column; gap: 5px; overflow: hidden; border: 2px solid transparent; border-radius: 6px; background: transparent; color: #525252; font-size: 11px; cursor: pointer; }
.thumbnail img { display: block; max-width: 106px; max-height: 86px; object-fit: contain; border: 1px solid #d4d4d4; background: white; box-shadow: 0 1px 3px rgb(0 0 0 / 12%); }
.thumbnail-placeholder { display: grid; place-items: center; width: 74px; height: 84px; border: 1px solid #d4d4d4; background: white; color: #a3a3a3; }
.thumbnail.active { border-color: #2563eb; background: #eff6ff; color: #1d4ed8; font-weight: 700; }
.pdf-frame-wrap { flex: 1; min-width: 0; min-height: 540px; overflow: auto; position: relative; padding: 24px; }
.pdf-page-stage { display: flex; min-width: min-content; min-height: 100%; align-items: flex-start; justify-content: center; position: relative; }
.pdf-page-render { width: fit-content; margin: 0 auto; }
.pdf-page-image { display: block; max-width: none; height: auto; border: 0; background: white; box-shadow: 0 3px 14px rgb(0 0 0 / 20%); }
.page-loading { position: absolute; z-index: 2; top: 12px; left: 50%; transform: translateX(-50%); border: 1px solid #bfdbfe; border-radius: 999px; background: #eff6ff; color: #1d4ed8; padding: 5px 12px; font-size: 12px; white-space: nowrap; }
.page-error { min-height: 180px; min-width: 300px; border-radius: 8px; background: white; padding: 24px; box-shadow: 0 2px 8px rgb(0 0 0 / 12%); }
.pdf-search-context { position: sticky; top: 0; z-index: 3; width: fit-content; max-width: min(720px, 90%); margin: 0 auto 10px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; border: 1px solid #bfdbfe; border-radius: 8px; background: #eff6ff; color: #1e3a8a; font-size: 12px; padding: 6px 10px; }
.pdf-search-context mark { background: #fde047; color: #111827; padding: 1px 2px; }
@media (max-width: 760px) {
  .pdf-body { flex-direction: column; }
  .pdf-thumbnails { width: auto; flex-basis: auto; display: flex; gap: 8px; border-right: 0; border-bottom: 1px solid var(--border, #d4d4d4); }
  .thumbnail { width: 108px; min-height: 104px; flex: 0 0 auto; margin-bottom: 0; }
  .pdf-frame-wrap { padding: 12px; }
  .pdf-search { order: 3; width: 100%; }
  .pdf-search input { flex: 1; width: auto; }
}
</style>
