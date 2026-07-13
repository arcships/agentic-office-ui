<template>
  <div
    ref="viewerRef"
    class="docx-viewer"
    :class="[{ 'docx-viewer--dark': darkTheme }, className]"
    data-testid="docx-viewer"
    :data-state="viewerState"
    :data-show-tracked-changes="String(effectiveShowTrackedChanges)"
    :data-show-comments="String(effectiveShowComments)"
    tabindex="0"
    @keydown="onViewerKeydown"
  >
    <input
      ref="uploadInputRef"
      class="docx-viewer__file-input"
      type="file"
      accept=".docx,.docm,.dotx,.dotm,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.ms-word.document.macroEnabled.12,application/vnd.openxmlformats-officedocument.wordprocessingml.template,application/vnd.ms-word.template.macroEnabled.12"
      @change="onUploadChange"
    />

    <DocxViewerToolbar
      v-if="showToolbar"
      ref="toolbarRef"
      :can-download="Boolean(effectiveFile)"
      :current-page="currentPage"
      :disabled="isLoading || Boolean(error) || !renderableModel"
      :file-name="displayFileName"
      :is-dark="darkTheme"
      :sidebar-open="sidebarOpen"
      :show-tracked-changes="effectiveShowTrackedChanges"
      :show-comments="effectiveShowComments"
      :tracked-change-count="viewerTrackedChanges.length"
      :comment-count="viewerComments.length"
      :total-pages="totalPages"
      :zoom="zoomPercent"
      :search-query="searchQuery"
      :search-result-count="searchState.matches.length"
      :search-result-index="searchState.activeIndex"
      @download="downloadFile"
      @select-page="selectPage"
      @toggle-sidebar="sidebarOpen = !sidebarOpen"
      @toggle-tracked-changes="toggleTrackedChanges"
      @toggle-comments="toggleComments"
      @update:search-query="searchQuery = $event"
      @search-next="selectSearchResult(1)"
      @search-previous="selectSearchResult(-1)"
      @search-close="closeSearch"
      @toggle-theme="toggleTheme"
      @update:zoom="zoomPercent = $event"
      @upload="uploadInputRef?.click()"
    />

    <div class="docx-viewer__body">
      <DocxThumbnailPanel
        v-if="sidebarOpen && renderableModel"
        :model="renderableModel"
        :layout-options="layoutOptions"
        :current-page="currentPage"
        :total-pages="totalPages"
        :show-toggle="false"
        @select-page="selectPageIndex"
      />

      <main class="docx-viewer__content">
        <div v-if="isLoading" class="docx-viewer-state" data-testid="docx-loading-state">
          <div class="docx-viewer-loading-spinner" />
          <div>
            <strong>Opening document</strong>
            <span>Loading DOCX…</span>
          </div>
        </div>

        <div
          v-else-if="error"
          class="docx-viewer-state docx-viewer-state--error"
          data-testid="load-error"
          :data-error-code="errorCode(error)"
          :data-error-phase="errorDetail(error, 'phase')"
          :data-error-limit="errorDetail(error, 'limit')"
          :data-error-actual="errorNumber(error, 'actual')"
          :data-error-allowed="errorNumber(error, 'allowed')"
        >
          <svg viewBox="0 0 24 24" aria-hidden="true">
            <path d="M12 3 2.8 19h18.4L12 3Z" /><path d="M12 9v4M12 17h.01" />
          </svg>
          <div>
            <strong>Unable to open this document</strong>
            <span v-if="errorCode(error)" class="docx-viewer-error-code">{{ errorCode(error) }}</span>
            <span>{{ error.message }}</span>
            <button type="button" @click="uploadInputRef?.click()">Choose another Word file</button>
          </div>
        </div>

        <div v-else-if="!renderableModel" class="docx-viewer-state docx-viewer-state--empty">
          <svg viewBox="0 0 24 24" aria-hidden="true">
            <path d="M6 2.8h8l4 4V21H6zM14 2.8V7h4" />
          </svg>
          <div>
            <strong>No document open</strong>
            <span>{{ emptyState }}</span>
            <button type="button" @click="uploadInputRef?.click()">Choose Word file</button>
          </div>
        </div>

        <DocxDocumentSurface
          v-else
          ref="documentSurfaceRef"
          :model="renderableModel"
          :layout-options="layoutOptions"
          :theme="darkTheme ? 'dark' : 'light'"
          :zoom="zoomPercent / 100"
          :editable="false"
          :show-tracked-changes="effectiveShowTrackedChanges"
          :show-comments="effectiveShowComments"
          @page-count-change="onPageCountChange"
          @visible-page-range="onVisiblePageRange"
          @search-state-change="onSearchStateChange"
          @update:zoom="zoomPercent = Math.round($event * 100)"
        />
      </main>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, shallowRef, watch } from "vue"
import type { DocModel, DocxImportResult, DocxRuntime, DocxRuntimeLoader, LayoutOptions } from "@arcships/docx-core"
import { collectCommentsFromModel, collectTrackedChangesFromModel, createDocxRuntime, DocxImportError } from "@arcships/docx-core"
import type { DocxSearchState } from "../composables/useDocxSearch"
import DocxDocumentSurface from "./DocxDocumentSurface.vue"
import DocxThumbnailPanel from "./DocxThumbnailPanel.vue"
import DocxViewerToolbar from "./DocxViewerToolbar.vue"

const props = withDefaults(
  defineProps<{
    file?: ArrayBuffer
    model?: DocModel
    className?: string
    layoutOptions?: LayoutOptions
    emptyState?: string
    runtime?: DocxRuntime
    fileName?: string
    showToolbar?: boolean
    defaultThumbnailsOpen?: boolean
    defaultZoom?: number
    isDark?: boolean
    defaultShowTrackedChanges?: boolean
    defaultShowComments?: boolean
    showTrackedChanges?: boolean | null
    showComments?: boolean | null
  }>(),
  {
    emptyState: "Upload or provide a DOCX file to start.",
    fileName: "document.docx",
    showToolbar: true,
    defaultThumbnailsOpen: false,
    defaultZoom: 100,
    isDark: false,
    defaultShowTrackedChanges: true,
    defaultShowComments: true,
    showTrackedChanges: null,
    showComments: null,
  },
)

const emit = defineEmits<{
  (event: "load-start"): void
  (event: "load-success", result: DocxImportResult): void
  (event: "load-error", error: Error): void
  (event: "file-selected", file: File): void
  (event: "update:isDark", isDark: boolean): void
  (event: "update:showTrackedChanges", visible: boolean): void
  (event: "update:showComments", visible: boolean): void
  (event: "searchStateChange", state: DocxSearchState): void
}>()

const isLoading = ref(false)
const error = ref<Error | undefined>()
const parsedModel = ref<DocModel | undefined>()
const layoutLimitExceeded = ref(false)
const pendingLoadResult = shallowRef<DocxImportResult | undefined>()
const internalFile = ref<ArrayBuffer | undefined>()
const internalFileName = ref("")
const uploadInputRef = ref<HTMLInputElement>()
const viewerRef = ref<HTMLElement>()
const toolbarRef = ref<InstanceType<typeof DocxViewerToolbar>>()
const documentSurfaceRef = ref<InstanceType<typeof DocxDocumentSurface>>()
const currentPage = ref(1)
const totalPages = ref(0)
const sidebarOpen = ref(props.defaultThumbnailsOpen)
const zoomPercent = ref(normalizeZoom(props.defaultZoom))
const darkTheme = ref(props.isDark)
const internalShowTrackedChanges = ref(props.defaultShowTrackedChanges)
const internalShowComments = ref(props.defaultShowComments)
const searchQuery = ref("")
const searchState = shallowRef<DocxSearchState>({
  status: "idle",
  query: "",
  matches: [],
  activeIndex: -1,
})

const effectiveFile = computed(() => internalFile.value ?? props.file)
const resolvedModel = computed(() => props.model ?? parsedModel.value)
const renderableModel = computed(() => layoutLimitExceeded.value ? undefined : resolvedModel.value)
const displayFileName = computed(() => internalFileName.value || props.fileName)
const effectiveShowTrackedChanges = computed(() => props.showTrackedChanges ?? internalShowTrackedChanges.value)
const effectiveShowComments = computed(() => props.showComments ?? internalShowComments.value)
const viewerTrackedChanges = computed(() => renderableModel.value ? collectTrackedChangesFromModel(renderableModel.value) : [])
const viewerComments = computed(() => renderableModel.value ? collectCommentsFromModel(renderableModel.value) : [])
const viewerState = computed(() => {
  if (isLoading.value || pendingLoadResult.value) return "loading"
  if (error.value) return "error"
  if (renderableModel.value) return "ready"
  return "idle"
})

const ownedRuntime = createDocxRuntime()
let loaderRuntime: DocxRuntime = props.runtime ?? ownedRuntime
let runtimeLoader: DocxRuntimeLoader = loaderRuntime.createLoader()

function resolveRuntimeLoader(): DocxRuntimeLoader {
  const nextRuntime = props.runtime ?? ownedRuntime
  if (nextRuntime !== loaderRuntime) {
    runtimeLoader.dispose()
    loaderRuntime = nextRuntime
    runtimeLoader = loaderRuntime.createLoader()
  }
  return runtimeLoader
}

watch(
  () => props.file,
  () => {
    internalFile.value = undefined
    internalFileName.value = ""
  },
)

watch(() => props.isDark, (value) => { darkTheme.value = value })
watch(() => props.defaultZoom, (value) => { zoomPercent.value = normalizeZoom(value) })
watch(() => props.defaultShowTrackedChanges, (value) => {
  if (props.showTrackedChanges == null) internalShowTrackedChanges.value = value
})
watch(() => props.defaultShowComments, (value) => {
  if (props.showComments == null) internalShowComments.value = value
})

watch(
  [searchQuery, renderableModel],
  async ([query, model], _previous, onCleanup) => {
    let current = true
    onCleanup(() => { current = false })
    await nextTick()
    if (!current) return
    const surface = documentSurfaceRef.value
    if (!surface || !model) {
      searchState.value = { status: "idle", query: "", matches: [], activeIndex: -1 }
      return
    }
    if (!query.trim()) {
      surface.clearSearch()
      return
    }
    void surface.search(query).catch(ignoreReplacedSearch)
  },
  { flush: "post" },
)

function ignoreReplacedSearch(error: unknown): void {
  if (error instanceof Error && error.name === "AbortError") return
  // Non-abort failures are published by the surface through searchStateChange.
}

function onSearchStateChange(next: DocxSearchState): void {
  searchState.value = next
  emit("searchStateChange", next)
}

function selectSearchResult(direction: -1 | 1): void {
  if (!searchState.value.matches.length) return
  const task = direction > 0
    ? documentSurfaceRef.value?.searchNext()
    : documentSurfaceRef.value?.searchPrevious()
  void task?.catch(ignoreReplacedSearch)
}

function onViewerKeydown(event: KeyboardEvent): void {
  if (!props.showToolbar || !(event.ctrlKey || event.metaKey) || event.key.toLocaleLowerCase() !== "f") return
  event.preventDefault()
  toolbarRef.value?.focusSearch()
}

function closeSearch(): void {
  searchQuery.value = ""
  documentSurfaceRef.value?.clearSearch()
  void nextTick(() => viewerRef.value?.focus())
}

watch(
  () => [effectiveFile.value, props.runtime, props.model] as const,
  async ([newFile], _oldValue, onCleanup) => {
    let active = true
    onCleanup(() => { active = false })
    const loader = resolveRuntimeLoader()
    pendingLoadResult.value = undefined
    layoutLimitExceeded.value = false
    currentPage.value = 1
    totalPages.value = 0

    if (!newFile) {
      loader.cancel()
      parsedModel.value = undefined
      error.value = undefined
      isLoading.value = false
      return
    }
    if (props.model) {
      loader.cancel()
      parsedModel.value = undefined
      error.value = undefined
      isLoading.value = false
      return
    }

    isLoading.value = true
    error.value = undefined
    emit("load-start")
    try {
      const result = await loader.load(newFile, { transferBuffer: false })
      if (!active) return
      pendingLoadResult.value = result
      parsedModel.value = result.model
    } catch (cause) {
      if (!active) return
      const normalized = cause instanceof Error ? cause : new Error("Unknown DOCX parse error")
      error.value = normalized
      pendingLoadResult.value = undefined
      emit("load-error", normalized)
    } finally {
      if (active) isLoading.value = false
    }
  },
  { immediate: true },
)

function onPageCountChange(count: number): void {
  const allowed = loaderRuntime.limits?.maxDocxPages
  if (allowed !== undefined && count > allowed) {
    const normalized = new DocxImportError(
      "LIMIT_EXCEEDED",
      `DOCX 页数 ${count} 超过允许值 ${allowed}。`,
      { phase: "layout", limit: "maxDocxPages", actual: count, allowed },
    )
    pendingLoadResult.value = undefined
    layoutLimitExceeded.value = true
    parsedModel.value = undefined
    error.value = normalized
    emit("load-error", normalized)
    return
  }

  totalPages.value = Math.max(0, count)
  if (pendingLoadResult.value) {
    const result = pendingLoadResult.value
    pendingLoadResult.value = undefined
    emit("load-success", result)
  }
}

function onVisiblePageRange(range: { startPageIndex: number; endPageIndex: number }): void {
  if (range.endPageIndex < range.startPageIndex) return
  currentPage.value = Math.floor((range.startPageIndex + range.endPageIndex) / 2) + 1
}

function selectPage(page: number): void {
  selectPageIndex(page - 1)
}

function selectPageIndex(pageIndex: number): void {
  documentSurfaceRef.value?.scrollToPage(pageIndex)
  currentPage.value = pageIndex + 1
}

async function onUploadChange(event: Event): Promise<void> {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  input.value = ""
  if (!file) return
  emit("file-selected", file)
  internalFileName.value = file.name
  internalFile.value = await file.arrayBuffer()
}

function downloadFile(): void {
  const bytes = effectiveFile.value
  if (!bytes) return
  const extension = /\.([^.]+)$/.exec(displayFileName.value)?.[1]?.toLowerCase()
  const mimeType = extension === "docm"
    ? "application/vnd.ms-word.document.macroEnabled.12"
    : extension === "dotx"
      ? "application/vnd.openxmlformats-officedocument.wordprocessingml.template"
      : extension === "dotm"
        ? "application/vnd.ms-word.template.macroEnabled.12"
        : "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  const url = URL.createObjectURL(new Blob([bytes], {
    type: mimeType,
  }))
  const anchor = document.createElement("a")
  anchor.href = url
  anchor.download = displayFileName.value || "document.docx"
  anchor.click()
  queueMicrotask(() => URL.revokeObjectURL(url))
}

function toggleTheme(): void {
  darkTheme.value = !darkTheme.value
  emit("update:isDark", darkTheme.value)
}

function toggleTrackedChanges(): void {
  const visible = !effectiveShowTrackedChanges.value
  if (props.showTrackedChanges == null) internalShowTrackedChanges.value = visible
  emit("update:showTrackedChanges", visible)
}

function toggleComments(): void {
  const visible = !effectiveShowComments.value
  if (props.showComments == null) internalShowComments.value = visible
  emit("update:showComments", visible)
}

function normalizeZoom(value: number): number {
  const options = [50, 75, 100, 125, 150, 175, 200]
  return options.reduce((closest, option) =>
    Math.abs(option - value) < Math.abs(closest - value) ? option : closest,
  100)
}

onBeforeUnmount(() => {
  runtimeLoader.dispose()
  ownedRuntime.dispose()
})

function errorCode(value: Error): string | undefined {
  const code = (value as Error & { code?: unknown }).code
  return typeof code === "string" ? code : undefined
}

function errorDetail(value: Error, key: "phase" | "limit"): string | undefined {
  const detail = (value as Error & Record<typeof key, unknown>)[key]
  return typeof detail === "string" ? detail : undefined
}

function errorNumber(value: Error, key: "actual" | "allowed"): number | undefined {
  const detail = (value as Error & Record<typeof key, unknown>)[key]
  return typeof detail === "number" ? detail : undefined
}
</script>

<style scoped>
.docx-viewer {
  --docx-border: #e4e4e7;
  --docx-foreground: #18181b;
  --docx-toolbar-bg: #fff;
  background: #f4f4f5;
  border: 1px solid var(--docx-border);
  border-radius: 10px;
  color: var(--docx-foreground);
  display: flex;
  flex-direction: column;
  height: 100%;
  min-height: 320px;
  min-width: 0;
  overflow: hidden;
  width: 100%;
}
.docx-viewer--dark {
  --docx-border: #3f3f46;
  --docx-foreground: #f4f4f5;
  --docx-toolbar-bg: #18181b;
  background: #09090b;
}
.docx-viewer__file-input { display: none; }
.docx-viewer__body { display: flex; flex: 1; min-height: 0; min-width: 0; }
.docx-viewer__content { display: flex; flex: 1; min-height: 0; min-width: 0; position: relative; }
.docx-viewer-state {
  align-items: center;
  color: #71717a;
  display: flex;
  flex: 1;
  gap: 14px;
  justify-content: center;
  min-height: 260px;
  padding: 32px;
  text-align: left;
}
.docx-viewer-state > div { display: flex; flex-direction: column; gap: 4px; max-width: 430px; }
.docx-viewer-state strong { color: var(--docx-foreground); font-size: 14px; }
.docx-viewer-state span { font-size: 13px; line-height: 1.5; }
.docx-viewer-state svg {
  fill: none;
  height: 32px;
  stroke: currentColor;
  stroke-linecap: round;
  stroke-linejoin: round;
  stroke-width: 1.5;
  width: 32px;
}
.docx-viewer-state button {
  align-self: flex-start;
  background: #fff;
  border: 1px solid #d4d4d8;
  border-radius: 7px;
  color: #18181b;
  cursor: pointer;
  font: inherit;
  font-size: 12px;
  margin-top: 8px;
  min-height: 32px;
  padding: 0 12px;
}
.docx-viewer-state--error { color: #b91c1c; }
.docx-viewer-error-code { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-weight: 600; }
.docx-viewer-loading-spinner {
  animation: docx-viewer-spin .7s linear infinite;
  border: 2px solid #d4d4d8;
  border-radius: 50%;
  border-top-color: #2563eb;
  height: 24px;
  width: 24px;
}
@keyframes docx-viewer-spin { to { transform: rotate(360deg); } }
</style>
