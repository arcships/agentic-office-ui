<template>
  <div class="page" data-testid="docx-viewer-page">
    <header class="product-header">
      <div>
        <h2>Word 文档查看器</h2>
        <p class="desc">打开文档后，可翻页、缩放、查看缩略图和下载原文件。</p>
      </div>
      <div class="controls control-panel product-actions">
        <label>
          示例文档
          <select v-model="selectedSample" data-testid="docx-sample-select" @change="loadSelectedSample">
            <option v-for="sample in samples" :key="sample.file" :value="sample.file">
              {{ sample.label }}
            </option>
          </select>
        </label>
        <button data-testid="docx-load-sample" @click="loadSelectedSample">打开示例</button>
        <label class="checkbox-action">
          <input v-model="showTrackedChanges" data-testid="docx-demo-show-tracked-changes" type="checkbox" />
          显示修订
        </label>
        <label class="checkbox-action">
          <input v-model="showComments" data-testid="docx-demo-show-comments" type="checkbox" />
          显示批注
        </label>
        <label class="file-action">
          打开本地文件
          <input data-testid="docx-file-input" type="file" accept=".docx,.docm,.dotx,.dotm" @change="onFileChange" />
        </label>
      </div>
    </header>

    <p v-if="error" class="error" data-testid="docx-page-error">
      {{ errorCode ? `${errorCode}: ` : "" }}{{ error }}
    </p>

    <fieldset class="props-panel" data-testid="docx-props-panel">
      <legend>组件 Props</legend>
      <label class="checkbox-label">
        <input v-model="showToolbar" data-testid="docx-prop-show-toolbar" type="checkbox" />
        showToolbar
      </label>
      <label class="checkbox-label">
        <input v-model="showUpload" data-testid="docx-prop-show-upload" type="checkbox" />
        showUpload
      </label>
      <label class="checkbox-label">
        <input v-model="showDownload" data-testid="docx-prop-show-download" type="checkbox" />
        showDownload
      </label>
      <label class="checkbox-label">
        <input v-model="defaultThumbnailsOpen" data-testid="docx-prop-default-thumbnails-open" type="checkbox" />
        defaultThumbnailsOpen
      </label>
      <label class="checkbox-label">
        <input v-model="isDark" data-testid="docx-prop-is-dark" type="checkbox" />
        isDark
      </label>
      <label class="select-label">
        defaultZoom
        <select v-model.number="defaultZoom" data-testid="docx-prop-default-zoom">
          <option v-for="zoom in zoomOptions" :key="zoom" :value="zoom">{{ zoom }}%</option>
        </select>
      </label>
    </fieldset>

    <div class="viewer-container product-surface">
      <DocxViewer
        v-if="fileBuffer"
        :key="viewerKey"
        :file="fileBuffer"
        :file-name="displayName"
        :runtime="runtime"
        :show-toolbar="showToolbar"
        :show-upload="showUpload"
        :show-download="showDownload"
        :default-thumbnails-open="defaultThumbnailsOpen"
        :default-zoom="defaultZoom"
        v-model:is-dark="isDark"
        v-model:show-tracked-changes="showTrackedChanges"
        v-model:show-comments="showComments"
        :layout-options="{ pageWidth: 816, pageHeight: 1056 }"
        style="height: 78vh;"
        @load-start="onViewerLoadStart"
        @load-success="onViewerLoadSuccess"
        @load-error="onViewerLoadError"
      />
      <div v-else class="empty">
        <p>选择示例或打开本地 Word 文档、模板文件。</p>
      </div>
    </div>

    <div class="status-grid info-grid verification-section">
      <div data-testid="page-status" :data-state="pageState"><strong>状态：</strong>{{ pageState }}</div>
      <div data-testid="loaded-file"><strong>文件：</strong>{{ displayName || "未打开" }}</div>
      <div data-testid="docx-import-source"><strong>执行位置：</strong>{{ sourceKind }}</div>
    </div>

    <details class="runtime-details" data-testid="docx-runtime-details">
      <summary>运行配置与诊断</summary>
      <fieldset class="runtime-config" data-testid="docx-runtime-config">
        <legend>DOCX 运行配置</legend>
        <label>
          WASM 地址
          <input data-testid="docx-wasm-url" v-model="wasmUrl" type="url" />
        </label>
        <label>
          Worker 地址（留空使用包内 Worker）
          <input data-testid="docx-worker-url" v-model="workerUrl" type="url" />
        </label>
        <label class="checkbox-label">
          <input data-testid="docx-allow-main-thread-fallback" v-model="allowMainThreadFallback" type="checkbox" />
          允许受字节上限约束的主线程回退
        </label>
        <label>
          主线程回退上限（字节）
          <input data-testid="docx-main-thread-fallback-max-bytes" v-model.number="mainThreadFallbackMaxBytes" min="0" type="number" />
        </label>
        <button data-testid="apply-docx-runtime" type="button" @click="applyRuntimeConfig">应用运行配置</button>
      </fieldset>
      <pre data-testid="docx-diagnostics">{{ JSON.stringify(diagnostics, null, 2) }}</pre>
    </details>
  </div>
</template>

<script setup lang="ts">
import { computed, onBeforeUnmount, ref, shallowRef } from "vue"
import {
  bundledDocxWasmUrl,
  createDocxRuntime,
  type DocxImportDiagnostic,
  type DocxImportResult,
} from "@arcships/docx-core"
import { DocxViewer } from "@arcships/vue-docx"

const samples = [
  { file: "demo.docx", label: "Demo — master services agreement" },
  { file: "legal-contract.docx", label: "Legal contract — multi-page" },
  { file: "invoice-table.docx", label: "Invoice table — merged cells and totals" },
  { file: "report-with-image.docx", label: "Report with image and rich text" },
  { file: "chinese-mixed.docx", label: "Chinese + English mixed text" },
  { file: "review-comments.docx", label: "Review — 修订、批注、脚注、尾注与搜索" },
  { file: "corrupted.docx", label: "Corrupted DOCX — negative test" },
]

type PageState = "idle" | "loading" | "ready" | "error"
type PageLoadSource = "sample" | "file"
type PageLoadDiagnostic = {
  type: "page-load-start" | "page-load-success" | "page-load-error" | "page-load-cancelled"
  requestId: string
  source: PageLoadSource
  fileName: string
}
type VisibleDiagnostic = DocxImportDiagnostic | PageLoadDiagnostic

const selectedSample = ref(samples[0].file)
const fileBuffer = ref<ArrayBuffer | null>(null)
const error = ref<string | null>(null)
const errorCode = ref<string | null>(null)
const displayName = ref("")
const sourceKind = ref("pending")
const loadCounter = ref(0)
const pageState = ref<PageState>("idle")
const wasmUrl = ref(bundledDocxWasmUrl)
const workerUrl = ref("")
const allowMainThreadFallback = ref(false)
const mainThreadFallbackMaxBytes = ref(0)
const showTrackedChanges = ref(true)
const showComments = ref(true)
const showToolbar = ref(true)
const showUpload = ref(true)
const showDownload = ref(true)
const defaultThumbnailsOpen = ref(false)
const isDark = ref(false)
const defaultZoom = ref(100)
const zoomOptions = [50, 75, 100, 125, 150, 175, 200]
const diagnostics = ref<VisibleDiagnostic[]>([])
let pageLoadSequence = 0
let activePageLoad:
  | {
      id: number
      source: PageLoadSource
      fileName: string
      fetchController?: AbortController
      reader?: FileReader
    }
  | undefined

function recordDiagnostic(event: DocxImportDiagnostic): void {
  diagnostics.value = [...diagnostics.value, event].slice(-50)
}

function recordPageLoadDiagnostic(
  type: PageLoadDiagnostic["type"],
  request: NonNullable<typeof activePageLoad>,
): void {
  diagnostics.value = [
    ...diagnostics.value,
    {
      type,
      requestId: `page:${request.id}`,
      source: request.source,
      fileName: request.fileName,
    },
  ].slice(-50)
}

function abandonActivePageLoad(): void {
  const request = activePageLoad
  if (!request) return
  activePageLoad = undefined
  request.fetchController?.abort()
  if (request.reader?.readyState === FileReader.LOADING) {
    request.reader.abort()
  }
  recordPageLoadDiagnostic("page-load-cancelled", request)
}

function beginPageLoad(source: PageLoadSource, fileName: string) {
  abandonActivePageLoad()
  const request = {
    id: ++pageLoadSequence,
    source,
    fileName,
  }
  activePageLoad = request
  pageState.value = "loading"
  error.value = null
  errorCode.value = null
  recordPageLoadDiagnostic("page-load-start", request)
  return request
}

function isCurrentPageLoad(request: NonNullable<typeof activePageLoad>): boolean {
  return activePageLoad?.id === request.id
}

function createPageRuntime() {
  return createDocxRuntime({
    wasmUrl: wasmUrl.value.trim() || bundledDocxWasmUrl,
    workerUrl: workerUrl.value.trim() || undefined,
    allowMainThreadFallback: allowMainThreadFallback.value,
    mainThreadFallbackMaxBytes: Math.max(0, mainThreadFallbackMaxBytes.value || 0),
    onDiagnostic: recordDiagnostic,
  })
}

const runtime = shallowRef(createPageRuntime())
const viewerKey = computed(() => `${displayName.value}-${loadCounter.value}-${String(defaultThumbnailsOpen.value)}`)

function setError(message: string, code: string): void {
  error.value = message
  errorCode.value = code
  pageState.value = "error"
}

function setBuffer(buffer: ArrayBuffer, name: string): void {
  fileBuffer.value = buffer
  displayName.value = name
  sourceKind.value = "pending"
  error.value = null
  errorCode.value = null
  pageState.value = "loading"
  loadCounter.value++
}

function onFileChange(event: Event): void {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  if (!file) return
  const request = beginPageLoad("file", file.name)
  if (!/\.(?:docx|docm|dotx|dotm)$/i.test(file.name)) {
    if (isCurrentPageLoad(request)) {
      activePageLoad = undefined
      setError("Please select a supported Word document", "PARSE_FAILED")
      recordPageLoadDiagnostic("page-load-error", request)
    }
    return
  }
  const reader = new FileReader()
  activePageLoad = { ...request, reader }
  reader.onload = () => {
    if (!isCurrentPageLoad(request)) return
    activePageLoad = undefined
    setBuffer(reader.result as ArrayBuffer, file.name)
    recordPageLoadDiagnostic("page-load-success", request)
  }
  reader.onerror = () => {
    if (!isCurrentPageLoad(request)) return
    activePageLoad = undefined
    setError("Failed to read file", "FETCH_FAILED")
    recordPageLoadDiagnostic("page-load-error", request)
  }
  reader.readAsArrayBuffer(file)
}

async function loadSelectedSample(): Promise<void> {
  const fileName = selectedSample.value
  const request = beginPageLoad("sample", fileName)
  const fetchController = new AbortController()
  activePageLoad = { ...request, fetchController }
  try {
    const response = await fetch(`/samples/${encodeURIComponent(fileName)}`, { signal: fetchController.signal })
    if (!isCurrentPageLoad(request)) return
    if (!response.ok) {
      activePageLoad = undefined
      setError(`Sample file not found: ${fileName}`, "FETCH_FAILED")
      recordPageLoadDiagnostic("page-load-error", request)
      return
    }
    const buffer = await response.arrayBuffer()
    if (!isCurrentPageLoad(request)) return
    activePageLoad = undefined
    setBuffer(buffer, fileName)
    recordPageLoadDiagnostic("page-load-success", request)
  } catch (cause) {
    if (!isCurrentPageLoad(request) || fetchController.signal.aborted) return
    activePageLoad = undefined
    setError(`Failed to load sample: ${String(cause)}`, "FETCH_FAILED")
    recordPageLoadDiagnostic("page-load-error", request)
  }
}

function applyRuntimeConfig(): void {
  abandonActivePageLoad()
  runtime.value.dispose()
  diagnostics.value = []
  runtime.value = createPageRuntime()
  void loadSelectedSample()
}

function onViewerLoadStart(): void {
  pageState.value = "loading"
}

function onViewerLoadSuccess(result: DocxImportResult): void {
  sourceKind.value = result.source
  error.value = null
  errorCode.value = null
  pageState.value = "ready"
}

function onViewerLoadError(cause: Error): void {
  const code = (cause as Error & { code?: unknown }).code
  setError(cause.message, typeof code === "string" ? code : "PARSE_FAILED")
}

onBeforeUnmount(() => {
  abandonActivePageLoad()
  runtime.value.dispose()
})

void loadSelectedSample()
</script>

<style scoped>
.page { padding: 16px; max-width: 1440px; margin: 0 auto; width: 100%; min-width: 0; }
h2 { margin-bottom: 4px; }
.desc { color: var(--muted-foreground); margin-bottom: 16px; font-size: 14px; }
.product-header { display: flex; align-items: end; justify-content: space-between; gap: 16px; flex-wrap: wrap; margin-bottom: 12px; }
.product-header .desc { margin-bottom: 0; }
.product-actions { margin-bottom: 0; }
.checkbox-action { align-items: center !important; flex-direction: row !important; white-space: nowrap; }
.checkbox-action input { padding: 0; }
.product-surface { box-shadow: 0 10px 30px rgb(15 23 42 / 8%); }
.verification-section { margin-top: 8px; }
.runtime-details { margin-top: 12px; border: 1px solid var(--border); border-radius: var(--radius); padding: 10px 12px; color: var(--muted-foreground); }
.runtime-details summary { cursor: pointer; font-size: 13px; font-weight: 600; color: var(--foreground); }
.controls, .runtime-config { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin-bottom: 12px; }
.controls label, .runtime-config label { display: flex; flex-direction: column; gap: 6px; align-items: flex-start; font-size: 13px; color: var(--muted-foreground); }
.controls button, .controls input, .controls select, .runtime-config button, .runtime-config input { padding: 8px 12px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--background); }
.runtime-config { padding: 12px; border: 1px solid var(--border); border-radius: var(--radius); }
.runtime-config legend { padding: 0 6px; font-size: 13px; color: var(--muted-foreground); }
.runtime-config input[type="url"] { min-width: 280px; }
.runtime-config .checkbox-label { flex-direction: row; align-items: center; }
.runtime-config .checkbox-label input { padding: 0; }
.props-panel { display: flex; gap: 18px; align-items: center; flex-wrap: wrap; padding: 12px 16px; margin-bottom: 16px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--muted); }
.props-panel legend { padding: 0 6px; font-size: 13px; font-weight: 600; color: var(--muted-foreground); }
.props-panel .checkbox-label { display: flex; flex-direction: row; align-items: center; gap: 6px; font-size: 13px; color: var(--foreground); white-space: nowrap; }
.props-panel .checkbox-label input { padding: 0; }
.props-panel .select-label { display: flex; flex-direction: row; align-items: center; gap: 8px; font-size: 13px; color: var(--foreground); }
.props-panel select { padding: 6px 10px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--background); }
.error { color: #ef4444; font-size: 13px; }
.status-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 8px; padding: 12px; margin-bottom: 16px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--muted); font-size: 13px; }
.viewer-container { border: 1px solid var(--border); border-radius: var(--radius); overflow: auto; margin-bottom: 24px; }
.empty { display: flex; align-items: center; justify-content: center; height: 200px; color: var(--muted-foreground); }
.diagnostics { margin-top: 24px; font-size: 13px; }
.diagnostics pre { max-height: 280px; overflow: auto; padding: 12px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--muted); white-space: pre-wrap; }
@media (max-width: 760px) {
  .page { padding: 10px; }
  .product-header { align-items: stretch; }
  .runtime-config input[type="url"] { min-width: min(280px, 76vw); }
}
</style>
