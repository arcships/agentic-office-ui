<template>
  <div class="page">
    <h2>📄 DOCX Viewer — Verification</h2>
    <p class="desc">
      使用公开 Runtime 配置验证真实 DOCX Worker、WASM、错误和恢复路径。
    </p>

    <div class="controls control-panel">
      <label>
        Sample DOCX
        <select v-model="selectedSample" data-testid="docx-sample-select" @change="loadSelectedSample">
          <option v-for="sample in samples" :key="sample.file" :value="sample.file">
            {{ sample.label }}
          </option>
        </select>
      </label>
      <button data-testid="docx-load-sample" @click="loadSelectedSample">Load sample</button>
      <input data-testid="docx-file-input" type="file" accept=".docx" @change="onFileChange" />
    </div>

    <fieldset class="runtime-config" data-testid="docx-runtime-config">
      <legend>DOCX Runtime configuration</legend>
      <label>
        WASM URL
        <input data-testid="docx-wasm-url" v-model="wasmUrl" type="url" />
      </label>
      <label>
        Worker URL（留空使用包内 Worker）
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
      <button data-testid="apply-docx-runtime" type="button" @click="applyRuntimeConfig">
        Apply runtime configuration
      </button>
    </fieldset>

    <div class="status-grid info-grid">
      <div data-testid="page-status" :data-state="pageState">
        <strong>Status:</strong> {{ pageState }}
      </div>
      <div data-testid="loaded-file"><strong>Loaded:</strong> {{ displayName || "None" }}</div>
      <div data-testid="docx-import-source"><strong>Source:</strong> {{ sourceKind }}</div>
      <div><strong>Expected coverage:</strong> Worker, WASM, paragraphs, tables, images and recovery</div>
    </div>

    <p v-if="error" class="error" data-testid="docx-page-error">
      {{ errorCode ? `${errorCode}: ` : "" }}{{ error }}
    </p>

    <div class="viewer-container">
      <DocxViewer
        v-if="fileBuffer"
        :key="viewerKey"
        :file="fileBuffer"
        :runtime="runtime"
        :layout-options="{ pageWidth: 816, pageHeight: 1056 }"
        style="height: 80vh; overflow: auto;"
        @load-start="onViewerLoadStart"
        @load-success="onViewerLoadSuccess"
        @load-error="onViewerLoadError"
      />
      <div v-else class="empty">
        <p>Load a DOCX sample or upload a .docx file to begin verification.</p>
      </div>
    </div>

    <details class="diagnostics" open>
      <summary>Runtime diagnostics</summary>
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
} from "@extend-ai/docx-core"
import { DocxViewer } from "@extend-ai/vue-docx"

const samples = [
  { file: "demo.docx", label: "Demo — master services agreement" },
  { file: "legal-contract.docx", label: "Legal contract — multi-page" },
  { file: "invoice-table.docx", label: "Invoice table — merged cells and totals" },
  { file: "report-with-image.docx", label: "Report with image and rich text" },
  { file: "chinese-mixed.docx", label: "Chinese + English mixed text" },
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
const viewerKey = computed(() => `${displayName.value}-${loadCounter.value}`)

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
  if (!file.name.toLowerCase().endsWith(".docx")) {
    if (isCurrentPageLoad(request)) {
      activePageLoad = undefined
      setError("Please select a .docx file", "PARSE_FAILED")
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
.page { padding: 24px; max-width: 1200px; margin: 0 auto; width: 100%; min-width: 0; }
h2 { margin-bottom: 4px; }
.desc { color: var(--muted-foreground); margin-bottom: 16px; font-size: 14px; }
.controls, .runtime-config { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin-bottom: 12px; }
.controls label, .runtime-config label { display: flex; flex-direction: column; gap: 6px; align-items: flex-start; font-size: 13px; color: var(--muted-foreground); }
.controls button, .controls input, .controls select, .runtime-config button, .runtime-config input { padding: 8px 12px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--background); }
.runtime-config { padding: 12px; border: 1px solid var(--border); border-radius: var(--radius); }
.runtime-config legend { padding: 0 6px; font-size: 13px; color: var(--muted-foreground); }
.runtime-config input[type="url"] { min-width: 280px; }
.runtime-config .checkbox-label { flex-direction: row; align-items: center; }
.runtime-config .checkbox-label input { padding: 0; }
.error { color: #ef4444; font-size: 13px; }
.status-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 8px; padding: 12px; margin-bottom: 16px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--muted); font-size: 13px; }
.viewer-container { border: 1px solid var(--border); border-radius: var(--radius); overflow: auto; margin-bottom: 24px; }
.empty { display: flex; align-items: center; justify-content: center; height: 200px; color: var(--muted-foreground); }
.diagnostics { margin-top: 24px; font-size: 13px; }
.diagnostics pre { max-height: 280px; overflow: auto; padding: 12px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--muted); white-space: pre-wrap; }
</style>
