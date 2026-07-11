<template>
  <div class="page" data-testid="pdf-viewer-page">
    <h2>📑 PDF Viewer — 真实渲染页</h2>
    <p class="desc">PDFium Worker 从已取得的文档字节渲染真实页面，支持缩略图、翻页、缩放、旋转、全文搜索和下载。</p>

    <div class="controls control-panel">
      <label>
        示例 PDF
        <select v-model="selectedSample" data-testid="pdf-sample-select" @change="loadSelectedSample">
          <option v-for="sample in samples" :key="sample.file" :value="sample.file">
            {{ sample.label }}
          </option>
        </select>
      </label>
      <button data-testid="pdf-load-sample" @click="loadSelectedSample">加载示例</button>
      <label>
        本地 PDF
        <input data-testid="pdf-file-input" type="file" accept=".pdf,application/pdf" @change="onFileChange" />
      </label>
    </div>

    <div class="policy-panel" data-testid="pdf-runtime-config">
      <h3>公开来源规则</h3>
      <label>
        PDF 地址
        <input v-model="urlDraft" data-testid="pdf-url-input" inputmode="url" />
      </label>
      <label>
        baseUrl
        <input v-model="baseUrl" data-testid="pdf-policy-base-url" inputmode="url" />
      </label>
      <label>
        允许协议（逗号分隔）
        <input v-model="allowedProtocolsText" data-testid="pdf-allowed-protocols" />
      </label>
      <label>
        允许来源（逗号分隔）
        <input v-model="allowedOriginsText" data-testid="pdf-allowed-origins" />
      </label>
      <label class="checkbox-label">
        <input v-model="allowHttpOnLocalhost" data-testid="pdf-allow-localhost-http" type="checkbox" />
        允许 localhost 的 HTTP
      </label>
      <label>
        maxFileSize（字节）
        <input
          v-model="maxFileSizeDraft"
          data-testid="pdf-max-file-size"
          type="number"
          min="1"
          step="1"
          inputmode="numeric"
        />
      </label>
      <button data-testid="pdf-apply-max-file-size" @click="applyMaxFileSize">应用文件上限</button>
      <button data-testid="pdf-apply-url" @click="applyUrl">应用地址和规则</button>
      <span
        data-testid="pdf-max-file-size-active"
        :data-bytes="maxFileSize"
      >当前上限：{{ formatFileSize(maxFileSize) }}</span>
      <span v-if="maxFileSizeInputError" class="error" data-testid="pdf-max-file-size-error">
        {{ maxFileSizeInputError }}
      </span>
    </div>

    <div class="status-grid info-grid">
      <div><strong>状态：</strong><span data-testid="page-status" :data-state="pageState">{{ pageState }}</span></div>
      <div><strong>已加载：</strong><span data-testid="loaded-file">{{ fileName || "无" }}</span></div>
      <div><strong>来源：</strong>{{ sourceKind }}</div>
      <div><strong>页面：</strong>{{ numPages || "等待加载" }}；当前 {{ activePage }}</div>
    </div>

    <div v-if="pageError" class="error" data-testid="pdf-page-load-error" :data-error-code="pageError.code">
      {{ pageError.code }}: {{ pageError.message }}
    </div>

    <div class="viewer-container">
      <PdfViewer
        v-if="source"
        :source="source"
        :url-policy="urlPolicy"
        :max-file-size="maxFileSize"
        :default-zoom="1"
        :file-name="fileName || 'verification.pdf'"
        style="height: 80vh;"
        @document-load-success="onLoadSuccess"
        @document-load-error="onLoadError"
        @active-page-change="activePage = $event"
        @diagnostic="onDiagnostic"
      />
      <div v-else class="empty">
        <p>请选择 PDF 文件或输入地址。</p>
      </div>
    </div>

    <div class="diagnostics" data-testid="pdf-diagnostics" aria-live="polite">
      <h3>公开诊断</h3>
      <ol>
        <li
          v-for="entry in diagnostics"
          :key="entry.id"
          data-testid="pdf-diagnostic"
          :data-diagnostic-type="entry.type"
          :data-request-id="entry.requestId"
          :data-runtime-id="entry.runtimeId"
          :data-error-code="entry.code"
        >
          #{{ entry.requestId }} {{ entry.type }}<span v-if="entry.taskId"> · {{ entry.taskId }}</span><span v-if="entry.code"> · {{ entry.code }}</span><span v-if="entry.bytes !== undefined"> · {{ entry.bytes }} bytes</span>
        </li>
      </ol>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, shallowRef } from "vue"
import { PdfViewer, type PdfDiagnostic, type PdfLoadError, type PdfSource, type PdfUrlPolicy } from "@arcships/vue-pdf"

const samples = [
  { file: "sample.pdf", label: "文本 PDF" },
  { file: "scanned-invoice.pdf", label: "扫描发票" },
  { file: "rotated-pages.pdf", label: "旋转页面" },
  { file: "large-contract.pdf", label: "长合同" },
  { file: "corrupted.pdf", label: "损坏文件（失败用例）" },
]

const currentOrigin = globalThis.location.origin
const DEFAULT_MAX_FILE_SIZE = 50 * 1024 * 1024
const selectedSample = ref(samples[0].file)
const source = shallowRef<PdfSource | undefined>()
const fileName = ref("")
const sourceKind = ref("未选择")
const urlDraft = ref(`/samples/${selectedSample.value}`)
const baseUrl = ref(globalThis.location.href)
const allowedProtocolsText = ref("https:")
const allowedOriginsText = ref(currentOrigin)
const allowHttpOnLocalhost = ref(true)
const maxFileSize = ref(DEFAULT_MAX_FILE_SIZE)
const maxFileSizeDraft = ref(String(DEFAULT_MAX_FILE_SIZE))
const maxFileSizeInputError = ref("")
const pageState = ref<"idle" | "loading" | "ready" | "error">("idle")
const pageError = ref<PdfLoadError | null>(null)
const numPages = ref(0)
const activePage = ref(1)
const diagnostics = ref<Array<PdfDiagnostic & { id: number; code?: string }>>([])
let diagnosticId = 0

function splitList(value: string): string[] {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean)
}

const urlPolicy = computed<PdfUrlPolicy>(() => ({
  baseUrl: baseUrl.value.trim(),
  allowRelativeUrl: true,
  allowedProtocols: splitList(allowedProtocolsText.value),
  allowedOrigins: splitList(allowedOriginsText.value),
  allowHttpOnLocalhost: allowHttpOnLocalhost.value,
}))

function formatFileSize(bytes: number): string {
  return `${(bytes / (1024 * 1024)).toFixed(2)} MiB（${bytes} 字节）`
}

function applyMaxFileSize() {
  const next = Number(maxFileSizeDraft.value)
  if (!Number.isSafeInteger(next) || next <= 0) {
    maxFileSizeInputError.value = "文件上限必须是大于 0 的整数。"
    return
  }
  maxFileSizeInputError.value = ""
  maxFileSize.value = next
}

function resetLoadState(nextKind: string, nextFileName: string) {
  sourceKind.value = nextKind
  fileName.value = nextFileName
  pageError.value = null
  pageState.value = "loading"
  numPages.value = 0
  activePage.value = 1
}

function loadSelectedSample() {
  const file = selectedSample.value
  urlDraft.value = `/samples/${file}`
  resetLoadState("示例地址", file)
  source.value = { kind: "url", url: urlDraft.value }
}

function applyUrl() {
  const url = urlDraft.value.trim()
  resetLoadState("地址", url || "未命名 PDF")
  source.value = { kind: "url", url }
}

function onFileChange(event: Event) {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  if (!file) return
  resetLoadState("本地文件", file.name)
  source.value = { kind: "blob", blob: file, fileName: file.name }
}

function onLoadSuccess(pages: number) {
  numPages.value = pages
  pageError.value = null
  pageState.value = "ready"
}

function onLoadError(error: PdfLoadError) {
  pageError.value = error
  pageState.value = "error"
}

function onDiagnostic(event: PdfDiagnostic) {
  diagnostics.value = [
    ...diagnostics.value.slice(-9),
    { ...event, id: ++diagnosticId, code: event.error?.code },
  ]
  if (event.type === "load-start") pageState.value = "loading"
  if (event.type === "load-success") pageState.value = "ready"
  if (event.type === "load-error") pageState.value = "error"
}

loadSelectedSample()
</script>

<style scoped>
.page { padding: 24px; max-width: 1200px; margin: 0 auto; width: 100%; min-width: 0; }
h2 { margin-bottom: 4px; }
.desc { color: var(--muted-foreground); margin-bottom: 16px; font-size: 14px; }
.controls, .policy-panel { display: flex; gap: 12px; align-items: end; flex-wrap: wrap; margin-bottom: 12px; }
.controls label, .policy-panel label { display: flex; flex-direction: column; gap: 6px; align-items: flex-start; font-size: 13px; color: var(--muted-foreground); }
.controls input, .controls select, .controls button, .policy-panel input, .policy-panel button { min-height: 36px; padding: 8px 12px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--background); }
.controls button, .policy-panel button { cursor: pointer; }
.policy-panel { align-items: end; padding: 12px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--muted); }
.policy-panel h3 { flex-basis: 100%; margin: 0; font-size: 14px; }
.policy-panel input:not([type="checkbox"]) { min-width: 180px; }
.policy-panel .checkbox-label { flex-direction: row; align-items: center; min-height: 36px; }
.policy-panel .checkbox-label input { min-height: auto; }
.error { color: #dc2626; font-size: 13px; margin: 0 0 12px; }
.status-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 8px; padding: 12px; margin-bottom: 16px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--muted); font-size: 13px; }
.viewer-container { border: 1px solid var(--border); border-radius: var(--radius); overflow: auto; margin-bottom: 24px; }
.empty { display: flex; align-items: center; justify-content: center; height: 200px; color: var(--muted-foreground); }
.diagnostics { margin-top: 24px; border-top: 1px solid var(--border); padding-top: 12px; font-size: 13px; }
.diagnostics h3 { margin: 0 0 8px; }
.diagnostics ol { margin: 0; padding-left: 20px; }
</style>
