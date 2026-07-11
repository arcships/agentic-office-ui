<template>
  <div class="page" data-testid="xlsx-viewer-page">
    <header class="product-header">
      <div>
        <h2>XLSX 工作簿</h2>
        <p class="desc">查看和编辑单元格、公式、样式、图表与图片。</p>
      </div>
      <div class="controls control-panel product-actions">
      <label>
        示例工作簿
        <select v-model="selectedSample" data-testid="xlsx-sample-select" @change="loadSelectedSample">
          <option v-for="sample in samples" :key="sample.file" :value="sample.file">{{ sample.label }}</option>
        </select>
      </label>
      <button data-testid="xlsx-load-sample" @click="loadSelectedSample">加载示例</button>
      <label class="inline"><input v-model="readOnly" data-testid="xlsx-read-only" type="checkbox" /> 只读</label>
      </div>
    </header>
    <input ref="fileInputRef" class="visually-hidden" data-testid="xlsx-file-input" type="file" accept=".xlsx,.xls" @change="onFileChange" />

    <div v-if="pageError" class="error" data-testid="load-error" :data-error-code="pageError.code">
      {{ pageError.code }}: {{ pageError.message }}
    </div>

    <div
      class="viewer-container product-surface"
      :class="{ 'viewer-container--drag-active': isDragActive }"
      @dragenter.prevent="onDragEnter"
      @dragover.prevent="onDragOver"
      @dragleave.prevent="onDragLeave"
      @drop.prevent="onDrop"
    >
      <div v-if="isDragActive" class="drag-overlay"><span>放下 .xlsx 文件</span></div>
      <XlsxViewerWrapper
        v-if="viewerKey"
        :key="viewerKey"
        :file="fileBuffer ?? undefined"
        :src="src ?? undefined"
        :file-name="displayName || 'verification.xlsx'"
        :read-only="readOnly"
        :use-worker="useWorker"
        :url-policy="urlPolicy"
        :defer-loading-above-bytes="deferLoadingAboveBytes"
        style="height: 76vh;"
        @controller-ready="onControllerReady"
        @diagnostic="onDiagnostic"
        @upload="openFilePicker"
        @update:read-only="onReadOnlyChange"
      />
      <div v-else class="empty"><p>请选择工作簿、拖入本地文件或输入地址。</p></div>
    </div>

    <div class="status-grid info-grid verification-section">
      <div><strong>状态：</strong><span data-testid="page-status" :data-state="pageState">{{ pageState }}</span></div>
      <div><strong>文件：</strong><span data-testid="loaded-file">{{ displayName || "未打开" }}</span></div>
      <div><strong>来源：</strong>{{ sourceKind }}</div>
      <div><strong>模式：</strong>{{ readOnly ? "只读" : "可编辑" }}</div>
      <div><strong>请求 Worker：</strong><span data-testid="xlsx-worker-requested">{{ useWorker ? "开启" : "关闭" }}</span></div>
      <div><strong>实际执行：</strong><span data-testid="xlsx-worker-actual">{{ actualWorkerMode }}</span></div>
    </div>

    <details class="runtime-details" data-testid="xlsx-runtime-details">
      <summary>地址规则与运行诊断</summary>
      <div class="policy-panel" data-testid="xlsx-runtime-config">
        <label class="checkbox-label"><input v-model="useWorker" data-testid="xlsx-use-worker" type="checkbox" />使用 Worker</label>
        <label>XLSX 地址 <input v-model="remoteUrl" data-testid="xlsx-url-input" inputmode="url" @keydown.enter="loadUrl" /></label>
        <label>baseUrl <input v-model="baseUrl" data-testid="xlsx-policy-base-url" inputmode="url" /></label>
        <label>允许协议 <input v-model="allowedProtocolsText" data-testid="xlsx-allowed-protocols" /></label>
        <label>允许来源 <input v-model="allowedOriginsText" data-testid="xlsx-allowed-origins" /></label>
        <label class="checkbox-label"><input v-model="allowHttpOnLocalhost" data-testid="xlsx-allow-localhost-http" type="checkbox" />允许 localhost 的 HTTP</label>
        <label>延迟解析阈值（字节）<input v-model.number="deferLoadingAboveBytes" data-testid="xlsx-defer-loading-above-bytes" min="0" type="number" /></label>
        <button data-testid="xlsx-apply-url" :disabled="!remoteUrl.trim()" @click="loadUrl">应用地址和规则</button>
        <button data-testid="xlsx-continue-deferred" :disabled="!controller?.canLoadDeferred" @click="continueDeferredLoad">继续解析</button>
        <button data-testid="xlsx-download-source" :disabled="!controller?.canDownload" @click="downloadSource">下载源文件</button>
      </div>
      <div class="diagnostics" data-testid="xlsx-diagnostics" aria-live="polite">
        <ol>
          <li v-for="entry in diagnostics" :key="entry.id" :data-diagnostic-type="entry.type">
            #{{ entry.requestId }} {{ entry.type }}<span v-if="entry.taskId"> · {{ entry.taskId }}</span><span v-if="entry.error?.code"> · {{ entry.error.code }}</span><span v-if="entry.bytes !== undefined"> · {{ entry.bytes }} bytes</span>
          </li>
        </ol>
      </div>
    </details>
  </div>
</template>

<script setup lang="ts">
import { computed, defineComponent, h, onBeforeUnmount, onMounted, ref, shallowRef, type PropType } from "vue"
import {
  useXlsxViewerController,
  XlsxViewer,
  type XlsxDiagnostic,
  type XlsxLoadError,
  type XlsxUrlPolicy,
} from "@arcships/vue-xlsx"
import type { XlsxViewerController } from "@arcships/xlsx-core"

const XlsxViewerWrapper = defineComponent({
  props: {
    file: { type: Object as PropType<ArrayBuffer | undefined> },
    src: { type: String as PropType<string | undefined> },
    fileName: { type: String as PropType<string | undefined> },
    readOnly: { type: Boolean, default: false },
    useWorker: { type: Boolean, default: true },
    deferLoadingAboveBytes: { type: Number, default: 0 },
    urlPolicy: { type: Object as PropType<XlsxUrlPolicy>, required: true },
  },
  emits: ["controller-ready", "diagnostic", "update:readOnly", "upload"],
  setup(props, { emit }) {
    const controller: XlsxViewerController = useXlsxViewerController({
      file: props.file,
      src: props.src,
      fileName: props.fileName,
      get readOnly() { return props.readOnly },
      useWorker: props.useWorker,
      deferLoadingAboveBytes: props.deferLoadingAboveBytes,
      urlPolicy: props.urlPolicy,
      onDiagnostic: (diagnostic) => emit("diagnostic", diagnostic),
    })
    onMounted(() => emit("controller-ready", controller))
    return () => h(XlsxViewer, {
      controller,
      readOnly: props.readOnly,
      showDefaultToolbar: true,
      showUpload: true,
      style: { height: "100%" },
      onUpload: () => emit("upload"),
      "onUpdate:readOnly": (value: boolean) => emit("update:readOnly", value),
    })
  },
})

const samples = [
  { file: "financial-model.xlsx", label: "财务模型" },
  { file: "sales-table.xlsx", label: "销售表" },
  { file: "charts-images.xlsx", label: "图表、图片和合并单元格" },
  { file: "large-grid.xlsx", label: "大表格" },
  { file: "corrupted.xlsx", label: "损坏文件（失败用例）" },
]

const currentOrigin = globalThis.location.origin
const selectedSample = ref(samples[0].file)
const src = ref<string | null>(null)
const fileBuffer = ref<ArrayBuffer | null>(null)
const displayName = ref("")
const inputError = ref<XlsxLoadError | null>(null)
const controller = shallowRef<XlsxViewerController | null>(null)
const fileInputRef = ref<HTMLInputElement | null>(null)
const readOnly = ref(false)
const useWorker = ref(true)
const sourceKind = ref("未选择")
const loadCounter = ref(0)
const remoteUrl = ref(`/samples/${selectedSample.value}`)
const baseUrl = ref(globalThis.location.href)
const allowedProtocolsText = ref("https:")
const allowedOriginsText = ref(currentOrigin)
const allowHttpOnLocalhost = ref(true)
const deferLoadingAboveBytes = ref(0)
const isDragActive = ref(false)
const diagnostics = ref<Array<XlsxDiagnostic & { id: number }>>([])
let diagnosticId = 0
let dragDepth = 0
let fileReadSequence = 0
let activeFileReader: FileReader | null = null

function splitList(value: string): string[] {
  return value.split(",").map((item) => item.trim()).filter(Boolean)
}

const urlPolicy = computed<XlsxUrlPolicy>(() => ({
  baseUrl: baseUrl.value.trim(),
  allowRelativeUrl: true,
  allowedProtocols: splitList(allowedProtocolsText.value),
  allowedOrigins: splitList(allowedOriginsText.value),
  allowHttpOnLocalhost: allowHttpOnLocalhost.value,
}))

const viewerKey = computed(() => {
  if (!src.value && !fileBuffer.value) return ""
  return `${sourceKind.value}:${displayName.value}-${loadCounter.value}-${useWorker.value}-${deferLoadingAboveBytes.value}`
})
const pageState = computed(() => {
  if (inputError.value) return "error"
  if (!viewerKey.value) return "idle"
  return controller.value?.sourceState ?? "loading"
})
const pageError = computed<XlsxLoadError | null>(() => inputError.value ?? controller.value?.sourceError ?? null)
const actualWorkerMode = computed(() => {
  const current = controller.value
  if (!current || current.sourceState === "idle") return "未初始化"
  if (current.sourceState === "loading") return "加载中"
  if (current.sourceState === "error") return "加载失败"
  return current.isWorkerBacked ? "Worker" : "主线程"
})

function beginLoad(kind: string, name: string) {
  controller.value = null
  inputError.value = null
  sourceKind.value = kind
  displayName.value = name
}

function invalidateFileRead(): number {
  const requestId = ++fileReadSequence
  if (activeFileReader?.readyState === FileReader.LOADING) {
    activeFileReader.abort()
  }
  activeFileReader = null
  return requestId
}

function loadSelectedSample() {
  invalidateFileRead()
  const file = selectedSample.value
  remoteUrl.value = `/samples/${file}`
  beginLoad("示例地址", file)
  fileBuffer.value = null
  src.value = remoteUrl.value
  loadCounter.value++
}

function loadUrl() {
  invalidateFileRead()
  const url = remoteUrl.value.trim()
  if (!url) return
  beginLoad("地址", url.split(/[?#]/)[0]?.split("/").pop() || "未命名工作簿")
  fileBuffer.value = null
  src.value = url
  loadCounter.value++
}

function processFile(file: File) {
  const requestId = invalidateFileRead()
  if (!file.name.toLowerCase().match(/\.xlsx?$/)) {
    inputError.value = { code: "INVALID_WORKBOOK", message: "请选择 .xlsx 或 .xls 文件。", sourceKind: "file" }
    return
  }
  beginLoad("本地文件", file.name)
  src.value = null
  fileBuffer.value = null
  const reader = new FileReader()
  activeFileReader = reader
  reader.onload = () => {
    if (requestId !== fileReadSequence || activeFileReader !== reader) return
    activeFileReader = null
    fileBuffer.value = reader.result as ArrayBuffer
    loadCounter.value++
  }
  reader.onerror = () => {
    if (requestId !== fileReadSequence || activeFileReader !== reader) return
    activeFileReader = null
    inputError.value = { code: "INVALID_WORKBOOK", message: "无法读取工作簿。", sourceKind: "file" }
  }
  reader.readAsArrayBuffer(file)
}

function onFileChange(event: Event) {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  if (file) processFile(file)
}

function openFilePicker() {
  fileInputRef.value?.click()
}

function onControllerReady(nextController: XlsxViewerController) {
  controller.value = nextController
}

function onReadOnlyChange(value: boolean) {
  readOnly.value = value
}

function onDiagnostic(diagnostic: XlsxDiagnostic) {
  diagnostics.value = [...diagnostics.value.slice(-9), { ...diagnostic, id: ++diagnosticId }]
}

function downloadSource() {
  controller.value?.download()
}

function continueDeferredLoad() {
  controller.value?.continueDeferredLoad()
}

function onDragEnter() {
  dragDepth++
  isDragActive.value = true
}

function onDragOver(event: DragEvent) {
  if (event.dataTransfer) event.dataTransfer.dropEffect = "copy"
}

function onDragLeave() {
  dragDepth--
  if (dragDepth <= 0) {
    dragDepth = 0
    isDragActive.value = false
  }
}

function onDrop(event: DragEvent) {
  dragDepth = 0
  isDragActive.value = false
  const file = event.dataTransfer?.files?.[0]
  if (file) processFile(file)
}

onBeforeUnmount(() => {
  invalidateFileRead()
})

loadSelectedSample()
</script>

<style scoped>
.page { padding: 16px; max-width: 1440px; margin: 0 auto; width: 100%; min-width: 0; }
h2 { margin-bottom: 4px; }
.desc { color: var(--muted-foreground); margin-bottom: 16px; font-size: 14px; }
.product-header { display: flex; align-items: end; justify-content: space-between; gap: 16px; flex-wrap: wrap; margin-bottom: 12px; }
.product-header .desc { margin-bottom: 0; }
.product-actions { margin-bottom: 0; }
.product-surface { box-shadow: 0 10px 30px rgb(15 23 42 / 8%); }
.verification-section { margin-top: 8px; }
.runtime-details { margin-top: 12px; border: 1px solid var(--border); border-radius: var(--radius); padding: 10px 12px; color: var(--muted-foreground); }
.runtime-details summary { cursor: pointer; font-size: 13px; font-weight: 600; color: var(--foreground); }
.controls, .policy-panel { display: flex; gap: 12px; align-items: end; flex-wrap: wrap; margin-bottom: 12px; }
.controls label, .policy-panel label { display: flex; gap: 6px; align-items: flex-start; flex-direction: column; font-size: 13px; color: var(--muted-foreground); }
.controls label.inline, .policy-panel .checkbox-label { flex-direction: row; align-items: center; min-height: 36px; }
.controls select, .controls button, .controls input[type="file"], .policy-panel input, .policy-panel button { min-height: 36px; padding: 8px 12px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--background); }
.controls button, .policy-panel button { cursor: pointer; }
.controls button:disabled, .policy-panel button:disabled { opacity: 0.5; cursor: not-allowed; }
.policy-panel { padding: 12px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--muted); }
.policy-panel h3 { flex-basis: 100%; margin: 0; font-size: 14px; }
.policy-panel input:not([type="checkbox"]) { min-width: 180px; }
.error { color: #dc2626; font-size: 13px; margin: 0 0 12px; }
.status-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 8px; padding: 12px; margin-bottom: 16px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--muted); font-size: 13px; }
.viewer-container { border: 1px solid var(--border); border-radius: var(--radius); overflow: auto; margin-bottom: 24px; background: white; position: relative; transition: border-color 0.2s, background 0.2s; }
.viewer-container--drag-active { border-color: #3b82f6; border-width: 2px; background: rgba(59, 130, 246, 0.04); }
.drag-overlay { position: absolute; inset: 0; display: flex; align-items: center; justify-content: center; background: rgba(59, 130, 246, 0.08); z-index: 100; pointer-events: none; }
.drag-overlay span { color: #3b82f6; font-size: 18px; font-weight: 600; }
.empty { display: flex; align-items: center; justify-content: center; height: 240px; color: var(--muted-foreground); }
.diagnostics { margin-top: 24px; border-top: 1px solid var(--border); padding-top: 12px; font-size: 13px; }
.diagnostics h3 { margin: 0 0 8px; }
.diagnostics ol { margin: 0; padding-left: 20px; }
.visually-hidden { position: fixed; inline-size: 1px; block-size: 1px; overflow: hidden; clip-path: inset(50%); white-space: nowrap; }
@media (max-width: 760px) {
  .page { padding: 10px; }
  .product-header { align-items: stretch; }
}
</style>
