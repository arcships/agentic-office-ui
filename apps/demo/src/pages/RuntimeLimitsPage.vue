<template>
  <main class="page" data-testid="runtime-limits-page">
    <h2>DOCX / XLSX 资源上限公开消费页</h2>
    <p>本页只使用公开 Runtime 和组件接口，可由宿主配置资源上限并查看结构化结果。</p>

    <div data-testid="page-status" data-state="ready">ready</div>

    <section class="panel" data-testid="runtime-limits-config">
      <label>
        格式
        <select v-model="format" data-testid="runtime-limits-format" @change="changeFormat">
          <option value="docx">DOCX</option>
          <option value="xlsx">XLSX</option>
        </select>
      </label>
      <label>
        Runtime limits（JSON）
        <textarea v-model="limitsText" data-testid="runtime-limits-json" rows="12" spellcheck="false" />
      </label>
      <button data-testid="runtime-limits-apply" type="button" @click="applyLimits">应用上限</button>
      <label>
        上传 {{ format.toUpperCase() }}
        <input
          :key="`${format}-${inputKey}`"
          data-testid="runtime-limits-file-input"
          type="file"
          :accept="format === 'docx' ? '.docx' : '.xlsx,.xls'"
          @change="onFileChange"
        />
      </label>
      <p v-if="configError" data-testid="runtime-limits-config-error">{{ configError }}</p>
      <pre data-testid="runtime-limits-active">{{ JSON.stringify(activeLimits, null, 2) }}</pre>
    </section>

    <section class="panel" aria-live="polite">
      <div data-testid="runtime-limits-status" :data-state="state">{{ state }}</div>
      <div data-testid="runtime-limits-loaded-file">{{ fileName || "未选择" }}</div>
      <div
        v-if="visibleError"
        data-testid="runtime-limits-error"
        :data-error-code="visibleError.code"
        :data-error-phase="visibleError.phase"
        :data-error-limit="visibleError.limit"
        :data-error-actual="visibleError.actual"
        :data-error-allowed="visibleError.allowed"
      >
        <div><span data-testid="runtime-limits-error-code">{{ visibleError.code }}</span>: {{ visibleError.message }}</div>
        <div>phase=<span data-testid="runtime-limits-error-phase">{{ visibleError.phase ?? "" }}</span></div>
        <div>limit=<span data-testid="runtime-limits-error-limit">{{ visibleError.limit ?? "" }}</span></div>
        <div>actual=<span data-testid="runtime-limits-error-actual">{{ visibleError.actual ?? "" }}</span></div>
        <div>allowed=<span data-testid="runtime-limits-error-allowed">{{ visibleError.allowed ?? "" }}</span></div>
      </div>
    </section>

    <section class="surface">
      <DocxViewer
        v-if="format === 'docx' && fileBuffer"
        :key="`docx-${generation}`"
        :file="fileBuffer"
        :runtime="docxRuntime"
        @load-start="state = 'loading'"
        @load-success="onDocxSuccess"
        @load-error="onLoadError"
      />
      <XlsxLimitsConsumer
        v-else-if="format === 'xlsx' && fileBuffer"
        :key="`xlsx-${generation}`"
        :file="fileBuffer"
        :file-name="fileName"
        :runtime="xlsxRuntime"
        @controller-ready="onXlsxControllerReady"
        @diagnostic="recordDiagnostic"
      />
      <p v-else>应用上限后上传文件。</p>
    </section>

    <section class="panel">
      <h3>公开诊断</h3>
      <pre data-testid="runtime-limits-diagnostics">{{ JSON.stringify(diagnostics, null, 2) }}</pre>
    </section>
  </main>
</template>

<script setup lang="ts">
import {
  computed,
  defineComponent,
  h,
  onBeforeUnmount,
  onMounted,
  ref,
  shallowRef,
  watch,
  type PropType,
} from "vue"
import {
  createDocxRuntime,
  type DocxImportDiagnostic,
  type DocxImportResult,
  type DocxRuntime,
  type DocxRuntimeLimits,
} from "@arcships/docx-core"
import {
  createXlsxRuntime,
  type XlsxRuntime,
  type XlsxRuntimeDiagnostic,
  type XlsxRuntimeLimits,
  type XlsxViewerController,
} from "@arcships/xlsx-core"
import { DocxViewer } from "@arcships/vue-docx"
import {
  useXlsxViewerController,
  XlsxViewer,
  type XlsxDiagnostic,
} from "@arcships/vue-xlsx"

type Format = "docx" | "xlsx"
type LoadState = "idle" | "loading" | "ready" | "error"
type StructuredError = {
  code: string
  message: string
  phase?: string
  limit?: string
  actual?: number
  allowed?: number
}

const COMMON_LIMIT_KEYS = [
  "maxInputBytes", "maxArchiveEntries", "maxUncompressedBytes",
  "maxSingleEntryBytes", "maxCompressionRatio", "maxArchivePathLength",
  "maxXmlBytes", "maxSingleXmlBytes", "maxXmlDepth",
  "maxXmlAttributeBytes", "maxTextNodeBytes", "maxRelationships", "maxParseMs",
  "maxSingleImageBytes", "maxTotalImageBytes", "maxImageWidth", "maxImageHeight",
  "maxSingleImagePixels", "maxTotalImagePixels", "maxConcurrentImageDecodes",
] as const
const DOCX_LIMIT_KEYS = new Set([...COMMON_LIMIT_KEYS, "maxDocxNodes", "maxDocxPages"])
const XLSX_LIMIT_KEYS = new Set([
  ...COMMON_LIMIT_KEYS, "maxWorksheetXmlBytes", "maxSharedStringsBytes",
  "maxWorksheetRows", "maxWorksheetColumns", "maxWorksheets",
  "maxSharedStrings", "maxFormulaCount",
])

const XlsxLimitsConsumer = defineComponent({
  props: {
    file: { type: Object as PropType<ArrayBuffer>, required: true },
    fileName: { type: String, required: true },
    runtime: { type: Object as PropType<XlsxRuntime>, required: true },
  },
  emits: ["controller-ready", "diagnostic"],
  setup(props, { emit }) {
    const controller = useXlsxViewerController({
      file: props.file,
      fileName: props.fileName,
      runtime: props.runtime,
      readOnly: true,
      useWorker: true,
      onDiagnostic: (diagnostic) => emit("diagnostic", diagnostic),
    })
    onMounted(() => emit("controller-ready", controller))
    return () => h(XlsxViewer, {
      controller,
      showDefaultToolbar: false,
      showFormulaBar: false,
      showRibbon: false,
      style: { height: "360px" },
    })
  },
})

const format = ref<Format>("docx")
const docxLimitsText = ref("{}")
const xlsxLimitsText = ref("{}")
const docxLimits = ref<DocxRuntimeLimits>({})
const xlsxLimits = ref<XlsxRuntimeLimits>({})
const state = ref<LoadState>("idle")
const configError = ref("")
const fileBuffer = shallowRef<ArrayBuffer | null>(null)
const fileName = ref("")
const inputKey = ref(0)
const generation = ref(0)
const runtimeError = shallowRef<StructuredError | null>(null)
const xlsxController = shallowRef<XlsxViewerController | null>(null)
const diagnostics = ref<unknown[]>([])
let fileReadSequence = 0

function recordDiagnostic(event: DocxImportDiagnostic | XlsxDiagnostic | XlsxRuntimeDiagnostic): void {
  diagnostics.value = [...diagnostics.value.slice(-49), event]
}

function makeDocxRuntime(): DocxRuntime {
  return createDocxRuntime({ limits: docxLimits.value, onDiagnostic: recordDiagnostic })
}

function makeXlsxRuntime(): XlsxRuntime {
  return createXlsxRuntime({ limits: xlsxLimits.value, onDiagnostic: recordDiagnostic })
}

const docxRuntime = shallowRef(makeDocxRuntime())
const xlsxRuntime = shallowRef(makeXlsxRuntime())
const limitsText = computed({
  get: () => format.value === "docx" ? docxLimitsText.value : xlsxLimitsText.value,
  set: (value: string) => {
    if (format.value === "docx") docxLimitsText.value = value
    else xlsxLimitsText.value = value
  },
})
const activeLimits = computed(() => format.value === "docx" ? docxRuntime.value.limits : xlsxRuntime.value.limits)
const visibleError = computed<StructuredError | null>(() => {
  if (runtimeError.value) return runtimeError.value
  if (format.value === "xlsx") return xlsxController.value?.sourceError ?? null
  return null
})

watch(
  () => [format.value, xlsxController.value?.sourceState, xlsxController.value?.sourceError] as const,
  ([currentFormat, sourceState, sourceError]) => {
    if (currentFormat !== "xlsx" || !sourceState) return
    state.value = sourceError ? "error" : sourceState
  },
)

function parseLimits(text: string, allowedKeys: ReadonlySet<string>): Record<string, number> {
  const value = JSON.parse(text) as unknown
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("limits 必须是 JSON 对象。")
  }
  const limits: Record<string, number> = {}
  for (const [key, raw] of Object.entries(value)) {
    if (!allowedKeys.has(key)) throw new Error(`当前格式不支持限制字段 ${key}。`)
    if (typeof raw !== "number" || !Number.isFinite(raw) || raw <= 0) {
      throw new Error(`${key} 必须是大于 0 的有限数字。`)
    }
    limits[key] = raw
  }
  return limits
}

function clearResult(): void {
  runtimeError.value = null
  xlsxController.value = null
  diagnostics.value = []
  state.value = fileBuffer.value ? "loading" : "idle"
}

function applyLimits(): void {
  try {
    if (format.value === "docx") {
      docxLimits.value = parseLimits(docxLimitsText.value, DOCX_LIMIT_KEYS) as DocxRuntimeLimits
      docxRuntime.value.dispose()
      docxRuntime.value = makeDocxRuntime()
    } else {
      xlsxLimits.value = parseLimits(xlsxLimitsText.value, XLSX_LIMIT_KEYS) as XlsxRuntimeLimits
      xlsxRuntime.value.dispose()
      xlsxRuntime.value = makeXlsxRuntime()
    }
    configError.value = ""
    clearResult()
    if (fileBuffer.value) generation.value += 1
  } catch (error) {
    configError.value = error instanceof Error ? error.message : String(error)
  }
}

function changeFormat(): void {
  fileReadSequence += 1
  fileBuffer.value = null
  fileName.value = ""
  inputKey.value += 1
  clearResult()
  configError.value = ""
}

async function onFileChange(event: Event): Promise<void> {
  const file = (event.target as HTMLInputElement).files?.[0]
  if (!file) return
  const expected = format.value === "docx" ? /\.docx$/i : /\.xlsx?$/i
  if (!expected.test(file.name)) {
    runtimeError.value = { code: "INVALID_SOURCE", message: `请选择 ${format.value.toUpperCase()} 文件。` }
    state.value = "error"
    return
  }
  const requestId = ++fileReadSequence
  state.value = "loading"
  runtimeError.value = null
  xlsxController.value = null
  diagnostics.value = []
  const buffer = await file.arrayBuffer()
  if (requestId !== fileReadSequence) return
  fileName.value = file.name
  fileBuffer.value = buffer
  generation.value += 1
}

function normalizeError(error: unknown): StructuredError {
  const value = error instanceof Error ? error : new Error(String(error))
  const details = value as Error & {
    code?: unknown; phase?: unknown; limit?: unknown; actual?: unknown; allowed?: unknown
  }
  return {
    code: typeof details.code === "string" ? details.code : "PARSE_FAILED",
    message: value.message,
    ...(typeof details.phase === "string" ? { phase: details.phase } : {}),
    ...(typeof details.limit === "string" ? { limit: details.limit } : {}),
    ...(typeof details.actual === "number" ? { actual: details.actual } : {}),
    ...(typeof details.allowed === "number" ? { allowed: details.allowed } : {}),
  }
}

function onLoadError(error: unknown): void {
  runtimeError.value = normalizeError(error)
  state.value = "error"
}

function onDocxSuccess(_result: DocxImportResult): void {
  runtimeError.value = null
  state.value = "ready"
}

function onXlsxControllerReady(controller: XlsxViewerController): void {
  xlsxController.value = controller
  state.value = controller.sourceState
}

onBeforeUnmount(() => {
  fileReadSequence += 1
  docxRuntime.value.dispose()
  xlsxRuntime.value.dispose()
})
</script>

<style scoped>
.page { max-width: 1100px; margin: 0 auto; padding: 24px; }
.panel { display: grid; gap: 8px; margin: 12px 0; }
label { display: grid; gap: 4px; }
textarea, select, button, input { font: inherit; padding: 6px; }
pre { overflow: auto; white-space: pre-wrap; }
.surface { min-height: 120px; border: 1px solid var(--border); padding: 8px; }
</style>
