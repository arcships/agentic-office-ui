<template>
  <div class="page">
    <h2>📊 XLSX Viewer — Verification</h2>
    <p class="desc">Verify real workbook loading, sheet tabs, selection, editing, zoom, read-only and large-grid behavior.</p>

    <div class="controls control-panel">
      <label>
        Sample workbook
        <select v-model="selectedSample" @change="loadSelectedSample">
          <option v-for="sample in samples" :key="sample.file" :value="sample.file">
            {{ sample.label }}
          </option>
        </select>
      </label>
      <button @click="loadSelectedSample">Load sample</button>
      <input type="file" accept=".xlsx,.xls" @change="onFileChange" />
      <label class="inline">
        <input v-model="readOnly" type="checkbox" /> Read only
      </label>
      <label class="inline">
        <input v-model="useWorker" type="checkbox" /> Worker
      </label>
      <p v-if="error" class="error">{{ error }}</p>
    </div>

    <div class="status-grid info-grid">
      <div><strong>Loaded:</strong> {{ displayName || "None" }}</div>
      <div><strong>Source:</strong> {{ sourceKind }}</div>
      <div><strong>Mode:</strong> {{ readOnly ? "Read only" : "Editable" }}</div>
      <div><strong>Worker:</strong> {{ useWorker ? "On" : "Off" }}</div>
      <div><strong>Expected coverage:</strong> tabs, selection, double-click edit, zoom, large grid</div>
    </div>

    <div class="viewer-container">
      <XlsxViewerWrapper
        v-if="viewerKey"
        :key="viewerKey"
        :file="fileBuffer ?? undefined"
        :src="src ?? undefined"
        :file-name="displayName || 'verification.xlsx'"
        :read-only="readOnly"
        :use-worker="useWorker"
        style="height: 70vh;"
      />
      <div v-else class="empty">
        <p>Load a workbook sample or upload an .xlsx file to begin verification.</p>
      </div>
    </div>

    <div class="api-verify">
      <h3>API Verification</h3>
      <table>
        <thead><tr><th>Check</th><th>Expected</th><th>Result</th></tr></thead>
        <tbody>
          <tr><td><code>XlsxViewer</code></td><td>Accepts controller prop; visible spreadsheet surface</td><td class="warn">⚠️ SCOPED</td></tr>
          <tr><td><code>useXlsxViewerController()</code></td><td>Selection, zoom, editing, tabs controller; some advanced APIs are partial/no-op</td><td class="warn">⚠️ PARTIAL</td></tr>
          <tr><td><code>columnLabel(26)</code></td><td>AA</td><td class="pass">{{ columnAA }}</td></tr>
          <tr><td><code>rangeToA1()</code></td><td>A1:C6</td><td class="pass">{{ rangeCheck }}</td></tr>
        </tbody>
      </table>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, h, defineComponent, type PropType } from "vue"
import { useXlsxViewerController, XlsxViewer } from "@extend-ai/vue-xlsx"
import type { XlsxViewerController } from "@extend-ai/xlsx-core"
import { columnLabel, rangeToA1 } from "@extend-ai/xlsx-core"

// Wrapper component: creates controller via useXlsxViewerController and passes it to XlsxViewer.
// Keyed by parent so it unmounts/remounts when the source changes, giving a fresh controller.
const XlsxViewerWrapper = defineComponent({
  props: {
    file: { type: Object as PropType<ArrayBuffer | undefined> },
    src: { type: String as PropType<string | undefined> },
    fileName: { type: String as PropType<string | undefined> },
    readOnly: { type: Boolean, default: false },
    useWorker: { type: Boolean, default: true },
  },
  setup(props) {
    const controller: XlsxViewerController = useXlsxViewerController({
      file: props.file,
      src: props.src,
      fileName: props.fileName,
      readOnly: props.readOnly,
      useWorker: props.useWorker,
    })
    return () => h(XlsxViewer, { controller, showDefaultToolbar: true, style: { height: "100%" } })
  },
})

const samples = [
  { file: "financial-model.xlsx", label: "Financial model — formulas + multiple sheets" },
  { file: "sales-table.xlsx", label: "Sales table — long text + 180 rows" },
  { file: "charts-images.xlsx", label: "Charts and images workbook" },
  { file: "large-grid.xlsx", label: "Large grid — 550 × 60" },
  { file: "corrupted.xlsx", label: "Corrupted workbook — negative test" },
]

const selectedSample = ref(samples[0].file)
const src = ref<string | null>(null)
const fileBuffer = ref<ArrayBuffer | null>(null)
const displayName = ref("")
const error = ref<string | null>(null)
const readOnly = ref(false)
const useWorker = ref(true)
const sourceKind = ref("sample")
const loadCounter = ref(0)

const viewerKey = computed(() => {
  if (!src.value && !fileBuffer.value) return ""
  return `${sourceKind.value}:${displayName.value}-${loadCounter.value}-${readOnly.value}-${useWorker.value}`
})
const columnAA = computed(() => columnLabel(26))
const rangeCheck = computed(() => rangeToA1({ start: { col: 0, row: 0 }, end: { col: 2, row: 5 } }))

function loadSelectedSample() {
  error.value = null
  fileBuffer.value = null
  src.value = `/samples/${selectedSample.value}`
  displayName.value = selectedSample.value
  sourceKind.value = "sample"
  loadCounter.value++
}

function onFileChange(e: Event) {
  const input = e.target as HTMLInputElement
  const file = input.files?.[0]
  if (!file) return
  if (!file.name.toLowerCase().match(/\.xlsx?$/)) {
    error.value = "Please select an .xlsx or .xls file"
    return
  }
  error.value = null
  sourceKind.value = "upload"
  displayName.value = file.name
  src.value = null
  const reader = new FileReader()
  reader.onload = () => {
    fileBuffer.value = reader.result as ArrayBuffer
    loadCounter.value++
  }
  reader.onerror = () => { error.value = "Failed to read workbook" }
  reader.readAsArrayBuffer(file)
}

loadSelectedSample()
</script>

<style scoped>
.page { padding: 24px; max-width: 1200px; margin: 0 auto; width: 100%; min-width: 0; }
h2 { margin-bottom: 4px; }
.desc { color: var(--muted-foreground); margin-bottom: 16px; font-size: 14px; }
.controls { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin-bottom: 12px; }
.controls label { display: flex; gap: 6px; align-items: center; font-size: 13px; color: var(--muted-foreground); }
.controls label:not(.inline) { flex-direction: column; align-items: flex-start; }
.controls select, .controls button, .controls input[type="file"] { padding: 8px 12px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--background); }
.controls button { cursor: pointer; }
.error { color: #ef4444; font-size: 13px; }
.status-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 8px; padding: 12px; margin-bottom: 16px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--muted); font-size: 13px; }
.viewer-container { border: 1px solid var(--border); border-radius: var(--radius); overflow: auto; margin-bottom: 24px; background: white; }
.empty { display: flex; align-items: center; justify-content: center; height: 240px; color: var(--muted-foreground); }
.api-verify { margin-top: 24px; }
.api-verify table { width: 100%; border-collapse: collapse; font-size: 13px; }
.api-verify th, .api-verify td { padding: 6px 12px; text-align: left; border-bottom: 1px solid var(--border); }
.pass { color: #16a34a; font-weight: 600; }
.warn { color: #b45309; font-weight: 600; }
</style>
