<template>
  <div class="page">
    <h2>📑 PDF Viewer — Verification</h2>
    <p class="desc">Verify PdfViewer with real sample PDFs: text, scanned, rotated, large and corrupted fixtures.</p>

    <div class="controls control-panel">
      <label>
        Sample PDF
        <select v-model="selectedSample" @change="loadSelectedSample">
          <option v-for="sample in samples" :key="sample.file" :value="sample.file">
            {{ sample.label }}
          </option>
        </select>
      </label>
      <button @click="loadSelectedSample">Load sample</button>
      <input type="file" accept=".pdf" @change="onFileChange" />
      <p v-if="error" class="error">{{ error }}</p>
    </div>

    <div class="status-grid info-grid">
      <div><strong>Loaded:</strong> {{ fileName || "None" }}</div>
      <div><strong>Source:</strong> {{ sourceKind }}</div>
      <div><strong>Expected coverage:</strong> thumbnails, zoom, search, rotate, download, error state</div>
    </div>

    <div class="viewer-container">
      <PdfViewer
        v-if="pdfUrl"
        :key="viewerKey"
        :src="pdfUrl"
        :default-zoom="1"
        :file-name="fileName || 'verification.pdf'"
        style="height: 80vh;"
        @document-load-success="numPages = $event"
        @active-page-change="activePage = $event"
      />
      <div v-else class="empty">
        <p>Upload or load a PDF file to begin verification.</p>
      </div>
    </div>

    <div class="api-verify">
      <h3>Runtime Status</h3>
      <table>
        <thead><tr><th>Check</th><th>Expected</th><th>Result</th></tr></thead>
        <tbody>
          <tr><td>PdfViewer component</td><td>Accepts <code>src</code>, <code>fileName</code>, <code>defaultZoom</code></td><td class="pass">✅ READY</td></tr>
          <tr><td>Loaded pages</td><td>Document load event reports page count</td><td>{{ numPages || "Pending" }}</td></tr>
          <tr><td>Active page</td><td>Page state updates after navigation</td><td>{{ activePage }}</td></tr>
          <tr><td>Negative fixture</td><td>corrupted.pdf should show engine/document error behavior</td><td class="pass">Available</td></tr>
        </tbody>
      </table>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, onBeforeUnmount, ref } from "vue"
import { PdfViewer } from "@extend-ai/vue-extend"

const samples = [
  { file: "sample.pdf", label: "Text PDF — 4 pages + searchable content" },
  { file: "scanned-invoice.pdf", label: "Scanned invoice — image pages" },
  { file: "rotated-pages.pdf", label: "Rotated pages — portrait + landscape" },
  { file: "large-contract.pdf", label: "Large contract — 31 pages" },
  { file: "corrupted.pdf", label: "Corrupted PDF — negative test" },
]

const selectedSample = ref(samples[0].file)
const pdfUrl = ref<string | null>(null)
const fileName = ref("")
const sourceKind = ref("sample")
const error = ref<string | null>(null)
const numPages = ref(0)
const activePage = ref(1)
const loadCounter = ref(0)
let objectUrl: string | null = null

function revokeObjectUrl() {
  if (objectUrl) {
    URL.revokeObjectURL(objectUrl)
    objectUrl = null
  }
}

const viewerKey = computed(() => `${fileName.value}-${loadCounter.value}`)

function loadSelectedSample() {
  revokeObjectUrl()
  error.value = null
  pdfUrl.value = `/samples/${selectedSample.value}`
  fileName.value = selectedSample.value
  sourceKind.value = "sample"
  numPages.value = 0
  activePage.value = 1
  loadCounter.value++
}

function onFileChange(e: Event) {
  const input = e.target as HTMLInputElement
  const file = input.files?.[0]
  if (!file) return
  if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
    error.value = "Please select a PDF file"
    return
  }
  revokeObjectUrl()
  objectUrl = URL.createObjectURL(file)
  pdfUrl.value = objectUrl
  fileName.value = file.name
  sourceKind.value = "upload"
  numPages.value = 0
  activePage.value = 1
  error.value = null
  loadCounter.value++
}

onBeforeUnmount(revokeObjectUrl)

loadSelectedSample()
</script>

<style scoped>
.page { padding: 24px; max-width: 1200px; margin: 0 auto; width: 100%; min-width: 0; }
h2 { margin-bottom: 4px; }
.desc { color: var(--muted-foreground); margin-bottom: 16px; font-size: 14px; }
.controls { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin-bottom: 12px; }
.controls label { display: flex; flex-direction: column; gap: 6px; align-items: flex-start; font-size: 13px; color: var(--muted-foreground); }
.controls input, .controls select, .controls button { padding: 8px 12px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--background); }
.controls button { cursor: pointer; }
.error { color: #ef4444; font-size: 13px; }
.status-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 8px; padding: 12px; margin-bottom: 16px; border: 1px solid var(--border); border-radius: var(--radius); background: var(--muted); font-size: 13px; }
.viewer-container { border: 1px solid var(--border); border-radius: var(--radius); overflow: auto; margin-bottom: 24px; }
.empty { display: flex; align-items: center; justify-content: center; height: 200px; color: var(--muted-foreground); }
.api-verify { margin-top: 24px; }
.api-verify table { width: 100%; border-collapse: collapse; font-size: 13px; }
.api-verify th, .api-verify td { padding: 6px 12px; border-bottom: 1px solid var(--border); }
.pass { color: #16a34a; font-weight: 600; }
</style>
