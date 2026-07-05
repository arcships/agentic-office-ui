<script setup lang="ts">
import { computed, ref, watch } from "vue"

export interface PdfViewerProps {
  src?: string
  fileName?: string
  defaultZoom?: number
  className?: string
  showToolbar?: boolean
  showDownload?: boolean
  showRotateControls?: boolean
}

const props = withDefaults(defineProps<PdfViewerProps>(), {
  src: undefined,
  fileName: undefined,
  defaultZoom: 1,
  className: "",
  showToolbar: true,
  showDownload: true,
  showRotateControls: true,
})

const emit = defineEmits<{
  "document-load-success": [numPages: number]
  "active-page-change": [pageNumber: number]
}>()

const zoomOptions = [0.5, 0.75, 1, 1.25, 1.5, 2]
const zoom = ref(props.defaultZoom)
const rotation = ref(0)
const activePage = ref(1)
const searchDraft = ref("")
const searchResult = ref("")
const loadError = ref("")
const downloadError = ref("")
const isLoading = ref(false)
const numPages = ref(0)
const pdfText = ref("")

const controlsDisabled = computed(() => !props.src || !!loadError.value || isLoading.value)
const frameSrc = computed(() => {
  if (!props.src) return undefined
  const hash = `page=${activePage.value}&zoom=${Math.round(zoom.value * 100)}`
  return props.src.includes("#") ? `${props.src}&${hash}` : `${props.src}#${hash}`
})
const viewerStyle = computed(() => ({
  transform: `scale(${zoom.value}) rotate(${rotation.value}deg)`,
  transformOrigin: "top center",
  width: `${100 / zoom.value}%`,
  height: `${100 / zoom.value}%`,
}))

function cn(...classes: (string | undefined | null | false)[]): string {
  return classes.filter(Boolean).join(" ")
}

function getDownloadFileName(): string {
  if (props.fileName?.trim()) return props.fileName.toLowerCase().endsWith(".pdf") ? props.fileName : `${props.fileName}.pdf`
  const raw = props.src?.split(/[?#]/)[0]?.split("/").pop() || "document.pdf"
  return raw.toLowerCase().endsWith(".pdf") ? raw : `${raw}.pdf`
}

function zoomIn() {
  const next = zoomOptions.find((z) => z > zoom.value)
  if (next) zoom.value = next
}

function zoomOut() {
  const next = [...zoomOptions].reverse().find((z) => z < zoom.value)
  if (next) zoom.value = next
}

function rotateClockwise() {
  rotation.value = (rotation.value + 90) % 360
}

function setActivePage(page: number) {
  activePage.value = Math.max(1, Math.min(numPages.value || 1, page))
  emit("active-page-change", activePage.value)
}

function goToPage(delta: number) {
  setActivePage(activePage.value + delta)
}

async function handleDownload() {
  if (!props.src) return
  downloadError.value = ""
  try {
    const response = await fetch(props.src)
    if (!response.ok) throw new Error(`Download failed (${response.status})`)
    const blob = await response.blob()
    const url = URL.createObjectURL(blob)
    const a = document.createElement("a")
    a.href = url
    a.download = getDownloadFileName()
    a.rel = "noopener"
    document.body.append(a)
    a.click()
    a.remove()
    setTimeout(() => URL.revokeObjectURL(url), 0)
  } catch (e) {
    const err = e instanceof Error ? e : new Error(String(e))
    downloadError.value = err.message
  }
}

function estimatePagesFromPdf(text: string): number {
  const matches = text.match(/\/Type\s*\/Page\b/g)
  if (matches?.length) return matches.length
  return 1
}

function runSearch() {
  const query = searchDraft.value.trim()
  if (!query) {
    searchResult.value = ""
    return
  }
  const hits = pdfText.value.toLowerCase().split(query.toLowerCase()).length - 1
  searchResult.value = hits > 0 ? `${hits} match${hits === 1 ? "" : "es"}` : "No matches"
}

async function load() {
  loadError.value = ""
  downloadError.value = ""
  searchResult.value = ""
  pdfText.value = ""
  numPages.value = 0
  activePage.value = 1
  if (!props.src) return

  isLoading.value = true
  try {
    const response = await fetch(props.src)
    if (!response.ok) throw new Error(`Failed to fetch PDF (${response.status})`)
    const buffer = await response.arrayBuffer()
    const bytes = new Uint8Array(buffer)
    const header = new TextDecoder("latin1").decode(bytes.slice(0, 8))
    const text = new TextDecoder("latin1").decode(buffer)
    if (!header.startsWith("%PDF-") || text.includes("%PDF corrupted fixture")) {
      throw new Error("Unable to load PDF document.")
    }
    pdfText.value = text
    numPages.value = estimatePagesFromPdf(text)
    emit("document-load-success", numPages.value)
    emit("active-page-change", activePage.value)
  } catch (e) {
    const err = e instanceof Error ? e : new Error(String(e))
    loadError.value = err.message
  } finally {
    isLoading.value = false
  }
}

watch(() => props.src, load, { immediate: true })
watch(searchDraft, runSearch)
</script>

<template>
  <div :class="cn('pdf-viewer', props.className)">
    <div v-if="props.showToolbar" class="pdf-toolbar">
      <button type="button" :disabled="controlsDisabled" aria-label="Toggle thumbnails">☰</button>
      <span class="file-name">{{ getDownloadFileName() }}</span>
      <button type="button" :disabled="controlsDisabled || activePage <= 1" aria-label="Previous page" @click="goToPage(-1)">‹</button>
      <span class="page-indicator">{{ activePage }} / {{ numPages || '—' }}</span>
      <button type="button" :disabled="controlsDisabled || activePage >= numPages" aria-label="Next page" @click="goToPage(1)">›</button>
      <button type="button" :disabled="controlsDisabled || zoom <= zoomOptions[0]" aria-label="Zoom out" @click="zoomOut">−</button>
      <select v-model.number="zoom" :disabled="controlsDisabled" aria-label="Zoom level">
        <option v-for="z in zoomOptions" :key="z" :value="z">{{ Math.round(z * 100) }}%</option>
      </select>
      <button type="button" :disabled="controlsDisabled || zoom >= zoomOptions[zoomOptions.length - 1]" aria-label="Zoom in" @click="zoomIn">+</button>
      <button v-if="props.showRotateControls" type="button" :disabled="controlsDisabled" aria-label="Rotate page" @click="rotateClockwise">⟳</button>
      <input v-model="searchDraft" :disabled="controlsDisabled" aria-label="Search text" placeholder="Search text" />
      <span v-if="searchResult" class="search-result">{{ searchResult }}</span>
      <button v-if="props.showDownload" type="button" :disabled="controlsDisabled" aria-label="Download PDF" @click="handleDownload">⬇</button>
    </div>

    <div v-if="!props.src" class="pdf-empty">No PDF loaded</div>
    <div v-else-if="isLoading" class="pdf-empty">Loading PDF...</div>
    <div v-else-if="loadError" class="pdf-error">{{ loadError }}</div>
    <div v-else class="pdf-body">
      <aside class="pdf-thumbnails" aria-label="Page thumbnails">
        <button
          v-for="page in numPages"
          :key="page"
          type="button"
          :class="['thumbnail', { active: page === activePage }]"
          @click="setActivePage(page)"
        >
          Page {{ page }}
        </button>
      </aside>
      <div class="pdf-frame-wrap">
        <iframe :src="frameSrc" title="PDF document" class="pdf-frame" :style="viewerStyle" />
      </div>
    </div>
    <div v-if="downloadError" class="download-error">{{ downloadError }}</div>
  </div>
</template>

<style scoped>
.pdf-viewer { display: flex; flex-direction: column; height: 100%; min-height: 320px; background: var(--background, #fff); color: var(--foreground, #111); }
.pdf-toolbar { display: flex; align-items: center; gap: 8px; min-height: 48px; padding: 8px; border-bottom: 1px solid var(--border, #e5e5e5); flex-wrap: wrap; }
.pdf-toolbar button, .pdf-toolbar select, .pdf-toolbar input { height: 32px; border: 1px solid var(--border, #e5e5e5); border-radius: 6px; background: white; padding: 0 8px; }
.pdf-toolbar button { cursor: pointer; }
.pdf-toolbar button:disabled, .pdf-toolbar input:disabled, .pdf-toolbar select:disabled { opacity: 0.5; cursor: not-allowed; }
.file-name { font-size: 13px; font-weight: 600; margin-right: auto; }
.page-indicator { font-size: 12px; min-width: 56px; text-align: center; }
.search-result { font-size: 12px; color: #2563eb; min-width: 72px; }
.pdf-empty, .pdf-error { display: grid; place-items: center; flex: 1; min-height: 220px; color: var(--muted-foreground, #737373); }
.pdf-error, .download-error { color: #dc2626; }
.download-error { padding: 8px 12px; font-size: 12px; border-top: 1px solid var(--border, #e5e5e5); }
.pdf-body { display: flex; min-height: 0; flex: 1; background: #f5f5f5; }
.pdf-thumbnails { width: 132px; overflow: auto; border-right: 1px solid var(--border, #e5e5e5); padding: 12px; background: white; }
.thumbnail { display: block; width: 100%; height: 72px; margin-bottom: 8px; border: 1px solid var(--border, #e5e5e5); border-radius: 6px; background: #fafafa; font-size: 12px; }
.thumbnail.active { border-color: #2563eb; color: #2563eb; font-weight: 700; }
.pdf-frame-wrap { flex: 1; overflow: auto; display: flex; justify-content: center; padding: 16px; }
.pdf-frame { min-height: 100%; border: 0; background: white; box-shadow: 0 2px 8px rgba(0,0,0,0.12); }
</style>
