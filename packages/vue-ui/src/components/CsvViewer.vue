<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import {
  computed,
  nextTick,
  onBeforeUnmount,
  onMounted,
  ref,
  watch,
} from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"
import type { OfficeMenuEntry } from "./ui/types"
import type { OfficeSearchStatus } from "./chrome/types"
import {
  decodeCsvBuffer,
  findCsvMatches,
  parseCsvText,
  resolveCsvVisibleRange,
  serializeCsv,
  type CsvMatch,
  type CsvParseResult,
} from "../csv"
import {
  provideOfficeUiMessages,
  useOfficeUiLocale,
  useOfficeUiMessages,
} from "../i18n"
import OfficeIconButton from "./ui/OfficeIconButton.vue"
import OfficeButton from "./ui/OfficeButton.vue"
import OfficeDropdownMenu from "./ui/OfficeDropdownMenu.vue"
import { officeClass, toCssLength } from "./ui/types"
import OfficeEmptyState from "./chrome/OfficeEmptyState.vue"
import OfficeErrorState from "./chrome/OfficeErrorState.vue"
import OfficeLoadingState from "./chrome/OfficeLoadingState.vue"
import OfficeSearchPopover from "./chrome/OfficeSearchPopover.vue"
import OfficeToolbar from "./chrome/OfficeToolbar.vue"
import OfficeToolbarGroup from "./chrome/OfficeToolbarGroup.vue"
import OfficeToolbarSeparator from "./chrome/OfficeToolbarSeparator.vue"
import OfficeZoomControl from "./chrome/OfficeZoomControl.vue"
import type {
  CsvViewerDownloadEvent,
  CsvViewerError,
  CsvViewerLoadEvent,
  CsvViewerProps,
  CsvViewerState,
} from "../types"

const props = withDefaults(defineProps<CsvViewerProps>(), {
  data: undefined,
  file: undefined,
  src: undefined,
  fileName: undefined,
  delimiter: "auto",
  search: true,
  searchDebounce: 300,
  height: undefined,
  maxFileSize: 50 * 1024 * 1024,
  loading: false,
  locale: undefined,
  messages: undefined,
  class: "",
})

const emit = defineEmits<{
  load: [event: CsvViewerLoadEvent]
  error: [error: CsvViewerError]
  download: [event: CsvViewerDownloadEvent]
  "update:data": [text: string]
}>()

const inheritedLocale = useOfficeUiLocale()
const inheritedMessages = useOfficeUiMessages()
const messagesOverride = computed(() => {
  if (props.messages !== undefined) return props.messages
  // The fallback catalog is English; passing it back as an override would
  // replace every translated key when `locale` switches to zh-CN. Only merge
  // an ambient catalog when the host did not choose a locale explicitly.
  if (props.locale !== undefined) return undefined
  return inheritedMessages.value
})
const m = provideOfficeUiMessages(
  () => props.locale ?? inheritedLocale.value,
  () => messagesOverride.value,
)

const ZOOM_OPTIONS = [75, 100, 125, 150, 200]
const BASE_ROW_HEIGHT = 34
const BASE_HEADER_HEIGHT = 36
const BASE_ZOOM = 100

const emptyParse = (): CsvParseResult => ({
  headers: [],
  rows: [],
  delimiter: ",",
  sourceHeaderCount: 0,
  columnCount: 0,
  error: null,
})

const parsed = ref<CsvParseResult>(emptyParse())
const errorState = ref<CsvViewerError | null>(null)
const fileName = ref<string | null>(props.fileName ?? null)
const isReading = ref(false)
const zoom = ref(BASE_ZOOM)
const scrollTop = ref(0)
const viewportHeight = ref(560)
const selectedCell = ref<{ row: number; col: number } | null>(null)
const viewportRef = ref<HTMLElement | null>(null)
const fileInputRef = ref<HTMLInputElement | null>(null)

const searchQuery = ref("")
const searchResults = ref<CsvMatch[]>([])
const activeResultIndex = ref(0)
const searchStatus = ref<OfficeSearchStatus>("idle")
const searchRequestId = ref(0)
let searchTimer: ReturnType<typeof setTimeout> | undefined
let readRequestId = 0
let activeFetchController: AbortController | null = null
let resizeObserver: ResizeObserver | null = null
const lastLoad = ref<
  | { kind: "file"; file: Blob; fileName: string | null }
  | { kind: "src"; src: string }
  | null
>(null)

const status = computed<CsvViewerState["status"]>(() => {
  if (props.loading || isReading.value) return "loading"
  if (errorState.value) return "error"
  if (parsed.value.rows.length === 0) return "empty"
  return "ready"
})

const zoomScale = computed(() => zoom.value / BASE_ZOOM)
const rowHeight = computed(() =>
  Math.max(1, Math.round(BASE_ROW_HEIGHT * zoomScale.value)),
)
const headerHeight = computed(() =>
  Math.max(1, Math.round(BASE_HEADER_HEIGHT * zoomScale.value)),
)
const rowHeaderWidth = computed(() =>
  Math.max(24, Math.round(44 * zoomScale.value)),
)
const columnWidths = computed(() =>
  parsed.value.headers.map((_, index) =>
    Math.max(80, Math.round((index === 0 ? 180 : 160) * zoomScale.value)),
  ),
)
const tableWidth = computed(
  () =>
    rowHeaderWidth.value +
    columnWidths.value.reduce((total, width) => total + width, 0),
)
const visibleRange = computed(() =>
  resolveCsvVisibleRange({
    scrollTop: scrollTop.value,
    viewportHeight: viewportHeight.value,
    rowCount: parsed.value.rows.length,
    rowHeight: rowHeight.value,
  }),
)

const renderedRows = computed(() =>
  parsed.value.rows
    .slice(visibleRange.value.start, visibleRange.value.end)
    .map((cells, offset) => ({
      index: visibleRange.value.start + offset,
      cells,
    })),
)

const rootClasses = computed(() =>
  officeClass("office-ui", "office-csv-viewer", props.class),
)

const rootStyle = computed(() =>
  props.height === undefined
    ? undefined
    : { blockSize: toCssLength(props.height) },
)

const tableStyle = computed(() => ({
  fontSize: `${Math.max(1, Math.round(13 * zoomScale.value))}px`,
  width: `${tableWidth.value}px`,
}))

const toolbarStatus = computed(() => {
  const csv = m.value.vueUi.csvViewer
  if (isReading.value) return csv.uploadPending
  if (status.value !== "ready") return ""
  return formatOfficeMessage(csv.status, {
    rows: parsed.value.rows.length,
    columns: parsed.value.columnCount,
  })
})

const canSearch = computed(() => status.value === "ready" && !isReading.value)
const canDownload = computed(
  () => status.value === "ready" && !isReading.value,
)

const errorMessage = computed(() => {
  const error = errorState.value
  if (!error) return ""
  const csv = m.value.vueUi.csvViewer
  switch (error.code) {
    case "INVALID_DELIMITER":
      return csv.invalidDelimiter
    case "UNTERMINATED_QUOTE":
      return csv.unterminatedQuote
    case "FILE_TOO_LARGE":
      return formatOfficeMessage(csv.fileTooLarge, {
        actual: error.actual ?? 0,
        allowed: error.allowed ?? 0,
      })
    case "READ_ERROR":
      return csv.readError
    case "FETCH_ERROR":
      return csv.loadError
    default:
      return csv.parseError
  }
})

const actionItems = computed<OfficeMenuEntry[]>(() => {
  const csv = m.value.vueUi.csvViewer
  return [
    {
      id: "download",
      label: csv.download,
      icon: "download",
      disabled: !canDownload.value,
      testId: "csv-download",
    },
    {
      id: "upload",
      label: csv.uploadShort,
      icon: "upload",
      disabled: isReading.value,
      testId: "csv-upload",
    },
  ]
})

function fileDisplayName(name: string | null | undefined): string | null {
  return name && name.trim() ? name : null
}

function delimiterForFile(name: string | null): string {
  if (props.delimiter !== "auto") return props.delimiter
  return name?.toLowerCase().endsWith(".tsv") ? "\t" : "auto"
}

function resetSearch(): void {
  if (searchTimer !== undefined) {
    clearTimeout(searchTimer)
    searchTimer = undefined
  }
  searchRequestId.value += 1
  searchQuery.value = ""
  searchResults.value = []
  activeResultIndex.value = 0
  searchStatus.value = "idle"
  selectedCell.value = null
}

function resetState(name: string | null = null): void {
  parsed.value = emptyParse()
  errorState.value = null
  fileName.value = fileDisplayName(name)
  selectedCell.value = null
  scrollTop.value = 0
  resetSearch()
}

function applyText(text: string, name: string | null): void {
  const result = parseCsvText(text, { delimiter: delimiterForFile(name) })
  fileName.value = fileDisplayName(name)
  selectedCell.value = null
  scrollTop.value = 0
  resetSearch()
  if (result.error) {
    parsed.value = emptyParse()
    errorState.value = {
      code: result.error.code,
      detail: result.error.detail,
    }
    emit("error", errorState.value)
    return
  }
  parsed.value = result
  errorState.value = null
  if (result.rows.length > 0) {
    emit("load", {
      fileName: fileName.value,
      headers: [...result.headers],
      rows: result.rows.map((row) => [...row]),
      delimiter: result.delimiter,
    })
  }
}

function failWith(error: CsvViewerError, name: string | null = null): void {
  parsed.value = emptyParse()
  fileName.value = fileDisplayName(name)
  resetSearch()
  errorState.value = error
  emit("error", error)
}

function sourceName(url: string): string | null {
  if (!url) return null
  try {
    const parsedUrl = new URL(url, "http://localhost")
    const last = parsedUrl.pathname.split("/").filter(Boolean).pop()
    return last ? decodeURIComponent(last) : null
  } catch {
    return null
  }
}

async function readBlob(
  blob: Blob,
  name: string | null,
  remember: boolean,
): Promise<void> {
  const requestId = ++readRequestId
  isReading.value = true
  errorState.value = null
  try {
    if (blob.size > props.maxFileSize) {
      const error: CsvViewerError = {
        code: "FILE_TOO_LARGE",
        actual: blob.size,
        allowed: props.maxFileSize,
      }
      failWith(error, name)
      return
    }
    const text = decodeCsvBuffer(await blob.arrayBuffer())
    if (requestId !== readRequestId) return
    if (remember) lastLoad.value = { kind: "file", file: blob, fileName: name }
    applyText(text, name)
    emit("update:data", text)
  } catch (error) {
    if (requestId !== readRequestId) return
    failWith(
      {
        code: "READ_ERROR",
        detail: error instanceof Error ? error.message : undefined,
      },
      name,
    )
  } finally {
    if (requestId === readRequestId) isReading.value = false
  }
}

async function loadSource(url: string): Promise<void> {
  if (typeof fetch !== "function") {
    failWith({ code: "FETCH_ERROR", detail: "fetch is unavailable" })
    return
  }
  const requestId = ++readRequestId
  activeFetchController?.abort()
  const controller = new AbortController()
  activeFetchController = controller
  isReading.value = true
  errorState.value = null
  try {
    const response = await fetch(url, { signal: controller.signal })
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`)
    }
    const buffer = await response.arrayBuffer()
    if (requestId !== readRequestId) return
    const text = decodeCsvBuffer(buffer)
    lastLoad.value = { kind: "src", src: url }
    applyText(text, sourceName(url))
    emit("update:data", text)
  } catch (error) {
    if (requestId !== readRequestId || controller.signal.aborted) return
    failWith({
      code: "FETCH_ERROR",
      detail: error instanceof Error ? error.message : undefined,
    })
  } finally {
    if (requestId === readRequestId) {
      isReading.value = false
      if (activeFetchController === controller) activeFetchController = null
    }
  }
}

function loadProps(): void {
  readRequestId += 1
  activeFetchController?.abort()
  activeFetchController = null
  if (props.data !== undefined) {
    isReading.value = false
    lastLoad.value = null
    applyText(props.data, props.fileName ?? null)
    return
  }
  if (props.file) {
    lastLoad.value = { kind: "file", file: props.file, fileName: props.fileName ?? null }
    void readBlob(props.file, props.fileName ?? null, false)
    return
  }
  if (props.src) {
    lastLoad.value = { kind: "src", src: props.src }
    void loadSource(props.src)
    return
  }
  isReading.value = false
  resetState(props.fileName ?? null)
}

watch(
  () => [props.data, props.file, props.src] as const,
  () => {
    loadProps()
  },
  { immediate: true },
)

watch(
  () => props.search,
  (enabled) => {
    if (!enabled) resetSearch()
  },
)

watch(
  () => props.fileName,
  (value) => {
    if (value !== undefined) fileName.value = fileDisplayName(value)
  },
)

watch(
  () => props.delimiter,
  () => loadProps(),
)

function openFilePicker(): void {
  if (isReading.value) return
  fileInputRef.value?.click()
}

function handleInputChange(event: Event): void {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  if (!file) return
  input.value = ""
  lastLoad.value = { kind: "file", file, fileName: file.name }
  void readBlob(file, file.name, true)
}

function retryLastLoad(): void {
  if (!lastLoad.value) {
    openFilePicker()
    return
  }
  if (lastLoad.value.kind === "src") {
    void loadSource(lastLoad.value.src)
  } else {
    void readBlob(lastLoad.value.file, lastLoad.value.fileName, false)
  }
}

function reload(): void {
  if (lastLoad.value?.kind === "src") {
    void loadSource(lastLoad.value.src)
    return
  }
  if (lastLoad.value?.kind === "file") {
    void readBlob(lastLoad.value.file, lastLoad.value.fileName, false)
    return
  }
  loadProps()
}

function download(): void {
  if (!canDownload.value) return
  const text = serializeCsv(parsed.value.headers, parsed.value.rows, {
    delimiter: parsed.value.delimiter,
  })
  const name = ensureCsvExtension(fileName.value ?? "data.csv")
  emit("download", { fileName: name, text })
  if (
    typeof document === "undefined"
    || typeof URL === "undefined"
    || typeof URL.createObjectURL !== "function"
  ) return
  const blob = new Blob([text], { type: "text/csv;charset=utf-8" })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement("a")
  anchor.href = url
  anchor.download = name
  anchor.rel = "noopener"
  document.body.append(anchor)
  anchor.click()
  anchor.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 0)
}

function ensureCsvExtension(name: string): string {
  const lower = name.toLowerCase()
  return lower.endsWith(".csv") || lower.endsWith(".tsv") ? name : `${name}.csv`
}

function handleActionSelect(entry: OfficeMenuEntry): void {
  if (entry.type === "separator" || entry.type === "label" || entry.type === "group") {
    return
  }
  if (!("id" in entry)) return
  if (entry.id === "download") download()
  else if (entry.id === "upload") openFilePicker()
}

function setZoom(value: number): void {
  const finite = Number.isFinite(value) ? value : BASE_ZOOM
  const clamped = Math.min(
    ZOOM_OPTIONS[ZOOM_OPTIONS.length - 1] as number,
    Math.max(ZOOM_OPTIONS[0] as number, Math.round(finite)),
  )
  zoom.value = clamped
}

function zoomIn(): void {
  setZoom(zoom.value + 25)
}

function zoomOut(): void {
  setZoom(zoom.value - 25)
}

function columnTitle(index: number): string {
  const raw = parsed.value.headers[index]?.trim()
  if (raw) return raw
  const csv = m.value.vueUi.csvViewer
  if (index >= parsed.value.sourceHeaderCount) {
    return formatOfficeMessage(csv.extraColumn, {
      index: index - parsed.value.sourceHeaderCount + 1,
    })
  }
  return formatOfficeMessage(csv.column, { index: index + 1 })
}

function valueAt(cells: readonly string[], column: number): string {
  return cells[column] ?? ""
}

function highlightSegments(
  value: string,
  text: string,
): Array<{ text: string; match: boolean }> {
  const query = text.trim().toLowerCase()
  if (!query) return [{ text: value, match: false }]
  const haystack = value.toLowerCase()
  const segments: Array<{ text: string; match: boolean }> = []
  let start = 0
  let found = haystack.indexOf(query, start)
  while (found >= 0) {
    if (found > start) {
      segments.push({ text: value.slice(start, found), match: false })
    }
    segments.push({
      text: value.slice(found, found + query.length),
      match: true,
    })
    start = found + query.length
    found = haystack.indexOf(query, start)
  }
  if (start < value.length) {
    segments.push({ text: value.slice(start), match: false })
  }
  return segments.length ? segments : [{ text: value, match: false }]
}

function selectCell(row: number, col: number): void {
  selectedCell.value = { row, col }
}

function isSelected(row: number, col: number): boolean {
  return selectedCell.value?.row === row && selectedCell.value?.col === col
}

function handleScroll(event: Event): void {
  const target = event.target as HTMLElement | null
  if (!target) return
  scrollTop.value = target.scrollTop ?? 0
  if (target.clientHeight > 0) viewportHeight.value = target.clientHeight
}

function measureViewport(): void {
  const element = viewportRef.value
  if (element && element.clientHeight > 0) {
    viewportHeight.value = element.clientHeight
  }
}

async function revealCell(row: number, col: number): Promise<void> {
  const viewport = viewportRef.value
  if (!viewport) return
  const height = viewport.clientHeight || viewportHeight.value
  const top = row * rowHeight.value
  const currentTop = viewport.scrollTop ?? scrollTop.value
  if (top < currentTop || top + rowHeight.value > currentTop + height) {
    const next = Math.max(
      0,
      top - Math.floor(height / 2) + Math.floor(rowHeight.value / 2),
    )
    viewport.scrollTop = next
    scrollTop.value = next
  }
  await nextTick()
  const cell = viewport.querySelector?.(
    `[data-csv-cell="${row}:${col}"]`,
  ) as HTMLElement | null
  cell?.scrollIntoView?.({ block: "nearest", inline: "nearest" })
}

function activateResult(index: number): void {
  if (!searchResults.value.length) {
    selectedCell.value = null
    return
  }
  const total = searchResults.value.length
  const next = ((index % total) + total) % total
  activeResultIndex.value = next
  const result = searchResults.value[next]
  if (!result) return
  selectedCell.value = { row: result.row, col: result.col }
  void revealCell(result.row, result.col)
}

function clearSearch(): void {
  resetSearch()
}

function cancelPendingSearch(): void {
  if (searchTimer !== undefined) {
    clearTimeout(searchTimer)
    searchTimer = undefined
  }
  searchRequestId.value += 1
}

function onSearchInput(value: string): void {
  searchQuery.value = value
  cancelPendingSearch()
  if (!value.trim()) {
    clearSearch()
    return
  }
  searchStatus.value = "searching"
  searchTimer = setTimeout(() => {
    searchTimer = undefined
    void runSearch(value)
  }, Math.max(0, props.searchDebounce))
}

async function runSearch(rawQuery: string): Promise<void> {
  const query = rawQuery.trim()
  if (!query) {
    clearSearch()
    return
  }
  cancelPendingSearch()
  searchQuery.value = query
  const requestId = searchRequestId.value
  searchStatus.value = "searching"
  const results = await findCsvMatches(
    parsed.value.headers,
    parsed.value.rows,
    query,
    {
      shouldCancel: () => requestId !== searchRequestId.value,
    },
  )
  if (requestId !== searchRequestId.value) return
  searchResults.value = results
  activeResultIndex.value = 0
  searchStatus.value = "ready"
  activateResult(0)
}

function nextResult(): void {
  if (!searchResults.value.length) return
  activateResult(activeResultIndex.value + 1)
}

function previousResult(): void {
  if (!searchResults.value.length) return
  activateResult(activeResultIndex.value - 1)
}

function onSearchSubmit(value: string): void {
  void runSearch(value)
}

function getState(): CsvViewerState {
  return {
    status: status.value,
    fileName: fileName.value,
    headers: [...parsed.value.headers],
    rows: parsed.value.rows.map((row) => [...row]),
    delimiter: parsed.value.delimiter,
    zoom: zoom.value,
    searchQuery: searchQuery.value,
    searchResults: searchResults.value.map((result) => ({ ...result })),
    activeResultIndex: activeResultIndex.value,
    error: errorState.value ? { ...errorState.value } : null,
  }
}

onMounted(() => {
  void nextTick(measureViewport)
  if (typeof ResizeObserver !== "undefined" && viewportRef.value) {
    resizeObserver = new ResizeObserver(measureViewport)
    resizeObserver.observe(viewportRef.value)
  }
})

onBeforeUnmount(() => {
  activeFetchController?.abort()
  activeFetchController = null
  readRequestId += 1
  cancelPendingSearch()
  resizeObserver?.disconnect()
  resizeObserver = null
})

defineExpose({
  reload,
  retry: retryLastLoad,
  download,
  openFilePicker,
  runSearch,
  clearSearch,
  nextResult,
  previousResult,
  setZoom,
  zoomIn,
  zoomOut,
  selectCell,
  getState,
})
</script>

<template>
  <div
    :class="rootClasses"
    :style="rootStyle"
    :data-testid="'csv-viewer'"
    :data-status="status"
    :data-file-name="fileName ?? undefined"
    :data-row-count="parsed.rows.length"
    :data-column-count="parsed.columnCount"
  >
    <OfficeToolbar>
      <template #leading>
        <span
          class="office-csv-viewer__meta"
          role="status"
          aria-live="polite"
          data-testid="csv-meta"
        >{{ toolbarStatus }}</span>
      </template>
      <template #trailing>
        <OfficeToolbarGroup>
          <OfficeZoomControl
            :model-value="zoom"
            :min="75"
            :max="200"
            :options="ZOOM_OPTIONS"
            :show-fit="false"
            @update:model-value="setZoom"
          />
        </OfficeToolbarGroup>
        <template v-if="props.search">
          <OfficeToolbarSeparator />
          <OfficeSearchPopover
            :model-value="searchQuery"
            :status="searchStatus"
            :result-count="searchResults.length"
            :active-index="activeResultIndex"
            :disabled="!canSearch"
            :show-options="false"
            :trigger-label="m.vueUi.csvViewer.searchLabel"
            :placeholder="m.vueUi.csvViewer.searchPlaceholder"
            test-id-prefix="csv-search"
            @update:model-value="onSearchInput"
            @submit="onSearchSubmit"
            @next="nextResult"
            @previous="previousResult"
          />
        </template>
        <OfficeToolbarSeparator />
        <input
          ref="fileInputRef"
          type="file"
          accept=".csv,.tsv,text/csv,text/tab-separated-values"
          class="office-csv-viewer__file-input"
          :aria-label="m.vueUi.csvViewer.upload"
          data-testid="csv-file-input"
          @change="handleInputChange"
        >
        <OfficeDropdownMenu
          :items="actionItems"
          :disabled="isReading"
          :menu-label="m.vueUi.csvViewer.actions"
          content-test-id="csv-actions-menu"
          @select="handleActionSelect"
        >
          <template #trigger>
            <OfficeIconButton
              :label="m.vueUi.csvViewer.actions"
              icon="more-horizontal"
              size="icon-sm"
              :disabled="isReading"
              data-testid="csv-actions-trigger"
            />
          </template>
        </OfficeDropdownMenu>
      </template>
    </OfficeToolbar>

    <div class="office-csv-viewer__body" data-testid="csv-body">
      <OfficeLoadingState
        v-if="status === 'loading'"
        class="office-csv-viewer__state"
        :title="m.vueUi.csvViewer.loadingTitle"
        :description="m.vueUi.csvViewer.loadingDescription"
        data-testid="csv-loading"
      />
      <OfficeErrorState
        v-else-if="status === 'error'"
        class="office-csv-viewer__state"
        :title="m.vueUi.csvViewer.errorTitle"
        :description="errorMessage"
        :show-retry="false"
        data-testid="csv-error"
        @retry="retryLastLoad"
      >
        <template #actions>
          <OfficeButton
            v-if="lastLoad"
            variant="outline"
            size="sm"
            data-testid="csv-error-retry"
            @click="retryLastLoad"
          >
            {{ m.common.chrome.retry }}
          </OfficeButton>
          <OfficeButton
            variant="outline"
            size="sm"
            data-testid="csv-error-upload"
            @click="openFilePicker"
          >
            {{ m.vueUi.csvViewer.upload }}
          </OfficeButton>
        </template>
      </OfficeErrorState>
      <OfficeEmptyState
        v-else-if="status === 'empty'"
        class="office-csv-viewer__state"
        :title="m.vueUi.csvViewer.emptyTitle"
        :description="m.vueUi.csvViewer.emptyDescription"
        data-testid="csv-empty"
      >
        <template #actions>
          <OfficeButton
            variant="outline"
            size="sm"
            data-testid="csv-empty-upload"
            @click="openFilePicker"
          >
            {{ m.vueUi.csvViewer.upload }}
          </OfficeButton>
        </template>
      </OfficeEmptyState>
      <div
        v-else
        ref="viewportRef"
        class="office-csv-viewer__viewport"
        data-testid="csv-viewport"
        @scroll="handleScroll"
      >
        <table
          class="office-csv-viewer__table"
          role="grid"
          :aria-label="m.vueUi.csvViewer.label"
          :aria-rowcount="parsed.rows.length + 1"
          :aria-colcount="parsed.columnCount + 1"
          :style="tableStyle"
          data-testid="csv-grid"
        >
          <colgroup>
            <col :style="{ width: `${rowHeaderWidth}px` }">
            <col
              v-for="(width, index) in columnWidths"
              :key="index"
              :style="{ width: `${width}px` }"
            >
          </colgroup>
          <thead>
            <tr
              role="row"
              :style="{ height: `${headerHeight}px` }"
            >
              <th
                scope="col"
                role="columnheader"
                class="office-csv-viewer__row-header"
                aria-hidden="true"
              >#</th>
              <th
                v-for="(_, index) in parsed.headers"
                :key="index"
                scope="col"
                role="columnheader"
                class="office-csv-viewer__header-cell"
                :title="columnTitle(index)"
              >{{ columnTitle(index) }}</th>
            </tr>
          </thead>
          <tbody>
            <tr
              v-if="visibleRange.padTop > 0"
              aria-hidden="true"
              class="office-csv-viewer__spacer"
            >
              <td :colspan="parsed.columnCount + 1" :style="{ height: `${visibleRange.padTop}px` }" />
            </tr>
            <tr
              v-for="row in renderedRows"
              :key="row.index"
              role="row"
              :style="{ height: `${rowHeight}px` }"
              :data-csv-row="row.index"
            >
              <td
                role="rowheader"
                class="office-csv-viewer__row-header"
              >{{ row.index + 1 }}</td>
              <td
                v-for="(_, columnIndex) in parsed.headers"
                :key="columnIndex"
                role="gridcell"
                class="office-csv-viewer__cell"
                :class="isSelected(row.index, columnIndex) && 'office-csv-viewer__cell--selected'"
                :data-csv-cell="`${row.index}:${columnIndex}`"
                :data-csv-row="row.index"
                :data-csv-col="columnIndex"
                :aria-selected="isSelected(row.index, columnIndex)"
                :title="valueAt(row.cells, columnIndex)"
                @click="selectCell(row.index, columnIndex)"
              >
                <template
                  v-for="(segment, segmentIndex) in highlightSegments(
                    valueAt(row.cells, columnIndex),
                    searchQuery,
                  )"
                  :key="segmentIndex"
                >
                  <mark
                    v-if="segment.match"
                    class="office-csv-viewer__match"
                  >{{ segment.text }}</mark>
                  <template v-else>{{ segment.text }}</template>
                </template>
              </td>
            </tr>
            <tr
              v-if="visibleRange.padBottom > 0"
              aria-hidden="true"
              class="office-csv-viewer__spacer"
            >
              <td :colspan="parsed.columnCount + 1" :style="{ height: `${visibleRange.padBottom}px` }" />
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>
</template>
