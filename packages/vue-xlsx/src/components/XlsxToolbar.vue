<template>
  <header class="xlsx-toolbar" :class="{ 'xlsx-toolbar--dark': isDark }" data-testid="xlsx-toolbar">
    <div class="xlsx-toolbar__identity">
      <svg class="xlsx-toolbar__file-icon" viewBox="0 0 24 24" aria-hidden="true">
        <path d="M6 2.8h8l4 4V21H6zM14 2.8V7h4" />
        <path d="m9 11 6 6M15 11l-6 6" />
      </svg>
      <span class="xlsx-toolbar__filename" :title="displayFileName">{{ displayFileName }}</span>
    </div>

    <div class="xlsx-toolbar__actions">
      <div class="xlsx-toolbar__zoom" aria-label="Workbook zoom">
        <button class="xlsx-toolbar__icon-button" :disabled="!controller.canZoomOut" title="Zoom out" aria-label="Zoom out" @click="controller.zoomOut()">
          <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9" /><path d="M8 12h8" /></svg>
        </button>
        <select class="xlsx-toolbar__zoom-select" aria-label="Zoom level" :value="zoomPercent" @change="onZoomChange">
          <option v-for="z in zoomPresets" :key="z" :value="z">{{ z }}%</option>
        </select>
        <button class="xlsx-toolbar__icon-button" :disabled="!controller.canZoomIn" title="Zoom in" aria-label="Zoom in" @click="controller.zoomIn()">
          <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9" /><path d="M8 12h8M12 8v8" /></svg>
        </button>
      </div>

      <span class="xlsx-toolbar__separator" />

      <div class="xlsx-toolbar__search-wrap">
        <button
          class="xlsx-toolbar__icon-button"
          data-testid="xlsx-search-toggle"
          :aria-expanded="searchOpen"
          aria-label="Search workbook"
          title="Search workbook"
          @click="toggleSearch"
        >
          <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="10.5" cy="10.5" r="6.5" /><path d="m15.5 15.5 5 5" /></svg>
        </button>
        <div v-if="searchOpen" class="xlsx-toolbar__search-panel" data-testid="xlsx-search-panel">
          <label>
            <span>Search workbook</span>
            <input
              ref="searchInputRef"
              v-model="searchQuery"
              data-testid="xlsx-search-input"
              placeholder="Find values"
              @input="scheduleSearch"
              @keydown.enter.prevent="selectRelativeResult(1)"
              @keydown.esc.prevent="closeSearch"
            />
          </label>
          <div class="xlsx-toolbar__search-summary">
            <span>{{ searchStatus }}</span>
            <div>
              <button :disabled="results.length === 0" aria-label="Previous result" @click="selectRelativeResult(-1)">↑</button>
              <button :disabled="results.length === 0" aria-label="Next result" @click="selectRelativeResult(1)">↓</button>
            </div>
          </div>
          <button
            v-for="(result, index) in results.slice(0, 8)"
            :key="`${result.workbookSheetIndex}:${result.cell.row}:${result.cell.col}`"
            class="xlsx-toolbar__search-result"
            :class="{ active: index === activeResultIndex }"
            @click="selectResult(index)"
          >
            <strong>{{ result.sheetName }}!{{ result.address }}</strong><span>{{ result.value }}</span>
          </button>
        </div>
      </div>

      <button v-if="showUpload" class="xlsx-toolbar__icon-button" data-testid="xlsx-toolbar-upload" title="Upload XLSX" aria-label="Upload XLSX" @click="emit('upload')">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 16V4M7.5 8.5 12 4l4.5 4.5M5 14v6h14v-6" /></svg>
      </button>
      <button v-if="controller.canDownload" class="xlsx-toolbar__icon-button" data-testid="xlsx-toolbar-download" title="Download source" aria-label="Download source" @click="controller.download()">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v12M7.5 11.5 12 16l4.5-4.5M5 19h14" /></svg>
      </button>
      <button v-if="controller.canExport" class="xlsx-toolbar__icon-button" data-testid="xlsx-toolbar-export" title="Export XLSX" aria-label="Export XLSX" @click="controller.exportXlsx()">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 3h12l2 2v16H5zM8 3v6h8V3M8 21v-7h8v7" /></svg>
      </button>
    </div>
  </header>
</template>

<script setup lang="ts">
import { computed, nextTick, onUnmounted, ref } from "vue";
import type { XlsxCellAddress, XlsxViewerController } from "@arcships/xlsx-core";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
  showUpload?: boolean;
}>();

const emit = defineEmits<{
  upload: [];
}>();

interface SearchResult {
  address: string;
  cell: XlsxCellAddress;
  sheetIndex: number;
  sheetName: string;
  value: string;
  workbookSheetIndex: number;
}

const zoomPresets = [50, 75, 100, 125, 150, 200];
const searchOpen = ref(false);
const searchInputRef = ref<HTMLInputElement | null>(null);
const searchQuery = ref("");
const results = ref<SearchResult[]>([]);
const activeResultIndex = ref(-1);
const isSearching = ref(false);
let searchTimer: ReturnType<typeof setTimeout> | null = null;
let searchRequest = 0;

const displayFileName = computed(() => props.controller.displayFileName);
const zoomPercent = computed(() => Math.round(props.controller.zoomScale));
const searchStatus = computed(() => {
  if (isSearching.value) return "Searching…";
  if (!searchQuery.value.trim()) return "Type to search";
  if (!results.value.length) return "No results";
  return `${Math.max(1, activeResultIndex.value + 1)} of ${results.value.length}`;
});

function cellAddress(cell: XlsxCellAddress): string {
  let col = "";
  let index = cell.col;
  while (index >= 0) {
    col = String.fromCharCode(65 + index % 26) + col;
    index = Math.floor(index / 26) - 1;
  }
  return `${col}${cell.row + 1}`;
}

function onZoomChange(event: Event): void {
  const value = Number((event.target as HTMLSelectElement).value);
  if (Number.isFinite(value)) props.controller.setZoomScale(value);
}

function toggleSearch(): void {
  searchOpen.value = !searchOpen.value;
  if (searchOpen.value) void nextTick(() => searchInputRef.value?.focus());
}

function closeSearch(): void {
  searchOpen.value = false;
}

function scheduleSearch(): void {
  if (searchTimer) clearTimeout(searchTimer);
  searchTimer = setTimeout(() => { void runSearch(); }, 180);
}

async function runSearch(): Promise<void> {
  const query = searchQuery.value.trim().toLocaleLowerCase();
  const request = ++searchRequest;
  results.value = [];
  activeResultIndex.value = -1;
  if (!query || !props.controller.sheets.length) return;
  isSearching.value = true;
  const matches: SearchResult[] = [];
  try {
    for (const [sheetIndex, sheet] of props.controller.sheets.entries()) {
      if (matches.length >= 200) break;
      if (props.controller.isWorkerBacked && props.controller.getRowsBatchAsync) {
        const rowLimit = Math.max(0, sheet.maxUsedRow + 1);
        for (let start = 0; start < rowLimit && matches.length < 200; start += 128) {
          const rows = await props.controller.getRowsBatchAsync(sheet.workbookSheetIndex, start, 128);
          if (request !== searchRequest) return;
          for (const rowEntry of rows ?? []) {
            if (!rowEntry || typeof rowEntry !== "object") continue;
            const row = Number(Reflect.get(rowEntry, "index"));
            const cells = Reflect.get(rowEntry, "cells");
            if (!Number.isInteger(row) || !Array.isArray(cells)) continue;
            for (const entry of cells) {
              const col = Number(Reflect.get(entry, "col"));
              const value = String(Reflect.get(entry, "value") ?? "");
              const formula = String(Reflect.get(entry, "formula") ?? "");
              if (Number.isInteger(col) && `${value}\n${formula}`.toLocaleLowerCase().includes(query)) {
                const cell = { row, col };
                matches.push({
                  address: cellAddress(cell),
                  cell,
                  sheetIndex,
                  sheetName: sheet.name,
                  value: value || formula,
                  workbookSheetIndex: sheet.workbookSheetIndex,
                });
                if (matches.length >= 200) break;
              }
            }
          }
        }
        continue;
      }

      const worksheet = props.controller.workbook?.getSheet(sheet.workbookSheetIndex);
      if (!worksheet) continue;
      try {
        const rows = sheet.visibleRows.filter((row) => row >= sheet.minUsedRow && row <= sheet.maxUsedRow);
        const columns = sheet.visibleCols.filter((col) => col >= sheet.minUsedCol && col <= sheet.maxUsedCol);
        for (const row of rows) {
          for (const col of columns) {
            const formatted = worksheet.getFormattedValueAt(row, col) ?? "";
            const formula = worksheet.getFormulaAt(row, col) ?? "";
            let value = formatted;
            if (!value) {
              const calculated = worksheet.getCalculatedValueAt(row, col);
              try {
                value = calculated.is_empty ? "" : calculated.toString();
              } finally {
                calculated.free();
              }
            }
            if (`${value}\n${formula}`.toLocaleLowerCase().includes(query)) {
              const cell = { row, col };
              matches.push({
                address: cellAddress(cell),
                cell,
                sheetIndex,
                sheetName: sheet.name,
                value: value || formula,
                workbookSheetIndex: sheet.workbookSheetIndex,
              });
              if (matches.length >= 200) break;
            }
          }
          if (matches.length >= 200) break;
        }
      } finally {
        worksheet.free();
      }
    }
    if (request !== searchRequest) return;
    results.value = matches;
    if (matches.length) selectResult(0);
  } finally {
    if (request === searchRequest) isSearching.value = false;
  }
}

function selectResult(index: number): void {
  const result = results.value[index];
  if (!result) return;
  activeResultIndex.value = index;
  if (props.controller.activeSheetIndex !== result.sheetIndex) {
    props.controller.setActiveSheetIndex(result.sheetIndex);
  }
  props.controller.selectCell(result.cell);
}

function selectRelativeResult(direction: 1 | -1): void {
  if (!results.value.length) {
    void runSearch();
    return;
  }
  const next = (activeResultIndex.value + direction + results.value.length) % results.value.length;
  selectResult(next);
}

onUnmounted(() => {
  if (searchTimer) clearTimeout(searchTimer);
  searchRequest += 1;
});
</script>

<style scoped>
.xlsx-toolbar {
  align-items: center;
  background: #fff;
  border-bottom: 1px solid #e4e4e7;
  color: #18181b;
  display: flex;
  flex: 0 0 auto;
  gap: 12px;
  justify-content: space-between;
  min-height: 48px;
  padding: 7px 12px;
  user-select: none;
}
.xlsx-toolbar--dark { background: #18181b; border-color: #3f3f46; color: #f4f4f5; }
.xlsx-toolbar__identity,
.xlsx-toolbar__actions,
.xlsx-toolbar__zoom,
.xlsx-toolbar__search-summary,
.xlsx-toolbar__search-summary div { align-items: center; display: flex; }
.xlsx-toolbar__identity { gap: 8px; min-width: 0; }
.xlsx-toolbar__actions { gap: 4px; }
.xlsx-toolbar__file-icon,
.xlsx-toolbar__icon-button svg {
  fill: none;
  height: 17px;
  stroke: currentColor;
  stroke-linecap: round;
  stroke-linejoin: round;
  stroke-width: 1.7;
  width: 17px;
}
.xlsx-toolbar__file-icon { color: #15803d; flex: 0 0 auto; }
.xlsx-toolbar__filename { font-size: 13px; font-weight: 600; max-width: 260px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.xlsx-toolbar__icon-button {
  align-items: center;
  background: transparent;
  border: 1px solid transparent;
  border-radius: 7px;
  color: inherit;
  cursor: pointer;
  display: inline-flex;
  height: 32px;
  justify-content: center;
  padding: 0;
  width: 32px;
}
.xlsx-toolbar__icon-button:hover:not(:disabled) { background: rgba(113, 113, 122, .12); }
.xlsx-toolbar__icon-button:disabled { cursor: default; opacity: .35; }
.xlsx-toolbar__icon-button:focus-visible,
.xlsx-toolbar__zoom-select:focus-visible,
.xlsx-toolbar__search-panel input:focus-visible { outline: 2px solid #2563eb; outline-offset: 2px; }
.xlsx-toolbar__zoom { gap: 3px; }
.xlsx-toolbar__zoom-select {
  background: inherit;
  border: 1px solid #d4d4d8;
  border-radius: 7px;
  color: inherit;
  font: inherit;
  font-size: 12px;
  height: 32px;
  min-width: 76px;
  padding: 0 8px;
}
.xlsx-toolbar__separator { background: #e4e4e7; height: 18px; margin: 0 3px; width: 1px; }
.xlsx-toolbar__search-wrap { position: relative; }
.xlsx-toolbar__search-panel {
  background: #fff;
  border: 1px solid #d4d4d8;
  border-radius: 10px;
  box-shadow: 0 12px 32px rgba(0, 0, 0, .14);
  color: #18181b;
  padding: 12px;
  position: absolute;
  right: 0;
  top: 38px;
  width: 288px;
  z-index: 40;
}
.xlsx-toolbar__search-panel label { display: flex; flex-direction: column; font-size: 11px; font-weight: 600; gap: 6px; }
.xlsx-toolbar__search-panel input { border: 1px solid #d4d4d8; border-radius: 7px; font: inherit; height: 34px; padding: 0 10px; }
.xlsx-toolbar__search-summary { color: #71717a; font-size: 11px; justify-content: space-between; padding: 8px 0 5px; }
.xlsx-toolbar__search-summary button { background: transparent; border: 1px solid #e4e4e7; cursor: pointer; height: 26px; width: 28px; }
.xlsx-toolbar__search-result { align-items: center; background: transparent; border: 0; border-radius: 6px; cursor: pointer; display: flex; font: inherit; gap: 8px; padding: 7px 8px; text-align: left; width: 100%; }
.xlsx-toolbar__search-result:hover,
.xlsx-toolbar__search-result.active { background: #f4f4f5; }
.xlsx-toolbar__search-result strong { flex: 0 0 96px; font-size: 11px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.xlsx-toolbar__search-result span { font-size: 11px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

@media (max-width: 620px) {
  .xlsx-toolbar { align-items: stretch; flex-direction: column; gap: 5px; }
  .xlsx-toolbar__filename { flex: 1; max-width: none; }
  .xlsx-toolbar__actions { justify-content: flex-end; overflow-x: auto; }
}
</style>
