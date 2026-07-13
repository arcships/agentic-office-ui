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
          <OfficeFindBar
            ref="findBarRef"
            :query="searchQuery"
            :status="searchState.status"
            :active-index="searchState.activeIndex"
            :result-count="results.length"
            input-test-id="xlsx-search-input"
            placeholder="Find values"
            @update:query="onSearchQuery"
            @previous="selectRelativeResult(-1)"
            @next="selectRelativeResult(1)"
            @close="closeSearch"
          />
          <div v-if="results.length" class="xlsx-toolbar__search-results">
            <button
              v-for="(result, index) in results"
              :key="`${result.workbookSheetIndex}:${result.address}:${result.start}:${index}`"
              type="button"
              class="xlsx-toolbar__search-result"
              :class="{ active: index === searchState.activeIndex }"
              @click="selectResult(index)"
            >
              <strong>{{ result.sheetName }}!{{ result.address }}</strong>
              <span>
                {{ result.text.slice(0, result.start) }}<mark>{{ result.text.slice(result.start, result.end) }}</mark>{{ result.text.slice(result.end) }}
              </span>
            </button>
          </div>
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
import { computed, nextTick, ref, watch } from "vue";
import type { XlsxViewerController } from "@arcships/xlsx-core";
import { OfficeFindBar } from "@arcships/vue-ui";
import { useXlsxSearch, type XlsxSurfaceSearch } from "../composables/useXlsxSearch";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
  search?: XlsxSurfaceSearch;
  showUpload?: boolean;
}>();

const emit = defineEmits<{
  searchClose: [];
  upload: [];
}>();

const zoomPresets = [50, 75, 100, 125, 150, 200];
const searchOpen = ref(false);
const findBarRef = ref<InstanceType<typeof OfficeFindBar> | null>(null);
const searchQuery = ref("");
const ownedSearch = useXlsxSearch(() => props.controller);
const effectiveSearch = computed(() => props.search ?? ownedSearch);
const searchState = computed(() => effectiveSearch.value.searchState.value);
const results = computed(() => searchState.value.matches);

const displayFileName = computed(() => props.controller.displayFileName);
const zoomPercent = computed(() => Math.round(props.controller.zoomScale));
function onZoomChange(event: Event): void {
  const value = Number((event.target as HTMLSelectElement).value);
  if (Number.isFinite(value)) props.controller.setZoomScale(value);
}

function toggleSearch(): void {
  if (searchOpen.value) closeSearch();
  else openSearch();
}

function openSearch(): void {
  searchOpen.value = true;
  void nextTick(() => findBarRef.value?.focus());
}

function closeSearch(): void {
  searchOpen.value = false;
  searchQuery.value = "";
  effectiveSearch.value.clearSearch();
  emit("searchClose");
}

function onSearchQuery(query: string): void {
  searchQuery.value = query;
  void effectiveSearch.value.search(query).catch(() => undefined);
}

function selectRelativeResult(direction: 1 | -1): void {
  if (direction === 1) void effectiveSearch.value.searchNext().catch(() => undefined);
  else void effectiveSearch.value.searchPrevious().catch(() => undefined);
}

function selectResult(index: number): void {
  void effectiveSearch.value.activateSearchMatch(index).catch(() => undefined);
}

watch(() => searchState.value.query, (query) => {
  if (query !== searchQuery.value.trim()) searchQuery.value = query;
});

defineExpose({ openSearch });
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
.xlsx-toolbar__search-results { max-height: 240px; overflow: auto; }
.xlsx-toolbar__search-result { align-items: center; background: transparent; border: 0; border-radius: 6px; cursor: pointer; display: flex; font: inherit; gap: 8px; padding: 7px 8px; text-align: left; width: 100%; }
.xlsx-toolbar__search-result:hover,
.xlsx-toolbar__search-result.active { background: #f4f4f5; }
.xlsx-toolbar__search-result strong { flex: 0 0 96px; font-size: 11px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.xlsx-toolbar__search-result span { font-size: 11px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.xlsx-toolbar__search-result mark { background: #fde047; color: inherit; padding: 0; }

@media (max-width: 620px) {
  .xlsx-toolbar { align-items: stretch; flex-direction: column; gap: 5px; }
  .xlsx-toolbar__filename { flex: 1; max-width: none; }
  .xlsx-toolbar__actions { justify-content: flex-end; overflow-x: auto; }
}
</style>
