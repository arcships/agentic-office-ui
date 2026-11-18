<template>
  <header class="xlsx-toolbar" :class="{ 'xlsx-toolbar--dark': isDark }" data-testid="xlsx-toolbar">
    <div class="xlsx-toolbar__identity">
      <svg class="xlsx-toolbar__file-icon" viewBox="0 0 24 24" aria-hidden="true">
        <path d="M6 2.8h8l4 4V21H6zM14 2.8V7h4" />
        <path d="m9 11 6 6M15 11l-6 6" />
      </svg>
      <span class="xlsx-toolbar__filename" :title="displayFileName">{{ displayFileName }}</span>
      <span
        v-if="controller.lastExportWarning"
        class="xlsx-toolbar__export-warning"
        data-testid="xlsx-export-warning"
        role="status"
      >{{ controller.lastExportWarning }}</span>
    </div>

    <div class="xlsx-toolbar__actions">
      <div class="xlsx-toolbar__zoom" :aria-label="m.xlsx.toolbar.zoomGroup">
        <button class="xlsx-toolbar__icon-button" :disabled="!controller.canZoomOut" :title="m.common.zoom.out" :aria-label="m.common.zoom.out" @click="controller.zoomOut()">
          <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9" /><path d="M8 12h8" /></svg>
        </button>
        <select class="xlsx-toolbar__zoom-select" :aria-label="m.common.zoom.level" :value="zoomPercent" @change="onZoomChange">
          <option v-for="z in zoomPresets" :key="z" :value="z">{{ z }}%</option>
        </select>
        <button class="xlsx-toolbar__icon-button" :disabled="!controller.canZoomIn" :title="m.common.zoom.in" :aria-label="m.common.zoom.in" @click="controller.zoomIn()">
          <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9" /><path d="M8 12h8M12 8v8" /></svg>
        </button>
      </div>

      <span class="xlsx-toolbar__separator" />

      <div class="xlsx-toolbar__search-wrap">
        <button
          class="xlsx-toolbar__icon-button"
          data-testid="xlsx-search-toggle"
          :aria-expanded="searchOpen"
          :aria-label="m.xlsx.toolbar.search"
          :title="m.xlsx.toolbar.search"
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
            :placeholder="m.xlsx.toolbar.searchPlaceholder"
            :case-sensitive="searchCaseSensitive ?? false"
            :whole-word="searchWholeWord ?? false"
            :regex="searchRegex ?? false"
            :show-replace="!controller.readOnly"
            :replacement="replaceText"
            :replace-disabled="!canReplace"
            :replace-all-disabled="!canReplaceAll"
            :replace-status="replaceStatus"
            @update:query="onSearchQuery"
            @update:case-sensitive="onSearchOption('caseSensitive', $event)"
            @update:whole-word="onSearchOption('wholeWord', $event)"
            @update:regex="onSearchOption('regex', $event)"
            @update:replacement="replaceText = $event"
            @previous="selectRelativeResult(-1)"
            @next="selectRelativeResult(1)"
            @close="closeSearch"
            @replace="replaceActiveMatch"
            @replace-all="replaceAllMatches"
          />
          <div v-if="searchOpen && !controller.readOnly" class="xlsx-toolbar__replace-options">
            <label class="xlsx-toolbar__replace-option">
              <input
                type="checkbox"
                data-testid="xlsx-replace-whole-cell"
                :checked="replaceWholeCell"
                @change="replaceWholeCell = ($event.target as HTMLInputElement).checked"
              />
              <span>{{ m.common.search.wholeCell }}</span>
            </label>
            <label class="xlsx-toolbar__replace-option">
              <input
                type="checkbox"
                data-testid="xlsx-replace-formulas"
                :checked="replaceFormulas"
                @change="replaceFormulas = ($event.target as HTMLInputElement).checked"
              />
              <span>{{ m.common.search.formulas }}</span>
            </label>
          </div>
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

      <button v-if="showUpload" class="xlsx-toolbar__icon-button" data-testid="xlsx-toolbar-upload" :title="m.xlsx.toolbar.upload" :aria-label="m.xlsx.toolbar.upload" @click="emit('upload')">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 16V4M7.5 8.5 12 4l4.5 4.5M5 14v6h14v-6" /></svg>
      </button>
      <button
        v-if="controller.activeSheet"
        class="xlsx-toolbar__icon-button"
        :class="{ active: controller.isPageBreakPreview }"
        data-testid="xlsx-toolbar-page-breaks"
        :aria-pressed="controller.isPageBreakPreview"
        :title="m.xlsx.toolbar.pageBreaks"
        :aria-label="m.xlsx.toolbar.pageBreaks"
        @click="controller.setPageBreakPreview(!controller.isPageBreakPreview)"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 5h16M4 12h16M4 19h16" stroke-dasharray="2 2" /></svg>
      </button>
      <button
        v-if="controller.activeSheet"
        class="xlsx-toolbar__icon-button"
        data-testid="xlsx-toolbar-print-preview"
        :title="m.xlsx.toolbar.printPreview"
        :aria-label="m.xlsx.toolbar.printPreview"
        @click="controller.openPrintPreview()"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 9V3h12v6M6 18H4V9h16v9h-2M8 14h8v7H8z" /></svg>
      </button>
      <button v-if="controller.canDownload" class="xlsx-toolbar__icon-button" data-testid="xlsx-toolbar-download" :title="m.xlsx.toolbar.download" :aria-label="m.xlsx.toolbar.download" @click="controller.download()">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v12M7.5 11.5 12 16l4.5-4.5M5 19h14" /></svg>
      </button>
      <button v-if="controller.canExport" class="xlsx-toolbar__icon-button" data-testid="xlsx-toolbar-export" :title="m.xlsx.toolbar.export" :aria-label="m.xlsx.toolbar.export" @click="controller.exportXlsx()">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 3h12l2 2v16H5zM8 3v6h8V3M8 21v-7h8v7" /></svg>
      </button>
    </div>
  </header>
</template>

<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue";
import { formatOfficeMessage } from "@arcships/office-interaction";
import type { XlsxViewerController } from "@arcships/xlsx-core";
import { OfficeFindBar } from "@arcships/vue-ui";
import { useXlsxSearch, type XlsxSearchOptions, type XlsxSurfaceSearch } from "../composables/useXlsxSearch";
import { useXlsxMessages } from "../i18n";

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
const m = useXlsxMessages();
const searchOpen = ref(false);
const findBarRef = ref<InstanceType<typeof OfficeFindBar> | null>(null);
const searchQuery = ref("");
const searchCaseSensitive = ref<boolean | undefined>();
const searchWholeWord = ref<boolean | undefined>();
const searchRegex = ref<boolean | undefined>();
const searchOptions = computed<XlsxSearchOptions>(() => ({
  ...(searchCaseSensitive.value === undefined ? {} : { caseSensitive: searchCaseSensitive.value }),
  ...(searchWholeWord.value === undefined ? {} : { wholeWord: searchWholeWord.value }),
  ...(searchRegex.value === undefined ? {} : { regex: searchRegex.value }),
}));
const ownedSearch = useXlsxSearch(() => props.controller);
const effectiveSearch = computed(() => props.search ?? ownedSearch);
const searchState = computed(() => effectiveSearch.value.searchState.value);
const results = computed(() => searchState.value.matches);

const replaceText = ref("");
const replaceWholeCell = ref(false);
const replaceFormulas = ref(true);
const replaceStatus = ref("");
const activeMatch = computed(() => {
  const match = searchState.value.matches[searchState.value.activeIndex];
  return match;
});
const canReplace = computed(
  () => !props.controller.readOnly && Boolean(activeMatch.value),
);
const canReplaceAll = computed(
  () => !props.controller.readOnly && results.value.length > 0,
);

function replaceMessage(count: number): string {
  if (count <= 0) return m.value.common.search.noResults;
  return formatOfficeMessage(m.value.common.search.replaced, { count });
}

/** Re-run the query so the result list follows the cells the command rewrote. */
async function refreshAfterReplace(): Promise<void> {
  // `useXlsxSearch` drops its results as soon as the workbook revision moves,
  // so wait for that flush and then search again.
  await nextTick();
  if (!searchQuery.value.trim()) return;
  await effectiveSearch.value.search(searchQuery.value, searchOptions.value).catch(() => undefined);
}

async function replaceActiveMatch(): Promise<void> {
  const match = activeMatch.value;
  if (!match || props.controller.readOnly) return;
  const result = props.controller.replaceInSheet(searchQuery.value, replaceText.value, {
    matchCase: searchCaseSensitive.value === true,
    wholeCell: replaceWholeCell.value,
    formula: replaceFormulas.value,
    sheetIndex: match.sheetIndex,
    occurrence: { cell: match.cell, start: match.start, end: match.end },
  });
  replaceStatus.value = replaceMessage(result.replacedOccurrences);
  await refreshAfterReplace();
}

async function replaceAllMatches(): Promise<void> {
  if (props.controller.readOnly || !searchQuery.value.trim()) return;
  const result = props.controller.replaceInSheet(searchQuery.value, replaceText.value, {
    matchCase: searchCaseSensitive.value === true,
    wholeCell: replaceWholeCell.value,
    formula: replaceFormulas.value,
  });
  replaceStatus.value = replaceMessage(result.replacedOccurrences);
  await refreshAfterReplace();
}

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
  replaceText.value = "";
  replaceStatus.value = "";
  effectiveSearch.value.clearSearch();
  emit("searchClose");
}

function onSearchQuery(query: string): void {
  searchQuery.value = query;
  replaceStatus.value = "";
  void effectiveSearch.value.search(query, searchOptions.value).catch(() => undefined);
}

function onSearchOption(
  key: "caseSensitive" | "wholeWord" | "regex",
  value: boolean,
): void {
  if (key === "caseSensitive") searchCaseSensitive.value = value;
  else if (key === "wholeWord") searchWholeWord.value = value;
  else searchRegex.value = value;
  if (searchQuery.value.trim()) onSearchQuery(searchQuery.value);
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
.xlsx-toolbar__export-warning { color: #b45309; font-size: 11px; max-width: 320px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.xlsx-toolbar--dark .xlsx-toolbar__export-warning { color: #fbbf24; }
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
.xlsx-toolbar__icon-button.active { background: rgba(37, 99, 235, .14); border-color: #2563eb; color: #2563eb; }
.xlsx-toolbar--dark .xlsx-toolbar__icon-button.active { background: rgba(96, 165, 250, .2); border-color: #60a5fa; color: #93c5fd; }
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
.xlsx-toolbar__replace-options { display: flex; flex-wrap: wrap; gap: 10px; padding-top: 8px; }
.xlsx-toolbar__replace-option { align-items: center; display: flex; flex-direction: row; font-size: 11px; gap: 5px; }
.xlsx-toolbar__search-panel .xlsx-toolbar__replace-option { flex-direction: row; }
.xlsx-toolbar__replace-option input { height: 14px; width: 14px; }

@container xlsx-viewer (max-width: 620px) {
  .xlsx-toolbar { align-items: stretch; flex-direction: column; gap: 5px; }
  .xlsx-toolbar__filename { flex: 1; max-width: none; }
  .xlsx-toolbar__actions { justify-content: flex-end; overflow-x: auto; }
}

@media (pointer: coarse) {
  .xlsx-toolbar__icon-button,
  .xlsx-toolbar__zoom-select,
  .xlsx-toolbar__search-panel input,
  .xlsx-toolbar__search-summary button,
  .xlsx-toolbar__search-result {
    min-height: 44px;
  }
  .xlsx-toolbar__icon-button {
    min-width: 44px;
  }
  .xlsx-toolbar__search-result { padding-block: 10px; }
}
</style>
