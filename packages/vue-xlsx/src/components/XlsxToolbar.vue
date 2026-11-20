<template>
  <OfficeThemeProvider :theme="isDark ? 'dark' : 'light'">
    <!--
      `xlsx-toolbar--dark` stays as a dark-mode marker: the visual-matrix
      harness and the blackbox suites assert it, while every colour now comes
      from the Office tokens the provider sets on this subtree.
    -->
    <OfficeToolbar
      :class="toolbarClasses"
      data-testid="xlsx-toolbar"
    >
      <template #leading>
        <OfficeToolbarGroup gap="sm" class="xlsx-toolbar__identity">
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
        </OfficeToolbarGroup>
      </template>

      <template #trailing>
        <OfficeToolbarGroup gap="sm" class="xlsx-toolbar__actions">
          <OfficeToolbarGroup
            gap="sm"
            class="xlsx-toolbar__zoom"
            :label="m.xlsx.toolbar.zoomGroup"
          >
            <OfficeIconButton
              class="xlsx-toolbar__icon-button"
              :label="m.common.zoom.out"
              icon="minus"
              size="icon-sm"
              data-testid="xlsx-zoom-out"
              :disabled="!controller.canZoomOut"
              @click="controller.zoomOut()"
            />
            <select
              class="xlsx-toolbar__zoom-select"
              :aria-label="m.common.zoom.level"
              :value="zoomPercent"
              @change="onZoomChange"
            >
              <option v-for="z in zoomPresets" :key="z" :value="z">{{ z }}%</option>
            </select>
            <OfficeIconButton
              class="xlsx-toolbar__icon-button"
              :label="m.common.zoom.in"
              icon="plus"
              size="icon-sm"
              data-testid="xlsx-zoom-in"
              :disabled="!controller.canZoomIn"
              @click="controller.zoomIn()"
            />
          </OfficeToolbarGroup>

          <OfficeToolbarSeparator class="xlsx-toolbar__separator" />

          <div class="xlsx-toolbar__search-wrap">
            <OfficeIconButton
              class="xlsx-toolbar__icon-button"
              :label="m.xlsx.toolbar.search"
              icon="search"
              size="icon-sm"
              data-testid="xlsx-search-toggle"
              :active="searchOpen"
              :aria-expanded="searchOpen"
              @click="toggleSearch"
            />
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

          <OfficeIconButton
            v-if="showUpload"
            class="xlsx-toolbar__icon-button"
            :label="m.xlsx.toolbar.upload"
            icon="upload"
            size="icon-sm"
            data-testid="xlsx-toolbar-upload"
            @click="emit('upload')"
          />
          <OfficeIconButton
            v-if="controller.activeSheet"
            class="xlsx-toolbar__icon-button"
            :label="m.xlsx.toolbar.pageBreaks"
            size="icon-sm"
            data-testid="xlsx-toolbar-page-breaks"
            :active="controller.isPageBreakPreview"
            @click="controller.setPageBreakPreview(!controller.isPageBreakPreview)"
          >
            <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 5h16M4 12h16M4 19h16" stroke-dasharray="2 2" /></svg>
          </OfficeIconButton>
          <OfficeIconButton
            v-if="controller.activeSheet"
            class="xlsx-toolbar__icon-button"
            :label="m.xlsx.toolbar.printPreview"
            icon="printer"
            size="icon-sm"
            data-testid="xlsx-toolbar-print-preview"
            @click="controller.openPrintPreview()"
          />
          <OfficeIconButton
            v-if="controller.canDownload"
            class="xlsx-toolbar__icon-button"
            :label="m.xlsx.toolbar.download"
            icon="download"
            size="icon-sm"
            data-testid="xlsx-toolbar-download"
            @click="controller.download()"
          />
          <OfficeIconButton
            v-if="controller.canExport"
            class="xlsx-toolbar__icon-button"
            :label="m.xlsx.toolbar.export"
            icon="save"
            size="icon-sm"
            data-testid="xlsx-toolbar-export"
            @click="controller.exportXlsx()"
          />
        </OfficeToolbarGroup>
      </template>
    </OfficeToolbar>
  </OfficeThemeProvider>
</template>

<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue";
import { formatOfficeMessage } from "@arcships/office-interaction";
import type { XlsxViewerController } from "@arcships/xlsx-core";
import {
  OfficeFindBar,
  OfficeIconButton,
  OfficeThemeProvider,
  OfficeToolbar,
  OfficeToolbarGroup,
  OfficeToolbarSeparator,
} from "@arcships/vue-ui";
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
const toolbarClasses = computed(() =>
  props.isDark ? "xlsx-toolbar xlsx-toolbar--dark" : "xlsx-toolbar",
);
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
/*
 * The cluster geometry comes from `OfficeToolbar`; the rules below only cover
 * spreadsheet-specific pieces (file identity, zoom select, search popover).
 * Every colour resolves from an Office token so light/dark follow the
 * provider instead of a local `.xlsx-toolbar--dark` override.
 */
.xlsx-toolbar {
  gap: var(--office-ui-space-3, 0.75rem);
}

.xlsx-toolbar__identity {
  min-inline-size: 0;
  font: inherit;
}

.xlsx-toolbar__file-icon {
  flex: 0 0 auto;
  color: var(--office-success-foreground, #15803d);
  fill: none;
  stroke: currentColor;
  stroke-linecap: round;
  stroke-linejoin: round;
  stroke-width: 1.7;
  height: 17px;
  width: 17px;
}

.xlsx-toolbar__filename {
  font-size: var(--office-text-sm, 0.875rem);
  font-weight: 600;
  max-inline-size: 260px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.xlsx-toolbar__export-warning {
  color: var(--office-warning-foreground, #b45309);
  font-size: var(--office-text-xs, 0.75rem);
  max-inline-size: 320px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.xlsx-toolbar__actions {
  margin-inline-start: auto;
}

.xlsx-toolbar__zoom {
  flex-wrap: nowrap;
}

.xlsx-toolbar__zoom-select {
  background: var(--office-background, #fff);
  border: var(--office-border-width, 1px) solid var(--office-input, #d4d4d8);
  border-radius: var(--office-radius-md, 0.5rem);
  color: var(--office-foreground, #18181b);
  font: inherit;
  font-size: var(--office-text-xs, 0.75rem);
  height: 28px;
  min-inline-size: 72px;
  padding-block: 0;
  padding-inline: 6px;
  box-shadow: var(--office-shadow-xs, 0 1px 2px rgb(0 0 0 / 0.05));
  cursor: pointer;
}

.xlsx-toolbar__zoom-select:focus-visible {
  border-color: var(--office-ring, #a1a1aa);
  outline: 2px solid color-mix(in oklab, var(--office-ring, #a1a1aa) 40%, transparent);
  outline-offset: var(--office-ring-offset-width, 1px);
}

.xlsx-toolbar__search-wrap {
  position: relative;
}

.xlsx-toolbar__search-panel {
  background: var(--office-popover, #fff);
  border: var(--office-border-width, 1px) solid var(--office-border, #e4e4e7);
  border-radius: var(--office-radius-lg, 0.625rem);
  box-shadow:
    0 12px 32px color-mix(in oklab, var(--office-foreground, #18181b) 16%, transparent);
  color: var(--office-popover-foreground, #18181b);
  padding: var(--office-ui-space-3, 0.75rem);
  position: absolute;
  inset-block-start: 36px;
  inset-inline-end: 0;
  inline-size: 288px;
  z-index: 40;
}

.xlsx-toolbar__search-panel label {
  display: flex;
  flex-direction: column;
  font-size: var(--office-text-xs, 0.75rem);
  font-weight: 600;
  gap: 6px;
}

.xlsx-toolbar__search-panel input:not([type="checkbox"]) {
  background: var(--office-background, #fff);
  border: var(--office-border-width, 1px) solid var(--office-input, #d4d4d8);
  border-radius: var(--office-radius-md, 0.5rem);
  color: inherit;
  font: inherit;
  height: 34px;
  padding-inline: 10px;
}

.xlsx-toolbar__search-panel input:focus-visible {
  border-color: var(--office-ring, #a1a1aa);
  outline: 2px solid color-mix(in oklab, var(--office-ring, #a1a1aa) 40%, transparent);
  outline-offset: var(--office-ring-offset-width, 1px);
}

.xlsx-toolbar__search-results {
  max-height: 240px;
  overflow: auto;
}

.xlsx-toolbar__search-result {
  align-items: center;
  background: transparent;
  border: 0;
  border-radius: var(--office-radius-sm, 0.375rem);
  color: inherit;
  cursor: pointer;
  display: flex;
  font: inherit;
  gap: var(--office-ui-space-2, 0.5rem);
  padding: 7px 8px;
  text-align: start;
  inline-size: 100%;
}

.xlsx-toolbar__search-result:hover,
.xlsx-toolbar__search-result.active {
  background: var(--office-accent, #f4f4f5);
}

.xlsx-toolbar__search-result strong {
  flex: 0 0 96px;
  font-size: var(--office-text-xs, 0.75rem);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.xlsx-toolbar__search-result span {
  font-size: var(--office-text-xs, 0.75rem);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.xlsx-toolbar__search-result mark {
  background: color-mix(in oklab, var(--office-warning, #fde047) 70%, transparent);
  color: inherit;
  padding: 0;
}

.xlsx-toolbar__replace-options {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  padding-block-start: 8px;
}

.xlsx-toolbar__replace-option {
  align-items: center;
  display: flex;
  flex-direction: row;
  font-size: var(--office-text-xs, 0.75rem);
  gap: 5px;
}

.xlsx-toolbar__search-panel .xlsx-toolbar__replace-option {
  flex-direction: row;
}

.xlsx-toolbar__replace-option input {
  height: 14px;
  width: 14px;
}

/*
 * Narrow hosts: the toolbar wraps into two rows (identity, then actions) and
 * the sheet keeps the remaining height because the toolbar is `flex: 0 0 auto`.
 */
@container xlsx-viewer (max-width: 620px) {
  .xlsx-toolbar {
    align-items: stretch;
    gap: var(--office-ui-space-1, 0.25rem);
  }

  /* Two rows: file identity above, then the scrolling action row. */
  .xlsx-toolbar :deep(.office-toolbar__leading),
  .xlsx-toolbar :deep(.office-toolbar__trailing) {
    flex: 1 1 100%;
  }

  .xlsx-toolbar :deep(.office-toolbar__trailing) {
    justify-content: flex-start;
  }

  .xlsx-toolbar__filename {
    flex: 1;
    max-inline-size: none;
  }

  .xlsx-toolbar__actions {
    flex-wrap: nowrap;
    justify-content: flex-end;
    margin-inline-start: auto;
    overflow-x: auto;
    padding-block-end: 2px;
  }

  .xlsx-toolbar__search-panel {
    inline-size: min(288px, calc(100vw - 24px));
  }
}

@media (pointer: coarse) {
  .xlsx-toolbar__icon-button {
    min-height: 44px;
    min-width: 44px;
  }

  .xlsx-toolbar__zoom-select,
  .xlsx-toolbar__search-panel input:not([type="checkbox"]),
  .xlsx-toolbar__search-result {
    min-height: 44px;
  }

  .xlsx-toolbar__search-result {
    padding-block: 10px;
  }
}
</style>
