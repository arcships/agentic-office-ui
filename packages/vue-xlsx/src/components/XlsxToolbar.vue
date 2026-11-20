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
      <template #trailing>
        <div
          class="xlsx-toolbar__zoom"
          role="group"
          :aria-label="m.xlsx.toolbar.zoomGroup"
        >
          <OfficeIconButton
            class="xlsx-toolbar__icon-button"
            :label="m.common.zoom.out"
            size="icon-sm"
            data-testid="xlsx-zoom-out"
            :disabled="!controller.canZoomOut"
            @click="controller.zoomOut()"
          >
            <HugeIcon :icon="CircleMinusIcon" class="xlsx-toolbar__icon" />
          </OfficeIconButton>
          <!--
            The resting SelectTrigger is painted here and a real <select>
            covers it: the upstream row shows the closed trigger, while the
            keyboard/native listbox behaviour stays with the platform control.
          -->
          <span class="xlsx-toolbar__zoom-select-wrap">
            <button
              type="button"
              class="xlsx-toolbar__zoom-trigger"
              tabindex="-1"
              aria-hidden="true"
            >
              <span class="xlsx-toolbar__zoom-value">{{ zoomPercent }}%</span>
              <OfficeIcon name="chevron-down" class="xlsx-toolbar__zoom-icon" />
            </button>
            <select
              class="office-ui xlsx-toolbar__zoom-native xlsx-toolbar__zoom-select"
              :aria-label="m.common.zoom.level"
              :value="zoomPercent"
              data-testid="xlsx-zoom-select"
              @change="onZoomChange"
            >
              <option v-for="z in zoomPresets" :key="z" :value="z">{{ z }}%</option>
            </select>
          </span>
          <OfficeIconButton
            class="xlsx-toolbar__icon-button"
            :label="m.common.zoom.in"
            size="icon-sm"
            data-testid="xlsx-zoom-in"
            :disabled="!controller.canZoomIn"
            @click="controller.zoomIn()"
          >
            <HugeIcon :icon="CirclePlusIcon" class="xlsx-toolbar__icon" />
          </OfficeIconButton>
        </div>

        <OfficeToolbarSeparator class="xlsx-toolbar__separator" />

        <div class="xlsx-toolbar__search-wrap">
          <OfficeIconButton
            class="xlsx-toolbar__icon-button"
            :label="m.xlsx.toolbar.search"
            size="icon-sm"
            data-testid="xlsx-search-toggle"
            :active="searchOpen"
            :aria-expanded="searchOpen"
            :disabled="controlsDisabled"
            @click="toggleSearch"
          >
            <HugeIcon :icon="Search01Icon" class="xlsx-toolbar__icon" />
          </OfficeIconButton>
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

        <OfficeToolbarSeparator class="xlsx-toolbar__separator" />

        <div ref="actionsRoot" class="xlsx-toolbar__actions-wrap">
          <OfficeIconButton
            class="xlsx-toolbar__icon-button"
            :label="m.xlsx.toolbar.moreActions"
            size="icon-sm"
            data-testid="xlsx-toolbar-actions"
            aria-haspopup="menu"
            :aria-expanded="actionsOpen"
            :active="actionsOpen"
            @click="toggleActions"
          >
            <HugeIcon :icon="MoreHorizontalIcon" class="xlsx-toolbar__icon" />
          </OfficeIconButton>
          <!--
            Upstream keeps everything that is not a zoom/search control in the
            workbook actions menu. `v-show` keeps the commands (and their test
            ids) in the DOM for hosts, component harnesses and assistive tech
            that address them without opening the menu.
          -->
          <div
            v-show="actionsOpen"
            class="xlsx-toolbar__menu"
            role="menu"
            :aria-label="m.xlsx.toolbar.moreActions"
            data-testid="xlsx-toolbar-menu"
            @keydown="onMenuKeydown"
          >
            <div class="xlsx-toolbar__menu-label" role="presentation">
              <span>{{ m.common.chrome.fileName }}</span>
              <span class="xlsx-toolbar__menu-file">{{ displayFileName }}</span>
            </div>
            <div class="xlsx-toolbar__menu-separator" role="separator" />
            <button
              v-for="action in menuActions"
              :key="action.testId"
              type="button"
              class="xlsx-toolbar__menu-item"
              :role="action.checked === undefined ? 'menuitem' : 'menuitemcheckbox'"
              :aria-checked="action.checked === undefined ? undefined : action.checked"
              :data-testid="action.testId"
              :disabled="action.disabled"
              @click="runAction(action)"
            >
              <OfficeIcon v-if="action.icon" :name="action.icon" />
              <span>{{ action.label }}</span>
            </button>
            <p
              v-if="controller.lastExportWarning"
              class="xlsx-toolbar__export-warning"
              data-testid="xlsx-export-warning"
              role="status"
            >{{ controller.lastExportWarning }}</p>
          </div>
        </div>
      </template>
    </OfficeToolbar>
  </OfficeThemeProvider>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, watch } from "vue";
import { formatOfficeMessage } from "@arcships/office-interaction";
import type { XlsxViewerController } from "@arcships/xlsx-core";
import {
  CircleMinusIcon,
  CirclePlusIcon,
  MoreHorizontalIcon,
  Search01Icon,
} from "@hugeicons/core-free-icons";
import {
  HugeIcon,
  OfficeFindBar,
  OfficeIcon,
  OfficeIconButton,
  OfficeThemeProvider,
  OfficeToolbar,
  OfficeToolbarSeparator,
  type OfficeIconName,
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

/** Upstream `ZOOM_OPTIONS` in `xlsx-viewer.tsx`. */
const zoomPresets = [10, 25, 50, 75, 100, 125, 150, 175, 200, 400];
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
/** Upstream disables search while the workbook is loading, failed or empty. */
const controlsDisabled = computed(
  () =>
    Boolean(props.controller.isLoading)
    || Boolean(props.controller.error)
    || (props.controller.sheets?.length ?? 0) === 0,
);
function onZoomChange(event: Event): void {
  const value = Number((event.target as HTMLSelectElement).value);
  if (Number.isFinite(value)) props.controller.setZoomScale(value);
}

function toggleSearch(): void {
  if (searchOpen.value) closeSearch();
  else openSearch();
}

function openSearch(): void {
  closeActions();
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

/**
 * Workbook actions menu.
 *
 * Upstream (`xlsx-viewer.tsx`, `WorkbookFileActionsMenu`) keeps every command
 * that is not zoom or search behind the `...` trigger. The commands and their
 * `data-testid`s are unchanged; only their location moved.
 */
interface ToolbarMenuAction {
  testId: string;
  label: string;
  icon?: OfficeIconName;
  checked?: boolean;
  disabled?: boolean
  run: () => void;
}

const actionsOpen = ref(false);
const actionsRoot = ref<HTMLElement | null>(null);

const menuActions = computed<ToolbarMenuAction[]>(() => {
  const controller = props.controller;
  const actions: ToolbarMenuAction[] = [];
  if (controller.canDownload) {
    actions.push({
      testId: "xlsx-toolbar-download",
      label: m.value.xlsx.toolbar.download,
      icon: "download",
      run: () => controller.download(),
    });
  }
  if (controller.canExport) {
    actions.push({
      testId: "xlsx-toolbar-export",
      label: m.value.xlsx.toolbar.export,
      icon: "save",
      run: () => controller.exportXlsx(),
    });
  }
  if (props.showUpload) {
    actions.push({
      testId: "xlsx-toolbar-upload",
      label: m.value.xlsx.toolbar.upload,
      icon: "upload",
      run: () => emit("upload"),
    });
  }
  if (controller.activeSheet) {
    actions.push({
      testId: "xlsx-toolbar-print-preview",
      label: m.value.xlsx.toolbar.printPreview,
      icon: "printer",
      run: () => controller.openPrintPreview(),
    });
    actions.push({
      testId: "xlsx-toolbar-page-breaks",
      label: m.value.xlsx.toolbar.pageBreaks,
      icon: "layout-grid",
      checked: controller.isPageBreakPreview,
      run: () => controller.setPageBreakPreview(!controller.isPageBreakPreview),
    });
  }
  return actions;
});

function openActions(): void {
  closeSearch();
  actionsOpen.value = true;
}

function closeActions(): void {
  if (!actionsOpen.value) return;
  actionsOpen.value = false;
  void nextTick(() => focusActionsTrigger());
}

function toggleActions(): void {
  if (actionsOpen.value) closeActions();
  else openActions();
}

function focusActionsTrigger(): void {
  const root = actionsRoot.value;
  if (!root || typeof root.querySelector !== "function") return;
  root.querySelector<HTMLButtonElement>("[data-testid='xlsx-toolbar-actions']")?.focus();
}

function runAction(action: ToolbarMenuAction): void {
  if (action.disabled) return;
  actionsOpen.value = false;
  action.run();
  void nextTick(() => focusActionsTrigger());
}

function menuItems(): HTMLButtonElement[] {
  const root = actionsRoot.value;
  if (!root || typeof root.querySelectorAll !== "function") return [];
  return Array.from(
    root.querySelectorAll<HTMLButtonElement>(".xlsx-toolbar__menu-item:not(:disabled)"),
  );
}

function onMenuKeydown(event: KeyboardEvent): void {
  const items = menuItems();
  if (!items.length) return;
  const current = items.indexOf(event.target as HTMLButtonElement);
  switch (event.key) {
    case "Escape":
      event.preventDefault();
      closeActions();
      break;
    case "Tab":
      closeActions();
      break;
    case "ArrowDown":
      event.preventDefault();
      items[(current + 1 + items.length) % items.length]?.focus();
      break;
    case "ArrowUp":
      event.preventDefault();
      items[(current - 1 + items.length) % items.length]?.focus();
      break;
    case "Home":
      event.preventDefault();
      items[0]?.focus();
      break;
    case "End":
      event.preventDefault();
      items[items.length - 1]?.focus();
      break;
    default:
      break;
  }
}

/** Close the menu on a pointer press outside the trigger/panel pair. */
function onDocumentPointerDown(event: PointerEvent): void {
  if (!actionsOpen.value) return;
  const root = actionsRoot.value;
  const target = event.target as Node | null;
  if (!root || !target || root.contains(target)) return;
  actionsOpen.value = false;
}

watch(actionsOpen, (open) => {
  if (typeof document === "undefined") return;
  const method = open ? "addEventListener" : "removeEventListener";
  if (typeof document[method] !== "function") return;
  document[method]("pointerdown", onDocumentPointerDown as EventListener);
});

onBeforeUnmount(() => {
  if (typeof document !== "undefined" && typeof document.removeEventListener === "function") {
    document.removeEventListener("pointerdown", onDocumentPointerDown as EventListener);
  }
});

defineExpose({ openSearch, closeSearch, openActions, closeActions });
</script>

<style scoped>
/*
 * Layout source: `tools/extend-reference/src/vendor/extend/xlsx-toolbar.tsx`
 * (Extend `WorkbookToolbar`, pinned 70fe70c6) — `min-h-12 flex-wrap px-3 py-2
 * gap-2 border-b`, a `gap-1` cluster on the right, then zoom / separator /
 * search / separator / workbook actions. The shared `OfficeToolbar` supplies
 * exactly that shell, so this block only covers the spreadsheet-specific
 * pieces (zoom select, search popover, actions menu).
 */
.xlsx-toolbar {
  /* A chrome row never shrinks: the viewer is a flex column and a squeezed
   * toolbar would clip a wrapped second row instead of growing. */
  flex: 0 0 auto;
}

/*
 * Upstream's row inherits the page type scale (16px / 24px) while the shared
 * `OfficeToolbar` sets `text-sm`; keep the row's own metrics at the reference
 * values so the computed style and the wrapped-line height match.
 */
.xlsx-toolbar.office-toolbar {
  font-size: var(--office-text-base);
  line-height: var(--office-leading-base);
}

/*
 * The toolbar has no leading cluster, and an empty flex item still takes a
 * wrapped line of its own (plus the row gap) on a narrow host. Hide it while it
 * is empty so the row is one line, exactly like the reference's single
 * `ml-auto` cluster.
 */
.xlsx-toolbar :deep(.office-toolbar__leading:empty) {
  display: none;
}

/* Upstream's cluster is `ml-auto`: the row is right-aligned even when the
 * leading cluster is not rendered. */
.xlsx-toolbar :deep(.office-toolbar__trailing) {
  margin-inline-start: auto;
}

.xlsx-toolbar__icon {
  display: block;
  flex-shrink: 0;
  opacity: 0.8;
}

.xlsx-toolbar__zoom {
  display: flex;
  flex: none;
  flex-wrap: nowrap;
  align-items: center;
  gap: var(--office-ui-space-1, 0.25rem);
}

/* Upstream `SelectTrigger size="sm" className="w-[84px] min-w-[84px]"`. */
.xlsx-toolbar__zoom-select-wrap {
  position: relative;
  display: inline-flex;
  flex: none;
  inline-size: 84px;
  min-inline-size: 84px;
  block-size: 28px;
}

.xlsx-toolbar__zoom-trigger {
  position: relative;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 6px;
  inline-size: 100%;
  block-size: 100%;
  margin: 0;
  padding: 0 9px;
  border: 1px solid var(--office-input);
  border-radius: var(--office-radius-lg);
  background: light-dark(
    var(--office-background),
    color-mix(in oklab, var(--office-input) 32%, transparent)
  );
  /* Upstream is `not-dark:bg-clip-padding`: in light the translucent border
   * paints over the row, in dark over the trigger's own fill. */
  background-clip: padding-box;
  color: var(--office-foreground);
  font: inherit;
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  text-align: start;
  box-shadow:
    0 0 0 0 transparent,
    0 0 0 0 transparent,
    0 0 0 0 transparent,
    0 0 0 0 transparent,
    color-mix(in oklab, black 5%, transparent) 0 1px 2px 0;
  cursor: pointer;
}

.xlsx-toolbar__zoom-trigger::before {
  content: "";
  position: absolute;
  inset: 0;
  border-radius: calc(var(--office-radius-lg) - 1px);
  box-shadow: 0 1px 0 0 color-mix(in srgb, black 4%, transparent);
  pointer-events: none;
}

/* Upstream swaps the light 1px bottom inset for a 1px top inset in dark. */
[data-office-theme="dark"] .xlsx-toolbar__zoom-trigger {
  background-clip: border-box;
}

[data-office-theme="dark"] .xlsx-toolbar__zoom-trigger::before {
  /* `oklab` keeps the same rounding as the reference's
   * `--alpha(var(--color-white) / 6%)` inset line; `srgb` lands 1/255 lower. */
  box-shadow: 0 -1px 0 0 color-mix(in oklab, white 6%, transparent);
}

.xlsx-toolbar__zoom-value {
  flex: 1 1 0%;
  min-inline-size: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.xlsx-toolbar__zoom-icon {
  flex: none;
  margin-inline-end: -4px;
  opacity: 0.8;
}

/* Below the `sm` breakpoint the upstream trigger is `min-h-8 text-base` with the
 * `size-4.5` chevron, and only shrinks to `sm:min-h-7 text-sm size-4` at >=640px. */
@media (max-width: 639.98px) {
  .xlsx-toolbar__zoom-select-wrap {
    block-size: 2rem;
  }

  .xlsx-toolbar__zoom-trigger {
    font-size: var(--office-text-base);
    line-height: var(--office-leading-base);
  }

  .xlsx-toolbar__zoom-icon {
    inline-size: 1.125rem;
    block-size: 1.125rem;
  }
}

/* The native listbox covers the painted trigger so the platform control keeps
 * its keyboard behaviour without changing the captured row. */
.xlsx-toolbar__zoom-native {
  position: absolute;
  inset: 0;
  inline-size: 100%;
  block-size: 100%;
  margin: 0;
  padding: 0;
  border: 0;
  background: transparent;
  color: transparent;
  font: inherit;
  appearance: none;
  outline: none;
  cursor: pointer;
}

.xlsx-toolbar__search-wrap {
  position: relative;
  display: inline-flex;
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
  inset-block-start: 32px;
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

/* ── Workbook actions menu (upstream `WorkbookFileActionsMenu`) ──────────── */
.xlsx-toolbar__actions-wrap {
  position: relative;
  display: inline-flex;
}

.xlsx-toolbar__menu {
  position: absolute;
  inset-block-start: 32px;
  inset-inline-end: 0;
  z-index: 40;
  display: flex;
  min-inline-size: 13rem;
  flex-direction: column;
  gap: 2px;
  padding: 4px;
  border: var(--office-border-width, 1px) solid var(--office-border, #e4e4e7);
  border-radius: var(--office-radius-lg, 0.625rem);
  background: var(--office-popover, #fff);
  color: var(--office-popover-foreground, #18181b);
  box-shadow: var(--office-shadow-md, 0 4px 6px -1px rgb(0 0 0 / 0.1));
}

.xlsx-toolbar__menu-label {
  display: flex;
  min-inline-size: 0;
  flex-direction: column;
  gap: 2px;
  padding: 6px 8px;
  color: var(--office-muted-foreground, #52525b);
  font-size: var(--office-text-xs, 0.75rem);
}

.xlsx-toolbar__menu-file {
  overflow: hidden;
  color: var(--office-foreground, #18181b);
  font-size: var(--office-text-sm, 0.875rem);
  text-overflow: ellipsis;
  white-space: nowrap;
}

.xlsx-toolbar__menu-separator {
  block-size: var(--office-border-width, 1px);
  margin: 4px -4px;
  background: var(--office-border, #e4e4e7);
}

.xlsx-toolbar__menu-item {
  display: flex;
  align-items: center;
  gap: var(--office-ui-space-2, 0.5rem);
  inline-size: 100%;
  padding: 6px 8px;
  border: 0;
  border-radius: var(--office-radius-sm, 0.375rem);
  background: transparent;
  color: inherit;
  font: inherit;
  font-size: var(--office-text-sm, 0.875rem);
  text-align: start;
  cursor: pointer;
}

.xlsx-toolbar__menu-item:hover:not(:disabled),
.xlsx-toolbar__menu-item:focus-visible {
  background: var(--office-accent, #f4f4f5);
  outline: none;
}

.xlsx-toolbar__menu-item:disabled {
  cursor: not-allowed;
  opacity: 0.64;
}

.xlsx-toolbar__menu-item .office-icon {
  flex: none;
  opacity: 0.8;
}

.xlsx-toolbar__export-warning {
  margin: 0;
  padding: 6px 8px;
  color: var(--office-warning-foreground, #b45309);
  font-size: var(--office-text-xs, 0.75rem);
}

/*
 * Narrow hosts: the actions cluster scrolls instead of growing the row past
 * the grid, and the popovers stay inside the viewport.
 */
@container xlsx-viewer (max-width: 620px) {
  .xlsx-toolbar__search-wrap,
  .xlsx-toolbar__actions-wrap {
    position: static;
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
