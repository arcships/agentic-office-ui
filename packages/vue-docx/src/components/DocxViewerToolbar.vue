<template>
  <header class="docx-viewer-toolbar" data-testid="docx-viewer-toolbar">
    <div class="docx-viewer-toolbar__leading">
      <button
        type="button"
        class="docx-viewer-toolbar__icon-button"
        data-testid="docx-sidebar-toggle"
        :aria-pressed="sidebarOpen"
        :aria-label="m.docx.toolbar.sidebarToggle"
        :title="m.docx.toolbar.pageThumbnails"
        :disabled="disabled"
        @click="emit('toggleSidebar')"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <rect x="3" y="4" width="18" height="16" rx="2" />
          <path d="M8 4v16M5.5 8h.01M5.5 12h.01M5.5 16h.01" />
        </svg>
      </button>

      <span class="docx-viewer-toolbar__file" :title="fileName">{{ fileName }}</span>

      <div class="docx-viewer-toolbar__page" :aria-label="m.common.page.navigation">
        <span>{{ m.common.page.label }}</span>
        <input
          ref="pageInputRef"
          v-model="pageDraft"
          data-testid="docx-page-current"
          :aria-label="m.common.page.current"
          inputmode="numeric"
          pattern="[0-9]*"
          :disabled="disabled || totalPages < 1"
          @focus="onPageFocus"
          @blur="onPageBlur"
          @keydown.enter.prevent="commitPage"
          @keydown.esc.prevent="resetPageDraft"
        />
        <span>{{ m.common.page.of }} {{ totalPages || "–" }}</span>
      </div>
    </div>

    <div class="docx-viewer-toolbar__actions">
      <div class="docx-viewer-toolbar__search" role="search">
        <input
          ref="searchInputRef"
          data-testid="docx-search-input"
          type="search"
          :value="searchQuery"
          :placeholder="m.docx.toolbar.searchPlaceholder"
          :aria-label="m.docx.toolbar.searchPlaceholder"
          :disabled="disabled"
          @input="onSearchInput"
          @compositionstart="searchComposing = true"
          @compositionend="onSearchCompositionEnd"
          @keydown.enter.prevent="onSearchEnter"
          @keydown.esc.prevent="emit('searchClose')"
        />
        <button
          type="button"
          class="docx-viewer-toolbar__search-option"
          :class="{ 'docx-viewer-toolbar__search-option--active': searchCaseSensitive }"
          data-testid="docx-search-case-sensitive"
          :aria-label="m.common.search.caseSensitive"
          :title="m.common.search.caseSensitive"
          :aria-pressed="searchCaseSensitive"
          :disabled="disabled"
          @click="emit('update:searchCaseSensitive', !searchCaseSensitive)"
        >&#65;&#97;</button>
        <button
          type="button"
          class="docx-viewer-toolbar__search-option"
          :class="{ 'docx-viewer-toolbar__search-option--active': searchWholeWord }"
          data-testid="docx-search-whole-word"
          :aria-label="m.common.search.wholeWord"
          :title="m.common.search.wholeWord"
          :aria-pressed="searchWholeWord"
          :disabled="disabled"
          @click="emit('update:searchWholeWord', !searchWholeWord)"
        >&#97;&#98;|</button>
        <button
          type="button"
          class="docx-viewer-toolbar__search-option"
          :class="{ 'docx-viewer-toolbar__search-option--active': searchRegex }"
          data-testid="docx-search-regex"
          :aria-label="m.common.search.regex"
          :title="m.common.search.regex"
          :aria-pressed="searchRegex"
          :disabled="disabled"
          @click="emit('update:searchRegex', !searchRegex)"
        >.*</button>
        <span
          data-testid="docx-search-status"
          role="status"
          aria-live="polite"
          aria-atomic="true"
        >{{ searchResultCount ? `${searchResultIndex + 1}/${searchResultCount}` : searchQuery ? '0/0' : '' }}</span>
        <button type="button" :aria-label="m.common.search.previous" data-testid="docx-search-previous" :disabled="disabled || !searchResultCount" @click="emit('searchPrevious')">‹</button>
        <button type="button" :aria-label="m.common.search.next" data-testid="docx-search-next" :disabled="disabled || !searchResultCount" @click="emit('searchNext')">›</button>
      </div>
      <button
        type="button"
        class="docx-viewer-toolbar__icon-button"
        :aria-label="m.common.zoom.out"
        :title="m.common.zoom.out"
        :disabled="disabled || zoom <= zoomOptions[0]"
        @click="emit('update:zoom', nextZoom(-1))"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <circle cx="12" cy="12" r="9" /><path d="M8 12h8" />
        </svg>
      </button>
      <select
        class="docx-viewer-toolbar__zoom"
        data-testid="docx-zoom-select"
        :aria-label="m.common.zoom.level"
        :disabled="disabled"
        :value="zoom"
        @change="onZoomChange"
      >
        <option v-for="option in zoomSelectOptions" :key="option" :value="option">
          {{ option }}%
        </option>
      </select>
      <button
        type="button"
        class="docx-viewer-toolbar__icon-button"
        :aria-label="m.common.zoom.in"
        :title="m.common.zoom.in"
        :disabled="disabled || zoom >= zoomOptions[zoomOptions.length - 1]"
        @click="emit('update:zoom', nextZoom(1))"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <circle cx="12" cy="12" r="9" /><path d="M8 12h8M12 8v8" />
        </svg>
      </button>

      <span class="docx-viewer-toolbar__separator" aria-hidden="true" />

      <button
        type="button"
        class="docx-viewer-toolbar__toggle-button"
        data-testid="docx-show-tracked-changes"
        :aria-pressed="showTrackedChanges"
        :disabled="disabled || trackedChangeCount === 0"
        :title="trackedChangesTitle"
        @click="emit('toggleTrackedChanges')"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path d="M5 3h10l4 4v14H5zM15 3v5h4M8 12h8M8 16h5" />
        </svg>
        <span>{{ m.docx.toolbar.trackedChanges }}</span>
        <span v-if="trackedChangeCount > 0" class="docx-viewer-toolbar__count">{{ trackedChangeCount }}</span>
      </button>
      <button
        type="button"
        class="docx-viewer-toolbar__toggle-button"
        data-testid="docx-show-comments"
        :aria-pressed="showComments"
        :disabled="disabled || commentCount === 0"
        :title="commentsTitle"
        @click="emit('toggleComments')"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 5h16v11H9l-5 4zM8 9h8M8 12h5" />
        </svg>
        <span>{{ m.docx.toolbar.comments }}</span>
        <span v-if="commentCount > 0" class="docx-viewer-toolbar__count">{{ commentCount }}</span>
      </button>

      <span class="docx-viewer-toolbar__separator" aria-hidden="true" />

      <button
        type="button"
        class="docx-viewer-toolbar__icon-button"
        data-testid="docx-theme-toggle"
        :aria-label="isDark ? m.docx.toolbar.lightThemeAction : m.docx.toolbar.darkThemeAction"
        :title="isDark ? m.docx.toolbar.lightDocument : m.docx.toolbar.darkDocument"
        :disabled="disabled"
        @click="emit('toggleTheme')"
      >
        <svg v-if="isDark" viewBox="0 0 24 24" aria-hidden="true">
          <circle cx="12" cy="12" r="4" /><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4" />
        </svg>
        <svg v-else viewBox="0 0 24 24" aria-hidden="true">
          <path d="M20.5 14.2A8.5 8.5 0 0 1 9.8 3.5 8.5 8.5 0 1 0 20.5 14.2Z" />
        </svg>
      </button>
      <button
        v-if="showUpload"
        type="button"
        class="docx-viewer-toolbar__icon-button"
        data-testid="docx-upload"
        :aria-label="m.docx.toolbar.upload"
        :title="m.docx.toolbar.upload"
        @click="emit('upload')"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path d="M12 16V4M7.5 8.5 12 4l4.5 4.5M5 14v5h14v-5" />
        </svg>
      </button>
      <button
        v-if="showDownload"
        type="button"
        class="docx-viewer-toolbar__icon-button"
        data-testid="docx-download"
        :aria-label="m.docx.toolbar.download"
        :title="m.docx.toolbar.download"
        :disabled="disabled || !canDownload"
        @click="emit('download')"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path d="M12 4v12M7.5 11.5 12 16l4.5-4.5M5 19h14" />
        </svg>
      </button>
    </div>
  </header>
</template>

<script setup lang="ts">
import { computed, ref, watch } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { useDocxMessages } from "../i18n"

const props = withDefaults(
  defineProps<{
    canDownload: boolean
    disabled: boolean
    fileName: string
    isDark: boolean
    sidebarOpen: boolean
    showTrackedChanges: boolean
    showComments: boolean
    trackedChangeCount: number
    commentCount: number
    currentPage: number
    totalPages: number
    zoom: number
    searchQuery: string
    searchResultCount: number
    searchResultIndex: number
    searchCaseSensitive?: boolean
    searchWholeWord?: boolean
    searchRegex?: boolean
    showUpload?: boolean
    showDownload?: boolean
  }>(),
  { showUpload: true, showDownload: true, searchCaseSensitive: false, searchWholeWord: false, searchRegex: false },
)

const emit = defineEmits<{
  download: []
  upload: []
  toggleSidebar: []
  toggleTheme: []
  toggleTrackedChanges: []
  toggleComments: []
  selectPage: [page: number]
  "update:zoom": [zoom: number]
  "update:searchQuery": [query: string]
  "update:searchCaseSensitive": [value: boolean]
  "update:searchWholeWord": [value: boolean]
  "update:searchRegex": [value: boolean]
  searchNext: []
  searchPrevious: []
  searchClose: []
}>()

const zoomOptions = [50, 75, 100, 125, 150, 175, 200]
const zoomSelectOptions = computed(() => {
  if (!Number.isFinite(props.zoom) || zoomOptions.includes(props.zoom)) return zoomOptions
  return [...zoomOptions, props.zoom].sort((a, b) => a - b)
})
const m = useDocxMessages()
const trackedChangesTitle = computed(() => {
  if (props.trackedChangeCount <= 0) return m.value.docx.toolbar.trackedChangesNone
  const template = props.showTrackedChanges
    ? m.value.docx.toolbar.trackedChangesHide
    : m.value.docx.toolbar.trackedChangesShow
  return formatOfficeMessage(template, { count: props.trackedChangeCount })
})
const commentsTitle = computed(() => {
  if (props.commentCount <= 0) return m.value.docx.toolbar.commentsNone
  const template = props.showComments
    ? m.value.docx.toolbar.commentsHide
    : m.value.docx.toolbar.commentsShow
  return formatOfficeMessage(template, { count: props.commentCount })
})
const pageDraft = ref(String(props.currentPage || 1))
const pageInputRef = ref<HTMLInputElement>()
const pageInputFocused = ref(false)
const searchInputRef = ref<HTMLInputElement>()
const searchComposing = ref(false)

function pageInputIsActive(): boolean {
  if (pageInputFocused.value) return true
  return typeof document !== "undefined" && pageInputRef.value === document.activeElement
}

watch(
  () => props.currentPage,
  (page) => {
    // Keep a focused draft intact while the user is typing a page number.
    if (pageInputIsActive()) return
    pageDraft.value = String(page || 1)
  },
)

function onPageFocus(): void {
  pageInputFocused.value = true
}

function onPageBlur(): void {
  pageInputFocused.value = false
  commitPage()
}

function resetPageDraft(): void {
  pageDraft.value = String(props.currentPage || 1)
}

function commitPage(): void {
  const parsed = Number(pageDraft.value)
  if (!Number.isInteger(parsed) || props.totalPages < 1) {
    resetPageDraft()
    return
  }
  const page = Math.min(Math.max(parsed, 1), props.totalPages)
  pageDraft.value = String(page)
  emit("selectPage", page)
}

function searchValue(event: Event): string {
  return (event.target as HTMLInputElement).value
}

function onSearchInput(event: Event): void {
  if (!searchComposing.value) emit("update:searchQuery", searchValue(event))
}

function onSearchCompositionEnd(event: CompositionEvent): void {
  searchComposing.value = false
  emit("update:searchQuery", searchValue(event))
}

function nextZoom(direction: -1 | 1): number {
  const currentIndex = zoomOptions.indexOf(props.zoom)
  if (currentIndex >= 0) {
    return zoomOptions[Math.min(Math.max(currentIndex + direction, 0), zoomOptions.length - 1)]
  }
  if (direction < 0) {
    return [...zoomOptions].reverse().find((option) => option < props.zoom) ?? zoomOptions[0]
  }
  return zoomOptions.find((option) => option > props.zoom) ?? zoomOptions[zoomOptions.length - 1]
}

function onZoomChange(event: Event): void {
  const value = Number((event.target as HTMLSelectElement).value)
  if (Number.isFinite(value) && value > 0) emit("update:zoom", value)
}

function onSearchEnter(event: KeyboardEvent): void {
  if (event.shiftKey) emit("searchPrevious")
  else emit("searchNext")
}

defineExpose({
  focusSearch: () => searchInputRef.value?.focus(),
})
</script>

<style scoped>
.docx-viewer-toolbar {
  align-items: center;
  background: var(--docx-toolbar-bg, #fff);
  border-bottom: 1px solid var(--docx-border, #e4e4e7);
  color: var(--docx-foreground, #18181b);
  display: flex;
  flex: 0 0 auto;
  gap: 12px;
  justify-content: space-between;
  min-height: 48px;
  padding: 7px 12px;
}
.docx-viewer-toolbar__leading,
.docx-viewer-toolbar__actions,
.docx-viewer-toolbar__page {
  align-items: center;
  display: flex;
}
.docx-viewer-toolbar__leading { gap: 10px; min-width: 0; }
.docx-viewer-toolbar__actions { gap: 4px; }
.docx-viewer-toolbar__search { align-items: center; display: flex; gap: 2px; margin-right: 4px; }
.docx-viewer-toolbar__search input { background: #fff; border: 1px solid #d4d4d8; border-radius: 7px; color: #18181b; font: inherit; font-size: 12px; height: 32px; padding: 0 8px; width: 132px; }
.docx-viewer-toolbar__search span { color: #71717a; font-size: 11px; min-width: 30px; text-align: center; }
.docx-viewer-toolbar__search .docx-viewer-toolbar__search-option {
  border-radius: 5px;
  font-size: 11px;
  font-weight: 700;
  height: 26px;
  line-height: 1;
  width: 26px;
}
.docx-viewer-toolbar__search .docx-viewer-toolbar__search-option--active {
  background: #e0e7ff;
  border-color: #818cf8;
  color: #312e81;
}
.docx-viewer-toolbar__search button { background: transparent; border: 1px solid transparent; border-radius: 5px; color: inherit; cursor: pointer; font-size: 20px; height: 28px; line-height: 20px; padding: 0; width: 24px; }
.docx-viewer-toolbar__search button:hover:not(:disabled) { background: #f4f4f5; }
.docx-viewer-toolbar__search button:disabled { opacity: .3; }
.docx-viewer-toolbar__file {
  font-size: 13px;
  font-weight: 600;
  max-width: 220px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.docx-viewer-toolbar__page {
  color: #52525b;
  font-size: 13px;
  gap: 5px;
  white-space: nowrap;
}
.docx-viewer-toolbar__page input {
  background: #fff;
  border: 1px solid #d4d4d8;
  border-radius: 6px;
  color: #18181b;
  font: inherit;
  height: 30px;
  padding: 0 5px;
  text-align: center;
  width: 42px;
}
.docx-viewer-toolbar__icon-button {
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
.docx-viewer-toolbar__toggle-button {
  align-items: center;
  background: transparent;
  border: 1px solid transparent;
  border-radius: 7px;
  color: inherit;
  cursor: pointer;
  display: inline-flex;
  font: inherit;
  font-size: 12px;
  gap: 4px;
  height: 32px;
  padding: 0 7px;
}
.docx-viewer-toolbar__toggle-button:hover:not(:disabled),
.docx-viewer-toolbar__toggle-button[aria-pressed="true"] { background: #eff6ff; border-color: #bfdbfe; color: #1d4ed8; }
.docx-viewer-toolbar__toggle-button:disabled { cursor: default; opacity: .35; }
.docx-viewer-toolbar__toggle-button svg { fill: none; height: 15px; stroke: currentColor; stroke-linecap: round; stroke-linejoin: round; stroke-width: 1.7; width: 15px; }
.docx-viewer-toolbar__count { align-items: center; background: currentColor; border-radius: 999px; color: #fff; display: inline-flex; font-size: 10px; height: 17px; justify-content: center; min-width: 17px; padding: 0 4px; }
.docx-viewer-toolbar__icon-button:hover:not(:disabled) { background: #f4f4f5; }
.docx-viewer-toolbar__icon-button:focus-visible,
.docx-viewer-toolbar__page input:focus-visible,
.docx-viewer-toolbar__zoom:focus-visible {
  outline: 2px solid #2563eb;
  outline-offset: 2px;
}
.docx-viewer-toolbar__icon-button:disabled { cursor: default; opacity: .35; }
.docx-viewer-toolbar__icon-button svg {
  fill: none;
  height: 17px;
  stroke: currentColor;
  stroke-linecap: round;
  stroke-linejoin: round;
  stroke-width: 1.7;
  width: 17px;
}
.docx-viewer-toolbar__zoom {
  background: #fff;
  border: 1px solid #d4d4d8;
  border-radius: 7px;
  color: #18181b;
  font: inherit;
  font-size: 12px;
  height: 32px;
  min-width: 76px;
  padding: 0 8px;
}
.docx-viewer-toolbar__separator { background: #e4e4e7; height: 18px; margin: 0 3px; width: 1px; }

@container docx-shell (max-width: 720px) {
  .docx-viewer-toolbar { align-items: stretch; flex-direction: column; gap: 6px; }
  .docx-viewer-toolbar__actions { justify-content: flex-end; }
  .docx-viewer-toolbar__file { flex: 1; max-width: none; }
}

@container docx-shell (max-width: 460px) {
  .docx-viewer-toolbar__file { display: none; }
  .docx-viewer-toolbar__leading { justify-content: space-between; }
  .docx-viewer-toolbar__actions { overflow-x: auto; }
}

@media (pointer: coarse) {
  .docx-viewer-toolbar__icon-button,
  .docx-viewer-toolbar__zoom,
  .docx-viewer-toolbar__page input,
  .docx-viewer-toolbar__search input,
  .docx-viewer-toolbar__search button {
    min-height: 44px;
    min-width: 44px;
  }
}
</style>
