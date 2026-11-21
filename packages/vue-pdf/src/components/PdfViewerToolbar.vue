<script setup lang="ts">
/**
 * PDF toolbar (issue #162).
 *
 * Layout source: `tools/extend-reference/src/vendor/extend/pdf-toolbar.tsx`
 * (Extend UI `PDFViewerToolbar`, pinned 70fe70c6). The default row is the
 * upstream row: sidebar toggle, Page + page button + of N on the left; rotate,
 * zoom, search and the actions trigger on the right. Everything the upstream
 * toolbar does not show (file name, first/last page, layout, download, print)
 * lives in the `...` actions menu; the commands and test ids are unchanged.
 */
import { computed, nextTick, onBeforeUnmount, ref, watch } from "vue"
import {
  ChevronDownIcon,
  CircleMinusIcon,
  CirclePlusIcon,
  MoreHorizontalIcon,
  RotateClockwiseIcon,
  Search01Icon,
  SidebarLeftIcon,
} from "@hugeicons/core-free-icons"
import {
  HugeIcon,
  OfficeButton,
  OfficeIcon,
  OfficeIconButton,
} from "@arcships/vue-ui"
import { usePdfMessages } from "../i18n"
import type { PdfPageLayout } from "../pdf/pdf-page-layout"

export interface PdfViewerToolbarProps {
  activePage: number
  totalPages: number
  disabled?: boolean
  sidebarOpen: boolean
  /** Zoom factor; 1 = 100%. */
  zoom: number
  pageLayout: PdfPageLayout
  fileName: string
  showDownload?: boolean
  showPrint?: boolean
  showRotateControls?: boolean
  canDownload?: boolean
  /** Controlled search draft. */
  searchQuery?: string
  /** Announced search result, for example `Result 1 / 1, page 3`. */
  searchResult?: string
  /** Disable the previous/next hit buttons when there are no hits. */
  searchHasHits?: boolean
}

const props = withDefaults(defineProps<PdfViewerToolbarProps>(), {
  disabled: false,
  showDownload: true,
  showPrint: true,
  showRotateControls: true,
  canDownload: false,
  searchQuery: "",
  searchResult: "",
  searchHasHits: false,
})

const emit = defineEmits<{
  toggleSidebar: []
  selectPage: [page: number]
  "update:zoom": [zoom: number]
  rotate: [direction: 1 | -1]
  "update:pageLayout": [layout: PdfPageLayout]
  "update:searchQuery": [query: string]
  submitSearch: []
  searchPrevious: []
  searchNext: []
  searchClose: []
  download: []
  print: []
}>()

const m = usePdfMessages()

const zoomOptions = [0.5, 0.75, 1, 1.25, 1.5, 2]
const pageLayouts: readonly PdfPageLayout[] = ["continuous", "single", "spread"]
const layoutLabels = computed<Record<PdfPageLayout, string>>(() => ({
  continuous: m.value.pdf.viewer.pageLayoutContinuous,
  single: m.value.pdf.viewer.pageLayoutSingle,
  spread: m.value.pdf.viewer.pageLayoutSpread,
}))
const layoutIcons: Record<PdfPageLayout, "list" | "file-text" | "columns-2"> = {
  continuous: "list",
  single: "file-text",
  spread: "columns-2",
}

const displayPage = computed(() => {
  const total = Math.max(0, Math.trunc(props.totalPages) || 0)
  if (total < 1) return 1
  return Math.min(Math.max(Math.trunc(props.activePage) || 1, 1), total)
})
const displayTotal = computed(() =>
  props.totalPages > 0 ? String(props.totalPages) : "–",
)
/** Upstream's page button width is the text width plus 18px padding and 2px border. */
const zoomPercent = computed(() => Math.round(props.zoom * 100))
const canZoomOut = computed(
  () => !props.disabled && props.zoom > zoomOptions[0],
)
const canZoomIn = computed(
  () => !props.disabled && props.zoom < zoomOptions[zoomOptions.length - 1],
)
const canGoPrevious = computed(() => !props.disabled && displayPage.value > 1)
const canGoNext = computed(
  () => !props.disabled && displayPage.value < (props.totalPages || 0),
)

const pageInput = ref(String(displayPage.value))
const pageFocused = ref(false)
watch(displayPage, (page) => {
  if (!pageFocused.value) pageInput.value = String(page)
})

const searchDraft = ref(props.searchQuery)
watch(
  () => props.searchQuery,
  (query) => {
    if (query !== searchDraft.value) searchDraft.value = query
  },
)
watch(searchDraft, (query) => {
  if (query !== props.searchQuery) emit("update:searchQuery", query)
})

const searchOpen = ref(false)
const actionsOpen = ref(false)
const searchRoot = ref<HTMLElement | null>(null)
const actionsRoot = ref<HTMLElement | null>(null)
const searchInput = ref<HTMLInputElement | null>(null)

function commitPage(raw: string | number = pageInput.value): void {
  const parsed = Number.parseInt(String(raw).trim(), 10)
  if (!Number.isInteger(parsed) || props.totalPages < 1) {
    pageInput.value = String(displayPage.value)
    return
  }
  const next = Math.min(Math.max(parsed, 1), props.totalPages)
  pageInput.value = String(next)
  if (next !== displayPage.value) emit("selectPage", next)
}

function onPageFocus(event: FocusEvent): void {
  pageFocused.value = true
  ;(event.target as HTMLInputElement | null)?.select?.()
}

function onPageBlur(): void {
  pageFocused.value = false
  commitPage()
}

function onPageKeydown(event: KeyboardEvent): void {
  if (props.disabled) return
  switch (event.key) {
    case "Enter":
      event.preventDefault()
      commitPage()
      break
    case "Escape":
      event.preventDefault()
      pageInput.value = String(displayPage.value)
      break
    case "ArrowUp":
      event.preventDefault()
      if (canGoNext.value) emit("selectPage", displayPage.value + 1)
      break
    case "ArrowDown":
      event.preventDefault()
      if (canGoPrevious.value) emit("selectPage", displayPage.value - 1)
      break
    case "Home":
      event.preventDefault()
      if (props.totalPages > 0) emit("selectPage", 1)
      break
    case "End":
      event.preventDefault()
      if (props.totalPages > 0) emit("selectPage", props.totalPages)
      break
    default:
      break
  }
}

function stepZoom(direction: -1 | 1): void {
  if (props.disabled) return
  const next = direction < 0
    ? [...zoomOptions].reverse().find((option) => option < props.zoom)
    : zoomOptions.find((option) => option > props.zoom)
  if (next !== undefined) emit("update:zoom", next)
}

function onZoomSelect(event: Event): void {
  const next = Number((event.target as HTMLSelectElement | null)?.value)
  if (Number.isFinite(next)) emit("update:zoom", next)
}

function selectLayout(layout: PdfPageLayout): void {
  if (props.disabled || layout === props.pageLayout) return
  emit("update:pageLayout", layout)
  closeActions()
}

function openSearch(): void {
  if (props.disabled) return
  actionsOpen.value = false
  searchOpen.value = true
  void nextTick(() => searchInput.value?.focus())
}

function closeSearch(): void {
  if (!searchOpen.value) return
  searchOpen.value = false
  emit("searchClose")
}

function toggleSearch(): void {
  if (searchOpen.value) closeSearch()
  else openSearch()
}

function submitSearch(): void {
  if (props.disabled) return
  emit("submitSearch")
}

function openActions(): void {
  if (props.disabled) return
  searchOpen.value = false
  actionsOpen.value = true
}

function closeActions(): void {
  actionsOpen.value = false
}

function toggleActions(): void {
  if (actionsOpen.value) closeActions()
  else openActions()
}

function runAction(action: "download" | "print" | "first" | "previous" | "next" | "last"): void {
  if (action === "download") emit("download")
  else if (action === "print") emit("print")
  else if (action === "first") emit("selectPage", 1)
  else if (action === "previous") emit("selectPage", displayPage.value - 1)
  else if (action === "next") emit("selectPage", displayPage.value + 1)
  else if (props.totalPages > 0) emit("selectPage", props.totalPages)
  closeActions()
  void nextTick(() => {
    const root = actionsRoot.value
    if (!root || typeof root.querySelector !== "function") return
    root.querySelector<HTMLButtonElement>("button")?.focus()
  })
}

function onDocumentPointerDown(event: PointerEvent): void {
  const target = event.target as Node | null
  if (searchOpen.value && searchRoot.value && target && !searchRoot.value.contains(target)) {
    closeSearch()
  }
  if (actionsOpen.value && actionsRoot.value && target && !actionsRoot.value.contains(target)) {
    closeActions()
  }
}

watch([searchOpen, actionsOpen], ([search, actions]) => {
  if (typeof document === "undefined") return
  const method = search || actions ? "addEventListener" : "removeEventListener"
  if (typeof document[method] !== "function") return
  document[method]("pointerdown", onDocumentPointerDown as EventListener)
})

onBeforeUnmount(() => {
  if (typeof document !== "undefined" && typeof document.removeEventListener === "function") {
    document.removeEventListener("pointerdown", onDocumentPointerDown as EventListener)
  }
})

defineExpose({
  openSearch,
  closeSearch,
  focusSearch: () => searchInput.value?.focus(),
  closeActions,
})
</script>

<template>
  <div
    class="pdf-toolbar"
    role="toolbar"
    :aria-label="m.common.chrome.toolbar"
    data-testid="pdf-toolbar"
    @keydown.esc.stop.prevent="closeActions(); closeSearch()"
  >
    <div class="pdf-toolbar__leading">
      <span class="pdf-toolbar__tooltip-trigger">
        <OfficeIconButton
          variant="ghost"
          size="icon-sm"
          data-testid="pdf-sidebar-toggle"
          :label="m.pdf.viewer.toggleThumbnails"
          :active="sidebarOpen"
          :disabled="disabled"
          @click="emit('toggleSidebar')"
        >
          <HugeIcon :icon="SidebarLeftIcon" class="pdf-toolbar__icon" />
        </OfficeIconButton>
      </span>
      <div class="pdf-toolbar__page">
        <span>{{ m.common.page.label }}</span>
        <span
          class="pdf-toolbar__page-control"
          :data-disabled="disabled ? 'true' : undefined"
        >
          <span class="pdf-toolbar__page-display" aria-hidden="true">{{ displayPage }}</span>
          <input
            v-model="pageInput"
            class="office-ui pdf-toolbar__page-input"
            type="number"
            inputmode="numeric"
            pattern="[0-9]*"
            min="1"
            :max="totalPages || 1"
            :disabled="disabled"
            :aria-label="m.pdf.viewer.pageInput"
            data-testid="pdf-page-input"
            @focus="onPageFocus"
            @blur="onPageBlur"
            @change="commitPage()"
            @keydown="onPageKeydown"
          />
        </span>
        <span class="pdf-toolbar__page-total">
          {{ m.common.page.of }}
          <span data-testid="pdf-page-count">{{ displayTotal }}</span>
        </span>
      </div>
    </div>

    <div class="pdf-toolbar__trailing">
      <template v-if="showRotateControls">
        <div class="pdf-toolbar__group">
          <span class="pdf-toolbar__tooltip-trigger">
            <OfficeIconButton
              variant="ghost"
              size="icon-sm"
              data-testid="pdf-rotate-counterclockwise"
              :label="m.pdf.viewer.rotateCounterclockwise"
              :disabled="disabled"
              @click="emit('rotate', -1)"
            >
              <HugeIcon :icon="RotateClockwiseIcon" class="pdf-toolbar__icon" />
            </OfficeIconButton>
          </span>
          <span class="pdf-toolbar__tooltip-trigger">
            <OfficeIconButton
              variant="ghost"
              size="icon-sm"
              data-testid="pdf-rotate"
              :label="m.pdf.viewer.rotateClockwise"
              :disabled="disabled"
              @click="emit('rotate', 1)"
            >
              <HugeIcon :icon="RotateClockwiseIcon" class="pdf-toolbar__icon pdf-toolbar__icon--mirror" />
            </OfficeIconButton>
          </span>
        </div>
        <div class="pdf-toolbar__separator" role="separator" aria-orientation="vertical" />
      </template>

      <div class="pdf-toolbar__group">
        <span class="pdf-toolbar__tooltip-trigger">
          <OfficeIconButton
            variant="ghost"
            size="icon-sm"
            data-testid="pdf-zoom-out"
            :label="m.common.zoom.out"
            :disabled="!canZoomOut"
            @click="stepZoom(-1)"
          >
            <HugeIcon :icon="CircleMinusIcon" class="pdf-toolbar__icon" />
          </OfficeIconButton>
        </span>
        <span class="pdf-toolbar__zoom" :data-disabled="disabled ? 'true' : undefined">
          <button
            type="button"
            class="office-ui pdf-toolbar__zoom-trigger"
            tabindex="-1"
            aria-hidden="true"
          >
            <span class="pdf-toolbar__zoom-value">{{ zoomPercent }}%</span>
            <HugeIcon :icon="ChevronDownIcon" class="pdf-toolbar__zoom-icon" />
          </button>
          <select
            class="office-ui pdf-toolbar__zoom-native"
            :value="zoom"
            :disabled="disabled"
            :aria-label="m.common.zoom.level"
            data-testid="pdf-zoom-select"
            @change="onZoomSelect"
          >
            <option v-for="value in zoomOptions" :key="value" :value="value">
              {{ Math.round(value * 100) }}%
            </option>
          </select>
        </span>
        <span class="pdf-toolbar__tooltip-trigger">
          <OfficeIconButton
            variant="ghost"
            size="icon-sm"
            data-testid="pdf-zoom-in"
            :label="m.common.zoom.in"
            :disabled="!canZoomIn"
            @click="stepZoom(1)"
          >
            <HugeIcon :icon="CirclePlusIcon" class="pdf-toolbar__icon" />
          </OfficeIconButton>
        </span>
      </div>

      <div class="pdf-toolbar__separator" role="separator" aria-orientation="vertical" />

      <span ref="searchRoot" class="pdf-toolbar__search-wrap">
        <span class="pdf-toolbar__tooltip-trigger">
          <OfficeIconButton
            variant="ghost"
            size="icon-sm"
            data-testid="pdf-search-toggle"
            :label="m.pdf.viewer.searchPlaceholder"
            :active="searchOpen"
            :disabled="disabled"
            @click="toggleSearch"
          >
            <HugeIcon :icon="Search01Icon" class="pdf-toolbar__icon" />
          </OfficeIconButton>
        </span>
        <form
          v-if="searchOpen"
          class="pdf-toolbar__popover pdf-toolbar__search"
          role="search"
          @submit.prevent="submitSearch"
          @keydown.esc.stop.prevent="closeSearch()"
        >
          <div class="pdf-toolbar__search-row">
            <input
              ref="searchInput"
              v-model="searchDraft"
              class="office-ui office-input pdf-toolbar__search-input"
              type="text"
              autocomplete="off"
              spellcheck="false"
              :disabled="disabled"
              :aria-label="m.pdf.viewer.searchLabel"
              :placeholder="m.pdf.viewer.searchPlaceholder"
              data-testid="pdf-search-input"
              @keydown.enter.prevent="submitSearch"
            />
            <OfficeButton
              variant="outline"
              size="sm"
              type="submit"
              data-testid="pdf-search-submit"
              :disabled="disabled || !searchDraft.trim()"
            >{{ m.pdf.viewer.search }}</OfficeButton>
          </div>
          <div class="pdf-toolbar__search-row">
            <span
              class="pdf-toolbar__search-result"
              data-testid="pdf-search-result"
              role="status"
              aria-live="polite"
              aria-atomic="true"
            >{{ searchResult }}</span>
            <OfficeIconButton
              variant="outline"
              size="icon-sm"
              data-testid="pdf-search-prev"
              :label="m.common.search.previous"
              :disabled="!searchHasHits"
              @click="emit('searchPrevious')"
            >
              <OfficeIcon name="chevron-left" />
            </OfficeIconButton>
            <OfficeIconButton
              variant="outline"
              size="icon-sm"
              data-testid="pdf-search-next"
              :label="m.common.search.next"
              :disabled="!searchHasHits"
              @click="emit('searchNext')"
            >
              <OfficeIcon name="chevron-right" />
            </OfficeIconButton>
          </div>
        </form>
      </span>

      <div class="pdf-toolbar__separator" role="separator" aria-orientation="vertical" />

      <span ref="actionsRoot" class="pdf-toolbar__actions-wrap">
        <span class="pdf-toolbar__tooltip-trigger">
          <OfficeIconButton
            variant="ghost"
            size="icon-sm"
            data-testid="pdf-actions"
            :label="m.pdf.viewer.moreActions"
            :active="actionsOpen"
            :disabled="disabled"
            aria-haspopup="menu"
            :aria-expanded="actionsOpen"
            @click="toggleActions"
          >
            <HugeIcon :icon="MoreHorizontalIcon" class="pdf-toolbar__icon" />
          </OfficeIconButton>
        </span>
        <div
          v-if="actionsOpen"
          class="pdf-toolbar__popover pdf-toolbar__menu"
          role="menu"
          :aria-label="m.pdf.viewer.moreActions"
          data-testid="pdf-actions-menu"
          @keydown.esc.stop.prevent="closeActions"
        >
          <div class="pdf-toolbar__menu-label" role="presentation">
            <span>{{ m.common.chrome.fileName }}</span>
            <span class="pdf-toolbar__menu-file">{{ fileName }}</span>
          </div>
          <div class="pdf-toolbar__menu-separator" role="separator" />
          <div
            class="pdf-toolbar__menu-group"
            role="group"
            :aria-label="m.pdf.viewer.pageLayout"
            data-testid="pdf-page-layout"
          >
            <button
              v-for="layout in pageLayouts"
              :key="layout"
              type="button"
              class="pdf-toolbar__menu-item"
              role="menuitemradio"
              :aria-checked="layout === pageLayout"
              :data-layout-active="layout === pageLayout ? 'true' : undefined"
              :data-testid="`pdf-page-layout-${layout}`"
              :disabled="disabled"
              @click="selectLayout(layout)"
            >
              <OfficeIcon
                :name="layoutIcons[layout]"
              />
              <span>{{ layoutLabels[layout] }}</span>
            </button>
          </div>
          <div class="pdf-toolbar__menu-separator" role="separator" />
          <button
            type="button"
            class="pdf-toolbar__menu-item"
            role="menuitem"
            data-testid="pdf-previous-page"
            :disabled="disabled || displayPage <= 1"
            @click="runAction('previous')"
          >{{ m.common.page.previous }}</button>
          <button
            type="button"
            class="pdf-toolbar__menu-item"
            role="menuitem"
            data-testid="pdf-next-page"
            :disabled="disabled || displayPage >= (props.totalPages || 0)"
            @click="runAction('next')"
          >{{ m.common.page.next }}</button>
          <button
            type="button"
            class="pdf-toolbar__menu-item"
            role="menuitem"
            data-testid="pdf-first-page"
            :disabled="disabled || displayPage <= 1"
            @click="runAction('first')"
          >{{ m.common.page.first }}</button>
          <button
            type="button"
            class="pdf-toolbar__menu-item"
            role="menuitem"
            data-testid="pdf-last-page"
            :disabled="disabled || props.totalPages < 1 || displayPage >= props.totalPages"
            @click="runAction('last')"
          >{{ m.common.page.last }}</button>
          <div class="pdf-toolbar__menu-separator" role="separator" />
          <button
            v-if="showPrint"
            type="button"
            class="pdf-toolbar__menu-item"
            role="menuitem"
            data-testid="pdf-print"
            :disabled="disabled"
            @click="runAction('print')"
          >
            <OfficeIcon name="printer" />
            <span>{{ m.pdf.viewer.print }}</span>
          </button>
          <button
            v-if="showDownload"
            type="button"
            class="pdf-toolbar__menu-item"
            role="menuitem"
            data-testid="pdf-download"
            :disabled="disabled || !canDownload"
            @click="runAction('download')"
          >
            <OfficeIcon name="download" />
            <span>{{ m.pdf.viewer.download }}</span>
          </button>
        </div>
      </span>
    </div>
  </div>
</template>

<style scoped>
/*
 * The row is the upstream Extend `min-h-12 px-3 py-2 gap-2` toolbar with
 * `justify-between`. Do not add OfficeToolbar's own padding/gap here: parity
 * measures this component against the pinned reference.
 */
.pdf-toolbar,
.pdf-toolbar * {
  box-sizing: border-box;
}

.pdf-toolbar {
  position: relative;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  min-block-size: 48px;
  padding: 8px 12px;
  border-block-end: 1px solid var(--office-border);
  background: var(--office-background);
  color: var(--office-foreground);
  font-family: inherit;
  font-size: var(--office-text-base);
  line-height: var(--office-leading-base);
}

.pdf-toolbar__leading,
.pdf-toolbar__trailing {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  min-inline-size: 0;
}

.pdf-toolbar__leading {
  gap: 8px;
}

.pdf-toolbar__trailing {
  justify-content: flex-end;
  gap: 4px;
}

.pdf-toolbar__tooltip-trigger {
  display: inline-flex;
}

.pdf-toolbar__icon {
  display: block;
  inline-size: 16px;
  block-size: 16px;
  flex-shrink: 0;
  opacity: 0.8;
}

.pdf-toolbar__icon--mirror {
  transform: scaleX(-1);
}

.pdf-toolbar .office-icon-button[data-active="true"] {
  box-shadow: inset 0 0 0 1px color-mix(in oklab, var(--office-ring) 40%, transparent);
}

.pdf-toolbar__page {
  display: flex;
  align-items: center;
  color: var(--office-primary);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  white-space: nowrap;
}

.pdf-toolbar__page-control {
  position: relative;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  block-size: 28px;
  inline-size: auto;
  flex: none;
  margin: 0;
  padding: 0 9px;
  border: 1px solid transparent;
  border-radius: var(--office-radius-lg);
  background: transparent;
  color: inherit;
  font: inherit;
  cursor: text;
}

.pdf-toolbar__page-display {
  display: block;
  line-height: 20px;
  pointer-events: none;
}

.pdf-toolbar__page-input {
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
  text-align: center;
  appearance: textfield;
  cursor: text;
}

.pdf-toolbar__page-input::-webkit-outer-spin-button,
.pdf-toolbar__page-input::-webkit-inner-spin-button {
  margin: 0;
  appearance: none;
}

.pdf-toolbar__page-control:hover:not([data-disabled="true"]) {
  background: var(--office-accent);
  color: var(--office-accent-foreground);
}

.pdf-toolbar__page-control:focus-within {
  box-shadow:
    0 0 0 1px var(--office-background),
    0 0 0 3px var(--office-ring);
}

.pdf-toolbar__page-control[data-disabled="true"] {
  cursor: not-allowed;
  opacity: 0.64;
}

.pdf-toolbar__page-input:disabled {
  cursor: not-allowed;
}

.pdf-toolbar__page-total {
  color: inherit;
}

.pdf-toolbar__group {
  display: flex;
  flex: none;
  align-items: center;
  gap: 4px;
}

.pdf-toolbar__separator {
  flex: none;
  inline-size: 1px;
  block-size: 16px;
  margin: 0 4px;
  background: var(--office-border);
}

.pdf-toolbar__zoom {
  position: relative;
  display: inline-flex;
  flex: none;
  inline-size: 104px;
  min-inline-size: 104px;
  block-size: 28px;
}

.pdf-toolbar__zoom-trigger {
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
  background: var(--office-background);
  background: light-dark(
    var(--office-background),
    color-mix(in oklab, var(--office-input) 32%, transparent)
  );
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

.pdf-toolbar__zoom-trigger::before {
  content: "";
  position: absolute;
  inset: 0;
  border-radius: calc(var(--office-radius-lg) - 1px);
  box-shadow: 0 1px 0 0 color-mix(in srgb, black 4%, transparent);
  pointer-events: none;
}

.pdf-toolbar__zoom-value {
  flex: 1 1 0%;
  min-inline-size: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.pdf-toolbar__zoom-icon {
  flex: none;
  inline-size: 16px;
  block-size: 16px;
  margin-inline-end: -4px;
  opacity: 0.8;
}

.pdf-toolbar__zoom-native {
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

.pdf-toolbar__zoom-native:disabled {
  display: none;
}

[data-office-theme="dark"] .pdf-toolbar__zoom-trigger {
  background-clip: border-box;
}

[data-office-theme="dark"] .pdf-toolbar__zoom-trigger::before {
  box-shadow: 0 -1px 0 0 color-mix(in oklab, white 6%, transparent);
}

.pdf-toolbar__zoom-native option {
  color: var(--office-foreground);
  background: var(--office-popover);
}

.pdf-toolbar__zoom[data-disabled="true"] .pdf-toolbar__zoom-trigger {
  opacity: 0.64;
  box-shadow: none;
}

.pdf-toolbar__zoom[data-disabled="true"] .pdf-toolbar__zoom-trigger::before {
  box-shadow: none;
}

.pdf-toolbar__zoom:focus-within .pdf-toolbar__zoom-trigger {
  border-color: var(--office-ring);
  box-shadow:
    0 0 0 1px var(--office-background),
    0 0 0 3px var(--office-ring);
}

.pdf-toolbar__search-wrap,
.pdf-toolbar__actions-wrap {
  position: relative;
  display: inline-flex;
}

.pdf-toolbar__popover {
  position: absolute;
  inset-block-start: calc(100% + 6px);
  inset-inline-end: 0;
  z-index: 50;
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-popover);
  color: var(--office-popover-foreground);
  box-shadow: var(--office-shadow-md);
}

.pdf-toolbar__search {
  display: flex;
  flex-direction: column;
  gap: 8px;
  inline-size: 288px;
  max-inline-size: calc(100vw - 24px);
  padding: 8px;
}

.pdf-toolbar__search-row {
  display: flex;
  align-items: center;
  gap: 4px;
  min-inline-size: 0;
}

.pdf-toolbar__search-input {
  flex: 1 1 auto;
  min-inline-size: 0;
  block-size: 28px;
  padding: 0 9px;
  border: 1px solid var(--office-input);
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  color: var(--office-foreground);
  font: inherit;
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
}

.pdf-toolbar__search-input:focus-visible {
  border-color: var(--office-ring);
  outline: 2px solid color-mix(in oklab, var(--office-ring) 40%, transparent);
  outline-offset: 1px;
}

.pdf-toolbar__search-result {
  flex: 1 1 auto;
  min-inline-size: 0;
  overflow: hidden;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  text-overflow: ellipsis;
  white-space: nowrap;
}

.pdf-toolbar__menu {
  display: flex;
  flex-direction: column;
  gap: 1px;
  inline-size: 224px;
  max-inline-size: calc(100vw - 24px);
  padding: 4px;
}

.pdf-toolbar__menu-group {
  display: flex;
  flex-direction: column;
  gap: 1px;
}

.pdf-toolbar__menu-label {
  display: flex;
  align-items: baseline;
  gap: 6px;
  min-inline-size: 0;
  padding: 6px 8px;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
}

.pdf-toolbar__menu-file {
  min-inline-size: 0;
  overflow: hidden;
  color: var(--office-foreground);
  text-overflow: ellipsis;
  white-space: nowrap;
}

.pdf-toolbar__menu-separator {
  block-size: 1px;
  margin: 4px -4px;
  background: var(--office-border);
}

.pdf-toolbar__menu-item {
  display: flex;
  align-items: center;
  gap: 8px;
  inline-size: 100%;
  min-block-size: 28px;
  padding: 4px 8px;
  border: 0;
  border-radius: var(--office-radius-sm);
  background: transparent;
  color: var(--office-foreground);
  font: inherit;
  font-size: var(--office-text-sm);
  text-align: start;
  cursor: pointer;
}

.pdf-toolbar__menu-item:hover:not(:disabled),
.pdf-toolbar__menu-item:focus-visible {
  background: var(--office-accent);
  color: var(--office-accent-foreground);
  outline: none;
}

.pdf-toolbar__menu-item[data-layout-active="true"] {
  background: color-mix(in oklab, var(--office-primary) 10%, transparent);
  color: var(--office-primary);
  font-weight: 600;
}

.pdf-toolbar__menu-item:disabled {
  cursor: not-allowed;
  opacity: 0.64;
}

.pdf-toolbar__menu-item svg.office-icon {
  flex: none;
  inline-size: 16px;
  block-size: 16px;
  opacity: 0.8;
}

/* Tailwind's `sm:` breakpoint: below 640px the upstream button/select sizes
   are h-8 (32px) instead of h-7 (28px). */
@media (max-width: 639px) {
  .pdf-toolbar__page-control,
  .pdf-toolbar__zoom,
  .pdf-toolbar__zoom-trigger {
    block-size: 32px;
  }

  .pdf-toolbar__page-control,
  .pdf-toolbar__zoom-trigger {
    font-size: var(--office-text-base);
    line-height: var(--office-leading-base);
  }

  .pdf-toolbar__page-display {
    line-height: var(--office-leading-base);
  }

  .pdf-toolbar__zoom-icon {
    inline-size: 18px;
    block-size: 18px;
  }
}

@media (pointer: coarse) {
  .pdf-toolbar button {
    min-block-size: 44px;
    min-inline-size: 44px;
  }

  .pdf-toolbar select,
  .pdf-toolbar input {
    min-block-size: 44px;
    min-inline-size: 44px;
  }

  .pdf-toolbar__zoom-native {
    inset-block-start: 50%;
    inset-block-end: auto;
    block-size: 44px;
    translate: 0 -50%;
  }
}

@media print {
  .pdf-toolbar {
    display: none !important;
  }
}
</style>
