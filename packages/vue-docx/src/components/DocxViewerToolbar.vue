<script setup lang="ts">
/**
 * DOCX viewer toolbar (issue #159).
 *
 * Layout source: `tools/extend-reference/src/vendor/extend/docx-toolbar.tsx`
 * (Extend UI `DocxToolbar`, pinned 70fe70c6). The default row is the upstream
 * row: thumbnails toggle, Page + page number + of N on the left; zoom out,
 * zoom level, zoom in, separator and the actions trigger on the right.
 *
 * Every other DOCX command the product ships (find, comments, tracked changes,
 * revision display, view mode, fit, theme, upload, download) lives in the
 * actions menu or behind an opt-in host prop, exactly like upstream's
 * `DocxFileActionsMenu` / `toolbarActions` slots. Commands, events and test ids
 * are unchanged.
 */
import { computed, nextTick, onBeforeUnmount, ref, watch } from "vue"
import {
  ArrowHorizontalIcon,
  ArrowVerticalIcon,
  CircleMinusIcon,
  CirclePlusIcon,
  Comment01Icon,
  FileDiffIcon,
  Moon02Icon,
  MoreHorizontalIcon,
  SidebarLeftIcon,
  Sun03Icon,
} from "@hugeicons/core-free-icons"
import {
  HugeIcon,
  OfficeIcon,
  OfficeIconButton,
  OfficeSearchPopover,
} from "@arcships/vue-ui"
import { useDocxMessages } from "../i18n"

const props = withDefaults(
  defineProps<{
    canDownload: boolean
    disabled: boolean
    fileName: string
    isDark: boolean
    sidebarOpen: boolean
    showTrackedChanges: boolean
    revisionDisplay: import("@arcships/docx-core").DocxRevisionDisplay
    showComments: boolean
    trackedChangeCount: number
    commentCount: number
    currentPage: number
    totalPages: number
    zoom: number
    /** Current page layout mode; the toolbar renders it as a segmented control. */
    viewMode?: import("../composables/docx-view-mode").DocxViewMode
    /** Current fit mode; `none` means the zoom factor is authoritative. */
    fitMode?: import("../composables/docx-view-mode").DocxFitMode | "none"
    searchQuery: string
    searchResultCount: number
    searchResultIndex: number
    searchCaseSensitive?: boolean
    searchWholeWord?: boolean
    searchRegex?: boolean
    showUpload?: boolean
    showDownload?: boolean
    /**
     * Opt-in: show the document file name next to the thumbnails toggle.
     * Off by default so the default row matches the upstream toolbar.
     */
    showFileName?: boolean
    /**
     * Opt-in: keep the find field inline in the row. Off by default; the
     * actions menu and Ctrl/Cmd+F open the same find popover either way.
     */
    showSearch?: boolean
    /** Opt-in: expose the view-mode and fit controls in the actions menu. */
    showViewControls?: boolean
    /** Opt-in: expose the review-display control in the actions menu. */
    showReviewControls?: boolean
  }>(),
  {
    showUpload: true,
    showDownload: true,
    searchCaseSensitive: false,
    searchWholeWord: false,
    searchRegex: false,
    viewMode: "continuous",
    fitMode: "none",
    showFileName: false,
    showSearch: false,
    showViewControls: false,
    showReviewControls: false,
  },
)

const emit = defineEmits<{
  download: []
  upload: []
  toggleSidebar: []
  toggleTheme: []
  toggleTrackedChanges: []
  toggleComments: []
  "update:revisionDisplay": [display: import("@arcships/docx-core").DocxRevisionDisplay]
  selectPage: [page: number]
  "update:zoom": [zoom: number]
  "update:viewMode": [mode: import("../composables/docx-view-mode").DocxViewMode]
  "update:fitMode": [mode: import("../composables/docx-view-mode").DocxFitMode | "none"]
  "update:searchQuery": [query: string]
  "update:searchCaseSensitive": [value: boolean]
  "update:searchWholeWord": [value: boolean]
  "update:searchRegex": [value: boolean]
  searchNext: []
  searchPrevious: []
  searchClose: []
}>()

const zoomOptions = [50, 75, 100, 125, 150, 175, 200]
const viewModes = ["continuous", "single", "spread"] as const
const revisionDisplays = ["inline", "balloons", "none"] as const

const m = useDocxMessages()

const displayPage = computed(() => {
  const total = Math.max(0, Math.trunc(props.totalPages) || 0)
  if (total < 1) return 1
  return Math.min(Math.max(Math.trunc(props.currentPage) || 1, 1), total)
})
const displayTotal = computed(() =>
  props.totalPages > 0 ? String(props.totalPages) : "-",
)
const zoomPercent = computed(() => Math.round(props.zoom))
/** Upstream appends the current gesture zoom when it is not a preset. */
const zoomSelectOptions = computed(() => {
  const current = Math.round(props.zoom * 100) / 100
  if (zoomOptions.includes(current)) return zoomOptions
  return [...zoomOptions, current].sort((left, right) => left - right)
})
const canZoomOut = computed(
  () => !props.disabled && props.zoom > zoomOptions[0]!,
)
const canZoomIn = computed(
  () => !props.disabled && props.zoom < zoomOptions[zoomOptions.length - 1]!,
)
const viewModeLabels = computed<Record<string, string>>(() => ({
  continuous: m.value.docx.toolbar.viewModeContinuous,
  single: m.value.docx.toolbar.viewModeSingle,
  spread: m.value.docx.toolbar.viewModeSpread,
}))
const revisionDisplayLabels = computed<Record<string, string>>(() => ({
  inline: m.value.docx.toolbar.revisionDisplayInline,
  balloons: m.value.docx.toolbar.revisionDisplayBalloons,
  none: m.value.docx.toolbar.revisionDisplayNone,
}))
const trackedChangesLabel = computed(() => {
  if (!props.showTrackedChanges) return m.value.docx.toolbar.trackedChangesShow.replace("{count}", String(props.trackedChangeCount))
  return m.value.docx.toolbar.trackedChangesHide.replace("{count}", String(props.trackedChangeCount))
})
const commentsLabel = computed(() => {
  if (!props.showComments) return m.value.docx.toolbar.commentsShow.replace("{count}", String(props.commentCount))
  return m.value.docx.toolbar.commentsHide.replace("{count}", String(props.commentCount))
})
const themeToggleLabel = computed(() =>
  props.isDark ? m.value.docx.toolbar.lightThemeAction : m.value.docx.toolbar.darkThemeAction,
)

// ── Page number: draft while focused, commit on Enter/blur ─────────
const pageDraft = ref(String(displayPage.value))
const pageFocused = ref(false)
watch(displayPage, (page) => {
  if (!pageFocused.value) pageDraft.value = String(page)
})

function commitPageDraft(): void {
  pageFocused.value = false
  const parsed = Number.parseInt(String(pageDraft.value).trim(), 10)
  if (!Number.isInteger(parsed) || props.totalPages < 1) {
    pageDraft.value = String(displayPage.value)
    return
  }
  const next = Math.min(Math.max(parsed, 1), props.totalPages)
  pageDraft.value = String(next)
  if (next !== displayPage.value) emit("selectPage", next)
}

function onPageFocus(event?: FocusEvent): void {
  pageFocused.value = true
  ;(event?.target as HTMLInputElement | null)?.select?.()
}

function onPageBlur(): void {
  commitPageDraft()
}

function onPageKeydown(event: KeyboardEvent): void {
  if (event.key === "Enter") {
    event.preventDefault()
    commitPageDraft()
  } else if (event.key === "Escape") {
    event.preventDefault()
    pageFocused.value = false
    pageDraft.value = String(displayPage.value)
  }
}

// ── Zoom ───────────────────────────────────────────────────────────
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

// ── Find ───────────────────────────────────────────────────────────
const searchOpen = ref(false)
const searchPopoverRef = ref<InstanceType<typeof OfficeSearchPopover>>()
const searchStatusText = computed(() =>
  props.searchResultCount
    ? `${props.searchResultIndex + 1}/${props.searchResultCount}`
    : props.searchQuery
      ? "0/0"
      : "",
)

function openSearch(): void {
  if (props.disabled) return
  actionsOpen.value = false
  searchOpen.value = true
  void nextTick(() => searchPopoverRef.value?.focus())
}

function closeSearch(): void {
  if (!searchOpen.value) return
  searchOpen.value = false
  emit("searchClose")
}

// ── Actions menu ───────────────────────────────────────────────────
const actionsOpen = ref(false)
const actionsRoot = ref<HTMLElement | null>(null)

function toggleActions(): void {
  if (props.disabled) return
  if (actionsOpen.value) {
    actionsOpen.value = false
    return
  }
  searchOpen.value = false
  actionsOpen.value = true
}

function closeActions(): void {
  actionsOpen.value = false
}

function runAction(action: "download" | "upload" | "theme" | "comments" | "trackedChanges"): void {
  if (action === "download") emit("download")
  else if (action === "upload") emit("upload")
  else if (action === "theme") emit("toggleTheme")
  else if (action === "comments") emit("toggleComments")
  else emit("toggleTrackedChanges")
  closeActions()
}

function selectViewMode(mode: import("../composables/docx-view-mode").DocxViewMode): void {
  if (props.disabled) return
  emit("update:viewMode", mode)
  closeActions()
}

function selectFitMode(
  mode: import("../composables/docx-view-mode").DocxFitMode | "none",
): void {
  if (props.disabled) return
  emit("update:fitMode", mode)
  closeActions()
}

function selectRevisionDisplay(
  display: import("@arcships/docx-core").DocxRevisionDisplay,
): void {
  if (props.disabled) return
  emit("update:revisionDisplay", display)
  closeActions()
}

function findInDocument(): void {
  openSearch()
}

function onDocumentPointerDown(event: PointerEvent): void {
  const target = event.target as Node | null
  if (actionsOpen.value && actionsRoot.value && target && !actionsRoot.value.contains(target)) {
    closeActions()
  }
}

watch(actionsOpen, (open) => {
  if (typeof document === "undefined") return
  const method = open ? "addEventListener" : "removeEventListener"
  if (typeof document[method] !== "function") return
  document[method]("pointerdown", onDocumentPointerDown as EventListener)
})

onBeforeUnmount(() => {
  if (typeof document !== "undefined" && typeof document.removeEventListener === "function") {
    document.removeEventListener("pointerdown", onDocumentPointerDown as EventListener)
  }
})

async function focusSearch(): Promise<void> {
  openSearch()
  await nextTick()
  searchPopoverRef.value?.focus()
}

defineExpose({
  focusSearch,
  closeSearch,
  closeActions,
})
</script>

<template>
  <div
    class="docx-viewer-toolbar"
    role="toolbar"
    :aria-label="m.common.chrome.toolbar"
    data-testid="docx-viewer-toolbar"
    :data-office-theme="isDark ? 'dark' : 'light'"
    data-office-density="comfortable"
    @keydown.esc.stop="closeActions(); closeSearch()"
  >
    <div class="docx-viewer-toolbar__leading">
      <span class="docx-viewer-toolbar__tooltip-trigger">
        <OfficeIconButton
          variant="ghost"
          size="icon-sm"
          class="docx-viewer-toolbar__icon-button"
          data-testid="docx-sidebar-toggle"
          :aria-pressed="sidebarOpen"
          :label="m.docx.toolbar.sidebarToggle"
          :disabled="disabled"
          @click="emit('toggleSidebar')"
        >
          <HugeIcon :icon="SidebarLeftIcon" class="docx-viewer-toolbar__icon" />
        </OfficeIconButton>
      </span>

      <span v-if="showFileName" class="docx-viewer-toolbar__file" :title="fileName">{{ fileName }}</span>

      <div class="docx-viewer-toolbar__page">
        <span>{{ m.common.page.label }}</span>
        <span
          class="docx-viewer-toolbar__page-control"
          :data-disabled="disabled || totalPages < 1 ? 'true' : undefined"
        >
          <span class="docx-viewer-toolbar__page-display" aria-hidden="true">{{ displayPage }}</span>
          <input
            class="docx-viewer-toolbar__page-input"
            type="number"
            inputmode="numeric"
            pattern="[0-9]*"
            min="1"
            :max="totalPages || 1"
            v-model="pageDraft"
            :disabled="disabled || totalPages < 1"
            @input="pageFocused = true"
            :aria-label="m.common.page.label"
            data-testid="docx-page-current"
            @focus="onPageFocus"
            @change="commitPageDraft"
            @blur="onPageBlur"
            @keydown="onPageKeydown"
          />
        </span>
        <span class="docx-viewer-toolbar__page-total">{{ m.common.page.of }} {{ displayTotal }}</span>
      </div>
    </div>

    <div class="docx-viewer-toolbar__trailing">
      <div class="docx-viewer-toolbar__group">
        <span class="docx-viewer-toolbar__tooltip-trigger">
          <OfficeIconButton
            variant="ghost"
            size="icon-sm"
            class="docx-viewer-toolbar__icon-button"
            data-testid="docx-zoom-out"
            :label="m.common.zoom.out"
            :disabled="!canZoomOut"
            @click="stepZoom(-1)"
          >
            <HugeIcon :icon="CircleMinusIcon" class="docx-viewer-toolbar__icon" />
          </OfficeIconButton>
        </span>
        <span class="docx-viewer-toolbar__zoom" :data-disabled="disabled ? 'true' : undefined">
          <button
            type="button"
            class="docx-viewer-toolbar__zoom-trigger"
            tabindex="-1"
            aria-hidden="true"
          >
            <span class="docx-viewer-toolbar__zoom-value">{{ zoomPercent }}%</span>
            <OfficeIcon name="chevron-down" class="docx-viewer-toolbar__zoom-icon" />
          </button>
          <select
            class="docx-viewer-toolbar__zoom-native"
            :value="zoom"
            :disabled="disabled"
            :aria-label="m.common.zoom.level"
            data-testid="docx-zoom-select"
            @change="onZoomSelect"
          >
            <option v-for="value in zoomSelectOptions" :key="value" :value="value">
              {{ value }}%
            </option>
          </select>
        </span>
        <span class="docx-viewer-toolbar__tooltip-trigger">
          <OfficeIconButton
            variant="ghost"
            size="icon-sm"
            class="docx-viewer-toolbar__icon-button"
            data-testid="docx-zoom-in"
            :label="m.common.zoom.in"
            :disabled="!canZoomIn"
            @click="stepZoom(1)"
          >
            <HugeIcon :icon="CirclePlusIcon" class="docx-viewer-toolbar__icon" />
          </OfficeIconButton>
        </span>
      </div>

      <span v-if="showSearch" class="docx-viewer-toolbar__search">
        <OfficeSearchPopover
          ref="searchPopoverRef"
          inline
          :model-value="searchQuery"
          :result-count="searchResultCount"
          :active-index="searchResultIndex"
          :status-text="searchStatusText"
          :case-sensitive="searchCaseSensitive"
          :whole-word="searchWholeWord"
          :regex="searchRegex"
          :show-options="true"
          :disabled="disabled"
          :placeholder="m.docx.toolbar.searchPlaceholder"
          input-type="search"
          test-id-prefix="docx-search"
          @update:model-value="emit('update:searchQuery', $event)"
          @update:case-sensitive="emit('update:searchCaseSensitive', $event)"
          @update:whole-word="emit('update:searchWholeWord', $event)"
          @update:regex="emit('update:searchRegex', $event)"
          @previous="emit('searchPrevious')"
          @next="emit('searchNext')"
          @close="closeSearch"
        />
      </span>

      <div class="docx-viewer-toolbar__separator" role="separator" aria-orientation="vertical" />

      <span ref="actionsRoot" class="docx-viewer-toolbar__actions-wrap">
        <span class="docx-viewer-toolbar__tooltip-trigger">
          <OfficeIconButton
            variant="ghost"
            size="icon-sm"
            class="docx-viewer-toolbar__icon-button"
            data-testid="docx-toolbar-actions"
            :label="m.docx.toolbar.moreActions"
            :active="actionsOpen"
            :disabled="disabled"
            aria-haspopup="menu"
            :aria-expanded="actionsOpen"
            @click="toggleActions"
          >
            <HugeIcon :icon="MoreHorizontalIcon" class="docx-viewer-toolbar__icon" />
          </OfficeIconButton>
        </span>

        <div
          v-if="actionsOpen"
          class="docx-viewer-toolbar__popover docx-viewer-toolbar__menu"
          role="menu"
          :aria-label="m.docx.toolbar.moreActions"
          data-testid="docx-actions-menu"
          @keydown.esc.stop.prevent="closeActions"
        >
          <button
            type="button"
            class="docx-viewer-toolbar__menu-item"
            role="menuitem"
            data-testid="docx-find"
            :disabled="disabled"
            @click="findInDocument"
          >
            <OfficeIcon name="search" />
            <span>{{ m.docx.toolbar.searchPlaceholder }}</span>
          </button>
          <div class="docx-viewer-toolbar__menu-separator" role="separator" />
          <button
            type="button"
            class="docx-viewer-toolbar__menu-item"
            role="menuitemcheckbox"
            data-testid="docx-show-comments"
            :aria-checked="showComments"
            :disabled="disabled"
            @click="runAction('comments')"
          >
            <HugeIcon :icon="Comment01Icon" class="docx-viewer-toolbar__menu-icon" />
            <span>{{ commentsLabel }}</span>
            <span v-if="showComments" class="docx-viewer-toolbar__menu-check" aria-hidden="true">✓</span>
          </button>
          <button
            type="button"
            class="docx-viewer-toolbar__menu-item"
            role="menuitemcheckbox"
            data-testid="docx-show-tracked-changes"
            :aria-checked="showTrackedChanges"
            :disabled="disabled"
            @click="runAction('trackedChanges')"
          >
            <HugeIcon :icon="FileDiffIcon" class="docx-viewer-toolbar__menu-icon" />
            <span>{{ trackedChangesLabel }}</span>
            <span v-if="showTrackedChanges" class="docx-viewer-toolbar__menu-check" aria-hidden="true">✓</span>
          </button>

          <template v-if="showReviewControls">
            <div class="docx-viewer-toolbar__menu-separator" role="separator" />
            <div
              class="docx-viewer-toolbar__menu-group"
              role="group"
              :aria-label="m.docx.toolbar.revisionDisplay"
              data-testid="docx-revision-display"
            >
              <button
                v-for="display in revisionDisplays"
                :key="display"
                type="button"
                class="docx-viewer-toolbar__menu-item"
                role="menuitemradio"
                :aria-checked="display === revisionDisplay"
                :data-display-active="display === revisionDisplay ? 'true' : undefined"
                :data-testid="`docx-revision-display-${display}`"
                :disabled="disabled"
                @click="selectRevisionDisplay(display)"
              >
                <span>{{ revisionDisplayLabels[display] }}</span>
              </button>
            </div>
          </template>

          <template v-if="showViewControls">
            <div class="docx-viewer-toolbar__menu-separator" role="separator" />
            <div
              class="docx-viewer-toolbar__menu-group"
              role="group"
              :aria-label="m.docx.toolbar.viewMode"
              data-testid="docx-view-mode"
            >
              <button
                v-for="mode in viewModes"
                :key="mode"
                type="button"
                class="docx-viewer-toolbar__menu-item"
                role="menuitemradio"
                :aria-checked="mode === viewMode"
                :data-view-mode-active="mode === viewMode ? 'true' : undefined"
                :data-testid="`docx-view-mode-${mode}`"
                :disabled="disabled"
                @click="selectViewMode(mode)"
              >
                <span>{{ viewModeLabels[mode] }}</span>
              </button>
            </div>
            <button
              type="button"
              class="docx-viewer-toolbar__menu-item"
              role="menuitemradio"
              data-testid="docx-fit-width"
              :aria-checked="fitMode === 'width'"
              :disabled="disabled"
              @click="selectFitMode(fitMode === 'width' ? 'none' : 'width')"
            >
              <HugeIcon :icon="ArrowHorizontalIcon" class="docx-viewer-toolbar__menu-icon" />
              <span>{{ m.docx.toolbar.fitWidth }}</span>
            </button>
            <button
              type="button"
              class="docx-viewer-toolbar__menu-item"
              role="menuitemradio"
              data-testid="docx-fit-page"
              :aria-checked="fitMode === 'page'"
              :disabled="disabled"
              @click="selectFitMode(fitMode === 'page' ? 'none' : 'page')"
            >
              <HugeIcon :icon="ArrowVerticalIcon" class="docx-viewer-toolbar__menu-icon" />
              <span>{{ m.docx.toolbar.fitPage }}</span>
            </button>
          </template>

          <div class="docx-viewer-toolbar__menu-separator" role="separator" />
          <button
            type="button"
            class="docx-viewer-toolbar__menu-item"
            role="menuitemcheckbox"
            data-testid="docx-theme-toggle"
            :aria-checked="isDark"
            :disabled="disabled"
            @click="runAction('theme')"
          >
            <HugeIcon
              :icon="isDark ? Sun03Icon : Moon02Icon"
              class="docx-viewer-toolbar__menu-icon"
            />
            <span>{{ themeToggleLabel }}</span>
            <span v-if="isDark" class="docx-viewer-toolbar__menu-check" aria-hidden="true">✓</span>
          </button>
          <button
            v-if="showUpload"
            type="button"
            class="docx-viewer-toolbar__menu-item"
            role="menuitem"
            data-testid="docx-upload"
            :disabled="disabled"
            @click="runAction('upload')"
          >
            <OfficeIcon name="upload" />
            <span>{{ m.docx.toolbar.upload }}</span>
          </button>
          <button
            v-if="showDownload"
            type="button"
            class="docx-viewer-toolbar__menu-item"
            role="menuitem"
            data-testid="docx-download"
            :disabled="disabled || !canDownload"
            @click="runAction('download')"
          >
            <OfficeIcon name="download" />
            <span>{{ m.docx.toolbar.download }}</span>
          </button>
        </div>
      </span>
    </div>
  </div>
</template>

<style scoped>
/*
 * The row is the upstream Extend `min-h-12 px-3 py-2 gap-2` toolbar with
 * `justify-between`. Parity measures this component against the pinned
 * reference, so it owns its padding/gap instead of using OfficeToolbar's.
 */
.docx-viewer-toolbar,
.docx-viewer-toolbar * {
  box-sizing: border-box;
}

.docx-viewer-toolbar {
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

.docx-viewer-toolbar__leading,
.docx-viewer-toolbar__trailing {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  min-inline-size: 0;
}

.docx-viewer-toolbar__leading {
  gap: 8px;
}

.docx-viewer-toolbar__trailing {
  margin-inline-start: auto;
  justify-content: flex-end;
  gap: 4px;
}

.docx-viewer-toolbar__tooltip-trigger {
  display: inline-flex;
}

.docx-viewer-toolbar__icon {
  display: block;
  inline-size: 16px;
  block-size: 16px;
  flex-shrink: 0;
  opacity: 0.8;
}

.docx-viewer-toolbar__file {
  color: var(--office-foreground);
  font-size: var(--office-text-sm);
  font-weight: 600;
  max-inline-size: 13.75rem;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.docx-viewer-toolbar .office-icon-button[data-active="true"] {
  box-shadow: inset 0 0 0 1px color-mix(in oklab, var(--office-ring) 40%, transparent);
}

.docx-viewer-toolbar__page {
  display: flex;
  align-items: center;
  color: var(--office-primary);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  white-space: nowrap;
}

.docx-viewer-toolbar__page-control {
  position: relative;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  block-size: 28px;
  flex: none;
  padding: 0 9px;
  border: 1px solid transparent;
  border-radius: var(--office-radius-lg);
  background: transparent;
  color: inherit;
  font: inherit;
  cursor: text;
}

.docx-viewer-toolbar__page-display {
  display: block;
  line-height: 20px;
  pointer-events: none;
}

.docx-viewer-toolbar__page-input {
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

.docx-viewer-toolbar__page-input::-webkit-outer-spin-button,
.docx-viewer-toolbar__page-input::-webkit-inner-spin-button {
  margin: 0;
  appearance: none;
}

.docx-viewer-toolbar__page-control:hover:not([data-disabled="true"]) {
  background: var(--office-accent);
  color: var(--office-accent-foreground);
}

.docx-viewer-toolbar__page-control:focus-within {
  box-shadow:
    0 0 0 1px var(--office-background),
    0 0 0 3px var(--office-ring);
}

.docx-viewer-toolbar__page-control[data-disabled="true"] {
  cursor: not-allowed;
  opacity: 0.64;
}

.docx-viewer-toolbar__page-input:disabled {
  cursor: not-allowed;
}

.docx-viewer-toolbar__page-total {
  color: inherit;
}

.docx-viewer-toolbar__group {
  display: flex;
  flex: none;
  align-items: center;
  gap: 4px;
}

.docx-viewer-toolbar__separator {
  flex: none;
  inline-size: 1px;
  block-size: 16px;
  margin: 0 4px;
  background: var(--office-border);
}

.docx-viewer-toolbar__zoom {
  position: relative;
  display: inline-flex;
  flex: none;
  inline-size: 104px;
  min-inline-size: 104px;
  block-size: 28px;
}

.docx-viewer-toolbar__zoom-trigger {
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

.docx-viewer-toolbar__zoom-trigger::before {
  content: "";
  position: absolute;
  inset: 0;
  border-radius: calc(var(--office-radius-lg) - 1px);
  box-shadow: 0 1px 0 0 color-mix(in srgb, black 4%, transparent);
  pointer-events: none;
}

.docx-viewer-toolbar__zoom-value {
  flex: 1 1 0%;
  min-inline-size: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.docx-viewer-toolbar__zoom-icon {
  flex: none;
  margin-inline-end: -4px;
  opacity: 0.8;
}

.docx-viewer-toolbar__zoom-native {
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

.docx-viewer-toolbar__zoom-native:disabled {
  display: none;
}

[data-office-theme="dark"] .docx-viewer-toolbar__zoom-trigger {
  background-clip: border-box;
}

[data-office-theme="dark"] .docx-viewer-toolbar__zoom-trigger::before {
  box-shadow: 0 -1px 0 0 color-mix(in oklab, white 6%, transparent);
}

.docx-viewer-toolbar__zoom-native option {
  color: var(--office-foreground);
  background: var(--office-popover);
}

.docx-viewer-toolbar__zoom[data-disabled="true"] .docx-viewer-toolbar__zoom-trigger {
  opacity: 0.64;
  box-shadow: none;
}

.docx-viewer-toolbar__zoom[data-disabled="true"] .docx-viewer-toolbar__zoom-trigger::before {
  box-shadow: none;
}

.docx-viewer-toolbar__zoom:focus-within .docx-viewer-toolbar__zoom-trigger {
  border-color: var(--office-ring);
  box-shadow:
    0 0 0 1px var(--office-background),
    0 0 0 3px var(--office-ring);
}

.docx-viewer-toolbar__search {
  display: flex;
  align-items: center;
  flex: 0 1 auto;
  gap: 4px;
  min-inline-size: 0;
}

.docx-viewer-toolbar__actions-wrap {
  position: relative;
  display: inline-flex;
}

.docx-viewer-toolbar__popover {
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

.docx-viewer-toolbar__menu {
  display: flex;
  flex-direction: column;
  gap: 1px;
  inline-size: 240px;
  max-inline-size: calc(100vw - 24px);
  padding: 4px;
}

.docx-viewer-toolbar__menu-group {
  display: flex;
  flex-direction: column;
  gap: 1px;
}

.docx-viewer-toolbar__menu-separator {
  block-size: 1px;
  margin: 4px -4px;
  background: var(--office-border);
}

.docx-viewer-toolbar__menu-item {
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

.docx-viewer-toolbar__menu-item:hover:not(:disabled),
.docx-viewer-toolbar__menu-item:focus-visible {
  background: var(--office-accent);
  color: var(--office-accent-foreground);
  outline: none;
}

.docx-viewer-toolbar__menu-item[data-view-mode-active="true"],
.docx-viewer-toolbar__menu-item[data-display-active="true"] {
  background: color-mix(in oklab, var(--office-primary) 10%, transparent);
  color: var(--office-primary);
  font-weight: 600;
}

.docx-viewer-toolbar__menu-item:disabled {
  cursor: not-allowed;
  opacity: 0.64;
}

.docx-viewer-toolbar__menu-item svg.office-icon,
.docx-viewer-toolbar__menu-icon {
  flex: none;
  inline-size: 16px;
  block-size: 16px;
  opacity: 0.8;
}

.docx-viewer-toolbar__menu-check {
  margin-inline-start: auto;
  color: var(--office-primary);
}

/* The upstream row wraps below the 640px Tailwind breakpoint; the container
 * queries below keep the DOCX-specific compact variants. */
@media (max-width: 639px) {
  .docx-viewer-toolbar__page-control,
  .docx-viewer-toolbar__zoom,
  .docx-viewer-toolbar__zoom-trigger {
    block-size: 32px;
  }

  .docx-viewer-toolbar__page-control,
  .docx-viewer-toolbar__zoom-trigger {
    font-size: var(--office-text-base);
    line-height: var(--office-leading-base);
  }

  .docx-viewer-toolbar__page-display {
    line-height: var(--office-leading-base);
  }

  .docx-viewer-toolbar__zoom-icon {
    inline-size: 18px;
    block-size: 18px;
  }
}

@container docx-shell (max-width: 720px) {
  .docx-viewer-toolbar__trailing {
    flex: 1 1 100%;
  }

  .docx-viewer-toolbar__leading {
    flex: 1 1 100%;
  }

  .docx-viewer-toolbar__file {
    flex: 1 1 auto;
    max-inline-size: none;
  }
}

@container docx-shell (max-width: 460px) {
  .docx-viewer-toolbar {
    padding-inline: 8px;
    gap: 4px;
  }

  .docx-viewer-toolbar__file {
    display: none;
  }

  .docx-viewer-toolbar__leading {
    justify-content: space-between;
  }
}

@media (pointer: coarse) {
  .docx-viewer-toolbar button {
    min-block-size: 44px;
    min-inline-size: 44px;
  }

  .docx-viewer-toolbar select,
  .docx-viewer-toolbar input {
    min-block-size: 44px;
    min-inline-size: 44px;
  }

  .docx-viewer-toolbar__zoom-native {
    inset-block-start: 50%;
    inset-block-end: auto;
    block-size: 44px;
    translate: 0 -50%;
  }
}
</style>
