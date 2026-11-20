<script setup lang="ts">
/**
 * PowerPoint toolbar (issue #161b).
 *
 * Layout source: `tools/extend-reference/src/vendor/extend/pptx-toolbar.tsx`
 * (Extend UI `PptxToolbar`, pinned 70fe70c6). The default row is the upstream
 * row: thumbnail toggle, previous, `Slide N of M`, next on the leading side;
 * zoom out / zoom select / zoom in and the `...` actions trigger on the trailing
 * side. Every PPTX-only command (browse layout, present, speaker notes,
 * fullscreen, find) lives inside the upstream-style actions menu, so the
 * resting row keeps the upstream controls and order.
 */
import { computed, nextTick, onBeforeUnmount, ref, watch } from "vue"
import {
  CircleMinusIcon,
  CirclePlusIcon,
  MoreHorizontalIcon,
  SidebarLeftIcon,
  ArrowLeft01Icon,
  ArrowRight01Icon,
} from "@hugeicons/core-free-icons"
import { HugeIcon, OfficeIcon, OfficeIconButton } from "@arcships/vue-ui"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { usePptxMessages } from "./i18n"

export interface PptxViewerToolbarProps {
  /** Zero-based index of the active slide. */
  activeSlideIndex: number
  /** Slide count; `0` renders the upstream `of -` empty state. */
  slideCount: number
  /** Disables every navigation/zoom control (loading, empty, error states). */
  disabled?: boolean
  /** Thumbnail rail is open. */
  sidebarOpen?: boolean
  /** Render the upstream thumbnail toggle (host `show-sidebar`). */
  showSidebarToggle?: boolean
  /** Active slide is hidden; renders the PPTX-only hidden marker. */
  hiddenSlide?: boolean
  /** Zoom percentage (100 = actual size). */
  zoom: number
  /** Percent presets offered by the zoom select. */
  zoomOptions?: number[]
  /** Allows the zoom select to change the zoom. */
  canZoom?: boolean
  /** Browse layout toggle is available (`show-view-mode`). */
  showViewMode?: boolean
  /** Current browse layout. */
  viewMode?: "single" | "sorter"
  /** Presentation mode toggle is available. */
  showPresentToggle?: boolean
  /** Speaker-notes toggle is available. */
  showNotesToggle?: boolean
  /** Notes/presenter rail is open. */
  notesOpen?: boolean
  /** Fullscreen command is available. */
  showFullscreenControl?: boolean
  /** The viewer currently owns fullscreen. */
  fullscreen?: boolean
  /** Find-in-slides command is available. */
  showSearch?: boolean
  /** The find bar below the toolbar is open. */
  searchOpen?: boolean
  /** Accessible label of the toolbar. */
  label?: string
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<PptxViewerToolbarProps>(), {
  disabled: false,
  sidebarOpen: false,
  showSidebarToggle: true,
  hiddenSlide: false,
  zoomOptions: () => [50, 75, 100, 125, 150, 200],
  canZoom: true,
  showViewMode: true,
  viewMode: "single",
  showPresentToggle: true,
  showNotesToggle: true,
  notesOpen: false,
  showFullscreenControl: true,
  fullscreen: false,
  showSearch: true,
  searchOpen: false,
  label: undefined,
  class: "",
})

const emit = defineEmits<{
  toggleSidebar: []
  selectSlide: [slideIndex: number]
  "update:zoom": [zoom: number]
  "update:viewMode": [mode: "single" | "sorter"]
  present: []
  "update:notesOpen": [open: boolean]
  toggleFullscreen: []
  toggleSearch: []
}>()

const m = usePptxMessages()

const zoomPercent = computed(() => Math.round(props.zoom))
const zoomPresets = computed(() => {
  const values = new Set<number>(props.zoomOptions)
  values.add(zoomPercent.value)
  return [...values].filter((value) => Number.isFinite(value)).sort((left, right) => left - right)
})
const canZoomOut = computed(
  () => !props.disabled && zoomPercent.value > (zoomPresets.value[0] ?? zoomPercent.value),
)
const canZoomIn = computed(
  () => !props.disabled && zoomPercent.value < (zoomPresets.value[zoomPresets.value.length - 1] ?? zoomPercent.value),
)

const total = computed(() => Math.max(0, Math.trunc(props.slideCount) || 0))
const displaySlide = computed(() => {
  if (total.value < 1) return 1
  return Math.min(Math.max(Math.trunc(props.activeSlideIndex) + 1, 1), total.value)
})
const displayTotal = computed(() => (total.value > 0 ? String(total.value) : "-"))
const canGoPrevious = computed(() => !props.disabled && total.value > 0 && displaySlide.value > 1)
const canGoNext = computed(() => !props.disabled && total.value > 0 && displaySlide.value < total.value)
const slideButtonLabel = computed(() =>
  formatOfficeMessage(m.value.pptx.viewer.editSlideNumber, { number: displaySlide.value }))

// ── Slide number editing (upstream keeps the display button until a click) ──
const editingSlide = ref(false)
const draftSlide = ref(String(displaySlide.value))
const slideInput = ref<HTMLInputElement | null>(null)

watch(displaySlide, (value) => {
  if (!editingSlide.value) draftSlide.value = String(value)
})

function startEditing(): void {
  if (props.disabled || total.value < 1) return
  actionsOpen.value = false
  draftSlide.value = String(displaySlide.value)
  editingSlide.value = true
  void nextTick(() => {
    slideInput.value?.focus()
    slideInput.value?.select()
  })
}

function commitSlide(raw: string = draftSlide.value): void {
  const parsed = Number.parseInt(raw.trim(), 10)
  if (!Number.isInteger(parsed) || total.value < 1) {
    draftSlide.value = String(displaySlide.value)
    return
  }
  const next = Math.min(Math.max(parsed, 1), total.value)
  draftSlide.value = String(next)
  if (next !== displaySlide.value) emit("selectSlide", next - 1)
}

function stopEditing(): void {
  if (!editingSlide.value) return
  editingSlide.value = false
  commitSlide()
}

function onSlideKeydown(event: KeyboardEvent): void {
  if (event.key === "Enter") {
    event.preventDefault()
    stopEditing()
  } else if (event.key === "Escape") {
    event.preventDefault()
    draftSlide.value = String(displaySlide.value)
    editingSlide.value = false
  }
}

function stepSlide(direction: -1 | 1): void {
  if (props.disabled || total.value < 1) return
  const next = Math.min(Math.max(displaySlide.value + direction, 1), total.value)
  if (next !== displaySlide.value) emit("selectSlide", next - 1)
}

// ── Zoom ──────────────────────────────────────────────────────────────────
function stepZoom(direction: -1 | 1): void {
  if (props.disabled) return
  const presets = zoomPresets.value
  const next = direction < 0
    ? [...presets].reverse().find((value) => value < zoomPercent.value)
    : presets.find((value) => value > zoomPercent.value)
  if (typeof next === "number") emit("update:zoom", next)
}

function onZoomSelect(event: Event): void {
  const next = Number((event.target as HTMLSelectElement | null)?.value)
  if (Number.isFinite(next)) emit("update:zoom", next)
}

// ── Actions menu ──────────────────────────────────────────────────────────
const actionsOpen = ref(false)
const actionsRoot = ref<HTMLElement | null>(null)

const showLayoutCommand = computed(() => props.showViewMode)
const hasActions = computed(
  () => props.showViewMode || props.showPresentToggle || props.showNotesToggle
    || props.showFullscreenControl || props.showSearch,
)

function toggleActions(): void {
  if (props.disabled && !props.showSearch) return
  actionsOpen.value = !actionsOpen.value
}

function closeActions(): void {
  actionsOpen.value = false
}

function runLayout(mode: "single" | "sorter"): void {
  closeActions()
  if (props.viewMode !== mode) emit("update:viewMode", mode)
}

function runPresent(): void {
  closeActions()
  emit("present")
}

function runNotes(): void {
  closeActions()
  emit("update:notesOpen", !props.notesOpen)
}

function runFullscreen(): void {
  closeActions()
  emit("toggleFullscreen")
}

function runSearch(): void {
  closeActions()
  emit("toggleSearch")
}

function onDocumentPointerDown(event: PointerEvent): void {
  if (!actionsOpen.value) return
  const target = event.target as Node | null
  if (target && actionsRoot.value && !actionsRoot.value.contains(target)) closeActions()
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

defineExpose({ closeActions })
</script>

<template>
  <div
    class="pptx-toolbar"
    :class="props.class"
    role="toolbar"
    :aria-label="props.label ?? m.pptx.viewer.toolbar"
    data-testid="pptx-viewer-toolbar"
    data-pptx-toolbar
    @keydown.esc.stop.prevent="closeActions"
  >
    <div class="pptx-toolbar__leading">
      <template v-if="props.showSidebarToggle">
        <span class="pptx-toolbar__tooltip-trigger">
          <OfficeIconButton
            variant="ghost"
            size="icon-sm"
            data-testid="pptx-sidebar-toggle"
            :label="m.pptx.viewer.sidebar"
            :disabled="props.disabled"
            @click="emit('toggleSidebar')"
          >
            <HugeIcon :icon="SidebarLeftIcon" class="pptx-toolbar__icon" />
          </OfficeIconButton>
        </span>
        <div class="pptx-toolbar__separator" role="separator" aria-orientation="vertical" />
      </template>

      <span class="pptx-toolbar__tooltip-trigger">
        <OfficeIconButton
          variant="ghost"
          size="icon-sm"
          data-testid="pptx-previous-slide"
          :label="m.pptx.viewer.previousSlide"
          :disabled="!canGoPrevious"
          @click="stepSlide(-1)"
        >
          <HugeIcon :icon="ArrowLeft01Icon" class="pptx-toolbar__icon" />
        </OfficeIconButton>
      </span>

      <div class="pptx-toolbar__slide" data-testid="pptx-page-counter">
        <span>{{ m.pptx.viewer.slideLabel }}</span>
        <span class="pptx-toolbar__slide-control" :data-disabled="props.disabled || total < 1 ? 'true' : undefined">
          <input
            v-if="editingSlide"
            ref="slideInput"
            v-model="draftSlide"
            class="pptx-toolbar__slide-input"
            type="number"
            inputmode="numeric"
            pattern="[0-9]*"
            min="1"
            :max="total || 1"
            :aria-label="m.pptx.viewer.slideNumberInput"
            data-testid="office-page-input"
            @blur="stopEditing"
            @change="commitSlide()"
            @keydown="onSlideKeydown"
          />
          <button
            v-else
            type="button"
            class="pptx-toolbar__slide-display"
            :aria-label="slideButtonLabel"
            :disabled="props.disabled || total < 1"
            data-testid="office-page-input"
            @click="startEditing"
          >{{ displaySlide }}</button>
        </span>
        <span class="pptx-toolbar__slide-total">
          {{ m.common.page.of }}
          <span data-testid="office-page-total">{{ displayTotal }}</span>
        </span>
      </div>

      <span class="pptx-toolbar__tooltip-trigger">
        <OfficeIconButton
          variant="ghost"
          size="icon-sm"
          data-testid="pptx-next-slide"
          :label="m.pptx.viewer.nextSlide"
          :disabled="!canGoNext"
          @click="stepSlide(1)"
        >
          <HugeIcon :icon="ArrowRight01Icon" class="pptx-toolbar__icon" />
        </OfficeIconButton>
      </span>

      <span
        v-if="props.hiddenSlide"
        class="pptx-toolbar__hidden"
        data-testid="pptx-hidden-badge"
      >{{ m.pptx.viewer.hiddenSlide }}</span>
    </div>

    <div class="pptx-toolbar__trailing">
      <div class="pptx-toolbar__group">
        <span class="pptx-toolbar__tooltip-trigger">
          <OfficeIconButton
            variant="ghost"
            size="icon-sm"
            data-testid="pptx-zoom-out"
            :label="m.common.zoom.out"
            :disabled="!canZoomOut"
            @click="stepZoom(-1)"
          >
            <HugeIcon :icon="CircleMinusIcon" class="pptx-toolbar__icon" />
          </OfficeIconButton>
        </span>
        <span
          class="pptx-toolbar__zoom"
          :data-disabled="props.disabled || !props.canZoom ? 'true' : undefined"
        >
          <button
            type="button"
            class="pptx-toolbar__zoom-trigger"
            tabindex="-1"
            aria-hidden="true"
            data-testid="pptx-zoom-reset"
          >
            <span class="pptx-toolbar__zoom-value">{{ zoomPercent }}%</span>
            <OfficeIcon name="chevron-down" class="pptx-toolbar__zoom-icon" />
          </button>
          <select
            class="pptx-toolbar__zoom-native"
            :value="zoomPercent"
            :disabled="props.disabled || !props.canZoom"
            :aria-label="m.common.zoom.level"
            data-testid="pptx-zoom-select"
            @change="onZoomSelect"
          >
            <option v-for="value in zoomPresets" :key="value" :value="value">
              {{ formatOfficeMessage(m.common.zoom.percent, { value }) }}
            </option>
          </select>
        </span>
        <span class="pptx-toolbar__tooltip-trigger">
          <OfficeIconButton
            variant="ghost"
            size="icon-sm"
            data-testid="pptx-zoom-in"
            :label="m.common.zoom.in"
            :disabled="!canZoomIn"
            @click="stepZoom(1)"
          >
            <HugeIcon :icon="CirclePlusIcon" class="pptx-toolbar__icon" />
          </OfficeIconButton>
        </span>
      </div>

      <template v-if="hasActions">
        <div class="pptx-toolbar__separator" role="separator" aria-orientation="vertical" />
        <span ref="actionsRoot" class="pptx-toolbar__actions-wrap">
          <span class="pptx-toolbar__tooltip-trigger">
            <OfficeIconButton
              variant="ghost"
              size="icon-sm"
              data-testid="pptx-actions"
              :label="m.pptx.viewer.moreActions"
              :active="actionsOpen"
              aria-haspopup="menu"
              :aria-expanded="actionsOpen"
              @click="toggleActions"
            >
              <HugeIcon :icon="MoreHorizontalIcon" class="pptx-toolbar__icon" />
            </OfficeIconButton>
          </span>
          <div
            v-if="actionsOpen"
            class="pptx-toolbar__menu"
            role="menu"
            :aria-label="m.pptx.viewer.moreActions"
            data-testid="pptx-actions-menu"
          >
            <template v-if="showLayoutCommand">
              <button
                type="button"
                class="pptx-toolbar__menu-item"
                role="menuitemcheckbox"
                :aria-checked="props.viewMode === 'single'"
                :data-checked="props.viewMode === 'single' ? 'true' : undefined"
                data-testid="pptx-view-mode-single"
                @click="runLayout('single')"
              >
                <OfficeIcon name="file-text" />
                <span>{{ m.pptx.viewer.viewModeSingle }}</span>
                <OfficeIcon v-if="props.viewMode === 'single'" name="check" class="pptx-toolbar__menu-check" />
              </button>
              <button
                type="button"
                class="pptx-toolbar__menu-item"
                role="menuitemcheckbox"
                :aria-checked="props.viewMode === 'sorter'"
                :data-checked="props.viewMode === 'sorter' ? 'true' : undefined"
                data-testid="pptx-view-mode-sorter"
                @click="runLayout('sorter')"
              >
                <OfficeIcon name="layout-grid" />
                <span>{{ m.pptx.viewer.viewModeSorter }}</span>
                <OfficeIcon v-if="props.viewMode === 'sorter'" name="check" class="pptx-toolbar__menu-check" />
              </button>
              <div class="pptx-toolbar__menu-separator" role="separator" />
            </template>

            <button
              v-if="props.showPresentToggle"
              type="button"
              class="pptx-toolbar__menu-item"
              role="menuitem"
              data-testid="pptx-view-mode-present"
              :disabled="props.disabled"
              @click="runPresent"
            >
              <OfficeIcon name="play" />
              <span>{{ m.pptx.viewer.viewModePresent }}</span>
            </button>

            <button
              v-if="props.showNotesToggle"
              type="button"
              class="pptx-toolbar__menu-item"
              role="menuitemcheckbox"
              :aria-checked="props.notesOpen"
              :data-checked="props.notesOpen ? 'true' : undefined"
              data-testid="pptx-notes-toggle"
              :disabled="props.disabled"
              @click="runNotes"
            >
              <OfficeIcon name="sticky-note" />
              <span>{{ m.pptx.viewer.notesToggle }}</span>
              <OfficeIcon v-if="props.notesOpen" name="check" class="pptx-toolbar__menu-check" />
            </button>

            <button
              v-if="props.showFullscreenControl"
              type="button"
              class="pptx-toolbar__menu-item"
              role="menuitem"
              data-testid="pptx-toolbar-fullscreen"
              @click="runFullscreen"
            >
              <OfficeIcon :name="props.fullscreen ? 'minimize' : 'maximize'" />
              <span>{{ m.pptx.viewer.fullscreen }}</span>
            </button>

            <template v-if="props.showSearch">
              <div v-if="props.showPresentToggle || props.showNotesToggle || props.showFullscreenControl" class="pptx-toolbar__menu-separator" role="separator" />
              <button
                type="button"
                class="pptx-toolbar__menu-item"
                role="menuitem"
                data-testid="pptx-search-toggle"
                :aria-expanded="props.searchOpen"
                :disabled="props.disabled"
                @click="runSearch"
              >
                <OfficeIcon name="search" />
                <span>{{ m.pptx.viewer.searchPlaceholder }}</span>
              </button>
            </template>
          </div>
        </span>
      </template>
    </div>
  </div>
</template>

<style scoped>
/*
 * The row is the upstream Extend `min-h-12 px-3 py-2 gap-2` toolbar with
 * `justify-between`. Do not add OfficeToolbar's own padding/gap here: parity
 * measures this component against the pinned reference.
 */
.pptx-toolbar,
.pptx-toolbar * {
  box-sizing: border-box;
}

.pptx-toolbar {
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
}

.pptx-toolbar__leading,
.pptx-toolbar__trailing {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  min-inline-size: 0;
}

.pptx-toolbar__leading {
  gap: 4px;
}

.pptx-toolbar__trailing {
  margin-inline-start: auto;
  justify-content: flex-end;
  gap: 4px;
}

.pptx-toolbar__tooltip-trigger {
  display: inline-flex;
}

.pptx-toolbar__icon {
  display: block;
  inline-size: 16px;
  block-size: 16px;
  flex-shrink: 0;
  opacity: 0.8;
}

.pptx-toolbar__separator {
  flex: none;
  inline-size: 1px;
  block-size: 16px;
  margin: 0 4px;
  background: var(--office-border);
}

.pptx-toolbar__group {
  display: flex;
  flex: none;
  align-items: center;
  gap: 4px;
}

/* Upstream `Slide [N] of M`: `text-sm whitespace-nowrap text-primary`. */
.pptx-toolbar__slide {
  display: flex;
  align-items: center;
  color: var(--office-primary);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  white-space: nowrap;
}

.pptx-toolbar__slide-control {
  position: relative;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-inline-size: 0;
  margin: 0;
  padding: 0;
  border: 0;
  background: transparent;
  color: inherit;
  font: inherit;
}

.pptx-toolbar__slide-display {
  display: flex;
  align-items: center;
  justify-content: center;
  block-size: 32px;
  margin: 0;
  padding: 0 9px;
  border: 1px solid transparent;
  border-radius: var(--office-radius-lg);
  background: transparent;
  color: inherit;
  font: inherit;
  font-size: var(--office-text-base);
  line-height: var(--office-leading-base);
  font-weight: 400;
  cursor: pointer;
}

.pptx-toolbar__slide-display:hover:not(:disabled) {
  background: var(--office-accent);
  color: var(--office-accent-foreground);
}

.pptx-toolbar__slide-display:focus-visible {
  outline: none;
  box-shadow:
    0 0 0 1px var(--office-background),
    0 0 0 3px var(--office-ring);
}

.pptx-toolbar__slide-display:disabled {
  cursor: not-allowed;
  opacity: 0.64;
}

.pptx-toolbar__slide-input {
  block-size: 32px;
  inline-size: 56px;
  min-inline-size: 56px;
  margin: 0 4px;
  padding: 0 8px;
  border: 1px solid var(--office-input);
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  color: var(--office-foreground);
  font: inherit;
  font-size: var(--office-text-base);
  text-align: center;
  appearance: textfield;
}

.pptx-toolbar__slide-input::-webkit-outer-spin-button,
.pptx-toolbar__slide-input::-webkit-inner-spin-button {
  margin: 0;
  appearance: none;
}

.pptx-toolbar__slide-input:focus-visible {
  outline: none;
  box-shadow:
    0 0 0 1px var(--office-background),
    0 0 0 3px var(--office-ring);
}

.pptx-toolbar__slide-total {
  color: inherit;
}

/* Upstream select trigger: `size=sm` = 28px, 9px inline padding, 104px wide. */
.pptx-toolbar__zoom {
  position: relative;
  display: inline-flex;
  flex: none;
  inline-size: 104px;
  min-inline-size: 104px;
  block-size: 32px;
}

.pptx-toolbar__zoom-trigger {
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
  font-size: var(--office-text-base);
  line-height: var(--office-leading-base);
  text-align: start;
  box-shadow:
    0 0 0 0 transparent,
    0 0 0 0 transparent,
    0 0 0 0 transparent,
    0 0 0 0 transparent,
    color-mix(in oklab, black 5%, transparent) 0 1px 2px 0;
  cursor: pointer;
}

.pptx-toolbar__zoom-trigger::before {
  content: "";
  position: absolute;
  inset: 0;
  border-radius: calc(var(--office-radius-lg) - 1px);
  box-shadow: 0 1px 0 0 color-mix(in srgb, black 4%, transparent);
  pointer-events: none;
}

.pptx-toolbar__zoom-value {
  flex: 1 1 0%;
  min-inline-size: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.pptx-toolbar__zoom-icon {
  flex: none;
  inline-size: 18px;
  block-size: 18px;
  margin-inline-end: -4px;
  opacity: 0.8;
}

/* Upstream button/select sizes at the `sm` breakpoint (640px). */
@media (min-width: 640px) {
  .pptx-toolbar__slide-display,
  .pptx-toolbar__slide-input {
    block-size: 28px;
    font-size: var(--office-text-sm);
    line-height: var(--office-leading-sm);
  }

  .pptx-toolbar__zoom {
    block-size: 28px;
  }

  .pptx-toolbar__zoom-trigger {
    font-size: var(--office-text-sm);
    line-height: var(--office-leading-sm);
  }

  .pptx-toolbar__zoom-icon {
    inline-size: 16px;
    block-size: 16px;
  }
}

.pptx-toolbar__zoom-native {
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

.pptx-toolbar__zoom-native:disabled {
  display: none;
}

.pptx-toolbar__zoom[data-disabled="true"] .pptx-toolbar__zoom-trigger {
  opacity: 0.64;
  box-shadow: none;
  cursor: not-allowed;
}

.pptx-toolbar__zoom[data-disabled="true"] .pptx-toolbar__zoom-trigger::before {
  box-shadow: none;
}

[data-office-theme="dark"] .pptx-toolbar__zoom-trigger {
  background-clip: border-box;
}

[data-office-theme="dark"] .pptx-toolbar__zoom-trigger::before {
  box-shadow: 0 -1px 0 0 color-mix(in oklab, white 6%, transparent);
}

.pptx-toolbar__zoom-native option {
  color: var(--office-foreground);
  background: var(--office-popover);
}

/* PPTX-only hidden-slide marker (the pinned reference has no hidden slides). */
.pptx-toolbar__hidden {
  display: inline-flex;
  align-items: center;
  block-size: 20px;
  padding: 0 8px;
  border: 1px solid var(--office-border);
  border-radius: 999px;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  white-space: nowrap;
}

/* ── Actions menu (upstream dropdown geometry, no portal) ─────────────────── */
.pptx-toolbar__actions-wrap {
  position: relative;
  display: inline-flex;
}

.pptx-toolbar__menu {
  position: absolute;
  inset-block-start: calc(100% + 4px);
  inset-inline-end: 0;
  z-index: 40;
  display: flex;
  flex-direction: column;
  inline-size: 11rem;
  padding: 4px;
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-popover);
  background-clip: padding-box;
  color: var(--office-popover-foreground);
  box-shadow: 0 10px 15px -3px color-mix(in srgb, black 10%, transparent),
    0 4px 6px -4px color-mix(in srgb, black 10%, transparent);
}

.pptx-toolbar__menu::before {
  content: "";
  position: absolute;
  inset: 0;
  border-radius: calc(var(--office-radius-lg) - 1px);
  box-shadow: 0 1px 0 0 color-mix(in srgb, black 4%, transparent);
  pointer-events: none;
}

[data-office-theme="dark"] .pptx-toolbar__menu::before {
  box-shadow: 0 -1px 0 0 color-mix(in oklab, white 6%, transparent);
}

.pptx-toolbar__menu-item {
  display: flex;
  align-items: center;
  gap: 8px;
  min-block-size: 28px;
  margin: 0;
  padding: 4px 8px;
  border: 0;
  border-radius: var(--office-radius-sm);
  background: transparent;
  color: inherit;
  font: inherit;
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  text-align: start;
  white-space: nowrap;
  cursor: default;
  user-select: none;
}

.pptx-toolbar__menu-item:hover:not(:disabled),
.pptx-toolbar__menu-item:focus-visible {
  background: var(--office-accent);
  color: var(--office-accent-foreground);
  outline: none;
}

.pptx-toolbar__menu-item:disabled {
  pointer-events: none;
  opacity: 0.64;
}

.pptx-toolbar__menu-item > svg.office-icon {
  inline-size: 16px;
  block-size: 16px;
  flex-shrink: 0;
  opacity: 0.8;
}

.pptx-toolbar__menu-check {
  margin-inline-start: auto;
}

.pptx-toolbar__menu-separator {
  block-size: 1px;
  margin: 4px -4px;
  background: var(--office-border);
}
</style>
