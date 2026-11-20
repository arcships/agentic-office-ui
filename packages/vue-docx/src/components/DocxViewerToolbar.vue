<template>
  <div ref="toolbarRef" class="docx-viewer-toolbar-host">
  <OfficeToolbar
    class="docx-viewer-toolbar"
    data-testid="docx-viewer-toolbar"
    :data-office-theme="isDark ? 'dark' : 'light'"
    data-office-density="comfortable"
    :data-toolbar-compact="compact ? 'true' : 'false'"
  >
    <template #leading>
      <OfficeToolbarGroup gap="sm" class="docx-viewer-toolbar__leading">
        <OfficeButton
          variant="ghost"
          size="icon-sm"
          class="docx-viewer-toolbar__icon-button"
          data-testid="docx-sidebar-toggle"
          :aria-pressed="sidebarOpen"
          :aria-label="m.docx.toolbar.sidebarToggle"
          :title="m.docx.toolbar.pageThumbnails"
          :disabled="disabled"
          @click="emit('toggleSidebar')"
        >
          <OfficeIcon name="columns-2" />
        </OfficeButton>

        <span v-if="!compact" class="docx-viewer-toolbar__file" :title="fileName">{{ fileName }}</span>

        <OfficePageControl
          class="docx-viewer-toolbar__page"
          :model-value="currentPage"
          :total-pages="totalPages"
          :disabled="disabled"
          :native-input="true"
          input-test-id="docx-page-current"
          @update:model-value="emit('selectPage', $event)"
        />
      </OfficeToolbarGroup>
    </template>

    <template #trailing>
      <OfficeToolbarGroup gap="sm" class="docx-viewer-toolbar__actions">
        <OfficeZoomControl
          :model-value="zoom"
          :options="zoomOptions"
          :show-fit="false"
          :native-select="true"
          select-test-id="docx-zoom-select"
          :disabled="disabled"
          @update:model-value="emit('update:zoom', $event)"
        />

        <OfficeSearchPopover
          v-if="!compact"
          ref="searchPopoverRef"
          class="docx-viewer-toolbar__search"
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
          @close="emit('searchClose')"
        />
        <span v-else class="docx-viewer-toolbar__search">
        <OfficeSearchPopover
          ref="searchPopoverRef"
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
          @close="emit('searchClose')"
        />
        </span>

        <DocxViewerToolbarControls
          v-if="!compact"
          v-bind="controlsProps"
          v-on="controlsListeners"
        />

        <OfficePopover
          v-else
          v-model:open="overflowOpen"
          side="bottom"
          align="end"
          :side-offset="6"
          content-class="docx-viewer-toolbar__overflow"
        >
          <template #trigger>
            <OfficeIconButton
              variant="ghost"
              size="icon-sm"
              class="docx-viewer-toolbar__icon-button"
              data-testid="docx-toolbar-overflow"
              :label="m.docx.toolbar.moreActions"
              :title="m.docx.toolbar.moreActions"
              icon="more-horizontal"
              :disabled="disabled"
            />
          </template>
          <DocxViewerToolbarControls
              v-bind="controlsProps"
            v-on="controlsListeners"
          />
        </OfficePopover>
      </OfficeToolbarGroup>
    </template>
  </OfficeToolbar>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue"
import {
  OfficeButton,
  OfficeIcon,
  OfficeIconButton,
  OfficePageControl,
  OfficePopover,
  OfficeSearchPopover,
  OfficeToolbar,
  OfficeToolbarGroup,
  OfficeZoomControl,
  useElementWidth,
} from "@arcships/vue-ui"
import { useDocxMessages } from "../i18n"
import DocxViewerToolbarControls from "./DocxViewerToolbarControls.vue"

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
  }>(),
  { showUpload: true, showDownload: true, searchCaseSensitive: false, searchWholeWord: false, searchRegex: false, viewMode: "continuous", fitMode: "none" },
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
const m = useDocxMessages()
const toolbarRef = ref<HTMLElement | null>(null)
const toolbarWidth = useElementWidth(toolbarRef)
const compact = computed(() => toolbarWidth.value > 0 && toolbarWidth.value < 1360)
const overflowOpen = ref(false)
const searchPopoverRef = ref<InstanceType<typeof OfficeSearchPopover>>()
const searchStatusText = computed(() =>
  props.searchResultCount
    ? `${props.searchResultIndex + 1}/${props.searchResultCount}`
    : props.searchQuery
      ? "0/0"
      : "",
)

const controlsProps = computed(() => ({
  canDownload: props.canDownload,
  disabled: props.disabled,
  isDark: props.isDark,
  showTrackedChanges: props.showTrackedChanges,
  revisionDisplay: props.revisionDisplay,
  showComments: props.showComments,
  trackedChangeCount: props.trackedChangeCount,
  commentCount: props.commentCount,
  viewMode: props.viewMode,
  fitMode: props.fitMode,
  searchQuery: props.searchQuery,
  searchResultCount: props.searchResultCount,
  searchResultIndex: props.searchResultIndex,
  searchCaseSensitive: props.searchCaseSensitive,
  searchWholeWord: props.searchWholeWord,
  searchRegex: props.searchRegex,
  showUpload: props.showUpload,
  showDownload: props.showDownload,
}))

const controlsListeners = {
  download: () => emit("download"),
  upload: () => emit("upload"),
  toggleTheme: () => emit("toggleTheme"),
  toggleTrackedChanges: () => emit("toggleTrackedChanges"),
  toggleComments: () => emit("toggleComments"),
  "update:revisionDisplay": (display: import("@arcships/docx-core").DocxRevisionDisplay) =>
    emit("update:revisionDisplay", display),
  "update:viewMode": (mode: import("../composables/docx-view-mode").DocxViewMode) =>
    emit("update:viewMode", mode),
  "update:fitMode": (
    mode: import("../composables/docx-view-mode").DocxFitMode | "none",
  ) => emit("update:fitMode", mode),
  "update:searchQuery": (query: string) => emit("update:searchQuery", query),
  "update:searchCaseSensitive": (value: boolean) =>
    emit("update:searchCaseSensitive", value),
  "update:searchWholeWord": (value: boolean) =>
    emit("update:searchWholeWord", value),
  "update:searchRegex": (value: boolean) => emit("update:searchRegex", value),
  searchNext: () => emit("searchNext"),
  searchPrevious: () => emit("searchPrevious"),
  searchClose: () => emit("searchClose"),
}

watch(compact, (value) => {
  if (!value) overflowOpen.value = false
})


async function focusSearch(): Promise<void> {
  if (compact.value) {
    searchPopoverRef.value?.open()
    await nextTick()
  }
  searchPopoverRef.value?.focus()
}

defineExpose({
  focusSearch,
})
</script>

<style scoped>
.docx-viewer-toolbar-host {
  flex: 0 0 auto;
  min-inline-size: 0;
}

:global(.docx-viewer-toolbar__overflow) {
  inline-size: 19rem;
  max-inline-size: calc(100vw - var(--office-ui-space-6));
  padding: var(--office-ui-space-2);
}

:global(.docx-viewer-toolbar__overflow .docx-viewer-toolbar__actions) {
  flex-direction: column;
  align-items: stretch;
  gap: var(--office-ui-space-2);
}

:global(.docx-viewer-toolbar__overflow .docx-viewer-toolbar__actions .docx-viewer-toolbar__search) {
  flex-wrap: wrap;
  inline-size: 100%;
}

.docx-viewer-toolbar {
  --office-ui-toolbar-min-height: 3rem;
  --office-ui-toolbar-padding-block: var(--office-ui-space-2);
  --office-ui-toolbar-padding-inline: var(--office-ui-space-3);
  font-size: var(--office-text-base, 1rem);
}

/* Match Extend's sm control row: 28px controls + 16px vertical padding + 1px
 * border stays inside the 48px min-height. */
.docx-viewer-toolbar :deep(.office-input),
.docx-viewer-toolbar :deep(.office-page-control__input),
.docx-viewer-toolbar :deep(.office-zoom-control__native),
.docx-viewer-toolbar__revision-select {
  height: var(--office-ui-height-xs, 1.75rem);
  min-height: var(--office-ui-height-xs, 1.75rem);
}

/* The verification demo applies a host-wide 36px min-height to page buttons.
 * The toolbar owns its 28/32px control contract, so restore it here and opt
 * back into the 44px coarse-pointer target explicitly. */
.docx-viewer-toolbar :deep(button),
.docx-viewer-toolbar :deep(input),
.docx-viewer-toolbar :deep(select) {
  min-height: 0;
  min-block-size: 0;
}

.docx-viewer-toolbar__leading {
  flex: 1 1 auto;
  min-inline-size: 0;
  gap: var(--office-ui-space-2);
}

.docx-viewer-toolbar :deep(.office-toolbar__leading),
.docx-viewer-toolbar :deep(.office-toolbar__trailing) {
  min-inline-size: 0;
}

.docx-viewer-toolbar :deep(.office-toolbar__trailing) {
  flex: 0 1 auto;
}

.docx-viewer-toolbar__actions {
  flex: 0 1 auto;
  justify-content: flex-end;
  max-inline-size: 100%;
  row-gap: var(--office-ui-space-1);
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

.docx-viewer-toolbar__page {
  flex: 0 0 auto;
}

.docx-viewer-toolbar__segments {
  flex-wrap: nowrap;
  gap: var(--office-ui-space-0-5);
  padding: 0;
  border: 0;
  border-radius: var(--office-radius-md);
}

.docx-viewer-toolbar__segment {
  inline-size: var(--office-ui-height-sm);
  block-size: var(--office-ui-height-sm);
  color: var(--office-muted-foreground);
}

.docx-viewer-toolbar__segment--active,
.docx-viewer-toolbar__segment[aria-pressed="true"] {
  background: var(--office-accent);
  color: var(--office-foreground);
}

.docx-viewer-toolbar__search {
  align-items: center;
  display: flex;
  flex: 0 1 auto;
  gap: var(--office-ui-space-0-5);
  min-inline-size: 0;
}

.docx-viewer-toolbar__search :deep(.office-input-control) {
  flex: 0 1 8.25rem;
  inline-size: 8.25rem;
  min-inline-size: 5rem;
}

.docx-viewer-toolbar__search :deep(.office-search-popover__option) {
  min-inline-size: var(--office-ui-height-xs);
  padding-inline: var(--office-ui-space-1);
  font-size: var(--office-text-xs);
  font-weight: 600;
}

.docx-viewer-toolbar__toggle-button {
  gap: var(--office-ui-space-1);
  white-space: nowrap;
}

.docx-viewer-toolbar__toggle-button[aria-pressed="true"] {
  background: var(--office-accent);
  border-color: var(--office-border);
  color: var(--office-foreground);
}

.docx-viewer-toolbar__count {
  align-items: center;
  background: var(--office-foreground);
  border-radius: var(--office-radius-2xl);
  color: var(--office-background);
  display: inline-flex;
  font-size: calc(var(--office-text-xs) * 0.85);
  block-size: var(--office-ui-space-4);
  justify-content: center;
  min-inline-size: var(--office-ui-space-4);
  padding-inline: var(--office-ui-space-1);
}

.docx-viewer-toolbar__revision {
  display: inline-flex;
  min-inline-size: 0;
}

.docx-viewer-toolbar__revision-select {
  inline-size: auto;
  max-inline-size: 7rem;
  font-size: var(--office-text-xs);
}

.docx-viewer-toolbar__icon-button {
  color: var(--office-muted-foreground);
}

.docx-viewer-toolbar__icon-button:hover:not(:disabled) {
  color: var(--office-foreground);
}

@container docx-shell (max-width: 720px) {
  .docx-viewer-toolbar :deep(.office-toolbar__trailing) {
    flex: 1 1 100%;
  }
  .docx-viewer-toolbar__actions {
    flex: 1 1 100%;
    justify-content: flex-end;
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
    --office-ui-toolbar-padding-inline: var(--office-ui-space-2);
    --office-ui-toolbar-gap: var(--office-ui-space-1);
  }
  .docx-viewer-toolbar__file {
    display: none;
  }
  .docx-viewer-toolbar__leading {
    justify-content: space-between;
  }
  .docx-viewer-toolbar__actions {
    overflow-x: auto;
    flex-wrap: nowrap;
  }
  .docx-viewer-toolbar__search :deep(.office-input-control) {
    inline-size: 6.5rem;
    min-inline-size: 4.5rem;
  }
  .docx-viewer-toolbar__toggle-button span:not(.docx-viewer-toolbar__count) {
    display: none;
  }
}

@media (pointer: coarse) {
  .docx-viewer-toolbar :deep(button),
  .docx-viewer-toolbar :deep(input),
  .docx-viewer-toolbar :deep(select) {
    min-height: 44px;
    min-block-size: 44px;
  }
  .docx-viewer-toolbar__search :deep(.office-search-popover__option) {
    min-inline-size: 44px;
  }
}
</style>
