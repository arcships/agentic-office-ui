<template>
  <OfficeToolbar
    class="docx-viewer-toolbar"
    data-testid="docx-viewer-toolbar"
    :data-office-theme="isDark ? 'dark' : 'light'"
    data-office-density="comfortable"
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
          <OfficeIcon name="panel-left" />
        </OfficeButton>

        <span class="docx-viewer-toolbar__file" :title="fileName">{{ fileName }}</span>

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
        <OfficeToolbarGroup
          gap="sm"
          class="docx-viewer-toolbar__segments"
          role="group"
          :aria-label="m.docx.toolbar.viewMode"
        >
          <OfficeButton
            variant="ghost"
            size="icon-sm"
            class="docx-viewer-toolbar__segment"
            :class="viewMode === 'continuous' ? 'docx-viewer-toolbar__segment--active' : undefined"
            data-testid="docx-view-mode-continuous"
            :aria-label="m.docx.toolbar.viewModeContinuous"
            :title="m.docx.toolbar.viewModeContinuous"
            :aria-pressed="viewMode === 'continuous'"
            :disabled="disabled"
            @click="emit('update:viewMode', 'continuous')"
          >
            <OfficeIcon name="rows-3" />
          </OfficeButton>
          <OfficeButton
            variant="ghost"
            size="icon-sm"
            class="docx-viewer-toolbar__segment"
            :class="viewMode === 'single' ? 'docx-viewer-toolbar__segment--active' : undefined"
            data-testid="docx-view-mode-single"
            :aria-label="m.docx.toolbar.viewModeSingle"
            :title="m.docx.toolbar.viewModeSingle"
            :aria-pressed="viewMode === 'single'"
            :disabled="disabled"
            @click="emit('update:viewMode', 'single')"
          >
            <OfficeIcon name="file" />
          </OfficeButton>
          <OfficeButton
            variant="ghost"
            size="icon-sm"
            class="docx-viewer-toolbar__segment"
            :class="viewMode === 'spread' ? 'docx-viewer-toolbar__segment--active' : undefined"
            data-testid="docx-view-mode-spread"
            :aria-label="m.docx.toolbar.viewModeSpread"
            :title="m.docx.toolbar.viewModeSpread"
            :aria-pressed="viewMode === 'spread'"
            :disabled="disabled"
            @click="emit('update:viewMode', 'spread')"
          >
            <OfficeIcon name="columns-2" />
          </OfficeButton>
        </OfficeToolbarGroup>

        <OfficeToolbarGroup
          gap="sm"
          class="docx-viewer-toolbar__segments"
          role="group"
          :aria-label="m.common.zoom.level"
        >
          <OfficeButton
            variant="ghost"
            size="icon-sm"
            class="docx-viewer-toolbar__segment"
            :class="fitMode === 'width' ? 'docx-viewer-toolbar__segment--active' : undefined"
            data-testid="docx-fit-width"
            :aria-label="m.docx.toolbar.fitWidth"
            :title="m.docx.toolbar.fitWidth"
            :aria-pressed="fitMode === 'width'"
            :disabled="disabled"
            @click="emit('update:fitMode', fitMode === 'width' ? 'none' : 'width')"
          >
            <OfficeIcon name="move-horizontal" />
          </OfficeButton>
          <OfficeButton
            variant="ghost"
            size="icon-sm"
            class="docx-viewer-toolbar__segment"
            :class="fitMode === 'page' ? 'docx-viewer-toolbar__segment--active' : undefined"
            data-testid="docx-fit-page"
            :aria-label="m.docx.toolbar.fitPage"
            :title="m.docx.toolbar.fitPage"
            :aria-pressed="fitMode === 'page'"
            :disabled="disabled"
            @click="emit('update:fitMode', fitMode === 'page' ? 'none' : 'page')"
          >
            <OfficeIcon name="scan" />
          </OfficeButton>
        </OfficeToolbarGroup>

        <div class="docx-viewer-toolbar__search" role="search">
          <input
            ref="searchInputRef"
            class="office-ui office-input office-input--sm docx-viewer-toolbar__search-input"
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
          <OfficeButton
            variant="outline"
            size="xs"
            class="docx-viewer-toolbar__search-option"
            data-testid="docx-search-case-sensitive"
            :aria-label="m.common.search.caseSensitive"
            :title="m.common.search.caseSensitive"
            :aria-pressed="searchCaseSensitive"
            :disabled="disabled"
            @click="emit('update:searchCaseSensitive', !searchCaseSensitive)"
          >Aa</OfficeButton>
          <OfficeButton
            variant="outline"
            size="xs"
            class="docx-viewer-toolbar__search-option"
            data-testid="docx-search-whole-word"
            :aria-label="m.common.search.wholeWord"
            :title="m.common.search.wholeWord"
            :aria-pressed="searchWholeWord"
            :disabled="disabled"
            @click="emit('update:searchWholeWord', !searchWholeWord)"
          >ab|</OfficeButton>
          <OfficeButton
            variant="outline"
            size="xs"
            class="docx-viewer-toolbar__search-option"
            data-testid="docx-search-regex"
            :aria-label="m.common.search.regex"
            :title="m.common.search.regex"
            :aria-pressed="searchRegex"
            :disabled="disabled"
            @click="emit('update:searchRegex', !searchRegex)"
          >.*</OfficeButton>
          <span
            class="docx-viewer-toolbar__search-status"
            data-testid="docx-search-status"
            role="status"
            aria-live="polite"
            aria-atomic="true"
          >{{ searchResultCount ? `${searchResultIndex + 1}/${searchResultCount}` : searchQuery ? '0/0' : '' }}</span>
          <OfficeIconButton
            variant="outline"
            size="icon-sm"
            class="docx-viewer-toolbar__icon-button"
            :label="m.common.search.previous"
            icon="chevron-left"
            data-testid="docx-search-previous"
            :disabled="disabled || !searchResultCount"
            @click="emit('searchPrevious')"
          />
          <OfficeIconButton
            variant="outline"
            size="icon-sm"
            class="docx-viewer-toolbar__icon-button"
            :label="m.common.search.next"
            icon="chevron-right"
            data-testid="docx-search-next"
            :disabled="disabled || !searchResultCount"
            @click="emit('searchNext')"
          />
        </div>

        <OfficeZoomControl
          :model-value="zoom"
          :options="zoomOptions"
          :show-fit="false"
          :native-select="true"
          select-test-id="docx-zoom-select"
          :disabled="disabled"
          @update:model-value="emit('update:zoom', $event)"
        />

        <OfficeToolbarSeparator />

        <OfficeButton
          variant="outline"
          size="sm"
          class="docx-viewer-toolbar__toggle-button"
          data-testid="docx-show-tracked-changes"
          :aria-pressed="showTrackedChanges"
          :disabled="disabled || trackedChangeCount === 0"
          :title="trackedChangesTitle"
          @click="emit('toggleTrackedChanges')"
        >
          <OfficeIcon name="history" />
          <span v-if="trackedChangeCount > 0" class="docx-viewer-toolbar__count">{{ trackedChangeCount }}</span>
        </OfficeButton>
        <label class="docx-viewer-toolbar__revision" :title="m.docx.toolbar.revisionDisplay">
          <select
            data-testid="docx-revision-display"
            class="office-ui office-zoom-control__native docx-viewer-toolbar__revision-select"
            :value="revisionDisplay"
            :aria-label="m.docx.toolbar.revisionDisplay"
            :disabled="disabled"
            @change="onRevisionDisplayChange"
          >
            <option value="inline">{{ m.docx.toolbar.revisionDisplayInline }}</option>
            <option value="balloons">{{ m.docx.toolbar.revisionDisplayBalloons }}</option>
            <option value="none">{{ m.docx.toolbar.revisionDisplayNone }}</option>
          </select>
        </label>
        <OfficeButton
          variant="outline"
          size="sm"
          class="docx-viewer-toolbar__toggle-button"
          data-testid="docx-show-comments"
          :aria-pressed="showComments"
          :disabled="disabled || commentCount === 0"
          :title="commentsTitle"
          @click="emit('toggleComments')"
        >
          <OfficeIcon name="message-square" />
          <span v-if="commentCount > 0" class="docx-viewer-toolbar__count">{{ commentCount }}</span>
        </OfficeButton>

        <OfficeToolbarSeparator />

        <OfficeIconButton
          variant="ghost"
          size="icon-sm"
          class="docx-viewer-toolbar__icon-button"
          data-testid="docx-theme-toggle"
          :label="isDark ? m.docx.toolbar.lightThemeAction : m.docx.toolbar.darkThemeAction"
          :title="isDark ? m.docx.toolbar.lightDocument : m.docx.toolbar.darkDocument"
          :icon="isDark ? 'sun' : 'moon'"
          :disabled="disabled"
          @click="emit('toggleTheme')"
        />
        <OfficeIconButton
          v-if="showUpload"
          variant="ghost"
          size="icon-sm"
          class="docx-viewer-toolbar__icon-button"
          data-testid="docx-upload"
          :label="m.docx.toolbar.upload"
          icon="upload"
          @click="emit('upload')"
        />
        <OfficeIconButton
          v-if="showDownload"
          variant="ghost"
          size="icon-sm"
          class="docx-viewer-toolbar__icon-button"
          data-testid="docx-download"
          :label="m.docx.toolbar.download"
          icon="download"
          :disabled="disabled || !canDownload"
          @click="emit('download')"
        />
      </OfficeToolbarGroup>
    </template>
  </OfficeToolbar>
</template>

<script setup lang="ts">
import { computed, ref } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"
import {
  OfficeButton,
  OfficeIcon,
  OfficeIconButton,
  OfficePageControl,
  OfficeToolbar,
  OfficeToolbarGroup,
  OfficeToolbarSeparator,
  OfficeZoomControl,
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
const searchInputRef = ref<HTMLInputElement>()
const searchComposing = ref(false)

function onRevisionDisplayChange(event: Event): void {
  const value = (event.target as HTMLSelectElement).value
  if (value === "inline" || value === "balloons" || value === "none") {
    emit("update:revisionDisplay", value)
  }
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
  --office-ui-toolbar-min-height: 3rem;
  --office-ui-toolbar-padding-block: var(--office-ui-space-2);
  --office-ui-toolbar-padding-inline: var(--office-ui-space-3);
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

.docx-viewer-toolbar__search-input {
  inline-size: 8.25rem;
  min-inline-size: 5rem;
}

.docx-viewer-toolbar__search-option {
  min-inline-size: var(--office-ui-height-xs);
  padding-inline: var(--office-ui-space-1);
  font-size: var(--office-text-xs);
  font-weight: 600;
}

.docx-viewer-toolbar__search-option[aria-pressed="true"] {
  background: var(--office-accent);
  border-color: var(--office-ring);
  color: var(--office-foreground);
}

.docx-viewer-toolbar__search-status {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  font-variant-numeric: tabular-nums;
  min-inline-size: 2rem;
  text-align: center;
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
  .docx-viewer-toolbar__search-input {
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
  .docx-viewer-toolbar__search-option {
    min-inline-size: 44px;
  }
}
</style>
