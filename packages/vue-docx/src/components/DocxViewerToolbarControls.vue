<template>
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
        <OfficeIcon name="list" />
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
        <OfficeIcon name="chevrons-up-down" />
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
        <OfficeIcon name="file" />
      </OfficeButton>
    </OfficeToolbarGroup>

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
      <OfficeIcon name="clock" />
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
      <OfficeIcon name="file-text" />
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

<script setup lang="ts">
import { computed } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"
import {
  OfficeButton,
  OfficeIcon,
  OfficeIconButton,
  OfficeToolbarGroup,
  OfficeToolbarSeparator,
} from "@arcships/vue-ui"
import { useDocxMessages } from "../i18n"

const props = withDefaults(
  defineProps<{
    canDownload: boolean
    disabled: boolean
    isDark: boolean
    showTrackedChanges: boolean
    revisionDisplay: import("@arcships/docx-core").DocxRevisionDisplay
    showComments: boolean
    trackedChangeCount: number
    commentCount: number
    viewMode?: import("../composables/docx-view-mode").DocxViewMode
    fitMode?: import("../composables/docx-view-mode").DocxFitMode | "none"
    showUpload?: boolean
    showDownload?: boolean
  }>(),
  {
    showUpload: true,
    showDownload: true,
    viewMode: "continuous",
    fitMode: "none",
  },
)

const emit = defineEmits<{
  download: []
  upload: []
  toggleTheme: []
  toggleTrackedChanges: []
  toggleComments: []
  "update:revisionDisplay": [display: import("@arcships/docx-core").DocxRevisionDisplay]
  "update:viewMode": [mode: import("../composables/docx-view-mode").DocxViewMode]
  "update:fitMode": [mode: import("../composables/docx-view-mode").DocxFitMode | "none"]
}>()

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
function onRevisionDisplayChange(event: Event): void {
  const value = (event.target as HTMLSelectElement).value
  if (value === "inline" || value === "balloons" || value === "none") {
    emit("update:revisionDisplay", value)
  }
}

</script>

<style scoped>
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
</style>
