<script setup lang="ts">
import { computed, useId } from "vue"
import { useOfficeUiMessages } from "../../i18n"
import OfficeSkeleton from "../ui/OfficeSkeleton.vue"
import { officeClass } from "../ui/types"

export interface DocumentViewerSidebarSkeletonProps {
  /** Number of placeholder page cards. */
  items?: number
  /**
   * `list` renders the generic multi-card skeleton; `thumbnail` mirrors the
   * upstream PDF sidebar skeleton (one 80×112 page block and a caption).
   */
  variant?: "list" | "thumbnail"
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<DocumentViewerSidebarSkeletonProps>(), {
  items: 3,
  variant: "list",
  class: "",
})

const m = useOfficeUiMessages()
const labelId = useId()

const classes = computed(() =>
  officeClass(
    "office-ui",
    "office-document-sidebar__skeleton",
    props.variant === "thumbnail" && "office-document-sidebar__skeleton--thumbnail",
    props.class,
  ),
)
const count = computed(() => Math.max(Math.trunc(props.items) || 0, 1))
</script>

<template>
  <div
    :class="classes"
    role="status"
    :aria-labelledby="props.variant === 'thumbnail' ? undefined : labelId"
    data-slot="sidebar-skeleton"
    data-testid="document-viewer-sidebar-skeleton"
  >
    <span v-if="props.variant !== 'thumbnail'" :id="labelId" class="office-document-sidebar__skeleton-label">
      {{ m.common.chrome.loadingTitle }}
    </span>
    <template v-if="props.variant === 'thumbnail'">
      <div class="office-document-sidebar__thumbnail-skeleton-thumb" data-slot="sidebar-skeleton-item">
        <span class="office-document-sidebar__thumbnail-skeleton-pulse" />
      </div>
      <div class="office-document-sidebar__thumbnail-skeleton-caption" />
    </template>
    <div
      v-for="index in props.variant === 'thumbnail' ? 0 : count"
      :key="index"
      class="office-document-sidebar__skeleton-item"
      data-slot="sidebar-skeleton-item"
    >
      <OfficeSkeleton class="office-document-sidebar__skeleton-thumb" />
      <OfficeSkeleton class="office-document-sidebar__skeleton-caption" />
    </div>
  </div>
</template>
