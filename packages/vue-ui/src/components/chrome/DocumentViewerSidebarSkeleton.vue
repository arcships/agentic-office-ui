<script setup lang="ts">
import { computed, useId } from "vue"
import { useOfficeUiMessages } from "../../i18n"
import OfficeSkeleton from "../ui/OfficeSkeleton.vue"
import { officeClass } from "../ui/types"

export interface DocumentViewerSidebarSkeletonProps {
  /** Number of placeholder page cards. */
  items?: number
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<DocumentViewerSidebarSkeletonProps>(), {
  items: 3,
  class: "",
})

const m = useOfficeUiMessages()
const labelId = useId()

const classes = computed(() =>
  officeClass(
    "office-ui",
    "office-document-sidebar__skeleton",
    props.class,
  ),
)
const count = computed(() => Math.max(Math.trunc(props.items) || 0, 1))
</script>

<template>
  <div
    :class="classes"
    role="status"
    :aria-labelledby="labelId"
    data-slot="sidebar-skeleton"
    data-testid="document-viewer-sidebar-skeleton"
  >
    <span :id="labelId" class="office-document-sidebar__skeleton-label">
      {{ m.common.chrome.loadingTitle }}
    </span>
    <div
      v-for="index in count"
      :key="index"
      class="office-document-sidebar__skeleton-item"
      data-slot="sidebar-skeleton-item"
    >
      <OfficeSkeleton class="office-document-sidebar__skeleton-thumb" />
      <OfficeSkeleton class="office-document-sidebar__skeleton-caption" />
    </div>
  </div>
</template>
