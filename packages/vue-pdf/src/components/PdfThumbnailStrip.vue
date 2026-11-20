<template>
  <aside
    class="pdf-thumbnails"
    :data-testid="testId"
    :aria-label="label"
  >
    <h2 v-if="heading" class="pdf-thumbnails__title">{{ heading }}</h2>
    <button
      v-for="page in pages"
      :key="page"
      :data-testid="itemTestId"
      type="button"
      :data-page-number="page"
      :aria-current="page === activePage ? 'page' : undefined"
      :class="['thumbnail', { active: page === activePage }]"
      @click="emit('select', page)"
    >
      <img
        v-if="urls[page]"
        :data-testid="`${itemTestId}-image`"
        :src="urls[page]"
        :alt="altFor(page)"
      />
      <span v-else class="thumbnail-placeholder">{{ labelFor(page) }}</span>
      <span>{{ labelTextFor(page) }}</span>
    </button>
  </aside>
</template>

<script setup lang="ts">
import { computed } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"

const props = withDefaults(
  defineProps<{
    /** Number of pages to show. */
    pageCount: number
    /** Page labels indexed from zero; missing entries fall back to the page number. */
    labels?: readonly string[]
    /** Thumbnail object URLs keyed by 1-based page number. */
    urls?: Record<number, string>
    /** Active page number (1-based). */
    activePage: number
    /** Accessible name of the strip. */
    label: string
    /** Template with a `{page}` placeholder. */
    thumbnailAlt: string
    /** Template with a `{page}` placeholder. */
    pageLabel: string
    /** Optional visible heading, rendered above the thumbnails. */
    heading?: string
    testId?: string
    itemTestId?: string
  }>(),
  {
    labels: () => [],
    urls: () => ({}),
    heading: undefined,
    testId: "pdf-thumbnails",
    itemTestId: "pdf-thumbnail",
  },
)

const emit = defineEmits<{
  select: [page: number]
}>()

const pages = computed(() =>
  Array.from({ length: Math.max(0, Math.floor(props.pageCount || 0)) }, (_, index) => index + 1),
)

function labelFor(page: number): string {
  const label = props.labels[page - 1]
  return typeof label === "string" && label.trim() ? label.trim() : String(page)
}

function labelTextFor(page: number): string {
  return formatOfficeMessage(props.pageLabel, { page: labelFor(page) })
}

function altFor(page: number): string {
  return formatOfficeMessage(props.thumbnailAlt, { page: labelFor(page) })
}
</script>

<style scoped>
/*
 * The strip is the sidebar body, so it fills its host instead of owning a
 * width. `DocumentViewerSidebar` owns the 160px rail, inline/overlay mode and
 * the loading skeleton; this file keeps the page buttons and their active /
 * focus states.
 */
.pdf-thumbnails {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-1);
  width: 100%;
  min-width: 0;
  box-sizing: border-box;
}
.pdf-thumbnails__title {
  margin: 0 0 var(--office-ui-space-1);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
}
.thumbnail {
  display: flex;
  width: 100%;
  min-height: 118px;
  margin-bottom: var(--office-ui-space-1);
  padding: var(--office-ui-space-1);
  align-items: center;
  justify-content: center;
  flex-direction: column;
  gap: var(--office-ui-space-1);
  overflow: hidden;
  border: 2px solid transparent;
  border-radius: var(--office-radius-md);
  background: transparent;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  cursor: pointer;
}
.thumbnail:hover { background: color-mix(in oklab, var(--office-foreground) 6%, transparent); }
.thumbnail img {
  display: block;
  max-width: 106px;
  max-height: 86px;
  object-fit: contain;
  border: var(--office-border-width) solid var(--office-border);
  background: var(--office-background);
  box-shadow: var(--office-shadow-xs);
  filter: var(--pdf-page-filter, none);
}
.thumbnail-placeholder {
  display: grid;
  place-items: center;
  width: 74px;
  height: 84px;
  border: var(--office-border-width) solid var(--office-border);
  background: var(--office-background);
  color: var(--office-muted-foreground);
}
.thumbnail.active {
  border-color: var(--office-ring);
  background: color-mix(in oklab, var(--office-primary) 10%, transparent);
  color: var(--office-primary);
  font-weight: 700;
}
.thumbnail:focus-visible {
  outline: var(--office-ring-width) solid var(--office-ring);
  outline-offset: 2px;
}
</style>
