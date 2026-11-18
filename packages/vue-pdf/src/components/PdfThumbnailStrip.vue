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
.pdf-thumbnails {
  width: 150px;
  flex: 0 0 150px;
  overflow: auto;
  border-right: 1px solid var(--pdf-border, #e5e5e5);
  padding: 10px;
  background: var(--pdf-panel-bg, #fafafa);
  box-sizing: border-box;
}
.pdf-thumbnails__title {
  margin: 0 0 8px;
  color: var(--pdf-muted, #737373);
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
}
.thumbnail {
  display: flex;
  width: 100%;
  min-height: 118px;
  margin-bottom: 10px;
  padding: 6px;
  align-items: center;
  justify-content: center;
  flex-direction: column;
  gap: 5px;
  overflow: hidden;
  border: 2px solid transparent;
  border-radius: 6px;
  background: transparent;
  color: var(--pdf-muted, #737373);
  font-size: 11px;
  cursor: pointer;
}
.thumbnail img {
  display: block;
  max-width: 106px;
  max-height: 86px;
  object-fit: contain;
  border: 1px solid var(--pdf-control-border, #d4d4d4);
  background: var(--pdf-control-bg, #fff);
  box-shadow: 0 1px 3px rgb(0 0 0 / 12%);
  filter: var(--pdf-page-filter, none);
}
.thumbnail-placeholder {
  display: grid;
  place-items: center;
  width: 74px;
  height: 84px;
  border: 1px solid var(--pdf-control-border, #d4d4d4);
  background: var(--pdf-control-bg, #fff);
  color: var(--pdf-muted, #737373);
}
.thumbnail.active {
  border-color: #2563eb;
  background: color-mix(in srgb, #2563eb 12%, var(--pdf-panel-bg, #fafafa));
  color: var(--pdf-link, #1d4ed8);
  font-weight: 700;
}
.thumbnail:focus-visible { outline: 2px solid #2563eb; outline-offset: 2px; }

@container pdf-viewer (max-width: 760px) {
  .pdf-thumbnails {
    width: auto;
    flex-basis: auto;
    display: flex;
    gap: 8px;
    border-right: 0;
    border-bottom: 1px solid var(--pdf-border, #e5e5e5);
  }
  .thumbnail {
    width: 108px;
    min-height: 104px;
    flex: 0 0 auto;
    margin-bottom: 0;
  }
}
</style>
