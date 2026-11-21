<template>
  <aside
    class="pdf-thumbnails"
    :data-testid="testId"
    :aria-label="label"
  >
    <h2 v-if="heading" class="pdf-thumbnails__title">{{ heading }}</h2>
    <div
      class="pdf-thumbnails__viewport"
      role="listbox"
      :aria-label="label"
      :aria-activedescendant="activePage > 0 ? `pdf-thumbnail-option-${activePage}` : undefined"
      tabindex="0"
      :style="{ paddingBlock: `${PANE_PADDING_Y}px` }"
    >
      <div class="pdf-thumbnails__list" :style="{ blockSize: `${totalHeight}px` }">
        <div
          v-for="meta in layout"
          :key="meta.page"
          class="pdf-thumbnails__slot"
          :style="{ insetBlockStart: `${meta.top}px`, blockSize: `${meta.wrapperHeight}px` }"
        >
          <button
            :id="`pdf-thumbnail-option-${meta.page}`"
            :data-testid="itemTestId"
            type="button"
            :data-page-number="meta.page"
            :aria-current="meta.page === activePage ? 'page' : undefined"
            :aria-label="altFor(meta.page)"
            :aria-selected="meta.page === activePage"
            :aria-posinset="meta.page"
            :aria-setsize="pages.length"
            role="option"
            :class="['thumbnail', { active: meta.page === activePage }]"
            @click="emit('select', meta.page)"
          >
            <span
              class="thumbnail__image"
              :style="{
                inlineSize: `${meta.width + IMAGE_PADDING * 2}px`,
                blockSize: `${meta.height + IMAGE_PADDING * 2}px`,
                padding: `${IMAGE_PADDING}px`,
              }"
            >
              <img
                v-if="urls[meta.page]"
                :data-testid="`${itemTestId}-image`"
                class="thumbnail__img"
                :src="urls[meta.page]"
                :alt="altFor(meta.page)"
                :style="{ inlineSize: `${meta.width}px`, blockSize: `${meta.height}px` }"
              />
              <span
                v-else
                class="thumbnail__placeholder"
                :style="{ inlineSize: `${meta.width}px`, blockSize: `${meta.height}px` }"
              >{{ labelFor(meta.page) }}</span>
            </span>
            <span class="thumbnail__label" :style="{ blockSize: `${LABEL_HEIGHT}px` }">
              <span class="thumbnail__label-text">{{ labelFor(meta.page) }}</span>
            </span>
          </button>
        </div>
      </div>
    </div>
  </aside>
</template>

<script setup lang="ts">
import { computed } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"

interface PdfThumbnailPageSize {
  width: number
  height: number
  /** Intrinsic page rotation in degrees; odd turns swap width/height. */
  rotation?: number
}

interface ThumbnailMeta {
  page: number
  width: number
  height: number
  wrapperHeight: number
  top: number
}

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
    /** Page boxes used for the thumbnail geometry; defaults to 600×800. */
    pageSizes?: readonly PdfThumbnailPageSize[]
    testId?: string
    itemTestId?: string
  }>(),
  {
    labels: () => [],
    urls: () => ({}),
    heading: undefined,
    pageSizes: () => [],
    testId: "pdf-thumbnails",
    itemTestId: "pdf-thumbnail",
  },
)

const emit = defineEmits<{
  select: [page: number]
}>()

/** Upstream PDF thumbnail geometry (pdf-viewer.tsx). */
const THUMBNAIL_PAGE_WIDTH = 92
const IMAGE_PADDING = 8
const LABEL_HEIGHT = 24
const PAGE_GAP = 12
const PANE_PADDING_Y = 16

const pages = computed(() =>
  Array.from({ length: Math.max(0, Math.floor(props.pageCount || 0)) }, (_, index) => index + 1),
)

const layout = computed<readonly ThumbnailMeta[]>(() => {
  let top = PANE_PADDING_Y
  return pages.value.map((page) => {
    const size = props.pageSizes[page - 1]
    const rotation = ((Math.round((size?.rotation ?? 0) / 90) % 4) + 4) % 4
    const pageWidth = size ? (rotation % 2 ? size.height : size.width) : 600
    const pageHeight = size ? (rotation % 2 ? size.width : size.height) : 800
    const width = THUMBNAIL_PAGE_WIDTH
    const height = Math.max(1, Math.round(width * (pageHeight / Math.max(1, pageWidth))))
    const wrapperHeight = IMAGE_PADDING + height + IMAGE_PADDING + LABEL_HEIGHT
    const meta = { page, width, height, wrapperHeight, top }
    top += wrapperHeight + PAGE_GAP
    return meta
  })
})

const totalHeight = computed(() => {
  if (layout.value.length === 0) return PANE_PADDING_Y * 2
  const last = layout.value[layout.value.length - 1]
  return last.top + last.wrapperHeight + PANE_PADDING_Y
})

function labelFor(page: number): string {
  const label = props.labels[page - 1]
  return typeof label === "string" && label.trim() ? label.trim() : String(page)
}

function altFor(page: number): string {
  return formatOfficeMessage(props.thumbnailAlt, { page: labelFor(page) })
}
</script>

<style scoped>
/*
 * Upstream `PDFViewerThumbnails` geometry: a listbox viewport with 16px
 * horizontal padding, absolutely positioned 163px thumbnail rows (for a 3:4
 * page at 92px wide), a 108×139 image well with 8px padding, and a 24px label.
 * `DocumentViewerSidebar` owns the 160px rail and its flush body.
 */
.pdf-thumbnails {
  display: flex;
  flex-direction: column;
  inline-size: 100%;
  block-size: 100%;
  min-inline-size: 0;
  min-block-size: 0;
  box-sizing: border-box;
}
.pdf-thumbnails__title {
  flex: none;
  margin: 0;
  padding: 8px 16px 0;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
}
.pdf-thumbnails__viewport {
  flex: 1 1 auto;
  min-block-size: 0;
  overflow-y: auto;
  padding-inline: 16px;
  box-sizing: border-box;
  outline: none;
}
.pdf-thumbnails__list {
  position: relative;
}
.pdf-thumbnails__slot {
  position: absolute;
  inset-inline: 0;
  display: flex;
  justify-content: center;
}
.pdf-thumbnails__slot:has(.thumbnail.active) { z-index: 1; }
.thumbnail {
  display: flex;
  inline-size: 100%;
  block-size: 100%;
  margin: 0;
  padding: 0 8px;
  align-items: center;
  justify-content: space-between;
  flex-direction: column;
  overflow: hidden;
  border: 0;
  border-radius: var(--office-radius-md);
  background: transparent;
  color: var(--office-muted-foreground);
  font: inherit;
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
  text-align: center;
  cursor: default;
  user-select: none;
  transition: box-shadow var(--office-ui-duration-fast) ease;
  box-sizing: border-box;
}
.thumbnail:hover {
  background: var(--office-sidebar-accent);
}
.thumbnail.active {
  background: var(--office-sidebar-accent);
  color: var(--office-foreground);
}
.pdf-thumbnails__viewport:focus-visible .thumbnail.active {
  box-shadow:
    0 0 0 2px var(--office-ring),
    0 0 0 3px var(--office-background);
}
.thumbnail__image {
  display: flex;
  flex: none;
  align-items: center;
  justify-content: center;
  overflow: hidden;
  border-radius: var(--office-radius-md);
  background: transparent;
  box-sizing: border-box;
}
.thumbnail__img {
  display: block;
  max-width: none;
  border-radius: var(--office-radius-sm);
  object-fit: contain;
}
.thumbnail__placeholder {
  display: grid;
  place-items: center;
  border-radius: var(--office-radius-sm);
  background: var(--office-background);
  color: var(--office-muted-foreground);
}
.thumbnail__label {
  display: flex;
  flex: none;
  align-items: center;
  justify-content: center;
  font-variant-numeric: tabular-nums;
}
.thumbnail__label-text {
  display: flex;
  min-inline-size: 1.25rem;
  align-items: center;
  justify-content: center;
  padding-inline: 6px;
  line-height: 1.25rem;
  text-align: center;
}
</style>
