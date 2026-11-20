<script setup lang="ts">
/**
 * Slide card used by the slide-sorter grid (issue #161b).
 *
 * Markup source: the pinned Extend PPTX thumbnail rail item
 * (`PptxSidebarThumbnail` + the option wrapper in `PptxThumbnailSidebarList`,
 * apps/v4/components/extend/pptx-viewer.tsx) and `FileThumbnail`:
 *
 *   div[role=option]  flex flex-col items-center gap-2 rounded-md p-2 text-xs
 *     FileThumbnail   w-full rounded-sm border-0 shadow-xs ring-0
 *                     (active: shadow-sm; preview: rounded-sm bg-white)
 *
 * The slide canvas is the default slot so the sorter can keep rendering
 * through `PptxThumbnail`/the document session.
 */
import { computed } from "vue"
import { OfficeBadge } from "@arcships/vue-ui"

export interface PptxSlideCardProps {
  /** Zero-based slide index; kept on `data-slide-index` for hosts and tests. */
  index: number
  /** One-based slide number shown under the preview. */
  number: number
  /** Current slide: upstream `bg-sidebar-accent text-foreground shadow-sm`. */
  active?: boolean
  /** Hidden slide marker (PPTX-only state; upstream has no hidden slides). */
  hidden?: boolean
  /** CSS `aspect-ratio` value, for example `16 / 9`. */
  aspectRatio?: string
  /** Accessible name of the option. */
  label?: string
  /** Copy of the hidden marker. */
  hiddenLabel?: string
  /** `data-testid` for hosts and DOM-free tests. */
  testId?: string
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<PptxSlideCardProps>(), {
  active: false,
  hidden: false,
  aspectRatio: "16 / 9",
  label: undefined,
  hiddenLabel: "Hidden",
  testId: "pptx-sorter-item",
  class: "",
})

const emit = defineEmits<{ select: [index: number] }>()

const frameStyle = computed(() => ({ aspectRatio: props.aspectRatio }))
</script>

<template>
  <div
    class="pptx-slide-card"
    :class="[props.class, props.active && 'pptx-slide-card--active']"
    role="option"
    :aria-label="props.label ?? `Slide ${props.number}`"
    :aria-selected="props.active"
    :data-slide-index="props.index"
    :data-active="props.active ? 'true' : 'false'"
    :data-testid="props.testId"
    @click="emit('select', props.index)"
  >
    <span class="pptx-slide-card__frame" :style="frameStyle">
      <slot />
    </span>
    <span class="pptx-slide-card__label">
      <span class="pptx-slide-card__number">{{ props.number }}</span>
      <OfficeBadge
        v-if="props.hidden"
        variant="secondary"
        class="pptx-slide-card__hidden"
        data-testid="pptx-thumbnail-hidden"
      >{{ props.hiddenLabel }}</OfficeBadge>
    </span>
  </div>
</template>

<style scoped>
.pptx-slide-card,
.pptx-slide-card * {
  box-sizing: border-box;
}

.pptx-slide-card {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 8px;
  inline-size: 100%;
  padding: 8px;
  border-radius: var(--office-radius-md);
  color: var(--office-muted-foreground);
  font-family: inherit;
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
  cursor: default;
  transition: background-color 150ms ease, color 150ms ease;
  user-select: none;
  outline: none;
}

.pptx-slide-card:hover {
  background: var(--office-sidebar-accent);
}

.pptx-slide-card--active {
  background: var(--office-sidebar-accent);
  color: var(--office-foreground);
}

/* Upstream FileThumbnail: rounded-sm, border-0, shadow-xs, white preview. */
.pptx-slide-card__frame {
  position: relative;
  display: block;
  inline-size: 100%;
  overflow: hidden;
  border-radius: var(--office-radius-sm);
  background: var(--office-background);
  box-shadow: var(--office-shadow-xs);
  transition: box-shadow 150ms ease;
}

.pptx-slide-card--active .pptx-slide-card__frame {
  box-shadow: var(--office-shadow-sm);
}

/* The thumbnail supplies the canvas only; the card owns the frame/label.
   The card frame and this inner clip both use `rounded-sm`, matching the
   nested FileThumbnail outer/inner rounded clips upstream. */
.pptx-slide-card__frame :deep(.pptx-thumbnail) {
  display: block;
  inline-size: 100%;
  block-size: 100%;
  margin: 0;
  padding: 0;
  border: 0;
  border-radius: var(--office-radius-sm);
  background: #fff;
  overflow: hidden;
  text-align: start;
}

.pptx-slide-card__frame :deep(.pptx-thumbnail__frame) {
  inline-size: 100%;
  block-size: 100%;
}

.pptx-slide-card__frame :deep(.pptx-thumbnail__canvas) {
  block-size: 100%;
  min-block-size: 0;
  background: transparent;
  overflow: visible;
}

.pptx-slide-card__frame :deep(.pptx-thumbnail__skeleton) {
  border-radius: 0;
}

.pptx-slide-card__label {
  display: inline-flex;
  align-items: center;
  gap: 6px;
}

.pptx-slide-card__hidden {
  font-size: var(--office-text-xs);
}
</style>
