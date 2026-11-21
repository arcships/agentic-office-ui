<template>
  <aside
    class="pptx-presenter"
    data-testid="pptx-presenter"
    data-pptx-presenter="true"
    :data-presenter-mode="present ? 'present' : 'browse'"
    :aria-label="present ? m.pptx.viewer.notes.presenter : m.pptx.viewer.notes.panel"
  >
    <OfficeCard class="pptx-presenter__card">
      <div v-if="present" class="pptx-presenter__meta">
        <p class="pptx-presenter__label">{{ m.pptx.viewer.notes.next }}</p>
        <PptxThumbnail
          v-if="nextSlide && session"
          class="pptx-presenter__next"
          :session="session"
          :slide="nextSlide"
          :width="thumbnailWidth"
          tabindex="-1"
        />
        <span v-else class="pptx-presenter__empty" data-testid="pptx-presenter-last">
          {{ m.pptx.viewer.notes.noNext }}
        </span>
        <div class="pptx-presenter__timer-row">
          <OfficeIcon name="clock" class="pptx-presenter__timer-icon" />
          <span
            class="pptx-presenter__timer"
            data-testid="pptx-presenter-timer"
            role="timer"
            :aria-label="m.pptx.viewer.notes.timer"
          >{{ elapsedText }}</span>
        </div>
      </div>
      <OfficeSeparator v-if="present" class="pptx-presenter__separator" />
      <div class="pptx-presenter__body" data-testid="pptx-presenter-notes" data-pptx-notes="true">
        <p class="pptx-presenter__label">{{ m.pptx.viewer.notes.title }}</p>
        <OfficeScrollArea class="pptx-presenter__scroll">
          <slot name="notes" :notes="notesText" :slide="slide?.index ?? 0">
            <p
              v-for="(paragraph, index) in paragraphs"
              :key="index"
              class="pptx-presenter__paragraph"
              data-testid="pptx-presenter-paragraph"
            >{{ paragraph }}</p>
            <p v-if="paragraphs.length === 0" class="pptx-presenter__empty" data-testid="pptx-presenter-empty">
              {{ m.pptx.viewer.notes.empty }}
            </p>
          </slot>
        </OfficeScrollArea>
      </div>
    </OfficeCard>
  </aside>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type { PptxSlideInfo } from "@arcships/pptx-core"
import type { PptxPreviewSession } from "@arcships/pptx-core/browser"
import {
  OfficeCard,
  OfficeIcon,
  OfficeScrollArea,
  OfficeSeparator,
} from "@arcships/vue-ui"
import { usePptxMessages } from "./i18n"
import PptxThumbnail from "./PptxThumbnail.vue"
import { formatPptxPresenterElapsed, usePptxPresenterClock } from "./composables/presenter-clock"

const props = withDefaults(defineProps<{
  session: PptxPreviewSession | null
  slide?: PptxSlideInfo | null
  nextSlide?: PptxSlideInfo | null
  /** Presenter rail: adds the next-slide preview and the elapsed-time readout. */
  present?: boolean
  /** Runs the elapsed-time clock; false keeps it at `00:00`. */
  running?: boolean
  thumbnailWidth?: number
}>(), {
  slide: null,
  nextSlide: null,
  present: false,
  running: false,
  thumbnailWidth: 160,
})

const m = usePptxMessages()
const elapsed = usePptxPresenterClock(() => props.present && props.running)
const elapsedText = computed(() => formatPptxPresenterElapsed(elapsed.value))
const notesText = computed(() => props.slide?.notes ?? "")
/** One paragraph per line, matching how notes are stored on the slide info. */
const paragraphs = computed(() => notesText.value.split("\n").filter((line) => line.trim().length > 0))
</script>
