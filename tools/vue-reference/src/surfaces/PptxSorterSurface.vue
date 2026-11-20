<script setup lang="ts">
/**
 * Issue #161 — PPTX slide-sorter card parity surface.
 *
 * Subject: the production `PptxSlideCard` + `PptxThumbnail` pair used by the
 * sorter grid. The fake session paints the shared fixture PNG so the card
 * chrome (padding, radius, shadow, label, hover/active) can be compared with
 * the vendored upstream card; the viewer's own renderer is not involved.
 */
import type { PptxSlideInfo } from "@arcships/pptx-core";
import type { PptxPreviewSession } from "@arcships/pptx-core/browser";
import PptxSlideCard from "../../../../packages/vue-pptx/src/PptxSlideCard.vue";
import PptxThumbnail from "../../../../packages/vue-pptx/src/PptxThumbnail.vue";
import { fixtures } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";
import slideUrl from "../../../parity-fixtures/assets/pptx-slide.png";

defineProps<{ cellWidth: number; theme: string }>();

const sorter = fixtures.pptxSorter;
const cardCount = 3;
const columns = 3;

const states = [
  { state: "default", matrixState: "default", activeSlide: 1 },
  { state: "hover", matrixState: "hover", activeSlide: 1 },
  { state: "selected", matrixState: "active", activeSlide: 3 },
];

const slides: PptxSlideInfo[] = Array.from({ length: cardCount }, (_, index) => ({
  index,
  number: index + 1,
  hidden: false,
})) as PptxSlideInfo[];

/**
 * Minimal preview session: `PptxThumbnail` only needs `renderThumbnail`, and
 * the fake handle paints the shared fixture image into the canvas.
 */
function fixtureSession(): PptxPreviewSession {
  return {
    renderThumbnail(_index: number, target: HTMLElement) {
      const image = document.createElement("img");
      image.src = slideUrl;
      image.alt = "";
      image.loading = "lazy";
      image.decoding = "async";
      image.draggable = false;
      image.style.cssText = "display:block;width:100%;height:100%;object-fit:cover";
      target.replaceChildren(image);
      return { ready: Promise.resolve(), dispose() {} };
    },
  } as unknown as PptxPreviewSession;
}

const session = fixtureSession();
const aspectRatio = `${sorter.width} / ${sorter.height}`;
</script>

<template>
  <ParityCell
    v-for="entry in states"
    :key="entry.state"
    :state="entry.state"
    :matrix-state="entry.matrixState"
    :width="cellWidth"
  >
    <div
      data-pptx-sorter-grid=""
      :style="{
        display: 'grid',
        gridTemplateColumns: `repeat(${columns}, minmax(0, 1fr))`,
        gap: '12px',
        width: '100%',
      }"
    >
      <PptxSlideCard
        v-for="(slide, index) in slides"
        :key="slide.index"
        :index="slide.index"
        :number="slide.number"
        :active="slide.number === entry.activeSlide"
        :aspect-ratio="aspectRatio"
        :label="`Slide ${slide.number}`"
      >
        <PptxThumbnail
          :session="session"
          :slide="slide"
          :active="slide.number === entry.activeSlide"
          :width="sorter.thumbnailWidth"
          :show-label="false"
          tabindex="-1"
        />
      </PptxSlideCard>
    </div>
  </ParityCell>
</template>
