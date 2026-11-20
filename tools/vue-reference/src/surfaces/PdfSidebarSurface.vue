<script setup lang="ts">
import { computed } from "vue";
import DocumentViewerSidebar from "../../../../packages/vue-ui/src/components/chrome/DocumentViewerSidebar.vue";
import DocumentViewerSidebarSkeleton from "../../../../packages/vue-ui/src/components/chrome/DocumentViewerSidebarSkeleton.vue";
import PdfThumbnailStrip from "../../../../packages/vue-pdf/src/components/PdfThumbnailStrip.vue";
import { fixtures } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";
import thumbnailUrl from "../../../parity-fixtures/assets/pdf-thumbnail.png";

const props = defineProps<{ cellWidth: number; theme: string }>();

const sidebar = fixtures.documentSidebar;
const frameHeight = 720;
const pageCount = 4;
const pageSize = { width: 600, height: 800 };
const labels = ["1", "2", "3", "4"];
const pageSizes = Array.from({ length: pageCount }, () => pageSize);
const urls = Object.fromEntries(
  Array.from({ length: pageCount }, (_, index) => [index + 1, thumbnailUrl]),
);

const inline = computed(() => window.innerWidth >= 768);

const states = [
  { state: "open", matrixState: "default", open: true, skeleton: false },
  { state: "closed", matrixState: "default", open: false, skeleton: false },
  { state: "skeleton", matrixState: "loading", open: true, skeleton: true },
];
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
      data-pdf-sidebar-frame=""
      :style="{
        width: `${sidebar.widthPx}px`,
        height: `${frameHeight}px`,
        overflow: 'hidden',
        position: 'relative',
      }"
    >
      <DocumentViewerSidebarSkeleton
        v-if="entry.skeleton"
        variant="thumbnail"
      />
      <DocumentViewerSidebar
        v-else
        :open="entry.open"
        :mode="inline ? 'inline' : 'overlay'"
        :breakpoint="768"
        :width="sidebar.widthPx"
        :close-on-escape="false"
        flush
        label="Page thumbnails"
        class="pdf-sidebar"
      >
        <PdfThumbnailStrip
          :page-count="pageCount"
          :labels="labels"
          :urls="urls"
          :active-page="1"
          :page-sizes="pageSizes"
          label="Page thumbnails"
          thumbnail-alt="Page {page} thumbnail"
          page-label="Page {page}"
        />
      </DocumentViewerSidebar>
    </div>
  </ParityCell>
</template>
