<script setup lang="ts">
/**
 * Issue #158 — generic `DocumentViewerSidebar` subject.
 *
 * Mirrors `tools/extend-reference/src/surfaces/DocumentSidebarSurface.tsx`: the
 * pinned Extend `DocumentViewerThumbnailSidebar` (160px `w-40` rail) hosting a
 * two-page fixture strip and the upstream `p-4` loading skeleton. States cover
 * inline open/closed/skeleton plus overlay open/closed, which is where the
 * 160px width, the `-ml-40` inline collapse, the `-translate-x-full` overlay
 * collapse and the `shadow-lg` edge are compared. The fixture strip is viewer
 * state rather than a sidebar primitive; each viewer keeps its own thumbnails.
 */
import DocumentViewerSidebar from "../../../../packages/vue-ui/src/components/chrome/DocumentViewerSidebar.vue";
import DocumentViewerSidebarSkeleton from "../../../../packages/vue-ui/src/components/chrome/DocumentViewerSidebarSkeleton.vue";
import FileThumbnail from "../../../../packages/vue-ui/src/components/FileThumbnail.vue";
import { fixtures } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";

const sidebar = fixtures.documentSidebar;
const pages = fixtures.documents.filter((entry) => entry.pages);
const frameHeight = 240;

defineProps<{ cellWidth: number }>();

const states = [
  { state: "open", matrixState: "default", open: true, overlay: false, skeleton: false },
  { state: "closed", matrixState: "default", open: false, overlay: false, skeleton: false },
  { state: "skeleton", matrixState: "loading", open: true, overlay: false, skeleton: true },
  { state: "overlay-open", matrixState: "default", open: true, overlay: true, skeleton: false },
  { state: "overlay-closed", matrixState: "default", open: false, overlay: true, skeleton: false },
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
      data-document-sidebar-frame=""
      :style="{
        width: `${entry.overlay ? cellWidth : sidebar.widthPx}px`,
        height: `${frameHeight}px`,
        overflow: 'hidden',
        position: 'relative',
        background: entry.skeleton
          ? undefined
          : entry.overlay
            ? 'var(--office-muted)'
            : 'var(--office-sidebar)',
      }"
    >
      <DocumentViewerSidebarSkeleton
        v-if="entry.skeleton"
        variant="thumbnail"
      />
      <DocumentViewerSidebar
        v-else
        :open="entry.open"
        :mode="entry.overlay ? 'overlay' : 'inline'"
        :width="sidebar.widthPx"
        :close-on-escape="false"
        flush
        label="Document pages"
      >
        <div class="document-sidebar-strip">
          <div
            v-for="(page, index) in pages"
            :key="page.id"
            class="document-sidebar-strip__page"
          >
            <FileThumbnail
              class="document-sidebar-strip__thumb"
              :file="{ name: page.name, type: page.type }"
              :size="80"
            />
            <span class="document-sidebar-strip__label">{{ index + 1 }}</span>
          </div>
        </div>
      </DocumentViewerSidebar>
    </div>
  </ParityCell>
</template>

<style scoped>
/* Upstream content markup: `flex h-full flex-col gap-3 overflow-hidden p-2`
 * with `flex flex-col items-center gap-1` cards and a `text-xs` label. */
.document-sidebar-strip {
  display: flex;
  flex-direction: column;
  gap: 12px;
  height: 100%;
  overflow: hidden;
  padding: 8px;
}

.document-sidebar-strip__page {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 4px;
}

/* Upstream `FileThumbnail className="h-28 w-20 rounded-md"`. */
.document-sidebar-strip__thumb {
  height: 112px;
  border-radius: var(--office-radius-md);
}

.document-sidebar-strip__label {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
</style>
