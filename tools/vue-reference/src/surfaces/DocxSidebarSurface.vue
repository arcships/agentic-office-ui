<script setup lang="ts">
import { fixtures, type ParityTheme } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";
import DocumentViewerSidebar from "../../../../packages/vue-ui/src/components/chrome/DocumentViewerSidebar.vue";
import FileThumbnail from "../../../../packages/vue-ui/src/components/FileThumbnail.vue";

const props = defineProps<{ cellWidth: number; theme: ParityTheme }>();
const sidebar = fixtures.documentSidebar;
const pages = fixtures.documents.filter((entry) => "pages" in entry && entry.pages);

const states = [
  { state: "open", matrixState: "default", open: true, loading: false },
  { state: "closed", matrixState: "default", open: false, loading: false },
  { state: "skeleton", matrixState: "loading", open: true, loading: true },
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
      :style="{
        width: `${sidebar.widthPx}px`,
        height: entry.loading ? '240px' : undefined,
        minHeight: '240px',
        overflow: 'hidden',
        position: 'relative',
        background: 'var(--office-sidebar)',
      }"
    >
      <DocumentViewerSidebar
        :open="entry.open"
        mode="inline"
        :width="sidebar.widthPx"
        :loading="entry.loading"
        label="Document pages"
        :style="entry.loading ? { borderInlineEndWidth: '0' } : undefined"
      >
        <div style="display:flex;flex-direction:column;align-items:center;gap:12px;padding:8px;overflow:hidden">
          <div
            v-for="(page, index) in pages"
            :key="page.id"
            style="display:flex;flex-direction:column;align-items:center;gap:4px"
          >
            <FileThumbnail
              :file="{ name: page.name, type: page.type }"
              :size="80"
              style="height:102px;border-radius:6px"
            />
            <span style="font-size:12px;color:var(--office-muted-foreground)">{{ index + 1 }}</span>
          </div>
        </div>
      </DocumentViewerSidebar>
    </div>
  </ParityCell>
</template>
