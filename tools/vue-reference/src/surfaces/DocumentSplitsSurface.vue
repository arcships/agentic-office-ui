<script setup lang="ts">

import fixtures from "../../../parity-fixtures/fixtures.json";
import DocumentSplits from "../../../../packages/vue-ui/src/components/DocumentSplits.vue";
import ParityCell from "../parity/ParityCell.vue";

const splits = fixtures.blocks.documentSplits;
const thumbnailImages = Object.fromEntries(
  splits.splits.flatMap((group) =>
    group.pages.map((pageId) => [pageId, splits.thumbnailImage]),
  ),
);

const states = [
  { state: "default", matrixState: "default", key: "splits" },
  { state: "hover", matrixState: "hover", key: "splits" },
  { state: "focus", matrixState: "focus", key: "splits" },
  { state: "single", matrixState: "default", key: "singleSplit" },
  { state: "empty-pages", matrixState: "empty", key: "emptyPagesSplit" },
  { state: "empty", matrixState: "empty", key: "none" },
  { state: "loading", matrixState: "loading", key: "none" },
  { state: "error", matrixState: "error", key: "none" },
];

defineProps<{ cellWidth: number }>();

// Lazy <img> decoding must not race the capture: warm the shared page image
// before either harness renders a thumbnail (same URL on both sides).
if (typeof window !== "undefined") {
  const preload = new Image();
  preload.src = splits.thumbnailImage;
}

function stateSplits(key: string) {
  if (key === "none") return [];
  return splits[key as "splits" | "singleSplit" | "emptyPagesSplit"];
}
</script>

<template>
  <ParityCell
    v-for="entry in states"
    :key="entry.state"
    :state="entry.state"
    :matrix-state="entry.matrixState"
    :width="cellWidth"
  >
    <div :style="{ height: '640px' }">
      <DocumentSplits
        :splits="stateSplits(entry.key)"
        :thumbnail-images="thumbnailImages"
        :with-frame-divider="false"
        :loading="entry.state === 'loading'"
        :error="
          entry.state === 'error'
            ? 'Document splits could not be loaded.'
            : undefined
        "
      />
    </div>
  </ParityCell>
</template>
