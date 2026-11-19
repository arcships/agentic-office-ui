<script setup lang="ts">
import FileUpload from "../../../../packages/vue-ui/src/components/FileUpload.vue";
import { fixtures } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";

const uploadCopy = fixtures.copy.fileUpload;
const uploadFixture = fixtures.fileUpload;

const states = [
  { state: "default", matrixState: "empty" },
  { state: "hover", matrixState: "hover" },
  { state: "focus", matrixState: "focus" },
  { state: "dragging", matrixState: "active" },
  { state: "rejected", matrixState: "error" },
  { state: "file-selected", matrixState: "default" },
];

defineProps<{ cellWidth: number }>();
</script>

<template>
  <ParityCell
    v-for="entry in states"
    :key="entry.state"
    :state="entry.state"
    :matrix-state="entry.matrixState"
    :width="cellWidth"
  >
    <FileUpload
      :accept="uploadFixture.accept"
      :accepted-file-types="uploadFixture.acceptedFileTypes"
      :browse-label="uploadCopy.browseLabel"
      :description="uploadCopy.description"
      :dragging-label="uploadCopy.draggingLabel"
      :multiple="uploadFixture.multiple"
      :max-size="uploadFixture.maxSizeBytes"
      show-file-list
      :title="uploadCopy.title"
    />
  </ParityCell>
</template>
