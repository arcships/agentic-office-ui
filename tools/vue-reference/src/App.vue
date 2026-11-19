<script setup lang="ts">
import manifest from "../../parity-fixtures/manifest.json";
import { readQuery } from "./fixtures";
import ParitySurface from "./parity/ParitySurface.vue";
import FileThumbnailSurface from "./surfaces/FileThumbnailSurface.vue";
import FileUploadSurface from "./surfaces/FileUploadSurface.vue";
import MissingSurface from "./surfaces/MissingSurface.vue";
import PdfToolbarSurface from "./surfaces/PdfToolbarSurface.vue";

type Surface = (typeof manifest.surfaces)[number];

const { surface: surfaceParam, theme } = readQuery();

const selected = manifest.surfaces.filter(
  (surface: Surface) =>
    surface.referenceStatus !== "fixtures-only" &&
    (!surfaceParam || surface.id === surfaceParam),
);

function subjectStatus(surface: Surface): "available" | "missing" {
  return surface.subjectStatus === "available" ? "available" : "missing";
}
</script>

<template>
  <ParitySurface
    v-for="surface in selected"
    :key="surface.id"
    :surface="surface.id"
    :theme="theme"
  >
    <MissingSurface
      v-if="subjectStatus(surface) === 'missing'"
      :cell-width="surface.cellWidth"
      :note="surface.notes"
      :pending-issue="surface.pendingIssue ?? '#unknown'"
      :states="surface.states"
    />
    <FileUploadSurface
      v-else-if="surface.id === 'file-upload'"
      :cell-width="surface.cellWidth"
    />
    <FileThumbnailSurface
      v-else-if="surface.id === 'file-thumbnail'"
      :cell-width="surface.cellWidth"
    />
    <PdfToolbarSurface
      v-else-if="surface.id === 'pdf-toolbar'"
      :cell-width="surface.cellWidth"
      :theme="theme"
    />
    <p v-else>No Vue renderer for {{ surface.id }}.</p>
  </ParitySurface>
</template>
