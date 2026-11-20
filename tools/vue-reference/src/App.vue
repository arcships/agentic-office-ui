<script setup lang="ts">
import manifest from "../../parity-fixtures/manifest.json";
import { readQuery } from "./fixtures";
import ParitySurface from "./parity/ParitySurface.vue";
import ButtonSurface from "./surfaces/ButtonSurface.vue";
import BoundingBoxCitationsSurface from "./surfaces/BoundingBoxCitationsSurface.vue";
import DialogSurface from "./surfaces/DialogSurface.vue";
import DocumentSidebarSurface from "./surfaces/DocumentSidebarSurface.vue";
import DocumentSplitsSurface from "./surfaces/DocumentSplitsSurface.vue";
import ESignatureSurface from "./surfaces/ESignatureSurface.vue";
import LayoutBlocksSurface from "./surfaces/LayoutBlocksSurface.vue";
import DocxSidebarSurface from "./surfaces/DocxSidebarSurface.vue";
import DocxToolbarSurface from "./surfaces/DocxToolbarSurface.vue";
import FileThumbnailSurface from "./surfaces/FileThumbnailSurface.vue";
import FileUploadSurface from "./surfaces/FileUploadSurface.vue";
import InputSurface from "./surfaces/InputSurface.vue";
import MissingSurface from "./surfaces/MissingSurface.vue";
import OfficeEmptyStateSurface from "./surfaces/OfficeEmptyStateSurface.vue";
import OfficeErrorStateSurface from "./surfaces/OfficeErrorStateSurface.vue";
import OfficeLoadingStateSurface from "./surfaces/OfficeLoadingStateSurface.vue";
import OfficeStatusBarSurface from "./surfaces/OfficeStatusBarSurface.vue";
import PdfSidebarSurface from "./surfaces/PdfSidebarSurface.vue";
import PdfToolbarSurface from "./surfaces/PdfToolbarSurface.vue";
import SelectSurface from "./surfaces/SelectSurface.vue";
import TooltipSurface from "./surfaces/TooltipSurface.vue";

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
    <ButtonSurface
      v-else-if="surface.id === 'button'"
      :cell-width="surface.cellWidth"
    />
    <InputSurface
      v-else-if="surface.id === 'input'"
      :cell-width="surface.cellWidth"
    />
    <SelectSurface
      v-else-if="surface.id === 'select'"
      :cell-width="surface.cellWidth"
    />
    <DialogSurface
      v-else-if="surface.id === 'dialog'"
      :cell-width="surface.cellWidth"
      :theme="theme"
    />
    <TooltipSurface
      v-else-if="surface.id === 'tooltip'"
      :cell-width="surface.cellWidth"
      :theme="theme"
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
    <PdfSidebarSurface
      v-else-if="surface.id === 'pdf-sidebar'"
      :cell-width="surface.cellWidth"
      :theme="theme"
    />
    <DocxToolbarSurface
      v-else-if="surface.id === 'docx-toolbar'"
      :cell-width="surface.cellWidth"
      :theme="theme"
    />
    <DocxSidebarSurface
      v-else-if="surface.id === 'docx-sidebar'"
      :cell-width="surface.cellWidth"
      :theme="theme"
    />
    <DocumentSidebarSurface
      v-else-if="surface.id === 'document-sidebar'"
      :cell-width="surface.cellWidth"
    />
    <OfficeEmptyStateSurface
      v-else-if="surface.id === 'office-empty-state'"
      :cell-width="surface.cellWidth"
    />
    <OfficeErrorStateSurface
      v-else-if="surface.id === 'office-error-state'"
      :cell-width="surface.cellWidth"
    />
    <OfficeLoadingStateSurface
      v-else-if="surface.id === 'office-loading-state'"
      :cell-width="surface.cellWidth"
    />
    <OfficeStatusBarSurface
      v-else-if="surface.id === 'office-status-bar'"
      :cell-width="surface.cellWidth"
    />
    <BoundingBoxCitationsSurface
      v-else-if="surface.id === 'bounding-box-citations'"
      :cell-width="surface.cellWidth"
    />
    <LayoutBlocksSurface
      v-else-if="surface.id === 'layout-blocks'"
      :cell-width="surface.cellWidth"
    />
    <ESignatureSurface
      v-else-if="surface.id === 'e-signature'"
      :cell-width="surface.cellWidth"
    />
    <DocumentSplitsSurface
      v-else-if="surface.id === 'document-splits'"
      :cell-width="surface.cellWidth"
    />
    <p v-else>No Vue renderer for {{ surface.id }}.</p>
  </ParitySurface>
</template>
