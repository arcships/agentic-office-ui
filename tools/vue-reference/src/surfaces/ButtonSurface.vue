<script setup lang="ts">
import { OfficeButton, OfficeIconButton } from "@arcships/vue-ui";
import { fixtures } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";

const copy = fixtures.copy.buttons;

const cells = [
  { state: "default", matrixState: "default", variant: "default", size: "default", label: copy.label },
  { state: "hover", matrixState: "hover", variant: "default", size: "default", label: copy.label },
  { state: "focus", matrixState: "focus", variant: "default", size: "default", label: copy.label },
  { state: "active", matrixState: "active", variant: "default", size: "default", label: copy.label },
  { state: "disabled", matrixState: "disabled", variant: "default", size: "default", label: copy.label, disabled: true },
  { state: "loading", matrixState: "loading", variant: "default", size: "default", label: copy.label, loading: true },
  { state: "outline", matrixState: "default", variant: "outline", size: "default", label: copy.outline },
  { state: "secondary", matrixState: "default", variant: "secondary", size: "default", label: copy.secondary },
  { state: "ghost", matrixState: "default", variant: "ghost", size: "default", label: copy.ghost },
  { state: "destructive", matrixState: "default", variant: "destructive", size: "default", label: copy.destructive },
  { state: "link", matrixState: "default", variant: "link", size: "default", label: copy.link },
  { state: "size-sm", matrixState: "default", variant: "default", size: "sm", label: copy.label },
  { state: "size-lg", matrixState: "default", variant: "default", size: "lg", label: copy.label },
] as const;

defineProps<{ cellWidth: number }>();
</script>

<template>
  <ParityCell
    v-for="cell in cells"
    :key="cell.state"
    :state="cell.state"
    :matrix-state="cell.matrixState"
    :width="cellWidth"
  >
    <OfficeButton
      :variant="cell.variant"
      :size="cell.size"
      :disabled="'disabled' in cell ? cell.disabled : false"
      :loading="'loading' in cell ? cell.loading : false"
    >
      {{ cell.label }}
    </OfficeButton>
  </ParityCell>
  <ParityCell state="icon-sm" matrix-state="default" :width="cellWidth">
    <OfficeIconButton
      :label="copy.iconAlt"
      icon="more-horizontal"
      size="icon-sm"
      variant="default"
    />
  </ParityCell>
</template>
