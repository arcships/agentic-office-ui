<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import {
  Download01Icon,
  FilePenIcon,
  Pen01Icon,
} from "@hugeicons/core-free-icons"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { HugeIcon } from "./HugeIcon"
import OfficeButton from "./ui/OfficeButton.vue"
import { useOfficeUiMessages } from "../i18n"
import type { ESignatureField, ESignatureFieldsProps } from "../types"

const props = withDefaults(defineProps<ESignatureFieldsProps>(), {
  canExport: false,
  exporting: false,
  showPreview: false,
  title: undefined,
  description: undefined,
  loading: false,
  error: undefined,
  className: "",
})

const emit = defineEmits<{
  sign: [field: ESignatureField]
  edit: [field: ESignatureField]
  clear: [field: ESignatureField]
  export: [field: ESignatureField]
}>()

const m = useOfficeUiMessages()

const heading = computed(
  () => props.title ?? m.value.vueUi.signaturePad.fieldsTitle,
)
const description = computed(
  () => props.description ?? m.value.vueUi.signaturePad.fieldsDescription,
)
const activeExportField = computed(
  () => props.fields.find((field) => isSigned(field)) ?? props.fields[0],
)

function isSigned(field: ESignatureField): boolean {
  return Boolean(field.imageDataUrl)
}

function metaLabel(field: ESignatureField): string {
  return formatOfficeMessage(m.value.vueUi.signaturePad.sizeOnPage, {
    width: Math.round(field.bbox.width),
    height: Math.round(field.bbox.height),
    page: field.page,
  })
}

function activate(field: ESignatureField) {
  if (isSigned(field)) emit("edit", field)
  else emit("sign", field)
}

function exportSigned(field: ESignatureField) {
  if (!props.canExport || props.exporting) return
  emit("export", field)
}
</script>

<template>
  <aside :class="['e-signature-fields', 'office-ui', props.className]" data-slot="e-signature-fields">
    <div class="e-signature-fields__scroll">
      <div class="e-signature-fields__content">
        <div class="e-signature-fields__heading">
          <h3 class="e-signature-fields__title">{{ heading }}</h3>
          <p class="e-signature-fields__description">{{ description }}</p>
        </div>

        <div v-if="props.error" class="e-signature-fields__state e-signature-fields__state--error" role="alert">
          {{ props.error }}
        </div>
        <div v-else-if="props.loading" class="e-signature-fields__state" role="status">
          {{ m.vueUi.signaturePad.loading }}
        </div>
        <div v-else-if="props.fields.length === 0" class="e-signature-fields__state" role="status">
          {{ m.vueUi.signaturePad.unsigned }}
        </div>

        <template v-else>
          <div
            v-for="field in props.fields"
            :key="field.id"
            class="e-signature-field"
            :data-signature-field-id="field.id"
            data-slot="e-signature-field"
          >
          <div class="e-signature-field__row">
            <div class="e-signature-field__icon">
              <HugeIcon :icon="FilePenIcon" class="e-signature-field__glyph" />
            </div>
            <div class="e-signature-field__body">
              <div class="e-signature-field__head">
                <div class="e-signature-field__label">{{ field.label }}</div>
                <div
                  :class="[
                    'e-signature-field__badge',
                    isSigned(field) ? 'is-signed' : 'is-unsigned',
                  ]"
                  data-slot="e-signature-field-status"
                >
                  {{ isSigned(field) ? m.vueUi.signaturePad.signed : m.vueUi.signaturePad.unsigned }}
                </div>
              </div>
              <div class="e-signature-field__meta">{{ metaLabel(field) }}</div>
              <div v-if="props.showPreview && field.imageDataUrl" class="e-signature-field__preview">
                <img
                  :src="field.imageDataUrl"
                  :alt="m.vueUi.signaturePad.previewAlt"
                  class="e-signature-field__preview-image"
                />
              </div>
              <div class="e-signature-field__actions">
                <OfficeButton
                  size="sm"
                  :variant="isSigned(field) ? 'outline' : 'default'"
                  class="e-signature-field__action"
                  data-slot="e-signature-field-sign"
                  @click="activate(field)"
                >
                  <HugeIcon :icon="Pen01Icon" class="e-signature-field__action-icon" />
                  {{ isSigned(field) ? m.vueUi.signaturePad.edit : m.vueUi.signaturePad.sign }}
                </OfficeButton>
                <OfficeButton
                  v-if="isSigned(field)"
                  size="sm"
                  variant="outline"
                  data-slot="e-signature-field-clear"
                  @click="emit('clear', field)"
                >
                  {{ m.vueUi.signaturePad.clear }}
                </OfficeButton>
              </div>
            </div>
          </div>
          </div>
        </template>

        <OfficeButton
          class="e-signature-fields__export"
          :disabled="!props.canExport || props.exporting"
          data-slot="e-signature-export"
          @click="exportSigned(activeExportField!)"
        >
          <HugeIcon :icon="Download01Icon" class="e-signature-field__action-icon" />
          {{ props.exporting ? m.vueUi.signaturePad.exporting : m.vueUi.signaturePad.export }}
        </OfficeButton>
      </div>
    </div>
  </aside>
</template>

<style scoped>
.e-signature-fields {
  display: flex;
  min-height: 0;
  flex-direction: column;
  background: var(--office-background);
  color: var(--office-foreground);
}
.e-signature-fields,
.e-signature-fields * {
  box-sizing: border-box;
}
.e-signature-fields__scroll {
  min-height: 0;
  flex: 1 1 auto;
  overflow-y: auto;
}
.e-signature-fields__content {
  padding: var(--office-ui-space-4);
}
.e-signature-fields__content > * + * {
  margin-top: var(--office-ui-space-4);
}
.e-signature-fields__heading {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-1);
}
.e-signature-fields__title {
  margin: 0;
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  font-weight: 500;
}
.e-signature-fields__description {
  margin: 0;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
.e-signature-fields__state {
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: color-mix(in oklab, var(--office-muted) 30%, transparent);
  padding: var(--office-ui-space-3);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
}
.e-signature-fields__state--error {
  border-color: color-mix(in oklab, var(--office-destructive) 45%, transparent);
  background: color-mix(in oklab, var(--office-destructive) 8%, transparent);
  color: var(--office-destructive);
}
.e-signature-field {
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-background);
  padding: var(--office-ui-space-3);
}
.e-signature-field__row {
  display: flex;
  align-items: flex-start;
  gap: var(--office-ui-space-3);
}
.e-signature-field__icon {
  display: grid;
  width: 36px;
  height: 36px;
  flex-shrink: 0;
  place-items: center;
  border-radius: var(--office-radius-md);
  background: color-mix(in oklab, oklch(62.3% .214 259.815) 10%, transparent);
  color: oklch(54.6% .245 262.881);
}
.e-signature-field__glyph {
  width: 16px;
  height: 16px;
}
.e-signature-field__body {
  min-width: 0;
  flex: 1 1 auto;
}
.e-signature-field__head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--office-ui-space-2);
}
.e-signature-field__label {
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  font-weight: 500;
}
.e-signature-field__badge {
  border-radius: calc(infinity * 1px);
  padding: var(--office-ui-space-0-5) var(--office-ui-space-2);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
.e-signature-field__badge.is-signed {
  background: color-mix(in oklab, oklch(69.6% .17 162.48) 10%, transparent);
  color: oklch(59.6% .145 163.225);
}
.e-signature-field__badge.is-unsigned {
  background: var(--office-muted);
  color: var(--office-muted-foreground);
}
.e-signature-field__meta {
  margin-top: var(--office-ui-space-2);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
.e-signature-field__preview {
  margin-top: var(--office-ui-space-3);
  overflow: hidden;
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: #fff;
}
.e-signature-field__preview-image {
  display: block;
  width: 100%;
  height: 58px;
  object-fit: contain;
}
.e-signature-field__actions {
  display: flex;
  gap: var(--office-ui-space-2);
  margin-top: var(--office-ui-space-3);
}
.e-signature-field__action {
  flex: 1 1 auto;
}
.e-signature-field__action-icon {
  width: 16px;
  height: 16px;
  flex-shrink: 0;
}
.e-signature-fields__export {
  width: 100%;
}
</style>
