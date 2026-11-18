<template>
  <div
    v-if="open"
    class="docx-style-preview"
    :class="{ 'docx-style-preview--dark': documentTheme === 'dark' }"
    role="listbox"
    :aria-label="m.docx.stylePreview.label"
    data-testid="docx-style-preview"
    :data-docx-style-preview="true"
    :style="popoverStyle"
    @pointerenter="emit('pointerenter')"
    @pointerleave="emit('pointerleave')"
    @keydown.esc.stop.prevent="emit('close')"
  >
    <div class="docx-style-preview__head">
      <span class="docx-style-preview__heading">{{ m.docx.stylePreview.label }}</span>
      <span v-if="activeStyle" class="docx-style-preview__name" data-testid="docx-style-preview-name">
        {{ activeStyle.name }}
      </span>
      <span v-else class="docx-style-preview__name" data-testid="docx-style-preview-name">
        {{ m.docx.stylePreview.normal }}
      </span>
    </div>

    <div
      class="docx-style-preview__sample"
      data-testid="docx-style-preview-sample"
      :style="sampleStyle"
      aria-hidden="true"
    >
      {{ m.docx.stylePreview.sample }}
    </div>

    <div class="docx-style-preview__list">
      <button
        type="button"
        role="option"
        class="docx-style-preview__item"
        :class="{ 'docx-style-preview__item--active': !activeStyle }"
        :aria-selected="!activeStyle"
        data-testid="docx-style-preview-item-normal"
        @pointerenter="emit('hover', undefined)"
        @focus="emit('hover', undefined)"
        @click="emit('select', undefined)"
      >
        {{ m.docx.stylePreview.normal }}
      </button>
      <button
        v-for="style of styles"
        :key="style.id"
        type="button"
        role="option"
        class="docx-style-preview__item"
        :class="{ 'docx-style-preview__item--active': activeStyle?.id === style.id }"
        :aria-selected="activeStyle?.id === style.id"
        :data-docx-style-preview-id="style.id"
        :data-testid="`docx-style-preview-item-${style.id}`"
        :style="itemStyle(style)"
        @pointerenter="emit('hover', style.id)"
        @focus="emit('hover', style.id)"
        @click="emit('select', style.id)"
      >
        {{ style.name }}
      </button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type { DocxDocumentTheme, ParagraphStyleDefinition } from "@arcships/docx-core"
import { useDocxMessages } from "../i18n"
import { docxStylePreviewCss } from "../composables/docx-style-preview"

const props = withDefaults(
  defineProps<{
    open?: boolean
    styles: ParagraphStyleDefinition[]
    /** Hovered style; undefined previews the document default ("Normal"). */
    previewStyleId?: string
    documentTheme?: DocxDocumentTheme
    x?: number
    y?: number
  }>(),
  { open: true, documentTheme: "light", x: 0, y: 0 }
)

const emit = defineEmits<{
  (event: "hover", styleId: string | undefined): void
  (event: "select", styleId: string | undefined): void
  (event: "close"): void
  (event: "pointerenter"): void
  (event: "pointerleave"): void
}>()

const m = useDocxMessages()

const activeStyle = computed(() =>
  props.styles.find((style) => style.id === props.previewStyleId)
)

const preview = computed(() => docxStylePreviewCss(activeStyle.value, props.documentTheme))
const sampleStyle = computed(() => preview.value.text)
const popoverStyle = computed(() => ({
  position: "fixed" as const,
  left: `${Math.max(0, Math.round(props.x))}px`,
  top: `${Math.max(0, Math.round(props.y))}px`,
  zIndex: 65,
}))

function itemStyle(style: ParagraphStyleDefinition): Record<string, string | number | undefined> {
  return docxStylePreviewCss(style, props.documentTheme).text
}
</script>

<style scoped>
.docx-style-preview {
  background: #ffffff;
  border: 1px solid #d1d5db;
  border-radius: 8px;
  box-shadow: 0 10px 30px rgba(15, 23, 42, 0.18);
  color: #111827;
  display: flex;
  flex-direction: column;
  gap: 8px;
  max-height: 22rem;
  overflow: auto;
  padding: 10px;
  width: 16rem;
}
.docx-style-preview__head {
  align-items: baseline;
  display: flex;
  gap: 6px;
}
.docx-style-preview__heading {
  color: #6b7280;
  font-size: 11px;
  text-transform: uppercase;
}
.docx-style-preview__name {
  font-size: 12px;
  font-weight: 600;
}
.docx-style-preview__sample {
  border: 1px dashed #e5e7eb;
  border-radius: 4px;
  font-size: 13px;
  line-height: 1.4;
  padding: 8px;
  white-space: pre-wrap;
  word-break: break-word;
}
.docx-style-preview__list {
  display: flex;
  flex-direction: column;
  gap: 2px;
}
.docx-style-preview__item {
  background: transparent;
  border: 1px solid transparent;
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  overflow: hidden;
  padding: 3px 6px;
  text-align: left;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.docx-style-preview__item:hover,
.docx-style-preview__item--active {
  background: #eff6ff;
  border-color: #bfdbfe;
}
.docx-style-preview--dark {
  background: #18181b;
  border-color: #3f3f46;
  color: #f4f4f5;
}
.docx-style-preview--dark .docx-style-preview__heading {
  color: #a1a1aa;
}
.docx-style-preview--dark .docx-style-preview__sample {
  border-color: #3f3f46;
}
.docx-style-preview--dark .docx-style-preview__item:hover,
.docx-style-preview--dark .docx-style-preview__item--active {
  background: #1e3a8a;
  border-color: #1d4ed8;
}
</style>
