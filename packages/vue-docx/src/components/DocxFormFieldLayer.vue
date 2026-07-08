<template>
  <div class="docx-form-field-layer">
    <div
      v-for="field of formFields"
      :key="field.id"
      class="docx-form-field"
      :style="fieldStyle(field)"
    >
      <input
        v-if="field.fieldType === 'text'"
        type="text"
        :value="field.value"
        :placeholder="field.placeholder"
        class="docx-form-text-input"
        @input="onFieldInput(field.id, $event)"
        @focus="onFieldFocus(field)"
      />
      <label
        v-else-if="field.fieldType === 'checkbox'"
        class="docx-form-checkbox-label"
      >
        <input
          type="checkbox"
          :checked="field.checked"
          @change="onCheckboxChange(field.id, $event)"
          @focus="onFieldFocus(field)"
        />
        <span class="docx-form-checkbox-text">{{ field.label }}</span>
      </label>
      <select
        v-else-if="field.fieldType === 'dropdown'"
        :value="field.value"
        class="docx-form-select"
        @change="onFieldInput(field.id, $event)"
        @focus="onFieldFocus(field)"
      >
        <option v-for="opt of field.options" :key="opt.value" :value="opt.value">
          {{ opt.label }}
        </option>
      </select>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type { DocxEditorController, LayoutPage } from "@extend-ai/docx-core"

const props = defineProps<{
  page: LayoutPage
  controller?: DocxEditorController
}>()

interface FormFieldEntry {
  id: string
  fieldType: string
  value?: string
  placeholder?: string
  checked?: boolean
  label?: string
  options?: { value: string; label: string }[]
  xPx: number
  yPx: number
  widthPx: number
  heightPx: number
}

const formFields = computed<FormFieldEntry[]>(() => [])

function fieldStyle(field: FormFieldEntry): Record<string, string> {
  return {
    position: "absolute",
    left: `${field.xPx}px`,
    top: `${field.yPx}px`,
    width: `${field.widthPx}px`,
    height: `${field.heightPx}px`,
  }
}

function onFieldInput(id: string, event: Event) {
  const target = event.target as HTMLInputElement | HTMLSelectElement
  // Dispatch form field value change via controller
}

function onFieldFocus(field: FormFieldEntry) {
  // Select form field
}

function onCheckboxChange(id: string, event: Event) {
  const target = event.target as HTMLInputElement
  // Toggle checkbox
}
</script>

<style scoped>
.docx-form-field-layer {
  position: absolute;
  inset: 0;
  pointer-events: none;
}
.docx-form-field {
  pointer-events: auto;
}
.docx-form-text-input {
  width: 100%;
  padding: 2px 4px;
  border: 1px solid #d1d5db;
  border-radius: 2px;
  font-size: inherit;
  font-family: inherit;
}
.docx-form-checkbox-label {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  cursor: pointer;
}
.docx-form-select {
  width: 100%;
  padding: 2px 4px;
  border: 1px solid #d1d5db;
  border-radius: 2px;
}
</style>
