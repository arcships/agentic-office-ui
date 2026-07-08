<template>
  <div v-if="formFields.length > 0" class="docx-form-field-layer">
    <div
      v-for="field of formFields"
      :key="field.id"
      class="docx-form-field"
      :class="{ 'docx-form-field--selected': field.selected }"
      :data-docx-form-field="true"
      :data-docx-form-field-location="field.locationKey"
      :style="fieldStyle(field)"
    >
      <!-- Text input -->
      <input
        v-if="field.fieldType === 'text'"
        type="text"
        :value="field.value"
        :placeholder="field.placeholder ?? 'Enter text...'"
        class="docx-form-text-input"
        @input="onFieldInput(field, ($event.target as HTMLInputElement).value)"
        @focus="onFieldFocus(field)"
        @dblclick="onFieldDoubleClick(field)"
      />

      <!-- Checkbox -->
      <label
        v-else-if="field.fieldType === 'checkbox'"
        class="docx-form-checkbox-label"
      >
        <input
          type="checkbox"
          :checked="field.checked"
          @change="onCheckboxToggle(field)"
          @focus="onFieldFocus(field)"
        />
        <span v-if="field.title" class="docx-form-checkbox-text">
          {{ field.title }}
        </span>
      </label>

      <!-- Dropdown -->
      <select
        v-else-if="field.fieldType === 'dropdown'"
        :value="field.value"
        class="docx-form-select"
        @change="onFieldInput(field, ($event.target as HTMLSelectElement).value)"
        @focus="onFieldFocus(field)"
      >
        <option
          v-for="opt of field.options"
          :key="opt.value"
          :value="opt.value"
        >
          {{ opt.label }}
        </option>
      </select>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type {
  DocxEditorController,
  DocxFormFieldLocation,
  FormFieldRunNode,
} from "@extend-ai/docx-core"
import {
  getParagraphAtLocation,
} from "@extend-ai/docx-core"

// ── Types ──────────────────────────────────────────────────────────
interface FormFieldEntry {
  id: string
  fieldType: string
  value: string
  placeholder?: string
  checked: boolean
  title?: string
  options: any[]
  locationKey: string
  location: DocxFormFieldLocation
  selected: boolean
}

// ── Props ──────────────────────────────────────────────────────────
const props = defineProps<{
  pageIndex: number
  controller: DocxEditorController
}>()

// ── Computed ───────────────────────────────────────────────────────
const formFields = computed<FormFieldEntry[]>(() => {
  const model = props.controller.model
  if (!model) return []

  const fields: FormFieldEntry[] = []
  const selectedLoc = props.controller.selectedFormField?.location

  model.nodes.forEach((node, nodeIndex) => {
    if (node.type !== "paragraph") return
    node.children.forEach((child, childIndex) => {
      if (child.type !== "form-field") return
      const ff = child as FormFieldRunNode

      const location: DocxFormFieldLocation = {
        kind: "paragraph",
        nodeIndex,
        childIndex,
      }

      const locationKey = `p:${nodeIndex}:${childIndex}`
      const isSelected =
        selectedLoc?.kind === "paragraph" &&
        selectedLoc.nodeIndex === nodeIndex &&
        selectedLoc.childIndex === childIndex

      fields.push({
        id: `ff-${nodeIndex}-${childIndex}`,
        fieldType: ff.fieldType ?? "text",
        value: ff.value ?? "",
        placeholder: ff.placeholder,
        checked: ff.checked ?? false,
        title: ff.title,
        options: (ff.options ?? []) as any[],
        locationKey,
        location,
        selected: isSelected,
      })
    })
  })

  return fields
})

// ── Style helpers ──────────────────────────────────────────────────
function fieldStyle(field: FormFieldEntry): Record<string, string> {
  return {
    display: field.fieldType === "checkbox" ? "inline-flex" : "block",
    marginBottom: "4px",
    outline: field.selected ? "2px solid #3b82f6" : "none",
    outlineOffset: "2px",
    borderRadius: "2px",
  }
}

// ── Event handlers ─────────────────────────────────────────────────
function onFieldInput(field: FormFieldEntry, value: string): void {
  props.controller.setFormFieldValue(field.location, value)
}

function onCheckboxToggle(field: FormFieldEntry): void {
  props.controller.toggleFormCheckbox(field.location)
}

function onFieldFocus(field: FormFieldEntry): void {
  props.controller.selectFormField(field.location)
}

function onFieldDoubleClick(field: FormFieldEntry): void {
  props.controller.selectFormField(field.location)
}
</script>

<style scoped>
.docx-form-field-layer {
  padding: 4px 0;
}
.docx-form-field {
  pointer-events: auto;
}
.docx-form-text-input {
  width: 100%;
  max-width: 300px;
  padding: 2px 4px;
  border: 1px solid #d1d5db;
  border-radius: 2px;
  font-size: inherit;
  font-family: inherit;
  outline: none;
}
.docx-form-text-input:focus {
  border-color: #3b82f6;
  box-shadow: 0 0 0 1px #3b82f6;
}
.docx-form-checkbox-label {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  cursor: pointer;
  font-size: inherit;
}
.docx-form-checkbox-text {
  user-select: none;
}
.docx-form-select {
  padding: 2px 4px;
  border: 1px solid #d1d5db;
  border-radius: 2px;
  font-size: inherit;
  font-family: inherit;
}
</style>
