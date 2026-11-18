<template>
  <div v-if="formFields.length > 0" class="docx-form-field-layer">
    <div
      v-for="field of formFields"
      :key="field.id"
      class="docx-form-field"
      :class="{ 'docx-form-field--selected': field.selected }"
      :data-docx-form-field="true"
      :data-docx-form-field-location="field.locationKey"
      :data-docx-form-field-required="field.required ? 'true' : undefined"
      :style="fieldStyle(field)"
    >
      <!-- Text input -->
      <input
        v-if="field.control === 'text'"
        type="text"
        :value="field.value"
        :maxlength="field.maxLength"
        :placeholder="field.placeholder ?? m.docx.formField.placeholder"
        :aria-invalid="fieldErrors[field.locationKey] ? 'true' : undefined"
        :aria-required="field.required ? 'true' : undefined"
        class="docx-form-text-input"
        @input="onTextInput(field, $event)"
        @focus="onFieldFocus(field)"
        @blur="onFieldBlur(field, $event)"
        @dblclick="onFieldDoubleClick(field)"
      />

      <!-- Date input: native control stores ISO, display follows the document format -->
      <div v-else-if="field.control === 'date'" class="docx-form-date">
        <input
          type="date"
          :value="field.dateValue"
          :aria-invalid="fieldErrors[field.locationKey] ? 'true' : undefined"
          :aria-required="field.required ? 'true' : undefined"
          :title="field.displayValue || m.docx.formField.selectDate"
          class="docx-form-text-input docx-form-date-input"
          @input="onDateInput(field, $event)"
          @focus="onFieldFocus(field)"
          @blur="onFieldBlur(field, $event)"
        />
        <span v-if="field.displayValue" class="docx-form-date-display">
          {{ field.displayValue }}
        </span>
      </div>

      <!-- Checkbox -->
      <label
        v-else-if="field.control === 'checkbox'"
        class="docx-form-checkbox-label"
      >
        <input
          type="checkbox"
          :checked="field.checked"
          :aria-required="field.required ? 'true' : undefined"
          @change="onCheckboxToggle(field)"
          @focus="onFieldFocus(field)"
          @blur="onFieldBlur(field, $event)"
        />
        <span v-if="field.title" class="docx-form-checkbox-text">
          {{ field.title }}
        </span>
      </label>

      <!-- Dropdown -->
      <select
        v-else-if="field.control === 'dropdown'"
        :value="field.value"
        :aria-invalid="fieldErrors[field.locationKey] ? 'true' : undefined"
        :aria-required="field.required ? 'true' : undefined"
        class="docx-form-select"
        @change="onSelectChange(field, $event)"
        @focus="onFieldFocus(field)"
        @blur="onFieldBlur(field, $event)"
      >
        <option
          v-for="(opt, optionIndex) of field.options"
          :key="`${optionIndex}:${opt.value ?? opt.displayText}`"
          :value="opt.value ?? opt.displayText"
        >
          {{ opt.displayText }}
        </option>
      </select>

      <span
        v-if="fieldErrors[field.locationKey]"
        class="docx-form-error"
        role="alert"
      >
        {{ fieldErrors[field.locationKey] }}
      </span>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref } from "vue"
import type {
  DocxEditorController,
  DocxFormFieldLocation,
  DocumentPageNodeSegment,
  FormFieldOption,
  FormFieldRunNode,
} from "@arcships/docx-core"
import {
  formFieldCurrentValue,
  formatFormFieldDateValue,
  isDateFormField,
  isoDateValueForField,
  validateFormFieldValue,
} from "../composables/form-field-value"
import { collectPageFormFields, docxFormFieldLocationKey } from "../composables/docx-form-fields"
import { useDocxMessages } from "../i18n"

// ── Types ──────────────────────────────────────────────────────────
type FormFieldControl = "text" | "date" | "checkbox" | "dropdown"

interface FormFieldEntry {
  id: string
  control: FormFieldControl
  node: FormFieldRunNode
  value: string
  displayValue: string
  dateValue: string
  placeholder?: string
  required: boolean
  checked: boolean
  title?: string
  options: FormFieldOption[]
  maxLength?: number
  locationKey: string
  location: DocxFormFieldLocation
  selected: boolean
}

// ── Props ──────────────────────────────────────────────────────────
const props = defineProps<{
  pageIndex: number
  controller: DocxEditorController
  pageNodeSegments: DocumentPageNodeSegment[]
}>()
const m = useDocxMessages()

const fieldErrors = ref<Record<string, string>>({})

// ── Computed ───────────────────────────────────────────────────────
function controlFor(node: FormFieldRunNode): FormFieldControl {
  if (node.fieldType === "checkbox") return "checkbox"
  if (node.fieldType === "dropdown") return "dropdown"
  if (isDateFormField(node)) return "date"
  return "text"
}

const formFields = computed<FormFieldEntry[]>(() => {
  const model = props.controller.model
  if (!model) return []

  const selectedLocation = props.controller.selectedFormField?.location
  const selectedKey = selectedLocation
    ? docxFormFieldLocationKey(selectedLocation)
    : undefined

  return collectPageFormFields(model, props.pageNodeSegments).map((entry) => {
    const ff = entry.field
    return {
      id: `ff-${entry.locationKey.replace(/:/g, "-")}`,
      control: controlFor(ff),
      node: ff,
      value: ff.value ?? "",
      displayValue: formatFormFieldDateValue(ff),
      dateValue: isoDateValueForField(ff),
      placeholder: ff.placeholder,
      required: ff.required === true,
      checked: ff.checked ?? false,
      title: ff.title,
      options: ff.options ?? [],
      maxLength: ff.widget?.text?.maxLength,
      locationKey: entry.locationKey,
      location: entry.location,
      selected: selectedKey === entry.locationKey,
    }
  })
})

// ── Style helpers ──────────────────────────────────────────────────
function fieldStyle(field: FormFieldEntry): Record<string, string> {
  return {
    display: field.control === "checkbox" ? "inline-flex" : "block",
    marginBottom: "4px",
    outline: field.selected ? "2px solid #3b82f6" : "none",
    outlineOffset: "2px",
    borderRadius: "2px",
  }
}

// ── Validation error state ─────────────────────────────────────────
function setFieldError(field: FormFieldEntry, error: string): void {
  fieldErrors.value = { ...fieldErrors.value, [field.locationKey]: error }
}

function clearFieldError(field: FormFieldEntry): void {
  if (!(field.locationKey in fieldErrors.value)) return
  const next = { ...fieldErrors.value }
  delete next[field.locationKey]
  fieldErrors.value = next
}

// ── Event handlers ─────────────────────────────────────────────────
function commitFieldValue(
  field: FormFieldEntry,
  rawValue: string,
  control: HTMLInputElement | HTMLSelectElement,
): void {
  const result = validateFormFieldValue(field.node, rawValue, {
    messages: m.value.docx.formField.errors,
  })
  if (!result.valid) {
    setFieldError(field, result.error ?? m.value.docx.formField.errors.invalidValue)
    // Keep the model untouched and restore the last accepted value.
    control.value = field.control === "date" ? field.dateValue : field.value
    return
  }
  clearFieldError(field)
  props.controller.setFormFieldValue(field.location, result.value)
}

function onTextInput(field: FormFieldEntry, event: Event): void {
  const input = event.target as HTMLInputElement
  commitFieldValue(field, input.value, input)
}

function onDateInput(field: FormFieldEntry, event: Event): void {
  const input = event.target as HTMLInputElement
  commitFieldValue(field, input.value, input)
}

function onSelectChange(field: FormFieldEntry, event: Event): void {
  const select = event.target as HTMLSelectElement
  commitFieldValue(field, select.value, select)
}

function onCheckboxToggle(field: FormFieldEntry): void {
  props.controller.toggleFormCheckbox(field.location)
}

function onFieldFocus(field: FormFieldEntry): void {
  props.controller.selectFormField(field.location)
}

function onFieldBlur(field: FormFieldEntry, event?: Event): void {
  props.controller.endFormFieldEdit(field.location)
  const control = event?.target as HTMLInputElement | HTMLSelectElement | undefined
  let value: string
  if (field.control === "checkbox") {
    const checked = (control as HTMLInputElement | undefined)?.checked ?? field.checked
    value = checked ? "1" : "0"
  } else {
    value = control?.value ?? formFieldCurrentValue(field.node)
  }
  const result = validateFormFieldValue(field.node, value, {
    enforceRequired: true,
    messages: m.value.docx.formField.errors,
  })
  if (!result.valid) {
    setFieldError(field, result.error ?? m.value.docx.formField.errors.required)
    return
  }
  clearFieldError(field)
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
.docx-form-text-input[aria-invalid="true"] {
  border-color: #dc2626;
  box-shadow: 0 0 0 1px #dc2626;
}
.docx-form-date {
  align-items: center;
  display: flex;
  flex-wrap: wrap;
  gap: 4px 8px;
}
.docx-form-date-input {
  width: auto;
  min-width: 132px;
}
.docx-form-date-display {
  color: #52525b;
  font-size: 0.85em;
}
.docx-form-error {
  color: #b91c1c;
  display: block;
  flex-basis: 100%;
  font-size: 0.8em;
  margin-top: 2px;
}
.docx-form-field[data-docx-form-field-required="true"] .docx-form-text-input,
.docx-form-field[data-docx-form-field-required="true"] .docx-form-select {
  border-left: 2px solid #2563eb;
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
.docx-form-select[aria-invalid="true"] {
  border-color: #dc2626;
}
</style>
