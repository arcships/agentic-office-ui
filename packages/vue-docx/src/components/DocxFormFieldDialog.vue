<template>
  <OfficeDialog
    :open="open"
    :title="m.docx.formFieldDialog.title"
    :close-label="m.docx.formFieldDialog.close"
    close-test-id="docx-form-field-dialog-close"
    content-test-id="docx-form-field-dialog"
    :content-data="contentData"
    @update:open="onDialogOpenChange"
  >
    <div
      class="docx-form-field-dialog"
      :class="{ 'docx-form-field-dialog--dark': isDark }"
    >
        <label class="docx-form-field-dialog__row">
          <span class="docx-form-field-dialog__label">{{ m.docx.formFieldDialog.fieldType }}</span>
          <select
            v-model="draft.fieldType"
            class="docx-form-field-dialog__control"
            data-testid="docx-form-field-dialog-type"
          >
            <option value="text">{{ m.docx.formFieldDialog.types.text }}</option>
            <option value="checkbox">{{ m.docx.formFieldDialog.types.checkbox }}</option>
            <option value="dropdown">{{ m.docx.formFieldDialog.types.dropdown }}</option>
            <option value="date">{{ m.docx.formFieldDialog.types.date }}</option>
          </select>
        </label>

        <label class="docx-form-field-dialog__row">
          <span class="docx-form-field-dialog__label">{{ m.docx.formFieldDialog.name }}</span>
          <input
            ref="firstFieldEl"
            v-model="draft.title"
            type="text"
            class="docx-form-field-dialog__control"
            :placeholder="m.docx.formFieldDialog.namePlaceholder"
            data-testid="docx-form-field-dialog-title"
          />
        </label>

        <label class="docx-form-field-dialog__row">
          <span class="docx-form-field-dialog__label">{{ m.docx.formFieldDialog.value }}</span>
          <input
            v-if="draft.fieldType === 'checkbox'"
            v-model="draft.checked"
            type="checkbox"
            class="docx-form-field-dialog__checkbox"
            data-testid="docx-form-field-dialog-checked"
          />
          <input
            v-else-if="draft.fieldType === 'date'"
            v-model="draft.fieldValue"
            type="date"
            class="docx-form-field-dialog__control"
            data-testid="docx-form-field-dialog-value"
          />
          <select
            v-else-if="draft.fieldType === 'dropdown'"
            v-model="draft.fieldValue"
            class="docx-form-field-dialog__control"
            data-testid="docx-form-field-dialog-value"
          >
            <option value="">—</option>
            <option
              v-for="(option, index) of draft.options"
              :key="`value-option-${index}`"
              :value="option.value || option.displayText"
            >
              {{ option.displayText }}
            </option>
          </select>
          <input
            v-else
            v-model="draft.fieldValue"
            type="text"
            class="docx-form-field-dialog__control"
            data-testid="docx-form-field-dialog-value"
          />
        </label>

        <template v-if="draft.fieldType === 'dropdown'">
          <div class="docx-form-field-dialog__options" data-testid="docx-form-field-dialog-options">
            <span class="docx-form-field-dialog__label">{{ m.docx.formFieldDialog.options }}</span>
            <p v-if="draft.options.length === 0" class="docx-form-field-dialog__hint">
              {{ m.docx.formFieldDialog.noOptions }}
            </p>
            <div
              v-for="(option, index) of draft.options"
              :key="`option-${index}`"
              class="docx-form-field-dialog__option"
              :data-docx-form-field-dialog-option="index"
            >
              <input
                v-model="option.displayText"
                type="text"
                class="docx-form-field-dialog__control"
                :aria-label="m.docx.formFieldDialog.option"
                :placeholder="m.docx.formFieldDialog.option"
                :data-testid="`docx-form-field-dialog-option-${index}`"
              />
              <input
                v-model="option.value"
                type="text"
                class="docx-form-field-dialog__control"
                :aria-label="m.docx.formFieldDialog.optionValue"
                :placeholder="m.docx.formFieldDialog.optionValue"
                :data-testid="`docx-form-field-dialog-option-value-${index}`"
              />
              <OfficeButton
                type="button"
                variant="outline"
                size="xs"
                class="docx-form-field-dialog__option-remove"
                :aria-label="m.docx.formFieldDialog.removeOption"
                :title="m.docx.formFieldDialog.removeOption"
                :data-testid="`docx-form-field-dialog-remove-option-${index}`"
                @click="removeOption(index)"
              >−</OfficeButton>
            </div>
            <OfficeButton
              type="button"
              variant="outline"
              size="xs"
              class="docx-form-field-dialog__add-option"
              data-testid="docx-form-field-dialog-add-option"
              @click="addOption"
            >
              {{ m.docx.formFieldDialog.addOption }}
            </OfficeButton>
          </div>
        </template>

        <label v-if="draft.fieldType === 'text' || draft.fieldType === 'date'" class="docx-form-field-dialog__row">
          <span class="docx-form-field-dialog__label">{{ m.docx.formFieldDialog.maxLength }}</span>
          <input
            v-model="draft.maxLength"
            type="number"
            min="0"
            class="docx-form-field-dialog__control"
            data-testid="docx-form-field-dialog-max-length"
          />
        </label>

        <label class="docx-form-field-dialog__row">
          <span class="docx-form-field-dialog__label">{{ m.docx.formFieldDialog.placeholder }}</span>
          <input
            v-model="draft.placeholder"
            type="text"
            class="docx-form-field-dialog__control"
            :placeholder="m.docx.formFieldDialog.placeholder"
            data-testid="docx-form-field-dialog-placeholder"
          />
        </label>

        <label class="docx-form-field-dialog__row docx-form-field-dialog__row--inline">
          <input
            v-model="draft.required"
            type="checkbox"
            class="docx-form-field-dialog__checkbox"
            data-testid="docx-form-field-dialog-required"
          />
          <span class="docx-form-field-dialog__label">{{ m.docx.formFieldDialog.required }}</span>
        </label>

      <p v-if="error" class="docx-form-field-dialog__error" role="alert">
        {{ error }}
      </p>
    </div>

    <template #footer>
      <OfficeButton
        type="button"
        size="sm"
        class="docx-form-field-dialog__save"
        data-testid="docx-form-field-dialog-save"
        @click="submit"
      >
        {{ m.docx.formFieldDialog.save }}
      </OfficeButton>
      <OfficeButton
        type="button"
        variant="outline"
        size="sm"
        class="docx-form-field-dialog__cancel"
        data-testid="docx-form-field-dialog-cancel"
        @click="close()"
      >
        {{ m.docx.formFieldDialog.cancel }}
      </OfficeButton>
    </template>
  </OfficeDialog>
</template>

<script setup lang="ts">
import { computed, nextTick, onMounted, ref, watch } from "vue"
import { OfficeButton, OfficeDialog } from "@arcships/vue-ui"
import type {
  DocxEditorController,
  DocxFormFieldLocation,
  DocxFormFieldUpdate,
  FormFieldOption,
  FormFieldRunNode,
  FormFieldType,
} from "@arcships/docx-core"
import { useDocxMessages } from "../i18n"
import {
  isDateFormField,
  isoDateValueForField,
  validateFormFieldValue,
} from "../composables/form-field-value"
import { resolveFormFieldAt } from "../composables/docx-form-fields"

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    open?: boolean
    /** Form field as it currently exists in the controller model. */
    field: FormFieldRunNode
    location: DocxFormFieldLocation
    controller: DocxEditorController
  }>(),
  { open: true }
)

const emit = defineEmits<{
  (event: "close"): void
  /** Fired after the controller accepted the patch. */
  (event: "saved", field: FormFieldRunNode): void
}>()

const m = useDocxMessages()
const isDark = computed(() => props.controller.documentTheme === "dark")
const contentData = computed(() => ({
  "data-docx-form-field-dialog": true,
  "data-office-theme": isDark.value ? "dark" : "light",
  "data-office-density": "comfortable",
  "data-docx-thumbnail-exclude": "true",
  // Reka's dialog sets role but not aria-modal; the editor contract keeps it.
  "aria-modal": "true",
}))

interface FormFieldDialogDraft {
  fieldType: FormFieldType
  title: string
  placeholder: string
  required: boolean
  checked: boolean
  fieldValue: string
  maxLength: string
  options: Array<{ displayText: string; value: string }>
}

const firstFieldEl = ref<HTMLInputElement | undefined>()
const error = ref("")

function draftFromField(field: FormFieldRunNode): FormFieldDialogDraft {
  return {
    fieldType: field.fieldType,
    title: field.title ?? "",
    placeholder: field.placeholder ?? "",
    required: field.required === true,
    checked: field.checked === true || field.value === "1",
    fieldValue:
      field.fieldType === "checkbox"
        ? ""
        : isDateFormField(field)
          ? isoDateValueForField(field)
          : field.value ?? "",
    maxLength:
      field.widget?.text?.maxLength === undefined
        ? ""
        : String(field.widget.text.maxLength),
    options: (field.options ?? []).map((option: FormFieldOption) => ({
      displayText: option.displayText ?? "",
      value: option.value ?? "",
    })),
  }
}

const draft = ref(draftFromField(props.field))

// The dialog is a view of the model: an external edit (undo, another host)
// refreshes the draft instead of letting the dialog save stale properties.
watch(
  () => props.field,
  (field) => {
    draft.value = draftFromField(field)
    error.value = ""
  }
)

watch(
  () => props.open,
  (open) => {
    if (!open) {
      error.value = ""
      return
    }
    draft.value = draftFromField(props.field)
    focusFirstField()
  }
)

onMounted(() => {
  if (props.open) focusFirstField()
})

/**
 * Reka's dialog focuses the first tabbable control (the field-type select),
 * while the form-field dialog has always landed on the name field. Wait for
 * the primitive's focus pass, then move focus to the first property field.
 */
function focusFirstField(): void {
  void nextTick(() => {
    void nextTick(() => firstFieldEl.value?.focus?.())
  })
}

/** Esc, backdrop and the close button all funnel into the public close event. */
function onDialogOpenChange(value: boolean): void {
  if (!value) close()
}

function addOption(): void {
  draft.value.options.push({ displayText: "", value: "" })
}

function removeOption(index: number): void {
  draft.value.options.splice(index, 1)
}

function patchFromDraft(): DocxFormFieldUpdate {
  const current = draft.value
  const dropdown = current.fieldType === "dropdown"
  const options = current.options
    .map((option) => ({
      displayText: option.displayText.trim(),
      value: option.value.trim() === "" ? undefined : option.value.trim(),
    }))
    .filter((option) => option.displayText !== "")

  const maxLength = Number.parseInt(current.maxLength, 10)
  const wantsMaxLength =
    (current.fieldType === "text" || current.fieldType === "date") &&
    Number.isFinite(maxLength) &&
    maxLength >= 0
  const patch: DocxFormFieldUpdate = {
    fieldType: current.fieldType,
    title: current.title,
    placeholder: current.placeholder,
    required: current.required,
    options: dropdown ? options : undefined,
    checked: current.fieldType === "checkbox" ? current.checked : undefined,
    value:
      current.fieldType === "checkbox"
        ? current.checked
          ? "1"
          : "0"
        : current.fieldValue,
    widget: wantsMaxLength ? { text: { maxLength } } : undefined,
  }
  return patch
}

function submit(): void {
  const patch = patchFromDraft()

  // Validate against the field the patch produces so a value and the option
  // list it belongs to are checked together.
  const candidate: FormFieldRunNode = {
    ...props.field,
    fieldType: patch.fieldType ?? props.field.fieldType,
    options: patch.options ?? props.field.options,
    widget: patch.widget
      ? { ...(props.field.widget ?? {}), text: { ...(props.field.widget?.text ?? {}), ...patch.widget.text } }
      : props.field.widget,
  }
  if (patch.value !== undefined && (patch.value !== "" || patch.required === true)) {
    const validation = validateFormFieldValue(candidate, patch.value, {
      messages: m.value.docx.formField.errors,
    })
    if (!validation.valid) {
      error.value = validation.error ?? m.value.docx.formField.errors.invalidValue
      return
    }
  }

  const saved = props.controller.updateFormField(props.location, patch)
  if (!saved) {
    error.value = m.value.docx.formField.errors.invalidValue
    return
  }
  error.value = ""
  emit("saved", resolveFormFieldAt(props.controller.model, props.location) ?? props.field)
  close()
}

function close(): void {
  error.value = ""
  emit("close")
}

defineExpose({
  submit,
  draft,
  error,
})
</script>

<style scoped>
/* `OfficeDialog` owns the modal shell (backdrop, header, body scroll, footer),
 * so this block only lays out the form fields inside the slot. */
.docx-form-field-dialog {
  display: flex;
  flex-direction: column;
  gap: 10px;
  min-inline-size: 0;
}
.docx-form-field-dialog__row {
  align-items: center;
  display: flex;
  gap: 8px;
}
.docx-form-field-dialog__row--inline {
  gap: 6px;
}
.docx-form-field-dialog__label {
  color: var(--office-muted-foreground, #4b5563);
  flex: 0 0 7.5rem;
  font-size: 12px;
}
.docx-form-field-dialog__control {
  border: 1px solid var(--office-border, #d1d5db);
  border-radius: 4px;
  flex: 1;
  font: inherit;
  font-size: 13px;
  min-width: 0;
  padding: 4px 6px;
}
.docx-form-field-dialog__control:focus {
  border-color: var(--office-primary, #3b82f6);
  outline: 1px solid var(--office-primary, #3b82f6);
}
.docx-form-field-dialog__options {
  display: flex;
  flex-direction: column;
  gap: 6px;
}
.docx-form-field-dialog__option {
  align-items: center;
  display: flex;
  gap: 6px;
}
.docx-form-field-dialog__option-remove,
.docx-form-field-dialog__add-option {
  cursor: pointer;
}
.docx-form-field-dialog__add-option {
  align-self: flex-start;
}
.docx-form-field-dialog__hint {
  color: var(--office-muted-foreground, #6b7280);
  font-size: 12px;
  margin: 0;
}
.docx-form-field-dialog__error {
  color: var(--office-destructive-foreground, #b91c1c);
  font-size: 12px;
  margin: 0;
}
.docx-form-field-dialog__save,
.docx-form-field-dialog__cancel {
  cursor: pointer;
}
.docx-form-field-dialog__save {
  background: var(--office-primary, #2563eb);
  border-color: var(--office-primary, #2563eb);
  color: var(--office-primary-foreground, #ffffff);
}
.docx-form-field-dialog--dark .docx-form-field-dialog__label {
  color: var(--office-muted-foreground, #a1a1aa);
}
.docx-form-field-dialog--dark .docx-form-field-dialog__control {
  background: var(--office-background, #09090b);
  border-color: var(--office-border, #3f3f46);
  color: var(--office-popover-foreground, #f4f4f5);
}
.docx-form-field-dialog--dark .docx-form-field-dialog__option-remove,
.docx-form-field-dialog--dark .docx-form-field-dialog__add-option,
.docx-form-field-dialog--dark .docx-form-field-dialog__cancel {
  background: var(--office-muted, #27272a);
  border-color: var(--office-border, #3f3f46);
  color: var(--office-popover-foreground, #e4e4e7);
}
.docx-form-field-dialog--dark .docx-form-field-dialog__hint {
  color: var(--office-muted-foreground, #a1a1aa);
}
</style>
