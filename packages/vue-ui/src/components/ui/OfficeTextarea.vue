<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import { officeClass, type OfficeControlSize } from "./types"

export interface OfficeTextareaProps {
  modelValue?: string
  /** Text size preset; the textarea grows vertically. */
  size?: OfficeControlSize
  invalid?: boolean
  disabled?: boolean
  readonly?: boolean
  placeholder?: string
  /** Initial visible rows (CSS still enforces a minimum height). */
  rows?: number
  /** Extra classes merged on the `<textarea>`. */
  class?: string
}

const props = withDefaults(defineProps<OfficeTextareaProps>(), {
  modelValue: undefined,
  size: "default",
  invalid: false,
  disabled: false,
  readonly: false,
  placeholder: undefined,
  rows: 3,
  class: "",
})

const emit = defineEmits<{
  "update:modelValue": [value: string]
}>()

const classes = computed(() =>
  officeClass(
    "office-ui",
    "office-textarea",
    props.size !== "default" && `office-input--${props.size}`,
    props.class,
  ),
)

function onInput(event: Event) {
  emit("update:modelValue", (event.target as HTMLTextAreaElement).value)
}
</script>

<template>
  <textarea
    v-bind="$attrs"
    :class="classes"
    :value="props.modelValue"
    :rows="props.rows"
    :placeholder="props.placeholder"
    :disabled="props.disabled"
    :readonly="props.readonly"
    :aria-invalid="props.invalid ? 'true' : undefined"
    data-slot="textarea"
    @input="onInput"
  />
</template>
