<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
defineOptions({ inheritAttrs: false })
import { computed, ref } from "vue"
import { officeClass, type OfficeControlSize } from "./types"

export interface OfficeInputProps {
  modelValue?: string | number
  /** Control height preset. */
  size?: OfficeControlSize
  /** Sets `aria-invalid` and the destructive border. */
  invalid?: boolean
  disabled?: boolean
  readonly?: boolean
  type?: string
  placeholder?: string
  /** Extra classes merged on the `<input>`. */
  class?: string
}

const props = withDefaults(defineProps<OfficeInputProps>(), {
  modelValue: undefined,
  size: "default",
  invalid: false,
  disabled: false,
  readonly: false,
  type: "text",
  placeholder: undefined,
  class: "",
})

const emit = defineEmits<{
  "update:modelValue": [value: string]
}>()

const inputRef = ref<HTMLInputElement | null>(null)

const controlClasses = computed(() =>
  officeClass(
    "office-ui",
    "office-input-control",
    props.size !== "default" && `office-input-control--${props.size}`,
    props.class,
  ),
)

const inputClasses = computed(() =>
  officeClass(
    "office-ui",
    "office-input",
    props.size !== "default" && `office-input--${props.size}`,
  ),
)

function onInput(event: Event) {
  emit("update:modelValue", (event.target as HTMLInputElement).value)
}

defineExpose({
  focus: () => inputRef.value?.focus(),
  input: inputRef,
})
</script>

<template>
  <span
    :class="controlClasses"
    :data-size="props.size"
    :data-disabled="props.disabled ? 'true' : undefined"
    data-slot="input-control"
  >
    <input
      ref="inputRef"
      v-bind="$attrs"
      :class="inputClasses"
      :type="props.type"
      :value="props.modelValue"
      :placeholder="props.placeholder"
      :disabled="props.disabled"
      :readonly="props.readonly"
      :aria-invalid="props.invalid ? 'true' : undefined"
      data-slot="input"
      @input="onInput"
    />
  </span>
</template>
