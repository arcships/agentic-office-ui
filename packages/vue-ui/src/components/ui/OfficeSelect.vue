<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import {
  SelectContent,
  SelectGroup,
  SelectIcon,
  SelectItem,
  SelectItemIndicator,
  SelectItemText,
  SelectLabel,
  SelectPortal,
  SelectRoot,
  SelectSeparator,
  SelectTrigger,
  SelectValue,
  SelectViewport,
} from "reka-ui"
import { useOfficePortalTarget } from "../../composables/office-portal"
import OfficeIcon from "../OfficeIcon.vue"
import {
  officeClass,
  type OfficeAlign,
  type OfficeControlSize,
  type OfficeSelectEntry,
  type OfficeSelectOption,
  type OfficeSide,
} from "./types"

export interface OfficeSelectProps {
  /** Selected value (`v-model`). */
  modelValue?: string
  /** Flat options, groups and separators. */
  options: OfficeSelectEntry[]
  placeholder?: string
  size?: OfficeControlSize
  disabled?: boolean
  invalid?: boolean
  /** Accessible name for the trigger when no external `<label>` is used. */
  ariaLabel?: string
  /** Native form field name (Reka renders a hidden select). */
  name?: string
  id?: string
  required?: boolean
  side?: OfficeSide
  align?: OfficeAlign
  sideOffset?: number
  /** Extra classes merged on the trigger. */
  class?: string
  /** Extra classes merged on the portalled content. */
  contentClass?: string
}

const props = withDefaults(defineProps<OfficeSelectProps>(), {
  modelValue: undefined,
  placeholder: undefined,
  size: "default",
  disabled: false,
  invalid: false,
  ariaLabel: undefined,
  name: undefined,
  id: undefined,
  required: false,
  side: "bottom",
  align: "start",
  sideOffset: 4,
  class: "",
  contentClass: "",
})

const emit = defineEmits<{
  "update:modelValue": [value: string]
  "update:open": [value: boolean]
}>()

const portalTarget = useOfficePortalTarget()

interface NormalizedGroup {
  kind: "group"
  label?: string
  options: OfficeSelectOption[]
}

interface NormalizedSeparator {
  kind: "separator"
}

const normalizedOptions = computed<Array<NormalizedGroup | NormalizedSeparator>>(() => {
  const result: Array<NormalizedGroup | NormalizedSeparator> = []
  for (const entry of props.options) {
    if ((entry as { type?: string }).type === "separator") {
      result.push({ kind: "separator" })
      continue
    }
    if ((entry as { type?: string }).type === "group") {
      const group = entry as { label?: string; options: OfficeSelectOption[] }
      result.push({ kind: "group", label: group.label, options: group.options })
      continue
    }
    result.push({ kind: "group", options: [entry as OfficeSelectOption] })
  }
  return result
})

const triggerClasses = computed(() =>
  officeClass(
    "office-ui",
    "office-select__trigger",
    props.size !== "default" && `office-select__trigger--${props.size}`,
    props.class,
  ),
)
</script>

<template>
  <SelectRoot
    :model-value="props.modelValue"
    :disabled="props.disabled"
    :name="props.name"
    :required="props.required"
    @update:model-value="emit('update:modelValue', String($event))"
    @update:open="emit('update:open', $event)"
  >
    <SelectTrigger
      :id="props.id"
      :class="triggerClasses"
      :aria-label="props.ariaLabel"
      :aria-invalid="props.invalid ? 'true' : undefined"
      data-slot="select-trigger"
    >
      <SelectValue class="office-select__value" :placeholder="props.placeholder" />
      <SelectIcon class="office-select__icon">
        <OfficeIcon name="chevron-down" />
      </SelectIcon>
    </SelectTrigger>
    <SelectPortal :to="portalTarget ?? undefined">
      <SelectContent
        :class="officeClass('office-ui', 'office-select__content', props.contentClass)"
        :side="props.side"
        :align="props.align"
        :side-offset="props.sideOffset"
        position="popper"
        data-slot="select-content"
      >
        <SelectViewport class="office-select__viewport">
          <template v-for="(entry, index) in normalizedOptions" :key="index">
            <SelectSeparator
              v-if="entry.kind === 'separator'"
              class="office-select__separator"
              data-slot="select-separator"
            />
            <SelectGroup v-else class="office-select__group" data-slot="select-group">
              <SelectLabel
                v-if="entry.label"
                class="office-select__label"
                data-slot="select-label"
              >
                {{ entry.label }}
              </SelectLabel>
              <SelectItem
                v-for="option in entry.options"
                :key="option.value"
                :value="option.value"
                :disabled="option.disabled"
                :text-value="option.label"
                class="office-select__item"
                data-slot="select-item"
              >
                <span class="office-select__item-indicator">
                  <SelectItemIndicator>
                    <OfficeIcon name="check" />
                  </SelectItemIndicator>
                </span>
                <OfficeIcon v-if="option.icon" :name="option.icon" />
                <SelectItemText>{{ option.label }}</SelectItemText>
              </SelectItem>
            </SelectGroup>
          </template>
        </SelectViewport>
      </SelectContent>
    </SelectPortal>
  </SelectRoot>
</template>
