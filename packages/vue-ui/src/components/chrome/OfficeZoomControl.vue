<script setup lang="ts">
import { computed, ref } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { useOfficeUiMessages } from "../../i18n"
import OfficeIconButton from "../ui/OfficeIconButton.vue"
import OfficeSelect from "../ui/OfficeSelect.vue"
import {
  officeClass,
  type OfficeSelectEntry,
} from "../ui/types"
import type { OfficeZoomChange, OfficeZoomFitMode } from "./types"

export interface OfficeZoomControlProps {
  /**
   * Current zoom percentage (100 = actual size). Leave undefined for
   * uncontrolled use; the component keeps its own value and still emits
   * `update:modelValue`.
   */
  modelValue?: number
  /** Initial zoom for uncontrolled use. */
  defaultZoom?: number
  /**
   * Active fit mode. `width` and `page` take over from the numeric zoom.
   * Leave undefined for uncontrolled use.
   */
  fitMode?: OfficeZoomFitMode
  /** Lower clamp for the numeric zoom. */
  min?: number
  /** Upper clamp for the numeric zoom. */
  max?: number
  /** Percent presets rendered in the select. */
  options?: number[]
  disabled?: boolean
  /** Render the fit-width / fit-page presets in the select. */
  showFit?: boolean
  /** Accessible name of the select; defaults to `common.zoom.level`. */
  label?: string
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<OfficeZoomControlProps>(), {
  modelValue: undefined,
  defaultZoom: 100,
  fitMode: undefined,
  min: 25,
  max: 400,
  options: () => [50, 75, 100, 125, 150, 200],
  disabled: false,
  showFit: true,
  label: undefined,
  class: "",
})

const emit = defineEmits<{
  "update:modelValue": [zoom: number]
  "update:fitMode": [fitMode: OfficeZoomFitMode]
  change: [value: OfficeZoomChange]
}>()

const m = useOfficeUiMessages()

const internalZoom = ref(clamp(props.defaultZoom))
const internalFit = ref<OfficeZoomFitMode>("none")

const zoom = computed(() => clamp(props.modelValue ?? internalZoom.value))
const fit = computed<OfficeZoomFitMode>(() => props.fitMode ?? internalFit.value)

const classes = computed(() =>
  officeClass(
    "office-ui",
    "office-zoom-control",
    fit.value !== "none" ? "office-zoom-control--fit" : undefined,
    props.class,
  ),
)

const percentLabel = (value: number) =>
  formatOfficeMessage(m.value.common.zoom.percent, { value })

const percentOptions = computed(() => {
  const values = new Set<number>(props.options)
  values.add(zoom.value)
  return [...values]
    .filter((value) => Number.isFinite(value))
    .map((value) => clamp(value))
    .filter((value, index, all) => all.indexOf(value) === index)
    .sort((left, right) => left - right)
})

const selectOptions = computed<OfficeSelectEntry[]>(() => {
  const entries: OfficeSelectEntry[] = []
  if (props.showFit) {
    entries.push(
      { value: "fit-width", label: m.value.common.zoom.fitWidth },
      { value: "fit-page", label: m.value.common.zoom.fitPage },
      { value: "actual", label: m.value.common.zoom.actualSize },
    )
  }
  for (const value of percentOptions.value) {
    entries.push({ value: String(value), label: percentLabel(value) })
  }
  return entries
})

const selectValue = computed(() => {
  if (fit.value === "width") return "fit-width"
  if (fit.value === "page") return "fit-page"
  return String(zoom.value)
})

function clamp(value: number): number {
  const finite = Number.isFinite(value) ? value : 100
  return Math.min(Math.max(finite, props.min), props.max)
}

function commit(nextZoom: number, nextFit: OfficeZoomFitMode): void {
  const rounded = clamp(Math.round(nextZoom))
  const previousZoom = zoom.value
  const previousFit = fit.value
  if (props.modelValue === undefined) internalZoom.value = rounded
  if (props.fitMode === undefined) internalFit.value = nextFit
  if (rounded !== previousZoom) emit("update:modelValue", rounded)
  if (nextFit !== previousFit) emit("update:fitMode", nextFit)
  if (rounded !== previousZoom || nextFit !== previousFit) {
    emit("change", { zoom: rounded, fitMode: nextFit })
  }
}

function setZoom(next: number): void {
  commit(next, "none")
}

function setFit(next: OfficeZoomFitMode): void {
  commit(zoom.value, next)
}

function onSelect(value: string): void {
  if (value === "fit-width") return setFit("width")
  if (value === "fit-page") return setFit("page")
  if (value === "actual") return setZoom(100)
  const next = Number(value)
  if (Number.isFinite(next)) setZoom(next)
}

function stepZoom(direction: -1 | 1): void {
  if (props.disabled) return
  const options = percentOptions.value
  if (!options.length) return
  const current = zoom.value
  const next = direction < 0
    ? [...options].reverse().find((option) => option < current)
    : options.find((option) => option > current)
  setZoom(typeof next === "number" ? next : current)
}

const canZoomOut = computed(
  () => !props.disabled && zoom.value > (percentOptions.value[0] ?? props.min),
)
const canZoomIn = computed(() => {
  if (props.disabled) return false
  const options = percentOptions.value
  return zoom.value < (options[options.length - 1] ?? props.max)
})
</script>

<template>
  <div
    :class="classes"
    role="group"
    :aria-label="props.label ?? m.common.zoom.level"
    data-slot="zoom-control"
    data-testid="office-zoom-control"
  >
    <OfficeIconButton
      :label="m.common.zoom.out"
      icon="minus"
      size="icon-sm"
      :disabled="!canZoomOut"
      data-testid="office-zoom-out"
      @click="stepZoom(-1)"
    />
    <span
      class="office-zoom-control__select"
      data-slot="zoom-select"
      data-testid="office-zoom-select"
    >
      <OfficeSelect
        :model-value="selectValue"
        :options="selectOptions"
        :aria-label="props.label ?? m.common.zoom.level"
        size="sm"
        :disabled="props.disabled"
        @update:model-value="onSelect"
      />
    </span>
    <OfficeIconButton
      :label="m.common.zoom.in"
      icon="plus"
      size="icon-sm"
      :disabled="!canZoomIn"
      data-testid="office-zoom-in"
      @click="stepZoom(1)"
    />
  </div>
</template>
