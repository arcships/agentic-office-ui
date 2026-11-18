<template>
  <Teleport to="body">
    <div
      class="docx-pagesetup-backdrop"
      data-testid="docx-page-setup-backdrop"
      @pointerdown="emit('close')"
    />
    <div
      ref="dialogRef"
      class="docx-pagesetup"
      :class="{ 'docx-pagesetup--dark': isDark }"
      data-testid="docx-page-setup-dialog"
      role="dialog"
      aria-modal="true"
      :aria-label="m.docx.pageSetup.title"
      tabindex="-1"
      @keydown.escape.stop.prevent="emit('close')"
      @pointerdown.stop
    >
      <header class="docx-pagesetup__header">
        <h2 class="docx-pagesetup__title">{{ m.docx.pageSetup.title }}</h2>
        <button
          type="button"
          class="docx-pagesetup__close"
          data-testid="docx-page-setup-close"
          :title="m.docx.pageSetup.close"
          :aria-label="m.docx.pageSetup.close"
          @click="emit('close')"
        >×</button>
      </header>

      <label v-if="sections.length > 1" class="docx-pagesetup__field">
        <span>{{ m.docx.pageSetup.sectionLabel }}</span>
        <select
          class="docx-pagesetup__input"
          data-testid="docx-page-setup-section"
          :value="String(sectionIndex)"
          @change="onSectionChange"
        >
          <option v-for="section of sections" :key="section.index" :value="String(section.index)">
            {{ formatOfficeMessage(m.docx.pageSetup.sectionOption, { index: section.sectionNumber }) }}
          </option>
        </select>
      </label>

      <section class="docx-pagesetup__group" data-testid="docx-page-setup-paper">
        <h3 class="docx-pagesetup__group-title">{{ m.docx.pageSetup.paperSize }}</h3>
        <label class="docx-pagesetup__field">
          <span>{{ m.docx.pageSetup.paperSize }}</span>
          <select
            class="docx-pagesetup__input"
            data-testid="docx-page-setup-paper-size"
            :value="preset"
            @change="onPresetChange"
          >
            <option v-for="option of paperSizeOptions" :key="option.value" :value="option.value">
              {{ option.label }}
            </option>
          </select>
        </label>
        <div class="docx-pagesetup__row">
          <label class="docx-pagesetup__field">
            <span>{{ m.docx.pageSetup.width }}</span>
            <input
              class="docx-pagesetup__input"
              type="number"
              min="0"
              step="0.01"
              inputmode="decimal"
              data-testid="docx-page-setup-width"
              :value="widthInches"
              @input="onSizeInput('width', $event)"
            />
          </label>
          <label class="docx-pagesetup__field">
            <span>{{ m.docx.pageSetup.height }}</span>
            <input
              class="docx-pagesetup__input"
              type="number"
              min="0"
              step="0.01"
              inputmode="decimal"
              data-testid="docx-page-setup-height"
              :value="heightInches"
              @input="onSizeInput('height', $event)"
            />
          </label>
        </div>
        <fieldset class="docx-pagesetup__orientation">
          <legend>{{ m.docx.pageSetup.orientation }}</legend>
          <label class="docx-pagesetup__radio">
            <input
              type="radio"
              name="docx-page-setup-orientation"
              value="portrait"
              data-testid="docx-page-setup-portrait"
              :checked="orientation === 'portrait'"
              @change="setOrientation('portrait')"
            />
            <span>{{ m.docx.pageSetup.orientationPortrait }}</span>
          </label>
          <label class="docx-pagesetup__radio">
            <input
              type="radio"
              name="docx-page-setup-orientation"
              value="landscape"
              data-testid="docx-page-setup-landscape"
              :checked="orientation === 'landscape'"
              @change="setOrientation('landscape')"
            />
            <span>{{ m.docx.pageSetup.orientationLandscape }}</span>
          </label>
        </fieldset>
      </section>

      <section class="docx-pagesetup__group" data-testid="docx-page-setup-margins">
        <h3 class="docx-pagesetup__group-title">{{ m.docx.pageSetup.margins }}</h3>
        <div class="docx-pagesetup__grid">
          <label
            v-for="field of marginFields"
            :key="field.key"
            class="docx-pagesetup__field"
          >
            <span>{{ field.label }}</span>
            <input
              class="docx-pagesetup__input"
              type="number"
              min="0"
              step="0.05"
              inputmode="decimal"
              :data-testid="`docx-page-setup-margin-${field.key}`"
              :value="marginInches[field.key]"
              @input="onMarginInput(field.key, $event)"
            />
          </label>
        </div>
        <p class="docx-pagesetup__hint">{{ m.docx.pageSetup.units }}</p>
      </section>

      <p
        v-if="validationMessage"
        class="docx-pagesetup__error"
        data-testid="docx-page-setup-error"
        role="alert"
      >{{ validationMessage }}</p>

      <div class="docx-pagesetup__actions">
        <button
          type="button"
          class="docx-pagesetup__button"
          data-testid="docx-page-setup-cancel"
          @click="emit('close')"
        >{{ m.docx.pageSetup.cancel }}</button>
        <button
          type="button"
          class="docx-pagesetup__button docx-pagesetup__button--primary"
          data-testid="docx-page-setup-apply"
          :disabled="Boolean(validationMessage) || readOnly"
          @click="apply"
        >{{ m.docx.pageSetup.apply }}</button>
      </div>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { computed, onMounted, ref, watch } from "vue"
import {
  TWIPS_PER_INCH,
  DOCX_PAGE_SIZE_PRESETS,
  matchDocxPageSizePreset,
  normalizePageSizeOrientation,
  type DocxPageMarginsUpdate,
  type DocxPageOrientation,
  type DocxPageSizePreset,
  type DocxPageSizeUpdate,
} from "@arcships/docx-core"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { useDocxMessages } from "../i18n"
import type {
  DocxPageSetupApplyPatch,
  DocxPageSetupDialogProps,
} from "./page-setup-dialog-types"

type MarginKey = "top" | "right" | "bottom" | "left"

const props = withDefaults(defineProps<DocxPageSetupDialogProps>(), {
  sections: () => [],
  readOnly: false,
  isDark: false,
})

const emit = defineEmits<{
  apply: [patch: DocxPageSetupApplyPatch]
  close: []
  "update:sectionIndex": [sectionIndex: number]
}>()

const m = useDocxMessages()

// Focus the dialog so Escape and the tab order start inside it.
const dialogRef = ref<HTMLElement | undefined>()
onMounted(() => {
  dialogRef.value?.focus?.()
})

const paperSizeOptions = computed(() => [
  { value: "custom", label: m.value.docx.pageSetup.paperSizeCustom },
  { value: "letter", label: m.value.docx.pageSetup.presetLetter },
  { value: "legal", label: m.value.docx.pageSetup.presetLegal },
  { value: "tabloid", label: m.value.docx.pageSetup.presetTabloid },
  { value: "a3", label: m.value.docx.pageSetup.presetA3 },
  { value: "a4", label: m.value.docx.pageSetup.presetA4 },
  { value: "a5", label: m.value.docx.pageSetup.presetA5 },
  { value: "b5", label: m.value.docx.pageSetup.presetB5 },
])

function toInches(twips: number): number {
  return Math.round((twips / TWIPS_PER_INCH) * 10000) / 10000
}

function toTwips(inches: number): number {
  return Math.max(0, Math.round(inches * TWIPS_PER_INCH))
}

/** Trim trailing zeros so "1.00" reads as "1" in the number inputs. */
function formatInches(inches: number): string {
  return String(Math.round(inches * 1000) / 1000)
}

const widthInches = ref(formatInches(toInches(props.pageSize.widthTwips)))
const heightInches = ref(formatInches(toInches(props.pageSize.heightTwips)))
const orientation = ref<DocxPageOrientation>(props.pageSize.orientation)
const marginInches = ref<Record<MarginKey, string>>({
  top: formatInches(toInches(props.margins.topTwips)),
  right: formatInches(toInches(props.margins.rightTwips)),
  bottom: formatInches(toInches(props.margins.bottomTwips)),
  left: formatInches(toInches(props.margins.leftTwips)),
})

const preset = computed(() => {
  const matched = matchDocxPageSizePreset({
    widthTwips: toTwips(Number(widthInches.value)),
    heightTwips: toTwips(Number(heightInches.value)),
    orientation: orientation.value,
  })
  return matched ?? "custom"
})

const marginFields = computed<
  Array<{ key: MarginKey; label: string }>
>(() => [
  { key: "top", label: m.value.docx.pageSetup.marginTop },
  { key: "right", label: m.value.docx.pageSetup.marginRight },
  { key: "bottom", label: m.value.docx.pageSetup.marginBottom },
  { key: "left", label: m.value.docx.pageSetup.marginLeft },
])

const unitsHint = computed(() => m.value.docx.pageSetup.units)

function syncFromProps(): void {
  widthInches.value = formatInches(toInches(props.pageSize.widthTwips))
  heightInches.value = formatInches(toInches(props.pageSize.heightTwips))
  orientation.value = props.pageSize.orientation
  marginInches.value = {
    top: formatInches(toInches(props.margins.topTwips)),
    right: formatInches(toInches(props.margins.rightTwips)),
    bottom: formatInches(toInches(props.margins.bottomTwips)),
    left: formatInches(toInches(props.margins.leftTwips)),
  }
}

watch(
  () => [props.pageSize, props.margins, props.sectionIndex] as const,
  () => syncFromProps()
)

const hasInvalidNumber = computed(() => {
  const values = [
    Number(widthInches.value),
    Number(heightInches.value),
    ...Object.values(marginInches.value).map((value) => Number(value)),
  ]
  return values.some((value) => !Number.isFinite(value) || value < 0)
})

const validationMessage = computed(() =>
  hasInvalidNumber.value ? m.value.docx.pageSetup.invalidValue : ""
)

function onSectionChange(event: Event): void {
  const value = Number((event.target as HTMLSelectElement).value)
  if (!Number.isFinite(value)) return
  emit("update:sectionIndex", Math.max(0, Math.round(value)))
}

function onPresetChange(event: Event): void {
  const value = (event.target as HTMLSelectElement).value as DocxPageSizePreset | "custom"
  if (value === "custom") return
  const size = DOCX_PAGE_SIZE_PRESETS[value]
  if (!size) return
  const normalized = normalizePageSizeOrientation({
    widthTwips: size.widthTwips,
    heightTwips: size.heightTwips,
    orientation: orientation.value,
  })
  widthInches.value = formatInches(toInches(normalized.widthTwips))
  heightInches.value = formatInches(toInches(normalized.heightTwips))
}

function setOrientation(next: DocxPageOrientation): void {
  if (orientation.value === next) return
  orientation.value = next
  // Paper dimensions follow the orientation, the way Word's radio button flips
  // an A4 portrait sheet into A4 landscape.
  const width = Number(widthInches.value)
  const height = Number(heightInches.value)
  if (next === "landscape" && width < height) {
    widthInches.value = formatInches(height)
    heightInches.value = formatInches(width)
  } else if (next === "portrait" && width > height) {
    widthInches.value = formatInches(height)
    heightInches.value = formatInches(width)
  }
}

function onSizeInput(edge: "width" | "height", event: Event): void {
  const value = (event.target as HTMLInputElement).value
  if (edge === "width") widthInches.value = value
  else heightInches.value = value
}

function onMarginInput(edge: MarginKey, event: Event): void {
  marginInches.value = {
    ...marginInches.value,
    [edge]: (event.target as HTMLInputElement).value,
  }
}

function apply(): void {
  if (validationMessage.value || props.readOnly) return
  const pageSize: DocxPageSizeUpdate = {
    widthTwips: toTwips(Number(widthInches.value)),
    heightTwips: toTwips(Number(heightInches.value)),
    orientation: orientation.value,
  }
  const margins: DocxPageMarginsUpdate = {
    topTwips: toTwips(Number(marginInches.value.top)),
    rightTwips: toTwips(Number(marginInches.value.right)),
    bottomTwips: toTwips(Number(marginInches.value.bottom)),
    leftTwips: toTwips(Number(marginInches.value.left)),
  }
  emit("apply", { pageSize, margins })
}
</script>

<style scoped>
.docx-pagesetup-backdrop {
  position: fixed;
  inset: 0;
  background: rgb(15 23 42 / 45%);
  z-index: 60;
}

.docx-pagesetup {
  position: fixed;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  z-index: 61;
  display: flex;
  flex-direction: column;
  gap: 12px;
  width: min(440px, calc(100vw - 32px));
  max-height: calc(100vh - 48px);
  overflow-y: auto;
  padding: 16px 18px 14px;
  border: 1px solid #d7dbe3;
  border-radius: 10px;
  background: #ffffff;
  box-shadow: 0 18px 48px rgb(15 23 42 / 22%);
  color: #1f2937;
  font: 13px/1.45 system-ui, -apple-system, "Segoe UI", sans-serif;
}

.docx-pagesetup--dark {
  border-color: #3f4756;
  background: #1c2129;
  color: #e5e7eb;
}

.docx-pagesetup__header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
}

.docx-pagesetup__title {
  margin: 0;
  font-size: 15px;
  font-weight: 600;
}

.docx-pagesetup__close {
  border: none;
  background: transparent;
  color: inherit;
  font-size: 18px;
  line-height: 1;
  cursor: pointer;
}

.docx-pagesetup__group {
  display: flex;
  flex-direction: column;
  gap: 8px;
  padding-top: 4px;
  border-top: 1px solid #e5e7eb;
}

.docx-pagesetup--dark .docx-pagesetup__group {
  border-top-color: #3f4756;
}

.docx-pagesetup__group-title {
  margin: 4px 0 0;
  font-size: 12px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  color: #6b7280;
}

.docx-pagesetup--dark .docx-pagesetup__group-title {
  color: #9ca3af;
}

.docx-pagesetup__field {
  display: flex;
  flex-direction: column;
  gap: 4px;
  font-size: 12px;
  color: inherit;
}

.docx-pagesetup__row {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 10px;
}

.docx-pagesetup__grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 10px;
}

.docx-pagesetup__input {
  padding: 5px 7px;
  border: 1px solid #cbd2dc;
  border-radius: 6px;
  background: #ffffff;
  color: inherit;
  font: inherit;
}

.docx-pagesetup--dark .docx-pagesetup__input {
  border-color: #4b5563;
  background: #11151b;
}

.docx-pagesetup__input:focus-visible {
  outline: 2px solid #2563eb;
  outline-offset: 1px;
}

.docx-pagesetup__orientation {
  display: flex;
  gap: 14px;
  margin: 0;
  padding: 0;
  border: none;
}

.docx-pagesetup__orientation legend {
  margin-bottom: 4px;
  padding: 0;
  font-size: 12px;
  color: #6b7280;
}

.docx-pagesetup--dark .docx-pagesetup__orientation legend {
  color: #9ca3af;
}

.docx-pagesetup__radio {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
}

.docx-pagesetup__hint {
  margin: 0;
  font-size: 11px;
  color: #6b7280;
}

.docx-pagesetup--dark .docx-pagesetup__hint {
  color: #9ca3af;
}

.docx-pagesetup__error {
  margin: 0;
  color: #b91c1c;
  font-size: 12px;
}

.docx-pagesetup__actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  padding-top: 4px;
}

.docx-pagesetup__button {
  padding: 6px 14px;
  border: 1px solid #cbd2dc;
  border-radius: 6px;
  background: #ffffff;
  color: inherit;
  font: inherit;
  cursor: pointer;
}

.docx-pagesetup--dark .docx-pagesetup__button {
  border-color: #4b5563;
  background: #11151b;
}

.docx-pagesetup__button--primary {
  border-color: #2563eb;
  background: #2563eb;
  color: #ffffff;
}

.docx-pagesetup__button:disabled {
  opacity: 0.55;
  cursor: default;
}
</style>
