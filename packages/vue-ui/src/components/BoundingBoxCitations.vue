<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue"
import {
  CancelCircleIcon,
  HashIcon,
  InputTextIcon,
  SecondBracketIcon,
  SourceCodeSquareIcon,
  TextCheckIcon,
  Undo02Icon,
} from "@hugeicons/core-free-icons"
import { HugeIcon } from "./HugeIcon"
import OfficeButton from "./ui/OfficeButton.vue"
import OfficeInput from "./ui/OfficeInput.vue"
import { useOfficeUiMessages } from "../i18n"
import type {
  BoundingBoxCitationsProps,
  BoundingBoxField,
  BoundingBoxReviewField,
  BoundingBoxReviewFieldSchema,
  BoundingBoxReviewLocation,
  OfficeJsonPrimitive,
  OfficeJsonValue,
} from "../types"

const props = withDefaults(defineProps<BoundingBoxCitationsProps>(), {
  file: undefined,
  fields: () => [],
  reviewFields: undefined,
  pageImage: undefined,
  pageWidth: undefined,
  pageHeight: undefined,
  activeFieldKey: undefined,
  showExpected: true,
  expectedValues: undefined,
  readOnly: false,
  view: "panel",
  loading: false,
  error: undefined,
  emptyMessage: undefined,
  className: "",
})

const emit = defineEmits<{
  "field-click": [field: BoundingBoxField]
  "field-focus": [field: BoundingBoxReviewField]
  "update:activeFieldKey": [key: string]
  "update:expected": [key: string, value: OfficeJsonValue]
}>()

const m = useOfficeUiMessages()

const activeTab = ref<"form" | "json">("form")
const localActiveKey = ref<string | null>(null)
const draft = ref<Record<string, OfficeJsonValue>>({})
const cardRefs = ref<Record<string, HTMLDivElement | null>>({})

const reviewMode = computed(() => props.reviewFields !== undefined)
const reviewFields = computed<BoundingBoxReviewField[]>(
  () => props.reviewFields ?? [],
)

const activeField = computed<BoundingBoxReviewField | undefined>(() => {
  const key = props.activeFieldKey ?? localActiveKey.value
  if (key) {
    const match = reviewFields.value.find((field) => field.key === key)
    if (match) return match
  }
  return reviewFields.value[0]
})

const activeLocation = computed<BoundingBoxReviewLocation | undefined>(
  () => activeField.value?.location,
)

const pageImage = computed(() => props.pageImage ?? props.file ?? "")
const splitLayout = computed(() => props.view === "split" && reviewMode.value)
const emptyCopy = computed(
  () => props.emptyMessage ?? m.value.vueUi.boundingBox.noFields,
)

watch(
  () => props.reviewFields,
  (fields) => {
    const next: Record<string, OfficeJsonValue> = {}
    for (const field of fields ?? []) next[field.key] = field.expected
    draft.value = next
  },
  { immediate: true, deep: false },
)

watch(
  () => props.expectedValues,
  (values) => {
    if (values) draft.value = { ...draft.value, ...values }
  },
  { immediate: true, deep: true },
)

// Legacy page-card state -------------------------------------------------
const selectedFieldId = ref<string | null>(null)
const groupedByPage = computed(() => {
  const map = new Map<number, BoundingBoxField[]>()
  for (const field of props.fields) {
    const existing = map.get(field.page) ?? []
    existing.push(field)
    map.set(field.page, existing)
  }
  return Array.from(map.entries()).sort(([a], [b]) => a - b)
})

function confidenceClass(confidence: number): string {
  if (confidence >= 0.8) return "confidence-high"
  if (confidence >= 0.5) return "confidence-medium"
  return "confidence-low"
}

function fieldStyle(field: BoundingBoxField) {
  const [left, top, right, bottom] = field.rect
  return {
    left: `${left * 100}%`,
    top: `${top * 100}%`,
    width: `${(right - left) * 100}%`,
    height: `${(bottom - top) * 100}%`,
  }
}

function selectLegacyField(field: BoundingBoxField) {
  selectedFieldId.value = field.id
  emit("field-click", field)
}

// Extend review mode -----------------------------------------------------
function isFieldActive(field: BoundingBoxReviewField): boolean {
  return activeField.value?.key === field.key
}

function isModified(field: BoundingBoxReviewField): boolean {
  const value = draft.value[field.key]
  return props.showExpected && !jsonEquals(value, field.expected)
}

function jsonEquals(left: OfficeJsonValue | undefined, right: OfficeJsonValue): boolean {
  if (left === right) return true
  if (typeof left !== typeof right || left === null || right === null) return false
  if (Array.isArray(left) && Array.isArray(right)) {
    return left.length === right.length && left.every((entry, index) => jsonEquals(entry, right[index]!))
  }
  if (typeof left === "object" && typeof right === "object") {
    const leftKeys = Object.keys(left)
    const rightKeys = Object.keys(right)
    if (leftKeys.length !== rightKeys.length) return false
    return leftKeys.every((key) =>
      jsonEquals(
        (left as Record<string, OfficeJsonValue>)[key],
        (right as Record<string, OfficeJsonValue>)[key]!,
      ),
    )
  }
  return false
}

function formatValue(value: OfficeJsonValue | undefined): string {
  if (value === null || value === undefined) return "null"
  if (typeof value === "string") return value
  if (typeof value === "number" || typeof value === "boolean") return String(value)
  return JSON.stringify(value)
}

function primitiveValue(field: BoundingBoxReviewField): OfficeJsonPrimitive {
  const value = draft.value[field.key] ?? field.expected
  if (value === null || typeof value !== "object") return value
  return formatValue(value)
}

function updateExpected(field: BoundingBoxReviewField, value: OfficeJsonValue) {
  draft.value = { ...draft.value, [field.key]: value }
  emit("update:expected", field.key, value)
}

function onInputValue(field: BoundingBoxReviewField, value: string) {
  updateExpected(field, value === "" ? null : value)
}

function onNumberValue(field: BoundingBoxReviewField, value: string) {
  updateExpected(field, value === "" ? null : Number(value))
}

function setNull(field: BoundingBoxReviewField) {
  updateExpected(field, null)
}

function revertField(field: BoundingBoxReviewField) {
  updateExpected(field, field.expected)
}

function fieldIcon(type: BoundingBoxReviewFieldSchema["type"]) {
  if (type === "number" || type === "integer") return HashIcon
  if (type === "boolean") return TextCheckIcon
  if (type === "array") return SecondBracketIcon
  if (type === "object") return SourceCodeSquareIcon
  return InputTextIcon
}

function focusField(field: BoundingBoxReviewField, select = false) {
  localActiveKey.value = field.key
  emit("update:activeFieldKey", field.key)
  emit("field-focus", field)
  if (select) emit("field-click", legacyFieldFor(field))
}

function legacyFieldFor(field: BoundingBoxReviewField): BoundingBoxField {
  const area = field.location?.area
  return {
    id: field.key,
    label: field.schema.title ?? field.key,
    page: field.location?.page ?? 1,
    rect: area
      ? [area.left / 100, area.top / 100, (area.left + area.width) / 100, (area.top + area.height) / 100]
      : [0, 0, 0, 0],
    value: formatValue(field.actual),
  }
}

async function moveFocus(field: BoundingBoxReviewField, direction: 1 | -1) {
  const index = reviewFields.value.findIndex((entry) => entry.key === field.key)
  const next = reviewFields.value[index + direction]
  if (!next) return
  focusField(next)
  await nextTick()
  cardRefs.value[next.key]?.focus?.()
}

function onCardKeydown(event: KeyboardEvent, field: BoundingBoxReviewField) {
  if (event.key === "ArrowDown") {
    event.preventDefault()
    void moveFocus(field, 1)
  } else if (event.key === "ArrowUp") {
    event.preventDefault()
    void moveFocus(field, -1)
  } else if (event.key === "Escape") {
    event.preventDefault()
    localActiveKey.value = null
  }
}

function setCardRef(element: unknown, key: string) {
  cardRefs.value[key] = (element as HTMLDivElement | null) ?? null
}

function highlightStyle(location: BoundingBoxReviewLocation) {
  const area = location.area
  return {
    left: `${area.left}%`,
    top: `${area.top}%`,
    width: `${area.width}%`,
    height: `${area.height}%`,
  }
}

const jsonValue = computed(() => {
  const output: Record<string, OfficeJsonValue> = {}
  for (const field of reviewFields.value) output[field.key] = field.actual
  return JSON.stringify(output, null, 2)
})
</script>

<template>
  <!-- Extend review mode: page highlight + HumanReviewPanel -->
  <div
    v-if="reviewMode"
    :class="[
      'bounding-box-citations',
      'office-ui',
      splitLayout ? 'bounding-box-citations--split' : 'bounding-box-citations--panel',
      props.className,
    ]"
    data-slot="bounding-box-citations"
  >
    <div v-if="splitLayout" class="review-page" data-slot="bounding-box-page">
      <div v-if="props.error" class="review-state review-state--error" role="alert">
        {{ props.error }}
      </div>
      <div v-else-if="props.loading" class="review-page__frame review-page__frame--loading" role="status">
        <div class="review-page__skeleton" />
      </div>
      <div v-else class="review-page__frame">
        <img
          v-if="pageImage"
          :src="pageImage"
          :alt="m.vueUi.boundingBox.imageAlt"
          class="review-page__image"
        />
        <div
          v-if="activeLocation"
          class="review-highlight"
          data-slot="bounding-box-highlight"
          :style="highlightStyle(activeLocation)"
        />
      </div>
    </div>

    <div class="review-panel" data-slot="bounding-box-panel">
      <div class="review-panel__toolbar">
        <div class="review-tabs" role="tablist" aria-label="Review view">
          <button
            type="button"
            role="tab"
            :aria-selected="activeTab === 'form'"
            :class="['review-tab', { 'is-active': activeTab === 'form' }]"
            data-slot="bounding-box-tab-form"
            @click="activeTab = 'form'"
          >
            <HugeIcon :icon="TextCheckIcon" class="review-tab__icon" />
            {{ m.vueUi.boundingBox.formTab }}
          </button>
          <button
            type="button"
            role="tab"
            :aria-selected="activeTab === 'json'"
            :class="['review-tab', { 'is-active': activeTab === 'json' }]"
            data-slot="bounding-box-tab-json"
            @click="activeTab = 'json'"
          >
            <HugeIcon :icon="SourceCodeSquareIcon" class="review-tab__icon" />
            {{ m.vueUi.boundingBox.jsonTab }}
          </button>
        </div>
      </div>

      <div v-if="activeTab === 'form'" class="review-scroll" data-slot="bounding-box-form">
        <div v-if="props.error" class="review-state-wrap">
          <div class="review-state review-state--error" role="alert">{{ props.error }}</div>
        </div>
        <div v-else-if="props.loading" class="review-state-wrap">
          <div class="review-state" role="status" data-slot="bounding-box-loading">
            {{ m.vueUi.boundingBox.loading }}
          </div>
        </div>
        <div v-else-if="reviewFields.length === 0" class="review-state-wrap">
          <div class="review-state" role="status" data-slot="bounding-box-empty">
            {{ emptyCopy }}
          </div>
        </div>
        <div v-else class="review-list" data-slot="bounding-box-list">
          <div
            v-for="field in reviewFields"
            :key="field.key"
            :ref="(element) => setCardRef(element, field.key)"
            :class="['review-card', { 'is-active': isFieldActive(field) }]"
            :tabindex="0"
            :data-review-field-key="field.key"
            data-slot="bounding-box-field"
            @mouseenter="focusField(field)"
            @focus.capture="focusField(field)"
            @keydown="onCardKeydown($event, field)"
          >
            <div class="review-card__header">
              <div class="review-card__heading">
                <div class="review-card__title-row">
                  <div class="review-card__title">{{ field.schema.title ?? field.key }}</div>
                  <span
                    :class="['review-card__modified-dot', { 'is-hidden': !isModified(field) }]"
                    :title="m.vueUi.boundingBox.modified"
                  />
                </div>
                <div class="review-card__key">{{ field.key }}</div>
              </div>
              <div class="review-card__actions">
                <OfficeButton
                  v-if="props.showExpected && !props.readOnly && isModified(field)"
                  variant="ghost"
                  size="icon-sm"
                  class="review-card__icon-button"
                  :aria-label="`${m.vueUi.boundingBox.revert} ${field.key}`"
                  data-slot="bounding-box-revert"
                  @click.stop="revertField(field)"
                >
                  <HugeIcon :icon="Undo02Icon" class="review-card__action-icon" />
                </OfficeButton>
                <OfficeButton
                  v-if="props.showExpected && !props.readOnly"
                  variant="ghost"
                  size="icon-sm"
                  class="review-card__icon-button"
                  :aria-label="`${m.vueUi.boundingBox.setNull} ${field.key}`"
                  data-slot="bounding-box-set-null"
                  @click.stop="setNull(field)"
                >
                  <HugeIcon :icon="CancelCircleIcon" class="review-card__action-icon" />
                </OfficeButton>
                <div class="review-card__type">
                  <HugeIcon :icon="fieldIcon(field.schema.type)" class="review-card__type-icon" />
                  {{ field.schema.type }}
                </div>
              </div>
            </div>

            <div :class="['review-values', { 'review-values--single': !props.showExpected }]">
              <div class="review-value">
                <p v-if="field.schema.description" class="review-value__description">
                  {{ field.schema.description }}
                </p>
                <div class="review-value__label">{{ m.vueUi.boundingBox.actual }}</div>
                <div class="review-value__static">{{ formatValue(field.actual) }}</div>
              </div>
              <div v-if="props.showExpected" class="review-value">
                <div class="review-value__label">{{ m.vueUi.boundingBox.expected }}</div>
                <OfficeInput
                  v-if="field.schema.type === 'number' || field.schema.type === 'integer'"
                  type="number"
                  :model-value="String(primitiveValue(field) ?? '')"
                  :readonly="props.readOnly"
                  @update:model-value="onNumberValue(field, $event)"
                />
                <OfficeInput
                  v-else
                  :model-value="String(primitiveValue(field) ?? '')"
                  :readonly="props.readOnly"
                  @update:model-value="onInputValue(field, $event)"
                />
              </div>
            </div>
          </div>
        </div>
      </div>

      <div v-else class="review-json" data-slot="bounding-box-json">
        <pre class="review-json__code">{{ jsonValue }}</pre>
      </div>
    </div>
  </div>

  <!-- Legacy page-card mode (pre-0.8 host API) -->
  <div v-else :class="['bounding-box-citations', 'office-ui', props.className]">
    <div
      v-for="[page, pageFields] in groupedByPage"
      :key="page"
      class="citation-page-card"
    >
      <h4 class="citation-page-title">
        Page {{ page }} &mdash; {{ pageFields.length }} field{{ pageFields.length > 1 ? "s" : "" }}
      </h4>
      <div class="citation-page-preview">
        <img :src="props.file" :alt="m.vueUi.boundingBox.imageAlt" class="citation-image" />
        <button
          v-for="field in pageFields"
          :key="field.id"
          type="button"
          :class="[
            'citation-box',
            confidenceClass(field.confidence ?? 0.5),
            { 'is-selected': selectedFieldId === field.id },
          ]"
          :style="fieldStyle(field)"
          :title="`${field.label}: ${field.value ?? ''}`"
          :aria-pressed="selectedFieldId === field.id"
          @click="selectLegacyField(field)"
        >
          <span class="citation-box-label">{{ field.label }}</span>
        </button>
      </div>
      <ul class="citation-list">
        <li v-for="field in pageFields" :key="field.id">
          <button
            type="button"
            :class="[
              'citation-list-item',
              confidenceClass(field.confidence ?? 0.5),
              { 'is-selected': selectedFieldId === field.id },
            ]"
            :aria-pressed="selectedFieldId === field.id"
            @click="selectLegacyField(field)"
          >
            <span class="citation-list-main">
              <span class="citation-label">{{ field.label }}</span>
              <span v-if="field.value" class="citation-value">{{ field.value }}</span>
            </span>
            <span v-if="field.confidence !== undefined" class="citation-confidence">
              {{ (field.confidence * 100).toFixed(0) }}%
            </span>
          </button>
        </li>
      </ul>
    </div>
    <p v-if="props.fields.length === 0" class="empty-state">
      {{ emptyCopy }}
    </p>
  </div>
</template>

<style scoped>
.bounding-box-citations {
  min-width: 0;
  color: var(--office-foreground);
}
.bounding-box-citations,
.bounding-box-citations * {
  box-sizing: border-box;
}
.bounding-box-citations--split {
  display: flex;
  min-height: 0;
  align-items: stretch;
  gap: var(--office-ui-space-3);
}
.review-page {
  min-width: 0;
  flex: 1 1 auto;
}
.review-page__frame {
  position: relative;
  width: 100%;
  min-height: 560px;
  overflow: hidden;
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-muted);
}
.review-page__frame--loading {
  background: var(--office-muted);
}
.review-page__skeleton {
  position: absolute;
  inset: 0;
  background: var(--office-muted);
  animation: review-pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite;
}
.review-page__image {
  display: block;
  width: 100%;
  height: 100%;
  object-fit: contain;
}
.review-highlight {
  pointer-events: none;
  position: absolute;
  z-index: 10;
  border: 1px solid color-mix(in oklab, oklch(62.3% .214 259.815) 70%, transparent);
  background: color-mix(in oklab, oklch(62.3% .214 259.815) 12%, transparent);
  box-shadow: 0 4px 16px color-mix(in oklab, oklch(62.3% .214 259.815) 10%, transparent);
}
.review-panel {
  display: flex;
  height: 560px;
  min-height: 0;
  flex-direction: column;
  gap: 0;
  background: var(--office-background);
}
.bounding-box-citations--split .review-panel {
  height: auto;
  min-height: 560px;
  flex: 0 0 42%;
  max-width: 60%;
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
}
.bounding-box-citations--panel .review-panel {
  width: 100%;
}
.review-panel__toolbar {
  display: flex;
  min-height: 48px;
  align-items: center;
  justify-content: space-between;
  gap: var(--office-ui-space-3);
  border-bottom: var(--office-border-width) solid var(--office-border);
  padding-inline: var(--office-ui-space-3);
}
.review-tabs {
  position: relative;
  z-index: 0;
  display: flex;
  width: fit-content;
  height: 28px;
  align-items: center;
  justify-content: center;
  column-gap: 2px;
  border-radius: var(--office-radius-lg);
  background: var(--office-muted);
  padding: 2px;
  color: color-mix(in oklab, var(--office-muted-foreground) 72%, transparent);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
.review-tab {
  position: relative;
  display: flex;
  height: 24px;
  flex-shrink: 0;
  flex-grow: 1;
  cursor: pointer;
  align-items: center;
  justify-content: center;
  gap: 6px;
  margin: 0;
  border: 1px solid transparent;
  border-radius: var(--office-radius-md);
  padding-inline: calc(var(--office-ui-space-2-5) - 1px);
  background: transparent;
  color: inherit;
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  font-weight: 500;
  white-space: nowrap;
  outline: none;
  transition:
    color 120ms ease-out,
    background-color 120ms ease-out,
    box-shadow 120ms ease-out;
}
.review-tab:hover {
  color: var(--office-muted-foreground);
}
.review-tab.is-active {
  background: var(--office-background);
  color: var(--office-foreground);
  box-shadow: 0 1px 2px 0 color-mix(in oklab, black 5%, transparent);
}
.review-tab:focus-visible {
  outline: none;
  box-shadow: 0 0 0 2px var(--office-ring);
}
.review-tab__icon {
  width: 16px;
  height: 16px;
  flex-shrink: 0;
  margin-inline: -2px;
}
.review-scroll {
  min-height: 0;
  flex: 1 1 auto;
  overflow-y: auto;
}
.review-list {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-3);
  padding: var(--office-ui-space-3);
}
.review-card {
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-background);
  padding: var(--office-ui-space-3);
  outline: none;
  transition:
    border-color 120ms ease-out,
    background-color 120ms ease-out,
    box-shadow 120ms ease-out;
}
.review-card:hover,
.review-card:focus-within {
  border-color: color-mix(in oklab, oklch(62.3% .214 259.815) 50%, transparent);
  background: color-mix(in oklab, oklch(62.3% .214 259.815) 5%, transparent);
}
.review-card:focus-visible {
  box-shadow: 0 0 0 2px color-mix(in oklab, oklch(62.3% .214 259.815) 20%, transparent);
}
.review-card.is-active {
  border-color: color-mix(in oklab, oklch(62.3% .214 259.815) 60%, transparent);
  background: color-mix(in oklab, oklch(62.3% .214 259.815) 5%, transparent);
  box-shadow: 0 0 0 1px color-mix(in oklab, oklch(62.3% .214 259.815) 8%, transparent);
}
.review-card__header {
  display: flex;
  min-height: 32px;
  align-items: flex-start;
  justify-content: space-between;
  gap: var(--office-ui-space-3);
  margin-bottom: var(--office-ui-space-3);
}
.review-card__heading {
  min-width: 0;
}
.review-card__title-row {
  display: flex;
  min-width: 0;
  align-items: center;
  gap: var(--office-ui-space-2);
}
.review-card__title {
  overflow: hidden;
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  font-weight: 500;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.review-card__modified-dot {
  width: 8px;
  height: 8px;
  flex-shrink: 0;
  border-radius: calc(infinity * 1px);
  background: oklch(82.8% .189 84.429);
}
.review-card__modified-dot.is-hidden {
  opacity: 0;
}
.review-card__key {
  overflow: hidden;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
  text-overflow: ellipsis;
  white-space: nowrap;
}
.review-card__actions {
  display: flex;
  flex-shrink: 0;
  align-items: center;
  gap: var(--office-ui-space-1);
}
.review-card__icon-button {
  color: var(--office-muted-foreground);
}
.review-card__action-icon {
  width: 16px;
  height: 16px;
}
.review-card__type {
  display: flex;
  height: 24px;
  align-items: center;
  gap: var(--office-ui-space-1);
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: color-mix(in oklab, var(--office-muted) 50%, transparent);
  padding-inline: 6px;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
.review-card__type-icon {
  width: 14px;
  height: 14px;
}
.review-values {
  display: grid;
  gap: var(--office-ui-space-2);
}
.review-values--single {
  grid-template-columns: minmax(0, 1fr);
}
.review-value {
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: color-mix(in oklab, var(--office-muted) 30%, transparent);
  padding: var(--office-ui-space-2);
}
.review-value__description {
  margin: 0 0 var(--office-ui-space-2);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
.review-value__label {
  margin-bottom: var(--office-ui-space-1);
  color: var(--office-muted-foreground);
  font-size: 11px;
  line-height: 16px;
  font-weight: 500;
}
.review-value__static {
  min-height: 28px;
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  padding: 6px var(--office-ui-space-2);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  overflow-wrap: anywhere;
}
.review-state-wrap {
  padding: var(--office-ui-space-3);
}
.review-state {
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: color-mix(in oklab, var(--office-muted) 30%, transparent);
  padding: var(--office-ui-space-3);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
}
.review-state--error {
  border-color: color-mix(in oklab, var(--office-destructive) 45%, transparent);
  background: color-mix(in oklab, var(--office-destructive) 8%, transparent);
  color: var(--office-destructive);
}
.review-json {
  min-height: 0;
  flex: 1 1 auto;
  overflow: auto;
  background: color-mix(in oklab, var(--office-muted) 30%, transparent);
  padding: var(--office-ui-space-3);
}
.review-json__code {
  margin: 0;
  color: var(--office-foreground);
  font-family: var(--office-font-mono);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}
@media (min-width: 640px) {
  .review-values {
    grid-template-columns: repeat(2, minmax(0, 1fr));
  }
  .review-values--single {
    grid-template-columns: minmax(0, 1fr);
  }
  .review-tabs {
    height: 28px;
  }
  .review-tab {
    height: 24px;
  }
}
@keyframes review-pulse {
  50% {
    opacity: 0.5;
  }
}
@media (prefers-reduced-motion: reduce) {
  .review-page__skeleton {
    animation: none;
  }
}
@media (max-width: 767px) {
  .bounding-box-citations--split {
    flex-direction: column;
  }
  .bounding-box-citations--split .review-panel {
    flex: 1 1 auto;
    max-width: none;
    width: 100%;
  }
}

/* Legacy page-card mode ------------------------------------------------- */
.citation-page-card {
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-background);
  padding: var(--office-ui-space-4);
}
.citation-page-title {
  margin: 0 0 var(--office-ui-space-3);
  color: var(--office-foreground);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  font-weight: 500;
}
.citation-page-preview {
  position: relative;
  margin-bottom: var(--office-ui-space-3);
  aspect-ratio: 3 / 4;
  width: 100%;
  max-width: 384px;
  overflow: hidden;
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: color-mix(in oklab, var(--office-muted) 35%, transparent);
}
.citation-image {
  display: block;
  width: 100%;
  height: 100%;
  object-fit: cover;
  opacity: 0.9;
}
.citation-box {
  position: absolute;
  cursor: crosshair;
  border-width: 2px;
  border-style: solid;
  border-radius: 0;
  background: transparent;
  padding: 0;
  transition:
    filter 120ms ease-out,
    outline-color 120ms ease-out;
}
.citation-box:hover {
  filter: brightness(1.12);
}
.citation-box:focus-visible,
.citation-list-item:focus-visible {
  outline: 2px solid var(--office-ring);
  outline-offset: 2px;
}
.citation-box.is-selected {
  box-shadow: none;
  outline: 2px solid var(--office-foreground);
  outline-offset: 1px;
  z-index: 2;
}
.citation-box-label {
  position: absolute;
  left: -2px;
  top: -18px;
  max-width: 160px;
  overflow: hidden;
  border-radius: 0;
  background: color-mix(in oklab, var(--office-foreground) 88%, transparent);
  color: var(--office-background);
  font-family: var(--office-font-mono);
  font-size: 10px;
  font-weight: 600;
  line-height: 16px;
  padding: 0 4px;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.citation-list {
  display: flex;
  flex-direction: column;
  gap: 6px;
  margin: 0;
  padding: 0;
  list-style: none;
}
.citation-list-item {
  display: flex;
  width: 100%;
  cursor: pointer;
  align-items: flex-start;
  gap: var(--office-ui-space-2);
  border-width: 0 0 0 3px;
  border-style: solid;
  border-radius: 0;
  background: transparent;
  padding: 6px 8px;
  color: inherit;
  text-align: left;
  transition: background-color 120ms ease-out;
}
.citation-list-item:hover {
  background: color-mix(in oklab, var(--office-muted-foreground) 12%, transparent);
}
.citation-list-item.is-selected {
  box-shadow: none;
  background: color-mix(in oklab, #0ea5e9 12%, transparent);
}
.citation-list-main {
  min-width: 0;
  flex: 1 1 auto;
  font-size: 13px;
}
.citation-label {
  color: var(--office-foreground);
  font-weight: 600;
}
.citation-value {
  margin-left: var(--office-ui-space-2);
  color: var(--office-muted-foreground);
}
.citation-confidence {
  flex-shrink: 0;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
  font-variant-numeric: tabular-nums;
}
.confidence-high {
  border-color: #22c55e;
  background: color-mix(in oklab, #22c55e 8%, transparent);
}
.confidence-medium {
  border-color: oklch(76.9% .188 70.08);
  background: color-mix(in oklab, oklch(76.9% .188 70.08) 8%, transparent);
}
.confidence-low {
  border-color: #ef4444;
  background: color-mix(in oklab, #ef4444 8%, transparent);
}
.empty-state {
  padding: var(--office-ui-space-8) 0;
  color: var(--office-muted-foreground);
  text-align: center;
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
}
.bounding-box-citations:not(.bounding-box-citations--split):not(.bounding-box-citations--panel) {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-4);
}
.bounding-box-citations:not(.bounding-box-citations--split):not(.bounding-box-citations--panel) :deep(.citation-page-card) {
  margin: 0;
}
</style>
