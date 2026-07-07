<script setup lang="ts">
import { computed, ref } from "vue"

export interface BoundingBoxField {
  /** Unique stable id for the field */
  id: string
  /** Display label */
  label: string
  /** Page number (1-indexed) */
  page: number
  /** Normalized bounding box [left, top, right, bottom] in 0..1 */
  rect: [number, number, number, number]
  /** Extracted value text */
  value?: string
  /** Confidence score 0..1 */
  confidence?: number
}

export interface BoundingBoxCitationsProps {
  /** PDF file URL */
  file: string
  /** Fields with bounding boxes to highlight */
  fields: BoundingBoxField[]
  className?: string
}

const props = withDefaults(defineProps<BoundingBoxCitationsProps>(), {
  className: "",
})

const emit = defineEmits<{
  "field-click": [field: BoundingBoxField]
}>()

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

function selectField(field: BoundingBoxField) {
  selectedFieldId.value = field.id
  emit("field-click", field)
}
</script>

<template>
  <div :class="['bounding-box-citations', props.className]">
    <div
      v-for="[page, pageFields] in groupedByPage"
      :key="page"
      class="citation-page-card"
    >
      <h4 class="citation-page-title">
        Page {{ page }} &mdash; {{ pageFields.length }} field{{ pageFields.length > 1 ? "s" : "" }}
      </h4>
      <div class="citation-page-preview">
        <img :src="props.file" alt="Citation source preview" class="citation-image" />
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
          @click="selectField(field)"
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
            @click="selectField(field)"
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
      No fields to display
    </p>
  </div>
</template>

<style scoped>
.bounding-box-citations { display: flex; min-width: 0; flex-direction: column; gap: 16px; }
.citation-page-card {
  border: 1px solid var(--border, #e5e7eb);
  border-radius: var(--radius, 10px);
  background: var(--background, #fff);
  padding: 16px;
}
.citation-page-title { margin: 0 0 12px; color: var(--foreground, #111827); font-size: 14px; font-weight: 500; }
.citation-page-preview {
  position: relative;
  margin-bottom: 12px;
  aspect-ratio: 3 / 4;
  width: 100%;
  max-width: 384px;
  overflow: hidden;
  border: 1px solid var(--border, #e5e7eb);
  border-radius: var(--radius-md, 8px);
  background: color-mix(in oklch, var(--muted, #f5f5f5) 35%, transparent);
}
.citation-image { display: block; width: 100%; height: 100%; object-fit: cover; opacity: 0.9; }
.citation-box {
  position: absolute;
  cursor: crosshair;
  border-width: 2px;
  border-style: solid;
  border-radius: 0 !important;
  background: transparent;
  padding: 0;
  transition: filter 0.12s ease, outline-color 0.12s ease;
}
.citation-box:hover { filter: brightness(1.12); }
.citation-box:focus-visible,
.citation-list-item:focus-visible { outline: 2px solid #0ea5e9; outline-offset: 2px; }
.citation-box.is-selected {
  box-shadow: none;
  outline: 2px solid #0f172a;
  outline-offset: 1px;
  z-index: 2;
}
.citation-box-label {
  position: absolute;
  left: -2px;
  top: -18px;
  max-width: 160px;
  overflow: hidden;
  border-radius: 0 !important;
  background: rgb(15 23 42 / 0.88);
  color: #fff;
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", monospace;
  font-size: 10px;
  font-weight: 600;
  line-height: 16px;
  padding: 0 4px;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.citation-list { display: flex; flex-direction: column; gap: 6px; margin: 0; padding: 0; list-style: none; }
.citation-list-item {
  display: flex;
  width: 100%;
  cursor: pointer;
  align-items: flex-start;
  gap: 8px;
  border-width: 0 0 0 3px;
  border-style: solid;
  border-radius: 0 !important;
  background: transparent;
  padding: 6px 8px;
  text-align: left;
  transition: background-color 0.12s ease;
}
.citation-list-item:hover { background: rgb(148 163 184 / 0.12); }
.citation-list-item.is-selected { box-shadow: none; background: rgb(14 165 233 / 0.12); }
.citation-list-main { min-width: 0; flex: 1; font-size: 13px; }
.citation-label { color: var(--foreground, #111827); font-weight: 600; }
.citation-value { margin-left: 8px; color: var(--muted-foreground, #737373); }
.citation-confidence { flex-shrink: 0; color: var(--muted-foreground, #737373); font-size: 12px; font-variant-numeric: tabular-nums; }
.confidence-high { border-color: rgb(34 197 94); background: rgb(34 197 94 / 0.08); }
.confidence-medium { border-color: rgb(245 158 11); background: rgb(245 158 11 / 0.08); }
.confidence-low { border-color: rgb(239 68 68); background: rgb(239 68 68 / 0.08); }
.empty-state { padding: 32px 0; color: var(--muted-foreground, #737373); text-align: center; font-size: 14px; }
</style>
