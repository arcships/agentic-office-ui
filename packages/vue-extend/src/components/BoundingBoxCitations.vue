<script setup lang="ts">
import { computed } from "vue"

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

const groupedByPage = computed(() => {
  const map = new Map<number, BoundingBoxField[]>()
  for (const field of props.fields) {
    const existing = map.get(field.page) ?? []
    existing.push(field)
    map.set(field.page, existing)
  }
  // Sort by page number
  return Array.from(map.entries()).sort(([a], [b]) => a - b)
})

function confidenceColor(confidence: number): string {
  if (confidence >= 0.8) return "border-green-500/60 bg-green-500/10"
  if (confidence >= 0.5) return "border-amber-500/60 bg-amber-500/10"
  return "border-red-500/60 bg-red-500/10"
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
</script>

<template>
  <div :class="['space-y-4', props.className]">
    <div
      v-for="[page, pageFields] in groupedByPage"
      :key="page"
      class="rounded-lg border border-border bg-background p-4"
    >
      <h4 class="mb-3 text-sm font-medium text-foreground">
        Page {{ page }} &mdash; {{ pageFields.length }} field{{ pageFields.length > 1 ? "s" : "" }}
      </h4>
      <!-- Page representation -->
      <div class="relative mb-3 aspect-[3/4] w-full max-w-sm overflow-hidden rounded-md border border-border bg-muted/30">
        <!-- Render each field as an overlay box -->
        <div
          v-for="field in pageFields"
          :key="field.id"
          :class="[
            'absolute cursor-pointer rounded-sm border transition-opacity hover:opacity-80',
            confidenceColor(field.confidence ?? 0.5),
          ]"
          :style="fieldStyle(field)"
          :title="`${field.label}: ${field.value ?? ''}`"
          @click="emit('field-click', field)"
        >
          <span class="absolute -top-5 left-0 truncate text-[10px] font-medium text-foreground">
            {{ field.label }}
          </span>
        </div>
      </div>
      <!-- Field list -->
      <ul class="space-y-2">
        <li
          v-for="field in pageFields"
          :key="field.id"
          :class="[
            'flex cursor-pointer items-start gap-2 rounded-md p-2 text-sm transition-colors hover:bg-accent',
            confidenceColor(field.confidence ?? 0.5),
          ]"
          @click="emit('field-click', field)"
        >
          <div class="min-w-0 flex-1">
            <span class="font-medium text-foreground">{{ field.label }}</span>
            <span v-if="field.value" class="ml-2 text-muted-foreground">
              {{ field.value }}
            </span>
          </div>
          <span
            v-if="field.confidence !== undefined"
            class="shrink-0 text-xs tabular-nums text-muted-foreground"
          >
            {{ (field.confidence * 100).toFixed(0) }}%
          </span>
        </li>
      </ul>
    </div>
    <p
      v-if="props.fields.length === 0"
      class="py-8 text-center text-sm text-muted-foreground"
    >
      No fields to display
    </p>
  </div>
</template>

<style scoped>
.responsive-fixture-image { max-width: 100%; height: auto; }
</style>
