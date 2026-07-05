<script setup lang="ts">
import { computed } from "vue"

export interface LayoutBlock {
  id: string
  /** Bounding box in absolute image coordinates */
  bbox: [number, number, number, number] // x, y, w, h
  /** Block type: text, title, table, figure, etc. */
  kind: string
  /** Text content */
  text?: string
  /** Confidence 0..1 */
  confidence?: number
  /** Parent/child relationship */
  parentId?: string
}

export interface ParsedOcrOutput {
  width: number
  height: number
  blocks: LayoutBlock[]
}

export interface LayoutBlocksProps {
  /** Image file URL */
  file: string
  /** Parsed OCR output with layout blocks */
  output: ParsedOcrOutput
  className?: string
}

const props = withDefaults(defineProps<LayoutBlocksProps>(), {
  className: "",
})

const emit = defineEmits<{
  "block-click": [block: LayoutBlock]
}>()

const imageStyle = computed(() => ({
  aspectRatio: `${props.output.width} / ${props.output.height}`,
  width: "100%",
  maxWidth: "100%",
  height: "auto",
}))

const kindColors: Record<string, string> = {
  text: "border-blue-500/60 bg-blue-500/10 text-blue-700 dark:text-blue-400",
  title: "border-violet-500/60 bg-violet-500/10 text-violet-700 dark:text-violet-400",
  table: "border-emerald-500/60 bg-emerald-500/10 text-emerald-700 dark:text-emerald-400",
  figure: "border-amber-500/60 bg-amber-500/10 text-amber-700 dark:text-amber-400",
  list: "border-cyan-500/60 bg-cyan-500/10 text-cyan-700 dark:text-cyan-400",
  header: "border-pink-500/60 bg-pink-500/10 text-pink-700 dark:text-pink-400",
  footer: "border-orange-500/60 bg-orange-500/10 text-orange-700 dark:text-orange-400",
}

function blockKindClass(kind: string): string {
  return kindColors[kind] ?? "border-border bg-muted/30 text-muted-foreground"
}

function blockStyle(block: LayoutBlock) {
  const [x, y, w, h] = block.bbox
  return {
    left: `${(x / props.output.width) * 100}%`,
    top: `${(y / props.output.height) * 100}%`,
    width: `${(w / props.output.width) * 100}%`,
    height: `${(h / props.output.height) * 100}%`,
  }
}

function truncatedText(text: string, maxLen = 80): string {
  return text.length > maxLen ? text.slice(0, maxLen) + "…" : text
}
</script>

<template>
  <div :class="['space-y-4', props.className]">
    <!-- Image with overlay blocks -->
    <div class="relative overflow-hidden rounded-lg border border-border">
      <img
        :src="props.file"
        alt="Layout preview"
        class="block w-full responsive-fixture-image"
        :style="imageStyle"
      />
      <!-- Overlay blocks -->
      <div
        v-for="block in props.output.blocks"
        :key="block.id"
        :class="[
          'absolute cursor-pointer rounded-sm border transition-opacity hover:opacity-90',
          blockKindClass(block.kind),
        ]"
        :style="blockStyle(block)"
        :title="block.kind"
        @click="emit('block-click', block)"
      />
    </div>

    <!-- Block list -->
    <div class="grid gap-2 sm:grid-cols-2">
      <div
        v-for="block in props.output.blocks"
        :key="block.id"
        :class="[
          'cursor-pointer rounded-md border p-3 text-sm transition-colors hover:bg-accent',
          blockKindClass(block.kind),
        ]"
        @click="emit('block-click', block)"
      >
        <div class="flex items-center justify-between gap-2">
          <span class="font-medium capitalize">{{ block.kind }}</span>
          <span
            v-if="block.confidence !== undefined"
            class="text-xs tabular-nums text-muted-foreground"
          >
            {{ (block.confidence * 100).toFixed(0) }}%
          </span>
        </div>
        <p v-if="block.text" class="mt-1 text-xs leading-relaxed text-muted-foreground">
          {{ truncatedText(block.text) }}
        </p>
      </div>
    </div>

    <p
      v-if="props.output.blocks.length === 0"
      class="py-8 text-center text-sm text-muted-foreground"
    >
      No layout blocks detected
    </p>
  </div>
</template>

<style scoped>
.responsive-fixture-image { max-width: 100%; height: auto; }
</style>
