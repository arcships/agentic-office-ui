<script setup lang="ts">
import { computed, ref } from "vue"

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

const selectedBlockId = ref<string | null>(null)

const imageStyle = computed(() => ({
  aspectRatio: `${props.output.width} / ${props.output.height}`,
  width: "100%",
  maxWidth: "100%",
  height: "auto",
}))

const kindClasses: Record<string, string> = {
  text: "kind-text",
  title: "kind-title",
  table: "kind-table",
  figure: "kind-figure",
  list: "kind-list",
  header: "kind-header",
  footer: "kind-footer",
}

function blockKindClass(kind: string): string {
  return kindClasses[kind] ?? "kind-default"
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

function selectBlock(block: LayoutBlock) {
  selectedBlockId.value = block.id
  emit("block-click", block)
}
</script>

<template>
  <div :class="['layout-blocks', props.className]">
    <div class="layout-image-frame">
      <img
        :src="props.file"
        alt="Layout preview"
        class="layout-image"
        :style="imageStyle"
      />
      <button
        v-for="block in props.output.blocks"
        :key="block.id"
        type="button"
        :class="[
          'layout-block-box',
          blockKindClass(block.kind),
          { 'is-selected': selectedBlockId === block.id },
        ]"
        :style="blockStyle(block)"
        :title="block.kind"
        :aria-pressed="selectedBlockId === block.id"
        @click="selectBlock(block)"
      />
    </div>

    <div class="layout-block-list">
      <button
        v-for="block in props.output.blocks"
        :key="block.id"
        type="button"
        :class="[
          'layout-block-card',
          blockKindClass(block.kind),
          { 'is-selected': selectedBlockId === block.id },
        ]"
        :aria-pressed="selectedBlockId === block.id"
        @click="selectBlock(block)"
      >
        <span class="layout-block-header">
          <span class="layout-kind">{{ block.kind }}</span>
          <span v-if="block.confidence !== undefined" class="layout-confidence">
            {{ (block.confidence * 100).toFixed(0) }}%
          </span>
        </span>
        <span v-if="block.text" class="layout-text">
          {{ truncatedText(block.text) }}
        </span>
      </button>
    </div>

    <p v-if="props.output.blocks.length === 0" class="empty-state">
      No layout blocks detected
    </p>
  </div>
</template>

<style scoped>
.layout-blocks { display: flex; min-width: 0; flex-direction: column; gap: 16px; }
.layout-image-frame {
  position: relative;
  overflow: hidden;
  border: 1px solid var(--border, #e5e7eb);
  border-radius: var(--radius, 10px);
  background: var(--muted, #f5f5f5);
}
.layout-image { display: block; width: 100%; max-width: 100%; height: auto; }
.layout-block-box {
  position: absolute;
  cursor: crosshair;
  border-width: 2px;
  border-style: solid;
  border-radius: 0 !important;
  background: transparent;
  padding: 0;
  transition: filter 0.12s ease, outline-color 0.12s ease;
}
.layout-block-box:hover { filter: brightness(1.12); }
.layout-block-box:focus-visible,
.layout-block-card:focus-visible { outline: 2px solid #0ea5e9; outline-offset: 2px; }
.layout-block-box.is-selected {
  box-shadow: none;
  outline: 2px solid #0f172a;
  outline-offset: 1px;
  z-index: 2;
}
.layout-block-list { display: grid; gap: 6px; }
@media (min-width: 640px) { .layout-block-list { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
.layout-block-card {
  display: flex;
  min-width: 0;
  cursor: pointer;
  flex-direction: column;
  gap: 4px;
  border-width: 0 0 0 3px;
  border-style: solid;
  border-radius: 0 !important;
  background: transparent;
  padding: 8px 10px;
  text-align: left;
  transition: background-color 0.12s ease;
}
.layout-block-card:hover { background: rgb(148 163 184 / 0.12); }
.layout-block-card.is-selected { box-shadow: none; background: rgb(14 165 233 / 0.12); }
.layout-block-header { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
.layout-kind { color: currentColor; font-size: 14px; font-weight: 600; text-transform: capitalize; }
.layout-confidence { color: var(--muted-foreground, #737373); font-size: 12px; font-variant-numeric: tabular-nums; }
.layout-text { color: var(--muted-foreground, #737373); font-size: 12px; line-height: 1.45; }
.kind-text { border-color: rgb(59 130 246 / 0.65); background: rgb(59 130 246 / 0.10); color: rgb(29 78 216); }
.kind-title { border-color: rgb(139 92 246 / 0.65); background: rgb(139 92 246 / 0.10); color: rgb(109 40 217); }
.kind-table { border-color: rgb(16 185 129 / 0.65); background: rgb(16 185 129 / 0.10); color: rgb(4 120 87); }
.kind-figure { border-color: rgb(245 158 11 / 0.65); background: rgb(245 158 11 / 0.10); color: rgb(180 83 9); }
.kind-list { border-color: rgb(6 182 212 / 0.65); background: rgb(6 182 212 / 0.10); color: rgb(14 116 144); }
.kind-header { border-color: rgb(236 72 153 / 0.65); background: rgb(236 72 153 / 0.10); color: rgb(190 24 93); }
.kind-footer { border-color: rgb(249 115 22 / 0.65); background: rgb(249 115 22 / 0.10); color: rgb(194 65 12); }
.kind-default { border-color: var(--border, #e5e7eb); background: color-mix(in oklch, var(--muted, #f5f5f5) 35%, transparent); color: var(--muted-foreground, #737373); }
.empty-state { padding: 32px 0; color: var(--muted-foreground, #737373); text-align: center; font-size: 14px; }
</style>
