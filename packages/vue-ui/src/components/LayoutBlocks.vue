<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed, nextTick, ref, type Component } from "vue"
import {
  AlignBoxBottomCenterIcon,
  Heading01Icon,
  ImageCompositionIcon,
  LeftToRightListBulletIcon,
  ParagraphIcon,
  Table01Icon,
  TextCenterlineCenterTopIcon,
  TextNumberSignIcon,
} from "@hugeicons/core-free-icons"
import { HugeIcon } from "./HugeIcon"
import { useOfficeUiMessages } from "../i18n"
import type {
  LayoutBlock,
  LayoutBlockType,
  LayoutBlocksProps,
  ParsedOcrOutput,
} from "../types"

const props = withDefaults(defineProps<LayoutBlocksProps>(), {
  file: undefined,
  output: undefined,
  blocks: undefined,
  pageImage: undefined,
  activeBlockId: undefined,
  view: "split",
  heading: undefined,
  emptyMessage: undefined,
  renderMarkdown: false,
  loading: false,
  error: undefined,
  ariaLabel: undefined,
  className: "",
})

const emit = defineEmits<{
  "block-click": [block: LayoutBlock]
  "block-focus": [block: LayoutBlock]
  "update:activeBlockId": [id: string]
}>()

const m = useOfficeUiMessages()
const localActiveId = ref<string | null>(props.activeBlockId ?? null)
const cardRefs = ref<Array<HTMLButtonElement | null>>([])

interface BlockStyle {
  label: string
  icon: Component
  tone: LayoutBlockType
}

const BLOCK_STYLES: Record<LayoutBlockType, BlockStyle> = {
  heading: { label: "Heading", icon: HugeIcon, tone: "heading" },
  paragraph: { label: "Paragraph", icon: HugeIcon, tone: "paragraph" },
  list: { label: "List", icon: HugeIcon, tone: "list" },
  table: { label: "Table", icon: HugeIcon, tone: "table" },
  figure: { label: "Figure", icon: HugeIcon, tone: "figure" },
  header: { label: "Header", icon: HugeIcon, tone: "header" },
  footer: { label: "Footer", icon: HugeIcon, tone: "footer" },
  page_number: { label: "Page number", icon: HugeIcon, tone: "page_number" },
}

const BLOCK_ICONS: Record<LayoutBlockType, typeof Heading01Icon> = {
  heading: Heading01Icon,
  paragraph: ParagraphIcon,
  list: LeftToRightListBulletIcon,
  table: Table01Icon,
  figure: ImageCompositionIcon,
  header: TextCenterlineCenterTopIcon,
  footer: AlignBoxBottomCenterIcon,
  page_number: TextNumberSignIcon,
}

const KIND_ALIASES: Record<string, LayoutBlockType> = {
  title: "heading",
  heading: "heading",
  header: "header",
  text: "paragraph",
  paragraph: "paragraph",
  list: "list",
  table: "table",
  figure: "figure",
  image: "figure",
  footer: "footer",
  page_number: "page_number",
  pageNumber: "page_number",
}

const resolvedOutput = computed<ParsedOcrOutput | null>(() => {
  if (props.blocks) {
    const firstBbox = props.blocks[0]?.bbox
    return {
      width: firstBbox ? Math.max(1, firstBbox[0] + firstBbox[2]) : 1000,
      height: firstBbox ? Math.max(1, firstBbox[1] + firstBbox[3]) : 1330,
      blocks: props.blocks,
    }
  }
  return props.output ?? null
})

const blocks = computed<LayoutBlock[]>(() => resolvedOutput.value?.blocks ?? [])

const activeId = computed(() => {
  if (props.activeBlockId !== undefined) {
    return blocks.value.some((block) => block.id === props.activeBlockId)
      ? props.activeBlockId
      : blocks.value[0]?.id
  }
  if (localActiveId.value && blocks.value.some((block) => block.id === localActiveId.value)) {
    return localActiveId.value
  }
  return blocks.value[0]?.id
})

const pageImage = computed(() => props.pageImage ?? props.file ?? "")

const imageStyle = computed(() => {
  const output = resolvedOutput.value
  if (!output) return { aspectRatio: "1000 / 1330" }
  return { aspectRatio: `${output.width} / ${output.height}` }
})

const panelEmptyMessage = computed(
  () => props.emptyMessage ?? m.value.vueUi.layoutBlocks.noBlocks,
)

const splitLayout = computed(() => props.view === "split")

function blockType(block: LayoutBlock): LayoutBlockType {
  if (block.type) return block.type
  return KIND_ALIASES[block.kind] ?? (block.kind as LayoutBlockType)
}

function blockIcon(block: LayoutBlock) {
  return BLOCK_ICONS[blockType(block)] ?? ParagraphIcon
}

function blockLabel(block: LayoutBlock): string {
  return block.label ?? BLOCK_STYLES[blockType(block)]?.label ?? block.kind
}

function confidenceLabel(block: LayoutBlock): string | null {
  if (block.confidence === undefined) return null
  return `${Math.round(block.confidence * 100)}%`
}

function blockStyle(block: LayoutBlock) {
  const output = resolvedOutput.value
  const [x, y, width, height] = block.bbox
  if (!output) return undefined
  return {
    left: `${(x / output.width) * 100}%`,
    top: `${(y / output.height) * 100}%`,
    width: `${(width / output.width) * 100}%`,
    height: `${(height / output.height) * 100}%`,
  }
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/gu, "&amp;")
    .replace(/</gu, "&lt;")
    .replace(/>/gu, "&gt;")
    .replace(/"/gu, "&quot;")
    .replace(/'/gu, "&#39;")
}

/**
 * Minimal, escape-first markdown subset for OCR `text` fields: bold, italic,
 * inline code, headings, list items and http(s) links. Everything is escaped
 * before formatting, so host-supplied OCR text can never inject markup.
 */
function renderMarkdown(value: string): string {
  const lines = escapeHtml(value).split(/\r?\n/)
  return lines
    .map((line) => {
      const heading = /^(#{1,6})\s+(.*)$/u.exec(line)
      if (heading) {
        return `<strong class="layout-block-markdown__heading">${heading[2]}</strong>`
      }
      const listItem = /^[-*]\s+(.*)$/u.exec(line)
      const source = listItem ? `• ${listItem[1]}` : line
      return source
        .replace(/`([^`]+)`/gu, '<code class="layout-block-markdown__code">$1</code>')
        .replace(/\*\*([^*]+)\*\*/gu, "<strong>$1</strong>")
        .replace(/\*([^*]+)\*/gu, "<em>$1</em>")
        .replace(
          /\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/gu,
          '<a class="layout-block-markdown__link" href="$2" rel="noreferrer noopener" target="_blank">$1</a>',
        )
    })
    .join("<br />")
}

function focusBlock(block: LayoutBlock) {
  localActiveId.value = block.id
  emit("update:activeBlockId", block.id)
  emit("block-focus", block)
}

function selectBlock(block: LayoutBlock) {
  focusBlock(block)
  emit("block-click", block)
}

async function moveFocus(block: LayoutBlock, direction: 1 | -1) {
  const index = blocks.value.findIndex((entry) => entry.id === block.id)
  const next = blocks.value[index + direction]
  if (!next) return
  focusBlock(next)
  await nextTick()
  const target = cardRefs.value[index + direction]
  target?.focus?.()
}

function onCardKeydown(event: KeyboardEvent, block: LayoutBlock) {
  if (event.key === "ArrowDown") {
    event.preventDefault()
    void moveFocus(block, 1)
  } else if (event.key === "ArrowUp") {
    event.preventDefault()
    void moveFocus(block, -1)
  } else if (event.key === "Home") {
    event.preventDefault()
    const first = blocks.value[0]
    if (first) {
      focusBlock(first)
      void nextTick(() => cardRefs.value[0]?.focus?.())
    }
  } else if (event.key === "End") {
    event.preventDefault()
    const last = blocks.value[blocks.value.length - 1]
    if (last) {
      focusBlock(last)
      void nextTick(() => cardRefs.value[blocks.value.length - 1]?.focus?.())
    }
  }
}

function setCardRef(element: unknown, index: number) {
  cardRefs.value[index] = (element as HTMLButtonElement | null) ?? null
}
</script>

<template>
  <div
    :class="[
      'layout-blocks',
      'office-ui',
      splitLayout ? 'layout-blocks--split' : 'layout-blocks--panel',
      props.className,
    ]"
    data-slot="layout-blocks"
  >
    <div v-if="splitLayout" class="layout-blocks__viewer" data-slot="layout-blocks-viewer">
      <slot name="viewer">
        <div v-if="props.error" class="layout-blocks__state layout-blocks__state--error" role="alert">
          {{ props.error }}
        </div>
        <div
          v-else-if="props.loading"
          class="layout-blocks__page layout-blocks__page--loading"
          role="status"
          :aria-label="m.vueUi.layoutBlocks.loading"
        >
          <div class="layout-blocks__skeleton" />
        </div>
        <div v-else class="layout-blocks__page" :style="imageStyle">
          <img
            v-if="pageImage"
            :src="pageImage"
            :alt="m.vueUi.layoutBlocks.imageAlt"
            class="layout-blocks__image"
          />
          <button
            v-for="block in blocks"
            :key="block.id"
            type="button"
            :class="[
              'layout-blocks__overlay',
              `layout-blocks__overlay--${blockType(block)}`,
              { 'is-active': activeId === block.id },
            ]"
            :style="blockStyle(block)"
            :title="blockLabel(block)"
            :aria-pressed="activeId === block.id"
            data-slot="layout-block-overlay"
            @click="selectBlock(block)"
          />
        </div>
      </slot>
    </div>

    <aside
      class="layout-blocks__panel"
      data-slot="layout-blocks-panel"
      :aria-label="props.ariaLabel"
    >
      <div v-if="props.heading" class="layout-blocks__heading" data-slot="layout-blocks-heading">
        <span>{{ props.heading }}</span>
        <span class="layout-blocks__count">{{ blocks.length }}</span>
      </div>
      <div class="layout-blocks__scroll" data-slot="layout-blocks-scroll">
        <div v-if="props.error" class="layout-blocks__state-wrap">
          <div class="layout-blocks__state layout-blocks__state--error" role="alert">
            {{ props.error }}
          </div>
        </div>
        <div v-else-if="props.loading" class="layout-blocks__state-wrap">
          <div
            class="layout-blocks__state"
            role="status"
            data-slot="layout-blocks-loading"
          >
            {{ m.vueUi.layoutBlocks.loading }}
          </div>
        </div>
        <div v-else-if="blocks.length" class="layout-blocks__list" data-slot="layout-blocks-list">
          <button
            v-for="(block, index) in blocks"
            :key="block.id"
            :ref="(element) => setCardRef(element, index)"
            type="button"
            :class="[
              'layout-blocks__card',
              `layout-blocks__card--${blockType(block)}`,
              { 'is-active': activeId === block.id },
            ]"
            :aria-pressed="activeId === block.id"
            :data-layout-block-id="block.id"
            :data-block-type="blockType(block)"
            data-slot="layout-block-card"
            @mouseenter="focusBlock(block)"
            @focus="focusBlock(block)"
            @click="selectBlock(block)"
            @keydown="onCardKeydown($event, block)"
          >
            <div class="layout-blocks__card-inner">
              <div class="layout-blocks__card-header">
                <div class="layout-blocks__card-meta">
                  <div class="layout-blocks__badge">
                    <HugeIcon :icon="blockIcon(block)" class="layout-blocks__badge-icon" />
                    {{ blockLabel(block) }}
                  </div>
                  <div v-if="confidenceLabel(block)" class="layout-blocks__confidence">
                    {{ confidenceLabel(block) }}
                  </div>
                </div>
                <div v-if="block.page !== undefined" class="layout-blocks__page-chip">
                  p. {{ block.page }}
                </div>
              </div>
              <div class="layout-blocks__text">
                <p
                  v-if="props.renderMarkdown && block.text"
                  class="layout-blocks__paragraph layout-blocks__paragraph--markdown"
                  data-slot="layout-block-markdown"
                  v-html="renderMarkdown(block.text)"
                />
                <p v-else class="layout-blocks__paragraph">
                  {{ block.text ?? "" }}
                </p>
              </div>
            </div>
          </button>
        </div>
        <div v-else class="layout-blocks__state-wrap">
          <div
            class="layout-blocks__state"
            role="status"
            data-slot="layout-blocks-empty"
          >
            {{ panelEmptyMessage }}
          </div>
        </div>
      </div>
    </aside>
  </div>
</template>

<style scoped>
.layout-blocks {
  min-width: 0;
  color: var(--office-foreground);
}
.layout-blocks,
.layout-blocks * {
  box-sizing: border-box;
}
.layout-blocks--split {
  display: flex;
  min-height: 0;
  align-items: stretch;
  gap: var(--office-ui-space-3);
}
.layout-blocks--split .layout-blocks__viewer {
  min-width: 0;
  flex: 1 1 auto;
}
.layout-blocks__page {
  position: relative;
  width: 100%;
  overflow: hidden;
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-muted);
}
.layout-blocks__page--loading {
  aspect-ratio: 1000 / 1330;
}
.layout-blocks__image {
  display: block;
  width: 100%;
  height: 100%;
  object-fit: contain;
  opacity: 0.92;
}
.layout-blocks__skeleton {
  position: absolute;
  inset: 0;
  background: var(--office-muted);
  animation: layout-blocks-pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite;
}
.layout-blocks__overlay {
  position: absolute;
  z-index: 10;
  margin: 0;
  padding: 0;
  border: 1px solid color-mix(in oklab, var(--office-muted-foreground) 35%, transparent);
  border-radius: 2px;
  background: color-mix(in oklab, var(--office-muted-foreground) 5%, transparent);
  cursor: pointer;
  transition:
    background-color 120ms ease-out,
    border-color 120ms ease-out;
}
.layout-blocks__overlay:hover {
  filter: brightness(1.08);
}
.layout-blocks__overlay:focus-visible {
  outline: 2px solid var(--office-ring);
  outline-offset: 2px;
}
.layout-blocks__overlay--heading {
  border-color: color-mix(in oklab, oklch(60.6% .25 292.717) 70%, transparent);
  background: color-mix(in oklab, oklch(60.6% .25 292.717) 10%, transparent);
}
.layout-blocks__overlay--paragraph {
  border-color: color-mix(in oklab, oklch(62.3% .214 259.815) 70%, transparent);
  background: color-mix(in oklab, oklch(62.3% .214 259.815) 10%, transparent);
}
.layout-blocks__overlay--list {
  border-color: color-mix(in oklab, oklch(69.6% .17 162.48) 70%, transparent);
  background: color-mix(in oklab, oklch(69.6% .17 162.48) 10%, transparent);
}
.layout-blocks__overlay--table {
  border-color: color-mix(in oklab, oklch(76.9% .188 70.08) 70%, transparent);
  background: color-mix(in oklab, oklch(76.9% .188 70.08) 10%, transparent);
}
.layout-blocks__overlay--figure {
  border-color: color-mix(in oklab, oklch(64.5% .246 16.439) 70%, transparent);
  background: color-mix(in oklab, oklch(64.5% .246 16.439) 10%, transparent);
}
.layout-blocks__overlay--header {
  border-color: color-mix(in oklab, oklch(71.5% .143 215.221) 70%, transparent);
  background: color-mix(in oklab, oklch(71.5% .143 215.221) 10%, transparent);
}
.layout-blocks__overlay--footer {
  border-color: color-mix(in oklab, oklch(55.4% .046 257.417) 70%, transparent);
  background: color-mix(in oklab, oklch(55.4% .046 257.417) 10%, transparent);
}
.layout-blocks__overlay--page_number {
  border-color: color-mix(in oklab, oklch(55.2% .016 285.938) 70%, transparent);
  background: color-mix(in oklab, oklch(55.2% .016 285.938) 10%, transparent);
}
.layout-blocks__overlay.is-active {
  z-index: 20;
  border-width: 2px;
  box-shadow: 0 4px 16px color-mix(in oklab, oklch(62.3% .214 259.815) 12%, transparent);
}
.layout-blocks__panel {
  display: flex;
  height: 420px;
  min-height: 0;
  flex-direction: column;
  background: var(--office-background);
}
.layout-blocks--split .layout-blocks__panel {
  height: auto;
  min-height: 420px;
  flex: 0 0 42%;
  max-width: 520px;
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
}
.layout-blocks__heading {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: var(--office-ui-space-3) var(--office-ui-space-4);
  border-bottom: var(--office-border-width) solid var(--office-border);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  font-weight: 500;
}
.layout-blocks__count {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
.layout-blocks__scroll {
  min-height: 0;
  flex: 1 1 auto;
  overflow-y: auto;
  outline: none;
}
.layout-blocks__list {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-2);
  padding: var(--office-ui-space-3);
}
.layout-blocks__card {
  display: block;
  margin: 0;
  width: 100%;
  cursor: pointer;
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-background);
  padding: var(--office-ui-space-3);
  color: inherit;
  text-align: left;
  outline: none;
}
.layout-blocks__card:hover {
  background: color-mix(in oklab, var(--office-muted) 40%, transparent);
}
.layout-blocks__card:focus-visible {
  box-shadow: 0 0 0 2px var(--office-ring);
}
.layout-blocks__card.is-active {
  box-shadow: 0 0 0 1px rgb(0 0 0 / 4%);
}
.layout-blocks__card.is-active:focus-visible {
  box-shadow:
    0 0 0 2px var(--office-ring),
    0 0 0 1px rgb(0 0 0 / 4%);
}
.layout-blocks__card-inner {
  min-width: 0;
}
.layout-blocks__card-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--office-ui-space-2);
}
.layout-blocks__card-meta {
  display: flex;
  min-width: 0;
  align-items: center;
  gap: var(--office-ui-space-2);
}
.layout-blocks__badge {
  display: inline-flex;
  flex-shrink: 0;
  align-items: center;
  gap: var(--office-ui-space-1);
  padding: var(--office-ui-space-0-5) var(--office-ui-space-2);
  border-radius: calc(infinity * 1px);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
  font-weight: 500;
  background: var(--office-muted);
  color: var(--office-muted-foreground);
}
.layout-blocks__badge-icon {
  width: 14px;
  height: 14px;
  flex-shrink: 0;
}
.layout-blocks__card--heading .layout-blocks__badge {
  background: oklch(96.9% .016 293.756);
  color: oklch(54.1% .281 293.009);
}
.layout-blocks__card--paragraph .layout-blocks__badge {
  background: oklch(97% .014 254.604);
  color: oklch(54.6% .245 262.881);
}
.layout-blocks__card--list .layout-blocks__badge {
  background: oklch(97.9% .021 166.113);
  color: oklch(59.6% .145 163.225);
}
.layout-blocks__card--table .layout-blocks__badge {
  background: oklch(98.7% .022 95.277);
  color: oklch(66.6% .179 58.318);
}
.layout-blocks__card--figure .layout-blocks__badge {
  background: oklch(96.9% .015 12.422);
  color: oklch(58.6% .253 17.585);
}
.layout-blocks__card--header .layout-blocks__badge {
  background: oklch(98.4% .019 200.873);
  color: oklch(60.9% .126 221.723);
}
.layout-blocks__card--footer .layout-blocks__badge {
  background: oklch(98.4% .003 247.858);
  color: oklch(44.6% .043 257.281);
}
.layout-blocks__card--page_number .layout-blocks__badge {
  background: oklch(98.5% 0 0);
  color: oklch(44.2% .017 285.786);
}
.layout-blocks__card--heading.is-active {
  border-color: color-mix(in oklab, oklch(60.6% .25 292.717) 60%, transparent);
  background: color-mix(in oklab, oklch(60.6% .25 292.717) 5%, transparent);
}
.layout-blocks__card--paragraph.is-active {
  border-color: color-mix(in oklab, oklch(62.3% .214 259.815) 60%, transparent);
  background: color-mix(in oklab, oklch(62.3% .214 259.815) 5%, transparent);
}
.layout-blocks__card--list.is-active {
  border-color: color-mix(in oklab, oklch(69.6% .17 162.48) 60%, transparent);
  background: color-mix(in oklab, oklch(69.6% .17 162.48) 5%, transparent);
}
.layout-blocks__card--table.is-active {
  border-color: color-mix(in oklab, oklch(76.9% .188 70.08) 60%, transparent);
  background: color-mix(in oklab, oklch(76.9% .188 70.08) 5%, transparent);
}
.layout-blocks__card--figure.is-active {
  border-color: color-mix(in oklab, oklch(64.5% .246 16.439) 60%, transparent);
  background: color-mix(in oklab, oklch(64.5% .246 16.439) 5%, transparent);
}
.layout-blocks__card--header.is-active {
  border-color: color-mix(in oklab, oklch(71.5% .143 215.221) 60%, transparent);
  background: color-mix(in oklab, oklch(71.5% .143 215.221) 5%, transparent);
}
.layout-blocks__card--footer.is-active {
  border-color: color-mix(in oklab, oklch(55.4% .046 257.417) 60%, transparent);
  background: color-mix(in oklab, oklch(55.4% .046 257.417) 5%, transparent);
}
.layout-blocks__card--page_number.is-active {
  border-color: color-mix(in oklab, oklch(55.2% .016 285.938) 60%, transparent);
  background: color-mix(in oklab, oklch(55.2% .016 285.938) 5%, transparent);
}
.layout-blocks__card:hover {
  background: color-mix(in oklab, var(--office-muted) 40%, transparent);
}
.layout-blocks__card.is-active:hover {
  background: color-mix(in oklab, var(--office-muted) 40%, transparent);
}
.layout-blocks__confidence {
  overflow: hidden;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
  text-overflow: ellipsis;
  white-space: nowrap;
}
.layout-blocks__page-chip {
  flex-shrink: 0;
  border-radius: calc(infinity * 1px);
  background: var(--office-muted);
  padding: var(--office-ui-space-0-5) var(--office-ui-space-2);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
.layout-blocks__text {
  margin-top: var(--office-ui-space-2);
  color: color-mix(in oklab, var(--office-foreground) 90%, transparent);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
}
.layout-blocks__paragraph {
  display: -webkit-box;
  margin: 0;
  overflow: hidden;
  -webkit-box-orient: vertical;
  -webkit-line-clamp: 3;
  font-size: 13px;
  line-height: var(--office-leading-sm);
  overflow-wrap: anywhere;
  white-space: pre-line;
}
.layout-blocks__paragraph--markdown {
  display: block;
  white-space: normal;
}
.layout-blocks__paragraph--markdown :deep(.layout-block-markdown__heading) {
  font-weight: 600;
}
.layout-blocks__paragraph--markdown :deep(.layout-block-markdown__code) {
  border-radius: var(--office-radius-sm);
  background: var(--office-muted);
  padding: 0 3px;
  font-family: var(--office-font-mono);
  font-size: 12px;
}
.layout-blocks__paragraph--markdown :deep(.layout-block-markdown__link) {
  color: inherit;
  text-decoration: underline;
}
.layout-blocks__state-wrap {
  padding: var(--office-ui-space-3);
}
.layout-blocks__state {
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: color-mix(in oklab, var(--office-muted) 30%, transparent);
  padding: var(--office-ui-space-3);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
}
.layout-blocks__state--error {
  border-color: color-mix(in oklab, var(--office-destructive) 45%, transparent);
  background: color-mix(in oklab, var(--office-destructive) 8%, transparent);
  color: var(--office-destructive);
}

[data-office-theme="dark"] .layout-blocks__card--heading .layout-blocks__badge {
  background: color-mix(in oklab, oklch(81.1% .111 293.571) 10%, transparent);
  color: oklch(81.1% .111 293.571);
}
[data-office-theme="dark"] .layout-blocks__card--paragraph .layout-blocks__badge {
  background: color-mix(in oklab, oklch(80.9% .105 251.813) 10%, transparent);
  color: oklch(80.9% .105 251.813);
}
[data-office-theme="dark"] .layout-blocks__card--list .layout-blocks__badge {
  background: color-mix(in oklab, oklch(84.5% .143 164.978) 10%, transparent);
  color: oklch(84.5% .143 164.978);
}
[data-office-theme="dark"] .layout-blocks__card--table .layout-blocks__badge {
  background: color-mix(in oklab, oklch(87.9% .169 91.605) 10%, transparent);
  color: oklch(87.9% .169 91.605);
}
[data-office-theme="dark"] .layout-blocks__card--figure .layout-blocks__badge {
  background: color-mix(in oklab, oklch(81% .117 11.638) 10%, transparent);
  color: oklch(81% .117 11.638);
}
[data-office-theme="dark"] .layout-blocks__card--header .layout-blocks__badge {
  background: color-mix(in oklab, oklch(86.5% .127 207.078) 10%, transparent);
  color: oklch(86.5% .127 207.078);
}
[data-office-theme="dark"] .layout-blocks__card--footer .layout-blocks__badge {
  background: color-mix(in oklab, oklch(86.9% .022 252.894) 10%, transparent);
  color: oklch(86.9% .022 252.894);
}
[data-office-theme="dark"] .layout-blocks__card--page_number .layout-blocks__badge {
  background: color-mix(in oklab, oklch(87.1% .006 286.286) 10%, transparent);
  color: oklch(87.1% .006 286.286);
}
@media (prefers-color-scheme: dark) {
  [data-office-theme="system"] .layout-blocks__card--heading .layout-blocks__badge {
    background: color-mix(in oklab, oklch(81.1% .111 293.571) 10%, transparent);
    color: oklch(81.1% .111 293.571);
  }
  [data-office-theme="system"] .layout-blocks__card--paragraph .layout-blocks__badge {
    background: color-mix(in oklab, oklch(80.9% .105 251.813) 10%, transparent);
    color: oklch(80.9% .105 251.813);
  }
  [data-office-theme="system"] .layout-blocks__card--list .layout-blocks__badge {
    background: color-mix(in oklab, oklch(84.5% .143 164.978) 10%, transparent);
    color: oklch(84.5% .143 164.978);
  }
  [data-office-theme="system"] .layout-blocks__card--table .layout-blocks__badge {
    background: color-mix(in oklab, oklch(87.9% .169 91.605) 10%, transparent);
    color: oklch(87.9% .169 91.605);
  }
  [data-office-theme="system"] .layout-blocks__card--figure .layout-blocks__badge {
    background: color-mix(in oklab, oklch(81% .117 11.638) 10%, transparent);
    color: oklch(81% .117 11.638);
  }
  [data-office-theme="system"] .layout-blocks__card--header .layout-blocks__badge {
    background: color-mix(in oklab, oklch(86.5% .127 207.078) 10%, transparent);
    color: oklch(86.5% .127 207.078);
  }
  [data-office-theme="system"] .layout-blocks__card--footer .layout-blocks__badge {
    background: color-mix(in oklab, oklch(86.9% .022 252.894) 10%, transparent);
    color: oklch(86.9% .022 252.894);
  }
  [data-office-theme="system"] .layout-blocks__card--page_number .layout-blocks__badge {
    background: color-mix(in oklab, oklch(87.1% .006 286.286) 10%, transparent);
    color: oklch(87.1% .006 286.286);
  }
}
@keyframes layout-blocks-pulse {
  50% {
    opacity: 0.5;
  }
}
@media (prefers-reduced-motion: reduce) {
  .layout-blocks__skeleton {
    animation: none;
  }
}
@media (max-width: 767px) {
  .layout-blocks--split {
    flex-direction: column;
  }
  .layout-blocks--split .layout-blocks__panel {
    flex: 1 1 auto;
    max-width: none;
    width: 100%;
  }
}
</style>
