<template>
  <div
    v-if="positioned.length > 0"
    class="docx-tracked-change-gutter"
    data-docx-thumbnail-exclude="true"
    data-docx-annotation-gutter="true"
    :data-docx-gutter-card-count="positioned.length"
    :data-docx-gutter-zoom="zoomFactor"
    :style="paneStyle"
  >
    <div
      v-for="entry of positioned"
      :key="entry.annotation.id"
      :ref="(element) => registerCardElement(entry.annotation.id, element)"
      class="docx-gutter-card"
      :class="cardClasses(entry)"
      :style="cardStyle(entry)"
      :data-docx-gutter-card-id="entry.annotation.id"
      :data-docx-gutter-card-index="entry.annotation.ordinal"
      :data-docx-gutter-anchor-y="entry.anchorY"
      :data-docx-gutter-card-top="entry.top"
      :data-docx-tracked-change-kind="entry.annotation.trackedChange?.kind"
      :data-docx-gutter-balloon="gutterBalloonFlag(entry)"
      role="button"
      tabindex="0"
      @click="emit('activate', entry.annotation.id)"
      @keydown.enter.prevent="emit('activate', entry.annotation.id)"
      @keydown.space.prevent="emit('activate', entry.annotation.id)"
    >
      <div class="docx-gutter-card-header">
        <span
          class="docx-gutter-card-index"
          :style="{ backgroundColor: accentColor(entry) }"
        >{{ entry.annotation.ordinal }}</span>
        <span class="docx-gutter-card-type">{{ cardTypeLabel(entry) }}</span>
        <span class="docx-gutter-card-author">{{ authorOf(entry) }}</span>
      </div>
      <div class="docx-gutter-card-text">
        {{ cardText(entry) }}
      </div>
      <div
        v-if="canDecide(entry)"
        class="docx-gutter-card-actions"
        @click.stop
        @keydown.enter.stop
        @keydown.space.stop
      >
        <button
          type="button"
          class="docx-gutter-card-action docx-gutter-card-action--accept"
          data-testid="docx-gutter-accept"
          :title="m.docx.gutter.acceptTitle"
          @click.stop="emit('decide', entry.annotation.id, 'accept')"
        >
          ✓ {{ m.docx.gutter.accept }}
        </button>
        <button
          type="button"
          class="docx-gutter-card-action docx-gutter-card-action--reject"
          data-testid="docx-gutter-reject"
          :title="m.docx.gutter.rejectTitle"
          @click.stop="emit('decide', entry.annotation.id, 'reject')"
        >
          ✕ {{ m.docx.gutter.reject }}
        </button>
      </div>
      <div v-if="entry.annotation.comment?.resolved" class="docx-gutter-card-resolved">
        ✓ {{ m.docx.trackedChange.resolved }}
      </div>
      <div v-if="dateOf(entry)" class="docx-gutter-card-date">
        {{ formatDate(dateOf(entry)) }}
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, onMounted, watch } from "vue"
import {
  commentAccentColor,
  trackedChangeAccentColor,
  trackedChangeUsesGutterBalloon,
  type DocxDocumentTheme,
  type DocxTrackedChange,
  type PositionedGutterAnnotation,
} from "@arcships/docx-core"
import {
  annotationGutterCardTopPx,
  annotationGutterPaneStyle,
} from "./gutter-metrics"
import { formatOfficeDate } from "@arcships/office-interaction"
import { useDocxLocale, useDocxMessages } from "../i18n"

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    /** Cards placed against their anchors, in document order. */
    positioned: readonly PositionedGutterAnnotation[]
    /** Viewer zoom factor of the page subtree this pane lives in. */
    zoomFactor: number
    /** Page height in page space, used to size the pane. */
    pageHeightPx: number
    documentTheme?: DocxDocumentTheme
    activeAnnotationId?: string
    /** True when the host can accept/reject revisions on the cards. */
    decidable?: boolean
  }>(),
  {
    zoomFactor: 1,
    pageHeightPx: 0,
    documentTheme: "light" as DocxDocumentTheme,
    decidable: false,
  }
)

const emit = defineEmits<{
  activate: [annotationId: string]
  cardHeight: [annotationId: string, heightPx: number]
  decide: [annotationId: string, decision: "accept" | "reject"]
}>()
const m = useDocxMessages()
const locale = useDocxLocale()

// ── Revision decisions ─────────────────────────────────────────────
// Only tracked-change cards expose decisions; comment cards stay activate-only.
function canDecide(entry: PositionedGutterAnnotation): boolean {
  return props.decidable && Boolean(entry.annotation.trackedChange)
}

// ── Pane geometry ──────────────────────────────────────────────────
// Counter-scaled so the pane keeps Word's fixed annotation-pane metrics at
// every zoom level instead of shrinking/growing with the paper.
const paneStyle = computed(() =>
  annotationGutterPaneStyle(props.zoomFactor, props.pageHeightPx)
)

function cardStyle(entry: PositionedGutterAnnotation): Record<string, string> {
  return {
    top: `${annotationGutterCardTopPx(entry.top, props.zoomFactor)}px`,
  }
}

// ── Card content ───────────────────────────────────────────────────
function changeLabel(change: DocxTrackedChange): string {
  switch (change.kind) {
    case "insertion":
      return m.value.docx.trackedChange.inserted
    case "deletion":
      return m.value.docx.trackedChange.deleted
    case "move-from":
      return m.value.docx.trackedChange.movedFrom
    case "move-to":
      return m.value.docx.trackedChange.movedTo
    case "format-change":
      return m.value.docx.trackedChange.format
    case "paragraph-format-change":
      return m.value.docx.trackedChange.paragraphFormatting
    case "table-format-change":
      return m.value.docx.trackedChange.tableFormatting
    case "row-format-change":
      return m.value.docx.trackedChange.rowFormatting
    case "cell-format-change":
      return m.value.docx.trackedChange.cellFormatting
    default:
      return change.kind
  }
}

function accentColor(entry: PositionedGutterAnnotation): string {
  if (entry.annotation.comment) {
    return commentAccentColor(
      props.documentTheme,
      entry.annotation.comment.commentId
    )
  }

  return entry.annotation.trackedChange
    ? trackedChangeAccentColor(entry.annotation.trackedChange.kind, props.documentTheme)
    : "#9ca3af"
}

function cardClasses(entry: PositionedGutterAnnotation): string[] {
  const classes = ["docx-gutter-card--annotation"]
  if (entry.annotation.comment) {
    classes.push("docx-gutter-card--comment")
    if (entry.annotation.comment.resolved) {
      classes.push("docx-gutter-card--resolved")
    }
  }

  const change = entry.annotation.trackedChange
  if (change) {
    switch (change.kind) {
      case "insertion":
        classes.push("docx-gutter-card--insertion")
        break
      case "deletion":
        classes.push("docx-gutter-card--deletion")
        break
      case "move-from":
      case "move-to":
        classes.push("docx-gutter-card--move")
        break
      case "format-change":
      case "paragraph-format-change":
      case "table-format-change":
      case "row-format-change":
      case "cell-format-change":
        classes.push("docx-gutter-card--format")
        break
      default:
        break
    }
  }

  if (props.activeAnnotationId === entry.annotation.id) {
    classes.push("docx-gutter-card--active")
  }

  return classes
}

function cardTypeLabel(entry: PositionedGutterAnnotation): string {
  if (entry.annotation.comment) {
    return m.value.docx.trackedChange.comment
  }

  const change = entry.annotation.trackedChange
  return change ? changeLabel(change) : m.value.docx.trackedChange.annotation
}

function gutterBalloonFlag(entry: PositionedGutterAnnotation): string | undefined {
  const change = entry.annotation.trackedChange
  return change
    ? trackedChangeUsesGutterBalloon(change)
      ? "true"
      : "false"
    : undefined
}

function authorOf(entry: PositionedGutterAnnotation): string {
  return (
    entry.annotation.comment?.author ??
    entry.annotation.trackedChange?.author ??
    ""
  )
}

function cardText(entry: PositionedGutterAnnotation): string {
  return (
    entry.annotation.comment?.text ??
    entry.annotation.trackedChange?.text ??
    ""
  )
}

function dateOf(entry: PositionedGutterAnnotation): string | undefined {
  return (
    entry.annotation.comment?.date ?? entry.annotation.trackedChange?.date
  )
}

function formatDate(dateStr?: string): string {
  if (!dateStr) return ""
  return formatOfficeDate(dateStr, locale.value) || dateStr
}

// ── Card measurement ───────────────────────────────────────────────
// Cards are counter-scaled, so `offsetHeight` (layout px) equals their screen
// height; the layout pipeline converts it back to page space.
const cardElements = new Map<string, HTMLElement>()

function isMeasurableElement(element: unknown): element is HTMLElement {
  return (
    typeof element === "object" &&
    element !== null &&
    typeof (element as HTMLElement).offsetHeight === "number"
  )
}

function registerCardElement(annotationId: string, element: unknown): void {
  if (isMeasurableElement(element)) {
    cardElements.set(annotationId, element)
    return
  }

  cardElements.delete(annotationId)
}

function measureCards(): void {
  for (const [annotationId, element] of cardElements) {
    emit("cardHeight", annotationId, element.offsetHeight)
  }

}

async function measureCardsAfterRender(): Promise<void> {
  await nextTick()
  measureCards()
}

onMounted(measureCardsAfterRender)
watch(() => props.positioned, measureCardsAfterRender, { flush: "post" })

// Activating an annotation from the page (marker or highlighted text) brings
// its card into view inside the pane.
watch(
  () => props.activeAnnotationId,
  async (annotationId) => {
    if (!annotationId) return
    await nextTick()
    cardElements.get(annotationId)?.scrollIntoView?.({ block: "nearest" })
  }
)
</script>

<style scoped>
.docx-tracked-change-gutter {
  box-sizing: border-box;
  /* Default beside-the-paper placement at 100% zoom. The pane counter-scales, so
     `annotationGutterPaneStyle` overrides `left`/`width`/`height`/`transform`
     per zoom level to keep the visible gap and the card text constant. */
  left: calc(100% + 16px);
  overflow-x: hidden;
  overflow-y: auto;
  padding: 0;
  /* Cards are absolutely positioned in screen px by
     `annotationGutterCardTopPx`; the pane only provides the scrolling frame. */
  position: absolute;
  top: 0;
  width: 240px;
  z-index: 8;
}
.docx-gutter-card {
  background: #f9fafb;
  border-left: 3px solid #d1d5db;
  border-radius: 4px;
  box-sizing: border-box;
  cursor: pointer;
  font-size: 12px;
  left: 0;
  margin-bottom: 4px;
  padding: 6px 8px;
  position: absolute;
  right: 0;
}
.docx-gutter-card:focus-visible {
  outline: 2px solid #2563eb;
  outline-offset: 1px;
}
.docx-gutter-card--active {
  box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.45);
}
.docx-gutter-card--insertion {
  background: #f0fdf4;
  border-left-color: #22c55e;
}
.docx-gutter-card--deletion {
  background: #fef2f2;
  border-left-color: #ef4444;
}
.docx-gutter-card--format {
  background: #eff6ff;
  border-left-color: #3b82f6;
}
.docx-gutter-card--move {
  background: #f0fdf4;
  border-left-color: #70ad47;
}
.docx-gutter-card--comment {
  background: #fffbeb;
  border-left-color: #f59e0b;
}
.docx-gutter-card--resolved {
  opacity: 0.7;
}
.docx-gutter-card-header {
  align-items: baseline;
  display: flex;
  gap: 6px;
  margin-bottom: 2px;
}
.docx-gutter-card-index {
  border-radius: 999px;
  color: #ffffff;
  flex: 0 0 auto;
  font-size: 10px;
  font-weight: 600;
  line-height: 14px;
  min-width: 14px;
  padding: 0 4px;
  text-align: center;
}
.docx-gutter-card-type {
  color: #374151;
  flex: 1 1 auto;
  font-weight: 600;
}
.docx-gutter-card-author {
  color: #6b7280;
  flex: 0 0 auto;
  font-size: 11px;
}
.docx-gutter-card-text {
  color: #4b5563;
  white-space: pre-wrap;
  word-break: break-word;
}
.docx-gutter-card-actions {
  display: flex;
  gap: 4px;
  margin-top: 6px;
}
.docx-gutter-card-action {
  background: #ffffff;
  border: 1px solid #d1d5db;
  border-radius: 3px;
  color: #374151;
  cursor: pointer;
  font-size: 11px;
  line-height: 1.4;
  padding: 1px 6px;
}
.docx-gutter-card-action--accept:hover {
  background: #ecfdf5;
  border-color: #22c55e;
}
.docx-gutter-card-action--reject:hover {
  background: #fef2f2;
  border-color: #ef4444;
}
.docx-gutter-card-action:focus-visible {
  outline: 2px solid #2563eb;
  outline-offset: 1px;
}
.docx-gutter-card-resolved {
  color: #22c55e;
  font-size: 11px;
  margin-top: 4px;
}
.docx-gutter-card-date {
  color: #9ca3af;
  font-size: 10px;
  margin-top: 2px;
}

/*
 * Keep the gutter beside the paper at every viewport width. The page frame is
 * exactly one paper tall and clips the page surface, so flowing the gutter
 * below the paper would overlap the next page; DocxViewerRoot reserves the
 * horizontal space for this gutter instead.
 */
</style>
