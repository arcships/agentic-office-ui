<template>
  <div
    v-if="positioned.length > 0"
    class="docx-gutter-anchor-layer"
    data-docx-gutter-anchor-layer="true"
  >
    <template v-for="entry of positioned" :key="entry.annotation.id">
      <div
        class="docx-gutter-anchor-leader"
        :style="leaderStyle(entry)"
        :data-docx-gutter-leader-id="entry.annotation.id"
        aria-hidden="true"
      />
      <button
        v-if="entry.annotation.comment"
        type="button"
        class="docx-gutter-anchor-marker"
        :class="{
          'docx-gutter-anchor-marker--active':
            entry.annotation.id === activeAnnotationId,
          'docx-gutter-anchor-marker--resolved': entry.annotation.comment.resolved,
        }"
        :style="markerStyle(entry)"
        :data-docx-comment-marker="entry.annotation.ordinal"
        :data-docx-gutter-marker-id="entry.annotation.id"
        :title="markerTitle(entry)"
        :aria-label="markerTitle(entry)"
        @click.stop="emit('activate', entry.annotation.id)"
      >
        {{ entry.annotation.ordinal }}
      </button>
      <div
        v-else
        class="docx-gutter-anchor-tick"
        :class="{ 'docx-gutter-anchor-tick--active': entry.annotation.id === activeAnnotationId }"
        :style="tickStyle(entry)"
        :data-docx-gutter-marker-id="entry.annotation.id"
        aria-hidden="true"
      />
    </template>
  </div>
</template>

<script setup lang="ts">
import {
  commentAccentColor,
  trackedChangeAccentColor,
  type DocxDocumentTheme,
  type PositionedGutterAnnotation,
} from "@arcships/docx-core"
import {
  annotationGutterLeaderStyle,
  annotationGutterMarkerStyle,
} from "./gutter-metrics"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { useDocxMessages } from "../i18n"

// ── Props ──────────────────────────────────────────────────────────
// The marker layer lives inside the scaled page, so its geometry is page space
// (unscaled CSS px); each chip counter-scales itself to keep a constant screen
// size. It is the visual link between a highlighted range and its card: the
// number on the chip is the card's index.
const props = withDefaults(
  defineProps<{
    positioned: readonly PositionedGutterAnnotation[]
    zoomFactor: number
    pageWidthPx: number
    /** Right edge of the text column; markers line up in the right margin. */
    contentRightPx: number
    documentTheme?: DocxDocumentTheme
    activeAnnotationId?: string
  }>(),
  {
    zoomFactor: 1,
    pageWidthPx: 0,
    contentRightPx: 0,
    documentTheme: "light" as DocxDocumentTheme,
  }
)

const emit = defineEmits<{
  activate: [annotationId: string]
}>()

function anchorStartXPx(entry: PositionedGutterAnnotation): number {
  // Keep the chip in the right margin, but never to the left of the anchor
  // (cells and indented paragraphs can start past the default text column).
  return Math.max(props.contentRightPx, Math.min(props.pageWidthPx, entry.anchorX))
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

function markerStyle(entry: PositionedGutterAnnotation): Record<string, string> {
  return {
    ...annotationGutterMarkerStyle(
      anchorStartXPx(entry),
      entry.anchorY,
      props.zoomFactor
    ),
    backgroundColor: accentColor(entry),
  }
}

function tickStyle(entry: PositionedGutterAnnotation): Record<string, string> {
  return {
    ...annotationGutterMarkerStyle(
      anchorStartXPx(entry),
      entry.anchorY,
      props.zoomFactor,
      10
    ),
    backgroundColor: accentColor(entry),
  }
}

function leaderStyle(entry: PositionedGutterAnnotation): Record<string, string> {
  return {
    ...annotationGutterLeaderStyle(
      anchorStartXPx(entry),
      entry.anchorY,
      props.pageWidthPx,
      props.zoomFactor
    ),
    backgroundColor: accentColor(entry),
  }
}

const m = useDocxMessages()

function markerTitle(entry: PositionedGutterAnnotation): string {
  const ordinal = entry.annotation.ordinal ?? 0
  const author = entry.annotation.comment?.author?.trim()
  return author
    ? formatOfficeMessage(m.value.docx.gutter.commentMarkerWithAuthor, { ordinal, author })
    : formatOfficeMessage(m.value.docx.gutter.commentMarker, { ordinal })
}
</script>

<style scoped>
.docx-gutter-anchor-layer {
  inset: 0;
  pointer-events: none;
  position: absolute;
  z-index: 7;
}
.docx-gutter-anchor-leader {
  opacity: 0.55;
  pointer-events: none;
  position: absolute;
}
.docx-gutter-anchor-marker,
.docx-gutter-anchor-tick {
  position: absolute;
}
.docx-gutter-anchor-marker {
  border: none;
  border-radius: 999px;
  box-sizing: border-box;
  color: #ffffff;
  cursor: pointer;
  font-family: inherit;
  font-size: 9px;
  font-weight: 600;
  height: 14px;
  line-height: 1;
  min-width: 14px;
  padding: 0 3px;
  pointer-events: auto;
}
.docx-gutter-anchor-marker:focus-visible {
  outline: 2px solid #2563eb;
  outline-offset: 1px;
}
.docx-gutter-anchor-marker--resolved {
  opacity: 0.55;
}
.docx-gutter-anchor-marker--active {
  box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.5);
}
.docx-gutter-anchor-tick {
  border-radius: 2px;
  height: 10px;
  opacity: 0.75;
  width: 3px;
}
.docx-gutter-anchor-tick--active {
  box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.4);
}
</style>
