<template>
  <div v-if="hasContent" class="docx-tracked-change-gutter">
    <!-- Tracked changes section -->
    <div v-if="showChanges && trackedChanges.length > 0" class="docx-gutter-section">
      <div class="docx-gutter-section-title">
        Tracked Changes ({{ trackedChanges.length }})
      </div>
      <div
        v-for="change of trackedChanges"
        :key="change.id"
        class="docx-gutter-card"
        :class="changeCardClass(change)"
      >
        <div class="docx-gutter-card-header">
          <span class="docx-gutter-card-type">{{ changeLabel(change) }}</span>
          <span class="docx-gutter-card-author">{{ change.author }}</span>
        </div>
        <div class="docx-gutter-card-text">
          {{ changeText(change) }}
        </div>
      </div>
    </div>

    <!-- Comments section -->
    <div v-if="showComments && comments.length > 0" class="docx-gutter-section">
      <div class="docx-gutter-section-title">
        Comments ({{ comments.length }})
      </div>
      <div
        v-for="comment of comments"
        :key="comment.id"
        class="docx-gutter-card docx-gutter-card--comment"
      >
        <div class="docx-gutter-card-header">
          <span class="docx-gutter-card-author">{{ comment.author }}</span>
          <span class="docx-gutter-card-date">{{ formatDate(comment.date) }}</span>
        </div>
        <div class="docx-gutter-card-text">
          {{ comment.text }}
        </div>
        <div v-if="comment.resolved" class="docx-gutter-card-resolved">
          ✓ Resolved
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type {
  DocxEditorController,
  DocxTrackedChange,
  DocxComment,
} from "@extend-ai/docx-core"

// ── Props ──────────────────────────────────────────────────────────
const props = defineProps<{
  pageIndex: number
  pageLayout: { pageWidthPx: number; pageHeightPx: number; marginsPx?: { top: number; bottom: number; left: number; right: number } }
  controller: DocxEditorController
}>()

// ── Computed ───────────────────────────────────────────────────────
const showChanges = computed(() => props.controller.showTrackedChanges)
const showComments = computed(() => props.controller.showComments)

const trackedChanges = computed<DocxTrackedChange[]>(() => {
  return props.controller.trackedChanges
})

const comments = computed<DocxComment[]>(() => {
  return props.controller.comments
})

const hasContent = computed(
  () =>
    (showChanges.value && trackedChanges.value.length > 0) ||
    (showComments.value && comments.value.length > 0)
)

// ── Helpers ────────────────────────────────────────────────────────
function changeCardClass(change: DocxTrackedChange): string {
  switch (change.kind) {
    case "insertion":
      return "docx-gutter-card--insertion"
    case "deletion":
      return "docx-gutter-card--deletion"
    case "format-change":
      return "docx-gutter-card--format"
    default:
      return ""
  }
}

function changeLabel(change: DocxTrackedChange): string {
  switch (change.kind) {
    case "insertion":
      return "+ Inserted"
    case "deletion":
      return "− Deleted"
    case "format-change":
      return "~ Format"
    default:
      return change.kind
  }
}

function changeText(change: DocxTrackedChange): string {
  if (change.kind === "insertion") {
    return change.text ?? ""
  }
  if (change.kind === "deletion") {
    return change.text ?? ""
  }
  return change.text ?? ""
}

function formatDate(dateStr?: string): string {
  if (!dateStr) return ""
  try {
    const date = new Date(dateStr)
    return date.toLocaleDateString()
  } catch {
    return dateStr
  }
}
</script>

<style scoped>
.docx-tracked-change-gutter {
  padding: 8px 0;
  border-top: 1px solid #e5e7eb;
  margin-top: 8px;
}
.docx-gutter-section {
  margin-bottom: 12px;
}
.docx-gutter-section-title {
  font-size: 11px;
  font-weight: 600;
  color: #6b7280;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  margin-bottom: 6px;
}
.docx-gutter-card {
  padding: 6px 8px;
  border-radius: 4px;
  margin-bottom: 4px;
  font-size: 12px;
  border-left: 3px solid #d1d5db;
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
.docx-gutter-card--comment {
  background: #fffbeb;
  border-left-color: #f59e0b;
}
.docx-gutter-card-header {
  display: flex;
  justify-content: space-between;
  margin-bottom: 2px;
}
.docx-gutter-card-type {
  font-weight: 600;
  color: #374151;
}
.docx-gutter-card-author {
  color: #6b7280;
  font-size: 11px;
}
.docx-gutter-card-date {
  color: #9ca3af;
  font-size: 10px;
}
.docx-gutter-card-text {
  color: #4b5563;
  white-space: pre-wrap;
  word-break: break-word;
}
.docx-gutter-card-resolved {
  color: #22c55e;
  font-size: 11px;
  margin-top: 4px;
}
</style>
