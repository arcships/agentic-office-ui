<template>
  <div v-if="hasContent" class="docx-tracked-change-gutter">
    <div
      v-for="change of trackedChanges"
      :key="change.id"
      class="docx-tracked-change-item"
      :class="{ 'docx-tracked-insertion': change.kind === 'insertion', 'docx-tracked-deletion': change.kind === 'deletion' }"
    >
      <span class="docx-tracked-change-author">{{ change.author }}</span>
      <span class="docx-tracked-change-text">{{ change.text }}</span>
    </div>
    <div
      v-for="comment of comments"
      :key="comment.id"
      class="docx-comment-item"
    >
      <div class="docx-comment-author">{{ comment.author }}</div>
      <div class="docx-comment-text">{{ comment.text }}</div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type { DocxEditorController, LayoutPage } from "@extend-ai/docx-core"

const props = defineProps<{
  page: LayoutPage
  controller?: DocxEditorController
}>()

const trackedChanges = computed(() => props.controller?.trackedChanges ?? [])
const comments = computed(() => props.controller?.comments ?? [])
const hasContent = computed(
  () =>
    (props.controller?.showTrackedChanges && trackedChanges.value.length > 0) ||
    (props.controller?.showComments && comments.value.length > 0)
)
</script>

<style scoped>
.docx-tracked-change-gutter {
  position: absolute;
  right: -180px;
  top: 0;
  width: 160px;
  font-size: 11px;
  color: #6b7280;
}
.docx-tracked-change-item {
  padding: 4px 8px;
  margin-bottom: 4px;
  border-radius: 4px;
}
.docx-tracked-insertion {
  background: #dcfce7;
  border-left: 3px solid #22c55e;
}
.docx-tracked-deletion {
  background: #fee2e2;
  border-left: 3px solid #ef4444;
  text-decoration: line-through;
}
.docx-tracked-change-author {
  display: block;
  font-weight: 600;
  font-size: 10px;
  color: #9ca3af;
}
.docx-comment-item {
  padding: 4px 8px;
  margin-bottom: 4px;
  background: #fef9c3;
  border-radius: 4px;
}
.docx-comment-author {
  font-weight: 600;
  font-size: 10px;
  color: #9ca3af;
}
</style>
