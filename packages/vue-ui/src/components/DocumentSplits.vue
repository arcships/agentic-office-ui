<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed, nextTick, ref } from "vue"
import {
  Add01Icon,
  Delete02Icon,
  DragDropVerticalIcon,
} from "@hugeicons/core-free-icons"
import { HugeIcon } from "./HugeIcon"
import FileThumbnail from "./FileThumbnail.vue"
import OfficeButton from "./ui/OfficeButton.vue"
import { useOfficeUiMessages } from "../i18n"
import { formatOfficeMessage } from "@arcships/office-interaction"
import type {
  DocumentSplit,
  DocumentSplitPageId,
  DocumentSplitsProps,
} from "../types"

const props = withDefaults(defineProps<DocumentSplitsProps>(), {
  thumbnailImages: () => ({}),
  withFrameDivider: true,
  hideAddSplit: false,
  disableRemove: false,
  loading: false,
  error: undefined,
  emptyMessage: undefined,
  className: "",
})

const emit = defineEmits<{
  "select-page": [pageNumber: number]
  "update:splits": [splits: DocumentSplit[]]
  "splits-change": [splits: DocumentSplit[]]
  "add-split": [split: DocumentSplit]
  "remove-split": [split: DocumentSplit]
}>()

const m = useOfficeUiMessages()

const THUMBNAIL_WIDTH = 72
const THUMBNAIL_HEIGHT = 92

const draggedPageId = ref<DocumentSplitPageId | null>(null)
const draggedSplitId = ref<string | null>(null)
const dropTargetKey = ref<string | null>(null)
const splitCounter = ref(props.splits.length)

const splitList = computed(() => props.splits)

const emptyCopy = computed(
  () => props.emptyMessage ?? m.value.vueUi.documentSplits.noSplits,
)

function pageNumber(pageId: DocumentSplitPageId | string): number {
  return Number(String(pageId).replace("page-", ""))
}

function pageRangeLabel(group: DocumentSplit): string {
  if (group.pages.length === 0) return m.value.vueUi.documentSplits.noPages
  return formatOfficeMessage(m.value.vueUi.documentSplits.pageRange, {
    pages: group.pages.map(pageNumber).join(", "),
  })
}

function pageTitle(pageId: DocumentSplitPageId): string {
  return formatOfficeMessage(m.value.vueUi.documentSplits.pageTitle, {
    number: pageNumber(pageId),
  })
}

function removeLabel(group: DocumentSplit): string {
  return formatOfficeMessage(m.value.vueUi.documentSplits.remove, {
    title: group.title,
  })
}

function reorderLabel(group: DocumentSplit): string {
  return formatOfficeMessage(m.value.vueUi.documentSplits.reorder, {
    title: group.title,
  })
}

function commit(nextSplits: DocumentSplit[], action: "add" | "remove" | "update" = "update") {
  if (action === "add") {
    const added = nextSplits[nextSplits.length - 1]
    if (added) emit("add-split", added)
  }
  if (action === "remove") {
    const removed = props.splits.find(
      (group) => !nextSplits.some((next) => next.id === group.id),
    )
    if (removed) emit("remove-split", removed)
  }
  emit("update:splits", nextSplits)
  emit("splits-change", nextSplits)
}

function addSplit() {
  splitCounter.value += 1
  const split: DocumentSplit = {
    id: `split-${Date.now()}-${splitCounter.value}`,
    title: formatOfficeMessage(m.value.vueUi.documentSplits.splitFallback, {
      number: splitList.value.length + 1,
    }),
    pages: [],
  }
  commit([...splitList.value, split], "add")
}

function removeSplit(group: DocumentSplit) {
  if (!canRemove()) return
  commit(
    splitList.value.filter((entry) => entry.id !== group.id),
    "remove",
  )
}

function canRemove(): boolean {
  return !props.disableRemove && splitList.value.length > 1
}

function findGroupId(pageId: DocumentSplitPageId | string): string | null {
  if (splitList.value.some((group) => group.id === pageId)) return pageId
  return (
    splitList.value.find((group) => group.pages.includes(pageId as DocumentSplitPageId))
      ?.id ?? null
  )
}

function movePageToGroup(
  pageId: DocumentSplitPageId,
  targetGroupId: string,
  insertIndex: number,
): DocumentSplit[] {
  return splitList.value.map((group) => {
    const pagesWithoutDragged = group.pages.filter((entry) => entry !== pageId)
    if (group.id !== targetGroupId) {
      return { ...group, pages: pagesWithoutDragged }
    }
    const nextPages = [...pagesWithoutDragged]
    nextPages.splice(Math.max(0, Math.min(insertIndex, nextPages.length)), 0, pageId)
    return { ...group, pages: nextPages }
  })
}

function reorderPage(
  pageId: DocumentSplitPageId,
  overPageId: DocumentSplitPageId,
): DocumentSplit[] {
  return splitList.value.map((group) => {
    const from = group.pages.indexOf(pageId)
    const to = group.pages.indexOf(overPageId)
    if (from === -1 || to === -1) return group
    const pages = [...group.pages]
    pages.splice(from, 1)
    pages.splice(to, 0, pageId)
    return { ...group, pages }
  })
}

function splitSignature(splits: DocumentSplit[]): string {
  return splits
    .map((group) => `${group.id}:${group.pages.join(",")}`)
    .join("|")
}

function applySplits(nextSplits: DocumentSplit[]) {
  if (splitSignature(nextSplits) === splitSignature(splitList.value)) return
  commit(nextSplits)
}

function onPageDragStart(event: DragEvent, pageId: DocumentSplitPageId) {
  draggedPageId.value = pageId
  draggedSplitId.value = null
  event.dataTransfer?.setData("text/plain", pageId)
  if (event.dataTransfer) event.dataTransfer.effectAllowed = "move"
}

function onSplitDragStart(event: DragEvent, group: DocumentSplit) {
  draggedSplitId.value = group.id
  draggedPageId.value = null
  event.dataTransfer?.setData("text/plain", group.id)
  if (event.dataTransfer) event.dataTransfer.effectAllowed = "move"
}

function onDragEnd() {
  draggedPageId.value = null
  draggedSplitId.value = null
  dropTargetKey.value = null
}

function onDragOverPage(event: DragEvent, key: string) {
  if (!draggedPageId.value && !draggedSplitId.value) return
  event.preventDefault()
  dropTargetKey.value = key
  if (event.dataTransfer) event.dataTransfer.dropEffect = "move"
}

function onDropOnPage(targetGroupId: string, overPageId: DocumentSplitPageId) {
  dropTargetKey.value = null
  const pageId = draggedPageId.value
  if (!pageId) {
    onDragEnd()
    return
  }
  const sourceGroupId = findGroupId(pageId)
  if (sourceGroupId === targetGroupId) {
    applySplits(reorderPage(pageId, overPageId))
  } else {
    const targetPages =
      splitList.value.find((group) => group.id === targetGroupId)?.pages ?? []
    const overIndex = targetPages.indexOf(overPageId)
    applySplits(
      movePageToGroup(pageId, targetGroupId, overIndex === -1 ? targetPages.length : overIndex),
    )
  }
  onDragEnd()
}

function onDropOnGroup(targetGroupId: string) {
  dropTargetKey.value = null
  const pageId = draggedPageId.value
  if (!pageId) {
    onDragEnd()
    return
  }
  const sourceGroupId = findGroupId(pageId)
  if (sourceGroupId !== targetGroupId) {
    const targetPages =
      splitList.value.find((group) => group.id === targetGroupId)?.pages ?? []
    applySplits(movePageToGroup(pageId, targetGroupId, targetPages.length))
  }
  onDragEnd()
}

function onDropOnSplit(targetGroupId: string) {
  dropTargetKey.value = null
  const activeId = draggedSplitId.value
  if (!activeId) {
    onDragEnd()
    return
  }
  const from = splitList.value.findIndex((group) => group.id === activeId)
  const to = splitList.value.findIndex((group) => group.id === targetGroupId)
  if (from !== -1 && to !== -1 && from !== to) {
    const next = [...splitList.value]
    const [moved] = next.splice(from, 1)
    if (moved) next.splice(to, 0, moved)
    applySplits(next)
  }
  onDragEnd()
}

async function reorderSplit(group: DocumentSplit, direction: 1 | -1) {
  const index = splitList.value.findIndex((entry) => entry.id === group.id)
  const target = index + direction
  if (target < 0 || target >= splitList.value.length) return
  const next = [...splitList.value]
  const [moved] = next.splice(index, 1)
  if (!moved) return
  next.splice(target, 0, moved)
  applySplits(next)
  await nextTick()
}

async function movePageAcrossGroups(
  pageId: DocumentSplitPageId,
  direction: 1 | -1,
) {
  const sourceGroupId = findGroupId(pageId)
  if (!sourceGroupId) return
  const sourceIndex = splitList.value.findIndex((group) => group.id === sourceGroupId)
  const target = splitList.value[sourceIndex + direction]
  if (!target) return
  applySplits(movePageToGroup(pageId, target.id, target.pages.length))
  await nextTick()
}

function reorderPageWithinGroup(pageId: DocumentSplitPageId, direction: 1 | -1) {
  const group = splitList.value.find((entry) => entry.pages.includes(pageId))
  if (!group) return
  const index = group.pages.indexOf(pageId)
  const target = index + direction
  if (target < 0 || target >= group.pages.length) return
  applySplits(reorderPage(pageId, group.pages[target]!))
}

function onPageClick(pageId: DocumentSplitPageId) {
  emit("select-page", pageNumber(pageId))
}

function onPageKeydown(event: KeyboardEvent, pageId: DocumentSplitPageId) {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault()
    onPageClick(pageId)
    return
  }
  if (event.key !== "ArrowLeft" && event.key !== "ArrowRight") return
  event.preventDefault()
  const direction = event.key === "ArrowRight" ? 1 : -1
  if (event.ctrlKey || event.metaKey) {
    void movePageAcrossGroups(pageId, direction)
  } else {
    reorderPageWithinGroup(pageId, direction)
  }
}

function onSplitHandleKeydown(event: KeyboardEvent, group: DocumentSplit) {
  if (event.key !== "ArrowUp" && event.key !== "ArrowDown") return
  event.preventDefault()
  void reorderSplit(group, event.key === "ArrowDown" ? 1 : -1)
}

function isDropTarget(key: string): boolean {
  return dropTargetKey.value === key
}
</script>

<template>
  <aside
    :class="[
      'doc-splits',
      'office-ui',
      { 'doc-splits--framed': props.withFrameDivider },
      props.className,
    ]"
    data-slot="document-splits"
  >
    <div v-if="!props.hideAddSplit" class="doc-splits__toolbar" data-slot="document-splits-toolbar">
      <OfficeButton variant="outline" size="sm" @click="addSplit">
        <HugeIcon :icon="Add01Icon" class="doc-splits__button-icon" />
        {{ m.vueUi.documentSplits.addSplit }}
      </OfficeButton>
    </div>

    <div class="doc-splits__body" data-slot="document-splits-body">
      <div v-if="props.error" class="doc-splits__state-wrap">
        <div class="doc-splits__state doc-splits__state--error" role="alert">
          {{ props.error }}
        </div>
      </div>
      <div v-else-if="props.loading" class="doc-splits__state-wrap">
        <div class="doc-splits__state" role="status" data-slot="document-splits-loading">
          {{ m.vueUi.documentSplits.loading }}
        </div>
      </div>
      <div v-else-if="splitList.length === 0" class="doc-splits__state-wrap">
        <div class="doc-splits__state" role="status" data-slot="document-splits-empty">
          {{ emptyCopy }}
        </div>
      </div>
      <div v-else class="doc-splits__list" data-slot="document-splits-list">
        <section
          v-for="group in splitList"
          :key="group.id"
          :class="[
            'doc-splits__card',
            {
              'is-dragging': draggedSplitId === group.id,
              'is-drop-target': isDropTarget(`split:${group.id}`),
            },
          ]"
          :data-split-id="group.id"
          data-slot="document-split-card"
          @dragover="onDragOverPage($event, `split:${group.id}`)"
          @drop.prevent="onDropOnSplit(group.id)"
        >
          <div class="doc-splits__card-header">
            <div class="doc-splits__card-heading">
              <button
                type="button"
                class="doc-splits__drag-handle"
                :aria-label="reorderLabel(group)"
                draggable="true"
                data-slot="document-split-drag-handle"
                @dragstart="onSplitDragStart($event, group)"
                @dragend="onDragEnd"
                @keydown="onSplitHandleKeydown($event, group)"
              >
                <HugeIcon :icon="DragDropVerticalIcon" class="doc-splits__handle-icon" />
              </button>
              <div class="doc-splits__card-copy">
                <div class="doc-splits__title">{{ group.title }}</div>
                <div class="doc-splits__range">{{ pageRangeLabel(group) }}</div>
              </div>
            </div>
            <div class="doc-splits__card-actions">
              <div class="doc-splits__count">{{ group.pages.length }}</div>
              <OfficeButton
                variant="ghost"
                size="icon-sm"
                :disabled="!canRemove()"
                :aria-label="removeLabel(group)"
                data-slot="document-split-remove"
                @click="removeSplit(group)"
              >
                <HugeIcon :icon="Delete02Icon" class="doc-splits__action-icon" />
              </OfficeButton>
            </div>
          </div>

          <div
            :class="[
              'doc-splits__dropzone',
              { 'is-drop-target': isDropTarget(`group:${group.id}`) },
            ]"
            data-slot="document-split-dropzone"
            @dragover="onDragOverPage($event, `group:${group.id}`)"
            @drop.prevent="onDropOnGroup(group.id)"
          >
            <div v-if="group.pages.length" class="doc-splits__strip">
              <div class="doc-splits__strip-inner">
                <button
                  v-for="pageId in group.pages"
                  :key="pageId"
                  type="button"
                  :class="[
                    'doc-splits__thumb',
                    { 'is-dragging': draggedPageId === pageId },
                  ]"
                  :style="{ width: `${THUMBNAIL_WIDTH}px`, height: `${THUMBNAIL_HEIGHT}px` }"
                  :aria-label="pageTitle(pageId)"
                  :data-page-id="pageId"
                  draggable="true"
                  data-slot="document-split-page"
                  @dragstart="onPageDragStart($event, pageId)"
                  @dragend="onDragEnd"
                  @dragover="onDragOverPage($event, `page:${pageId}`)"
                  @drop.stop.prevent="onDropOnPage(group.id, pageId)"
                  @click="onPageClick(pageId)"
                  @keydown="onPageKeydown($event, pageId)"
                >
                  <FileThumbnail
                    :file="{ name: `${pageId}.pdf`, type: 'application/pdf' }"
                    :preview-image-url="props.thumbnailImages[pageId]"
                    :is-loading="!props.thumbnailImages[pageId]"
                    fill
                    preview-fill
                    class="doc-splits__thumb-file"
                  />
                  <span class="doc-splits__page-number">{{ pageNumber(pageId) }}</span>
                </button>
              </div>
            </div>
            <div v-else class="doc-splits__empty-drop" data-slot="document-split-dropzone-empty">
              {{ m.vueUi.documentSplits.dropPages }}
            </div>
          </div>
        </section>
      </div>
    </div>
  </aside>
</template>

<style scoped>
.doc-splits {
  display: flex;
  height: 100%;
  min-height: 0;
  flex-direction: column;
  background: color-mix(in oklab, var(--office-muted) 20%, transparent);
  color: var(--office-foreground);
}
.doc-splits,
.doc-splits * {
  box-sizing: border-box;
}
.doc-splits--framed {
  border-top: var(--office-border-width) solid var(--office-border);
}
@media (min-width: 768px) {
  .doc-splits--framed {
    border-top: 0;
    border-left: var(--office-border-width) solid var(--office-border);
  }
}
.doc-splits__toolbar {
  display: flex;
  min-height: 48px;
  align-items: center;
  justify-content: flex-end;
  gap: var(--office-ui-space-3);
  border-bottom: var(--office-border-width) solid var(--office-border);
  background: var(--office-background);
  padding-inline: var(--office-ui-space-3);
}
.doc-splits__button-icon,
.doc-splits__action-icon,
.doc-splits__handle-icon {
  width: 16px;
  height: 16px;
  flex-shrink: 0;
}
.doc-splits__body {
  min-height: 0;
  flex: 1 1 auto;
  overflow-y: auto;
}
.doc-splits__list {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-3);
  padding: var(--office-ui-space-3);
}
.doc-splits__state-wrap {
  padding: var(--office-ui-space-3);
}
.doc-splits__state {
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: color-mix(in oklab, var(--office-muted) 30%, transparent);
  padding: var(--office-ui-space-3);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
}
.doc-splits__state--error {
  border-color: color-mix(in oklab, var(--office-destructive) 45%, transparent);
  background: color-mix(in oklab, var(--office-destructive) 8%, transparent);
  color: var(--office-destructive);
}
.doc-splits__card {
  width: 100%;
  overflow: hidden;
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-background);
  transition:
    border-color 120ms ease-out,
    box-shadow 120ms ease-out;
}
.doc-splits__card.is-dragging {
  opacity: 0.5;
}
.doc-splits__card.is-drop-target {
  border-color: color-mix(in oklab, var(--office-ring) 40%, transparent);
  box-shadow: 0 0 0 2px color-mix(in oklab, var(--office-ring) 35%, transparent);
}
.doc-splits__card-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--office-ui-space-3);
  border-bottom: var(--office-border-width) solid var(--office-border);
  padding: var(--office-ui-space-3);
}
.doc-splits__card-heading {
  display: flex;
  min-width: 0;
  align-items: center;
  gap: var(--office-ui-space-2);
}
.doc-splits__drag-handle {
  display: inline-flex;
  width: 28px;
  height: 28px;
  flex-shrink: 0;
  cursor: grab;
  align-items: center;
  justify-content: center;
  margin: 0;
  padding: 0;
  border: 0;
  border-radius: var(--office-radius-md);
  background: transparent;
  color: var(--office-muted-foreground);
}
.doc-splits__drag-handle:hover {
  background: var(--office-muted);
}
.doc-splits__drag-handle:active {
  cursor: grabbing;
}
.doc-splits__drag-handle:focus-visible {
  outline: 2px solid var(--office-ring);
  outline-offset: 1px;
}
.doc-splits__card-copy {
  min-width: 0;
}
.doc-splits__title {
  overflow: hidden;
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  font-weight: 500;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.doc-splits__range {
  margin-top: var(--office-ui-space-1);
  overflow: hidden;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
  text-overflow: ellipsis;
  white-space: nowrap;
}
.doc-splits__card-actions {
  display: flex;
  flex-shrink: 0;
  align-items: center;
  gap: var(--office-ui-space-2);
}
.doc-splits__count {
  border-radius: calc(infinity * 1px);
  background: var(--office-muted);
  padding: var(--office-ui-space-0-5) var(--office-ui-space-2);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
.doc-splits__dropzone {
  min-height: 116px;
  border-radius: var(--office-radius-lg);
  padding: var(--office-ui-space-2);
  transition:
    background-color 120ms ease-out,
    box-shadow 120ms ease-out;
}
.doc-splits__dropzone.is-drop-target {
  background: color-mix(in oklab, var(--office-accent) 30%, transparent);
  box-shadow: inset 0 0 0 1px var(--office-border);
}
.doc-splits__strip {
  height: 110px;
  width: 100%;
  overflow-x: auto;
  overflow-y: hidden;
  padding-bottom: 14px;
}
.doc-splits__strip-inner {
  display: flex;
  width: max-content;
  gap: var(--office-ui-space-2);
  overflow-y: hidden;
  padding: var(--office-ui-space-1) 24px var(--office-ui-space-1) 0;
}
.doc-splits__thumb {
  position: relative;
  flex-shrink: 0;
  cursor: grab;
  margin: 0;
  padding: 0;
  overflow: hidden;
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-background);
  box-shadow: var(--office-shadow-xs);
  text-align: left;
  transition:
    border-color 120ms ease-out,
    box-shadow 120ms ease-out,
    opacity 120ms ease-out;
  outline: none;
}
.doc-splits__thumb:hover {
  border-color: color-mix(in oklab, var(--office-ring) 40%, transparent);
  box-shadow: 0 0 0 2px var(--office-ring);
}
.doc-splits__thumb:focus-visible {
  border-color: color-mix(in oklab, var(--office-ring) 40%, transparent);
  box-shadow: 0 0 0 2px var(--office-ring);
}
.doc-splits__thumb:active {
  cursor: grabbing;
}
.doc-splits__thumb.is-dragging {
  opacity: 0;
}
.doc-splits__thumb-file {
  border: 0;
  border-radius: inherit;
  background: transparent;
}
.doc-splits__page-number {
  position: absolute;
  right: var(--office-ui-space-1);
  bottom: var(--office-ui-space-1);
  border-radius: var(--office-radius-sm);
  background: color-mix(in oklab, var(--office-background) 95%, transparent);
  padding: var(--office-ui-space-0-5) var(--office-ui-space-1-5);
  color: var(--office-foreground);
  font-size: 10px;
  line-height: 15px;
  font-weight: 600;
  box-shadow: var(--office-shadow-xs);
  outline: 1px solid color-mix(in oklab, var(--office-border) 80%, transparent);
  outline-offset: -1px;
}
.doc-splits__empty-drop {
  display: grid;
  height: 104px;
  place-items: center;
  border-radius: var(--office-radius-lg);
  background: color-mix(in oklab, var(--office-muted) 35%, transparent);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
</style>
