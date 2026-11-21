<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed, nextTick, ref } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"
import OfficeIcon from "./OfficeIcon.vue"
import OfficeIconButton from "./ui/OfficeIconButton.vue"
import FileThumbnail from "./FileThumbnail.vue"
import { useOfficeUiLocale, useOfficeUiMessages } from "../i18n"
import {
  buildFileSystemIndex,
  fileSystemAncestors,
  fileSystemEntries,
  fileSystemPreviewUrl,
  formatFileSystemSize,
  resolveFileSystemFolder,
  type FileSystemNode,
} from "../file-system"
import type { FileSystemItem, FileSystemProps, FileSystemView } from "../types"

const props = withDefaults(defineProps<FileSystemProps>(), {
  currentFolderId: undefined,
  defaultFolderId: null,
  view: undefined,
  defaultView: "grid",
  selectedId: undefined,
  title: undefined,
  readOnly: false,
  loading: false,
  error: undefined,
  emptyMessage: undefined,
  className: "",
})

const emit = defineEmits<{
  "update:currentFolderId": [folderId: string | null]
  "update:view": [view: FileSystemView]
  "update:selectedId": [itemId: string | null]
  "selection-change": [item: FileSystemItem | null]
  "open-folder": [item: FileSystemItem]
  "open-file": [item: FileSystemItem]
  "delete-item": [item: FileSystemItem]
}>()

const m = useOfficeUiMessages()
const locale = useOfficeUiLocale()
const copy = computed(() => m.value.vueUi.fileSystem)
const title = computed(() => props.title ?? copy.value.title)

const internalFolderId = ref<string | null>(props.defaultFolderId)
const internalView = ref<FileSystemView>(props.defaultView)
const internalSelectedId = ref<string | null>(null)

const index = computed(() => buildFileSystemIndex(props.items))

const requestedFolderId = computed<string | null>(() =>
  props.currentFolderId !== undefined ? props.currentFolderId : internalFolderId.value,
)
const currentFolderId = computed<string | null>(() =>
  resolveFileSystemFolder(index.value, requestedFolderId.value),
)

const entries = computed(() => fileSystemEntries(index.value, currentFolderId.value))
const ancestors = computed(() => fileSystemAncestors(index.value, currentFolderId.value))
const breadcrumb = computed(() => [
  { id: null as string | null, name: title.value },
  ...ancestors.value.map((node) => ({ id: node.item.id, name: node.item.name })),
])

const view = computed<FileSystemView>(() => props.view ?? internalView.value)
const emptyCopy = computed(() => props.emptyMessage ?? copy.value.empty)
const selectedId = computed<string | null>(() =>
  props.selectedId !== undefined ? props.selectedId : internalSelectedId.value,
)
const visibleSelectedId = computed<string | null>(() =>
  entries.value.some((node) => node.item.id === selectedId.value) ? selectedId.value : null,
)
const activeItemId = computed<string | null>(
  () => visibleSelectedId.value ?? entries.value[0]?.item.id ?? null,
)
const showContent = computed(() => !props.error && !props.loading)

const history = ref<Array<string | null>>([currentFolderId.value])
const historyIndex = ref(0)
const canGoBack = computed(() => historyIndex.value > 0)
const canGoForward = computed(() => historyIndex.value < history.value.length - 1)

const rowRefs = new Map<string, HTMLElement>()

const dateFormatter = computed(() => {
  try {
    return new Intl.DateTimeFormat(locale.value, {
      year: "numeric",
      month: "short",
      day: "numeric",
    })
  } catch {
    return null
  }
})

function setRowRef(itemId: string, element: unknown): void {
  if (element instanceof HTMLElement) rowRefs.set(itemId, element)
  else rowRefs.delete(itemId)
}

function focusRow(itemId: string | null): void {
  if (!itemId) return
  void nextTick(() => rowRefs.get(itemId)?.focus())
}

function emitSelection(item: FileSystemItem | null): void {
  const nextId = item?.id ?? null
  if (props.selectedId === undefined) internalSelectedId.value = nextId
  emit("update:selectedId", nextId)
  emit("selection-change", item)
}

function selectItem(item: FileSystemItem, options: { focus?: boolean } = {}): void {
  emitSelection(item)
  if (options.focus) focusRow(item.id)
}

function goToFolder(folderId: string | null, options: { push?: boolean } = {}): void {
  const target = resolveFileSystemFolder(index.value, folderId)
  const previous = currentFolderId.value
  const { push = true } = options
  if (props.currentFolderId === undefined) internalFolderId.value = target
  emit("update:currentFolderId", target)
  if (push && target !== previous) {
    const stack = history.value.slice(0, historyIndex.value + 1)
    stack.push(target)
    history.value = stack
    historyIndex.value = stack.length - 1
  }
  if (visibleSelectedId.value) emitSelection(null)
  rowRefs.clear()
}

function goBack(): void {
  if (!canGoBack.value) return
  historyIndex.value -= 1
  goToFolder(history.value[historyIndex.value] ?? null, { push: false })
}

function goForward(): void {
  if (!canGoForward.value) return
  historyIndex.value += 1
  goToFolder(history.value[historyIndex.value] ?? null, { push: false })
}

function setView(next: FileSystemView): void {
  if (props.view === undefined) internalView.value = next
  emit("update:view", next)
}

function openItem(item: FileSystemItem): void {
  if (item.kind === "folder") {
    emit("open-folder", item)
    goToFolder(item.id)
    return
  }
  emitSelection(item)
  emit("open-file", item)
}

function deleteItem(item: FileSystemItem): void {
  if (props.readOnly) return
  if (selectedId.value === item.id) emitSelection(null)
  emit("delete-item", item)
}

function moveSelection(delta: number): void {
  if (entries.value.length === 0) return
  const current = entries.value.findIndex((node) => node.item.id === activeItemId.value)
  const nextIndex = Math.max(0, Math.min(entries.value.length - 1, current + delta))
  const next = entries.value[nextIndex]
  if (!next) return
  selectItem(next.item, { focus: true })
}

function selectEdge(edge: "start" | "end"): void {
  const node = edge === "start" ? entries.value[0] : entries.value[entries.value.length - 1]
  if (!node) return
  selectItem(node.item, { focus: true })
}

function goUp(): void {
  const parentId = ancestors.value[ancestors.value.length - 1]?.parentId ?? null
  goToFolder(parentId)
}

function onItemKeydown(event: KeyboardEvent, item: FileSystemItem): void {
  switch (event.key) {
    case "ArrowDown":
      event.preventDefault()
      moveSelection(1)
      return
    case "ArrowUp":
      event.preventDefault()
      moveSelection(-1)
      return
    case "Home":
      event.preventDefault()
      selectEdge("start")
      return
    case "End":
      event.preventDefault()
      selectEdge("end")
      return
    case "ArrowRight":
      if (item.kind === "folder") {
        event.preventDefault()
        openItem(item)
      }
      return
    case "ArrowLeft":
    case "Backspace":
      if (currentFolderId.value) {
        event.preventDefault()
        goUp()
      }
      return
    case "Enter":
    case " ":
      event.preventDefault()
      openItem(item)
      return
    case "Delete":
      if (!props.readOnly) {
        event.preventDefault()
        deleteItem(item)
      }
      return
    default:
      return
  }
}

function itemLabel(item: FileSystemItem): string {
  return formatOfficeMessage(
    item.kind === "folder" ? copy.value.openFolder : copy.value.openFile,
    { name: item.name },
  )
}

function deleteLabel(item: FileSystemItem): string {
  return formatOfficeMessage(copy.value.deleteItem, { name: item.name })
}

function formatDate(iso: string | undefined): string | null {
  if (!iso) return null
  const date = new Date(iso)
  if (Number.isNaN(date.getTime())) return null
  return dateFormatter.value?.format(date) ?? null
}

function itemMeta(item: FileSystemItem): string {
  const size = formatFileSystemSize(item.size)
  const date = formatDate(item.updatedAt)
  if (size && date) return formatOfficeMessage(copy.value.itemMeta, { size, date })
  return size ?? date ?? ""
}

function typeLabel(item: FileSystemItem): string {
  if (item.kind === "folder") return copy.value.folderKind
  return item.mime ?? ""
}

function isSelected(item: FileSystemItem): boolean {
  return visibleSelectedId.value === item.id
}

function isActive(item: FileSystemItem): boolean {
  return activeItemId.value === item.id
}

function rowTabIndex(item: FileSystemItem): number {
  return isActive(item) ? 0 : -1
}

function previewUrl(item: FileSystemItem): string | null {
  return fileSystemPreviewUrl(item)
}

function isFolder(item: FileSystemItem): boolean {
  return item.kind === "folder"
}

function hasPreview(item: FileSystemItem): boolean {
  return previewUrl(item) !== null
}

function crumbLabel(crumb: { id: string | null; name: string }): string {
  return crumb.name
}

function thumbnailFile(item: FileSystemItem): { name: string; type: string } {
  return { name: item.name, type: item.mime ?? "" }
}

function isCurrentCrumb(folderId: string | null): boolean {
  return folderId === currentFolderId.value
}

function childCount(node: FileSystemNode): number {
  return index.value.children.get(node.item.id)?.length ?? 0
}

defineExpose({
  /** Navigate to a folder id (or `null` for the root) from a host command. */
  navigate: (folderId: string | null) => goToFolder(folderId),
  /** Current folder id after resolving stale ids to the root. */
  currentFolder: () => currentFolderId.value,
})
</script>

<template>
  <div
    :class="['file-system', 'office-ui', props.className]"
    role="group"
    :aria-label="title"
    :data-view="view"
    :data-current-folder-id="currentFolderId ?? ''"
    :data-read-only="props.readOnly ? 'true' : undefined"
    data-slot="file-system"
  >
    <div class="file-system__toolbar" data-slot="file-system-toolbar">
      <div class="file-system__nav">
        <OfficeIconButton
          :label="copy.back"
          icon="arrow-left"
          size="icon-sm"
          :disabled="!canGoBack"
          data-slot="file-system-back"
          @click="goBack"
        />
        <OfficeIconButton
          :label="copy.forward"
          icon="arrow-right"
          size="icon-sm"
          :disabled="!canGoForward"
          data-slot="file-system-forward"
          @click="goForward"
        />
      </div>

      <nav class="file-system__breadcrumb" :aria-label="copy.breadcrumb" data-slot="file-system-breadcrumb">
        <ol class="file-system__crumbs">
          <li v-for="(crumb, crumbIndex) in breadcrumb" :key="crumb.id ?? 'root'" class="file-system__crumb-item">
            <button
              type="button"
              :class="[
                'file-system__crumb',
                { 'is-current': isCurrentCrumb(crumb.id) },
              ]"
              :aria-current="isCurrentCrumb(crumb.id) ? 'page' : undefined"
              :data-folder-id="crumb.id ?? ''"
              data-slot="file-system-crumb"
              @click="goToFolder(crumb.id)"
            >
              {{ crumbLabel(crumb) }}
            </button>
            <OfficeIcon
              v-if="crumbIndex < breadcrumb.length - 1"
              name="chevron-right"
              class="file-system__crumb-separator"
            />
          </li>
        </ol>
      </nav>

      <div class="file-system__view-toggle" role="group" :aria-label="copy.view" data-slot="file-system-view-toggle">
        <button
          type="button"
          :class="['file-system__view-button', { 'is-active': view === 'list' }]"
          :aria-pressed="view === 'list'"
          :aria-label="copy.listView"
          data-slot="file-system-view-list"
          @click="setView('list')"
        >
          <OfficeIcon name="list" />
        </button>
        <button
          type="button"
          :class="['file-system__view-button', { 'is-active': view === 'grid' }]"
          :aria-pressed="view === 'grid'"
          :aria-label="copy.gridView"
          data-slot="file-system-view-grid"
          @click="setView('grid')"
        >
          <OfficeIcon name="layout-grid" />
        </button>
      </div>
    </div>

    <div class="file-system__body" data-slot="file-system-body">
      <div v-if="props.error" class="file-system__state file-system__state--error" role="alert" data-slot="file-system-error">
        {{ props.error }}
      </div>
      <div
        v-else-if="props.loading"
        class="file-system__state"
        role="status"
        data-slot="file-system-loading"
      >
        {{ copy.loading }}
      </div>
      <div
        v-else-if="entries.length === 0"
        class="file-system__state"
        role="status"
        data-slot="file-system-empty"
      >
        {{ emptyCopy }}
      </div>

      <div
        v-else-if="showContent && view === 'list'"
        class="file-system__list"
        role="list"
        data-slot="file-system-list"
      >
        <div
          v-for="node in entries"
          :key="node.item.id"
          role="listitem"
          :class="['file-system__row', { 'is-selected': isSelected(node.item) }]"
          :data-item-id="node.item.id"
          :data-kind="node.item.kind"
          data-slot="file-system-item"
        >
          <button
            :ref="(element) => setRowRef(node.item.id, element)"
            type="button"
            :class="['file-system__row-main', { 'is-selected': isSelected(node.item) }]"
            :aria-label="itemLabel(node.item)"
            :aria-current="isSelected(node.item) ? 'true' : undefined"
            :tabindex="rowTabIndex(node.item)"
            :data-item-id="node.item.id"
            data-slot="file-system-item-main"
            @click="selectItem(node.item)"
            @dblclick="openItem(node.item)"
            @keydown="onItemKeydown($event, node.item)"
          >
            <span class="file-system__glyph">
              <FileThumbnail
                v-if="hasPreview(node.item)"
                :file="thumbnailFile(node.item)"
                :preview-image-url="previewUrl(node.item)"
                size="sm"
                alt=""
                class="file-system__thumb"
              />
              <OfficeIcon v-else :name="isFolder(node.item) ? 'folder' : 'file'" />
            </span>
            <span class="file-system__row-copy">
              <span class="file-system__name">{{ node.item.name }}</span>
              <span class="file-system__meta">{{ typeLabel(node.item) }}</span>
            </span>
            <span class="file-system__row-meta">{{ itemMeta(node.item) }}</span>
          </button>
          <button
            v-if="!props.readOnly"
            type="button"
            class="file-system__delete"
            :aria-label="deleteLabel(node.item)"
            data-slot="file-system-item-delete"
            @click.stop="deleteItem(node.item)"
          >
            <OfficeIcon name="trash-2" />
          </button>
        </div>
      </div>

      <div
        v-else-if="showContent"
        class="file-system__grid"
        role="list"
        data-slot="file-system-grid"
      >
        <div
          v-for="node in entries"
          :key="node.item.id"
          role="listitem"
          :class="['file-system__card', { 'is-selected': isSelected(node.item) }]"
          :data-item-id="node.item.id"
          :data-kind="node.item.kind"
          data-slot="file-system-item"
        >
          <button
            :ref="(element) => setRowRef(node.item.id, element)"
            type="button"
            :class="['file-system__card-main', { 'is-selected': isSelected(node.item) }]"
            :aria-label="itemLabel(node.item)"
            :aria-current="isSelected(node.item) ? 'true' : undefined"
            :tabindex="rowTabIndex(node.item)"
            :data-item-id="node.item.id"
            data-slot="file-system-item-main"
            @click="selectItem(node.item)"
            @dblclick="openItem(node.item)"
            @keydown="onItemKeydown($event, node.item)"
          >
            <span class="file-system__preview">
              <FileThumbnail
                v-if="hasPreview(node.item)"
                :file="thumbnailFile(node.item)"
                :preview-image-url="previewUrl(node.item)"
                fill
                preview-fill
                alt=""
                class="file-system__preview-image"
              />
              <OfficeIcon
                v-else
                :name="isFolder(node.item) ? 'folder' : 'file'"
                class="file-system__preview-glyph"
              />
            </span>
            <span class="file-system__card-copy">
              <span class="file-system__name">{{ node.item.name }}</span>
              <span class="file-system__meta">{{ itemMeta(node.item) }}</span>
            </span>
          </button>
          <button
            v-if="!props.readOnly"
            type="button"
            class="file-system__delete file-system__delete--card"
            :aria-label="deleteLabel(node.item)"
            data-slot="file-system-item-delete"
            @click.stop="deleteItem(node.item)"
          >
            <OfficeIcon name="trash-2" />
          </button>
          <span
            v-if="isFolder(node.item) && childCount(node) > 0"
            class="file-system__count"
            :data-count="childCount(node)"
          >{{ childCount(node) }}</span>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.file-system {
  display: flex;
  min-width: 0;
  height: 480px;
  max-height: 100%;
  flex-direction: column;
  overflow: hidden;
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-background);
  color: var(--office-foreground);
}
.file-system,
.file-system * {
  box-sizing: border-box;
}
.file-system__toolbar {
  display: flex;
  flex-shrink: 0;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--office-ui-space-2);
  border-bottom: var(--office-border-width) solid var(--office-border);
  background: color-mix(in oklab, var(--office-muted) 40%, var(--office-background));
  padding: var(--office-ui-space-2) var(--office-ui-space-3);
  min-height: 48px;
}
.file-system__nav {
  display: flex;
  flex-shrink: 0;
  align-items: center;
  gap: var(--office-ui-space-1);
}
.file-system__breadcrumb {
  min-width: 0;
  flex: 1 1 auto;
  overflow: hidden;
}
.file-system__crumbs {
  display: flex;
  min-width: 0;
  align-items: center;
  gap: var(--office-ui-space-1);
  margin: 0;
  padding: 0;
  list-style: none;
  overflow-x: auto;
  scrollbar-width: none;
}
.file-system__crumb-item {
  display: inline-flex;
  flex-shrink: 0;
  align-items: center;
  gap: var(--office-ui-space-1);
}
.file-system__crumb {
  margin: 0;
  border: var(--office-border-width) solid transparent;
  border-radius: var(--office-radius-md);
  background: transparent;
  color: var(--office-muted-foreground);
  cursor: pointer;
  font: inherit;
  font-size: var(--office-text-sm);
  line-height: 1.25;
  padding: var(--office-ui-space-1) var(--office-ui-space-2);
  white-space: nowrap;
}
.file-system__crumb:hover {
  background: var(--office-accent);
  color: var(--office-foreground);
}
.file-system__crumb.is-current {
  color: var(--office-foreground);
  font-weight: 600;
}
.file-system__crumb:focus-visible {
  outline: 2px solid var(--office-ring);
  outline-offset: 1px;
}
.file-system__crumb-separator {
  width: 12px;
  height: 12px;
  flex-shrink: 0;
  color: var(--office-muted-foreground);
  opacity: 0.7;
}
.file-system__view-toggle {
  display: inline-flex;
  flex-shrink: 0;
  align-items: center;
  gap: 2px;
  margin-inline-start: auto;
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  padding: 2px;
}
.file-system__view-button {
  display: inline-flex;
  width: var(--office-ui-height-sm);
  height: var(--office-ui-height-sm);
  align-items: center;
  justify-content: center;
  margin: 0;
  border: 0;
  border-radius: var(--office-radius-sm);
  background: transparent;
  color: var(--office-muted-foreground);
  cursor: pointer;
  padding: 0;
}
.file-system__view-button:hover {
  background: var(--office-accent);
  color: var(--office-foreground);
}
.file-system__view-button.is-active {
  background: var(--office-secondary);
  color: var(--office-foreground);
}
.file-system__view-button:focus-visible {
  outline: 2px solid var(--office-ring);
  outline-offset: 1px;
}
.file-system__body {
  min-height: 0;
  flex: 1 1 auto;
  overflow: auto;
  padding: var(--office-ui-space-2);
}
.file-system__state {
  display: flex;
  height: 100%;
  min-height: 160px;
  align-items: center;
  justify-content: center;
  padding: var(--office-ui-space-6);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-sm);
  text-align: center;
}
.file-system__state--error {
  color: var(--office-destructive);
}
.file-system__list {
  display: flex;
  min-width: 0;
  flex-direction: column;
  gap: 2px;
}
.file-system__row {
  display: flex;
  min-width: 0;
  align-items: center;
  gap: var(--office-ui-space-1);
  border-radius: var(--office-radius-md);
}
.file-system__row:hover {
  background: color-mix(in oklab, var(--office-accent) 60%, transparent);
}
.file-system__row.is-selected {
  background: var(--office-secondary);
}
.file-system__row-main {
  display: flex;
  min-width: 0;
  flex: 1 1 auto;
  align-items: center;
  gap: var(--office-ui-space-2);
  margin: 0;
  border: var(--office-border-width) solid transparent;
  border-radius: var(--office-radius-md);
  background: transparent;
  color: inherit;
  cursor: pointer;
  font: inherit;
  padding: var(--office-ui-space-1-5) var(--office-ui-space-2);
  text-align: start;
}
.file-system__row-main:focus-visible {
  outline: 2px solid var(--office-ring);
  outline-offset: -2px;
}
.file-system__glyph {
  display: inline-flex;
  width: 28px;
  height: 28px;
  flex-shrink: 0;
  align-items: center;
  justify-content: center;
  color: var(--office-muted-foreground);
}
.file-system__thumb {
  overflow: hidden;
  border-radius: var(--office-radius-sm);
}
.file-system__row-copy,
.file-system__card-copy {
  display: flex;
  min-width: 0;
  flex-direction: column;
  gap: 2px;
}
.file-system__row-copy {
  flex: 1 1 auto;
}
.file-system__name {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-size: var(--office-text-sm);
  font-weight: 500;
}
.file-system__meta {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
}
.file-system__row-meta {
  flex-shrink: 0;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  white-space: nowrap;
}
.file-system__delete {
  display: inline-flex;
  width: var(--office-ui-height-sm);
  height: var(--office-ui-height-sm);
  flex-shrink: 0;
  align-items: center;
  justify-content: center;
  margin: 0;
  border: 0;
  border-radius: var(--office-radius-sm);
  background: transparent;
  color: var(--office-muted-foreground);
  cursor: pointer;
  padding: 0;
}
.file-system__delete:hover {
  background: color-mix(in oklab, var(--office-destructive) 12%, transparent);
  color: var(--office-destructive);
}
.file-system__delete:focus-visible {
  outline: 2px solid var(--office-ring);
  outline-offset: 1px;
}
.file-system__grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(148px, 1fr));
  gap: var(--office-ui-space-2);
}
.file-system__card {
  position: relative;
  display: flex;
  min-width: 0;
  flex-direction: column;
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: var(--office-background);
}
.file-system__card:hover {
  border-color: color-mix(in oklab, var(--office-ring) 45%, var(--office-border));
}
.file-system__card.is-selected {
  border-color: var(--office-ring);
  background: color-mix(in oklab, var(--office-secondary) 60%, var(--office-background));
}
.file-system__card-main {
  display: flex;
  min-width: 0;
  flex-direction: column;
  align-items: stretch;
  gap: var(--office-ui-space-2);
  margin: 0;
  border: 0;
  border-radius: var(--office-radius-md);
  background: transparent;
  color: inherit;
  cursor: pointer;
  font: inherit;
  padding: var(--office-ui-space-2);
  text-align: start;
}
.file-system__card-main:focus-visible {
  outline: 2px solid var(--office-ring);
  outline-offset: -2px;
}
.file-system__preview {
  display: flex;
  aspect-ratio: 4 / 3;
  align-items: center;
  justify-content: center;
  overflow: hidden;
  border-radius: var(--office-radius-sm);
  background: color-mix(in oklab, var(--office-muted) 70%, var(--office-background));
  color: var(--office-muted-foreground);
}
.file-system__preview-image {
  width: 100%;
  height: 100%;
}
.file-system__preview-glyph {
  width: 28px;
  height: 28px;
}
.file-system__delete--card {
  position: absolute;
  top: 4px;
  right: 4px;
  background: color-mix(in oklab, var(--office-background) 88%, transparent);
}
.file-system__count {
  position: absolute;
  top: 4px;
  left: 4px;
  border-radius: var(--office-radius-full);
  background: color-mix(in oklab, var(--office-background) 88%, transparent);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: 1;
  padding: 3px 6px;
}
@media (max-width: 640px) {
  .file-system {
    height: 420px;
  }
  .file-system__toolbar {
    gap: var(--office-ui-space-1);
    padding: var(--office-ui-space-1-5) var(--office-ui-space-2);
  }
  .file-system__row-meta {
    display: none;
  }
  .file-system__grid {
    grid-template-columns: repeat(auto-fill, minmax(120px, 1fr));
  }
}
@media (min-width: 768px) and (max-width: 1023px) {
  .file-system__grid {
    grid-template-columns: repeat(auto-fill, minmax(132px, 1fr));
  }
}
</style>
