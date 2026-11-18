<template>
  <nav
    class="pdf-outline"
    data-testid="pdf-surface-outline"
    :aria-label="label"
  >
    <h2 class="pdf-outline__title">{{ label }}</h2>
    <ul class="pdf-outline__tree" role="tree" :aria-label="label">
      <li v-for="entry in visibleEntries" :key="entry.id" role="none">
        <div
          :ref="(element) => setItemRef(entry.id, element)"
          class="pdf-outline__item"
          :class="{ 'pdf-outline__item--active': entry.id === activeEntryId }"
          data-testid="pdf-outline-item"
          role="treeitem"
          :aria-level="entry.level"
          :aria-expanded="entry.hasChildren ? entry.expanded : undefined"
          :aria-selected="entry.id === activeEntryId ? 'true' : 'false'"
          :data-outline-id="entry.id"
          :data-target-page-index="entry.item.pageIndex"
          :data-outline-url="entry.item.url"
          :tabindex="entry.id === focusedId ? 0 : -1"
          :style="{ paddingLeft: `${(entry.level - 1) * 12 + 8}px` }"
          @click="activate(entry)"
          @keydown="onItemKeydown($event, entry)"
        >
          <span
            v-if="entry.hasChildren"
            class="pdf-outline__chevron"
            data-testid="pdf-outline-toggle"
            aria-hidden="true"
            @click.stop="activateToggle(entry)"
          >
            {{ entry.expanded ? "▾" : "▸" }}
          </span>
          <span v-else class="pdf-outline__chevron pdf-outline__chevron--leaf" aria-hidden="true" />
          <span class="pdf-outline__label">{{ entry.item.title || labelForEntry(entry) }}</span>
        </div>
      </li>
    </ul>
  </nav>
</template>

<script setup lang="ts">
import { computed, ref, watch } from "vue"
import type { PdfOutlineItem } from "../pdf/pdf-outline"

interface OutlineEntry {
  id: string
  /** Outline item this row renders. */
  item: PdfOutlineItem
  /** 1-based depth used by `aria-level`. */
  level: number
  parentId?: string
  hasChildren: boolean
}

interface VisibleOutlineEntry extends OutlineEntry {
  expanded: boolean
}

const props = defineProps<{
  items: readonly PdfOutlineItem[]
  /** Zero-based page currently in view. */
  activePageIndex?: number
  /** Accessible name and visible heading of the panel. */
  label: string
}>()

const emit = defineEmits<{
  activate: [item: PdfOutlineItem]
}>()

const expandedIds = ref<Set<string>>(new Set())
const focusedId = ref<string | undefined>(undefined)
const itemRefs = new Map<string, HTMLElement>()

function collectEntries(
  items: readonly PdfOutlineItem[],
  parentId: string | undefined,
  level: number,
  result: OutlineEntry[],
): void {
  for (let index = 0; index < items.length; index++) {
    const item = items[index]
    const id = parentId ? `${parentId}-${index}` : String(index)
    const hasChildren = item.children.length > 0
    result.push({ id, item, level, parentId, hasChildren })
    if (hasChildren) collectEntries(item.children, id, level + 1, result)
  }
}

const allEntries = computed<OutlineEntry[]>(() => {
  const entries: OutlineEntry[] = []
  collectEntries(props.items, undefined, 1, entries)
  return entries
})

const visibleEntries = computed<VisibleOutlineEntry[]>(() => {
  const visible: VisibleOutlineEntry[] = []
  let hiddenBelowLevel: number | undefined
  for (const entry of allEntries.value) {
    if (hiddenBelowLevel !== undefined && entry.level > hiddenBelowLevel) continue
    hiddenBelowLevel = undefined
    const expanded = expandedIds.value.has(entry.id)
    visible.push({ ...entry, expanded })
    if (entry.hasChildren && !expanded) hiddenBelowLevel = entry.level
  }
  return visible
})

const activeEntryId = computed<string | undefined>(() => {
  const activePage = props.activePageIndex ?? -1
  let active: string | undefined
  let activePageIndex = -1
  for (const entry of visibleEntries.value) {
    const pageIndex = entry.item.pageIndex
    if (pageIndex === undefined || pageIndex > activePage) continue
    if (pageIndex >= activePageIndex) {
      activePageIndex = pageIndex
      active = entry.id
    }
  }
  return active
})

watch(
  () => props.items,
  (items) => {
    const roots = new Set<string>()
    items.forEach((item, index) => {
      if (item.children.length > 0) roots.add(String(index))
    })
    expandedIds.value = roots
    focusedId.value = items.length > 0 ? "0" : undefined
  },
  { immediate: true },
)

watch(visibleEntries, (entries) => {
  if (entries.length === 0) {
    focusedId.value = undefined
    return
  }
  const focused = focusedId.value
  if (focused && entries.some((entry) => entry.id === focused)) return
  const active = activeEntryId.value
  focusedId.value = active && entries.some((entry) => entry.id === active) ? active : entries[0].id
})

function labelForEntry(entry: OutlineEntry): string {
  const pageIndex = entry.item.pageIndex
  return pageIndex === undefined ? "" : String(pageIndex + 1)
}

function setItemRef(id: string, element: unknown): void {
  if (element) itemRefs.set(id, element as HTMLElement)
  else itemRefs.delete(id)
}

function focusEntry(id: string): void {
  focusedId.value = id
  itemRefs.get(id)?.focus?.()
}

function toggleExpanded(entry: OutlineEntry, expanded: boolean): void {
  if (!entry.hasChildren) return
  const next = new Set(expandedIds.value)
  if (expanded) next.add(entry.id)
  else next.delete(entry.id)
  expandedIds.value = next
}

function activate(entry: OutlineEntry): void {
  focusEntry(entry.id)
  if (entry.hasChildren && !expandedIds.value.has(entry.id)) toggleExpanded(entry, true)
  emit("activate", entry.item)
}

/** Mouse users toggle a branch from its chevron without leaving the tree. */
function activateToggle(entry: OutlineEntry): void {
  focusEntry(entry.id)
  toggleExpanded(entry, !expandedIds.value.has(entry.id))
}

function firstChild(entry: OutlineEntry): VisibleOutlineEntry | undefined {
  const visible = visibleEntries.value
  const index = visible.findIndex((candidate) => candidate.id === entry.id)
  if (index < 0 || !entry.hasChildren) return undefined
  const next = visible[index + 1]
  return next && next.level === entry.level + 1 ? next : undefined
}

function onItemKeydown(event: KeyboardEvent, entry: OutlineEntry): void {
  if (!handleItemKey(event, entry)) return
  event.preventDefault()
  event.stopPropagation?.()
}

function handleItemKey(event: KeyboardEvent, entry: OutlineEntry): boolean {
  const visible = visibleEntries.value
  const index = visible.findIndex((candidate) => candidate.id === entry.id)
  if (index < 0) return false
  switch (event.key) {
    case "ArrowDown": {
      const next = visible[index + 1]
      if (!next) return false
      focusEntry(next.id)
      return true
    }
    case "ArrowUp": {
      const previous = visible[index - 1]
      if (!previous) return false
      focusEntry(previous.id)
      return true
    }
    case "ArrowRight": {
      if (!entry.hasChildren) return false
      if (!expandedIds.value.has(entry.id)) {
        toggleExpanded(entry, true)
        return true
      }
      const child = firstChild(entry)
      if (!child) return false
      focusEntry(child.id)
      return true
    }
    case "ArrowLeft": {
      if (entry.hasChildren && expandedIds.value.has(entry.id)) {
        toggleExpanded(entry, false)
        return true
      }
      if (!entry.parentId) return false
      focusEntry(entry.parentId)
      return true
    }
    case "Home": {
      const first = visible[0]
      if (!first || first.id === focusedId.value) return false
      focusEntry(first.id)
      return true
    }
    case "End": {
      const last = visible[visible.length - 1]
      if (!last || last.id === focusedId.value) return false
      focusEntry(last.id)
      return true
    }
    case "Enter":
    case " ":
    case "Spacebar": {
      activate(entry)
      return true
    }
    default:
      return false
  }
}
</script>

<style scoped>
.pdf-outline {
  display: flex;
  min-width: 0;
  flex-direction: column;
  gap: 4px;
}
.pdf-outline__title {
  margin: 0;
  padding: 4px 8px;
  color: var(--pdf-muted, #6b7280);
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
}
.pdf-outline__tree {
  display: flex;
  min-width: 0;
  flex-direction: column;
  gap: 1px;
  margin: 0;
  padding: 0;
  list-style: none;
}
.pdf-outline__item {
  display: flex;
  min-width: 0;
  align-items: baseline;
  gap: 4px;
  border-radius: 4px;
  color: var(--pdf-fg, #111827);
  cursor: pointer;
  font-size: 12px;
  line-height: 1.4;
  padding-bottom: 4px;
  padding-right: 8px;
  padding-top: 4px;
  user-select: none;
}
.pdf-outline__item:hover { background: rgb(37 99 235 / 0.08); }
.pdf-outline__item--active {
  background: rgb(37 99 235 / 0.14);
  color: var(--pdf-link, #1d4ed8);
  font-weight: 600;
}
.pdf-outline__item:focus-visible { outline: 2px solid #2563eb; outline-offset: -2px; }
.pdf-outline__chevron {
  flex: 0 0 12px;
  width: 12px;
  color: var(--pdf-muted, #6b7280);
  text-align: center;
}
.pdf-outline__chevron--leaf { visibility: hidden; }
.pdf-outline__label {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
</style>
