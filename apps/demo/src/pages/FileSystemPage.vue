<script setup lang="ts">
import { computed, ref } from "vue"
import {
  FileSystem,
  OfficeThemeProvider,
  type FileSystemItem,
  type FileSystemView,
  type OfficeTheme,
} from "@arcships/vue-ui"
import { FILE_SYSTEM_FIXTURE } from "../lib/file-system-fixture"
import { readDemoTheme } from "../lib/demo-visual-settings"

const theme = ref<OfficeTheme>(readDemoTheme())
const view = ref<FileSystemView>("grid")
const readOnly = ref(false)
const currentFolderId = ref<string | null>(null)
const selectedId = ref<string | null>(null)
const items = ref<FileSystemItem[]>(structuredClone(FILE_SYSTEM_FIXTURE))
const lastEvent = ref("—")
const openedFile = ref<FileSystemItem | null>(null)

const selectionLabel = computed(() => {
  const find = (list: FileSystemItem[]): FileSystemItem | null => {
    for (const item of list) {
      if (item.id === selectedId.value) return item
      const child = item.children ? find(item.children) : null
      if (child) return child
    }
    return null
  }
  return find(items.value)?.name ?? "nothing selected"
})

function describe(item: FileSystemItem | null): string {
  return item ? `${item.kind} · ${item.name}` : "none"
}

function removeItem(list: FileSystemItem[], id: string): FileSystemItem[] {
  return list
    .filter((item) => item.id !== id)
    .map((item) =>
      item.children ? { ...item, children: removeItem(item.children, id) } : item,
    )
}

function onOpenFolder(item: FileSystemItem): void {
  lastEvent.value = `open-folder → ${item.name}`
  openedFile.value = null
}

function onOpenFile(item: FileSystemItem): void {
  lastEvent.value = `open-file → ${item.name}`
  openedFile.value = item
}

function onDeleteItem(item: FileSystemItem): void {
  items.value = removeItem(items.value, item.id)
  lastEvent.value = `delete-item → ${item.name}`
  if (openedFile.value?.id === item.id) openedFile.value = null
}

function onSelectionChange(item: FileSystemItem | null): void {
  lastEvent.value = `selection-change → ${describe(item)}`
}

function reset(): void {
  items.value = structuredClone(FILE_SYSTEM_FIXTURE)
  currentFolderId.value = null
  selectedId.value = null
  openedFile.value = null
  lastEvent.value = "reset"
}
</script>

<template>
  <OfficeThemeProvider :theme="theme">
    <section class="fs-page" data-testid="file-system-demo">
      <header class="fs-page__header">
        <div>
          <p class="fs-page__eyebrow">@arcships/vue-ui</p>
          <h2>FileSystem</h2>
          <p class="fs-page__lead">
            Finder-style browser: breadcrumb navigation, list/grid toggle, image
            thumbnails via FileThumbnail, selection and open events. Resize to
            390 / 768 / 1440 to inspect the responsive states.
          </p>
        </div>
        <div class="fs-page__controls">
          <label>
            Theme
            <select v-model="theme" data-testid="file-system-theme">
              <option value="light">Light</option>
              <option value="dark">Dark</option>
              <option value="system">System</option>
            </select>
          </label>
          <label>
            View
            <select v-model="view" data-testid="file-system-view">
              <option value="grid">Grid</option>
              <option value="list">List</option>
            </select>
          </label>
          <label class="fs-page__toggle">
            <input v-model="readOnly" type="checkbox" data-testid="file-system-readonly">
            Read only
          </label>
          <button type="button" class="fs-page__reset" data-testid="file-system-reset" @click="reset">
            Reset data
          </button>
        </div>
      </header>

      <div class="fs-page__layout">
        <FileSystem
          v-model:current-folder-id="currentFolderId"
          v-model:selected-id="selectedId"
          v-model:view="view"
          :items="items"
          :read-only="readOnly"
          title="Workspace"
          data-testid="file-system-block"
          @open-folder="onOpenFolder"
          @open-file="onOpenFile"
          @delete-item="onDeleteItem"
          @selection-change="onSelectionChange"
        />

        <aside class="fs-page__panel">
          <h3>Events</h3>
          <dl>
            <dt>Last event</dt>
            <dd data-testid="file-system-last-event">{{ lastEvent }}</dd>
            <dt>Current folder</dt>
            <dd data-testid="file-system-current-folder">{{ currentFolderId ?? "root" }}</dd>
            <dt>Selection</dt>
            <dd data-testid="file-system-selection">{{ selectionLabel }}</dd>
            <dt>Opened file</dt>
            <dd data-testid="file-system-opened">{{ openedFile?.name ?? "—" }}</dd>
          </dl>
          <p class="fs-page__hint">
            Click selects, double click opens, Enter/Space opens the focused row.
            Arrow keys move the selection, Backspace/Left goes up one folder, and
            Delete removes the focused entry (disabled when read-only).
          </p>
        </aside>
      </div>
    </section>
  </OfficeThemeProvider>
</template>

<style scoped>
.fs-page {
  display: flex;
  flex-direction: column;
  gap: 20px;
  padding: 24px;
}
.fs-page__header {
  display: flex;
  flex-wrap: wrap;
  align-items: flex-start;
  justify-content: space-between;
  gap: 16px;
}
.fs-page__eyebrow {
  margin: 0 0 4px;
  color: var(--office-muted-foreground);
  font-size: 12px;
  letter-spacing: .08em;
  text-transform: uppercase;
}
.fs-page h2 { margin: 0 0 6px; font-size: 22px; letter-spacing: -.02em; }
.fs-page__lead { margin: 0; max-width: 720px; color: var(--office-muted-foreground); font-size: 13px; line-height: 1.5; }
.fs-page__controls { display: flex; flex-wrap: wrap; align-items: center; gap: 12px; font-size: 13px; }
.fs-page__controls label { display: inline-flex; align-items: center; gap: 6px; color: var(--office-muted-foreground); }
.fs-page__controls select {
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  color: var(--office-foreground);
  font: inherit;
  padding: 5px 8px;
}
.fs-page__reset {
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  color: var(--office-foreground);
  font: inherit;
  padding: 5px 10px;
  cursor: pointer;
}
.fs-page__layout {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 260px;
  align-items: start;
  gap: 20px;
}
.fs-page__panel {
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-background);
  padding: 16px;
}
.fs-page__panel h3 { margin: 0 0 12px; font-size: 14px; font-weight: 650; }
.fs-page__panel dl { display: grid; grid-template-columns: 1fr; gap: 4px; margin: 0; }
.fs-page__panel dt { color: var(--office-muted-foreground); font-size: 11px; text-transform: uppercase; letter-spacing: .06em; }
.fs-page__panel dd { margin: 0 0 10px; font-size: 13px; overflow-wrap: anywhere; }
.fs-page__hint { margin: 0; color: var(--office-muted-foreground); font-size: 12px; line-height: 1.5; }
@media (max-width: 900px) {
  .fs-page__layout { grid-template-columns: minmax(0, 1fr); }
}
@media (max-width: 640px) {
  .fs-page { padding: 16px 12px; }
}
</style>
