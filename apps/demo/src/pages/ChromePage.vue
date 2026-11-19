<script setup lang="ts">
import { ref } from "vue"
import {
  DocumentViewerSidebar,
  OfficeButton,
  OfficeEmptyState,
  OfficeErrorState,
  OfficeIconButton,
  OfficeLoadingState,
  OfficePageControl,
  OfficeSearchPopover,
  OfficeStatusBar,
  OfficeThemeProvider,
  OfficeToolbar,
  OfficeToolbarGroup,
  OfficeToolbarSeparator,
  OfficeZoomControl,
  type OfficeSearchStatus,
  type OfficeTheme,
  type OfficeZoomFitMode,
} from "@arcships/vue-ui"

const theme = ref<OfficeTheme>("light")
const sidebarOpen = ref(true)
const sidebarLoading = ref(false)
const page = ref(3)
const totalPages = 12
const zoom = ref(100)
const fitMode = ref<OfficeZoomFitMode>("none")
const query = ref("")
const resultCount = ref(0)
const activeIndex = ref(0)
const searchStatus = ref<OfficeSearchStatus>("idle")
const caseSensitive = ref(false)
const wholeWord = ref(false)
const regex = ref(false)
const statusMode = ref<"normal" | "warning" | "error">("normal")
const fileName = "Q3-review.pdf"

const thumbnails = Array.from({ length: 6 }, (_, index) => index + 1)

function runSearch(value: string) {
  query.value = value
  if (!value.trim()) {
    resultCount.value = 0
    activeIndex.value = 0
    searchStatus.value = "idle"
    return
  }
  searchStatus.value = "ready"
  resultCount.value = 4
  activeIndex.value = 0
}

function nextResult() {
  if (!resultCount.value) return
  activeIndex.value = (activeIndex.value + 1) % resultCount.value
}

function previousResult() {
  if (!resultCount.value) return
  activeIndex.value = (activeIndex.value - 1 + resultCount.value) % resultCount.value
}

function submitSearch() {
  if (query.value.trim()) runSearch(query.value)
}
</script>

<template>
  <OfficeThemeProvider :theme="theme">
    <section class="chrome-page" data-testid="office-chrome-gallery">
      <header class="chrome-page__header">
        <div>
          <p class="chrome-page__eyebrow">@arcships/vue-ui</p>
          <h2>Shared Office chrome</h2>
          <p class="chrome-page__lead">
            Toolbar, page/zoom/search controls, status bar and the canonical
            document sidebar. Resize the window to inspect inline/overlay and
            narrow-toolbar wrapping.
          </p>
        </div>
        <div class="chrome-page__controls">
          <label>
            Theme
            <select v-model="theme" data-testid="chrome-gallery-theme">
              <option value="light">Light</option>
              <option value="dark">Dark</option>
              <option value="system">System</option>
            </select>
          </label>
          <label class="chrome-page__toggle">
            <input v-model="sidebarOpen" type="checkbox" data-testid="chrome-gallery-sidebar-open">
            Sidebar open
          </label>
          <label class="chrome-page__toggle">
            <input v-model="sidebarLoading" type="checkbox" data-testid="chrome-gallery-sidebar-loading">
            Sidebar skeleton
          </label>
          <label>
            Status
            <select v-model="statusMode" data-testid="chrome-gallery-status">
              <option value="normal">Normal</option>
              <option value="warning">Warning</option>
              <option value="error">Error</option>
            </select>
          </label>
        </div>
      </header>

      <section class="chrome-page__section">
        <h3>Viewer shell</h3>
        <div class="chrome-demo" data-testid="office-chrome-shell">
          <OfficeToolbar label="Demo document toolbar" data-testid="chrome-gallery-toolbar">
            <OfficeToolbarGroup gap="sm" grow>
              <OfficeIconButton
                :label="sidebarOpen ? 'Hide pages' : 'Show pages'"
                icon="file-text"
                size="icon-sm"
                :active="sidebarOpen"
                data-testid="chrome-sidebar-toggle"
                @click="sidebarOpen = !sidebarOpen"
              />
              <OfficePageControl v-model="page" :total-pages="totalPages" />
              <span class="chrome-demo__file" :title="fileName">{{ fileName }}</span>
            </OfficeToolbarGroup>
            <template #trailing>
              <OfficeToolbarGroup gap="sm">
                <OfficeSearchPopover
                  v-model="query"
                  v-model:case-sensitive="caseSensitive"
                  v-model:whole-word="wholeWord"
                  v-model:regex="regex"
                  :result-count="resultCount"
                  :active-index="activeIndex"
                  :status="searchStatus"
                  @update:model-value="runSearch"
                  @next="nextResult"
                  @previous="previousResult"
                  @submit="submitSearch"
                />
                <OfficeToolbarSeparator />
                <OfficeZoomControl v-model="zoom" v-model:fit-mode="fitMode" />
                <OfficeToolbarSeparator />
                <OfficeIconButton
                  label="Download"
                  icon="download"
                  size="icon-sm"
                  data-testid="chrome-download"
                />
              </OfficeToolbarGroup>
            </template>
          </OfficeToolbar>

          <div class="chrome-demo__body">
            <DocumentViewerSidebar
              v-model:open="sidebarOpen"
              :loading="sidebarLoading"
              :label="'Document pages'"
            >
              <template #header>
                <span class="chrome-demo__sidebar-title">Pages</span>
              </template>
              <div class="chrome-demo__thumbs">
                <button
                  v-for="index in thumbnails"
                  :key="index"
                  type="button"
                  class="chrome-demo__thumb"
                  :class="{ 'chrome-demo__thumb--active': index === page }"
                  :aria-label="`Page ${index}`"
                  data-testid="chrome-gallery-thumbnail"
                  @click="page = index"
                >
                  <span class="chrome-demo__thumb-page" />
                  <span class="chrome-demo__thumb-label">{{ index }}</span>
                </button>
              </div>
              <template #footer>
                <span class="chrome-demo__sidebar-foot">{{ totalPages }} pages</span>
              </template>
            </DocumentViewerSidebar>

            <div class="chrome-demo__main">
              <div class="chrome-demo__surface" data-testid="chrome-content">
                <article class="chrome-demo__paper" :style="{ transform: `scale(${zoom / 100})` }">
                  <h4>Quarterly review</h4>
                  <p>Document content sits below the toolbar and above the status bar.</p>
                  <p class="chrome-demo__muted">
                    Page {{ page }} of {{ totalPages }} · {{ zoom }}%
                    <template v-if="fitMode !== 'none'"> · fit {{ fitMode }}</template>
                  </p>
                  <div class="chrome-demo__lines">
                    <span v-for="line in 8" :key="line" :style="{ width: `${52 + (line * 5) % 40}%` }" />
                  </div>
                </article>
              </div>
            </div>
          </div>

          <OfficeStatusBar
            :page="page"
            :total-pages="totalPages"
            :file-name="fileName"
            :status="statusMode === 'normal' ? 'Ready' : undefined"
            :warning="statusMode === 'warning' ? 'Some images were optimized' : undefined"
            :error="statusMode === 'error' ? 'Page 5 render failed; retrying in the background' : undefined"
            :busy="sidebarLoading"
          />
        </div>
      </section>

      <section class="chrome-page__section">
        <h3>Document states</h3>
        <div class="chrome-states" data-testid="office-chrome-states">
          <article class="chrome-state-card" data-testid="chrome-state-empty">
            <h4>Empty</h4>
            <OfficeEmptyState compact :description="'Open a PDF, DOCX, XLSX or PPTX file to begin.'">
              <template #actions>
                <OfficeButton size="sm">Open file</OfficeButton>
              </template>
            </OfficeEmptyState>
          </article>
          <article class="chrome-state-card" data-testid="chrome-state-error">
            <h4>Error</h4>
            <OfficeErrorState
              compact
              :error="'The worker returned an invalid page tree.'"
              @retry="() => {}"
            />
          </article>
          <article class="chrome-state-card" data-testid="chrome-state-loading">
            <h4>Loading</h4>
            <OfficeLoadingState compact :progress="42" :description="'Rendering page 3…'" />
          </article>
        </div>
      </section>

      <section class="chrome-page__section">
        <h3>Toolbar structure</h3>
        <OfficeToolbar label="Toolbar structure demo">
          <OfficeToolbarGroup gap="md" grow>
            <OfficeIconButton label="Open file" icon="folder-open" size="icon-sm" />
            <OfficeToolbarSeparator />
            <OfficeToolbarGroup gap="sm">
              <OfficeIconButton label="Undo" icon="undo-2" size="icon-sm" disabled />
              <OfficeIconButton label="Redo" icon="redo-2" size="icon-sm" disabled />
            </OfficeToolbarGroup>
            <OfficeToolbarSeparator />
            <span class="chrome-demo__file">Four viewers, one chrome</span>
          </OfficeToolbarGroup>
          <template #center>
            <OfficePageControl :model-value="1" :total-pages="1" disabled />
          </template>
          <template #trailing>
            <OfficeToolbarGroup gap="sm">
              <OfficeIconButton label="Search" icon="search" size="icon-sm" active />
              <OfficeZoomControl :model-value="100" />
            </OfficeToolbarGroup>
          </template>
        </OfficeToolbar>
      </section>
    </section>
  </OfficeThemeProvider>
</template>

<style scoped>
.chrome-page {
  display: flex;
  flex-direction: column;
  gap: 24px;
  padding: 24px;
  min-width: 0;
  background: var(--office-surface, var(--background));
  color: var(--office-foreground, var(--foreground));
}
.chrome-page__header {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 16px;
}
.chrome-page__eyebrow {
  margin: 0 0 4px;
  color: var(--office-muted-foreground, var(--muted-foreground));
  font-size: 12px;
  font-weight: 600;
  letter-spacing: .08em;
  text-transform: uppercase;
}
.chrome-page h2 { margin: 0 0 6px; font-size: 22px; letter-spacing: -.02em; }
.chrome-page__lead { margin: 0; max-width: 60ch; color: var(--office-muted-foreground, var(--muted-foreground)); font-size: 14px; line-height: 1.55; }
.chrome-page h3 { margin: 0 0 10px; font-size: 14px; font-weight: 650; }
.chrome-page__controls { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; }
.chrome-page__controls label { display: inline-flex; align-items: center; gap: 6px; color: var(--office-muted-foreground, var(--muted-foreground)); font-size: 12px; }
.chrome-page__controls select { height: 30px; padding: 0 8px; border: 1px solid var(--office-input, var(--border)); border-radius: 7px; background: var(--office-background, var(--background)); color: inherit; }
.chrome-page__toggle { white-space: nowrap; }
.chrome-page__section { min-width: 0; }
.chrome-demo {
  position: relative;
  display: flex;
  flex-direction: column;
  align-items: stretch;
  min-width: 0;
  min-height: 520px;
  overflow: hidden;
  border: 1px solid var(--office-border, var(--border));
  border-radius: 12px;
  background: var(--office-background, var(--background));
}
.chrome-demo__body {
  position: relative;
  display: flex;
  flex: 1 1 auto;
  align-items: stretch;
  min-width: 0;
  min-height: 0;
  overflow: hidden;
}
.chrome-demo__main {
  display: flex;
  flex: 1 1 auto;
  flex-direction: column;
  min-width: 0;
  min-height: 0;
}
.chrome-demo__sidebar-title { font-size: 12px; font-weight: 650; }
.chrome-demo__sidebar-foot { color: var(--office-muted-foreground, var(--muted-foreground)); font-size: 11px; }
.chrome-demo__thumbs { display: flex; flex-direction: column; gap: 10px; }
.chrome-demo__thumb {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 4px;
  padding: 0;
  border: 0;
  background: transparent;
  color: var(--office-muted-foreground, var(--muted-foreground));
  cursor: pointer;
  font: inherit;
}
.chrome-demo__thumb-page {
  display: block;
  width: 96px;
  height: 124px;
  border: 1px solid var(--office-sidebar-border, var(--border));
  border-radius: 6px;
  background: var(--office-background, var(--background));
  box-shadow: 0 1px 2px rgb(0 0 0 / .06);
}
.chrome-demo__thumb--active .chrome-demo__thumb-page { border-color: var(--office-ring, var(--border)); outline: 2px solid var(--office-ring, var(--border)); outline-offset: 1px; }
.chrome-demo__thumb-label { font-size: 11px; }
.chrome-demo__thumb:focus-visible { outline: 2px solid var(--office-ring, var(--border)); outline-offset: 2px; border-radius: 6px; }
.chrome-demo__file {
  overflow: hidden;
  color: var(--office-muted-foreground, var(--muted-foreground));
  font-size: 13px;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.chrome-demo__surface {
  display: flex;
  flex: 1 1 auto;
  align-items: flex-start;
  justify-content: center;
  min-height: 0;
  overflow: auto;
  padding: 24px;
  background: var(--office-muted, var(--muted));
}
.chrome-demo__paper {
  width: min(100%, 420px);
  min-height: 320px;
  padding: 28px;
  border: 1px solid var(--office-border, var(--border));
  border-radius: 8px;
  background: var(--office-background, var(--background));
  box-shadow: var(--office-shadow-sm, 0 1px 3px rgb(0 0 0 / .1));
  transform-origin: top center;
  transition: transform 160ms ease-out;
}
.chrome-demo__paper h4 { margin: 0 0 10px; font-size: 18px; }
.chrome-demo__paper p { margin: 0 0 8px; font-size: 13px; line-height: 1.5; }
.chrome-demo__muted { color: var(--office-muted-foreground, var(--muted-foreground)); }
.chrome-demo__lines { display: flex; flex-direction: column; gap: 8px; margin-top: 18px; }
.chrome-demo__lines span { display: block; height: 8px; border-radius: 4px; background: var(--office-muted, var(--muted)); }
.chrome-states { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 14px; }
.chrome-state-card { overflow: hidden; border: 1px solid var(--office-border, var(--border)); border-radius: 10px; background: var(--office-background, var(--background)); }
.chrome-state-card h4 { margin: 0; padding: 10px 12px; border-bottom: 1px solid var(--office-border, var(--border)); color: var(--office-muted-foreground, var(--muted-foreground)); font-size: 11px; font-weight: 650; letter-spacing: .06em; text-transform: uppercase; }
@media (prefers-reduced-motion: reduce) {
  .chrome-demo__paper { transition: none; }
}
@media (max-width: 900px) {
  .chrome-page { padding: 16px 12px; gap: 18px; }
  .chrome-states { grid-template-columns: minmax(0, 1fr); }
  .chrome-demo { min-height: 460px; }
  .chrome-demo__surface { padding: 16px; }
}
@media (max-width: 480px) {
  .chrome-demo { min-height: 520px; }
  .chrome-demo__paper { padding: 18px; }
}
</style>
