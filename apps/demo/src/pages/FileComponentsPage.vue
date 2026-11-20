<script setup lang="ts">
import { nextTick, onMounted, ref } from "vue"
import {
  FileThumbnail,
  FileUpload,
  OfficeThemeProvider,
  type OfficeTheme,
} from "@arcships/vue-ui"
import { readDemoTheme } from "../lib/demo-visual-settings"

const theme = ref<OfficeTheme>(readDemoTheme())
const defaultRef = ref<InstanceType<typeof FileUpload> | null>(null)
const listRef = ref<InstanceType<typeof FileUpload> | null>(null)
const errorRef = ref<InstanceType<typeof FileUpload> | null>(null)
const dragRef = ref<InstanceType<typeof FileUpload> | null>(null)
const lastAction = ref("—")

const longName =
  "quarterly-board-review-consolidated-financial-statements-and-appendix-2026.pdf"

function makeFile(name: string, type: string, content = "parity fixture"): File {
  return new File([content], name, { type })
}

async function sampleFile(name: string, type: string, url: string): Promise<File> {
  try {
    const response = await fetch(url)
    const blob = await response.blob()
    return new File([blob], name, { type })
  } catch {
    return makeFile(name, type, "image")
  }
}

function rootOf(component: unknown): ParentNode | null {
  const instance = component as { $el?: ParentNode } | null
  return instance?.$el ?? null
}

function dispatchFiles(component: unknown, files: File[]): void {
  const root = rootOf(component)
  const input = root?.querySelector<HTMLInputElement>("input[type='file']")
  if (!input || typeof DataTransfer === "undefined") return
  const transfer = new DataTransfer()
  for (const file of files) transfer.items.add(file)
  Object.defineProperty(input, "files", {
    configurable: true,
    value: transfer.files,
  })
  input.dispatchEvent(new Event("change", { bubbles: true }))
}

function dispatchDrag(component: unknown, active: boolean): void {
  const root = rootOf(component)
  const zone = root?.querySelector<HTMLElement>("[role='button']")
  if (!zone) return
  const transfer = new DataTransfer()
  transfer.items.add(makeFile("drag-preview.pdf", "application/pdf"))
  zone.dispatchEvent(
    new DragEvent(active ? "dragenter" : "dragleave", {
      bubbles: true,
      cancelable: true,
      dataTransfer: transfer,
    }),
  )
}

async function seedStories(): Promise<void> {
  await nextTick()
  const image = await sampleFile(
    "contract-page.png",
    "image/png",
    "/samples/contract-page.png",
  )
  dispatchFiles(listRef.value, [
    makeFile("quarterly-report.pdf", "application/pdf"),
    image,
    makeFile(
      "legal-contract.docx",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "document",
    ),
  ])
  dispatchFiles(errorRef.value, [
    makeFile("meeting-notes.txt", "text/plain", "notes"),
  ])
  dispatchDrag(dragRef.value, true)
}

function selectFiles(): void {
  dispatchFiles(defaultRef.value, [
    makeFile("quarterly-report.pdf", "application/pdf"),
  ])
  lastAction.value = "Selected quarterly-report.pdf"
}

function rejectFiles(): void {
  dispatchFiles(errorRef.value, [
    makeFile("meeting-notes.txt", "text/plain", "notes"),
  ])
  lastAction.value = "Rejected meeting-notes.txt"
}

function startDrag(): void {
  dispatchDrag(dragRef.value, true)
  lastAction.value = "Dragging a PDF over the drop zone"
}

function stopDrag(): void {
  dispatchDrag(dragRef.value, false)
  lastAction.value = "Drag left the drop zone"
}

onMounted(seedStories)
</script>

<template>
  <OfficeThemeProvider :theme="theme">
    <section class="file-story" data-testid="file-components-story">
      <header class="hero">
        <div>
          <p class="eyebrow">Extend parity · issue #163</p>
          <h2>File components</h2>
          <p class="lead">
            FileUpload and FileThumbnail aligned with the pinned Extend UI
            reference: tokenized radius, borders, muted surfaces and
            interaction states, plus the file list, ready state and preview
            reveal layers.
          </p>
        </div>
        <div class="controls">
          <label>
            Theme
            <select v-model="theme" data-testid="file-story-theme">
              <option value="light">Light</option>
              <option value="dark">Dark</option>
              <option value="system">System</option>
            </select>
          </label>
          <span class="action-log" data-testid="file-story-action">{{ lastAction }}</span>
        </div>
      </header>

      <section class="story-section" data-testid="section-upload">
        <div class="section-heading">
          <h3>FileUpload</h3>
          <p>Default, drag, file list and rejected states.</p>
        </div>
        <div class="upload-grid">
          <article class="story-card" data-testid="story-upload-default">
            <h4>Default</h4>
            <FileUpload
              ref="defaultRef"
              accept=".pdf,.docx,.png"
              show-file-list
              :max-size="25 * 1024 * 1024"
              title="Click to upload or drop files"
              description="PDF, DOC/DOCX, XLSX, CSV, PNG, or JPG"
              browse-label="Browse files"
              dragging-label="Drop to add"
              @files-accepted="lastAction = `Accepted ${$event.length} file(s)`"
              @files-rejected="lastAction = `Rejected ${$event.length} file(s)`"
            />
            <button class="story-button" type="button" @click="selectFiles">
              Simulate file selection
            </button>
          </article>

          <article class="story-card" data-testid="story-upload-drag">
            <h4>Dragging</h4>
            <FileUpload
              ref="dragRef"
              show-file-list
              accept=".pdf,.png"
              title="Drop to add"
              description="The icon cluster lifts and the border switches to the active token."
              browse-label="Browse files"
              dragging-label="Drop to add"
            />
            <div class="story-actions">
              <button class="story-button" type="button" @click="startDrag">
                Start drag
              </button>
              <button class="story-button" type="button" @click="stopDrag">
                End drag
              </button>
            </div>
          </article>

          <article class="story-card" data-testid="story-upload-list">
            <h4>File list / Ready</h4>
            <FileUpload
              ref="listRef"
              show-file-list
              accept=".pdf,.docx,.png"
              title="Click to upload or drop files"
              description="Long names truncate without moving the Ready badge."
            />
          </article>

          <article class="story-card" data-testid="story-upload-error">
            <h4>Rejected</h4>
            <FileUpload
              ref="errorRef"
              show-file-list
              accept=".pdf,.docx,.png"
              title="Click to upload or drop files"
              description="Unsupported files stay out of the list."
            />
            <button class="story-button" type="button" @click="rejectFiles">
              Simulate rejected file
            </button>
          </article>
        </div>
      </section>

      <section class="story-section" data-testid="section-thumbnail">
        <div class="section-heading">
          <h3>FileThumbnail</h3>
          <p>
            Square and custom aspect-ratio preview cards with blur-in,
            loading, error and file-type glyph layers.
          </p>
        </div>
        <div class="thumbnail-grid">
          <article class="story-card thumbnail-story" data-testid="story-thumbnail-empty">
            <h4>Empty fallback</h4>
            <FileThumbnail
              :file="{ name: 'quarterly-report.pdf', type: 'application/pdf' }"
              :size="96"
            />
          </article>
          <article class="story-card thumbnail-story" data-testid="story-thumbnail-image">
            <h4>Image blur-in</h4>
            <FileThumbnail
              :file="{ name: 'contract-page.png', type: 'image/png' }"
              :size="96"
              preview-image-url="/samples/contract-page.png"
            />
          </article>
          <article class="story-card thumbnail-story" data-testid="story-thumbnail-loading">
            <h4>Loading overlay</h4>
            <FileThumbnail
              :file="{ name: 'loading.pdf', type: 'application/pdf' }"
              :size="96"
              is-loading
            />
          </article>
          <article class="story-card thumbnail-story" data-testid="story-thumbnail-error">
            <h4>Error + type glyph</h4>
            <FileThumbnail
              :file="{ name: 'financial-model.xlsx', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }"
              :size="96"
              has-error
              show-type-glyph
            />
          </article>
          <article class="story-card thumbnail-story" data-testid="story-thumbnail-aspect">
            <h4>16:9 preview</h4>
            <FileThumbnail
              :file="{ name: 'contract-page.png', type: 'image/png' }"
              :size="160"
              :preview-aspect-ratio="16 / 9"
              preview-image-url="/samples/contract-page.png"
            />
          </article>
          <article class="story-card thumbnail-story" data-testid="story-thumbnail-long">
            <h4>Long name</h4>
            <FileThumbnail
              :file="{ name: longName, type: 'application/pdf' }"
              :size="120"
              show-meta
              show-type-glyph
            />
          </article>
        </div>
      </section>

      <section class="story-section seam" data-testid="document-sidebar-seam">
        <div>
          <h3>DocumentViewerSidebar integration seam</h3>
          <p>
            The canonical <code>DocumentViewerSidebar</code> is delivered by
            issue #158 at
            <code>packages/vue-ui/src/components/chrome/</code>. #163 does not
            duplicate it: the thumbnail strip will compose these
            <code>FileThumbnail</code> cards inside the sidebar's default slot,
            with the viewer owning open state, width measurement and the
            inline/overlay switch. The sidebar-dependent demo and parity cells
            land with #158, then the four viewer packages reuse the same
            component.
          </p>
        </div>
      </section>
    </section>
  </OfficeThemeProvider>
</template>

<style scoped>
.file-story {
  width: 100%;
  max-width: 1240px;
  margin: 0 auto;
  padding: 24px;
  color: var(--office-foreground);
  background: var(--office-background);
}
.hero {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 24px;
  flex-wrap: wrap;
  margin-bottom: 28px;
}
.eyebrow {
  margin: 0 0 6px;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  letter-spacing: 0.08em;
  text-transform: uppercase;
}
h2 {
  margin: 0 0 8px;
  font-size: 1.5rem;
  line-height: 1.2;
}
.lead,
.section-heading p,
.seam p {
  margin: 0;
  max-width: 720px;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
}
.controls {
  display: flex;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
  font-size: var(--office-text-xs);
}
.controls label {
  display: flex;
  align-items: center;
  gap: 8px;
  color: var(--office-muted-foreground);
}
.controls select {
  min-height: 2.25rem;
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  color: var(--office-foreground);
  font: inherit;
  padding: 0 8px;
}
.action-log {
  color: var(--office-muted-foreground);
}
.story-section {
  margin-top: 28px;
}
.section-heading {
  margin-bottom: 14px;
}
.section-heading h3,
.seam h3 {
  margin: 0 0 4px;
  font-size: 1rem;
}
.upload-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(min(300px, 100%), 1fr));
  gap: 16px;
}
.story-card {
  min-width: 0;
  padding: 16px;
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-xl);
  background: var(--office-card);
}
.story-card h4 {
  margin: 0 0 12px;
  font-size: var(--office-text-sm);
  font-weight: 600;
}
.story-button {
  margin-top: 12px;
  min-height: 2.25rem;
  padding: 0 12px;
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  color: var(--office-foreground);
  font: inherit;
  font-size: var(--office-text-xs);
  cursor: pointer;
}
.story-button:hover {
  background: var(--office-muted);
}
.story-actions {
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
}
.thumbnail-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(min(190px, 100%), 1fr));
  gap: 16px;
}
.thumbnail-story {
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  gap: 8px;
}
.seam {
  padding: 16px;
  border: 1px dashed var(--office-border);
  border-radius: var(--office-radius-xl);
  background: color-mix(in oklab, var(--office-muted) 35%, transparent);
}
code {
  padding: 1px 4px;
  border-radius: 4px;
  background: var(--office-muted);
  font-family: var(--office-font-mono);
  font-size: 0.9em;
}
@media (max-width: 640px) {
  .file-story {
    padding: 14px;
  }
  .hero {
    gap: 12px;
  }
  .upload-grid,
  .thumbnail-grid {
    grid-template-columns: minmax(0, 1fr);
  }
}
</style>
