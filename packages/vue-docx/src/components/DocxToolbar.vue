<template>
  <div class="docx-toolbar-shell" data-testid="docx-editor-toolbar">
    <div class="docx-toolbar-topbar">
      <div class="docx-toolbar-topbar__leading">
        <button
          type="button"
          class="docx-toolbar-btn"
          :class="{ active: showThumbnails }"
          aria-label="Toggle page thumbnails"
          :aria-pressed="showThumbnails"
          title="Page thumbnails"
          @click="emit('toggleThumbnails')"
        >
          <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true">
            <rect x="3" y="4" width="18" height="16" rx="2" />
            <path d="M8 4v16M5.5 8h.01M5.5 12h.01M5.5 16h.01" />
          </svg>
        </button>
        <span class="docx-toolbar-filename" :title="controller.fileName || 'document.docx'">
          {{ controller.fileName || "document.docx" }}
        </span>
        <div class="docx-toolbar-page-control">
          <span>Page</span>
          <input
            :value="currentPage"
            aria-label="Current page"
            inputmode="numeric"
            :disabled="(totalPages ?? 0) < 1"
            @change="onPageChange"
          />
          <span>of {{ totalPages || "–" }}</span>
        </div>
      </div>

      <div class="docx-toolbar-topbar__actions">
        <label class="docx-toolbar-toggle" title="Show tracked changes">
          <input type="checkbox" :checked="controller.showTrackedChanges" @change="controller.toggleShowTrackedChanges()" />
          <span>Edits</span>
        </label>
        <label class="docx-toolbar-toggle" title="Show comments">
          <input type="checkbox" :checked="controller.showComments" @change="controller.toggleShowComments()" />
          <span>Comments</span>
        </label>
        <span v-if="readOnly" class="docx-toolbar-readonly">Read only</span>
        <button type="button" class="docx-toolbar-btn" title="Toggle document theme" @click="controller.setDocumentTheme(controller.documentTheme === 'light' ? 'dark' : 'light')">
          <svg v-if="controller.documentTheme === 'light'" class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true">
            <path d="M20.5 14.2A8.5 8.5 0 0 1 9.8 3.5 8.5 8.5 0 1 0 20.5 14.2Z" />
          </svg>
          <svg v-else class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true">
            <circle cx="12" cy="12" r="4" /><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2" />
          </svg>
        </button>
        <span class="docx-toolbar-separator" />
        <button data-testid="editor-new" type="button" class="docx-toolbar-action" title="New document" :disabled="readOnly" @click="controller.newDocument()">
          <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h9l4 4v14H6zM15 3v4h4M9 12h7M12.5 8.5v7" /></svg>
          <span>New</span>
        </button>
        <button data-testid="editor-import" type="button" class="docx-toolbar-action" title="Import DOCX" :disabled="readOnly" @click="onImport">
          <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 15V3M7.5 7.5 12 3l4.5 4.5M5 13v6h14v-6" /></svg>
          <span>Import</span>
        </button>
        <button data-testid="editor-export" type="button" class="docx-toolbar-action docx-toolbar-action--primary" title="Export DOCX" @click="controller.exportDocx()">
          <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v12M7.5 11.5 12 16l4.5-4.5M5 19h14" /></svg>
          <span>Export</span>
        </button>
      </div>
    </div>

    <fieldset class="docx-toolbar" :disabled="readOnly">
    <!-- Undo/Redo -->
    <div class="docx-toolbar-group">
      <button data-testid="editor-undo" class="docx-toolbar-btn" :disabled="!controller.canUndo" title="Undo (⌘Z)" @click="controller.undo()">
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M9 7 4 12l5 5M5 12h8a6 6 0 0 1 6 6" /></svg>
      </button>
      <button data-testid="editor-redo" class="docx-toolbar-btn" :disabled="!controller.canRedo" title="Redo (⇧⌘Z)" @click="controller.redo()">
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="m15 7 5 5-5 5M19 12h-8a6 6 0 0 0-6 6" /></svg>
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Font family -->
    <div class="docx-toolbar-group">
      <select
        class="docx-toolbar-select"
        :value="selectedRunStyle?.fontFamily ?? ''"
        @change="onFontFamilyChange"
        title="Font Family"
      >
        <option value="">Font</option>
        <option value="Calibri">Calibri</option>
        <option value="Arial">Arial</option>
        <option value="Times New Roman">Times New Roman</option>
        <option value="Georgia">Georgia</option>
        <option value="Helvetica">Helvetica</option>
        <option value="Courier New">Courier New</option>
      </select>
    </div>

    <!-- Font size -->
    <div class="docx-toolbar-group">
      <select
        class="docx-toolbar-select docx-toolbar-select-sm"
        :value="selectedRunStyle?.fontSizePt ?? ''"
        @change="onFontSizeChange"
        title="Font Size"
      >
        <option value="">Size</option>
        <option v-for="size of fontSizes" :key="size" :value="size">{{ size }}</option>
      </select>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Text style toggles -->
    <div class="docx-toolbar-group">
      <button
        data-testid="editor-bold"
        class="docx-toolbar-btn"
        :class="{ active: selectedRunStyle?.bold }"
        title="Bold (⌘B)"
        @click="controller.toggleBold()"
      >
        <strong>B</strong>
      </button>
      <button
        class="docx-toolbar-btn"
        :class="{ active: selectedRunStyle?.italic }"
        title="Italic (⌘I)"
        @click="controller.toggleItalic()"
      >
        <em>I</em>
      </button>
      <button
        class="docx-toolbar-btn"
        :class="{ active: selectedRunStyle?.underline }"
        title="Underline (⌘U)"
        @click="controller.toggleUnderline()"
      >
        <u>U</u>
      </button>
      <button
        class="docx-toolbar-btn"
        :class="{ active: selectedRunStyle?.strike }"
        title="Strikethrough"
        @click="controller.toggleStrike()"
      >
        <s>S</s>
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Superscript / Subscript -->
    <div class="docx-toolbar-group">
      <button
        class="docx-toolbar-btn"
        :class="{ active: selectedRunStyle?.verticalAlign === 'superscript' }"
        title="Superscript"
        aria-label="Superscript"
        @click="controller.toggleSuperscript()"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="m4 7 7 10M11 7 4 17M15 8c.4-1.3 1.4-2 2.8-2 1.3 0 2.2.7 2.2 1.8 0 1.6-1.7 2.3-4.8 4.2H20" /></svg>
      </button>
      <button
        class="docx-toolbar-btn"
        :class="{ active: selectedRunStyle?.verticalAlign === 'subscript' }"
        title="Subscript"
        aria-label="Subscript"
        @click="controller.toggleSubscript()"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="m4 6 7 10M11 6 4 16M15 15c.4-1.3 1.4-2 2.8-2 1.3 0 2.2.7 2.2 1.8 0 1.6-1.7 2.3-4.8 4.2H20" /></svg>
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Text color + Highlight color -->
    <div class="docx-toolbar-group">
      <div class="docx-toolbar-color-picker" title="Text Color">
        <input
          type="color"
          class="docx-toolbar-color-input"
          :value="textColorValue"
          @input="onTextColorChange"
        />
        <span class="docx-toolbar-color-label">A</span>
      </div>
      <div class="docx-toolbar-color-picker" title="Highlight Color">
        <input
          type="color"
          class="docx-toolbar-color-input"
          :value="highlightColorValue"
          @input="onHighlightColorChange"
        />
        <span class="docx-toolbar-color-label docx-toolbar-highlight-label">H</span>
      </div>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Hyperlink -->
    <div class="docx-toolbar-group">
      <button
        class="docx-toolbar-btn"
        :class="{ active: !!controller.selectedLink }"
        title="Link"
        @click="onLinkClick"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M9.5 14.5 14.5 9.5M7.3 16.7l-1 1a3.5 3.5 0 0 1-5-5l4-4a3.5 3.5 0 0 1 5 0M16.7 7.3l1-1a3.5 3.5 0 0 1 5 5l-4 4a3.5 3.5 0 0 1-5 0" /></svg>
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Alignment -->
    <div class="docx-toolbar-group">
      <button class="docx-toolbar-btn" title="Align Left" @click="controller.setAlignment('left')">
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 10h11M4 14h16M4 18h9" /></svg>
      </button>
      <button class="docx-toolbar-btn" title="Align Center" @click="controller.setAlignment('center')">
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M7 10h10M4 14h16M8 18h8" /></svg>
      </button>
      <button class="docx-toolbar-btn" title="Align Right" @click="controller.setAlignment('right')">
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M9 10h11M4 14h16M11 18h9" /></svg>
      </button>
      <button class="docx-toolbar-btn" title="Justify" @click="controller.setAlignment('justify')">
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 10h16M4 14h16M4 18h16" /></svg>
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Lists -->
    <div class="docx-toolbar-group">
      <button
        class="docx-toolbar-btn"
        :class="{ active: controller.hasUnorderedList }"
        title="Bullet List"
        aria-label="Bullet list"
        @click="controller.toggleList('unordered')"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><circle cx="5" cy="7" r="1" /><circle cx="5" cy="12" r="1" /><circle cx="5" cy="17" r="1" /><path d="M9 7h11M9 12h11M9 17h11" /></svg>
      </button>
      <button
        class="docx-toolbar-btn"
        :class="{ active: controller.hasOrderedList }"
        title="Numbered List"
        aria-label="Numbered list"
        @click="controller.toggleList('ordered')"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 5h1v4M3.5 9H6M3.5 13h2L3.5 17H6M10 7h10M10 15h10" /></svg>
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Line spacing -->
    <div class="docx-toolbar-group">
      <select
        class="docx-toolbar-select docx-toolbar-select-sm"
        :value="lineSpacingValue"
        @change="onLineSpacingChange"
        title="Line Spacing"
      >
        <option value="1">1.00</option>
        <option value="1.15">1.15</option>
        <option value="1.2">1.20</option>
        <option value="1.5">1.50</option>
        <option value="2">2.00</option>
        <option value="2.5">2.50</option>
        <option value="3">3.00</option>
      </select>
    </div>

    <!-- Heading styles -->
    <div class="docx-toolbar-group">
      <select
        class="docx-toolbar-select"
        :value="controller.selectedParagraphStyleId ?? ''"
        @change="onStyleChange"
        title="Paragraph Style"
      >
        <option value="">Normal</option>
        <option
          v-for="style of controller.availableParagraphStyles"
          :key="style.id"
          :value="style.id"
        >
          {{ style.name }}
        </option>
      </select>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Borders -->
    <div class="docx-toolbar-group">
      <select
        class="docx-toolbar-select"
        :value="activeBorderPresetValue"
        @change="onBorderPresetChange"
        title="Borders"
      >
        <option value="">Borders</option>
        <option v-for="preset of borderPresets" :key="preset.value" :value="preset.value">
          {{ preset.label }}
        </option>
      </select>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Insert image + Insert table -->
    <div class="docx-toolbar-group">
      <button class="docx-toolbar-btn" title="Insert Image" @click="onInsertImage">
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="2" /><circle cx="9" cy="10" r="2" /><path d="m5 18 5-5 3 3 2-2 4 4" /></svg>
      </button>
      <button class="docx-toolbar-btn" title="Insert Table" @click="controller.insertTable()">
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="1" /><path d="M3 10h18M3 15h18M9 4v16M15 4v16" /></svg>
      </button>
    </div>

    <div class="docx-toolbar-spacer" />

    <!-- Zoom -->
    <div class="docx-toolbar-group">
      <select
        class="docx-toolbar-select docx-toolbar-select-sm"
        :value="zoomPercent"
        @change="onZoomChange"
        title="Zoom"
      >
        <option v-for="z of zoomLevels" :key="z" :value="z">{{ z }}%</option>
      </select>
    </div>

    </fieldset>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, watch } from "vue"
import type { DocxEditorController, DocxBorderPreset, HeadingLevel } from "@arcships/docx-core"

const props = defineProps<{
  controller: DocxEditorController
  zoom?: number
  readOnly?: boolean
  showThumbnails?: boolean
  currentPage?: number
  totalPages?: number
}>()

const emit = defineEmits<{
  (e: "update:zoom", value: number): void
  (e: "toggleThumbnails"): void
  (e: "selectPage", page: number): void
}>()

// ── Font sizes ──────────────────────────────────────────────────────
const fontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48]

// ── Zoom ────────────────────────────────────────────────────────────
const zoomLevels = [50, 60, 70, 80, 90, 100, 110, 120, 130, 150, 175, 200]
const zoomPercent = ref(props.zoom ?? 100)

watch(() => props.zoom, (v) => { zoomPercent.value = v ?? 100 })

function onZoomChange(event: Event): void {
  const target = event.target as HTMLSelectElement
  const value = Number(target.value)
  zoomPercent.value = value
  emit("update:zoom", value)
}

function onPageChange(event: Event): void {
  const value = Number((event.target as HTMLInputElement).value)
  const total = Math.max(1, props.totalPages ?? 1)
  if (!Number.isInteger(value)) return
  emit("selectPage", Math.min(Math.max(value, 1), total))
}

// ── Derived values ──────────────────────────────────────────────────
const selectedRunStyle = computed(() => props.controller.selectedRunStyle)

const textColorValue = computed(() => {
  const color = selectedRunStyle.value?.color
  if (color && /^#[0-9a-fA-F]{6}$/.test(color)) return color
  return "#000000"
})

const highlightColorValue = computed(() => {
  const highlight = selectedRunStyle.value?.highlight
  if (highlight && /^#[0-9a-fA-F]{6}$/.test(highlight)) return highlight
  return "#ffff00"
})

const lineSpacingValue = computed(() => {
  const ls = props.controller.selectedLineSpacing
  const m = ls.multiple
  const presets = [1, 1.15, 1.2, 1.5, 2, 2.5, 3]
  const closest = presets.reduce((prev, curr) =>
    Math.abs(curr - m) < Math.abs(prev - m) ? curr : prev
  )
  return String(closest)
})

// ── Border presets ──────────────────────────────────────────────────
const borderPresets: { value: DocxBorderPreset; label: string }[] = [
  { value: "bottom", label: "Bottom Border" },
  { value: "top", label: "Top Border" },
  { value: "left", label: "Left Border" },
  { value: "right", label: "Right Border" },
  { value: "none", label: "No Border" },
  { value: "all", label: "All Borders" },
  { value: "outside", label: "Outside Borders" },
  { value: "inside", label: "Inside Borders" },
  { value: "inside-horizontal", label: "Inside Horizontal" },
  { value: "inside-vertical", label: "Inside Vertical" },
  { value: "diagonal-down", label: "Diagonal Down" },
  { value: "diagonal-up", label: "Diagonal Up" },
  { value: "horizontal-line", label: "Horizontal Line" },
]

const activeBorderPresetValue = computed(() => {
  const presets = props.controller.activeBorderPresets
  for (const p of borderPresets) {
    if (presets[p.value]) return p.value
  }
  return ""
})

// ── Event handlers ──────────────────────────────────────────────────
function onFontFamilyChange(event: Event): void {
  const target = event.target as HTMLSelectElement
  if (target.value) props.controller.setFontFamily(target.value)
}

function onFontSizeChange(event: Event): void {
  const target = event.target as HTMLSelectElement
  const value = Number(target.value)
  if (value) props.controller.setFontSize(value)
}

function onTextColorChange(event: Event): void {
  const target = event.target as HTMLInputElement
  props.controller.setTextColor(target.value || undefined)
}

function onHighlightColorChange(event: Event): void {
  const target = event.target as HTMLInputElement
  props.controller.setHighlight(target.value || undefined)
}

function onLinkClick(): void {
  const existing = props.controller.selectedLink
  if (existing) {
    if (confirm(`Remove link "${existing}"?`)) {
      props.controller.setLink(undefined)
    }
    return
  }
  const url = window.prompt("Enter URL:", "https://")
  if (url) props.controller.setLink(url)
}

function onLineSpacingChange(event: Event): void {
  const target = event.target as HTMLSelectElement
  props.controller.setLineSpacing(Number(target.value))
}

function onBorderPresetChange(event: Event): void {
  const target = event.target as HTMLSelectElement
  if (target.value) props.controller.applyBorderPreset(target.value as DocxBorderPreset)
}

function onStyleChange(event: Event): void {
  const target = event.target as HTMLSelectElement
  const value = target.value || undefined
  if (value && value.startsWith("Heading")) {
    const level = parseInt(value.replace("Heading ", ""))
    if (level >= 1 && level <= 6) {
      props.controller.setHeading(level as HeadingLevel)
    } else {
      props.controller.setParagraphStyle(value)
    }
  } else {
    props.controller.setParagraphStyle(value)
  }
}

function onInsertImage(): void {
  const input = document.createElement("input")
  input.type = "file"
  input.accept = "image/*"
  input.onchange = async () => {
    const file = input.files?.[0]
    if (file) {
      await props.controller.insertImageFile(file)
    }
  }
  input.click()
}

function onImport(): void {
  const input = document.createElement("input")
  input.type = "file"
  input.accept = ".docx"
  input.onchange = async () => {
    const file = input.files?.[0]
    if (file) {
      await props.controller.importDocxFile(file)
    }
  }
  input.click()
}
</script>

<style scoped>
.docx-toolbar-shell {
  background: #fff;
  border-bottom: 1px solid #e4e4e7;
  color: #18181b;
  flex: 0 0 auto;
  min-width: 0;
}
.docx-toolbar-topbar {
  align-items: center;
  border-bottom: 1px solid #e4e4e7;
  display: flex;
  gap: 12px;
  justify-content: space-between;
  min-height: 44px;
  padding: 6px 12px;
}
.docx-toolbar-topbar__leading,
.docx-toolbar-topbar__actions,
.docx-toolbar-page-control {
  align-items: center;
  display: flex;
}
.docx-toolbar-topbar__leading { gap: 9px; min-width: 0; }
.docx-toolbar-topbar__actions { gap: 4px; }
.docx-toolbar-filename {
  font-size: 13px;
  font-weight: 600;
  max-width: 220px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.docx-toolbar-page-control {
  color: #52525b;
  font-size: 12px;
  gap: 5px;
  white-space: nowrap;
}
.docx-toolbar-page-control input {
  background: #fff;
  border: 1px solid #d4d4d8;
  border-radius: 6px;
  color: #18181b;
  font: inherit;
  height: 28px;
  padding: 0 4px;
  text-align: center;
  width: 38px;
}
.docx-toolbar-action {
  align-items: center;
  background: #fff;
  border: 1px solid #d4d4d8;
  border-radius: 7px;
  color: #27272a;
  cursor: pointer;
  display: inline-flex;
  font: inherit;
  font-size: 12px;
  font-weight: 500;
  gap: 6px;
  height: 32px;
  padding: 0 10px;
}
.docx-toolbar-action:hover { background: #f4f4f5; }
.docx-toolbar-action--primary { background: #18181b; border-color: #18181b; color: #fff; }
.docx-toolbar-action--primary:hover { background: #27272a; }
.docx-toolbar-icon {
  fill: none;
  height: 16px;
  stroke: currentColor;
  stroke-linecap: round;
  stroke-linejoin: round;
  stroke-width: 1.7;
  width: 16px;
}
.docx-toolbar-readonly {
  background: #f4f4f5;
  border-radius: 999px;
  color: #71717a;
  font-size: 11px;
  font-weight: 600;
  padding: 4px 8px;
}
.docx-toolbar {
  display: flex;
  align-items: center;
  gap: 2px;
  padding: 6px 12px;
  background: #fff;
  border: 0;
  border-bottom: 0;
  overflow-x: visible;
  flex-shrink: 0;
  flex-wrap: wrap;
  row-gap: 4px;
  margin: 0;
  min-inline-size: 0;
}
.docx-toolbar-group {
  display: flex;
  align-items: center;
  gap: 1px;
  flex-shrink: 0;
}
.docx-toolbar-btn {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 32px;
  height: 32px;
  border: none;
  background: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 13px;
  color: #374151;
}
.docx-toolbar-btn:hover {
  background: #f3f4f6;
}
.docx-toolbar-btn.active {
  background: #e5e7eb;
  color: #2563eb;
}
.docx-toolbar-btn:disabled {
  opacity: 0.3;
  cursor: default;
}
.docx-toolbar-btn:focus-visible,
.docx-toolbar-action:focus-visible,
.docx-toolbar-select:focus-visible,
.docx-toolbar-page-control input:focus-visible {
  outline: 2px solid #2563eb;
  outline-offset: 2px;
}
.docx-toolbar-separator {
  width: 1px;
  height: 24px;
  background: #e5e7eb;
  margin: 0 4px;
  flex-shrink: 0;
}
.docx-toolbar-select {
  height: 32px;
  padding: 0 8px;
  border: 1px solid #e5e7eb;
  border-radius: 4px;
  font-size: 13px;
  background: #fff;
  cursor: pointer;
}
.docx-toolbar-select-sm {
  width: 60px;
  padding: 0 4px;
}
.docx-toolbar-spacer {
  flex: 1;
}
.docx-toolbar-color-picker {
  position: relative;
  width: 32px;
  height: 32px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 4px;
  cursor: pointer;
}
.docx-toolbar-color-picker:hover {
  background: #f3f4f6;
}
.docx-toolbar-color-input {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  opacity: 0;
  cursor: pointer;
}
.docx-toolbar-color-label {
  font-size: 14px;
  font-weight: 700;
  color: #374151;
  pointer-events: none;
}
.docx-toolbar-highlight-label {
  font-size: 12px;
}
.docx-toolbar-toggle {
  display: flex;
  align-items: center;
  gap: 4px;
  font-size: 12px;
  cursor: pointer;
  user-select: none;
  padding: 2px 6px;
  border-radius: 4px;
}
.docx-toolbar-toggle input {
  appearance: none;
  background: #d4d4d8;
  border-radius: 999px;
  cursor: pointer;
  height: 16px;
  margin: 0;
  position: relative;
  transition: background .15s ease;
  width: 28px;
}
.docx-toolbar-toggle input::after {
  background: #fff;
  border-radius: 50%;
  box-shadow: 0 1px 2px rgba(0, 0, 0, .2);
  content: "";
  height: 12px;
  left: 2px;
  position: absolute;
  top: 2px;
  transition: transform .15s ease;
  width: 12px;
}
.docx-toolbar-toggle input:checked { background: #2563eb; }
.docx-toolbar-toggle input:checked::after { transform: translateX(12px); }
.docx-toolbar-toggle:hover {
  background: #f3f4f6;
}
.docx-toolbar-toggle-label {
  color: #6b7280;
  font-weight: 500;
}

@media (max-width: 900px) {
  .docx-toolbar-topbar { align-items: stretch; flex-direction: column; gap: 5px; }
  .docx-toolbar-topbar__actions { justify-content: flex-end; overflow-x: auto; }
  .docx-toolbar-filename { flex: 1; max-width: none; }
}

@media (max-width: 560px) {
  .docx-toolbar-topbar__actions .docx-toolbar-toggle { display: none; }
  .docx-toolbar-action span { display: none; }
  .docx-toolbar-action { padding: 0; width: 32px; }
}
</style>
