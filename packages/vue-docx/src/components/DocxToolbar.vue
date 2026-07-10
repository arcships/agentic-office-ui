<template>
  <div class="docx-toolbar">
    <!-- Undo/Redo -->
    <div class="docx-toolbar-group">
      <button data-testid="editor-undo" class="docx-toolbar-btn" :disabled="!controller.canUndo" title="Undo (⌘Z)" @click="controller.undo()">
        ↩
      </button>
      <button data-testid="editor-redo" class="docx-toolbar-btn" :disabled="!controller.canRedo" title="Redo (⇧⌘Z)" @click="controller.redo()">
        ↪
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
        @click="controller.toggleSuperscript()"
      >
        x²
      </button>
      <button
        class="docx-toolbar-btn"
        :class="{ active: selectedRunStyle?.verticalAlign === 'subscript' }"
        title="Subscript"
        @click="controller.toggleSubscript()"
      >
        x₂
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
        <span class="docx-toolbar-color-label docx-toolbar-highlight-label">🖌</span>
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
        🔗
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Alignment -->
    <div class="docx-toolbar-group">
      <button class="docx-toolbar-btn" title="Align Left" @click="controller.setAlignment('left')">
        ⬅
      </button>
      <button class="docx-toolbar-btn" title="Align Center" @click="controller.setAlignment('center')">
        ⬌
      </button>
      <button class="docx-toolbar-btn" title="Align Right" @click="controller.setAlignment('right')">
        ➡
      </button>
      <button class="docx-toolbar-btn" title="Justify" @click="controller.setAlignment('justify')">
        ☰
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Lists -->
    <div class="docx-toolbar-group">
      <button
        class="docx-toolbar-btn"
        :class="{ active: controller.hasUnorderedList }"
        title="Bullet List"
        @click="controller.toggleList('unordered')"
      >
        •
      </button>
      <button
        class="docx-toolbar-btn"
        :class="{ active: controller.hasOrderedList }"
        title="Numbered List"
        @click="controller.toggleList('ordered')"
      >
        1.
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
        🖼
      </button>
      <button class="docx-toolbar-btn" title="Insert Table" @click="controller.insertTable()">
        ⊞
      </button>
    </div>

    <div class="docx-toolbar-spacer" />

    <!-- Show edits / Show comments toggles -->
    <div class="docx-toolbar-group">
      <label class="docx-toolbar-toggle" title="Show Track Changes">
        <input
          type="checkbox"
          :checked="controller.showTrackedChanges"
          @change="controller.toggleShowTrackedChanges()"
        />
        <span class="docx-toolbar-toggle-label">TC</span>
      </label>
      <label class="docx-toolbar-toggle" title="Show Comments">
        <input
          type="checkbox"
          :checked="controller.showComments"
          @change="controller.toggleShowComments()"
        />
        <span class="docx-toolbar-toggle-label">💬</span>
      </label>
    </div>

    <div class="docx-toolbar-separator" />

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

    <!-- Columns toggle -->
    <div class="docx-toolbar-group">
      <button
        class="docx-toolbar-btn"
        :class="{ active: showColumns }"
        title="Columns"
        @click="showColumns = !showColumns"
      >
        ▥
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Document theme -->
    <div class="docx-toolbar-group">
      <button
        class="docx-toolbar-btn"
        title="Toggle Theme"
        @click="controller.setDocumentTheme(controller.documentTheme === 'light' ? 'dark' : 'light')"
      >
        {{ controller.documentTheme === 'light' ? '🌙' : '☀️' }}
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- File operations -->
    <div class="docx-toolbar-group">
      <button class="docx-toolbar-btn" title="Import DOCX" @click="onImport">
        📂
      </button>
      <button data-testid="editor-export-toolbar" class="docx-toolbar-btn" title="Export DOCX" @click="controller.exportDocx()">
        💾
      </button>
      <button class="docx-toolbar-btn" title="New Document" @click="controller.newDocument()">
        ✨
      </button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, watch } from "vue"
import type { DocxEditorController, DocxBorderPreset, HeadingLevel } from "@extend-ai/docx-core"

const props = defineProps<{
  controller: DocxEditorController
  zoom?: number
}>()

const emit = defineEmits<{
  (e: "update:zoom", value: number): void
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

// ── Columns placeholder ─────────────────────────────────────────────
const showColumns = ref(false)

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
.docx-toolbar {
  display: flex;
  align-items: center;
  gap: 2px;
  padding: 6px 12px;
  background: #fff;
  border-bottom: 1px solid #e5e7eb;
  overflow-x: auto;
  flex-shrink: 0;
  flex-wrap: nowrap;
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
.docx-toolbar-toggle:hover {
  background: #f3f4f6;
}
.docx-toolbar-toggle-label {
  color: #6b7280;
  font-weight: 500;
}
</style>
