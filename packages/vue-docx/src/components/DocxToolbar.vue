<template>
  <div class="docx-toolbar">
    <!-- Undo/Redo -->
    <div class="docx-toolbar-group">
      <button class="docx-toolbar-btn" :disabled="!controller.canUndo" title="Undo" @click="controller.undo()">
        ↩
      </button>
      <button class="docx-toolbar-btn" :disabled="!controller.canRedo" title="Redo" @click="controller.redo()">
        ↪
      </button>
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

    <!-- Heading styles -->
    <div class="docx-toolbar-group">
      <select
        class="docx-toolbar-select"
        :value="controller.selectedParagraphStyleId ?? ''"
        @change="onStyleChange"
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

    <div class="docx-toolbar-spacer" />

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
      <button class="docx-toolbar-btn" title="Export DOCX" @click="controller.exportDocx()">
        💾
      </button>
      <button class="docx-toolbar-btn" title="New Document" @click="controller.newDocument()">
        ✨
      </button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type { DocxEditorController, HeadingLevel } from "@extend-ai/docx-core"

const props = defineProps<{
  controller: DocxEditorController
}>()

const selectedRunStyle = computed(() => props.controller?.selectedRunStyle)

function onStyleChange(event: Event) {
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

function onImport() {
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
}
.docx-toolbar-group {
  display: flex;
  align-items: center;
  gap: 1px;
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
.docx-toolbar-spacer {
  flex: 1;
}
</style>
