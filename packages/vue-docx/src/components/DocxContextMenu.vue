<template>
  <Teleport to="body">
    <div
      v-if="visible"
      class="docx-context-menu-backdrop"
      @click="close"
      @contextmenu.prevent="close"
    />
    <div
      v-if="visible"
      ref="menuRef"
      class="docx-context-menu"
      :style="menuStyle"
      @contextmenu.prevent
    >
      <!-- Clipboard actions -->
      <button
        class="docx-context-menu-item"
        @click="onAction('cut')"
      >
        <span class="docx-context-menu-icon">✂</span>
        <span>Cut</span>
        <span class="docx-context-menu-shortcut">⌘X</span>
      </button>
      <button
        class="docx-context-menu-item"
        @click="onAction('copy')"
      >
        <span class="docx-context-menu-icon">📋</span>
        <span>Copy</span>
        <span class="docx-context-menu-shortcut">⌘C</span>
      </button>
      <button
        class="docx-context-menu-item"
        @click="onAction('paste')"
      >
        <span class="docx-context-menu-icon">📄</span>
        <span>Paste</span>
        <span class="docx-context-menu-shortcut">⌘V</span>
      </button>

      <div class="docx-context-menu-separator" />

      <!-- Text formatting -->
      <button
        class="docx-context-menu-item"
        :class="{ 'docx-context-menu-item--active': selectedRunStyle?.bold }"
        @click="onAction('bold')"
      >
        <span class="docx-context-menu-icon docx-context-menu-icon--bold">B</span>
        <span>Bold</span>
        <span class="docx-context-menu-shortcut">⌘B</span>
      </button>
      <button
        class="docx-context-menu-item"
        :class="{ 'docx-context-menu-item--active': selectedRunStyle?.italic }"
        @click="onAction('italic')"
      >
        <span class="docx-context-menu-icon docx-context-menu-icon--italic">I</span>
        <span>Italic</span>
        <span class="docx-context-menu-shortcut">⌘I</span>
      </button>
      <button
        class="docx-context-menu-item"
        :class="{ 'docx-context-menu-item--active': selectedRunStyle?.underline }"
        @click="onAction('underline')"
      >
        <span class="docx-context-menu-icon docx-context-menu-icon--underline">U</span>
        <span>Underline</span>
        <span class="docx-context-menu-shortcut">⌘U</span>
      </button>
      <button
        class="docx-context-menu-item"
        :class="{ 'docx-context-menu-item--active': selectedRunStyle?.strike }"
        @click="onAction('strike')"
      >
        <span class="docx-context-menu-icon docx-context-menu-icon--strike">S</span>
        <span>Strikethrough</span>
      </button>

      <div class="docx-context-menu-separator" />

      <!-- Highlight colors -->
      <div class="docx-context-menu-section-title">Highlight</div>
      <div class="docx-context-menu-color-row">
        <button
          v-for="color of highlightColors"
          :key="color.value"
          class="docx-context-menu-color-btn"
          :style="{ backgroundColor: color.css }"
          :title="color.label"
          @click="onAction('highlight', color.value)"
        />
        <button
          class="docx-context-menu-color-btn docx-context-menu-color-btn--none"
          title="No highlight"
          @click="onAction('highlight', undefined)"
        >
          ✕
        </button>
      </div>

      <!-- Text colors -->
      <div class="docx-context-menu-section-title">Text Color</div>
      <div class="docx-context-menu-color-row">
        <button
          v-for="color of textColors"
          :key="color.value"
          class="docx-context-menu-color-btn"
          :style="{ backgroundColor: color.css }"
          :title="color.label"
          @click="onAction('textColor', color.value)"
        />
      </div>

      <div class="docx-context-menu-separator" />

      <!-- Paragraph actions -->
      <button class="docx-context-menu-item" @click="onAction('splitParagraph')">
        <span class="docx-context-menu-icon">↵</span>
        <span>Split Paragraph</span>
      </button>
      <button class="docx-context-menu-item" @click="onAction('deleteParagraph')">
        <span class="docx-context-menu-icon docx-context-menu-icon--danger">🗑</span>
        <span class="docx-context-menu-item--danger">Delete Paragraph</span>
      </button>

      <template v-if="tableContext">
        <div class="docx-context-menu-separator" />
        <div class="docx-context-menu-section-title">Table</div>
        <button class="docx-context-menu-item" @click="onAction('insertRowAbove')">
          Insert Row Above
        </button>
        <button class="docx-context-menu-item" @click="onAction('insertRowBelow')">
          Insert Row Below
        </button>
        <button class="docx-context-menu-item" @click="onAction('insertColumnLeft')">
          Insert Column Left
        </button>
        <button class="docx-context-menu-item" @click="onAction('insertColumnRight')">
          Insert Column Right
        </button>
        <button class="docx-context-menu-item" @click="onAction('deleteRow')">
          <span class="docx-context-menu-item--danger">Delete Row</span>
        </button>
        <button class="docx-context-menu-item" @click="onAction('deleteColumn')">
          <span class="docx-context-menu-item--danger">Delete Column</span>
        </button>
        <button class="docx-context-menu-item" @click="onAction('deleteTable')">
          <span class="docx-context-menu-item--danger">Delete Table</span>
        </button>
      </template>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { ref, computed, watch, onMounted, onUnmounted } from "vue"
import type {
  DocxEditorController,
  DocxContextMenuContext,
  DocxTableContextMenuContext,
} from "@extend-ai/docx-core"

// ── Props ──────────────────────────────────────────────────────────
const props = defineProps<{
  context: DocxContextMenuContext | undefined
  controller: DocxEditorController
}>()

const emit = defineEmits<{
  close: []
}>()

// ── State ──────────────────────────────────────────────────────────
const menuRef = ref<HTMLElement | undefined>()
const x = ref(0)
const y = ref(0)

// ── Computed ───────────────────────────────────────────────────────
const visible = computed(() => !!props.context)

const tableContext = computed<DocxTableContextMenuContext | undefined>(() => {
  const ctx = props.context
  if (ctx && "tableIndex" in ctx) {
    return ctx as unknown as DocxTableContextMenuContext
  }
  return undefined
})

const selectedRunStyle = computed(() => props.controller.selectedRunStyle)

const menuStyle = computed(() => {
  // Adjust position to stay within viewport
  let posX = x.value
  let posY = y.value
  if (typeof window !== "undefined") {
    const w = window.innerWidth
    const h = window.innerHeight
    const menuW = 240
    const menuH = 420
    if (posX + menuW > w) posX = Math.max(0, w - menuW - 8)
    if (posY + menuH > h) posY = Math.max(0, h - menuH - 8)
  }
  return {
    position: "fixed" as const,
    left: `${posX}px`,
    top: `${posY}px`,
    zIndex: 10000,
  }
})

// ── Colors ─────────────────────────────────────────────────────────
const highlightColors = [
  { value: "yellow", label: "Yellow", css: "#fff59d" },
  { value: "green", label: "Green", css: "#bbf7d0" },
  { value: "cyan", label: "Cyan", css: "#a5f3fc" },
  { value: "magenta", label: "Magenta", css: "#f5d0fe" },
  { value: "blue", label: "Blue", css: "#bfdbfe" },
  { value: "red", label: "Red", css: "#fecaca" },
]

const textColors = [
  { value: "#111827", label: "Black", css: "#111827" },
  { value: "#ef4444", label: "Red", css: "#ef4444" },
  { value: "#f59e0b", label: "Orange", css: "#f59e0b" },
  { value: "#22c55e", label: "Green", css: "#22c55e" },
  { value: "#3b82f6", label: "Blue", css: "#3b82f6" },
  { value: "#8b5cf6", label: "Purple", css: "#8b5cf6" },
]

// ── Actions ────────────────────────────────────────────────────────
function onAction(action: string, value?: string): void {
  const ctrl = props.controller
  switch (action) {
    case "cut": ctrl.copy(); break
    case "copy": ctrl.copy(); break
    case "paste": ctrl.paste(); break
    case "bold": ctrl.toggleBold(); break
    case "italic": ctrl.toggleItalic(); break
    case "underline": ctrl.toggleUnderline(); break
    case "strike": ctrl.toggleStrike(); break
    case "highlight": ctrl.setHighlight(value); break
    case "textColor": ctrl.setTextColor(value!); break
    case "splitParagraph": ctrl.splitParagraphAtSelection("", 0); break
    case "deleteParagraph": ctrl.deleteExpandedSelection(undefined); break
    // Table actions
    case "insertRowAbove":
      if (tableContext.value) {
        ctrl.insertTableRow(
          tableContext.value.tableIndex,
          Math.max(0, tableContext.value.rowIndex),
          "above"
        )
      }
      break
    case "insertRowBelow":
      if (tableContext.value) {
        ctrl.insertTableRow(
          tableContext.value.tableIndex,
          tableContext.value.rowIndex + 1,
          "below"
        )
      }
      break
    case "insertColumnLeft":
      if (tableContext.value) {
        ctrl.insertTableColumn(
          tableContext.value.tableIndex,
          Math.max(0, tableContext.value.cellIndex),
          "left"
        )
      }
      break
    case "insertColumnRight":
      if (tableContext.value) {
        ctrl.insertTableColumn(
          tableContext.value.tableIndex,
          tableContext.value.cellIndex + 1,
          "right"
        )
      }
      break
    case "deleteRow":
      if (tableContext.value) {
        ctrl.deleteTableRow(tableContext.value.tableIndex, tableContext.value.rowIndex)
      }
      break
    case "deleteColumn":
      if (tableContext.value) {
        ctrl.deleteTableColumn(tableContext.value.tableIndex, tableContext.value.cellIndex)
      }
      break
    case "deleteTable":
      if (tableContext.value) {
        ctrl.deleteTable(tableContext.value.tableIndex)
      }
      break
  }
  close()
}

function close(): void {
  emit("close")
}

// ── Position from context ──────────────────────────────────────────
watch(() => props.context, (ctx) => {
  // Position is set externally when showing the menu,
  // context itself doesn't carry position info
})

// Keyboard shortcut: Escape to close
function onKeydown(event: KeyboardEvent): void {
  if (event.key === "Escape" && visible.value) {
    close()
  }
}

onMounted(() => {
  window.addEventListener("keydown", onKeydown)
})

onUnmounted(() => {
  window.removeEventListener("keydown", onKeydown)
})
</script>

<style scoped>
.docx-context-menu-backdrop {
  position: fixed;
  inset: 0;
  z-index: 9999;
}
.docx-context-menu {
  min-width: 200px;
  background: #fff;
  border: 1px solid #e5e7eb;
  border-radius: 8px;
  box-shadow: 0 12px 40px rgba(0, 0, 0, 0.12);
  padding: 4px;
  font-size: 13px;
  color: #374151;
}
.docx-context-menu-item {
  display: flex;
  align-items: center;
  gap: 8px;
  width: 100%;
  padding: 6px 8px;
  border: none;
  background: none;
  border-radius: 4px;
  cursor: pointer;
  text-align: left;
  font-size: 13px;
  color: inherit;
}
.docx-context-menu-item:hover {
  background: #f3f4f6;
}
.docx-context-menu-item--active {
  background: #eff6ff;
  color: #2563eb;
}
.docx-context-menu-item--danger {
  color: #ef4444;
}
.docx-context-menu-icon {
  width: 20px;
  text-align: center;
  font-size: 12px;
}
.docx-context-menu-icon--bold {
  font-weight: 700;
}
.docx-context-menu-icon--italic {
  font-style: italic;
}
.docx-context-menu-icon--underline {
  text-decoration: underline;
}
.docx-context-menu-icon--strike {
  text-decoration: line-through;
}
.docx-context-menu-icon--danger {
  color: #ef4444;
}
.docx-context-menu-shortcut {
  margin-left: auto;
  color: #9ca3af;
  font-size: 11px;
}
.docx-context-menu-separator {
  height: 1px;
  background: #e5e7eb;
  margin: 4px 8px;
}
.docx-context-menu-section-title {
  padding: 4px 8px;
  font-size: 11px;
  font-weight: 600;
  color: #6b7280;
  text-transform: uppercase;
}
.docx-context-menu-color-row {
  display: flex;
  gap: 3px;
  padding: 4px 8px;
  flex-wrap: wrap;
}
.docx-context-menu-color-btn {
  width: 20px;
  height: 20px;
  border: 1px solid #d1d5db;
  border-radius: 3px;
  cursor: pointer;
}
.docx-context-menu-color-btn:hover {
  outline: 2px solid #3b82f6;
  outline-offset: 1px;
}
.docx-context-menu-color-btn--none {
  background: #fff;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 10px;
  color: #6b7280;
}
</style>
