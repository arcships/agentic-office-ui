<template>
  <Teleport to="body">
    <div
      v-if="visible"
      class="docx-context-menu-backdrop"
      @pointerdown="close"
      @contextmenu.prevent="close"
    />
    <div
      v-if="visible"
      ref="menuRef"
      class="docx-context-menu"
      role="menu"
      aria-orientation="vertical"
      :aria-label="m.docx.contextMenu.label"
      tabindex="-1"
      :style="menuStyle"
      @contextmenu.prevent
      @keydown="onMenuKeydown"
    >
      <!-- Clipboard actions -->
      <button
        class="docx-context-menu-item"
        type="button"
        role="menuitem"
        tabindex="-1"
        @click="onAction('cut')"
      >
        <span class="docx-context-menu-icon">✂</span>
        <span>{{ m.docx.contextMenu.cut }}</span>
        <span class="docx-context-menu-shortcut">⌘X</span>
      </button>
      <button
        class="docx-context-menu-item"
        type="button"
        role="menuitem"
        tabindex="-1"
        @click="onAction('copy')"
      >
        <span class="docx-context-menu-icon">📋</span>
        <span>{{ m.docx.contextMenu.copy }}</span>
        <span class="docx-context-menu-shortcut">⌘C</span>
      </button>
      <button
        class="docx-context-menu-item"
        type="button"
        role="menuitem"
        tabindex="-1"
        @click="onAction('paste')"
      >
        <span class="docx-context-menu-icon">📄</span>
        <span>{{ m.docx.contextMenu.paste }}</span>
        <span class="docx-context-menu-shortcut">⌘V</span>
      </button>

      <div class="docx-context-menu-separator" role="separator" />

      <!-- Text formatting -->
      <button
        class="docx-context-menu-item"
        type="button"
        role="menuitem"
        tabindex="-1"
        :class="{ 'docx-context-menu-item--active': selectedRunStyle?.bold }"
        @click="onAction('bold')"
      >
        <span class="docx-context-menu-icon docx-context-menu-icon--bold">B</span>
        <span>{{ m.docx.contextMenu.bold }}</span>
        <span class="docx-context-menu-shortcut">⌘B</span>
      </button>
      <button
        class="docx-context-menu-item"
        type="button"
        role="menuitem"
        tabindex="-1"
        :class="{ 'docx-context-menu-item--active': selectedRunStyle?.italic }"
        @click="onAction('italic')"
      >
        <span class="docx-context-menu-icon docx-context-menu-icon--italic">I</span>
        <span>{{ m.docx.contextMenu.italic }}</span>
        <span class="docx-context-menu-shortcut">⌘I</span>
      </button>
      <button
        class="docx-context-menu-item"
        type="button"
        role="menuitem"
        tabindex="-1"
        :class="{ 'docx-context-menu-item--active': selectedRunStyle?.underline }"
        @click="onAction('underline')"
      >
        <span class="docx-context-menu-icon docx-context-menu-icon--underline">U</span>
        <span>{{ m.docx.contextMenu.underline }}</span>
        <span class="docx-context-menu-shortcut">⌘U</span>
      </button>
      <button
        class="docx-context-menu-item"
        type="button"
        role="menuitem"
        tabindex="-1"
        :class="{ 'docx-context-menu-item--active': selectedRunStyle?.strike }"
        @click="onAction('strike')"
      >
        <span class="docx-context-menu-icon docx-context-menu-icon--strike">S</span>
        <span>{{ m.docx.contextMenu.strikethrough }}</span>
      </button>

      <div class="docx-context-menu-separator" role="separator" />

      <!-- Highlight colors -->
      <div class="docx-context-menu-section-title" role="presentation">{{ m.docx.contextMenu.highlight }}</div>
      <div class="docx-context-menu-color-row">
        <button
          v-for="color of highlightColors"
          :key="color.value"
          class="docx-context-menu-color-btn"
          type="button"
          role="menuitem"
          tabindex="-1"
          :aria-label="color.label"
          :style="{ backgroundColor: color.css }"
          :title="color.label"
          @click="onAction('highlight', color.value)"
        />
        <button
          class="docx-context-menu-color-btn docx-context-menu-color-btn--none"
          type="button"
          role="menuitem"
          tabindex="-1"
          :aria-label="m.docx.contextMenu.noHighlight"
          :title="m.docx.contextMenu.noHighlight"
          @click="onAction('highlight', undefined)"
        >
          ✕
        </button>
      </div>

      <!-- Text colors -->
      <div class="docx-context-menu-section-title" role="presentation">{{ m.docx.contextMenu.textColor }}</div>
      <div class="docx-context-menu-color-row">
        <button
          v-for="color of textColors"
          :key="color.value"
          class="docx-context-menu-color-btn"
          type="button"
          role="menuitem"
          tabindex="-1"
          :aria-label="color.label"
          :style="{ backgroundColor: color.css }"
          :title="color.label"
          @click="onAction('textColor', color.value)"
        />
      </div>

      <div class="docx-context-menu-separator" role="separator" />

      <!-- Paragraph actions -->
      <button class="docx-context-menu-item" type="button" role="menuitem" tabindex="-1" @click="onAction('splitParagraph')">
        <span class="docx-context-menu-icon">↵</span>
        <span>{{ m.docx.contextMenu.splitParagraph }}</span>
      </button>
      <button class="docx-context-menu-item" type="button" role="menuitem" tabindex="-1" @click="onAction('deleteParagraph')">
        <span class="docx-context-menu-icon docx-context-menu-icon--danger">🗑</span>
        <span class="docx-context-menu-item--danger">{{ m.docx.contextMenu.deleteParagraph }}</span>
      </button>

      <template v-if="tableContext">
        <div class="docx-context-menu-separator" role="separator" />
        <div class="docx-context-menu-section-title" role="presentation">{{ m.docx.contextMenu.table }}</div>
        <button class="docx-context-menu-item" type="button" role="menuitem" tabindex="-1" @click="onAction('insertRowAbove')">
          {{ m.docx.contextMenu.insertRowAbove }}
        </button>
        <button class="docx-context-menu-item" type="button" role="menuitem" tabindex="-1" @click="onAction('insertRowBelow')">
          {{ m.docx.contextMenu.insertRowBelow }}
        </button>
        <button class="docx-context-menu-item" type="button" role="menuitem" tabindex="-1" @click="onAction('insertColumnLeft')">
          {{ m.docx.contextMenu.insertColumnLeft }}
        </button>
        <button class="docx-context-menu-item" type="button" role="menuitem" tabindex="-1" @click="onAction('insertColumnRight')">
          {{ m.docx.contextMenu.insertColumnRight }}
        </button>
        <button class="docx-context-menu-item" type="button" role="menuitem" tabindex="-1" @click="onAction('deleteRow')">
          <span class="docx-context-menu-item--danger">{{ m.docx.contextMenu.deleteRow }}</span>
        </button>
        <button class="docx-context-menu-item" type="button" role="menuitem" tabindex="-1" @click="onAction('deleteColumn')">
          <span class="docx-context-menu-item--danger">{{ m.docx.contextMenu.deleteColumn }}</span>
        </button>
        <button
          data-testid="docx-context-menu-merge-cells"
          class="docx-context-menu-item"
          type="button"
          role="menuitem"
          tabindex="-1"
          :disabled="!canMergeTableCells"
          @click="onAction('mergeCells')"
        >
          Merge cells
        </button>
        <button
          data-testid="docx-context-menu-split-cell"
          class="docx-context-menu-item"
          type="button"
          role="menuitem"
          tabindex="-1"
          :disabled="!canSplitTableCell"
          @click="onAction('splitCell')"
        >
          Split cell
        </button>
        <button class="docx-context-menu-item" type="button" role="menuitem" tabindex="-1" @click="onAction('deleteTable')">
          <span class="docx-context-menu-item--danger">{{ m.docx.contextMenu.deleteTable }}</span>
        </button>
      </template>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { ref, computed, nextTick, watch } from "vue"
import type {
  DocxEditorController,
  DocxContextMenuContext,
  DocxTableContextMenuContext,
} from "@arcships/docx-core"
import { compareTextRangeBoundaries, paragraphText } from "@arcships/docx-core"
import { resolveOfficeMenuKeyCommand } from "@arcships/vue-ui"
import { useDocxMessages } from "../i18n"

// ── Props ──────────────────────────────────────────────────────────
const props = defineProps<{
  context: DocxContextMenuContext | undefined
  controller: DocxEditorController
  /** Viewport coordinates where the menu was opened. */
  x?: number
  y?: number
  /** Current DOM text of the text host under the pointer, if any. */
  text?: string
  /** Caret offset inside `text` when the menu was opened. */
  caretOffset?: number
}>()
const m = useDocxMessages()

const emit = defineEmits<{
  close: []
}>()

// ── Constants ──────────────────────────────────────────────────────
const MENU_FALLBACK_WIDTH = 240
const MENU_FALLBACK_HEIGHT = 420
const MENU_VIEWPORT_GAP = 8

// Lightweight element contract: satisfied by real DOM nodes and by the
// non-DOM node objects used by the component test renderer.
interface MenuNode {
  children?: ArrayLike<MenuNode>
  props?: Record<string, unknown>
  getAttribute?: (name: string) => string | null
  hasAttribute?: (name: string) => boolean
  setAttribute?: (name: string, value: string) => void
  focus?: (options?: { preventScroll?: boolean }) => void
}

// ── State ──────────────────────────────────────────────────────────
const menuRef = ref<HTMLElement | undefined>()
const menuWidth = ref(MENU_FALLBACK_WIDTH)
const menuHeight = ref(MENU_FALLBACK_HEIGHT)
const activeItemIndex = ref(-1)

// ── Anchor ─────────────────────────────────────────────────────────
// `x`/`y` are the editor's pointer wiring; `context.position` carries the same
// point for consumers that open the menu themselves.
const anchorX = computed(() => props.x ?? props.context?.position?.x ?? 0)
const anchorY = computed(() => props.y ?? props.context?.position?.y ?? 0)
// ── Computed ───────────────────────────────────────────────────────
const visible = computed(() => !!props.context)

const tableContext = computed<DocxTableContextMenuContext | undefined>(() => (
  props.context?.tableContext
))

const selectedTableSelection = computed(() => {
  const selection = props.controller.selection
  const context = tableContext.value
  if (!context || selection.kind !== "table-cell") return undefined
  if (selection.tableIndex !== context.tableIndex) return undefined
  if (selection.tableRuntimeKey) return undefined
  return selection
})

const selectedTableRange = computed(() => selectedTableSelection.value?.cellRange)

const canMergeTableCells = computed(() => {
  const range = selectedTableRange.value
  return Boolean(
    range &&
      (range.anchorRowIndex !== range.focusRowIndex ||
        range.anchorCellIndex !== range.focusCellIndex)
  )
})

const canSplitTableCell = computed(() => {
  const context = tableContext.value
  const selection = selectedTableSelection.value
  if (!context) return false
  const rowIndex = selection?.rowIndex ?? context.rowIndex
  const cellIndex = selection?.cellIndex ?? context.cellIndex
  const table = props.controller.model.nodes[context.tableIndex]
  if (!table || table.type !== "table") return false
  const cell = table.rows[rowIndex]?.cells[cellIndex]
  return Boolean(
    cell &&
      ((cell.style?.gridSpan ?? 1) > 1 ||
        (cell.style?.rowSpan ?? 1) > 1 ||
        cell.style?.vMergeContinuation)
  )
})

const selectedRunStyle = computed(() => props.controller.selectedRunStyle)

const menuStyle = computed(() => {
  // Clamp against the measured menu box so tall menus (the table section alone
  // can exceed 20 items) never overflow the viewport.
  let posX = anchorX.value
  let posY = anchorY.value
  if (typeof window !== "undefined") {
    const w = window.innerWidth
    const h = window.innerHeight
    if (posX + menuWidth.value + MENU_VIEWPORT_GAP > w) {
      posX = Math.max(0, w - menuWidth.value - MENU_VIEWPORT_GAP)
    }
    if (posY + menuHeight.value + MENU_VIEWPORT_GAP > h) {
      posY = Math.max(0, h - menuHeight.value - MENU_VIEWPORT_GAP)
    }
  }
  return {
    position: "fixed" as const,
    left: `${posX}px`,
    top: `${posY}px`,
    zIndex: 10000,
  }
})

// ── Colors ─────────────────────────────────────────────────────────
const highlightColors = computed(() => [
  { value: "yellow", label: m.value.docx.colors.yellow, css: "#fff59d" },
  { value: "green", label: m.value.docx.colors.green, css: "#bbf7d0" },
  { value: "cyan", label: m.value.docx.colors.cyan, css: "#a5f3fc" },
  { value: "magenta", label: m.value.docx.colors.magenta, css: "#f5d0fe" },
  { value: "blue", label: m.value.docx.colors.blue, css: "#bfdbfe" },
  { value: "red", label: m.value.docx.colors.red, css: "#fecaca" },
])

const textColors = computed(() => [
  { value: "#111827", label: m.value.docx.colors.black, css: "#111827" },
  { value: "#ef4444", label: m.value.docx.colors.red, css: "#ef4444" },
  { value: "#f59e0b", label: m.value.docx.colors.orange, css: "#f59e0b" },
  { value: "#22c55e", label: m.value.docx.colors.green, css: "#22c55e" },
  { value: "#3b82f6", label: m.value.docx.colors.blue, css: "#3b82f6" },
  { value: "#8b5cf6", label: m.value.docx.colors.purple, css: "#8b5cf6" },
])

// ── Actions ────────────────────────────────────────────────────────
function activeRange() {
  return props.controller.activeTextRange ?? props.context?.activeTextRange
}

function currentParagraphCaret(): number | undefined {
  const range = activeRange()
  if (!range || range.start.location.kind !== "paragraph") return undefined
  return range.start.offset
}

function splitParagraph(): void {
  const ctrl = props.controller
  const paragraph = ctrl.selectedParagraph
  const draftText = props.text ?? (paragraph ? paragraphText(paragraph) : "")
  const caret = props.caretOffset ?? currentParagraphCaret() ?? draftText.length
  ctrl.splitParagraphAtSelection(draftText, caret, caret)
}

function deleteParagraph(): void {
  const ctrl = props.controller
  const range = activeRange()
  if (range && compareTextRangeBoundaries(range.start, range.end) < 0) {
    ctrl.deleteExpandedSelection(range)
    return
  }

  const location = range?.start.location
  if (!location || location.kind !== "paragraph") return
  const paragraph = ctrl.model.nodes[location.nodeIndex]
  if (!paragraph || paragraph.type !== "paragraph") return

  const text = props.text ?? paragraphText(paragraph)
  const rawCaret = props.caretOffset ?? range?.start.offset ?? text.length
  const caret = Math.max(0, Math.min(Math.round(rawCaret), text.length))
  const start = { kind: "paragraph" as const, nodeIndex: location.nodeIndex }
  const nextNode = ctrl.model.nodes[location.nodeIndex + 1]

  // Prefer the paragraph-merge path: deleting to the next paragraph's start
  // removes the break instead of leaving a collapsed no-op.
  if (nextNode && nextNode.type === "paragraph") {
    ctrl.deleteExpandedSelection({
      start: { location: start, offset: caret },
      end: { location: { kind: "paragraph", nodeIndex: location.nodeIndex + 1 }, offset: 0 },
    })
    return
  }

  if (caret < text.length) {
    ctrl.deleteExpandedSelection({
      start: { location: start, offset: caret },
      end: { location: start, offset: text.length },
    })
  } else if (text.length > 0) {
    ctrl.deleteExpandedSelection({
      start: { location: start, offset: 0 },
      end: { location: start, offset: text.length },
    })
  }
}

async function onAction(action: string, value?: string): Promise<void> {
  const ctrl = props.controller
  switch (action) {
    case "cut": {
      const range = activeRange()
      if (!range || compareTextRangeBoundaries(range.start, range.end) >= 0) break
      ctrl.setActiveTextRange(range)
      await ctrl.copy()
      ctrl.deleteExpandedSelection(range)
      break
    }
    case "copy": await ctrl.copy(); break
    case "paste": await ctrl.paste(); break
    case "bold": ctrl.toggleBold(); break
    case "italic": ctrl.toggleItalic(); break
    case "underline": ctrl.toggleUnderline(); break
    case "strike": ctrl.toggleStrike(); break
    case "highlight": ctrl.setHighlight(value); break
    case "textColor": ctrl.setTextColor(value!); break
    case "splitParagraph": splitParagraph(); break
    case "deleteParagraph": deleteParagraph(); break
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
          Math.max(0, tableContext.value.rowIndex),
          "below"
        )
      }
      break
    case "insertColumnLeft":
      if (tableContext.value) {
        ctrl.insertTableColumn(
          tableContext.value.tableIndex,
          Math.max(0, tableContext.value.cellIndex),
          "left",
          tableContext.value.rowIndex
        )
      }
      break
    case "insertColumnRight":
      if (tableContext.value) {
        ctrl.insertTableColumn(
          tableContext.value.tableIndex,
          Math.max(0, tableContext.value.cellIndex),
          "right",
          tableContext.value.rowIndex
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
        ctrl.deleteTableColumn(
          tableContext.value.tableIndex,
          tableContext.value.cellIndex,
          tableContext.value.rowIndex
        )
      }
      break
    case "mergeCells":
      if (tableContext.value) {
        const range = selectedTableRange.value ?? {
          tableIndex: tableContext.value.tableIndex,
          anchorRowIndex: tableContext.value.rowIndex,
          anchorCellIndex: tableContext.value.cellIndex,
          focusRowIndex: tableContext.value.rowIndex,
          focusCellIndex: tableContext.value.cellIndex,
        }
        ctrl.mergeTableCells(tableContext.value.tableIndex, range)
      }
      break
    case "splitCell":
      if (tableContext.value) {
        const selection = selectedTableSelection.value
        ctrl.splitTableCell(
          tableContext.value.tableIndex,
          selection?.rowIndex ?? tableContext.value.rowIndex,
          selection?.cellIndex ?? tableContext.value.cellIndex
        )
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
  restoreInvokerFocus()
  emit("close")
}

// ── Focus return ───────────────────────────────────────────────────
// Closing must hand focus back to the surface that opened the menu; otherwise
// keyboard users land on <body> once the teleported menu disappears.
let invoker: HTMLElement | null = null

function captureInvoker(): void {
  if (typeof document === "undefined") return
  const active = document.activeElement as HTMLElement | null
  if (!active || typeof active.focus !== "function") return
  const menu = menuRef.value
  if (menu && typeof menu.contains === "function" && menu.contains(active)) return
  invoker = active
}

function restoreInvokerFocus(): void {
  const target = invoker
  invoker = null
  target?.focus?.({ preventScroll: true })
}

// ── Position + focus from context ──────────────────────────────────
function collectMenuItems(): MenuNode[] {
  const items: MenuNode[] = []
  const visit = (node: MenuNode | undefined): void => {
    if (!node?.children) return
    for (const child of Array.from(node.children)) {
      const role = child.getAttribute?.("role") ?? child.props?.role
      const disabled =
        child.hasAttribute?.("disabled") ?? Boolean(child.props?.disabled)
      if (role === "menuitem" && !disabled) items.push(child)
      visit(child)
    }
  }
  visit(menuRef.value as MenuNode | undefined)
  return items
}

function focusMenuItem(index: number): void {
  const items = collectMenuItems()
  if (items.length === 0) return
  const nextIndex = ((index % items.length) + items.length) % items.length
  activeItemIndex.value = nextIndex
  items.forEach((item, itemIndex) => {
    item.setAttribute?.("tabindex", itemIndex === nextIndex ? "0" : "-1")
    if (itemIndex === nextIndex) item.focus?.({ preventScroll: true })
  })
}

function measureMenu(): void {
  const menu = menuRef.value
  if (!menu) return
  if (menu.offsetWidth > 0) menuWidth.value = menu.offsetWidth
  if (menu.offsetHeight > 0) menuHeight.value = menu.offsetHeight
}

watch(
  () => props.context,
  async (context) => {
    if (!context) {
      activeItemIndex.value = -1
      return
    }
    menuWidth.value = MENU_FALLBACK_WIDTH
    menuHeight.value = MENU_FALLBACK_HEIGHT
    captureInvoker()
    await nextTick()
    measureMenu()
    focusMenuItem(0)
  },
  { immediate: true }
)

// ── Keyboard ───────────────────────────────────────────────────────
// The key semantics come from the shared Office menu keyboard so DOCX and XLSX
// menus answer Escape, Tab and the arrow keys the same way. Item collection and
// the roving tabindex stay local: the menu must work with both real DOM nodes
// and the non-DOM node objects used by the component test renderer.
function onMenuKeydown(event: KeyboardEvent): void {
  if (!visible.value) return
  const command = resolveOfficeMenuKeyCommand(
    event.key,
    collectMenuItems().length,
    activeItemIndex.value,
  )
  if (!command) return
  event.preventDefault()
  event.stopPropagation()
  if (command.type === "focus") {
    focusMenuItem(command.index)
    return
  }
  close()
}
</script>

<style scoped>
.docx-context-menu-backdrop {
  position: fixed;
  inset: 0;
  z-index: 9999;
}
.docx-context-menu {
  min-width: 200px;
  max-height: 100vh;
  overflow-y: auto;
  outline: none;
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
.docx-context-menu-item:focus-visible,
.docx-context-menu-color-btn:focus-visible {
  outline: 2px solid #3b82f6;
  outline-offset: 1px;
}
@media (pointer: coarse) {
  .docx-context-menu-item { min-height: 44px; }
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
