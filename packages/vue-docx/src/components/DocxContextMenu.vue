<template>
  <Teleport to="body">
    <div
      v-if="visible"
      class="docx-context-menu-backdrop"
      @pointerdown="close"
      @contextmenu.prevent="close"
    />
    <!--
      Host-rendered menu. The editor still owns the backdrop, pointer position
      and closing, so a custom menu behaves like the built-in one.
    -->
    <div
      v-if="visible && customMenu"
      class="docx-context-menu-host"
      data-testid="docx-context-menu-host"
      :data-docx-context-menu-kind="context?.kind"
      :style="menuStyle"
      @contextmenu.prevent
      @keydown.esc.stop.prevent="close"
    >
      <HostMenu :render="customMenu.render" :render-props="customMenu.renderProps" />
    </div>
    <div
      v-if="visible && !customMenu"
      ref="menuRef"
      class="docx-context-menu"
      role="menu"
      aria-orientation="vertical"
      :aria-label="m.docx.contextMenu.label"
      :data-docx-context-menu-read-only="readOnly ? 'true' : undefined"
      tabindex="-1"
      :style="menuStyle"
      @contextmenu.prevent
      @keydown="onMenuKeydown"
    >
      <!--
        Read-only documents keep a Copy menu: selecting text must stay
        copyable even though every editing command is hidden.
      -->
      <template v-if="readOnly">
        <button
          data-testid="docx-context-menu-copy"
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
      </template>
      <template v-else>
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
            {{ m.docx.contextMenu.mergeCells }}
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
            {{ m.docx.contextMenu.splitCell }}
          </button>
          <button
            data-testid="docx-context-menu-select-table"
            class="docx-context-menu-item"
            type="button"
            role="menuitem"
            tabindex="-1"
            @click="onAction('selectTable')"
          >
            {{ m.docx.contextMenu.selectTable }}
          </button>
          <button
            data-testid="docx-context-menu-clear-cells"
            class="docx-context-menu-item"
            type="button"
            role="menuitem"
            tabindex="-1"
            :disabled="!canClearTableCells"
            @click="onAction('clearTableCells')"
          >
            <span class="docx-context-menu-item--danger">{{ m.docx.contextMenu.clearContents }}</span>
          </button>
          <button class="docx-context-menu-item" type="button" role="menuitem" tabindex="-1" @click="onAction('deleteTable')">
            <span class="docx-context-menu-item--danger">{{ m.docx.contextMenu.deleteTable }}</span>
          </button>
        </template>

        <template v-if="imageContext">
          <div class="docx-context-menu-separator" role="separator" />
          <div class="docx-context-menu-section-title" role="presentation">{{ m.docx.contextMenu.layer }}</div>
          <button
            v-for="action of imageLayerActions"
            :key="action.actionId"
            :data-testid="action.testId"
            class="docx-context-menu-item"
            type="button"
            role="menuitem"
            tabindex="-1"
            :disabled="action.disabled"
            @click="onAction(action.command)"
          >
            {{ action.label }}
          </button>
        </template>
      </template>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { computed, defineComponent, nextTick, ref, watch } from "vue"
import type { PropType, VNodeChild } from "vue"
import type {
  DocxContextMenuAction,
  DocxContextMenuActionId,
  DocxContextMenuRenderProps,
  DocxEditorController,
  DocxContextMenuContext,
  DocxImageLayerOrderCommand,
  DocxTableContextMenuAction,
  DocxTableContextMenuActionId,
  DocxTableContextMenuContext,
  DocxTableContextMenuRenderProps,
} from "@arcships/docx-core"
import {
  compareTextRangeBoundaries,
  docxImageLayerOrderCommandFromActionId,
  getParagraphAtLocation,
  paragraphText,
  resolveDocxImageLayerOrderMenuActions,
  reviewDisplayParagraph,
} from "@arcships/docx-core"
import { resolveOfficeMenuKeyCommand } from "@arcships/vue-ui"
import { useDocxMessages } from "../i18n"
import {
  effectiveControllerRevisionDisplay,
  paragraphDisplayOffsetToSourceOffset,
  paragraphSourceOffsetToDisplayOffset,
} from "../composables/editor-selection"

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
  /**
   * True when the document is read-only (or editing is locked). The built-in
   * menu then only offers Copy.
   */
  readOnly?: boolean
  /**
   * Host renderer that replaces the built-in text, image and table menu.
   * Omit to keep the default menu.
   */
  renderContextMenu?: (props: DocxContextMenuRenderProps) => unknown
  /** Host renderer for table menus; wins over `renderContextMenu`. */
  renderTableContextMenu?: (
    props: DocxTableContextMenuRenderProps
  ) => unknown
}>()
const m = useDocxMessages()

const emit = defineEmits<{
  close: []
}>()

// ── Host menu ──────────────────────────────────────────────────────
// A host renderer returns whatever it wants to show; wrapping it in a
// component keeps strings, arrays and VNodes all renderable.
type HostMenuRenderer =
  | ((props: DocxContextMenuRenderProps) => unknown)
  | ((props: DocxTableContextMenuRenderProps) => unknown)

const HostMenu = defineComponent({
  name: "DocxContextMenuHost",
  props: {
    render: {
      type: Function as PropType<HostMenuRenderer>,
      required: true,
    },
    renderProps: {
      type: Object as PropType<
        DocxContextMenuRenderProps | DocxTableContextMenuRenderProps
      >,
      required: true,
    },
  },
  setup(menu) {
    return () =>
      (menu.render as unknown as (props: unknown) => VNodeChild)(
        menu.renderProps
      )
  },
})

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

const imageContext = computed(() => (
  props.context?.kind === "image" ? props.context.image : undefined
))

function imageLayerOrderLabel(command: DocxImageLayerOrderCommand): string {
  const menu = m.value.docx.contextMenu
  switch (command) {
    case "bringToFront": return menu.bringToFront
    case "bringForward": return menu.bringForward
    case "sendBackward": return menu.sendBackward
    case "sendToBack": return menu.sendToBack
    default: return command
  }
}

// The four layer commands are the flattened leaves of
// DEFAULT_CONTEXT_MENU_IMAGE_LAYER_ACTIONS; the helper resolves their item
// order and Word's disabled rules from the image that opened the menu.
const imageLayerActions = computed(() =>
  resolveDocxImageLayerOrderMenuActions(imageContext.value?.floating).map(
    (action) => ({
      ...action,
      label: imageLayerOrderLabel(action.command),
      testId: `docx-context-menu-${action.actionId}`,
    })
  )
)

// ── Built-in actions for host menus ────────────────────────────────
// Host renderers receive the same commands the default menu shows, with ids
// that `runAction()` understands. Kebab-case ids match `DocxContextMenuActionId`;
// color entries use `highlight-<color>` and `text-color-<css color>`.
const ACTION_COMMANDS: Record<string, string> = {
  cut: "cut",
  copy: "copy",
  paste: "paste",
  bold: "bold",
  italic: "italic",
  underline: "underline",
  strike: "strike",
  "split-paragraph": "splitParagraph",
  "delete-paragraph": "deleteParagraph",
  "insert-row-above": "insertRowAbove",
  "insert-row-below": "insertRowBelow",
  "insert-column-left": "insertColumnLeft",
  "insert-column-right": "insertColumnRight",
  "delete-row": "deleteRow",
  "delete-column": "deleteColumn",
  "merge-cells": "mergeCells",
  "split-cell": "splitCell",
  "select-table": "selectTable",
  "clear-table-cells": "clearTableCells",
  "delete-table": "deleteTable",
}

const HIGHLIGHT_ACTION_PATTERN = /^highlight-(yellow|green|cyan|magenta|blue|red)$/u
const TEXT_COLOR_ACTION_PATTERN = /^text-color-(#[0-9a-f]{3,8})$/iu

function highlightColorActions(): DocxContextMenuAction[] {
  return highlightColors.value.map((color) => ({
    id: `highlight-${color.value}`,
    label: color.label,
  }))
}

function textColorActions(): DocxContextMenuAction[] {
  return textColors.value.map((color) => ({
    id: `text-color-${color.value}`,
    label: color.label,
  }))
}

const tableMenuActions = computed<DocxTableContextMenuAction[]>(() => {
  const menu = m.value.docx.contextMenu
  return [
    { id: "insert-row-above", label: menu.insertRowAbove },
    { id: "insert-row-below", label: menu.insertRowBelow },
    { id: "insert-column-left", label: menu.insertColumnLeft },
    { id: "insert-column-right", label: menu.insertColumnRight },
    { id: "delete-row", label: menu.deleteRow, destructive: true },
    { id: "delete-column", label: menu.deleteColumn, destructive: true },
    { id: "merge-cells", label: menu.mergeCells },
    { id: "split-cell", label: menu.splitCell },
    { id: "select-table", label: menu.selectTable },
    { id: "clear-table-cells", label: menu.clearContents },
    { id: "delete-table", label: menu.deleteTable, destructive: true },
  ]
})

const builtInActions = computed<DocxContextMenuAction[]>(() => {
  const menu = m.value.docx.contextMenu
  const readOnly = Boolean(props.readOnly)
  const actions: DocxContextMenuAction[] = []
  if (!readOnly) {
    actions.push({ id: "cut", label: menu.cut, shortcut: "⌘X" })
  }
  actions.push({ id: "copy", label: menu.copy, shortcut: "⌘C" })
  if (readOnly) return actions

  actions.push({ id: "paste", label: menu.paste, shortcut: "⌘V" })
  const runStyle = selectedRunStyle.value
  actions.push(
    {
      id: "bold",
      label: menu.bold,
      shortcut: "⌘B",
      separatorBefore: true,
      checked: Boolean(runStyle?.bold),
    },
    {
      id: "italic",
      label: menu.italic,
      shortcut: "⌘I",
      checked: Boolean(runStyle?.italic),
    },
    {
      id: "underline",
      label: menu.underline,
      shortcut: "⌘U",
      checked: Boolean(runStyle?.underline),
    },
    {
      id: "strike",
      label: menu.strikethrough,
      checked: Boolean(runStyle?.strike),
    },
    {
      id: "highlight",
      label: menu.highlight,
      separatorBefore: true,
      children: highlightColorActions(),
    },
    { id: "text-color", label: menu.textColor, children: textColorActions() },
    {
      id: "split-paragraph",
      label: menu.splitParagraph,
      separatorBefore: true,
    },
    { id: "delete-paragraph", label: menu.deleteParagraph, destructive: true }
  )

  if (tableContext.value) {
    actions.push(
      ...tableMenuActions.value.map((action, index) => ({
        ...action,
        separatorBefore: index === 0,
      }))
    )
  }

  if (imageContext.value) {
    actions.push(
      ...imageLayerActions.value.map((action, index) => ({
        id: action.actionId,
        label: action.label,
        disabled: action.disabled,
        separatorBefore: index === 0,
      }))
    )
  }

  return actions
})

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

const canClearTableCells = computed(() => canMergeTableCells.value)

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

interface CustomMenuEntry {
  render: HostMenuRenderer
  renderProps: DocxContextMenuRenderProps | DocxTableContextMenuRenderProps
}

// A host renderer replaces the whole menu body. The table-specific hook wins
// for table targets; other targets fall back to the general hook.
const customMenu = computed<CustomMenuEntry | undefined>(() => {
  const context = props.context
  if (!context) return undefined
  const position = { x: anchorX.value, y: anchorY.value }
  const documentTheme = props.controller.documentTheme

  if (!props.readOnly && tableContext.value && props.renderTableContextMenu) {
    return {
      render: props.renderTableContextMenu,
      renderProps: {
        context: tableContext.value,
        actions: tableMenuActions.value,
        runAction: (actionId: DocxTableContextMenuActionId) => {
          runAction(actionId)
        },
        closeMenu: close,
        position,
        documentTheme,
      } satisfies DocxTableContextMenuRenderProps,
    }
  }

  if (props.renderContextMenu) {
    return {
      render: props.renderContextMenu,
      renderProps: {
        context,
        actions: builtInActions.value,
        runAction,
        closeMenu: close,
        position,
        documentTheme,
        readOnly: Boolean(props.readOnly),
      } satisfies DocxContextMenuRenderProps,
    }
  }

  return undefined
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

function revisionMode() {
  return effectiveControllerRevisionDisplay(props.controller)
}

function activeParagraph() {
  const range = activeRange()
  if (!range) return undefined
  return getParagraphAtLocation(props.controller.model, range.start.location).paragraph
}

function currentParagraphCaret(): number | undefined {
  const range = activeRange()
  const paragraph = activeParagraph()
  if (!range || !paragraph || range.start.location.kind !== "paragraph") return undefined
  return paragraphSourceOffsetToDisplayOffset(
    paragraph,
    revisionMode(),
    range.start.offset
  )
}

function displayParagraphText(paragraph: ReturnType<typeof activeParagraph>): string {
  return paragraph
    ? paragraphText(reviewDisplayParagraph(paragraph, revisionMode()))
    : ""
}

function splitParagraph(): void {
  const ctrl = props.controller
  const paragraph = ctrl.selectedParagraph
  const draftText = props.text ?? displayParagraphText(paragraph)
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

  const sourceText = paragraphText(paragraph)
  const displayText = props.text ?? displayParagraphText(paragraph)
  const rawCaret = props.caretOffset ?? currentParagraphCaret() ?? displayText.length
  const displayCaret = Math.max(0, Math.min(Math.round(rawCaret), displayText.length))
  const caret = paragraphDisplayOffsetToSourceOffset(
    paragraph,
    revisionMode(),
    displayCaret
  )
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

  if (caret < sourceText.length) {
    ctrl.deleteExpandedSelection({
      start: { location: start, offset: caret },
      end: { location: start, offset: sourceText.length },
    })
  } else if (sourceText.length > 0) {
    ctrl.deleteExpandedSelection({
      start: { location: start, offset: 0 },
      end: { location: start, offset: sourceText.length },
    })
  }
}

/**
 * Runs a built-in action by its public id. Host menus use this instead of the
 * editor's internal command names; `closeMenu()` is called once the command
 * finishes, matching the default menu.
 */
function runAction(actionId: DocxContextMenuActionId | (string & {})): void {
  const id = String(actionId)
  const highlight = HIGHLIGHT_ACTION_PATTERN.exec(id)
  if (highlight) {
    void onAction("highlight", highlight[1])
    return
  }
  const textColor = TEXT_COLOR_ACTION_PATTERN.exec(id)
  if (textColor) {
    void onAction("textColor", textColor[1])
    return
  }
  const layerCommand = docxImageLayerOrderCommandFromActionId(id)
  if (layerCommand) {
    void onAction(layerCommand)
    return
  }
  void onAction(ACTION_COMMANDS[id] ?? id)
}

async function onAction(action: string, value?: string): Promise<void> {
  const ctrl = props.controller
  switch (action) {
    case "cut": {
      const range = activeRange()
      if (range && compareTextRangeBoundaries(range.start, range.end) < 0) {
        ctrl.setActiveTextRange(range)
        await ctrl.copy()
        ctrl.deleteExpandedSelection(range)
      } else if (selectedTableRange.value) {
        await ctrl.copy()
        ctrl.clearTableSelection()
      }
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
    case "selectTable":
      if (tableContext.value) {
        ctrl.selectWholeTable(tableContext.value.tableIndex)
      }
      break
    case "clearTableCells":
      ctrl.clearTableSelection()
      break
    case "deleteTable":
      if (tableContext.value) {
        ctrl.deleteTable(tableContext.value.tableIndex)
      }
      break
    // Image layer order actions
    case "bringToFront":
    case "bringForward":
    case "sendBackward":
    case "sendToBack":
      if (imageContext.value) {
        ctrl[action](imageContext.value.location)
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
/* The host menu owns its own visuals; this layer only anchors it at the
   pointer and lets the backdrop close it. */
.docx-context-menu-host {
  outline: none;
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
