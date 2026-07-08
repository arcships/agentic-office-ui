<template>
  <Teleport to="body">
    <div
      v-if="visible"
      ref="menuEl"
      class="docx-context-menu"
      :style="menuStyle"
      @click.stop
    >
      <button
        v-for="action of actions"
        :key="action.id"
        class="docx-context-menu-item"
        :class="{ 'docx-context-menu-disabled': action.disabled }"
        :disabled="action.disabled"
        @click="onAction(action)"
      >
        <span class="docx-context-menu-label">{{ action.label }}</span>
        <span v-if="action.shortcut" class="docx-context-menu-shortcut">{{ action.shortcut }}</span>
      </button>
    </div>
    <div v-if="visible" class="docx-context-menu-backdrop" @click="onClose" />
  </Teleport>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, onUnmounted } from "vue"
import type { DocxEditorController, DocxContextMenuContext, DocxContextMenuAction } from "@extend-ai/docx-core"

const props = defineProps<{
  context?: DocxContextMenuContext
  position?: { x: number; y: number }
  controller?: DocxEditorController
}>()

const emit = defineEmits<{ close: [] }>()

const visible = computed(() => !!props.context)
const menuEl = ref<HTMLElement | undefined>()

interface MenuAction {
  id: string
  label: string
  shortcut?: string
  disabled?: boolean
  onSelect?: () => void
}

const actions = computed<MenuAction[]>(() => [
  { id: "copy", label: "Copy", shortcut: "⌘C" },
  { id: "cut", label: "Cut", shortcut: "⌘X" },
  { id: "paste", label: "Paste", shortcut: "⌘V" },
  { id: "select-all", label: "Select All", shortcut: "⌘A" },
])

const menuStyle = computed(() => ({
  position: "fixed" as const,
  left: `${props.position?.x ?? 0}px`,
  top: `${props.position?.y ?? 0}px`,
  zIndex: 10000,
}))

function onAction(action: MenuAction) {
  action.onSelect?.()
  onClose()
}

function onClose() {
  emit("close")
}

function onKeydown(e: KeyboardEvent) {
  if (e.key === "Escape") onClose()
}

onMounted(() => document.addEventListener("keydown", onKeydown))
onUnmounted(() => document.removeEventListener("keydown", onKeydown))
</script>

<style scoped>
.docx-context-menu {
  background: #fff;
  border: 1px solid #e5e7eb;
  border-radius: 8px;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.12);
  padding: 4px;
  min-width: 180px;
}
.docx-context-menu-item {
  display: flex;
  align-items: center;
  gap: 8px;
  width: 100%;
  padding: 6px 12px;
  border: none;
  background: none;
  font-size: 13px;
  cursor: pointer;
  border-radius: 4px;
  text-align: left;
  color: #1f2937;
}
.docx-context-menu-item:hover:not(:disabled) {
  background: #f3f4f6;
}
.docx-context-menu-disabled {
  opacity: 0.4;
  cursor: default;
}
.docx-context-menu-label {
  flex: 1;
}
.docx-context-menu-shortcut {
  color: #9ca3af;
  font-size: 11px;
}
.docx-context-menu-backdrop {
  position: fixed;
  inset: 0;
  z-index: 9999;
}
</style>
