<template>
  <Teleport to="body">
    <div
      v-if="visible"
      class="xlsx-contextmenu-backdrop"
      @pointerdown="close"
      @contextmenu.prevent="close"
    />
    <div
      v-if="visible"
      class="xlsx-contextmenu"
      :style="menuStyle"
      role="menu"
    >
      <button class="xlsx-contextmenu__item" role="menuitem" @click="onCopy">复制</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        role="menuitem"
        @click="onPaste"
      >粘贴</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        role="menuitem"
        @click="onClear"
      >清除内容</button>
      <div class="xlsx-contextmenu__separator" />
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        role="menuitem"
        @click="onUndo"
      >撤销</button>
      <button
        v-if="!isReadOnly"
        class="xlsx-contextmenu__item"
        role="menuitem"
        @click="onRedo"
      >重做</button>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, onUnmounted, type CSSProperties } from "vue";
import type { XlsxViewerController } from "@extend-ai/xlsx-core";

const props = defineProps<{
  controller: XlsxViewerController;
}>();

const visible = ref(false);
const position = ref({ x: 0, y: 0 });

const isReadOnly = computed(() => props.controller.readOnly);

const menuStyle = computed<CSSProperties>(() => ({
  left: `${position.value.x}px`,
  position: "fixed",
  top: `${position.value.y}px`,
  zIndex: "10000",
}));

function onContextMenu(event: MouseEvent) {
  position.value = { x: event.clientX, y: event.clientY };
  visible.value = true;
}

function close() { visible.value = false; }
function onCopy() { props.controller.copySelectionToClipboard(); close(); }
function onPaste() { props.controller.pasteFromClipboard(); close(); }
function onClear() { props.controller.clearSelectedCells(); close(); }
function onUndo() { props.controller.undo(); close(); }
function onRedo() { props.controller.redo(); close(); }

let boundHandler: ((e: MouseEvent) => void) | null = null;

onMounted(() => {
  boundHandler = (e: MouseEvent) => onContextMenu(e);
  document.addEventListener("contextmenu", boundHandler);
});

onUnmounted(() => {
  if (boundHandler) document.removeEventListener("contextmenu", boundHandler);
});
</script>

<style scoped>
.xlsx-contextmenu-backdrop { bottom: 0; left: 0; position: fixed; right: 0; top: 0; z-index: 9999; }
.xlsx-contextmenu {
  background: #ffffff; border: 1px solid #d4d4d8; border-radius: 6px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.12); min-width: 160px; padding: 4px;
}
.xlsx-contextmenu__item {
  align-items: center; background: transparent; border: none; border-radius: 4px;
  color: #18181b; cursor: pointer; display: flex; font-family: inherit;
  font-size: 12px; height: 28px; padding: 0 8px; width: 100%;
}
.xlsx-contextmenu__item:hover { background: #f4f4f5; }
.xlsx-contextmenu__separator { background: #e4e4e7; height: 1px; margin: 4px 0; }
</style>
