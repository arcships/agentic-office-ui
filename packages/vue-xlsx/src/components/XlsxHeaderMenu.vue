<template>
  <Teleport to="body">
    <div class="xlsx-headermenu-backdrop" data-testid="xlsx-headermenu-backdrop" @pointerdown="emit('close')" />
    <div
      ref="menuRef"
      class="xlsx-headermenu"
      :data-testid="kind === 'column' ? 'xlsx-column-menu-list' : 'xlsx-row-menu-list'"
      :style="menuStyle"
      role="menu"
      :aria-label="kind === 'column' ? m.xlsx.contextMenu.selectColumn : m.xlsx.contextMenu.selectRow"
      @keydown="onKeydown"
    >
      <button
        v-for="item in items"
        :key="item.id"
        class="xlsx-headermenu__item"
        :data-testid="item.testId"
        role="menuitem"
        @click="run(item)"
      >{{ item.label }}</button>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { computed, nextTick, onMounted, ref, type CSSProperties } from "vue";
import { resolveOfficeMenuKeyCommand } from "@arcships/vue-ui";
import type { XlsxViewerController } from "@arcships/xlsx-core";
import { useXlsxMessages } from "../i18n";

type HeaderMenuItem = {
  id: string;
  label: string;
  run: () => void;
  testId: string;
};

const props = defineProps<{
  controller: XlsxViewerController;
  index: number;
  kind: "row" | "column";
  x: number;
  y: number;
}>();

const emit = defineEmits<{ close: [] }>();

const m = useXlsxMessages();
const menuRef = ref<HTMLElement | null>(null);
const menuStyle = computed<CSSProperties>(() => ({
  left: `${props.x}px`,
  position: "fixed",
  top: `${props.y}px`,
  zIndex: "10000"
}));

const items = computed<HeaderMenuItem[]>(() => {
  const controller = props.controller;
  if (props.kind === "column") {
    const table = controller.tables.find(
      (candidate) => props.index >= candidate.start.col && props.index <= candidate.end.col
    );
    const columnMenu: HeaderMenuItem[] = [];
    if (table) {
      columnMenu.push(
        {
          id: "sort-ascending",
          label: m.value.xlsx.contextMenu.sortAscending,
          run: () => controller.sortTable(table.name, props.index - table.start.col, "ascending"),
          testId: "xlsx-column-menu-sort-ascending"
        },
        {
          id: "sort-descending",
          label: m.value.xlsx.contextMenu.sortDescending,
          run: () => controller.sortTable(table.name, props.index - table.start.col, "descending"),
          testId: "xlsx-column-menu-sort-descending"
        }
      );
    }
    return [
      ...columnMenu,
      {
        id: "insert-left",
        label: m.value.xlsx.contextMenu.insertColumnLeft,
        run: () => controller.insertColumns?.(props.index),
        testId: "xlsx-column-menu-insert-left"
      },
      {
        id: "insert-right",
        label: m.value.xlsx.contextMenu.insertColumnRight,
        run: () => controller.insertColumns?.(props.index + 1),
        testId: "xlsx-column-menu-insert-right"
      },
      {
        id: "delete",
        label: m.value.xlsx.contextMenu.deleteColumn,
        run: () => controller.deleteColumns?.(props.index),
        testId: "xlsx-column-menu-delete"
      },
      {
        id: "select",
        label: m.value.xlsx.contextMenu.selectColumn,
        run: () => controller.selectColumn?.(props.index),
        testId: "xlsx-column-menu-select"
      },
      {
        id: "clear",
        label: m.value.xlsx.contextMenu.clearContents,
        run: () => {
          controller.selectColumn?.(props.index);
          controller.clearSelectedCells();
        },
        testId: "xlsx-column-menu-clear"
      }
    ];
  }
  return [
    {
      id: "insert-above",
      label: m.value.xlsx.contextMenu.insertRowAbove,
      run: () => controller.insertRows?.(props.index),
      testId: "xlsx-row-menu-insert-above"
    },
    {
      id: "insert-below",
      label: m.value.xlsx.contextMenu.insertRowBelow,
      run: () => controller.insertRows?.(props.index + 1),
      testId: "xlsx-row-menu-insert-below"
    },
    {
      id: "delete",
      label: m.value.xlsx.contextMenu.deleteRow,
      run: () => controller.deleteRows?.(props.index),
      testId: "xlsx-row-menu-delete"
    },
    {
      id: "select",
      label: m.value.xlsx.contextMenu.selectRow,
      run: () => controller.selectRow?.(props.index),
      testId: "xlsx-row-menu-select"
    },
    {
      id: "clear",
      label: m.value.xlsx.contextMenu.clearContents,
      run: () => {
        controller.selectRow?.(props.index);
        controller.clearSelectedCells();
      },
      testId: "xlsx-row-menu-clear"
    }
  ];
});

function run(item: HeaderMenuItem) {
  if (props.controller.isEditingLocked ?? props.controller.readOnly) {
    emit("close");
    return;
  }
  item.run();
  emit("close");
}

function focusables() {
  return menuRef.value && typeof menuRef.value.querySelectorAll === "function"
    ? Array.from(menuRef.value.querySelectorAll('[role="menuitem"]')) as HTMLElement[]
    : [];
}

function onKeydown(event: KeyboardEvent) {
  const focused = typeof document !== "undefined" ? document.activeElement as HTMLElement | null : null;
  const elements = focusables();
  const currentIndex = focused ? elements.indexOf(focused) : -1;
  if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
    const key = event.key.toLowerCase();
    for (let offset = 1; offset <= elements.length; offset += 1) {
      const candidate = elements[(currentIndex + offset + elements.length) % elements.length];
      if ((candidate?.textContent ?? "").trim().toLowerCase().startsWith(key)) {
        event.preventDefault();
        candidate?.focus?.();
        return;
      }
    }
    return;
  }
  const command = resolveOfficeMenuKeyCommand(event.key, elements.length, currentIndex);
  if (!command) return;
  event.preventDefault();
  if (command.type === "focus") {
    elements[command.index]?.focus?.();
  } else {
    emit("close");
  }
}

onMounted(() => {
  void nextTick(() => focusables()[0]?.focus?.());
});
</script>

<style scoped>
.xlsx-headermenu-backdrop { bottom: 0; left: 0; position: fixed; right: 0; top: 0; z-index: 9999; }
.xlsx-headermenu {
  background: #ffffff; border: 1px solid #d4d4d8; border-radius: 6px;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.12); min-width: 168px; padding: 4px;
}
.xlsx-headermenu__item {
  align-items: center; background: transparent; border: none; border-radius: 4px;
  color: #18181b; cursor: pointer; display: flex; font-family: inherit;
  font-size: 12px; height: 28px; padding: 0 8px; text-align: left; width: 100%;
}
.xlsx-headermenu__item:hover { background: #f4f4f5; }
@media (pointer: coarse) {
  .xlsx-headermenu__item { height: 44px; }
}
</style>
