<template>
  <Teleport to="body">
    <div class="xlsx-headermenu-backdrop" data-testid="xlsx-headermenu-backdrop" @pointerdown="emit('close')" />
    <div
      ref="menuRef"
      class="office-ui office-menu__content office-menu__viewport xlsx-headermenu"
      :data-testid="isColumn ? 'xlsx-column-menu-list' : 'xlsx-row-menu-list'"
      :style="menuStyle"
      role="menu"
      :aria-label="isColumn ? m.xlsx.contextMenu.selectColumn : m.xlsx.contextMenu.selectRow"
      @keydown="onKeydown"
    >
      <button
        v-for="item in items"
        :key="item.id"
        class="office-menu__item xlsx-headermenu__item"
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
const isColumn = computed(() => props.kind === "column")

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
/*
 * Row/column header menu: the shared `.office-menu__content` surface with the
 * imperative positioning layer. Keyboard handling already delegates to the
 * shared `resolveOfficeMenuKeyCommand`, so arrow/Escape/Tab semantics match
 * the cell context menu and the Office dropdown menus.
 */
.xlsx-headermenu-backdrop {
  inset: 0;
  position: fixed;
  z-index: 9999;
}

.xlsx-headermenu {
  min-inline-size: 168px;
  position: fixed;
}

.xlsx-headermenu__item {
  border: 0;
  cursor: pointer;
  font-family: inherit;
  inline-size: 100%;
  text-align: start;
}

.xlsx-headermenu__item:hover:not(:disabled),
.xlsx-headermenu__item:focus-visible:not(:disabled) {
  background: var(--office-accent);
  color: var(--office-accent-foreground);
  outline: none;
}

@media (pointer: coarse) {
  .xlsx-headermenu__item {
    min-block-size: 44px;
  }
}
</style>
