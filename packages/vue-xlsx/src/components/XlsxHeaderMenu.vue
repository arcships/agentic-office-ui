<template>
  <!--
    Header menu: a real `OfficeDropdownMenu` anchored at the pointer. The item
    model keeps the historical `data-testid`s, the read-only guard and the
    controller commands; keyboard, typeahead, focus and dismissal come from the
    shared primitive layer.
  -->
  <OfficeDropdownMenu
    :open="open"
    :items="menuItems"
    :content-test-id="listTestId"
    :content-aria-label="isColumn ? m.xlsx.contextMenu.selectColumn : m.xlsx.contextMenu.selectRow"
    :content-class="'xlsx-headermenu'"
    :side-offset="-4"
    :restore-focus="restoreFocusTarget"
    @update:open="onOpenChange"
  >
    <template #trigger>
      <span class="xlsx-headermenu__anchor" :style="anchorStyle" aria-hidden="true" tabindex="-1" />
    </template>
  </OfficeDropdownMenu>
</template>

<script setup lang="ts">
import { computed, onMounted, ref, type CSSProperties } from "vue";
import { OfficeDropdownMenu, type OfficeMenuEntry } from "@arcships/vue-ui";
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
  /** Element that receives focus when the menu closes (the grid). */
  restoreTarget?: HTMLElement | null;
}>();

const emit = defineEmits<{ close: [] }>();

const m = useXlsxMessages();
const open = ref(false);
const isColumn = computed(() => props.kind === "column");
const listTestId = computed(() => (isColumn.value ? "xlsx-column-menu-list" : "xlsx-row-menu-list"));
onMounted(() => {
  open.value = true;
});

const anchorStyle = computed<CSSProperties>(() => ({
  blockSize: "0px",
  inlineSize: "0px",
  insetBlockStart: `${props.y}px`,
  insetInlineStart: `${props.x}px`,
  pointerEvents: "none",
  position: "fixed",
}));

const items = computed<HeaderMenuItem[]>(() => {
  const controller = props.controller;
  if (isColumn.value) {
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

const menuItems = computed<OfficeMenuEntry[]>(() => items.value.map((item) => ({
  type: "item",
  id: item.id,
  testId: item.testId,
  label: item.label,
  onSelect: () => run(item),
})));

/**
 * Reka closes the menu after the select, which funnels through
 * `onOpenChange(false)` → `close`; the handler only guards and runs the command.
 */
function run(item: HeaderMenuItem): void {
  if (props.controller.isEditingLocked === true || props.controller.readOnly === true) return;
  item.run();
}

function restoreFocusTarget(): HTMLElement | null {
  return props.restoreTarget ?? null;
}

function onOpenChange(value: boolean): void {
  open.value = value;
  if (!value) emit("close");
}
</script>

<style scoped>
/*
 * The menu surface and its rows come from `OfficeDropdownMenu`
 * (`.office-menu__content` / `.office-menu__item`); only the invisible pointer
 * anchor is local.
 */
.xlsx-headermenu__anchor {
  border: 0;
  margin: 0;
  padding: 0;
}
</style>
