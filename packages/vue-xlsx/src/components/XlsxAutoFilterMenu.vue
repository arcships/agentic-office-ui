<template>
  <!--
    AutoFilter panel, rendered by the shared `OfficeDropdownMenu` (Reka): the
    header row, checkable values and the Apply/Clear rows are menu entries, so
    keyboard, focus, dismissal and scrolling come from the primitive layer.
    Value rows keep the menu open (`keepOpen`) because filtering is multi-select.
  -->
  <OfficeDropdownMenu
    :open="open"
    :items="menuItems"
    :content-test-id="'xlsx-autofilter-menu'"
    :menu-label="columnName ? m.xlsx.autoFilter.label + ': ' + columnName : m.xlsx.autoFilter.label"
    :content-class="'xlsx-autofilter'"
    :side-offset="-4"
    :restore-focus="restoreFocusTarget"
    @update:open="onOpenChange"
  >
    <template #trigger>
      <span class="xlsx-autofilter__anchor" :style="anchorStyle" aria-hidden="true" tabindex="-1" />
    </template>
  </OfficeDropdownMenu>
</template>

<script setup lang="ts">
import { computed, onMounted, ref, watch, type CSSProperties } from "vue";
import {
  OfficeDropdownMenu,
  type OfficeMenuCheckboxItem,
  type OfficeMenuEntry,
  type OfficeMenuItem,
} from "@arcships/vue-ui";
import type {
  XlsxAutoFilterRange,
  XlsxAutoFilterValueOption,
  XlsxTableSortDirection,
  XlsxViewerController,
} from "@arcships/xlsx-core";
import { useXlsxMessages } from "../i18n";

const props = defineProps<{
  controller: XlsxViewerController;
  range: XlsxAutoFilterRange;
  columnIndex: number;
  columnName?: string | null;
  x: number;
  y: number;
  /** Element that receives focus when the menu closes (the grid). */
  restoreTarget?: HTMLElement | null;
}>();

const emit = defineEmits<{ close: [] }>();

const m = useXlsxMessages();
const open = ref(false);

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

const sortDisabled = computed(() => props.controller.readOnly || props.controller.isEditingLocked);

const options = computed<XlsxAutoFilterValueOption[]>(() =>
  props.controller.getAutoFilterColumnValues(props.range, props.columnIndex)
);

const activeColumn = computed(() => {
  const state = props.controller.autoFilterStates.find(
    (candidate) => candidate.reference === props.range.reference
  );
  return state?.columns.find((column) => column.columnIndex === props.columnIndex) ?? null;
});

// Until the user picks values, every listed value is allowed. When a filter is
// already active, the checks reflect the stored criterion.
const selectedValues = ref<Set<string>>(new Set());
const includeBlank = ref(false);

function seedSelection() {
  const active = activeColumn.value;
  if (!active) {
    selectedValues.value = new Set(options.value.filter((option) => !option.blank).map((option) => option.value));
    includeBlank.value = options.value.some((option) => option.blank);
    return;
  }
  selectedValues.value = new Set(active.values);
  includeBlank.value = active.includeBlank;
}

seedSelection();

// A Worker-backed range may deliver its values after the menu mounts; reseed
// once they arrive so Apply does not start from an empty selection.
watch(options, (nextOptions) => {
  if (nextOptions.length > 0) seedSelection();
});

const allSelected = computed(() => {
  const allValues = options.value.filter((option) => !option.blank);
  const everyValue = allValues.every((option) => selectedValues.value.has(option.value));
  const blankOption = options.value.some((option) => option.blank);
  return everyValue && (!blankOption || includeBlank.value);
});

function isSelected(option: XlsxAutoFilterValueOption) {
  return option.blank ? includeBlank.value : selectedValues.value.has(option.value);
}

function toggleValue(option: XlsxAutoFilterValueOption) {
  if (option.blank) {
    includeBlank.value = !includeBlank.value;
    return;
  }
  const next = new Set(selectedValues.value);
  if (next.has(option.value)) {
    next.delete(option.value);
  } else {
    next.add(option.value);
  }
  selectedValues.value = next;
}

function toggleAll() {
  if (allSelected.value) {
    selectedValues.value = new Set();
    includeBlank.value = false;
    return;
  }
  selectedValues.value = new Set(options.value.filter((option) => !option.blank).map((option) => option.value));
  includeBlank.value = options.value.some((option) => option.blank);
}

function optionLabel(option: XlsxAutoFilterValueOption) {
  return option.blank ? m.value.xlsx.autoFilter.blanks : option.value;
}

function sort(direction: XlsxTableSortDirection) {
  props.controller.sortAutoFilter(props.range.reference, props.columnIndex, direction);
}

function apply() {
  const listed = options.value.filter((option) => !option.blank);
  const selectsEverything = listed.every((option) => selectedValues.value.has(option.value))
    && (!options.value.some((option) => option.blank) || includeBlank.value);
  props.controller.applyAutoFilterValues(
    props.range.reference,
    props.columnIndex,
    selectsEverything ? null : [...selectedValues.value],
    { includeBlank: selectsEverything ? false : includeBlank.value }
  );
}

function clear() {
  props.controller.clearAutoFilter(props.range.reference);
}

const menuItems = computed<OfficeMenuEntry[]>(() => {
  const entries: OfficeMenuEntry[] = [
    {
      type: "item",
      id: "sort-ascending",
      testId: "xlsx-autofilter-sort-ascending",
      label: m.value.xlsx.contextMenu.sortAscending,
      disabled: sortDisabled.value,
      onSelect: () => sort("ascending"),
    } satisfies OfficeMenuItem,
    {
      type: "item",
      id: "sort-descending",
      testId: "xlsx-autofilter-sort-descending",
      label: m.value.xlsx.contextMenu.sortDescending,
      disabled: sortDisabled.value,
      onSelect: () => sort("descending"),
    } satisfies OfficeMenuItem,
    { type: "separator" },
    {
      type: "label",
      id: "filter-by-value",
      testId: "xlsx-autofilter-values-label",
      label: m.value.xlsx.autoFilter.filterByValue,
    },
    {
      type: "checkbox",
      id: "select-all",
      testId: "xlsx-autofilter-select-all",
      label: m.value.xlsx.autoFilter.selectAll,
      checked: allSelected.value,
      keepOpen: true,
      onSelect: toggleAll,
    } satisfies OfficeMenuCheckboxItem,
  ];
  options.value.forEach((option, index) => {
    entries.push({
      type: "checkbox",
      id: option.blank ? "__blank" : `value-${option.value}`,
      testId: `xlsx-autofilter-value-${index}`,
      label: optionLabel(option),
      checked: isSelected(option),
      keepOpen: true,
      onSelect: () => toggleValue(option),
    } satisfies OfficeMenuCheckboxItem);
  });
  if (options.value.length === 0) {
    entries.push({
      type: "label",
      id: "no-values",
      testId: "xlsx-autofilter-empty",
      label: m.value.xlsx.autoFilter.noValues,
    });
  }
  entries.push(
    { type: "separator" },
    {
      type: "item",
      id: "clear",
      testId: "xlsx-autofilter-clear",
      label: m.value.xlsx.autoFilter.clear,
      onSelect: clear,
    } satisfies OfficeMenuItem,
    {
      type: "item",
      id: "apply",
      testId: "xlsx-autofilter-apply",
      label: m.value.xlsx.autoFilter.apply,
      onSelect: apply,
    } satisfies OfficeMenuItem,
  );
  return entries;
});

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
 * The panel is the shared menu surface; the local rule only keeps the column
 * of values readable for wide value sets.
 */
.xlsx-autofilter {
  min-inline-size: 188px;
}

.xlsx-autofilter__anchor {
  border: 0;
  margin: 0;
  padding: 0;
}
</style>
