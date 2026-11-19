<template>
  <Teleport to="body">
    <div class="xlsx-autofilter-backdrop" data-testid="xlsx-autofilter-backdrop" @pointerdown="emit('close')" />
    <div
      ref="menuRef"
      class="office-ui office-popover__content xlsx-autofilter"
      data-testid="xlsx-autofilter-menu"
      :style="menuStyle"
      role="dialog"
      :aria-label="dialogLabel"
      @keydown="onKeydown"
    >
      <button
        class="office-menu__item xlsx-autofilter__item"
        data-testid="xlsx-autofilter-sort-ascending"
        :disabled="sortDisabled"
        type="button"
        @click="sort('ascending')"
      >{{ m.xlsx.contextMenu.sortAscending }}</button>
      <button
        class="office-menu__item xlsx-autofilter__item"
        data-testid="xlsx-autofilter-sort-descending"
        :disabled="sortDisabled"
        type="button"
        @click="sort('descending')"
      >{{ m.xlsx.contextMenu.sortDescending }}</button>
      <div class="office-menu__separator xlsx-autofilter__separator" />
      <div class="office-menu__label xlsx-autofilter__section" role="group" :aria-label="m.xlsx.autoFilter.filterByValue">
        {{ m.xlsx.autoFilter.filterByValue }}
      </div>
      <label class="xlsx-autofilter__check xlsx-autofilter__check--all">
        <input
          :checked="allSelected"
          data-testid="xlsx-autofilter-select-all"
          type="checkbox"
          @change="toggleAll"
        />
        <span>{{ m.xlsx.autoFilter.selectAll }}</span>
      </label>
      <div class="xlsx-autofilter__values" data-testid="xlsx-autofilter-values">
        <label
          v-for="(option, index) in options"
          :key="option.blank ? '__blank' : option.value"
          class="xlsx-autofilter__check"
        >
          <input
            :checked="isSelected(option)"
            :data-testid="`xlsx-autofilter-value-${index}`"
            type="checkbox"
            @change="toggleValue(option)"
          />
          <span :class="{ 'xlsx-autofilter__blank': option.blank }">{{ optionLabel(option) }}</span>
        </label>
        <span v-if="options.length === 0" class="xlsx-autofilter__empty">
          {{ m.xlsx.autoFilter.noValues }}
        </span>
      </div>
      <div class="office-menu__separator xlsx-autofilter__separator" />
      <div class="xlsx-autofilter__actions">
        <button
          class="office-button office-button--outline office-button--size-sm xlsx-autofilter__button"
          data-testid="xlsx-autofilter-clear"
          type="button"
          @click="clear"
        >{{ m.xlsx.autoFilter.clear }}</button>
        <button
          class="office-button office-button--default office-button--size-sm xlsx-autofilter__button xlsx-autofilter__button--primary"
          data-testid="xlsx-autofilter-apply"
          type="button"
          @click="apply"
        >{{ m.xlsx.autoFilter.apply }}</button>
      </div>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { computed, nextTick, onMounted, ref, watch, type CSSProperties } from "vue";
import type {
  XlsxAutoFilterRange,
  XlsxAutoFilterValueOption,
  XlsxTableSortDirection,
  XlsxViewerController
} from "@arcships/xlsx-core";
import { useXlsxMessages } from "../i18n";

const props = defineProps<{
  columnIndex: number;
  columnName: string;
  controller: XlsxViewerController;
  range: XlsxAutoFilterRange;
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

const dialogLabel = computed(() => props.columnName
  ? `${m.value.xlsx.autoFilter.label}: ${props.columnName}`
  : m.value.xlsx.autoFilter.label
);

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

function toggleAll(event: Event) {
  const checked = (event.target as HTMLInputElement | null)?.checked ?? !allSelected.value;
  if (checked) {
    selectedValues.value = new Set(options.value.filter((option) => !option.blank).map((option) => option.value));
    includeBlank.value = options.value.some((option) => option.blank);
    return;
  }
  selectedValues.value = new Set();
  includeBlank.value = false;
}

function optionLabel(option: XlsxAutoFilterValueOption) {
  return option.blank ? m.value.xlsx.autoFilter.blanks : option.value;
}

function sort(direction: XlsxTableSortDirection) {
  props.controller.sortAutoFilter(props.range.reference, props.columnIndex, direction);
  emit("close");
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
  emit("close");
}

function clear() {
  props.controller.clearAutoFilter(props.range.reference);
  emit("close");
}

function onKeydown(event: KeyboardEvent) {
  if (event.key === "Escape") {
    event.preventDefault();
    emit("close");
  }
}

onMounted(() => {
  void nextTick(() => {
    const menu = menuRef.value;
    if (!menu || typeof menu.querySelector !== "function") return;
    const focusTarget = menu.querySelector<HTMLElement>("button, input");
    focusTarget?.focus?.();
  });
});
</script>

<style scoped>
/*
 * The panel is the shared `.office-popover__content` surface; rows reuse the
 * `.office-menu__*` geometry so the AutoFilter popover and the context menus
 * read as one design-system family. Only the value list and the checkbox rows
 * need local rules.
 */
.xlsx-autofilter-backdrop {
  inset: 0;
  position: fixed;
  z-index: 9999;
}

.xlsx-autofilter {
  display: flex;
  flex-direction: column;
  max-block-size: 320px;
  min-inline-size: 188px;
  padding: var(--office-ui-space-1);
  position: fixed;
}

.xlsx-autofilter__item {
  border: 0;
  cursor: pointer;
  font-family: inherit;
  inline-size: 100%;
  text-align: start;
}

.xlsx-autofilter__item:hover:not(:disabled),
.xlsx-autofilter__button:hover:not(:disabled) {
  background: var(--office-accent);
  color: var(--office-accent-foreground);
}

.xlsx-autofilter__item:disabled {
  cursor: default;
  opacity: 0.64;
}

.xlsx-autofilter__item:disabled:hover {
  background: transparent;
}

.xlsx-autofilter__button {
  font-size: var(--office-text-xs);
}

.xlsx-autofilter__section {
  padding-block: 2px;
}

.xlsx-autofilter__values {
  display: flex;
  flex-direction: column;
  max-block-size: 160px;
  overflow-y: auto;
  padding-inline: var(--office-ui-space-1);
}

.xlsx-autofilter__check {
  align-items: center;
  border-radius: var(--office-radius-sm);
  display: flex;
  font-size: var(--office-text-xs);
  gap: var(--office-ui-space-1-5);
  min-block-size: 24px;
  padding-inline: var(--office-ui-space-1);
  white-space: nowrap;
}

.xlsx-autofilter__check:hover {
  background: var(--office-accent);
}

.xlsx-autofilter__check input {
  accent-color: var(--office-primary);
  margin: 0;
}

.xlsx-autofilter__check--all {
  padding-inline: var(--office-ui-space-2);
}

.xlsx-autofilter__blank,
.xlsx-autofilter__empty {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
}

.xlsx-autofilter__blank {
  font-style: italic;
}

.xlsx-autofilter__empty {
  padding: var(--office-ui-space-1);
}

.xlsx-autofilter__actions {
  display: flex;
  gap: var(--office-ui-space-1);
  justify-content: flex-end;
  padding-block-start: var(--office-ui-space-1);
}

@media (pointer: coarse) {
  .xlsx-autofilter__item,
  .xlsx-autofilter__button {
    min-block-size: 44px;
  }

  .xlsx-autofilter__check {
    min-block-size: 36px;
  }
}
</style>
