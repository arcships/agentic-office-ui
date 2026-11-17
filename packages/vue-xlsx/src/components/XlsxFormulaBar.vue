<template>
  <div class="xlsx-formula-bar" :style="barStyle">
    <input
      class="xlsx-formula-bar__name-box"
      data-testid="xlsx-name-box"
      :value="nameBoxValue"
      readonly
      title="Active cell"
    />
    <span class="xlsx-formula-bar__fx">fx</span>
    <input
      ref="formulaInputRef"
      class="xlsx-formula-bar__formula-input"
      data-testid="xlsx-formula-input"
      :disabled="!hasFormulaTarget || isReadOnly"
      :value="formulaDraft"
      placeholder="Enter a formula or value"
      @blur="commitFormula(); focusedField = null"
      @focus="onFormulaFocus"
      @input="onFormulaInput"
      @keydown="onFormulaKeydown"
      @compositionstart="onFormulaCompositionStart"
      @compositionend="onFormulaCompositionEnd"
    />
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, ref, watch, type CSSProperties } from "vue";
import type { XlsxViewerController } from "@arcships/xlsx-core";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
  readOnly?: boolean;
}>();

const formulaInputRef = ref<HTMLInputElement | null>(null);
const formulaDraft = ref("");
const formulaInitialValue = ref("");
const focusedField = ref<"formula" | null>(null);
const isFormulaComposing = ref(false);
let commitFormulaOnCompositionEnd = false;

const isReadOnly = computed(() => props.readOnly === true || props.controller.isEditingLocked === true || (props.readOnly === undefined && props.controller.readOnly));

const hasFormulaTarget = computed(() => {
  const addr = props.controller.activeCellAddress;
  const sf = props.controller.selectedFormulaTarget;
  return !!addr || sf?.kind === "chartSeries";
});

const nameBoxValue = computed(() => {
  const sf = props.controller.selectedFormulaTarget;
  if (sf?.kind === "chartSeries") {
    return `SERIES ${sf.seriesIndex + 1}`;
  }
  return props.controller.activeCellAddress ?? "";
});

const barStyle = computed<CSSProperties>(() => ({
  alignItems: "center",
  backgroundColor: props.isDark ? "#27272a" : "#f8f9fa",
  borderBottom: `1px solid ${props.isDark ? "#3f3f46" : "#e4e4e7"}`,
  display: "flex",
  flexShrink: "0",
  gap: "1px",
  padding: "4px 8px",
}));

// Sync formula draft from controller when not focused
watch(
  () => [props.controller.selectedFormula, props.controller.selectedValue, props.controller.activeCellAddress],
  () => {
    if (focusedField.value === "formula") return;
    formulaDraft.value = props.controller.selectedFormula || props.controller.selectedValue;
  },
  { immediate: true }
);

function onFormulaFocus() {
  formulaInitialValue.value = formulaDraft.value;
  focusedField.value = "formula";
}

function onFormulaInput(event: Event) {
  const target = event.target as HTMLInputElement;
  formulaDraft.value = target.value;
}

function commitFormula(nextFormula?: string) {
  const resolvedFormula = nextFormula ?? formulaDraft.value;
  if (!hasFormulaTarget.value) return;
  if (resolvedFormula === formulaInitialValue.value) return;
  props.controller.setSelectedFormula(resolvedFormula);
}

function onFormulaKeydown(event: KeyboardEvent) {
  if (isFormulaComposing.value || event.isComposing || event.keyCode === 229) {
    // Let the IME consume the key. Enter confirms the composition; remember it
    // so the formula is committed once the final composition text is in place.
    if (event.key === "Enter") commitFormulaOnCompositionEnd = true;
    return;
  }

  if (event.key === "Enter") {
    event.preventDefault();
    commitFormula();
    focusedField.value = null;
    formulaInputRef.value?.blur();
  }
}

function onFormulaCompositionStart() {
  isFormulaComposing.value = true;
  commitFormulaOnCompositionEnd = false;
}

function onFormulaCompositionEnd() {
  isFormulaComposing.value = false;
  if (!commitFormulaOnCompositionEnd) return;
  commitFormulaOnCompositionEnd = false;
  // The final composition text lands via @input before this tick completes.
  void nextTick(() => {
    commitFormula();
    focusedField.value = null;
    formulaInputRef.value?.blur();
  });
}
</script>

<style scoped>
.xlsx-formula-bar__name-box {
  background: transparent;
  border: 1px solid rgba(128, 128, 128, 0.2);
  border-right-color: rgba(128, 128, 128, 0.3);
  border-radius: 4px;
  color: inherit;
  font-family: "SF Mono", "Menlo", "Monaco", monospace;
  font-size: 12px;
  height: 28px;
  outline: none;
  padding: 0 8px;
  width: 90px;
  flex-shrink: 0;
}

.xlsx-formula-bar__fx {
  align-items: center;
  color: #71717a;
  display: flex;
  flex-shrink: 0;
  font-family: "SF Mono", "Menlo", "Monaco", monospace;
  font-size: 11px;
  font-style: italic;
  font-weight: 600;
  height: 28px;
  justify-content: center;
  width: 28px;
}

.xlsx-formula-bar__formula-input {
  background: transparent;
  border: 1px solid rgba(128, 128, 128, 0.15);
  border-radius: 4px;
  color: inherit;
  flex: 1;
  font-family: "SF Mono", "Menlo", "Monaco", monospace;
  font-size: 12px;
  height: 28px;
  outline: none;
  padding: 0 8px;
  min-width: 0;
}

.xlsx-formula-bar__formula-input:focus {
  border-color: rgba(59, 130, 246, 0.5);
}

.xlsx-formula-bar__formula-input:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

@media (pointer: coarse) {
  .xlsx-formula-bar button,
  .xlsx-formula-bar input,
  .xlsx-formula-bar select {
    min-height: 44px;
  }
}
</style>
