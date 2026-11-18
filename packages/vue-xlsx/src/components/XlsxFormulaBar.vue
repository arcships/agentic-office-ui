<template>
  <div
    class="xlsx-formula-bar"
    :style="barStyle"
    :data-formula-reference-count="formulaReferenceRanges.length || undefined"
    :data-formula-reference-ranges="formulaReferenceRangeList || undefined"
    :data-formula-reference-active="activeFormulaReferenceAddress || undefined"
  >
    <input
      class="xlsx-formula-bar__name-box"
      data-testid="xlsx-name-box"
      :value="nameBoxValue"
      readonly
      :title="m.xlsx.formulaBar.activeCell"
    />
    <span class="xlsx-formula-bar__fx">fx</span>
    <div class="xlsx-formula-bar__field">
      <!-- The input keeps its native caret, selection and IME behaviour; this
           backdrop only paints the reference text in Excel's colours. -->
      <div
        v-if="showReferenceHighlight"
        class="xlsx-formula-bar__reference-layer"
        data-testid="xlsx-formula-reference-highlight"
        aria-hidden="true"
        :data-formula-reference-count="formulaReferenceRanges.length"
        :data-formula-reference-ranges="formulaReferenceRangeList"
        :data-formula-reference-active="activeFormulaReferenceAddress || undefined"
        :style="referenceLayerStyle"
      >
        <span class="xlsx-formula-bar__reference-text" :style="referenceTextStyle"><span
          v-for="(segment, index) in referenceSegments"
          :key="index"
          :data-formula-reference="segment.reference?.rangeAddress"
          :data-formula-reference-active="segment.active ? 'true' : undefined"
          :style="segment.color ? { color: segment.color, textDecoration: segment.active ? 'underline' : undefined } : undefined"
        >{{ segment.text }}</span></span>
      </div>
      <input
        ref="formulaInputRef"
        class="xlsx-formula-bar__formula-input"
        data-testid="xlsx-formula-input"
        :disabled="!hasFormulaTarget || isReadOnly"
        :value="formulaDraft"
        :placeholder="m.xlsx.formulaBar.placeholder"
        :style="formulaInputStyle"
        @blur="onFormulaBlur"
        @focus="onFormulaFocus"
        @input="onFormulaInput"
        @keydown="onFormulaKeydown"
        @keyup="syncFormulaCaret"
        @click="syncFormulaCaret"
        @select="syncFormulaCaret"
        @scroll="syncFormulaScroll"
        @compositionstart="onFormulaCompositionStart"
        @compositionend="onFormulaCompositionEnd"
      />
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, watch, type CSSProperties } from "vue";
import type { XlsxViewerController } from "@arcships/xlsx-core";
import { useXlsxMessages } from "../i18n";
import {
  collectFormulaReferences,
  formulaReferenceSegments,
  type FormulaReferenceCollection,
} from "../composables/formula-reference-highlight";
import {
  createNamedRangeResolver,
  FORMULA_REFERENCE_OWNER_FORMULA_BAR,
  useFormulaReferenceEditing,
} from "../composables/useFormulaReferenceEditing";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
  readOnly?: boolean;
}>();

const formulaInputRef = ref<HTMLInputElement | null>(null);
const m = useXlsxMessages();
const formulaDraft = ref("");
const formulaInitialValue = ref("");
const formulaCaret = ref(0);
const formulaScrollLeft = ref(0);
const focusedField = ref<"formula" | null>(null);
const isFormulaComposing = ref(false);
let commitFormulaOnCompositionEnd = false;

const formulaReferenceEditing = useFormulaReferenceEditing(props.controller);
const resolveFormulaReferenceName = createNamedRangeResolver(props.controller);
const EMPTY_FORMULA_REFERENCES: FormulaReferenceCollection = {
  activeIndex: -1,
  activeReference: null,
  references: [],
};

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

// ── Formula reference highlighting ────────────────────────────────────
// The references are parsed from the draft at the caret, never from the
// workbook, so typing a half-finished `=SUM(A1:A` stays harmless until commit.
const formulaReferences = computed<FormulaReferenceCollection>(() => {
  if (
    focusedField.value !== "formula"
    || !formulaReferenceEditing.session.active.value
    || formulaReferenceEditing.session.owner.value !== FORMULA_REFERENCE_OWNER_FORMULA_BAR
  ) return EMPTY_FORMULA_REFERENCES;
  return collectFormulaReferences(formulaDraft.value, formulaCaret.value, {
    activeSheetName: props.controller.activeSheet?.name ?? null,
    resolveName: resolveFormulaReferenceName,
  });
});
const showReferenceHighlight = computed(() =>
  formulaReferences.value.references.length > 0 && !isFormulaComposing.value,
);
const formulaReferenceRanges = computed(() => formulaReferences.value.references);
const formulaReferenceRangeList = computed(() =>
  formulaReferenceRanges.value.map((reference) => reference.rangeAddress).join(" "),
);
const activeFormulaReferenceAddress = computed(
  () => formulaReferences.value.activeReference?.rangeAddress ?? "",
);
const referenceSegments = computed(() =>
  formulaReferenceSegments(
    formulaDraft.value,
    formulaReferences.value.references,
    formulaReferences.value.activeIndex,
  ),
);
const referenceLayerStyle = computed<CSSProperties>(() => ({
  alignItems: "center",
  border: "1px solid transparent",
  boxSizing: "border-box",
  color: props.isDark ? "#e4e4e7" : "#18181b",
  display: "flex",
  fontFamily: '"SF Mono", "Menlo", "Monaco", monospace',
  fontSize: "12px",
  height: "28px",
  inset: "0",
  overflow: "hidden",
  padding: "0 8px",
  pointerEvents: "none",
  position: "absolute",
  whiteSpace: "pre",
}));
const referenceTextStyle = computed<CSSProperties>(() => ({
  display: "inline-block",
  transform: `translateX(${-formulaScrollLeft.value}px)`,
  whiteSpace: "pre",
}));
const formulaInputStyle = computed<CSSProperties>(() => ({
  position: "relative",
  zIndex: "1",
  ...(showReferenceHighlight.value
    ? { caretColor: props.isDark ? "#e4e4e7" : "#18181b", color: "transparent" }
    : {}),
}));
function readFormulaCaret(input: HTMLInputElement | null = formulaInputRef.value): number {
  const value = input?.selectionStart;
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.max(0, Math.min(value, formulaDraft.value.length));
  }
  return formulaDraft.value.length;
}

function syncFormulaScroll(event?: Event) {
  const input = (event?.target as HTMLInputElement | null) ?? formulaInputRef.value;
  formulaScrollLeft.value = input?.scrollLeft ?? 0;
}

function syncFormulaCaret(event?: Event) {
  const input = (event?.target as HTMLInputElement | null) ?? formulaInputRef.value;
  formulaCaret.value = readFormulaCaret(input);
  syncFormulaScroll();
  if (focusedField.value === "formula") {
    formulaReferenceEditing.update(
      FORMULA_REFERENCE_OWNER_FORMULA_BAR,
      formulaDraft.value,
      formulaCaret.value,
    );
  }
}

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
  formulaCaret.value = readFormulaCaret();
  syncFormulaScroll();
  formulaReferenceEditing.begin(
    FORMULA_REFERENCE_OWNER_FORMULA_BAR,
    formulaDraft.value,
    formulaCaret.value,
  );
}

function onFormulaBlur() {
  formulaReferenceEditing.end(FORMULA_REFERENCE_OWNER_FORMULA_BAR);
  commitFormula();
  focusedField.value = null;
}

function onFormulaInput(event: Event) {
  const target = event.target as HTMLInputElement;
  formulaDraft.value = target.value;
  formulaCaret.value = readFormulaCaret(target);
  formulaScrollLeft.value = target.scrollLeft ?? 0;
  formulaReferenceEditing.update(
    FORMULA_REFERENCE_OWNER_FORMULA_BAR,
    formulaDraft.value,
    formulaCaret.value,
  );
}

/** Drops the draft back to the cell's value, like Escape in Excel. */
function cancelFormulaEdit() {
  formulaDraft.value = formulaInitialValue.value;
  formulaCaret.value = formulaDraft.value.length;
  formulaReferenceEditing.end(FORMULA_REFERENCE_OWNER_FORMULA_BAR);
  focusedField.value = null;
  formulaInputRef.value?.blur();
}

function commitFormula(nextFormula?: string) {
  const resolvedFormula = nextFormula ?? formulaDraft.value;
  if (!hasFormulaTarget.value) return;
  if (resolvedFormula === formulaInitialValue.value) return;
  props.controller.setSelectedFormula(resolvedFormula);
  // The blur that follows Enter must not write the same formula a second time.
  formulaInitialValue.value = resolvedFormula;
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
    formulaReferenceEditing.end(FORMULA_REFERENCE_OWNER_FORMULA_BAR);
    commitFormula();
    focusedField.value = null;
    formulaInputRef.value?.blur();
  } else if (event.key === "Escape") {
    event.preventDefault();
    cancelFormulaEdit();
  }
}

function onFormulaCompositionStart() {
  isFormulaComposing.value = true;
  commitFormulaOnCompositionEnd = false;
}

function onFormulaCompositionEnd() {
  isFormulaComposing.value = false;
  if (!commitFormulaOnCompositionEnd) {
    void nextTick(syncFormulaCaret);
    return;
  }
  commitFormulaOnCompositionEnd = false;
  // The final composition text lands via @input before this tick completes.
  void nextTick(() => {
    formulaReferenceEditing.end(FORMULA_REFERENCE_OWNER_FORMULA_BAR);
    commitFormula();
    focusedField.value = null;
    formulaInputRef.value?.blur();
  });
}

watch(isReadOnly, (readOnly) => {
  if (!readOnly) return;
  formulaReferenceEditing.end(FORMULA_REFERENCE_OWNER_FORMULA_BAR);
  if (focusedField.value === "formula") focusedField.value = null;
});

onBeforeUnmount(() => {
  formulaReferenceEditing.end(FORMULA_REFERENCE_OWNER_FORMULA_BAR);
});
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
  width: 100%;
}

.xlsx-formula-bar__field {
  display: flex;
  flex: 1;
  min-width: 0;
  position: relative;
}

.xlsx-formula-bar__reference-layer {
  left: 0;
  right: 0;
  top: 0;
  bottom: 0;
}

.xlsx-formula-bar__reference-text {
  display: inline-block;
  white-space: pre;
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
