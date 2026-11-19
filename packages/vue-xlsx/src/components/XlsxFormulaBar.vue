<template>
  <div
    class="xlsx-formula-bar"
    :style="barStyle"
    :data-formula-reference-count="formulaReferenceRanges.length || undefined"
    :data-formula-reference-ranges="formulaReferenceRangeList || undefined"
    :data-formula-reference-active="activeFormulaReferenceAddress || undefined"
  >
    <div class="xlsx-formula-bar__name-box-wrap">
      <input
        class="xlsx-formula-bar__name-box"
        data-testid="xlsx-name-box"
        :value="nameBoxFocused ? nameBoxDraft : nameBoxValue"
        :aria-expanded="showNameBoxList ? 'true' : 'false'"
        aria-autocomplete="list"
        role="combobox"
        :title="m.xlsx.formulaBar.goToName"
        @blur="onNameBoxBlur"
        @focus="onNameBoxFocus"
        @input="onNameBoxInput"
        @keydown="onNameBoxKeydown"
      />
      <div
        v-if="showNameBoxList"
        class="xlsx-formula-bar__name-box-list"
        data-testid="xlsx-name-box-list"
        role="listbox"
        :aria-label="m.xlsx.formulaBar.definedNames"
      >
        <button
          v-for="(entry, index) in nameBoxMatches"
          :key="`${entry.scope}:${entry.sheetName ?? ''}:${entry.name}`"
          type="button"
          class="xlsx-formula-bar__name-box-option"
          :class="{ 'xlsx-formula-bar__name-box-option--active': index === nameBoxActiveIndex }"
          data-name-box-option
          :data-name="entry.name"
          :data-scope="entry.scope"
          :data-sheet-name="entry.sheetName || undefined"
          role="option"
          :aria-selected="index === nameBoxActiveIndex ? 'true' : 'false'"
          @mousedown.prevent
          @click="onNameBoxOption(entry)"
        >
          <span class="xlsx-formula-bar__name-box-option-name">{{ entry.name }}</span>
          <span v-if="entry.scope === 'sheet' && entry.sheetName" class="xlsx-formula-bar__name-box-option-scope">{{ entry.sheetName }}</span>
        </button>
      </div>
    </div>
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
import type { XlsxDefinedName, XlsxViewerController } from "@arcships/xlsx-core";
import { useXlsxMessages } from "../i18n";
import {
  collectFormulaReferences,
  formulaReferenceSegments,
  resolveNamedRangeTarget,
  type FormulaReferenceCollection,
} from "../composables/formula-reference-highlight";
import { parseA1RangeReference } from "../composables/selection";
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

// ── Name box ────────────────────────────────────────────────────────────
// The name box lists the workbook's defined names and navigates to the range
// they resolve to; typing an A1 address or range also jumps there, matching
// Excel. The list is a plain combobox so hosts stay free to restyle it.
const nameBoxFocused = ref(false);
const nameBoxDraft = ref("");
const nameBoxActiveIndex = ref(-1);
const definedNames = computed<XlsxDefinedName[]>(() => props.controller.definedNames ?? []);
const nameBoxMatches = computed(() => {
  const query = nameBoxDraft.value.trim().toLowerCase();
  return definedNames.value.filter((entry) =>
    !query || entry.name.toLowerCase().includes(query)
  );
});
const showNameBoxList = computed(() => nameBoxFocused.value && nameBoxMatches.value.length > 0);

watch(nameBoxValue, (value) => {
  if (!nameBoxFocused.value) {
    nameBoxDraft.value = value;
  }
}, { immediate: true });

function onNameBoxFocus() {
  nameBoxFocused.value = true;
  nameBoxDraft.value = nameBoxValue.value;
  nameBoxActiveIndex.value = -1;
}

function onNameBoxBlur() {
  nameBoxFocused.value = false;
  nameBoxDraft.value = nameBoxValue.value;
  nameBoxActiveIndex.value = -1;
}

function onNameBoxInput(event: Event) {
  nameBoxDraft.value = (event.target as HTMLInputElement).value;
  nameBoxActiveIndex.value = -1;
}

function onNameBoxKeydown(event: KeyboardEvent) {
  if (event.key === "ArrowDown" || event.key === "ArrowUp") {
    if (nameBoxMatches.value.length === 0) return;
    event.preventDefault();
    const delta = event.key === "ArrowDown" ? 1 : -1;
    const next = nameBoxActiveIndex.value + delta;
    nameBoxActiveIndex.value = next < 0
      ? nameBoxMatches.value.length - 1
      : next % nameBoxMatches.value.length;
    return;
  }
  if (event.key === "Escape") {
    event.preventDefault();
    nameBoxDraft.value = nameBoxValue.value;
    nameBoxActiveIndex.value = -1;
    (event.target as HTMLInputElement).blur();
    return;
  }
  if (event.key === "Enter") {
    event.preventDefault();
    const active = nameBoxMatches.value[nameBoxActiveIndex.value];
    if (active) {
      goToDefinedName(active);
    } else {
      commitNameBoxText(nameBoxDraft.value);
    }
    nameBoxDraft.value = nameBoxValue.value;
    nameBoxActiveIndex.value = -1;
  }
}

function onNameBoxOption(entry: XlsxDefinedName) {
  if (goToDefinedName(entry)) {
    nameBoxDraft.value = nameBoxValue.value;
    nameBoxActiveIndex.value = -1;
  }
}

function selectRangeOnSheet(
  sheetName: string | null | undefined,
  range: { start: { col: number; row: number }; end: { col: number; row: number } },
) {
  const targetSheet = sheetName?.toLowerCase();
  if (targetSheet && targetSheet !== props.controller.activeSheet?.name.toLowerCase()) {
    const index = props.controller.sheets.findIndex(
      (sheet) => sheet.name.toLowerCase() === targetSheet,
    );
    if (index >= 0) {
      props.controller.setActiveSheetIndex(index);
    }
  }
  props.controller.selectRange(range);
}

function goToDefinedName(entry: XlsxDefinedName) {
  const target = resolveNamedRangeTarget(
    entry.range,
    entry.sheetName ?? props.controller.activeSheet?.name ?? null,
  );
  if (!target) return false;
  selectRangeOnSheet(target.sheetName, target.range);
  return true;
}

function commitNameBoxText(text: string) {
  const trimmed = text.trim();
  if (!trimmed) return false;
  const entry = definedNames.value.find(
    (candidate) => candidate.name.toLowerCase() === trimmed.toLowerCase(),
  );
  if (entry) {
    return goToDefinedName(entry);
  }
  const range = parseA1RangeReference(trimmed.replace(/\$/g, ""));
  if (!range) return false;
  props.controller.selectRange(range);
  return true;
}

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
.xlsx-formula-bar__name-box-wrap { flex-shrink: 0; position: relative; }
.xlsx-formula-bar__name-box-list {
  background: #ffffff;
  border: 1px solid #d4d4d8;
  border-radius: 4px;
  box-shadow: 0 8px 20px rgba(15, 23, 42, 0.18);
  display: flex;
  flex-direction: column;
  left: 0;
  max-height: 240px;
  min-width: 180px;
  overflow-y: auto;
  position: absolute;
  top: calc(100% + 2px);
  z-index: 40;
}
.xlsx-formula-bar__name-box-option {
  align-items: baseline;
  background: transparent;
  border: 0;
  color: #18181b;
  cursor: pointer;
  display: flex;
  font: inherit;
  font-size: 12px;
  gap: 8px;
  justify-content: space-between;
  padding: 4px 8px;
  text-align: left;
  white-space: nowrap;
}
.xlsx-formula-bar__name-box-option:hover,
.xlsx-formula-bar__name-box-option--active { background: #dbeafe; }
.xlsx-formula-bar__name-box-option-scope { color: #71717a; font-size: 11px; }

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
