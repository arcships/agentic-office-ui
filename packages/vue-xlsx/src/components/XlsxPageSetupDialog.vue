<template>
  <OfficeThemeProvider :theme="isDark ? 'dark' : 'light'">
    <OfficeDialog
      :open="true"
      :title="m.xlsx.authoring.pageSetup.title"
      :close-label="m.xlsx.authoring.close"
      close-test-id="xlsx-page-setup-close"
      content-test-id="xlsx-page-setup-dialog"
      :content-data="{ 'data-kind': 'pageSetup' }"
      class="xlsx-authoring xlsx-page-setup"
      @update:open="onDialogOpenChange"
    >
      <p class="xlsx-authoring__hint" data-testid="xlsx-page-setup-sheet">{{ sheetLabel }}</p>
      <p
        v-if="locked"
        class="xlsx-page-setup__locked"
        data-testid="xlsx-page-setup-locked"
        role="status"
      >
        {{ m.xlsx.authoring.pageSetup.locked }}
      </p>

      <!-- Page tab -->
      <section class="xlsx-page-setup__section">
        <div class="xlsx-page-setup__grid">
          <label class="xlsx-authoring__field">
            <span>{{ m.xlsx.authoring.pageSetup.orientation }}</span>
            <select
              v-model="orientation"
              data-testid="xlsx-page-setup-orientation"
              :disabled="locked"
            >
              <option value="portrait">{{ m.xlsx.authoring.pageSetup.orientationPortrait }}</option>
              <option value="landscape">{{ m.xlsx.authoring.pageSetup.orientationLandscape }}</option>
            </select>
          </label>
          <label class="xlsx-authoring__field">
            <span>{{ m.xlsx.authoring.pageSetup.paperSize }}</span>
            <select
              v-model.number="paperSize"
              data-testid="xlsx-page-setup-paper"
              :disabled="locked"
            >
              <option v-for="paper in paperOptions" :key="paper.id" :value="paper.id">
                {{ paper.name }}
              </option>
            </select>
          </label>
        </div>

        <label class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.pageSetup.scaling }}</span>
          <select
            v-model="scaleMode"
            data-testid="xlsx-page-setup-scale-mode"
            :disabled="locked"
          >
            <option value="adjust">{{ m.xlsx.authoring.pageSetup.scaleAdjust }}</option>
            <option value="fit">{{ m.xlsx.authoring.pageSetup.scaleFit }}</option>
          </select>
        </label>
        <div v-if="scaleMode === 'adjust'" class="xlsx-page-setup__grid">
          <label class="xlsx-authoring__field">
            <span>{{ m.xlsx.authoring.pageSetup.scalePercent }}</span>
            <input
              v-model.number="scale"
              type="number"
              min="10"
              max="400"
              step="1"
              data-testid="xlsx-page-setup-scale"
              :disabled="locked"
            />
          </label>
        </div>
        <div v-else class="xlsx-page-setup__grid">
          <label class="xlsx-authoring__field">
            <span>{{ m.xlsx.authoring.pageSetup.fitToWidth }}</span>
            <input
              v-model.number="fitToWidth"
              type="number"
              min="1"
              max="32767"
              step="1"
              data-testid="xlsx-page-setup-fit-width"
              :disabled="locked"
            />
          </label>
          <label class="xlsx-authoring__field">
            <span>{{ m.xlsx.authoring.pageSetup.fitToHeight }}</span>
            <input
              v-model.number="fitToHeight"
              type="number"
              min="0"
              max="32767"
              step="1"
              data-testid="xlsx-page-setup-fit-height"
              :disabled="locked"
            />
          </label>
        </div>
        <p v-if="scaleMode === 'fit'" class="xlsx-authoring__hint">
          {{ m.xlsx.authoring.pageSetup.fitToHeightHint }}
        </p>
      </section>

      <!-- Margins tab -->
      <section class="xlsx-page-setup__section">
        <p class="xlsx-page-setup__legend">{{ m.xlsx.authoring.pageSetup.margins }}</p>
        <div class="xlsx-page-setup__grid">
          <label
            v-for="field in marginFields"
            :key="field.key"
            class="xlsx-authoring__field"
          >
            <span>{{ field.label }}</span>
            <input
              v-model.number="margins[field.key]"
              type="number"
              min="0"
              max="10"
              step="0.1"
              :data-testid="`xlsx-page-setup-margin-${field.key}`"
              :disabled="locked"
            />
          </label>
        </div>
        <div class="xlsx-page-setup__checks">
          <label class="xlsx-authoring__check">
            <input
              v-model="centerHorizontally"
              type="checkbox"
              data-testid="xlsx-page-setup-center-horizontal"
              :disabled="locked"
            />
            <span>{{ m.xlsx.authoring.pageSetup.centerHorizontally }}</span>
          </label>
          <label class="xlsx-authoring__check">
            <input
              v-model="centerVertically"
              type="checkbox"
              data-testid="xlsx-page-setup-center-vertical"
              :disabled="locked"
            />
            <span>{{ m.xlsx.authoring.pageSetup.centerVertically }}</span>
          </label>
        </div>
      </section>

      <!-- Sheet tab -->
      <section class="xlsx-page-setup__section">
        <p class="xlsx-page-setup__legend">{{ m.xlsx.authoring.pageSetup.printOptions }}</p>
        <div class="xlsx-page-setup__checks">
          <label class="xlsx-authoring__check">
            <input
              v-model="printGridLines"
              type="checkbox"
              data-testid="xlsx-page-setup-grid-lines"
              :disabled="locked"
            />
            <span>{{ m.xlsx.authoring.pageSetup.printGridLines }}</span>
          </label>
          <label class="xlsx-authoring__check">
            <input
              v-model="printHeadings"
              type="checkbox"
              data-testid="xlsx-page-setup-headings"
              :disabled="locked"
            />
            <span>{{ m.xlsx.authoring.pageSetup.printHeadings }}</span>
          </label>
        </div>
        <label class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.pageSetup.pageOrder }}</span>
          <select
            v-model="pageOrder"
            data-testid="xlsx-page-setup-page-order"
            :disabled="locked"
          >
            <option value="downThenOver">{{ m.xlsx.authoring.pageSetup.pageOrderDownThenOver }}</option>
            <option value="overThenDown">{{ m.xlsx.authoring.pageSetup.pageOrderOverThenDown }}</option>
          </select>
        </label>
        <label class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.pageSetup.printArea }}</span>
          <input
            v-model="printAreaText"
            type="text"
            :placeholder="m.xlsx.authoring.pageSetup.printAreaPlaceholder"
            data-testid="xlsx-page-setup-print-area"
            :aria-invalid="printAreaInvalid ? 'true' : undefined"
            :disabled="locked"
          />
        </label>
        <div class="xlsx-authoring__actions">
          <button
            type="button"
            class="xlsx-authoring__button"
            data-testid="xlsx-page-setup-print-area-selection"
            :disabled="locked || !selectionRange"
            @click="applySelectionAsPrintArea"
          >{{ m.xlsx.authoring.pageSetup.printAreaUseSelection }}</button>
          <button
            type="button"
            class="xlsx-authoring__button"
            data-testid="xlsx-page-setup-print-area-clear"
            :disabled="locked || printAreaText.trim() === ''"
            @click="printAreaText = ''"
          >{{ m.xlsx.authoring.pageSetup.printAreaClear }}</button>
        </div>
        <p
          v-if="printAreaInvalid"
          class="xlsx-page-setup__error"
          data-testid="xlsx-page-setup-print-area-error"
          role="alert"
        >
          {{ m.xlsx.authoring.pageSetup.printAreaInvalid }}
        </p>
      </section>

      <p class="xlsx-authoring__hint" data-testid="xlsx-page-setup-summary">{{ summary }}</p>

      <div class="xlsx-authoring__actions xlsx-page-setup__actions">
        <button
          type="button"
          class="xlsx-authoring__button xlsx-authoring__button--primary"
          data-testid="xlsx-page-setup-apply"
          :disabled="locked || printAreaInvalid"
          @click="onApply"
        >{{ m.xlsx.authoring.pageSetup.apply }}</button>
        <button
          type="button"
          class="xlsx-authoring__button"
          data-testid="xlsx-page-setup-cancel"
          @click="emit('close')"
        >{{ m.xlsx.authoring.pageSetup.cancel }}</button>
      </div>
    </OfficeDialog>
  </OfficeThemeProvider>
</template>

<script setup lang="ts">
import { computed, reactive, ref, watch } from "vue";
import { formatOfficeMessage } from "@arcships/office-interaction";
import {
  PAPER_SIZES,
  buildWorksheetPrintLayout,
  parseA1RangeReference,
  resolvePageSetupInput,
  type XlsxCellRange,
  type XlsxPageMargins,
  type XlsxPageOrder,
  type XlsxPageOrientation,
  type XlsxPageSetupInput,
  type XlsxViewerController,
} from "@arcships/xlsx-core";
import { OfficeDialog, OfficeThemeProvider } from "@arcships/vue-ui";
import { useXlsxMessages } from "../i18n";
import { rangeToA1 } from "../composables/selection";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
}>();

const emit = defineEmits<{ close: [] }>();

const m = useXlsxMessages();

/** Paper sizes Excel lists in `Page Layout > Size`, resolved from OOXML ids. */
const COMMON_PAPER_SIZES = [1, 5, 7, 3, 9, 8, 11, 13, 12];

const margins = reactive<XlsxPageMargins>({
  bottom: 0.75,
  footer: 0.3,
  header: 0.3,
  left: 0.7,
  right: 0.7,
  top: 0.75,
});
const orientation = ref<XlsxPageOrientation>("portrait");
const paperSize = ref(1);
const pageOrder = ref<XlsxPageOrder>("downThenOver");
const scaleMode = ref<"adjust" | "fit">("adjust");
const scale = ref(100);
const fitToWidth = ref(1);
const fitToHeight = ref(1);
const centerHorizontally = ref(false);
const centerVertically = ref(false);
const printGridLines = ref(false);
const printHeadings = ref(false);
const printAreaText = ref("");

const sheet = computed(() => props.controller.activeSheet);
const selectionRange = computed<XlsxCellRange | null>(() => {
  const selection = props.controller.selection;
  if (selection) return selection;
  const cell = props.controller.activeCell;
  return cell ? { end: cell, start: cell } : null;
});
const locked = computed(() =>
  props.controller.readOnly === true
  || props.controller.isEditingLocked === true
  || props.controller.isWorkerBacked === true
);
const sheetLabel = computed(() => sheet.value?.name ?? "");

/**
 * Loads the dialog's draft from the workbook.
 *
 * Excel initializes every field from the worksheet's current values, so a sheet
 * stored as landscape/A3 with 0.5in margins opens exactly like that instead of
 * showing the OOXML defaults.
 */
function resetDraft() {
  const current = sheet.value;
  if (!current) return;
  const pageSetup = current.pageSetup;
  orientation.value = pageSetup.orientation;
  paperSize.value = pageSetup.paperSize;
  pageOrder.value = pageSetup.pageOrder;
  scaleMode.value = pageSetup.fitToPage ? "fit" : "adjust";
  scale.value = pageSetup.scale;
  fitToWidth.value = pageSetup.fitToWidth;
  fitToHeight.value = pageSetup.fitToHeight;
  Object.assign(margins, current.pageMargins);
  centerHorizontally.value = current.printOptions.horizontalCentered;
  centerVertically.value = current.printOptions.verticalCentered;
  printGridLines.value = current.printOptions.gridLines;
  printHeadings.value = current.printOptions.headings;
  printAreaText.value = current.printAreas.map((range) => rangeToA1(range)).join(",");
}

watch(
  () => sheet.value?.workbookSheetIndex,
  () => resetDraft(),
  { immediate: true },
);

const paperOptions = computed(() => {
  const ids = COMMON_PAPER_SIZES.includes(paperSize.value)
    ? COMMON_PAPER_SIZES
    : [paperSize.value, ...COMMON_PAPER_SIZES];
  return ids.map((id) => ({
    id,
    name: PAPER_SIZES[id]?.name ?? String(id),
  }));
});

const marginFields = computed<Array<{ key: keyof XlsxPageMargins; label: string }>>(() => [
  { key: "top", label: m.value.xlsx.authoring.pageSetup.marginTop },
  { key: "bottom", label: m.value.xlsx.authoring.pageSetup.marginBottom },
  { key: "left", label: m.value.xlsx.authoring.pageSetup.marginLeft },
  { key: "right", label: m.value.xlsx.authoring.pageSetup.marginRight },
  { key: "header", label: m.value.xlsx.authoring.pageSetup.marginHeader },
  { key: "footer", label: m.value.xlsx.authoring.pageSetup.marginFooter },
]);

/** Parsed print area; `invalid` blocks Apply instead of silently dropping it. */
const printArea = computed<{ invalid: boolean; ranges: XlsxCellRange[] }>(() => {
  const text = printAreaText.value.trim();
  if (!text) return { invalid: false, ranges: [] };
  const ranges: XlsxCellRange[] = [];
  for (const part of text.split(",")) {
    const trimmed = part.trim();
    if (!trimmed) continue;
    const parsed = parseA1RangeReference(trimmed.replace(/\$/g, ""));
    if (!parsed) return { invalid: true, ranges: [] };
    ranges.push(parsed);
  }
  return { invalid: ranges.length === 0, ranges };
});
const printAreaInvalid = computed(() => printArea.value.invalid);

function coerceNumber(value: unknown, fallback: number) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

/** The patch the dialog would write; also drives the live page summary. */
function buildInput(): XlsxPageSetupInput {
  const pageSetup = scaleMode.value === "fit"
    ? {
        fitToHeight: coerceNumber(fitToHeight.value, 1),
        fitToPage: true,
        fitToWidth: coerceNumber(fitToWidth.value, 1),
        orientation: orientation.value,
        pageOrder: pageOrder.value,
        paperSize: coerceNumber(paperSize.value, 1),
      }
    : {
        fitToPage: false,
        orientation: orientation.value,
        pageOrder: pageOrder.value,
        paperSize: coerceNumber(paperSize.value, 1),
        scale: coerceNumber(scale.value, 100),
      };
  return {
    pageMargins: {
      bottom: coerceNumber(margins.bottom, 0.75),
      footer: coerceNumber(margins.footer, 0.3),
      header: coerceNumber(margins.header, 0.3),
      left: coerceNumber(margins.left, 0.7),
      right: coerceNumber(margins.right, 0.7),
      top: coerceNumber(margins.top, 0.75),
    },
    pageSetup,
    printAreas: printArea.value.ranges,
    printOptions: {
      gridLines: printGridLines.value,
      headings: printHeadings.value,
      horizontalCentered: centerHorizontally.value,
      verticalCentered: centerVertically.value,
    },
  };
}

/**
 * Page count and effective scale the draft would print.
 *
 * It runs the real pagination over a draft sheet, so the dialog cannot promise
 * a layout the print preview then contradicts.
 */
const summary = computed(() => {
  const current = sheet.value;
  if (!current) return "";
  const resolved = resolvePageSetupInput(current, buildInput());
  const layout = buildWorksheetPrintLayout({
    ...current,
    pageMargins: resolved.pageMargins,
    pageSetup: resolved.pageSetup,
    printAreas: resolved.printAreas ?? current.printAreas,
    printOptions: resolved.printOptions,
  });
  return formatOfficeMessage(m.value.xlsx.authoring.pageSetup.summary, {
    count: String(layout.pages.length),
    orientation: layout.pageSetup.orientation === "landscape"
      ? m.value.xlsx.print.orientationLandscape
      : m.value.xlsx.print.orientationPortrait,
    paper: layout.paperSize.name,
    scale: String(layout.pages[0]?.scale ?? layout.scale),
  });
});

function applySelectionAsPrintArea() {
  const range = selectionRange.value;
  if (!range) return;
  printAreaText.value = rangeToA1(range);
}

function onApply() {
  if (locked.value || printAreaInvalid.value) return;
  const setPageSetup = props.controller.setPageSetup;
  if (typeof setPageSetup === "function") {
    setPageSetup.call(props.controller, buildInput());
  }
  emit("close");
}

/** Reka owns Esc/backdrop dismissal; both funnel into the public close event. */
function onDialogOpenChange(value: boolean) {
  if (!value) emit("close");
}
</script>

<style scoped>
/*
 * `OfficeDialog` owns the shell and `XlsxAuthoringDialog`'s unscoped block owns
 * the field chrome (`.xlsx-authoring__*`), so the dialog only has to lay its own
 * sections out. Every grid is `minmax(0, 1fr)` based so a 390px viewport wraps
 * instead of overflowing; colours and font sizes come from tokens without
 * literal fallbacks (style guard).
 */
.xlsx-page-setup__section {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-2, 0.5rem);
}
.xlsx-page-setup__grid {
  display: grid;
  gap: var(--office-ui-space-2, 0.5rem);
  grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
}
.xlsx-page-setup__checks {
  display: flex;
  flex-wrap: wrap;
  gap: var(--office-ui-space-3, 0.75rem);
}
.xlsx-page-setup__legend {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  font-weight: 600;
  margin: 0;
}
.xlsx-page-setup__actions {
  justify-content: flex-end;
}
.xlsx-page-setup__error {
  color: var(--office-destructive-foreground);
  font-size: var(--office-text-xs);
  margin: 0;
}
.xlsx-page-setup__locked {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  margin: 0;
}
</style>
