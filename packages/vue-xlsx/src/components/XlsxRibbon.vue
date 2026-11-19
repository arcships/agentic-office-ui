<template>
  <div class="xlsx-ribbon" :class="{ 'xlsx-ribbon--dark': isDark }" data-testid="xlsx-format-toolbar">
    <fieldset class="xlsx-ribbon__controls" :disabled="isReadOnly || !hasWorkbook">
      <div class="xlsx-ribbon__group">
        <button data-testid="xlsx-undo" class="xlsx-ribbon__icon-button" :disabled="!controller.canUndo" :title="m.xlsx.ribbon.undo" :aria-label="m.xlsx.ribbon.undo" @click="controller.undo()">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 7 4 12l5 5M5 12h8a6 6 0 0 1 6 6" /></svg>
        </button>
        <button data-testid="xlsx-redo" class="xlsx-ribbon__icon-button" :disabled="!controller.canRedo" :title="m.xlsx.ribbon.redo" :aria-label="m.xlsx.ribbon.redo" @click="controller.redo()">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m15 7 5 5-5 5M19 12h-8a6 6 0 0 0-6 6" /></svg>
        </button>
      </div>
      <span class="xlsx-ribbon__separator" />
      <div class="xlsx-ribbon__group">
        <button data-testid="xlsx-ribbon-cut" class="xlsx-ribbon__action-button" :disabled="!hasSelection" :title="m.xlsx.ribbon.cut" @click="cutSelection">{{ m.xlsx.ribbon.cut }}</button>
        <button data-testid="xlsx-ribbon-copy" class="xlsx-ribbon__action-button" :disabled="!hasSelection" :title="m.xlsx.ribbon.copy" @click="copySelection">{{ m.xlsx.ribbon.copy }}</button>
        <button data-testid="xlsx-ribbon-paste" class="xlsx-ribbon__action-button" :disabled="!hasActiveCell" :title="m.xlsx.ribbon.paste" @click="controller.pasteFromClipboard()">{{ m.xlsx.ribbon.paste }}</button>
        <button data-testid="xlsx-ribbon-paste-special" class="xlsx-ribbon__action-button" :disabled="!hasSelection" :title="m.xlsx.ribbon.pasteSpecial" @click="emit('authoring', 'pasteSpecial')">{{ m.xlsx.ribbon.pasteSpecial }}</button>
      </div>
      <span class="xlsx-ribbon__separator" />
      <div class="xlsx-ribbon__group">
        <button data-testid="xlsx-ribbon-insert-row-above" class="xlsx-ribbon__icon-button" :disabled="!hasActiveCell" :title="m.xlsx.ribbon.insertRowAbove" :aria-label="m.xlsx.ribbon.insertRowAbove" @click="insertRow('above')"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 4h16M4 20h16M12 8v8M8 12h8" /></svg></button>
        <button data-testid="xlsx-ribbon-insert-row-below" class="xlsx-ribbon__icon-button" :disabled="!hasActiveCell" :title="m.xlsx.ribbon.insertRowBelow" :aria-label="m.xlsx.ribbon.insertRowBelow" @click="insertRow('below')"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 4h16M4 20h16M12 6v8M8 10h8" /></svg></button>
        <button data-testid="xlsx-ribbon-delete-row" class="xlsx-ribbon__icon-button" :disabled="!hasActiveCell" :title="m.xlsx.ribbon.deleteRow" :aria-label="m.xlsx.ribbon.deleteRow" @click="deleteRow"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 4h16M4 20h16M8 12h8" /></svg></button>
        <button data-testid="xlsx-ribbon-insert-column-left" class="xlsx-ribbon__icon-button" :disabled="!hasActiveCell" :title="m.xlsx.ribbon.insertColumnLeft" :aria-label="m.xlsx.ribbon.insertColumnLeft" @click="insertColumn('left')"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 4v16M20 4v16M8 12h8M12 8v8" /></svg></button>
        <button data-testid="xlsx-ribbon-insert-column-right" class="xlsx-ribbon__icon-button" :disabled="!hasActiveCell" :title="m.xlsx.ribbon.insertColumnRight" :aria-label="m.xlsx.ribbon.insertColumnRight" @click="insertColumn('right')"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 4v16M20 4v16M8 12h8M12 8v8" /></svg></button>
        <button data-testid="xlsx-ribbon-delete-column" class="xlsx-ribbon__icon-button" :disabled="!hasActiveCell" :title="m.xlsx.ribbon.deleteColumn" :aria-label="m.xlsx.ribbon.deleteColumn" @click="deleteColumn"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 4v16M20 4v16M8 12h8" /></svg></button>
      </div>
      <span class="xlsx-ribbon__separator" />
      <div class="xlsx-ribbon__group">
        <select :value="fontFamilyValue" :aria-label="m.xlsx.ribbon.fontFamily" :title="m.xlsx.ribbon.fontFamily" @change="setFontFamily">
          <option value="Calibri">Calibri</option>
          <option value="Aptos">Aptos</option>
          <option value="Arial">Arial</option>
          <option value="Georgia">Georgia</option>
          <option value="Times New Roman">Times New Roman</option>
          <option value="Courier New">Courier New</option>
        </select>
        <select :value="fontSizeValue" class="xlsx-ribbon__size" :aria-label="m.xlsx.ribbon.fontSize" :title="m.xlsx.ribbon.fontSize" @change="setFontSize">
          <option v-if="!fontSizes.includes(fontSizeValue)" :value="fontSizeValue">{{ fontSizeValue }}</option>
          <option v-for="size in fontSizes" :key="size" :value="size">{{ size }}</option>
        </select>
        <button class="xlsx-ribbon__text-button" :class="{ active: fontState.bold }" :aria-pressed="fontState.bold" :title="m.xlsx.ribbon.bold" :aria-label="m.xlsx.ribbon.bold" @click="applyFont({ bold: !fontState.bold })"><strong>B</strong></button>
        <button class="xlsx-ribbon__text-button" :class="{ active: fontState.italic }" :aria-pressed="fontState.italic" :title="m.xlsx.ribbon.italic" :aria-label="m.xlsx.ribbon.italic" @click="applyFont({ italic: !fontState.italic })"><em>I</em></button>
        <button class="xlsx-ribbon__text-button" :class="{ active: fontState.underline }" :aria-pressed="fontState.underline" :title="m.xlsx.ribbon.underline" :aria-label="m.xlsx.ribbon.underline" @click="applyFont({ underline: fontState.underline ? 'none' : 'single' })"><u>U</u></button>
        <button class="xlsx-ribbon__text-button" :class="{ active: fontState.strikethrough }" :aria-pressed="fontState.strikethrough" :title="m.xlsx.ribbon.strikethrough" :aria-label="m.xlsx.ribbon.strikethrough" @click="applyFont({ strikethrough: !fontState.strikethrough })"><s>S</s></button>
        <label class="xlsx-ribbon__color" :title="m.xlsx.ribbon.textColor">
          <span>A</span><i :style="{ backgroundColor: textColor }" />
          <input v-model="textColor" type="color" :aria-label="m.xlsx.ribbon.textColor" @change="applyTextColor" />
        </label>
        <label class="xlsx-ribbon__color" :title="m.xlsx.ribbon.fillColor">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m7 4 10 10-5 5-7-7zM14 6l2-2 4 4-2 2" /></svg><i :style="{ backgroundColor: fillColor }" />
          <input v-model="fillColor" type="color" :aria-label="m.xlsx.ribbon.fillColor" @change="applyFillColor" />
        </label>
      </div>
      <span class="xlsx-ribbon__separator" />
      <div class="xlsx-ribbon__group">
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.horizontal === 'left' }" :aria-pressed="alignmentState.horizontal === 'left'" :title="m.xlsx.ribbon.alignLeft" :aria-label="m.xlsx.ribbon.alignLeft" @click="applyAlignment({ horizontal: 'left' })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 10h11M4 14h16M4 18h9" /></svg></button>
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.horizontal === 'center' }" :aria-pressed="alignmentState.horizontal === 'center'" :title="m.xlsx.ribbon.alignCenter" :aria-label="m.xlsx.ribbon.alignCenter" @click="applyAlignment({ horizontal: 'center' })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M7 10h10M4 14h16M8 18h8" /></svg></button>
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.horizontal === 'right' }" :aria-pressed="alignmentState.horizontal === 'right'" :title="m.xlsx.ribbon.alignRight" :aria-label="m.xlsx.ribbon.alignRight" @click="applyAlignment({ horizontal: 'right' })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M9 10h11M4 14h16M11 18h9" /></svg></button>
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.vertical === 'top' }" :aria-pressed="alignmentState.vertical === 'top'" :title="m.xlsx.ribbon.alignTop" :aria-label="m.xlsx.ribbon.alignTop" @click="applyAlignment({ vertical: 'top' })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 4h14M7 8h10M9 12h6" /></svg></button>
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.vertical === 'center' }" :aria-pressed="alignmentState.vertical === 'center'" :title="m.xlsx.ribbon.alignMiddle" :aria-label="m.xlsx.ribbon.alignMiddle" @click="applyAlignment({ vertical: 'center' })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 5h14M7 9h10M7 15h10M5 19h14" /></svg></button>
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.vertical === 'bottom' }" :aria-pressed="alignmentState.vertical === 'bottom'" :title="m.xlsx.ribbon.alignBottom" :aria-label="m.xlsx.ribbon.alignBottom" @click="applyAlignment({ vertical: 'bottom' })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 12h6M7 16h10M5 20h14" /></svg></button>
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.wrapText }" :aria-pressed="alignmentState.wrapText" :title="m.xlsx.ribbon.wrapText" :aria-label="m.xlsx.ribbon.wrapText" @click="applyAlignment({ wrapText: !alignmentState.wrapText })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 10h12a3 3 0 0 1 0 6h-3M16 13l-3 3 3 3M4 20h7" /></svg></button>
      </div>
      <span class="xlsx-ribbon__separator" />
      <div class="xlsx-ribbon__group">
        <select :value="numberFormatValue" class="xlsx-ribbon__number" :aria-label="m.xlsx.ribbon.numberFormat" :title="m.xlsx.ribbon.numberFormat" @change="setNumberFormat">
          <option value="general">{{ m.xlsx.ribbon.numberFormats.general }}</option>
          <option value="number">{{ m.xlsx.ribbon.numberFormats.number }}</option>
          <option value="currency">{{ m.xlsx.ribbon.numberFormats.currency }}</option>
          <option value="percent">{{ m.xlsx.ribbon.numberFormats.percent }}</option>
          <option value="date">{{ m.xlsx.ribbon.numberFormats.date }}</option>
          <option value="text">{{ m.xlsx.ribbon.numberFormats.text }}</option>
        </select>
        <button class="xlsx-ribbon__text-button" :title="m.xlsx.ribbon.currency" :aria-label="m.xlsx.ribbon.currency" @click="applyStyle(numberFormats.currency)">$</button>
        <button class="xlsx-ribbon__text-button" :title="m.xlsx.ribbon.percent" :aria-label="m.xlsx.ribbon.percent" @click="applyStyle(numberFormats.percent)">%</button>
        <button data-testid="xlsx-ribbon-number-format" class="xlsx-ribbon__action-button" :title="m.xlsx.ribbon.numberFormatDialog" @click="emit('authoring', 'numberFormat')">{{ m.xlsx.ribbon.numberFormatDialog }}</button>
      </div>
      <span class="xlsx-ribbon__separator" />
      <div class="xlsx-ribbon__group">
        <button data-testid="xlsx-ribbon-merge-center" class="xlsx-ribbon__action-button" :disabled="!hasSelection" :title="m.xlsx.ribbon.mergeAndCenter" @click="mergeAndCenter">{{ m.xlsx.ribbon.mergeAndCenter }}</button>
        <button data-testid="xlsx-ribbon-unmerge" class="xlsx-ribbon__action-button" :disabled="!hasSelection" @click="controller.unmergeSelection()">{{ m.xlsx.ribbon.unmerge }}</button>
        <select data-testid="xlsx-ribbon-border-preset" class="xlsx-ribbon__borders" :value="''" :aria-label="m.xlsx.ribbon.borders" :title="m.xlsx.ribbon.borders" @change="applyBorderPreset">
          <option value="">{{ m.xlsx.ribbon.borders }}</option>
          <option v-for="preset in borderPresets" :key="preset" :value="preset">{{ borderLabels[preset] }}</option>
        </select>
        <label class="xlsx-ribbon__color" :title="m.xlsx.ribbon.borderColor">
          <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="4" width="16" height="16" /></svg><i :style="{ backgroundColor: borderColor }" />
          <input v-model="borderColor" data-testid="xlsx-ribbon-border-color" type="color" :aria-label="m.xlsx.ribbon.borderColor" @change="applyBorderColor" />
        </label>
        <button data-testid="xlsx-ribbon-borders-all" class="xlsx-ribbon__icon-button" :title="m.xlsx.ribbon.allBorders" :aria-label="m.xlsx.ribbon.allBorders" @click="applyBorders('all')"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="4" width="16" height="16" /><path d="M12 4v16M4 12h16" /></svg></button>
      </div>
      <span class="xlsx-ribbon__separator" />
      <div class="xlsx-ribbon__group" data-testid="xlsx-ribbon-authoring">
        <button class="xlsx-ribbon__icon-button" data-testid="xlsx-ribbon-chart" :title="m.xlsx.authoring.chart.title" :aria-label="m.xlsx.authoring.chart.title" @click="emit('authoring', 'chart')"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20V4M4 20h16M8 16v-4M12 16V8M16 16v-6" /></svg></button>
        <button class="xlsx-ribbon__icon-button" data-testid="xlsx-ribbon-comment" :title="m.xlsx.authoring.comment.title" :aria-label="m.xlsx.authoring.comment.title" :disabled="!hasActiveCell" @click="emit('authoring', 'comment')"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 5h16v10H9l-5 4z" /></svg></button>
        <button class="xlsx-ribbon__icon-button" data-testid="xlsx-ribbon-hyperlink" :title="m.xlsx.authoring.hyperlink.title" :aria-label="m.xlsx.authoring.hyperlink.title" :disabled="!hasActiveCell" @click="emit('authoring', 'hyperlink')"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M10 14a4 4 0 0 1 0-5l3-3a4 4 0 0 1 5 5l-2 2M14 10a4 4 0 0 1 0 5l-3 3a4 4 0 0 1-5-5l2-2" /></svg></button>
        <button class="xlsx-ribbon__icon-button" data-testid="xlsx-ribbon-table" :title="m.xlsx.authoring.table.title" :aria-label="m.xlsx.authoring.table.title" :disabled="!hasSelection" @click="emit('authoring', 'table')"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="5" width="16" height="14" /><path d="M4 10h16M10 10v9M4 15h16" /></svg></button>
        <button class="xlsx-ribbon__icon-button" data-testid="xlsx-ribbon-conditional" :title="m.xlsx.authoring.conditional.title" :aria-label="m.xlsx.authoring.conditional.title" :disabled="!hasSelection" @click="emit('authoring', 'conditional')"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="4" width="7" height="7" /><rect x="13" y="4" width="7" height="7" /><rect x="4" y="13" width="7" height="7" /><rect x="13" y="13" width="7" height="7" /></svg></button>
        <button class="xlsx-ribbon__icon-button" data-testid="xlsx-ribbon-validation" :title="m.xlsx.authoring.validation.title" :aria-label="m.xlsx.authoring.validation.title" :disabled="!hasSelection" @click="emit('authoring', 'validation')"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="m5 12 4 4 10-10" /></svg></button>
        <button class="xlsx-ribbon__icon-button" :class="{ active: isFrozen }" data-testid="xlsx-ribbon-freeze" :aria-pressed="isFrozen" :title="isFrozen ? m.xlsx.authoring.freeze.unfreeze : m.xlsx.authoring.freeze.freezeHere" :aria-label="isFrozen ? m.xlsx.authoring.freeze.unfreeze : m.xlsx.authoring.freeze.freezeHere" :disabled="!hasActiveCell" @click="toggleFreeze"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v18M6 6l12 12M18 6 6 18M4 12h16" /></svg></button>
        <button class="xlsx-ribbon__action-button" data-testid="xlsx-ribbon-names" :title="m.xlsx.ribbon.nameManager" @click="emit('authoring', 'names')">{{ m.xlsx.ribbon.nameManager }}</button>
      </div>
    </fieldset>

    <label class="xlsx-ribbon__readonly" :title="m.xlsx.ribbon.readOnlyTitle">
      <input data-testid="xlsx-ribbon-read-only" type="checkbox" :checked="isReadOnly" @change="emit('update:readOnly', ($event.target as HTMLInputElement).checked)" />
      <span>{{ m.common.readOnly }}</span>
    </label>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, watch } from "vue";
import {
  resolveWorkbookColor,
  resolveWorkbookFillStyle,
  type XlsxBorderPreset,
  type XlsxCellAlignmentInput,
  type XlsxCellFontStyleInput,
  type XlsxCellStyleInput,
  type XlsxResolvedCellStyle,
  type XlsxViewerController,
} from "@arcships/xlsx-core";
import { classifyNumberFormat } from "../composables/cell-input";
import { DEFAULT_BORDER_COLOR_HEX, XLSX_BORDER_PRESETS } from "../composables/border-presets";
import { useXlsxMessages } from "../i18n";
import type { XlsxAuthoringDialogKind } from "./authoring-dialog";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
  readOnly?: boolean;
}>();

const emit = defineEmits<{
  authoring: [kind: XlsxAuthoringDialogKind];
  "update:readOnly": [value: boolean];
}>();

const fontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];
const m = useXlsxMessages();
const textColor = ref("#18181b");
const fillColor = ref("#fff2cc");
const borderColor = ref(`#${DEFAULT_BORDER_COLOR_HEX.toLowerCase()}`);
// The colour picker re-applies the preset it belongs to, so it follows the
// last border the user asked for rather than the dropdown's placeholder.
const lastBorderPreset = ref<XlsxBorderPreset>("all");
const borderPresets = XLSX_BORDER_PRESETS;
const borderLabels = computed<Record<XlsxBorderPreset, string>>(() => ({
  all: m.value.xlsx.ribbon.allBorders,
  bottom: m.value.xlsx.ribbon.bottomBorder,
  left: m.value.xlsx.ribbon.leftBorder,
  none: m.value.xlsx.ribbon.noBorders,
  outside: m.value.xlsx.ribbon.outsideBorders,
  right: m.value.xlsx.ribbon.rightBorder,
  top: m.value.xlsx.ribbon.topBorder
}));
const hasWorkbook = computed(() => props.controller.tabs.length > 0);
const hasSelection = computed(() => Boolean(props.controller.selection));
const hasActiveCell = computed(() => Boolean(props.controller.activeCell));
const isFrozen = computed(() => Boolean(props.controller.freezePanes));
const isReadOnly = computed(() => props.readOnly === true || props.controller.isEditingLocked === true || (props.readOnly === undefined && props.controller.readOnly));

function asRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" ? value as Record<string, unknown> : {};
}

const activeCellStyle = computed<XlsxResolvedCellStyle>(() => {
  // The controller revision is the public invalidation signal after a style edit.
  props.controller.revision;
  const cell = props.controller.activeCell;
  const sheet = props.controller.activeSheet;
  if (!cell || !sheet) return {};

  if (!props.controller.isWorkerBacked) {
    const worksheet = props.controller.getActiveWorksheet();
    if (worksheet) {
      try {
        const style = worksheet.getCellStyleAt(cell.row, cell.col) as XlsxResolvedCellStyle | null | undefined;
        if (style) return style;
      } catch {
        // Row/column styles below remain a useful public fallback.
      } finally {
        worksheet.free?.();
      }
    }
  }

  const rowStyleId = sheet.rowStyleIds[cell.row];
  const columnStyleId = sheet.colStyleIds[cell.col];
  return sheet.styleById[rowStyleId ?? columnStyleId ?? -1] ?? {};
});

const activeFont = computed(() => asRecord(activeCellStyle.value.font));
const activeAlignment = computed(() => asRecord(activeCellStyle.value.alignment));
const fontState = computed(() => ({
  bold: activeFont.value.bold === true,
  italic: activeFont.value.italic === true,
  strikethrough: activeFont.value.strikethrough === true,
  underline: Boolean(activeFont.value.underline && activeFont.value.underline !== "none"),
}));
const alignmentState = computed(() => ({
  horizontal: typeof activeAlignment.value.horizontal === "string" ? activeAlignment.value.horizontal : "general",
  vertical: typeof activeAlignment.value.vertical === "string" ? activeAlignment.value.vertical : "bottom",
  wrapText: activeAlignment.value.wrapText === true,
}));
const fontFamilyValue = computed(() => typeof activeFont.value.name === "string" && activeFont.value.name.trim()
  ? activeFont.value.name
  : "Calibri");
const fontSizeValue = computed(() => typeof activeFont.value.size === "number" && Number.isFinite(activeFont.value.size)
  ? activeFont.value.size
  : 11);
const numberFormatValue = computed(() => {
  // Built-in formats may arrive as an id without a formatString (e.g. numFmtId
  // 14), and colour sections such as `[Red]` must not read as date tokens.
  switch (classifyNumberFormat(asRecord(activeCellStyle.value.numberFormat))) {
    case "percent":
      return "percent";
    case "currency":
      return "currency";
    case "date":
    case "time":
      return "date";
    case "text":
      return "text";
    case "general":
      return "general";
    default:
      return "number";
  }
});
const selectedTextColor = computed(() => resolveWorkbookColor(
  asRecord(activeFont.value.color),
  props.controller.activeSheet?.themePalette,
) ?? "#18181b");
const selectedFillColor = computed(() => resolveWorkbookFillStyle(
  asRecord(activeCellStyle.value.fill),
  props.controller.activeSheet?.themePalette,
).backgroundColor ?? "#ffffff");

watch(selectedTextColor, (value) => { textColor.value = value; }, { immediate: true });
watch(selectedFillColor, (value) => { fillColor.value = value; }, { immediate: true });

const numberFormats = {
  general: { numberFormat: { formatType: "general" as const } },
  number: { numberFormat: { formatType: "custom" as const, formatString: "#,##0.00" } },
  currency: { numberFormat: { formatType: "custom" as const, formatString: "$#,##0.00" } },
  percent: { numberFormat: { formatType: "custom" as const, formatString: "0.00%" } },
  date: { numberFormat: { formatType: "custom" as const, formatString: "m/d/yyyy" } },
  text: { numberFormat: { formatType: "custom" as const, formatString: "@" } },
};

function applyStyle(style: XlsxCellStyleInput): void {
  if (isReadOnly.value || !props.controller.activeCell) return;
  if (props.controller.selection) props.controller.setRangeStyle(props.controller.selection, style);
  else props.controller.setSelectedCellStyle(style);
}

function toggleFreeze(): void {
  if (isReadOnly.value) return;
  props.controller.toggleFreezePanesAtActiveCell?.();
}

function applyFont(font: XlsxCellFontStyleInput): void { applyStyle({ font }); }
function applyAlignment(alignment: XlsxCellAlignmentInput): void { applyStyle({ alignment }); }

function setFontFamily(event: Event): void {
  applyFont({ name: (event.target as HTMLSelectElement).value });
}

function setFontSize(event: Event): void {
  applyFont({ size: Number((event.target as HTMLSelectElement).value) });
}

function applyTextColor(): void {
  applyFont({ color: { colorType: "rgb", hex: textColor.value.slice(1).toUpperCase() } });
}

function applyFillColor(): void {
  applyStyle({ fill: { fillType: "solid", color: { colorType: "rgb", hex: fillColor.value.slice(1).toUpperCase() } } });
}

function setNumberFormat(event: Event): void {
  const key = (event.target as HTMLSelectElement).value as keyof typeof numberFormats;
  applyStyle(numberFormats[key] ?? numberFormats.general);
}

function applyBorders(preset: XlsxBorderPreset = "all"): void {
  if (isReadOnly.value) return;
  lastBorderPreset.value = preset;
  const color = { colorType: "rgb" as const, hex: borderColor.value.slice(1).toUpperCase() };
  if (typeof props.controller.setSelectionBorders === "function") {
    props.controller.setSelectionBorders({ preset, color });
    return;
  }
  // Controllers proxied from an older host keep the plain All-borders path.
  if (preset !== "all") return;
  const edge = { style: "thin" as const, color };
  applyStyle({ border: { top: edge, right: edge, bottom: edge, left: edge } });
}

function applyBorderPreset(event: Event): void {
  const select = event.target as HTMLSelectElement | null;
  const preset = (select?.value ?? "") as XlsxBorderPreset | "";
  // The first option is the dropdown's placeholder, not a command.
  if (!preset) return;
  applyBorders(preset);
  if (select) select.value = "";
}

function applyBorderColor(): void {
  applyBorders(lastBorderPreset.value);
}

function mergeAndCenter(): void {
  if (isReadOnly.value) return;
  props.controller.mergeAndCenterSelection?.();
}

function copySelection(): void {
  if (isReadOnly.value) return;
  void props.controller.copySelectionToClipboard?.();
}

function cutSelection(): void {
  if (isReadOnly.value) return;
  props.controller.cutSelection?.();
}

function insertRow(position: "above" | "below"): void {
  if (isReadOnly.value || !props.controller.activeCell) return;
  props.controller.insertRows?.(props.controller.activeCell.row + (position === "below" ? 1 : 0));
}

function deleteRow(): void {
  if (isReadOnly.value || !props.controller.activeCell) return;
  props.controller.deleteRows?.(props.controller.activeCell.row);
}

function insertColumn(position: "left" | "right"): void {
  if (isReadOnly.value || !props.controller.activeCell) return;
  props.controller.insertColumns?.(props.controller.activeCell.col + (position === "right" ? 1 : 0));
}

function deleteColumn(): void {
  if (isReadOnly.value || !props.controller.activeCell) return;
  props.controller.deleteColumns?.(props.controller.activeCell.col);
}
</script>

<style scoped>
.xlsx-ribbon {
  align-items: center;
  background: #fff;
  border-bottom: 1px solid #e4e4e7;
  color: #18181b;
  display: flex;
  flex: 0 0 auto;
  gap: 8px;
  min-height: 48px;
  padding: 7px 10px;
  user-select: none;
}
.xlsx-ribbon--dark { background: #18181b; border-color: #3f3f46; color: #f4f4f5; }
.xlsx-ribbon__controls { align-items: center; border: 0; display: flex; flex: 1; flex-wrap: wrap; gap: 4px 6px; margin: 0; min-inline-size: 0; overflow: visible; padding: 0; }
.xlsx-ribbon__group { align-items: center; display: flex; flex: 0 0 auto; gap: 3px; }
.xlsx-ribbon__separator { background: #e4e4e7; flex: 0 0 auto; height: 20px; width: 1px; }
.xlsx-ribbon button,
.xlsx-ribbon select { color: inherit; font: inherit; }
.xlsx-ribbon select { background: transparent; border: 1px solid #d4d4d8; border-radius: 6px; font-size: 11px; height: 30px; max-width: 118px; padding: 0 7px; }
.xlsx-ribbon__size { width: 52px; }
.xlsx-ribbon__number { width: 86px; }
.xlsx-ribbon__borders { width: 104px; }
.xlsx-ribbon__icon-button,
.xlsx-ribbon__text-button,
.xlsx-ribbon__action-button,
.xlsx-ribbon__color {
  align-items: center;
  background: transparent;
  border: 1px solid transparent;
  border-radius: 6px;
  box-sizing: border-box;
  cursor: pointer;
  display: inline-flex;
  height: 30px;
  justify-content: center;
}
.xlsx-ribbon__icon-button,
.xlsx-ribbon__text-button,
.xlsx-ribbon__color { padding: 0; width: 30px; }
.xlsx-ribbon__action-button { font-size: 11px; padding: 0 8px; }
.xlsx-ribbon button:hover:not(:disabled),
.xlsx-ribbon__color:hover { background: rgba(113, 113, 122, .12); }
.xlsx-ribbon button.active { background: rgba(37, 99, 235, .14); color: #1d4ed8; }
.xlsx-ribbon button:disabled,
.xlsx-ribbon__controls:disabled { cursor: default; opacity: .45; }
.xlsx-ribbon svg { fill: none; height: 16px; stroke: currentColor; stroke-linecap: round; stroke-linejoin: round; stroke-width: 1.7; width: 16px; }
.xlsx-ribbon__color { flex-direction: column; font-size: 13px; font-weight: 700; position: relative; }
.xlsx-ribbon__color i { border-radius: 2px; height: 3px; width: 17px; }
.xlsx-ribbon__color input { cursor: pointer; inset: 0; opacity: 0; position: absolute; }
.xlsx-ribbon__readonly { align-items: center; color: #71717a; display: flex; flex: 0 0 auto; font-size: 11px; gap: 5px; }
.xlsx-ribbon__readonly input { accent-color: #2563eb; }
.xlsx-ribbon button:focus-visible,
.xlsx-ribbon select:focus-visible,
.xlsx-ribbon__color:focus-within { outline: 2px solid #2563eb; outline-offset: 2px; }

@container xlsx-viewer (max-width: 560px) {
  .xlsx-ribbon__readonly span { display: none; }
}

@media (pointer: coarse) {
  .xlsx-ribbon button,
  .xlsx-ribbon select,
  .xlsx-ribbon input {
    min-height: 44px;
  }
  .xlsx-ribbon__icon-button,
  .xlsx-ribbon__text-button {
    min-width: 44px;
  }
}
</style>
