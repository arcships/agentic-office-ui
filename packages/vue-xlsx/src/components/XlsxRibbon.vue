<template>
  <div class="xlsx-ribbon" :class="{ 'xlsx-ribbon--dark': isDark }" data-testid="xlsx-format-toolbar">
    <fieldset class="xlsx-ribbon__controls" :disabled="isReadOnly || !hasWorkbook">
      <div class="xlsx-ribbon__group">
        <button data-testid="xlsx-undo" class="xlsx-ribbon__icon-button" :disabled="!controller.canUndo" title="Undo" aria-label="Undo" @click="controller.undo()">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 7 4 12l5 5M5 12h8a6 6 0 0 1 6 6" /></svg>
        </button>
        <button data-testid="xlsx-redo" class="xlsx-ribbon__icon-button" :disabled="!controller.canRedo" title="Redo" aria-label="Redo" @click="controller.redo()">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m15 7 5 5-5 5M19 12h-8a6 6 0 0 0-6 6" /></svg>
        </button>
      </div>
      <span class="xlsx-ribbon__separator" />
      <div class="xlsx-ribbon__group">
        <select :value="fontFamilyValue" aria-label="Font family" title="Font family" @change="setFontFamily">
          <option value="Calibri">Calibri</option>
          <option value="Aptos">Aptos</option>
          <option value="Arial">Arial</option>
          <option value="Georgia">Georgia</option>
          <option value="Times New Roman">Times New Roman</option>
          <option value="Courier New">Courier New</option>
        </select>
        <select :value="fontSizeValue" class="xlsx-ribbon__size" aria-label="Font size" title="Font size" @change="setFontSize">
          <option v-if="!fontSizes.includes(fontSizeValue)" :value="fontSizeValue">{{ fontSizeValue }}</option>
          <option v-for="size in fontSizes" :key="size" :value="size">{{ size }}</option>
        </select>
        <button class="xlsx-ribbon__text-button" :class="{ active: fontState.bold }" :aria-pressed="fontState.bold" title="Bold" aria-label="Bold" @click="applyFont({ bold: !fontState.bold })"><strong>B</strong></button>
        <button class="xlsx-ribbon__text-button" :class="{ active: fontState.italic }" :aria-pressed="fontState.italic" title="Italic" aria-label="Italic" @click="applyFont({ italic: !fontState.italic })"><em>I</em></button>
        <button class="xlsx-ribbon__text-button" :class="{ active: fontState.underline }" :aria-pressed="fontState.underline" title="Underline" aria-label="Underline" @click="applyFont({ underline: fontState.underline ? 'none' : 'single' })"><u>U</u></button>
        <button class="xlsx-ribbon__text-button" :class="{ active: fontState.strikethrough }" :aria-pressed="fontState.strikethrough" title="Strikethrough" aria-label="Strikethrough" @click="applyFont({ strikethrough: !fontState.strikethrough })"><s>S</s></button>
        <label class="xlsx-ribbon__color" title="Text color">
          <span>A</span><i :style="{ backgroundColor: textColor }" />
          <input v-model="textColor" type="color" aria-label="Text color" @change="applyTextColor" />
        </label>
        <label class="xlsx-ribbon__color" title="Fill color">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m7 4 10 10-5 5-7-7zM14 6l2-2 4 4-2 2" /></svg><i :style="{ backgroundColor: fillColor }" />
          <input v-model="fillColor" type="color" aria-label="Fill color" @change="applyFillColor" />
        </label>
      </div>
      <span class="xlsx-ribbon__separator" />
      <div class="xlsx-ribbon__group">
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.horizontal === 'left' }" :aria-pressed="alignmentState.horizontal === 'left'" title="Align left" aria-label="Align left" @click="applyAlignment({ horizontal: 'left' })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 10h11M4 14h16M4 18h9" /></svg></button>
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.horizontal === 'center' }" :aria-pressed="alignmentState.horizontal === 'center'" title="Align center" aria-label="Align center" @click="applyAlignment({ horizontal: 'center' })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M7 10h10M4 14h16M8 18h8" /></svg></button>
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.horizontal === 'right' }" :aria-pressed="alignmentState.horizontal === 'right'" title="Align right" aria-label="Align right" @click="applyAlignment({ horizontal: 'right' })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M9 10h11M4 14h16M11 18h9" /></svg></button>
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.vertical === 'top' }" :aria-pressed="alignmentState.vertical === 'top'" title="Align top" aria-label="Align top" @click="applyAlignment({ vertical: 'top' })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 4h14M7 8h10M9 12h6" /></svg></button>
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.vertical === 'center' }" :aria-pressed="alignmentState.vertical === 'center'" title="Align middle" aria-label="Align middle" @click="applyAlignment({ vertical: 'center' })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 5h14M7 9h10M7 15h10M5 19h14" /></svg></button>
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.vertical === 'bottom' }" :aria-pressed="alignmentState.vertical === 'bottom'" title="Align bottom" aria-label="Align bottom" @click="applyAlignment({ vertical: 'bottom' })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 12h6M7 16h10M5 20h14" /></svg></button>
        <button class="xlsx-ribbon__icon-button" :class="{ active: alignmentState.wrapText }" :aria-pressed="alignmentState.wrapText" title="Wrap text" aria-label="Wrap text" @click="applyAlignment({ wrapText: !alignmentState.wrapText })"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 10h12a3 3 0 0 1 0 6h-3M16 13l-3 3 3 3M4 20h7" /></svg></button>
      </div>
      <span class="xlsx-ribbon__separator" />
      <div class="xlsx-ribbon__group">
        <select :value="numberFormatValue" class="xlsx-ribbon__number" aria-label="Number format" title="Number format" @change="setNumberFormat">
          <option value="general">General</option>
          <option value="number">Number</option>
          <option value="currency">Currency</option>
          <option value="percent">Percent</option>
          <option value="date">Date</option>
        </select>
        <button class="xlsx-ribbon__text-button" title="Currency" aria-label="Currency" @click="applyStyle(numberFormats.currency)">$</button>
        <button class="xlsx-ribbon__text-button" title="Percent" aria-label="Percent" @click="applyStyle(numberFormats.percent)">%</button>
      </div>
      <span class="xlsx-ribbon__separator" />
      <div class="xlsx-ribbon__group">
        <button class="xlsx-ribbon__action-button" :disabled="!hasSelection" @click="controller.mergeSelection()">Merge</button>
        <button class="xlsx-ribbon__action-button" :disabled="!hasSelection" @click="controller.unmergeSelection()">Unmerge</button>
        <button class="xlsx-ribbon__icon-button" title="All borders" aria-label="All borders" @click="applyBorders"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="4" width="16" height="16" /><path d="M12 4v16M4 12h16" /></svg></button>
      </div>
    </fieldset>

    <label class="xlsx-ribbon__readonly" title="Read-only mode">
      <input data-testid="xlsx-ribbon-read-only" type="checkbox" :checked="isReadOnly" @change="emit('update:readOnly', ($event.target as HTMLInputElement).checked)" />
      <span>Read only</span>
    </label>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, watch } from "vue";
import {
  resolveWorkbookColor,
  resolveWorkbookFillStyle,
  type XlsxCellAlignmentInput,
  type XlsxCellFontStyleInput,
  type XlsxCellStyleInput,
  type XlsxResolvedCellStyle,
  type XlsxViewerController,
} from "@arcships/xlsx-core";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
  readOnly?: boolean;
}>();

const emit = defineEmits<{
  "update:readOnly": [value: boolean];
}>();

const fontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];
const textColor = ref("#18181b");
const fillColor = ref("#fff2cc");
const hasWorkbook = computed(() => props.controller.tabs.length > 0);
const hasSelection = computed(() => Boolean(props.controller.selection));
const isReadOnly = computed(() => props.readOnly ?? props.controller.readOnly);

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
  const numberFormat = asRecord(activeCellStyle.value.numberFormat);
  const format = String(numberFormat.formatString ?? "").toLowerCase();
  if (format.includes("%")) return "percent";
  if (/[$€£¥]/.test(format)) return "currency";
  if (/[dmy]/.test(format)) return "date";
  if (format && format !== "general") return "number";
  return "general";
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
};

function applyStyle(style: XlsxCellStyleInput): void {
  if (isReadOnly.value || !props.controller.activeCell) return;
  if (props.controller.selection) props.controller.setRangeStyle(props.controller.selection, style);
  else props.controller.setSelectedCellStyle(style);
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

function applyBorders(): void {
  const edge = { style: "thin" as const, color: { colorType: "rgb" as const, hex: "A1A1AA" } };
  applyStyle({ border: { top: edge, right: edge, bottom: edge, left: edge } });
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

@media (max-width: 560px) {
  .xlsx-ribbon__readonly span { display: none; }
}
</style>
