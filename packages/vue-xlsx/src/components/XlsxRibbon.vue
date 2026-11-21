<template>
  <OfficeThemeProvider :theme="isDark ? 'dark' : 'light'">
    <!-- `xlsx-ribbon--dark` remains the dark-mode marker the blackbox suite greps for. -->
    <OfficeToolbar
      :class="ribbonClasses"
      data-testid="xlsx-format-toolbar"
    >
      <OfficeToolbarGroup
        gap="sm"
        class="xlsx-ribbon__controls"
      >
        <OfficeToolbarGroup gap="sm" class="xlsx-ribbon__group">
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-undo"
            :label="m.xlsx.ribbon.undo"
            size="icon-sm"
            :disabled="controlsDisabled || !controller.canUndo"
            @click="controller.undo()"
          ><HugeIcon :icon="Undo02Icon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-redo"
            :label="m.xlsx.ribbon.redo"
            size="icon-sm"
            :disabled="controlsDisabled || !controller.canRedo"
            @click="controller.redo()"
          ><HugeIcon :icon="Redo02Icon" class="xlsx-ribbon__icon" /></OfficeIconButton>
      </OfficeToolbarGroup>

        <OfficeToolbarSeparator class="xlsx-ribbon__separator" />

        <OfficeToolbarGroup gap="sm" class="xlsx-ribbon__group">
          <OfficeButton
            data-testid="xlsx-ribbon-cut"
            class="xlsx-ribbon__action-button"
            variant="ghost"
            size="sm"
            :title="m.xlsx.ribbon.cut"
            :disabled="controlsDisabled || !hasSelection"
            @click="cutSelection"
          >{{ m.xlsx.ribbon.cut }}</OfficeButton>
          <OfficeButton
            data-testid="xlsx-ribbon-copy"
            class="xlsx-ribbon__action-button"
            variant="ghost"
            size="sm"
            :title="m.xlsx.ribbon.copy"
            :disabled="controlsDisabled || !hasSelection"
            @click="copySelection"
          >{{ m.xlsx.ribbon.copy }}</OfficeButton>
          <OfficeButton
            data-testid="xlsx-ribbon-paste"
            class="xlsx-ribbon__action-button"
            variant="ghost"
            size="sm"
            :title="m.xlsx.ribbon.paste"
            :disabled="controlsDisabled || !hasActiveCell"
            @click="controller.pasteFromClipboard()"
          >{{ m.xlsx.ribbon.paste }}</OfficeButton>
          <OfficeButton
            data-testid="xlsx-ribbon-paste-special"
            class="xlsx-ribbon__action-button"
            variant="ghost"
            size="sm"
            :title="m.xlsx.ribbon.pasteSpecial"
            :disabled="controlsDisabled || !hasSelection"
            @click="emit('authoring', 'pasteSpecial')"
          >{{ m.xlsx.ribbon.pasteSpecial }}</OfficeButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator class="xlsx-ribbon__separator" />

        <OfficeToolbarGroup gap="sm" class="xlsx-ribbon__group">
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-ribbon-insert-row-above"
            :label="m.xlsx.ribbon.insertRowAbove"
            size="icon-sm"
            :disabled="controlsDisabled || !hasActiveCell"
            @click="insertRow('above')"
          ><HugeIcon :icon="InsertRowUpIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-ribbon-insert-row-below"
            :label="m.xlsx.ribbon.insertRowBelow"
            size="icon-sm"
            :disabled="controlsDisabled || !hasActiveCell"
            @click="insertRow('below')"
          ><HugeIcon :icon="InsertRowDownIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-ribbon-delete-row"
            :label="m.xlsx.ribbon.deleteRow"
            size="icon-sm"
            :disabled="controlsDisabled || !hasActiveCell"
            @click="deleteRow"
          ><HugeIcon :icon="DeleteRowIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-ribbon-insert-column-left"
            :label="m.xlsx.ribbon.insertColumnLeft"
            size="icon-sm"
            :disabled="controlsDisabled || !hasActiveCell"
            @click="insertColumn('left')"
          ><HugeIcon :icon="ColumnInsertIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-ribbon-insert-column-right"
            :label="m.xlsx.ribbon.insertColumnRight"
            size="icon-sm"
            :disabled="controlsDisabled || !hasActiveCell"
            @click="insertColumn('right')"
          ><HugeIcon :icon="ColumnInsertIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-ribbon-delete-column"
            :label="m.xlsx.ribbon.deleteColumn"
            size="icon-sm"
            :disabled="controlsDisabled || !hasActiveCell"
            @click="deleteColumn"
          ><HugeIcon :icon="ColumnDeleteIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator class="xlsx-ribbon__separator" />

        <OfficeToolbarGroup gap="sm" class="xlsx-ribbon__group">
          <select
            class="office-ui xlsx-ribbon__select xlsx-ribbon__font"
            :value="fontFamilyValue"
            :aria-label="m.xlsx.ribbon.fontFamily"
            :title="m.xlsx.ribbon.fontFamily"
            :disabled="controlsDisabled"
            @change="setFontFamily"
          >
            <option value="Calibri">Calibri</option>
            <option value="Aptos">Aptos</option>
            <option value="Arial">Arial</option>
            <option value="Georgia">Georgia</option>
            <option value="Times New Roman">Times New Roman</option>
            <option value="Courier New">Courier New</option>
          </select>
          <select
            class="office-ui xlsx-ribbon__select xlsx-ribbon__size"
            :value="fontSizeValue"
            :aria-label="m.xlsx.ribbon.fontSize"
            :title="m.xlsx.ribbon.fontSize"
            :disabled="controlsDisabled"
            @change="setFontSize"
          >
            <option v-if="!fontSizes.includes(fontSizeValue)" :value="fontSizeValue">{{ fontSizeValue }}</option>
            <option v-for="size in fontSizes" :key="size" :value="size">{{ size }}</option>
          </select>
          <OfficeIconButton
            class="xlsx-ribbon__text-button"
            :active="fontState.bold"
            :aria-pressed="fontState.bold"
            :label="m.xlsx.ribbon.bold"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyFont({ bold: !fontState.bold })"
          ><HugeIcon :icon="TextBoldIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__text-button"
            :active="fontState.italic"
            :aria-pressed="fontState.italic"
            :label="m.xlsx.ribbon.italic"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyFont({ italic: !fontState.italic })"
          ><HugeIcon :icon="TextItalicIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__text-button"
            :active="fontState.underline"
            :aria-pressed="fontState.underline"
            :label="m.xlsx.ribbon.underline"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyFont({ underline: fontState.underline ? 'none' : 'single' })"
          ><HugeIcon :icon="TextUnderlineIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__text-button"
            :active="fontState.strikethrough"
            :aria-pressed="fontState.strikethrough"
            :label="m.xlsx.ribbon.strikethrough"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyFont({ strikethrough: !fontState.strikethrough })"
          ><HugeIcon :icon="TextStrikethroughIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <label class="xlsx-ribbon__color" :title="m.xlsx.ribbon.textColor">
            <span aria-hidden="true">A</span><i :style="{ backgroundColor: textColor }" />
            <input v-model="textColor" type="color" :aria-label="m.xlsx.ribbon.textColor" :disabled="controlsDisabled" @change="applyTextColor" />
          </label>
          <label class="xlsx-ribbon__color" :title="m.xlsx.ribbon.fillColor">
            <HugeIcon :icon="PaintBucketIcon" class="xlsx-ribbon__icon" /><i :style="{ backgroundColor: fillColor }" />
            <input v-model="fillColor" type="color" :aria-label="m.xlsx.ribbon.fillColor" :disabled="controlsDisabled" @change="applyFillColor" />
          </label>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator class="xlsx-ribbon__separator" />

        <OfficeToolbarGroup gap="sm" class="xlsx-ribbon__group">
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            :active="alignmentState.horizontal === 'left'"
            :aria-pressed="alignmentState.horizontal === 'left'"
            :label="m.xlsx.ribbon.alignLeft"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyAlignment({ horizontal: 'left' })"
          ><HugeIcon :icon="TextAlignLeft01Icon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            :active="alignmentState.horizontal === 'center'"
            :aria-pressed="alignmentState.horizontal === 'center'"
            :label="m.xlsx.ribbon.alignCenter"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyAlignment({ horizontal: 'center' })"
          ><HugeIcon :icon="TextAlignCenterIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            :active="alignmentState.horizontal === 'right'"
            :aria-pressed="alignmentState.horizontal === 'right'"
            :label="m.xlsx.ribbon.alignRight"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyAlignment({ horizontal: 'right' })"
          ><HugeIcon :icon="TextAlignRight01Icon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            :active="alignmentState.vertical === 'top'"
            :aria-pressed="alignmentState.vertical === 'top'"
            :label="m.xlsx.ribbon.alignTop"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyAlignment({ vertical: 'top' })"
          ><HugeIcon :icon="AlignBoxTopCenterIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            :active="alignmentState.vertical === 'center'"
            :aria-pressed="alignmentState.vertical === 'center'"
            :label="m.xlsx.ribbon.alignMiddle"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyAlignment({ vertical: 'center' })"
          ><HugeIcon :icon="AlignBoxMiddleCenterIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            :active="alignmentState.vertical === 'bottom'"
            :aria-pressed="alignmentState.vertical === 'bottom'"
            :label="m.xlsx.ribbon.alignBottom"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyAlignment({ vertical: 'bottom' })"
          ><HugeIcon :icon="AlignBoxBottomCenterIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            :active="alignmentState.wrapText"
            :aria-pressed="alignmentState.wrapText"
            :label="m.xlsx.ribbon.wrapText"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyAlignment({ wrapText: !alignmentState.wrapText })"
          ><HugeIcon :icon="TextWrapIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator class="xlsx-ribbon__separator" />

        <OfficeToolbarGroup gap="sm" class="xlsx-ribbon__group">
          <select
            class="office-ui xlsx-ribbon__select xlsx-ribbon__number"
            :value="numberFormatValue"
            :aria-label="m.xlsx.ribbon.numberFormat"
            :title="m.xlsx.ribbon.numberFormat"
            :disabled="controlsDisabled"
            @change="setNumberFormat"
          >
            <option value="general">{{ m.xlsx.ribbon.numberFormats.general }}</option>
            <option value="number">{{ m.xlsx.ribbon.numberFormats.number }}</option>
            <option value="currency">{{ m.xlsx.ribbon.numberFormats.currency }}</option>
            <option value="percent">{{ m.xlsx.ribbon.numberFormats.percent }}</option>
            <option value="date">{{ m.xlsx.ribbon.numberFormats.date }}</option>
            <option value="text">{{ m.xlsx.ribbon.numberFormats.text }}</option>
          </select>
          <OfficeIconButton
            class="xlsx-ribbon__text-button"
            :label="m.xlsx.ribbon.currency"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyStyle(numberFormats.currency)"
          ><span aria-hidden="true">$</span></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__text-button"
            :label="m.xlsx.ribbon.percent"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyStyle(numberFormats.percent)"
          ><span aria-hidden="true">%</span></OfficeIconButton>
          <OfficeButton
            data-testid="xlsx-ribbon-number-format"
            class="xlsx-ribbon__action-button"
            variant="ghost"
            size="sm"
            :title="m.xlsx.ribbon.numberFormatDialog"
            :disabled="controlsDisabled"
            @click="emit('authoring', 'numberFormat')"
          >{{ m.xlsx.ribbon.numberFormatDialog }}</OfficeButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator class="xlsx-ribbon__separator" />

        <OfficeToolbarGroup gap="sm" class="xlsx-ribbon__group">
          <OfficeButton
            data-testid="xlsx-ribbon-merge-center"
            class="xlsx-ribbon__action-button"
            variant="ghost"
            size="sm"
            :title="m.xlsx.ribbon.mergeAndCenter"
            :disabled="controlsDisabled || !hasSelection"
            @click="mergeAndCenter"
          >{{ m.xlsx.ribbon.mergeAndCenter }}</OfficeButton>
          <OfficeButton
            data-testid="xlsx-ribbon-unmerge"
            class="xlsx-ribbon__action-button"
            variant="ghost"
            size="sm"
            :disabled="controlsDisabled || !hasSelection"
            @click="controller.unmergeSelection()"
          >{{ m.xlsx.ribbon.unmerge }}</OfficeButton>
          <select
            data-testid="xlsx-ribbon-border-preset"
            class="office-ui xlsx-ribbon__select xlsx-ribbon__borders"
            :value="''"
            :aria-label="m.xlsx.ribbon.borders"
            :title="m.xlsx.ribbon.borders"
            :disabled="controlsDisabled"
            @change="applyBorderPreset"
          >
            <option value="">{{ m.xlsx.ribbon.borders }}</option>
            <option v-for="preset in borderPresets" :key="preset" :value="preset">{{ borderLabels[preset] }}</option>
          </select>
          <label class="xlsx-ribbon__color" :title="m.xlsx.ribbon.borderColor">
            <HugeIcon :icon="BorderAll01Icon" class="xlsx-ribbon__icon" /><i :style="{ backgroundColor: borderColor }" />
            <input v-model="borderColor" data-testid="xlsx-ribbon-border-color" type="color" :aria-label="m.xlsx.ribbon.borderColor" :disabled="controlsDisabled" @change="applyBorderColor" />
          </label>
          <OfficeIconButton
            data-testid="xlsx-ribbon-borders-all"
            class="xlsx-ribbon__icon-button"
            :label="m.xlsx.ribbon.allBorders"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="applyBorders('all')"
          ><HugeIcon :icon="BorderAll01Icon" class="xlsx-ribbon__icon" /></OfficeIconButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator class="xlsx-ribbon__separator" />

        <OfficeToolbarGroup gap="sm" class="xlsx-ribbon__group" data-testid="xlsx-ribbon-authoring">
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-ribbon-chart"
            :label="m.xlsx.authoring.chart.title"
            size="icon-sm"
            :disabled="controlsDisabled"
            @click="emit('authoring', 'chart')"
          ><HugeIcon :icon="ChartColumnIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-ribbon-comment"
            :label="m.xlsx.authoring.comment.title"
            size="icon-sm"
            :disabled="controlsDisabled || !hasActiveCell"
            @click="emit('authoring', 'comment')"
          ><HugeIcon :icon="Comment01Icon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-ribbon-hyperlink"
            :label="m.xlsx.authoring.hyperlink.title"
            size="icon-sm"
            :disabled="controlsDisabled || !hasActiveCell"
            @click="emit('authoring', 'hyperlink')"
          ><HugeIcon :icon="Link02Icon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-ribbon-table"
            :label="m.xlsx.authoring.table.title"
            size="icon-sm"
            :disabled="controlsDisabled || !hasSelection"
            @click="emit('authoring', 'table')"
          ><HugeIcon :icon="Table01Icon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-ribbon-conditional"
            :label="m.xlsx.authoring.conditional.title"
            size="icon-sm"
            :disabled="controlsDisabled || !hasSelection"
            @click="emit('authoring', 'conditional')"
          ><HugeIcon :icon="Grid2X2Icon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            data-testid="xlsx-ribbon-validation"
            :label="m.xlsx.authoring.validation.title"
            size="icon-sm"
            :disabled="controlsDisabled || !hasSelection"
            @click="emit('authoring', 'validation')"
          ><HugeIcon :icon="CheckmarkCircle01Icon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeIconButton
            class="xlsx-ribbon__icon-button"
            :active="isFrozen"
            :aria-pressed="isFrozen"
            data-testid="xlsx-ribbon-freeze"
            :label="isFrozen ? m.xlsx.authoring.freeze.unfreeze : m.xlsx.authoring.freeze.freezeHere"
            size="icon-sm"
            :disabled="controlsDisabled || !hasActiveCell"
            @click="toggleFreeze"
          ><HugeIcon :icon="SnowflakeIcon" class="xlsx-ribbon__icon" /></OfficeIconButton>
          <OfficeButton
            data-testid="xlsx-ribbon-names"
            class="xlsx-ribbon__action-button"
            variant="ghost"
            size="sm"
            :title="m.xlsx.ribbon.nameManager"
            :disabled="controlsDisabled"
            @click="emit('authoring', 'names')"
          >{{ m.xlsx.ribbon.nameManager }}</OfficeButton>
          <OfficeButton
            data-testid="xlsx-ribbon-page-setup"
            class="xlsx-ribbon__action-button"
            variant="ghost"
            size="sm"
            :title="m.xlsx.authoring.pageSetup.title"
            :disabled="controlsDisabled"
            @click="emit('authoring', 'pageSetup')"
          >{{ m.xlsx.authoring.pageSetup.title }}</OfficeButton>
        </OfficeToolbarGroup>

        <label class="xlsx-ribbon__readonly" :title="m.xlsx.ribbon.readOnlyTitle">
          <input data-testid="xlsx-ribbon-read-only" type="checkbox" :checked="isReadOnly" @change="emit('update:readOnly', ($event.target as HTMLInputElement).checked)" />
          <span>{{ m.common.readOnly }}</span>
        </label>
      </OfficeToolbarGroup>
    </OfficeToolbar>
  </OfficeThemeProvider>
</template>

<script setup lang="ts">
import { computed, ref, watch } from "vue";
import {
  AlignBoxBottomCenterIcon,
  AlignBoxMiddleCenterIcon,
  AlignBoxTopCenterIcon,
  BorderAll01Icon,
  ChartColumnIcon,
  CheckmarkCircle01Icon,
  ColumnDeleteIcon,
  ColumnInsertIcon,
  Comment01Icon,
  DeleteRowIcon,
  Grid2X2Icon,
  InsertRowDownIcon,
  InsertRowUpIcon,
  Link02Icon,
  PaintBucketIcon,
  Redo02Icon,
  SnowflakeIcon,
  Table01Icon,
  TextAlignCenterIcon,
  TextAlignLeft01Icon,
  TextAlignRight01Icon,
  TextBoldIcon,
  TextItalicIcon,
  TextStrikethroughIcon,
  TextUnderlineIcon,
  TextWrapIcon,
  Undo02Icon,
} from "@hugeicons/core-free-icons";
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
import {
  HugeIcon,
  OfficeButton,
  OfficeIconButton,
  OfficeThemeProvider,
  OfficeToolbar,
  OfficeToolbarGroup,
  OfficeToolbarSeparator,
} from "@arcships/vue-ui";
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
const ribbonClasses = computed(() =>
  props.isDark ? "xlsx-ribbon xlsx-ribbon--dark" : "xlsx-ribbon",
);
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
// Replaces the disabled `<fieldset>` wrapper: every ribbon control honours the
// same "no workbook / read-only" gate the fieldset used to enforce.
const controlsDisabled = computed(() => isReadOnly.value || !hasWorkbook.value);

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
/*
 * Geometry comes from `OfficeToolbar` / `OfficeToolbarGroup`: the ribbon is a
 * dense second toolbar row (48px min height, 4px cluster gap, 8px group gap).
 * Only spreadsheet-specific pieces (native selects, colour wells, the
 * read-only switch) need local rules, and every colour is an Office token.
 */
.xlsx-ribbon {
  --office-ui-toolbar-min-height: 3rem;
  flex: 0 0 auto;
  --office-ui-toolbar-padding-block: 6px;
  border-block-start: var(--office-border-width) solid var(--office-border);
}

.xlsx-ribbon__controls {
  flex: 1 1 auto;
  flex-wrap: wrap;
  min-inline-size: 0;
}

.xlsx-ribbon__group {
  /* Shrinkable so a 390px host lets each cluster wrap instead of overflowing. */
  flex: 0 1 auto;
  min-inline-size: 0;
}

.xlsx-ribbon__separator {
  block-size: 20px;
  margin-inline: 2px;
}

.xlsx-ribbon button,
.xlsx-ribbon select {
  color: inherit;
  font: inherit;
}

.xlsx-ribbon__select {
  background: var(--office-background);
  border: var(--office-border-width) solid var(--office-input);
  border-radius: var(--office-radius-sm);
  color: var(--office-foreground);
  font-size: var(--office-text-xs);
  height: 28px;
  max-inline-size: 118px;
  padding-block: 0;
  padding-inline: 6px;
}

.xlsx-ribbon__select:focus-visible {
  border-color: var(--office-ring);
  outline: 2px solid color-mix(in oklab, var(--office-ring) 40%, transparent);
  outline-offset: var(--office-ring-offset-width);
}

.xlsx-ribbon__font { inline-size: 112px; }
.xlsx-ribbon__size { inline-size: 52px; }
.xlsx-ribbon__number { inline-size: 92px; }
.xlsx-ribbon__borders { inline-size: 104px; }

.xlsx-ribbon__text-button span {
  font-size: var(--office-text-sm);
  line-height: 1;
}

/*
 * Hugeicons render the upstream glyph data unchanged: the shared Office button
 * already supplies the 16px box, the 2px optical margin and the 80% opacity
 * that Extend's `[&_svg]` utilities apply, so no stroke-width override here.
 */
.xlsx-ribbon__icon {
  display: block;
  flex-shrink: 0;
}

.xlsx-ribbon__action-button {
  font-size: var(--office-text-xs);
  padding-inline: var(--office-ui-space-2, 0.5rem);
}

.xlsx-ribbon__color {
  align-items: center;
  border: var(--office-border-width) solid transparent;
  border-radius: var(--office-radius-sm);
  box-sizing: border-box;
  cursor: pointer;
  display: inline-flex;
  flex-direction: column;
  font-size: var(--office-text-sm);
  font-weight: 700;
  block-size: 28px;
  inline-size: 28px;
  justify-content: center;
  position: relative;
}

.xlsx-ribbon__color:hover {
  background: var(--office-accent);
}

.xlsx-ribbon__color:focus-within {
  outline: 2px solid var(--office-ring);
  outline-offset: var(--office-ring-offset-width);
}

.xlsx-ribbon__color svg {
  block-size: 14px;
  inline-size: 14px;
}

.xlsx-ribbon__color i {
  border-radius: 2px;
  block-size: 3px;
  inline-size: 17px;
}

.xlsx-ribbon__color input {
  cursor: pointer;
  inset: 0;
  opacity: 0;
  position: absolute;
}

.xlsx-ribbon__readonly {
  align-items: center;
  color: var(--office-muted-foreground);
  display: flex;
  flex: 0 0 auto;
  font-size: var(--office-text-xs);
  gap: 5px;
  white-space: nowrap;
}

.xlsx-ribbon__readonly input {
  accent-color: var(--office-primary);
}

/* Tablet/narrow hosts: wrapped clusters stay scrollable (about four rows)
 * instead of eating the grid height. */
@container xlsx-viewer (max-width: 900px) {
  .xlsx-ribbon {
    max-block-size: 132px;
    overflow-y: auto;
  }
}

@container xlsx-viewer (max-width: 560px) {
  .xlsx-ribbon__readonly span {
    display: none;
  }
}

@media (pointer: coarse) {
  .xlsx-ribbon :deep(.office-button),
  .xlsx-ribbon select,
  .xlsx-ribbon .xlsx-ribbon__color {
    min-block-size: 44px;
  }

  .xlsx-ribbon :deep(.office-button--size-icon-sm) {
    min-inline-size: 44px;
  }

  .xlsx-ribbon .xlsx-ribbon__color {
    min-inline-size: 44px;
  }
}
</style>
