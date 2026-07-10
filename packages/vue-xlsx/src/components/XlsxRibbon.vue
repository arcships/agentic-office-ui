<template>
  <div class="xlsx-ribbon" :style="ribbonStyle">
    <div class="xlsx-ribbon__tabs">
      <button
        v-for="tab in RIBBON_TABS"
        :key="tab"
        class="xlsx-ribbon__tab"
        :class="{ 'xlsx-ribbon__tab--active': activeTab === tab }"
        @click="activeTab = tab"
      >
        {{ tab }}
      </button>
    </div>
    <div class="xlsx-ribbon__content">
      <!-- Home -->
      <template v-if="activeTab === 'Home'">
        <div class="xlsx-ribbon__group">
          <div class="xlsx-ribbon__group-label">Clipboard</div>
          <div class="xlsx-ribbon__group-buttons">
            <button
              class="xlsx-ribbon__btn"
              data-testid="xlsx-undo"
              :disabled="!controller.canUndo"
              title="Undo"
              @click="controller.undo()"
            >↩</button>
            <button
              class="xlsx-ribbon__btn"
              data-testid="xlsx-redo"
              :disabled="!controller.canRedo"
              title="Redo"
              @click="controller.redo()"
            >↪</button>
          </div>
        </div>
        <div class="xlsx-ribbon__group">
          <div class="xlsx-ribbon__group-label">Font</div>
          <div class="xlsx-ribbon__group-buttons">
            <button class="xlsx-ribbon__btn" :disabled="!canEditSelection" title="Bold" @click="applyFontStyle({ bold: true })"><b>B</b></button>
            <button class="xlsx-ribbon__btn" :disabled="!canEditSelection" title="Italic" @click="applyFontStyle({ italic: true })"><i>I</i></button>
            <button class="xlsx-ribbon__btn" :disabled="!canEditSelection" title="Underline" @click="applyFontStyle({ underline: 'single' })"><u>U</u></button>
          </div>
        </div>
        <div class="xlsx-ribbon__group">
          <div class="xlsx-ribbon__group-label">Alignment</div>
          <div class="xlsx-ribbon__group-buttons">
            <button class="xlsx-ribbon__btn" :disabled="!canEditSelection" title="Align left" @click="applyAlignment({ horizontal: 'left' })">⫷</button>
            <button class="xlsx-ribbon__btn" :disabled="!canEditSelection" title="Center" @click="applyAlignment({ horizontal: 'center' })">⬌</button>
            <button class="xlsx-ribbon__btn" :disabled="!canEditSelection" title="Align right" @click="applyAlignment({ horizontal: 'right' })">⫸</button>
          </div>
        </div>
        <div class="xlsx-ribbon__group">
          <div class="xlsx-ribbon__group-label">Number</div>
          <div class="xlsx-ribbon__group-buttons">
            <button class="xlsx-ribbon__btn" :disabled="!canEditSelection" title="Currency" @click="applyStyle(NUMBER_FORMATS[2].style)">$</button>
            <button class="xlsx-ribbon__btn" :disabled="!canEditSelection" title="Percent" @click="applyStyle(NUMBER_FORMATS[3].style)">%</button>
          </div>
        </div>
        <div class="xlsx-ribbon__group">
          <div class="xlsx-ribbon__group-label">Cells</div>
          <div class="xlsx-ribbon__group-buttons">
            <button class="xlsx-ribbon__btn" :disabled="!hasSelection || isReadOnly" title="Merge" @click="controller.mergeSelection()">Merge</button>
            <button class="xlsx-ribbon__btn" :disabled="!hasSelection || isReadOnly" title="Unmerge" @click="controller.unmergeSelection()">Unmerge</button>
          </div>
        </div>
      </template>

      <!-- Insert -->
      <template v-if="activeTab === 'Insert'">
        <div class="xlsx-ribbon__group">
          <div class="xlsx-ribbon__group-label">Workbook</div>
          <div class="xlsx-ribbon__group-buttons">
            <button class="xlsx-ribbon__btn" :disabled="!hasWorkbook || isReadOnly" title="Add sheet" @click="controller.addSheet()">+ Sheet</button>
            <button class="xlsx-ribbon__btn" :disabled="controller.tabs.length <= 1 || isReadOnly" title="Delete sheet" @click="controller.removeActiveSheet()">✕ Delete</button>
          </div>
        </div>
      </template>

      <!-- Page Layout -->
      <template v-if="activeTab === 'Page Layout'">
        <div class="xlsx-ribbon__group">
          <div class="xlsx-ribbon__group-label">Export</div>
          <div class="xlsx-ribbon__group-buttons">
            <button class="xlsx-ribbon__btn" :disabled="!controller.canDownload" title="Download source" @click="controller.download()">⬇ Source</button>
            <button class="xlsx-ribbon__btn" :disabled="!controller.canExport" title="Export XLSX" @click="controller.exportXlsx()">💾 XLSX</button>
            <button class="xlsx-ribbon__btn" :disabled="!controller.canExport" title="Export CSV" @click="controller.exportCsv()">📄 CSV</button>
          </div>
        </div>
      </template>

      <!-- Formulas -->
      <template v-if="activeTab === 'Formulas'">
        <div class="xlsx-ribbon__group">
          <div class="xlsx-ribbon__group-label">Calculation</div>
          <div class="xlsx-ribbon__group-buttons">
            <button class="xlsx-ribbon__btn" :disabled="!controller.canExport" title="Recalculate" @click="controller.recalculate()">↻ Recalc</button>
          </div>
        </div>
      </template>

      <!-- Data -->
      <template v-if="activeTab === 'Data'">
        <div class="xlsx-ribbon__group">
          <div class="xlsx-ribbon__group-label">Refresh</div>
          <div class="xlsx-ribbon__group-buttons">
            <button class="xlsx-ribbon__btn" :disabled="!controller.canExport" title="Recalculate" @click="controller.recalculate()">↻ Recalc</button>
          </div>
        </div>
      </template>

      <!-- View -->
      <template v-if="activeTab === 'View'">
        <div class="xlsx-ribbon__group">
          <div class="xlsx-ribbon__group-label">Zoom</div>
          <div class="xlsx-ribbon__group-buttons">
            <button class="xlsx-ribbon__btn" :disabled="!hasWorkbook || !controller.canZoomOut" title="Zoom out" @click="controller.zoomOut()">−</button>
            <select
              class="xlsx-ribbon__select"
              :disabled="!hasWorkbook"
              :value="Math.round(controller.zoomScale)"
              @change="onZoomSelect"
            >
              <option v-for="z in zoomPresets" :key="z" :value="z">{{ z }}%</option>
            </select>
            <button class="xlsx-ribbon__btn" :disabled="!hasWorkbook || !controller.canZoomIn" title="Zoom in" @click="controller.zoomIn()">+</button>
            <button class="xlsx-ribbon__btn" :disabled="!hasWorkbook" title="Reset zoom" @click="controller.resetZoom()">Reset</button>
          </div>
        </div>
        <div class="xlsx-ribbon__group">
          <div class="xlsx-ribbon__group-label">Display</div>
          <div class="xlsx-ribbon__group-buttons">
            <label class="xlsx-ribbon__toggle">
              <input type="checkbox" :checked="readOnly" @change="emit('update:readOnly', ($event.target as HTMLInputElement).checked)" />
              Read only
            </label>
          </div>
        </div>
      </template>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, type CSSProperties } from "vue";
import type { XlsxViewerController, XlsxCellStyleInput } from "@extend-ai/xlsx-core";

const RIBBON_TABS = ["Home", "Insert", "Page Layout", "Formulas", "Data", "View"] as const;

const NUMBER_FORMATS = [
  { label: "General", style: { numberFormat: { formatType: "general" as const } } },
  { label: "Number", style: { numberFormat: { formatType: "custom" as const, formatString: "#,##0.00" } } },
  { label: "Currency", style: { numberFormat: { formatType: "custom" as const, formatString: "$#,##0.00" } } },
  { label: "Percent", style: { numberFormat: { formatType: "custom" as const, formatString: "0.00%" } } },
];

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
  readOnly?: boolean;
}>();

const emit = defineEmits<{
  "update:readOnly": [value: boolean];
}>();

const activeTab = ref<(typeof RIBBON_TABS)[number]>("Home");
const zoomPresets = [50, 75, 100, 125, 150, 200];

const hasWorkbook = computed(() => props.controller.tabs.length > 0);
const hasSelection = computed(() => !!props.controller.selection);
const isReadOnly = computed(() => props.readOnly ?? props.controller.readOnly);
const canEditSelection = computed(() => !!props.controller.activeCell && !isReadOnly.value);

const ribbonStyle = computed<CSSProperties>(() => ({
  backgroundColor: props.isDark ? "#1a1a1a" : "#f8f9fa",
  borderBottom: `1px solid ${props.isDark ? "#3f3f46" : "#e4e4e7"}`,
  flexShrink: "0",
  userSelect: "none",
}));

function onZoomSelect(event: Event) {
  const target = event.target as HTMLSelectElement;
  const value = Number(target.value);
  if (Number.isFinite(value)) {
    props.controller.setZoomScale(value);
  }
}

function applyStyle(style: XlsxCellStyleInput) {
  if (!canEditSelection.value) return;
  if (props.controller.selection) {
    props.controller.setRangeStyle(props.controller.selection, style);
  } else {
    props.controller.setSelectedCellStyle(style);
  }
}

function applyFontStyle(fontStyle: NonNullable<XlsxCellStyleInput["font"]>) {
  applyStyle({ font: fontStyle });
}

function applyAlignment(alignment: NonNullable<XlsxCellStyleInput["alignment"]>) {
  applyStyle({ alignment });
}
</script>

<style scoped>
.xlsx-ribbon {
  font-size: 12px;
}

.xlsx-ribbon__tabs {
  display: flex;
  gap: 0;
  border-bottom: 1px solid var(--border, #e4e4e7);
  padding: 0 8px;
  background: var(--background, #fff);
}

.xlsx-ribbon__tab {
  background: transparent;
  border: 1px solid transparent;
  border-bottom: none;
  border-radius: 4px 4px 0 0;
  color: var(--muted-foreground, #71717a);
  cursor: pointer;
  font-family: inherit;
  font-size: 11px;
  font-weight: 500;
  height: 28px;
  padding: 0 12px;
  white-space: nowrap;
}

.xlsx-ribbon__tab:hover {
  background: rgba(128, 128, 128, 0.08);
  color: inherit;
}

.xlsx-ribbon__tab--active {
  background: var(--muted, #f4f4f5);
  border-color: var(--border, #e4e4e7);
  color: inherit;
}

.xlsx-ribbon__content {
  display: flex;
  gap: 4px;
  padding: 6px 8px;
  min-height: 64px;
  overflow-x: auto;
}

.xlsx-ribbon__group {
  border: 1px solid rgba(128, 128, 128, 0.15);
  border-radius: 6px;
  display: flex;
  flex-direction: column;
  gap: 4px;
  min-width: fit-content;
  padding: 4px 6px;
}

.xlsx-ribbon__group-label {
  color: var(--muted-foreground, #71717a);
  font-size: 10px;
  font-weight: 500;
  text-align: center;
  white-space: nowrap;
}

.xlsx-ribbon__group-buttons {
  align-items: center;
  display: flex;
  gap: 2px;
}

.xlsx-ribbon__btn {
  align-items: center;
  background: transparent;
  border: 1px solid transparent;
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  display: flex;
  font-family: inherit;
  font-size: 12px;
  height: 26px;
  justify-content: center;
  min-width: 26px;
  padding: 0 6px;
  white-space: nowrap;
}

.xlsx-ribbon__btn:hover:not(:disabled) {
  background: rgba(128, 128, 128, 0.12);
  border-color: rgba(128, 128, 128, 0.2);
}

.xlsx-ribbon__btn:disabled {
  opacity: 0.35;
  cursor: default;
}

.xlsx-ribbon__select {
  background: transparent;
  border: 1px solid rgba(128, 128, 128, 0.2);
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  font-family: inherit;
  font-size: 11px;
  height: 26px;
  outline: none;
  padding: 0 4px;
}

.xlsx-ribbon__toggle {
  align-items: center;
  cursor: pointer;
  display: flex;
  font-size: 11px;
  gap: 4px;
  padding: 2px 6px;
  border: 1px solid rgba(128, 128, 128, 0.15);
  border-radius: 4px;
  white-space: nowrap;
}
</style>
