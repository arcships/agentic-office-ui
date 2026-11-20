<template>
  <Teleport to="body">
    <OfficeThemeProvider :theme="isDark ? 'dark' : 'light'">
    <div class="office-overlay__backdrop xlsx-authoring-backdrop" data-testid="xlsx-authoring-backdrop" @pointerdown="emit('close')" />
    <div
      ref="dialogRef"
      class="office-ui office-dialog__content office-dialog__content--lg xlsx-authoring"
      :class="{ 'xlsx-authoring--dark': isDark }"
      data-testid="xlsx-authoring-dialog"
      :data-kind="kind"
      role="dialog"
      aria-modal="true"
      :aria-label="dialogTitle"
      tabindex="-1"
      @keydown.escape.stop.prevent="emit('close')"
      @pointerdown.stop
    >
      <header class="office-dialog__header xlsx-authoring__header">
        <h2 class="office-dialog__title xlsx-authoring__title">{{ dialogTitle }}</h2>
        <button
          type="button"
          class="office-button office-button--ghost office-button--size-icon-sm office-dialog__close xlsx-authoring__close"
          data-testid="xlsx-authoring-close"
          :title="m.xlsx.authoring.close"
          :aria-label="m.xlsx.authoring.close"
          @click="emit('close')"
        ><span class="office-button__content" aria-hidden="true">×</span></button>
      </header>
      <div class="office-dialog__body xlsx-authoring__body">
      <p class="xlsx-authoring__range" data-testid="xlsx-authoring-range">{{ rangeLabel }}</p>
      <p v-if="statusMessage" class="xlsx-authoring__status" data-testid="xlsx-authoring-status" role="status">
        {{ statusMessage }}
      </p>

      <!-- Chart -->
      <section v-if="kind === 'chart'" class="xlsx-authoring__section">
        <div class="xlsx-authoring__tabs">
          <button
            type="button"
            class="xlsx-authoring__tab"
            :class="{ 'xlsx-authoring__tab--active': chartMode === 'insert' }"
            data-testid="xlsx-chart-mode-insert"
            @click="chartMode = 'insert'"
          >{{ m.xlsx.authoring.chart.insertMode }}</button>
          <button
            type="button"
            class="xlsx-authoring__tab"
            :class="{ 'xlsx-authoring__tab--active': chartMode === 'edit' }"
            data-testid="xlsx-chart-mode-edit"
            :disabled="!selectedChart"
            @click="chartMode = 'edit'"
          >{{ m.xlsx.authoring.chart.editMode }}</button>
        </div>
        <p v-if="chartMode === 'edit' && selectedChart" class="xlsx-authoring__hint" data-testid="xlsx-chart-selected">
          {{ formatOfficeMessage(m.xlsx.authoring.chart.selectedChart, { title: selectedChart.title || selectedChart.name || selectedChart.chartType }) }}
        </p>
        <p v-else-if="chartMode === 'edit'" class="xlsx-authoring__hint">{{ m.xlsx.authoring.chart.noneSelected }}</p>
        <label v-if="chartMode === 'insert'" class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.chart.family }}</span>
          <select v-model="chartFamily" data-testid="xlsx-chart-family">
            <option value="bar">{{ m.xlsx.authoring.chart.familyBar }}</option>
            <option value="line">{{ m.xlsx.authoring.chart.familyLine }}</option>
            <option value="pie">{{ m.xlsx.authoring.chart.familyPie }}</option>
          </select>
        </label>
        <label class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.chart.chartTitle }}</span>
          <input
            v-model="chartTitle"
            type="text"
            data-testid="xlsx-chart-title"
            :placeholder="m.xlsx.authoring.chart.chartTitlePlaceholder"
          />
        </label>
        <label v-if="chartMode === 'edit'" class="xlsx-authoring__check">
          <input v-model="chartShowDataLabels" type="checkbox" data-testid="xlsx-chart-data-labels" />
          <span>{{ m.xlsx.authoring.chart.showDataLabels }}</span>
        </label>
        <template v-if="chartMode === 'insert'">
          <label class="xlsx-authoring__check">
            <input v-model="chartHeaderRow" type="checkbox" data-testid="xlsx-chart-header-row" />
            <span>{{ m.xlsx.authoring.chart.headerRow }}</span>
          </label>
          <label class="xlsx-authoring__check">
            <input v-model="chartHeaderColumn" type="checkbox" data-testid="xlsx-chart-header-column" />
            <span>{{ m.xlsx.authoring.chart.headerColumn }}</span>
          </label>
        </template>
        <div class="xlsx-authoring__actions">
          <button
            v-if="chartMode === 'insert'"
            type="button"
            class="office-button office-button--default office-button--size-sm xlsx-authoring__button xlsx-authoring__button--primary"
            data-testid="xlsx-chart-insert"
            @click="insertChart"
          >{{ m.xlsx.authoring.chart.insert }}</button>
          <button
            v-else
            type="button"
            class="office-button office-button--default office-button--size-sm xlsx-authoring__button xlsx-authoring__button--primary"
            data-testid="xlsx-chart-update"
            :disabled="!selectedChart"
            @click="updateChart"
          >{{ m.xlsx.authoring.chart.update }}</button>
          <button
            v-if="selectedChart"
            type="button"
            class="office-button office-button--outline office-button--size-sm xlsx-authoring__button"
            data-testid="xlsx-chart-delete"
            @click="deleteChart"
          >{{ m.xlsx.authoring.chart.delete }}</button>
        </div>
      </section>

      <!-- Comment -->
      <section v-else-if="kind === 'comment'" class="xlsx-authoring__section">
        <label class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.comment.text }}</span>
          <textarea
            v-model="commentText"
            rows="4"
            data-testid="xlsx-comment-text"
            :placeholder="m.xlsx.authoring.comment.textPlaceholder"
          />
        </label>
        <label class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.comment.author }}</span>
          <input
            v-model="commentAuthor"
            type="text"
            data-testid="xlsx-comment-author"
            :placeholder="m.xlsx.authoring.comment.authorPlaceholder"
          />
        </label>
        <div class="xlsx-authoring__actions">
          <button
            type="button"
            class="office-button office-button--default office-button--size-sm xlsx-authoring__button xlsx-authoring__button--primary"
            data-testid="xlsx-comment-save"
            :disabled="!commentText.trim()"
            @click="saveComment"
          >{{ m.xlsx.authoring.comment.save }}</button>
          <button
            v-if="existingComment"
            type="button"
            class="office-button office-button--outline office-button--size-sm xlsx-authoring__button"
            data-testid="xlsx-comment-delete"
            @click="deleteComment"
          >{{ m.xlsx.authoring.comment.delete }}</button>
        </div>
      </section>

      <!-- Hyperlink -->
      <section v-else-if="kind === 'hyperlink'" class="xlsx-authoring__section">
        <label class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.hyperlink.target }}</span>
          <input
            v-model="hyperlinkTarget"
            type="text"
            data-testid="xlsx-hyperlink-target"
            :placeholder="m.xlsx.authoring.hyperlink.targetPlaceholder"
          />
        </label>
        <label class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.hyperlink.tooltip }}</span>
          <input
            v-model="hyperlinkTooltip"
            type="text"
            data-testid="xlsx-hyperlink-tooltip"
            :placeholder="m.xlsx.authoring.hyperlink.tooltipPlaceholder"
          />
        </label>
        <div class="xlsx-authoring__actions">
          <button
            type="button"
            class="office-button office-button--default office-button--size-sm xlsx-authoring__button xlsx-authoring__button--primary"
            data-testid="xlsx-hyperlink-save"
            :disabled="!hyperlinkTarget.trim()"
            @click="saveHyperlink"
          >{{ m.xlsx.authoring.hyperlink.save }}</button>
          <button
            v-if="existingHyperlink"
            type="button"
            class="office-button office-button--outline office-button--size-sm xlsx-authoring__button"
            data-testid="xlsx-hyperlink-remove"
            @click="removeHyperlink"
          >{{ m.xlsx.authoring.hyperlink.remove }}</button>
        </div>
      </section>

      <!-- Table -->
      <section v-else-if="kind === 'table'" class="xlsx-authoring__section">
        <p v-if="tableAtCell" class="xlsx-authoring__hint" data-testid="xlsx-table-existing">
          {{ formatOfficeMessage(m.xlsx.authoring.table.existing, { name: tableAtCell.name, range: tableAtCell.reference }) }}
        </p>
        <label class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.table.name }}</span>
          <input
            v-model="tableName"
            type="text"
            data-testid="xlsx-table-name"
            :placeholder="m.xlsx.authoring.table.namePlaceholder"
          />
        </label>
        <label class="xlsx-authoring__check">
          <input v-model="tableHasHeaders" type="checkbox" data-testid="xlsx-table-has-headers" />
          <span>{{ m.xlsx.authoring.table.hasHeaders }}</span>
        </label>
        <div class="xlsx-authoring__actions">
          <button
            type="button"
            class="office-button office-button--default office-button--size-sm xlsx-authoring__button xlsx-authoring__button--primary"
            data-testid="xlsx-table-create"
            @click="createTable"
          >{{ m.xlsx.authoring.table.create }}</button>
          <button
            v-if="tableAtCell"
            type="button"
            class="office-button office-button--outline office-button--size-sm xlsx-authoring__button"
            data-testid="xlsx-table-remove"
            @click="removeTable"
          >{{ m.xlsx.authoring.table.remove }}</button>
        </div>
      </section>

      <!-- Conditional formatting -->
      <section v-else-if="kind === 'conditional'" class="xlsx-authoring__section">
        <ul class="xlsx-authoring__rules" data-testid="xlsx-conditional-rules">
          <li v-for="entry in conditionalRules" :key="entry.index" class="xlsx-authoring__rule" data-testid="xlsx-conditional-rule">
            <span class="xlsx-authoring__rule-label">
              {{ formatOfficeMessage(m.xlsx.authoring.conditional.ruleSummary, { kind: ruleKindLabel(entry.rule.kind), ranges: ruleRangeLabel(entry.rule) }) }}
            </span>
            <span class="xlsx-authoring__rule-actions">
              <button type="button" class="xlsx-authoring__link" data-testid="xlsx-conditional-edit" @click="editConditionalRule(entry.index)">
                {{ m.xlsx.authoring.conditional.edit }}
              </button>
              <button type="button" class="xlsx-authoring__link" data-testid="xlsx-conditional-delete" @click="deleteConditionalRule(entry.index)">
                {{ m.xlsx.authoring.conditional.delete }}
              </button>
            </span>
          </li>
          <li v-if="conditionalRules.length === 0" class="xlsx-authoring__hint" data-testid="xlsx-conditional-empty">
            {{ m.xlsx.authoring.conditional.empty }}
          </li>
        </ul>
        <label class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.conditional.kind }}</span>
          <select v-model="conditionalKind" data-testid="xlsx-conditional-kind">
            <option value="colorScale">{{ m.xlsx.authoring.conditional.kindColorScale }}</option>
            <option value="dataBar">{{ m.xlsx.authoring.conditional.kindDataBar }}</option>
            <option value="iconSet">{{ m.xlsx.authoring.conditional.kindIconSet }}</option>
          </select>
        </label>
        <label v-if="conditionalKind === 'iconSet'" class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.conditional.iconSet }}</span>
          <select v-model="conditionalIconSet" data-testid="xlsx-conditional-iconset">
            <option value="3Arrows">3Arrows</option>
            <option value="3TrafficLights1">3TrafficLights1</option>
            <option value="4Arrows">4Arrows</option>
            <option value="5Arrows">5Arrows</option>
          </select>
        </label>
        <fieldset class="xlsx-authoring__thresholds">
          <legend>{{ m.xlsx.authoring.conditional.thresholds }}</legend>
          <div v-for="(threshold, index) in thresholdRows" :key="index" class="xlsx-authoring__threshold">
            <select v-model="threshold.type" data-testid="xlsx-conditional-threshold-type">
              <option value="min">{{ m.xlsx.authoring.conditional.typeMin }}</option>
              <option value="num">{{ m.xlsx.authoring.conditional.typeNumber }}</option>
              <option value="percent">{{ m.xlsx.authoring.conditional.typePercent }}</option>
              <option value="percentile">{{ m.xlsx.authoring.conditional.typePercentile }}</option>
              <option value="max">{{ m.xlsx.authoring.conditional.typeMax }}</option>
            </select>
            <input
              v-model="threshold.value"
              type="text"
              inputmode="decimal"
              data-testid="xlsx-conditional-threshold-value"
              :disabled="threshold.type === 'min' || threshold.type === 'max'"
              :placeholder="m.xlsx.authoring.conditional.value"
            />
          </div>
        </fieldset>
        <div class="xlsx-authoring__colors">
          <label v-for="(_, index) in conditionalColors" :key="index" class="xlsx-authoring__color">
            <span>{{ m.xlsx.authoring.conditional.color }}</span>
            <input v-model="conditionalColors[index]" type="color" data-testid="xlsx-conditional-color" />
          </label>
        </div>
        <label class="xlsx-authoring__check">
          <input v-model="conditionalShowValue" type="checkbox" data-testid="xlsx-conditional-show-value" />
          <span>{{ m.xlsx.authoring.conditional.showValue }}</span>
        </label>
        <label v-if="conditionalKind === 'iconSet'" class="xlsx-authoring__check">
          <input v-model="conditionalReverse" type="checkbox" data-testid="xlsx-conditional-reverse" />
          <span>{{ m.xlsx.authoring.conditional.reverse }}</span>
        </label>
        <div class="xlsx-authoring__actions">
          <button
            type="button"
            class="office-button office-button--default office-button--size-sm xlsx-authoring__button xlsx-authoring__button--primary"
            data-testid="xlsx-conditional-add"
            @click="submitConditionalRule"
          >{{ editingConditionalIndex === null ? m.xlsx.authoring.conditional.add : m.xlsx.authoring.conditional.update }}</button>
          <button
            v-if="editingConditionalIndex !== null"
            type="button"
            class="office-button office-button--outline office-button--size-sm xlsx-authoring__button"
            data-testid="xlsx-conditional-cancel"
            @click="resetConditionalForm"
          >{{ m.xlsx.authoring.close }}</button>
        </div>
      </section>

      <!-- Data validation -->
      <section v-else-if="kind === 'validation'" class="xlsx-authoring__section">
        <p v-if="existingValidation" class="xlsx-authoring__hint" data-testid="xlsx-validation-existing">
          {{ formatOfficeMessage(m.xlsx.authoring.validation.existing, { type: existingValidation.validationType, ranges: rangeLabel }) }}
        </p>
        <label class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.validation.type }}</span>
          <select v-model="validationType" data-testid="xlsx-validation-type">
            <option value="list">{{ m.xlsx.authoring.validation.typeList }}</option>
            <option value="whole">{{ m.xlsx.authoring.validation.typeWhole }}</option>
            <option value="decimal">{{ m.xlsx.authoring.validation.typeDecimal }}</option>
            <option value="date">{{ m.xlsx.authoring.validation.typeDate }}</option>
            <option value="textLength">{{ m.xlsx.authoring.validation.typeTextLength }}</option>
          </select>
        </label>
        <label v-if="validationType === 'list'" class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.validation.source }}</span>
          <input
            v-model="validationSource"
            type="text"
            data-testid="xlsx-validation-source"
            :placeholder="m.xlsx.authoring.validation.sourcePlaceholder"
          />
        </label>
        <template v-else>
          <label class="xlsx-authoring__field">
            <span>{{ m.xlsx.authoring.validation.value1 }}</span>
            <input v-model="validationValue1" type="text" data-testid="xlsx-validation-value1" />
          </label>
          <label class="xlsx-authoring__field">
            <span>{{ m.xlsx.authoring.validation.value2 }}</span>
            <input v-model="validationValue2" type="text" data-testid="xlsx-validation-value2" />
          </label>
        </template>
        <label class="xlsx-authoring__check">
          <input v-model="validationAllowBlank" type="checkbox" data-testid="xlsx-validation-allow-blank" />
          <span>{{ m.xlsx.authoring.validation.allowBlank }}</span>
        </label>
        <label v-if="validationType === 'list'" class="xlsx-authoring__check">
          <input v-model="validationShowDropdown" type="checkbox" data-testid="xlsx-validation-show-dropdown" />
          <span>{{ m.xlsx.authoring.validation.showDropdown }}</span>
        </label>
        <label class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.validation.inputMessage }}</span>
          <input v-model="validationInputMessage" type="text" data-testid="xlsx-validation-input-message" />
        </label>
        <label class="xlsx-authoring__field">
          <span>{{ m.xlsx.authoring.validation.errorMessage }}</span>
          <input v-model="validationErrorMessage" type="text" data-testid="xlsx-validation-error-message" />
        </label>
        <div class="xlsx-authoring__actions">
          <button
            type="button"
            class="office-button office-button--default office-button--size-sm xlsx-authoring__button xlsx-authoring__button--primary"
            data-testid="xlsx-validation-apply"
            @click="applyValidation"
          >{{ m.xlsx.authoring.validation.apply }}</button>
          <button
            v-if="existingValidation"
            type="button"
            class="office-button office-button--outline office-button--size-sm xlsx-authoring__button"
            data-testid="xlsx-validation-remove"
            @click="removeValidation"
          >{{ m.xlsx.authoring.validation.remove }}</button>
        </div>
      </section>

      <!-- Paste special -->
      <XlsxPasteSpecialSection
        v-else-if="kind === 'pasteSpecial'"
        :controller="controller"
        @applied="emit('close')"
      />

      <!-- Number format -->
      <XlsxNumberFormatSection
        v-else-if="kind === 'numberFormat'"
        :controller="controller"
        @applied="emit('close')"
      />

      <!-- Name manager -->
      <XlsxNameManagerSection
        v-else-if="kind === 'names'"
        :controller="controller"
      />
      </div>
    </div>
    </OfficeThemeProvider>
  </Teleport>
</template>

<script setup lang="ts">
import { computed, nextTick, onMounted, ref, watch } from "vue";
import { formatOfficeMessage } from "@arcships/office-interaction";
import type {
  XlsxCellAddress,
  XlsxCellRange,
  XlsxChartFamily,
  XlsxConditionalFormatRule,
  XlsxViewerController,
} from "@arcships/xlsx-core";
import { OfficeThemeProvider } from "@arcships/vue-ui";
import { useXlsxMessages } from "../i18n";
import XlsxNameManagerSection from "./XlsxNameManagerSection.vue";
import XlsxNumberFormatSection from "./XlsxNumberFormatSection.vue";
import XlsxPasteSpecialSection from "./XlsxPasteSpecialSection.vue";
import type { XlsxAuthoringDialogKind } from "./authoring-dialog";

const props = defineProps<{
  controller: XlsxViewerController;
  kind: XlsxAuthoringDialogKind;
  isDark?: boolean;
}>();

const emit = defineEmits<{ close: [] }>();

const m = useXlsxMessages();
const dialogRef = ref<HTMLElement | null>(null);
const statusMessage = ref("");
const activeCell = computed<XlsxCellAddress | null>(() => props.controller.activeCell ?? null);
const selectionRange = computed<XlsxCellRange | null>(() => {
  if (props.controller.selection) {
    return props.controller.selection;
  }
  const cell = activeCell.value;
  return cell ? { end: cell, start: cell } : null;
});
const singleCellSelection = computed(() => {
  const range = selectionRange.value;
  return Boolean(range && range.start.row === range.end.row && range.start.col === range.end.col);
});
const rangeLabel = computed(() => {
  const address = props.controller.selection
    ? props.controller.selectedRangeAddress
    : props.controller.activeCellAddress;
  return formatOfficeMessage(m.value.xlsx.authoring.rangeLabel, { range: address ?? "" });
});

const dialogTitle = computed(() => {
  switch (props.kind) {
    case "chart":
      return m.value.xlsx.authoring.chart.title;
    case "comment":
      return m.value.xlsx.authoring.comment.title;
    case "hyperlink":
      return m.value.xlsx.authoring.hyperlink.title;
    case "table":
      return m.value.xlsx.authoring.table.title;
    case "conditional":
      return m.value.xlsx.authoring.conditional.title;
    case "pasteSpecial":
      return m.value.xlsx.authoring.pasteSpecial.title;
    case "numberFormat":
      return m.value.xlsx.authoring.numberFormat.title;
    case "names":
      return m.value.xlsx.authoring.names.title;
    default:
      return m.value.xlsx.authoring.validation.title;
  }
});

onMounted(() => {
  void nextTick(() => dialogRef.value?.focus?.());
});

// ── Chart ────────────────────────────────────────────────────────────
const selectedChart = computed(() => {
  props.controller.revision;
  return props.controller.selectedChartId ? props.controller.getChartById?.(props.controller.selectedChartId) ?? null : null;
});
const chartMode = ref<"edit" | "insert">(selectedChart.value ? "edit" : "insert");
const chartFamily = ref<XlsxChartFamily>("bar");
const chartTitle = ref(selectedChart.value?.title ?? "");
const chartHeaderRow = ref(true);
const chartHeaderColumn = ref(true);
const chartShowDataLabels = ref(false);

watch(selectedChart, (chart) => {
  if (!chart) {
    return;
  }
  chartMode.value = "edit";
  chartTitle.value = chart.title ?? "";
  chartShowDataLabels.value = chart.dataLabels?.showValue === true;
}, { immediate: true });

function insertChart() {
  const insert = props.controller.insertChart;
  if (typeof insert !== "function") {
    return;
  }
  const chart = insert({
    family: chartFamily.value,
    hasHeaderColumn: chartHeaderColumn.value,
    hasHeaderRow: chartHeaderRow.value,
    // A single-cell selection means "chart this sheet": let the controller
    // fall back to the used range instead of failing on one cell.
    range: singleCellSelection.value ? undefined : selectionRange.value ?? undefined,
    title: chartTitle.value.trim() || undefined
  }) as { id?: string } | null;
  if (!chart) {
    statusMessage.value = m.value.xlsx.authoring.status.actionFailed;
    return;
  }
  if (chart.id) {
    props.controller.selectChart?.(chart.id);
  }
  statusMessage.value = m.value.xlsx.authoring.status.chartInserted;
}

function updateChart() {
  const chart = selectedChart.value;
  const update = props.controller.updateChart;
  if (!chart || typeof update !== "function") {
    return;
  }
  update(chart.id, {
    dataLabels: { ...(chart.dataLabels ?? {}), showValue: chartShowDataLabels.value },
    title: chartTitle.value.trim() || undefined
  });
  statusMessage.value = m.value.xlsx.authoring.status.chartUpdated;
}

function deleteChart() {
  const chart = selectedChart.value;
  const remove = props.controller.deleteChart;
  if (!chart || typeof remove !== "function") {
    return;
  }
  if (remove(chart.id)) {
    statusMessage.value = m.value.xlsx.authoring.status.chartDeleted;
  }
}

// ── Comment ──────────────────────────────────────────────────────────
const existingComment = computed(() => {
  props.controller.revision;
  return (props.controller.getCellComment?.() ?? null) as { author?: string; text: string } | null;
});
const commentText = ref("");
const commentAuthor = ref("");

watch(existingComment, (comment) => {
  commentText.value = comment?.text ?? "";
  commentAuthor.value = comment?.author ?? "";
}, { immediate: true });

function saveComment() {
  const cell = activeCell.value;
  const save = props.controller.setCellComment;
  if (!cell || typeof save !== "function") {
    return;
  }
  const changed = save(cell, { author: commentAuthor.value.trim() || undefined, text: commentText.value });
  statusMessage.value = changed
    ? m.value.xlsx.authoring.status.commentSaved
    : m.value.xlsx.authoring.status.actionFailed;
}

function deleteComment() {
  const cell = activeCell.value;
  const save = props.controller.setCellComment;
  if (!cell || typeof save !== "function") {
    return;
  }
  if (save(cell, null)) {
    statusMessage.value = m.value.xlsx.authoring.status.commentDeleted;
  }
}

// ── Hyperlink ────────────────────────────────────────────────────────
const existingHyperlink = computed(() => {
  props.controller.revision;
  return (props.controller.getCellHyperlink?.() ?? null) as { target: string; tooltip?: string } | null;
});
const hyperlinkTarget = ref("");
const hyperlinkTooltip = ref("");

watch(existingHyperlink, (hyperlink) => {
  hyperlinkTarget.value = hyperlink?.target ?? "";
  hyperlinkTooltip.value = hyperlink?.tooltip ?? "";
}, { immediate: true });

function saveHyperlink() {
  const cell = activeCell.value;
  const save = props.controller.setCellHyperlink;
  if (!cell || typeof save !== "function") {
    return;
  }
  const changed = save(cell, {
    target: hyperlinkTarget.value.trim(),
    tooltip: hyperlinkTooltip.value.trim() || undefined
  });
  statusMessage.value = changed
    ? m.value.xlsx.authoring.status.hyperlinkSaved
    : m.value.xlsx.authoring.status.actionFailed;
}

function removeHyperlink() {
  const cell = activeCell.value;
  const save = props.controller.setCellHyperlink;
  if (!cell || typeof save !== "function") {
    return;
  }
  if (save(cell, null)) {
    statusMessage.value = m.value.xlsx.authoring.status.hyperlinkRemoved;
  }
}

// ── Table ────────────────────────────────────────────────────────────
const tableAtCell = computed(() => {
  props.controller.revision;
  return (props.controller.getTableAtCell?.(activeCell.value) ?? null) as { name: string; reference: string } | null;
});
const tableName = ref("");
const tableHasHeaders = ref(true);

watch(tableAtCell, (table) => {
  tableName.value = table?.name ?? "";
}, { immediate: true });

function createTable() {
  const create = props.controller.createTable;
  if (typeof create !== "function") {
    return;
  }
  const table = create({
    hasHeaders: tableHasHeaders.value,
    name: tableName.value.trim() || undefined,
    range: selectionRange.value ?? undefined
  }) as { name?: string } | null;
  statusMessage.value = table
    ? m.value.xlsx.authoring.status.tableCreated
    : m.value.xlsx.authoring.status.actionFailed;
}

function removeTable() {
  const table = tableAtCell.value;
  const remove = props.controller.removeTable;
  if (!table || typeof remove !== "function") {
    return;
  }
  if (remove(table.name)) {
    statusMessage.value = m.value.xlsx.authoring.status.tableRemoved;
  }
}

// ── Conditional formatting ───────────────────────────────────────────
const conditionalRules = computed(() => {
  props.controller.revision;
  const rules = (props.controller.getConditionalFormatRules?.() ?? []) as XlsxConditionalFormatRule[];
  return rules.map((rule, index) => ({ index, rule }));
});
const conditionalKind = ref<"colorScale" | "dataBar" | "iconSet">("colorScale");
const conditionalIconSet = ref("3TrafficLights1");
const conditionalColors = ref<string[]>(["#f8696b", "#ffeb84", "#63be7b"]);
const conditionalShowValue = ref(true);
const conditionalReverse = ref(false);
const editingConditionalIndex = ref<number | null>(null);
/** Set while a stored rule loads into the form so defaults do not overwrite its cfvos. */
let loadingConditionalRule = false;
const thresholdRows = ref<Array<{ type: string; value: string }>>([
  { type: "min", value: "" },
  { type: "percentile", value: "50" },
  { type: "max", value: "" }
]);

const iconSetThresholdCount = computed(() => {
  const match = /^(\d)/.exec(conditionalIconSet.value);
  const count = match ? Number(match[1]) : 3;
  return count >= 3 && count <= 5 ? count : 3;
});

watch([conditionalKind, iconSetThresholdCount], () => {
  if (loadingConditionalRule) {
    return;
  }
  const rows: Array<{ type: string; value: string }> = [];
  if (conditionalKind.value === "colorScale") {
    rows.push({ type: "min", value: "" }, { type: "percentile", value: "50" }, { type: "max", value: "" });
    conditionalColors.value = conditionalColors.value.length === 3
      ? conditionalColors.value
      : ["#f8696b", "#ffeb84", "#63be7b"];
  } else if (conditionalKind.value === "dataBar") {
    rows.push({ type: "min", value: "" }, { type: "max", value: "" });
    conditionalColors.value = [conditionalColors.value[0] ?? "#5b9bd5"];
  } else {
    const count = iconSetThresholdCount.value;
    for (let index = 0; index < count; index += 1) {
      rows.push(index === 0
        ? { type: "percent", value: "0" }
        : { type: "percent", value: String(Math.round((index * 100) / count)) });
    }
    conditionalColors.value = [conditionalColors.value[0] ?? "#5b9bd5"];
  }
  thresholdRows.value = rows;
});

function ruleKindLabel(kind: string) {
  if (kind === "colorScale") return m.value.xlsx.authoring.conditional.kindColorScale;
  if (kind === "dataBar") return m.value.xlsx.authoring.conditional.kindDataBar;
  return m.value.xlsx.authoring.conditional.kindIconSet;
}

function ruleRangeLabel(rule: XlsxConditionalFormatRule) {
  return rule.ranges.map((range) => rangeAddress(range)).join(", ");
}

function rangeAddress(range: XlsxCellRange) {
  const start = cellAddress(range.start);
  const end = cellAddress(range.end);
  return start === end ? start : `${start}:${end}`;
}

function cellAddress(cell: XlsxCellAddress) {
  let label = "";
  let value = Math.max(0, cell.col);
  do {
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26) - 1;
  } while (value >= 0);
  return `${label}${Math.max(0, cell.row) + 1}`;
}

function normalizeColor(value: string) {
  const hex = value.trim().replace(/^#/, "");
  return hex.length === 8 ? hex.slice(2).toUpperCase() : hex.toUpperCase();
}

function buildConditionalDraft() {
  const thresholds = thresholdRows.value.map((row) => ({
    type: row.type,
    value: row.value.trim() === "" ? undefined : row.value.trim()
  }));
  return {
    kind: conditionalKind.value,
    colors: conditionalColors.value.map(normalizeColor),
    iconSet: conditionalIconSet.value,
    reverse: conditionalReverse.value,
    showValue: conditionalShowValue.value,
    thresholds
  };
}

function editConditionalRule(index: number) {
  const rule = props.controller.getConditionalFormatRules?.()[index];
  if (!rule) {
    return;
  }
  editingConditionalIndex.value = index;
  loadingConditionalRule = true;
  conditionalKind.value = rule.kind;
  if (rule.kind === "iconSet") {
    conditionalIconSet.value = rule.icons[0]?.iconSet ?? "3TrafficLights1";
  }
  conditionalShowValue.value = rule.kind === "colorScale" ? true : rule.showValue ?? true;
  conditionalReverse.value = rule.kind === "iconSet" ? rule.reverse === true : false;
  thresholdRows.value = rule.cfvos.map((cfvo) => ({
    type: cfvo.type,
    value: cfvo.value === undefined ? "" : String(cfvo.value)
  }));
  conditionalColors.value = rule.kind === "colorScale"
    ? rule.colors.map((color) => String(color.rgb ?? "#5b9bd5"))
    : [String((rule.kind === "dataBar" ? rule.color?.rgb : undefined) ?? "#5b9bd5")];
  void nextTick(() => { loadingConditionalRule = false; });
}

function resetConditionalForm() {
  editingConditionalIndex.value = null;
  conditionalKind.value = "colorScale";
  conditionalIconSet.value = "3TrafficLights1";
  conditionalShowValue.value = true;
  conditionalReverse.value = false;
  conditionalColors.value = ["#f8696b", "#ffeb84", "#63be7b"];
  thresholdRows.value = [
    { type: "min", value: "" },
    { type: "percentile", value: "50" },
    { type: "max", value: "" }
  ];
}

function submitConditionalRule() {
  const draft = buildConditionalDraft();
  if (editingConditionalIndex.value === null) {
    statusMessage.value = props.controller.addConditionalFormatRule?.(draft)
      ? m.value.xlsx.authoring.status.ruleAdded
      : m.value.xlsx.authoring.status.actionFailed;
    return;
  }
  if (props.controller.updateConditionalFormatRule?.(editingConditionalIndex.value, draft)) {
    statusMessage.value = m.value.xlsx.authoring.status.ruleUpdated;
    resetConditionalForm();
  } else {
    statusMessage.value = m.value.xlsx.authoring.status.actionFailed;
  }
}

function deleteConditionalRule(index: number) {
  if (props.controller.deleteConditionalFormatRule?.(index)) {
    statusMessage.value = m.value.xlsx.authoring.status.ruleRemoved;
    if (editingConditionalIndex.value === index) {
      resetConditionalForm();
    }
  }
}

// ── Data validation ──────────────────────────────────────────────────
const existingValidation = computed(() => {
  props.controller.revision;
  return (props.controller.getCellDataValidation?.(activeCell.value) ?? null) as
    | { listSource?: string; validationType: string; value1?: string }
    | null;
});
const validationType = ref<"list" | "whole" | "decimal" | "date" | "textLength">("list");
const validationSource = ref("");
const validationValue1 = ref("");
const validationValue2 = ref("");
const validationAllowBlank = ref(true);
const validationShowDropdown = ref(true);
const validationInputMessage = ref("");
const validationErrorMessage = ref("");

function applyValidation() {
  const apply = props.controller.setSelectionDataValidation;
  if (typeof apply !== "function") {
    return;
  }
  const changed = apply({
    allowBlank: validationAllowBlank.value,
    errorMessage: validationErrorMessage.value.trim() || undefined,
    inputMessage: validationInputMessage.value.trim() || undefined,
    listSource: validationType.value === "list" ? validationSource.value : undefined,
    ranges: selectionRange.value ? [selectionRange.value] : undefined,
    showDropdown: validationShowDropdown.value,
    validationType: validationType.value,
    value1: validationValue1.value.trim() || undefined,
    value2: validationValue2.value.trim() || undefined
  });
  statusMessage.value = changed
    ? m.value.xlsx.authoring.status.validationSet
    : m.value.xlsx.authoring.status.actionFailed;
}

function removeValidation() {
  const apply = props.controller.setSelectionDataValidation;
  if (typeof apply !== "function") {
    return;
  }
  if (apply(null)) {
    statusMessage.value = m.value.xlsx.authoring.status.validationRemoved;
  }
}
</script>

/* Shared section controls. The paste-special, number-format and Name Manager
   sections are child components, so they need these rules unscoped. */
.xlsx-authoring__field { display: flex; flex-direction: column; font-size: var(--office-text-xs, 0.75rem); gap: 4px; }
.xlsx-authoring__field select,
.xlsx-authoring__field input,
.xlsx-authoring__field textarea {
  background: var(--office-background, #fff);
  border: var(--office-border-width, 1px) solid var(--office-input, #d4d4d8);
  border-radius: var(--office-radius-md, 0.5rem);
  color: var(--office-foreground, #18181b);
  font: inherit;
  font-size: var(--office-text-sm, 0.875rem);
  padding: 6px 8px;
}
.xlsx-authoring__field select:focus-visible,
.xlsx-authoring__field input:focus-visible,
.xlsx-authoring__field textarea:focus-visible {
  border-color: var(--office-ring, #a1a1aa);
  outline: 2px solid color-mix(in oklab, var(--office-ring, #a1a1aa) 40%, transparent);
  outline-offset: var(--office-ring-offset-width, 1px);
}
.xlsx-authoring__check { align-items: center; display: flex; font-size: var(--office-text-xs, 0.75rem); gap: 6px; }
.xlsx-authoring__check input { accent-color: var(--office-primary, #2563eb); }
.xlsx-authoring__actions { display: flex; flex-wrap: wrap; gap: var(--office-ui-space-2, 0.5rem); }
.xlsx-authoring__hint { color: var(--office-muted-foreground, #71717a); font-size: var(--office-text-xs, 0.75rem); margin: 0; }
.xlsx-authoring__rules { display: flex; flex-direction: column; gap: 6px; list-style: none; margin: 0; padding: 0; }
.xlsx-authoring__rule { align-items: center; display: flex; font-size: var(--office-text-xs, 0.75rem); gap: var(--office-ui-space-2, 0.5rem); justify-content: space-between; }
.xlsx-authoring__link {
  background: transparent;
  border: 0;
  color: var(--office-info-foreground, #2563eb);
  cursor: pointer;
  font: inherit;
  font-size: var(--office-text-xs, 0.75rem);
  padding: 0;
}


<style scoped>
/*
 * The panel is the shared `.office-dialog__content` surface (centred modal,
 * `lg` width, shared header/body/footer parts). Only the spreadsheet-specific
 * pieces — range chip, section stack, conditional-format thresholds — need
 * local rules, and every colour resolves from an Office token.
 */
.xlsx-authoring-backdrop {
  backdrop-filter: none;
}

.xlsx-authoring {
  gap: var(--office-ui-space-2, 0.5rem);
  max-block-size: min(82vh, 720px);
  max-inline-size: min(92vw, 480px);
}

.xlsx-authoring__header {
  flex-direction: row;
  align-items: center;
  justify-content: space-between;
  padding-block-end: var(--office-ui-space-2, 0.5rem);
}

.xlsx-authoring__title {
  font-size: var(--office-text-lg, 1.125rem);
}

.xlsx-authoring__close {
  border-radius: var(--office-radius-md, 0.5rem);
  font-size: 18px;
  line-height: 1;
  position: static;
}

.xlsx-authoring__body {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-2, 0.5rem);
  padding-block-start: 0;
}

.xlsx-authoring__range {
  color: var(--office-muted-foreground, #71717a);
  font-size: var(--office-text-xs, 0.75rem);
  margin: 0;
}

.xlsx-authoring__status {
  color: var(--office-success-foreground, #15803d);
  font-size: var(--office-text-xs, 0.75rem);
  margin: 0;
}

.xlsx-authoring__section {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-2, 0.5rem);
}

.xlsx-authoring__tabs {
  display: flex;
  gap: var(--office-ui-space-1, 0.25rem);
}

.xlsx-authoring__tab {
  background: transparent;
  border: var(--office-border-width, 1px) solid var(--office-border, #e4e4e7);
  border-radius: var(--office-radius-md, 0.5rem);
  color: inherit;
  cursor: pointer;
  font: inherit;
  font-size: var(--office-text-xs, 0.75rem);
  padding: 5px 10px;
}

.xlsx-authoring__tab--active {
  background: var(--office-accent, rgba(37, 99, 235, 0.14));
  border-color: var(--office-primary, #2563eb);
  color: var(--office-accent-foreground, #18181b);
}

.xlsx-authoring__rule-actions {
  display: flex;
  gap: var(--office-ui-space-2, 0.5rem);
}

.xlsx-authoring__thresholds {
  border: var(--office-border-width, 1px) solid var(--office-border, #e4e4e7);
  border-radius: var(--office-radius-md, 0.5rem);
  display: flex;
  flex-direction: column;
  gap: 6px;
  margin: 0;
  padding: 8px;
}

.xlsx-authoring__thresholds legend {
  color: var(--office-muted-foreground, #71717a);
  font-size: var(--office-text-xs, 0.75rem);
  padding: 0 4px;
}

.xlsx-authoring__threshold {
  display: flex;
  gap: var(--office-ui-space-2, 0.5rem);
}

.xlsx-authoring__threshold select {
  flex: 1 1 55%;
}

.xlsx-authoring__threshold input {
  flex: 1 1 45%;
  min-width: 0;
}

.xlsx-authoring__colors {
  display: flex;
  flex-wrap: wrap;
  gap: var(--office-ui-space-2, 0.5rem);
}

.xlsx-authoring__color {
  align-items: center;
  display: flex;
  font-size: var(--office-text-xs, 0.75rem);
  gap: 6px;
}

.xlsx-authoring__color input {
  block-size: 24px;
  border: none;
  cursor: pointer;
  inline-size: 32px;
  padding: 0;
}

.xlsx-authoring button:focus-visible,
.xlsx-authoring input:focus-visible,
.xlsx-authoring select:focus-visible,
.xlsx-authoring textarea:focus-visible {
  outline: 2px solid var(--office-ring, #2563eb);
  outline-offset: var(--office-ring-offset-width, 2px);
}

@media (pointer: coarse) {
  .xlsx-authoring__button,
  .xlsx-authoring__tab {
    min-height: 44px;
  }
}
</style>
