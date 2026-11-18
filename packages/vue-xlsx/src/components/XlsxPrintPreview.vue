<template>
  <section
    class="xlsx-print-preview"
    :class="{ 'xlsx-print-preview--dark': isDark }"
    data-testid="xlsx-print-preview"
    role="dialog"
    aria-modal="true"
    :aria-label="m.xlsx.print.title"
    :style="{ '--xlsx-print-preview-scale': String(previewScale) }"
  >
    <header class="xlsx-print-preview__toolbar">
      <div class="xlsx-print-preview__identity">
        <strong>{{ m.xlsx.print.title }}</strong>
        <span data-testid="xlsx-print-summary">{{ summary }}</span>
      </div>
      <div class="xlsx-print-preview__actions">
        <button
          type="button"
          class="xlsx-print-preview__button xlsx-print-preview__button--primary"
          data-testid="xlsx-print-print"
          @click="onPrint"
        >{{ m.xlsx.print.print }}</button>
        <button
          type="button"
          class="xlsx-print-preview__button"
          data-testid="xlsx-print-close"
          @click="emit('close')"
        >{{ m.xlsx.print.close }}</button>
      </div>
    </header>
    <div ref="scrollRef" class="xlsx-print-preview__pages" data-testid="xlsx-print-pages">
      <article
        v-for="view in pageViews"
        :key="view.page.index"
        class="xlsx-print-page"
        data-testid="xlsx-print-page"
        :data-page-number="view.page.pageNumber"
        :data-page-range="view.rangeLabel"
        :data-page-paper="view.page.paperSizeName"
        :style="view.frameStyle"
      >
        <div class="xlsx-print-page__sheet" :style="view.sheetStyle">
          <div class="xlsx-print-page__content" :style="view.contentStyle">
            <table
              class="xlsx-print-page__table"
              :class="{ 'xlsx-print-page__table--grid': view.gridLines }"
              :style="view.tableStyle"
            >
              <colgroup>
                <col v-if="view.headings" style="width: 34px" />
                <col
                  v-for="column in view.columns"
                  :key="`col-${column.index}`"
                  :style="{ width: `${column.widthPx}px` }"
                />
              </colgroup>
              <tbody>
                <tr v-if="view.headings" class="xlsx-print-page__headings">
                  <th class="xlsx-print-page__corner" />
                  <th
                    v-for="column in view.columns"
                    :key="`heading-${column.index}`"
                    class="xlsx-print-page__column-heading"
                  >{{ column.label }}</th>
                </tr>
                <tr
                  v-for="row in view.rows"
                  :key="`row-${row.index}`"
                  :style="{ height: `${row.heightPx}px` }"
                >
                  <th
                    v-if="view.headings"
                    class="xlsx-print-page__row-heading"
                  >{{ row.index + 1 }}</th>
                  <td
                    v-for="cell in row.cells"
                    :key="`cell-${cell.row}-${cell.col}`"
                    :rowspan="cell.rowSpan"
                    :colspan="cell.colSpan"
                    :style="cell.style"
                  >{{ cell.text }}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </article>
      <p v-if="pageViews.length === 0" class="xlsx-print-preview__empty">{{ m.xlsx.print.noPages }}</p>
    </div>
  </section>
</template>

<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref, watch, type CSSProperties } from "vue";
import {
  buildPrintPageGrid,
  resolveWorkbookColor,
  resolveWorkbookFillStyle,
  type XlsxPrintLayout,
  type XlsxPrintPage,
  type XlsxResolvedCellStyle,
  type XlsxThemePalette,
  type XlsxViewerController,
} from "@arcships/xlsx-core";
import { formatOfficeMessage } from "@arcships/office-interaction";
import { useXlsxMessages } from "../i18n";

const props = withDefaults(defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
}>(), {
  isDark: false,
});

const emit = defineEmits<{
  close: [];
}>();

const m = useXlsxMessages();
const scrollRef = ref<HTMLElement | null>(null);
const previewScale = ref(0.8);
let resizeObserver: ResizeObserver | null = null;
let workerHydrationToken = 0;
const workerCellValues = new Map<string, string>();
const workerValuesRevision = ref(0);

interface PrintCellView {
  col: number;
  colSpan: number;
  row: number;
  rowSpan: number;
  style: CSSProperties;
  text: string;
}

interface PrintRowView {
  cells: PrintCellView[];
  heightPx: number;
  index: number;
}

interface PrintColumnView {
  index: number;
  label: string;
  widthPx: number;
}

interface PrintPageView {
  columns: PrintColumnView[];
  contentStyle: CSSProperties;
  frameStyle: CSSProperties;
  gridLines: boolean;
  headings: boolean;
  page: XlsxPrintPage;
  rangeLabel: string;
  rows: PrintRowView[];
  sheetStyle: CSSProperties;
  tableStyle: CSSProperties;
}

const layout = computed<XlsxPrintLayout | null>(() => props.controller.printLayout ?? null);
const pages = computed(() => layout.value?.pages ?? []);
const maxPageWidth = computed(() => pages.value.reduce((width, page) => Math.max(width, page.widthPx), 0));

const summary = computed(() => {
  const current = layout.value;
  if (!current) return "";
  return formatOfficeMessage(m.value.xlsx.print.summary, {
    count: String(current.pages.length),
    orientation: current.pageSetup.orientation === "landscape"
      ? m.value.xlsx.print.orientationLandscape
      : m.value.xlsx.print.orientationPortrait,
    paper: current.paperSize.name,
    scale: String(current.pages[0]?.scale ?? current.scale),
  });
});

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return value && typeof value === "object" ? value as Record<string, unknown> : undefined;
}

function borderEdge(edge: unknown, themePalette: XlsxThemePalette): string | undefined {
  const style = asRecord(edge);
  if (!style?.style || style.style === "none") return undefined;
  const rawStyle = String(style.style);
  const cssStyle = rawStyle.includes("Dash")
    ? "dashed"
    : rawStyle === "dotted" || rawStyle === "hair"
      ? "dotted"
      : rawStyle === "double"
        ? "double"
        : "solid";
  const width = rawStyle.startsWith("medium") ? 2 : rawStyle === "thick" ? 3 : 1;
  const color = resolveWorkbookColor(asRecord(style.color), themePalette) ?? "#52525b";
  return `${width}px ${cssStyle} ${color}`;
}

function cellStyle(
  row: number,
  col: number,
  scale: number,
  gridLines: boolean,
  sheet: NonNullable<XlsxViewerController["activeSheet"]>,
  worksheet: ReturnType<XlsxViewerController["getActiveWorksheet"]>,
): CSSProperties {
  let workbookStyle: XlsxResolvedCellStyle | undefined;
  if (worksheet) {
    try {
      workbookStyle = worksheet.getCellStyleAt(row, col) as XlsxResolvedCellStyle | undefined;
    } catch {
      workbookStyle = undefined;
    }
  }
  if (!workbookStyle) {
    const rowStyleId = sheet.rowStyleIds[row];
    const colStyleId = sheet.colStyleIds[col];
    workbookStyle = rowStyleId !== undefined
      ? sheet.styleById[rowStyleId]
      : colStyleId !== undefined ? sheet.styleById[colStyleId] : undefined;
  }
  const font = asRecord(workbookStyle?.font);
  const fill = asRecord(workbookStyle?.fill);
  const alignment = asRecord(workbookStyle?.alignment);
  const border = asRecord(workbookStyle?.border);
  const fillStyle = resolveWorkbookFillStyle(fill, sheet.themePalette);
  const fontSizePx = Math.max(6, (typeof font?.size === "number" ? font.size * 4 / 3 : 12) * scale / 100);
  const decoration = [
    font?.underline && font.underline !== "none" ? "underline" : "",
    font?.strikethrough ? "line-through" : "",
  ].filter(Boolean).join(" ");
  return {
    backgroundColor: fillStyle.backgroundColor ?? "#ffffff",
    backgroundImage: fillStyle.backgroundImage ?? undefined,
    borderBottom: borderEdge(border?.bottom, sheet.themePalette) ?? (gridLines ? "1px solid #d4d4d8" : undefined),
    borderLeft: borderEdge(border?.left, sheet.themePalette) ?? (gridLines ? "1px solid #d4d4d8" : undefined),
    borderRight: borderEdge(border?.right, sheet.themePalette) ?? (gridLines ? "1px solid #d4d4d8" : undefined),
    borderTop: borderEdge(border?.top, sheet.themePalette) ?? (gridLines ? "1px solid #d4d4d8" : undefined),
    color: resolveWorkbookColor(asRecord(font?.color), sheet.themePalette) ?? "#18181b",
    fontFamily: typeof font?.name === "string" && font.name.trim() ? font.name : "Segoe UI",
    fontSize: `${fontSizePx}px`,
    fontStyle: font?.italic ? "italic" : "normal",
    fontWeight: font?.bold ? "700" : "400",
    overflow: "hidden",
    padding: "0 3px",
    textAlign: alignment?.horizontal === "center" ? "center" : alignment?.horizontal === "right" ? "right" : "left",
    textDecoration: decoration || undefined,
    verticalAlign: alignment?.vertical === "top" ? "top" : alignment?.vertical === "center" ? "middle" : "bottom",
    whiteSpace: alignment?.wrapText ? "pre-wrap" : "nowrap",
    wordBreak: "break-word",
  };
}

function pageRangeLabel(range: XlsxPrintPage["range"]) {
  const start = `${columnLabel(range.start.col)}${range.start.row + 1}`;
  const end = `${columnLabel(range.end.col)}${range.end.row + 1}`;
  return start === end ? start : `${start}:${end}`;
}

function columnLabel(col: number) {
  let label = "";
  let value = col + 1;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function cellKey(row: number, col: number) {
  return `${row}:${col}`;
}

const pageViews = computed<PrintPageView[]>(() => {
  const current = layout.value;
  const sheet = props.controller.activeSheet;
  if (!current || !sheet) return [];
  const gridLines = current.printOptions.gridLines;
  const marginLeft = Math.round(current.pageMargins.left * 96);
  const marginTop = Math.round(current.pageMargins.top * 96);
  void workerValuesRevision.value;
  const worksheet = props.controller.getActiveWorksheet();
  return current.pages.map((page) => {
    const grid = buildPrintPageGrid(sheet, page, current.printTitles);
    const scale = page.scale / 100;
    const cellsByRow = new Map<number, PrintCellView[]>();
    for (const cell of grid.cells) {
      if (cell.covered) continue;
      const list = cellsByRow.get(cell.row) ?? [];
      list.push({
        col: cell.col,
        colSpan: cell.colSpan,
        row: cell.row,
        rowSpan: cell.rowSpan,
        style: cellStyle(cell.row, cell.col, page.scale, gridLines, sheet, worksheet),
        text: workerCellValues.get(cellKey(cell.row, cell.col))
          ?? props.controller.getCellDisplayValue({ col: cell.col, row: cell.row }),
      });
      cellsByRow.set(cell.row, list);
    }
    const columns = grid.columns.map((column) => ({
      index: column.index,
      label: columnLabel(column.index),
      widthPx: column.sizePx * scale,
    }));
    const rows = grid.rows.map((row) => ({
      cells: (cellsByRow.get(row.index) ?? []).sort((left, right) => left.col - right.col),
      heightPx: row.sizePx * scale,
      index: row.index,
    }));
    const tableWidth = columns.reduce((total, column) => total + column.widthPx, 0);
    const headings = current.printOptions.headings;
    const headerWidth = headings ? 34 : 0;
    return {
      columns,
      contentStyle: {
        alignItems: current.printOptions.horizontalCentered ? "center" : "flex-start",
        display: "flex",
        flexDirection: "column",
        height: `${page.contentHeightPx}px`,
        justifyContent: current.printOptions.verticalCentered ? "center" : "flex-start",
        left: `${marginLeft}px`,
        overflow: "hidden",
        position: "absolute",
        top: `${marginTop}px`,
        width: `${page.contentWidthPx}px`,
      },
      frameStyle: {
        "--xlsx-page-height": `${page.heightPx}px`,
        "--xlsx-page-width": `${page.widthPx}px`,
      } as CSSProperties,
      gridLines,
      headings,
      page,
      rangeLabel: pageRangeLabel(page.range),
      rows,
      sheetStyle: {
        height: `${page.heightPx}px`,
        width: `${page.widthPx}px`,
      },
      tableStyle: {
        borderCollapse: "collapse",
        tableLayout: "fixed",
        width: `${Math.min(page.contentWidthPx, Math.max(1, tableWidth + headerWidth))}px`,
      },
    };
  });
});

async function hydrateWorkerCells() {
  const current = layout.value;
  const sheet = props.controller.activeSheet;
  const snapshot = props.controller.getCellSnapshotAsync;
  if (!current || !sheet || !props.controller.isWorkerBacked || !snapshot) return;
  const token = ++workerHydrationToken;
  workerCellValues.clear();
  workerValuesRevision.value += 1;
  const targets: Array<{ col: number; row: number }> = [];
  const seen = new Set<string>();
  const hydrationLimit = 20_000;
  for (const page of current.pages) {
    const grid = buildPrintPageGrid(sheet, page, current.printTitles);
    for (const cell of grid.cells) {
      if (cell.covered) continue;
      const key = cellKey(cell.row, cell.col);
      if (seen.has(key)) continue;
      seen.add(key);
      targets.push({ col: cell.col, row: cell.row });
      if (targets.length >= hydrationLimit) break;
    }
    if (targets.length >= hydrationLimit) break;
  }
  let index = 0;
  let applied = 0;
  const worker = async () => {
    while (index < targets.length && token === workerHydrationToken) {
      const cell = targets[index++]!;
      try {
        const value = await snapshot.call(
          props.controller,
          sheet.workbookSheetIndex,
          cell.row,
          cell.col,
        );
        if (token !== workerHydrationToken) return;
        workerCellValues.set(cellKey(cell.row, cell.col), value.displayValue);
        applied += 1;
        if (applied % 50 === 0) workerValuesRevision.value += 1;
      } catch {
        // A missing snapshot leaves the cached display value in place.
      }
    }
  };
  await Promise.all(Array.from({ length: Math.min(8, targets.length) }, () => worker()));
  if (token === workerHydrationToken) workerValuesRevision.value += 1;
}

watch(
  () => [pages.value.length, layout.value?.sheetName, props.controller.revision] as const,
  () => { void hydrateWorkerCells(); },
  { immediate: true },
);

function updatePreviewScale() {
  const container = scrollRef.value;
  const width = container?.clientWidth ?? 0;
  if (!width || !maxPageWidth.value) {
    previewScale.value = 0.8;
    return;
  }
  previewScale.value = Math.max(0.25, Math.min(1, (width - 48) / maxPageWidth.value));
}

onMounted(() => {
  updatePreviewScale();
  if (typeof ResizeObserver !== "undefined" && scrollRef.value) {
    resizeObserver = new ResizeObserver(updatePreviewScale);
    resizeObserver.observe(scrollRef.value);
  }
});

let ownsPrintPageStyle = false;
let printPageStyle: HTMLStyleElement | undefined;

function syncPrintPageSize() {
  if (typeof document === "undefined") return;
  const firstPage = layout.value?.pages[0];
  if (!firstPage) return;
  if (!printPageStyle) {
    const existing = document.querySelector?.('style[data-office-print-page="xlsx"]') as HTMLStyleElement | null;
    printPageStyle = existing ?? document.createElement("style");
    ownsPrintPageStyle = !existing;
    if (ownsPrintPageStyle) {
      printPageStyle.dataset.officePrintPage = "xlsx";
      document.head.append(printPageStyle);
    }
  }
  printPageStyle.textContent = `@page { size: ${(firstPage.widthPx / 96).toFixed(2)}in ${(firstPage.heightPx / 96).toFixed(2)}in; margin: 0; }`;
}

watch(
  () => [layout.value?.pages[0]?.widthPx, layout.value?.pages[0]?.heightPx, layout.value?.pages.length] as const,
  syncPrintPageSize,
  { immediate: true },
);

onBeforeUnmount(() => {
  workerHydrationToken += 1;
  resizeObserver?.disconnect();
  resizeObserver = null;
  if (ownsPrintPageStyle) printPageStyle?.remove?.();
});

watch(maxPageWidth, updatePreviewScale);

function onPrint() {
  props.controller.print();
}
</script>

<style scoped>
.xlsx-print-preview {
  background: #525659;
  display: flex;
  flex: 1 1 auto;
  flex-direction: column;
  inset: 0;
  min-height: 0;
  position: absolute;
  z-index: 60;
}
.xlsx-print-preview--dark { background: #27272a; }
.xlsx-print-preview__toolbar {
  align-items: center;
  background: #ffffff;
  border-bottom: 1px solid #d4d4d8;
  color: #18181b;
  display: flex;
  flex: 0 0 auto;
  gap: 12px;
  justify-content: space-between;
  padding: 8px 12px;
}
.xlsx-print-preview--dark .xlsx-print-preview__toolbar { background: #18181b; border-color: #3f3f46; color: #f4f4f5; }
.xlsx-print-preview__identity { align-items: baseline; display: flex; gap: 10px; min-width: 0; }
.xlsx-print-preview__identity strong { font-size: 13px; }
.xlsx-print-preview__identity span { color: #71717a; font-size: 12px; }
.xlsx-print-preview__actions { display: flex; gap: 8px; }
.xlsx-print-preview__button {
  background: #f4f4f5;
  border: 1px solid #d4d4d8;
  border-radius: 6px;
  color: inherit;
  cursor: pointer;
  font: inherit;
  font-size: 12px;
  padding: 5px 10px;
}
.xlsx-print-preview__button:hover { background: #e4e4e7; }
.xlsx-print-preview__button--primary { background: #2563eb; border-color: #2563eb; color: #ffffff; }
.xlsx-print-preview__button--primary:hover { background: #1d4ed8; }
.xlsx-print-preview__pages {
  align-items: flex-start;
  display: flex;
  flex: 1 1 auto;
  flex-direction: column;
  gap: 18px;
  min-height: 0;
  overflow: auto;
  padding: 24px;
}
.xlsx-print-preview__empty { color: #e4e4e7; font-size: 13px; margin: 24px auto; }
.xlsx-print-page {
  background: #ffffff;
  box-shadow: 0 2px 14px rgba(0, 0, 0, .35);
  flex: 0 0 auto;
  height: calc(var(--xlsx-page-height) * var(--xlsx-print-preview-scale, .8));
  overflow: hidden;
  position: relative;
  width: calc(var(--xlsx-page-width) * var(--xlsx-print-preview-scale, .8));
}
.xlsx-print-page__sheet {
  background: #ffffff;
  color: #18181b;
  transform: scale(var(--xlsx-print-preview-scale, .8));
  transform-origin: top left;
}
.xlsx-print-page__table { table-layout: fixed; }
.xlsx-print-page__table td { overflow: hidden; }
.xlsx-print-page__table th { background: #f4f4f5; border: 1px solid #d4d4d8; color: #52525b; font-size: 9px; font-weight: 400; overflow: hidden; padding: 0; text-align: center; vertical-align: middle; }

@media print {
  .xlsx-print-preview { background: #ffffff; display: block; inset: auto; overflow: visible; position: static; }
  .xlsx-print-preview__toolbar { display: none !important; }
  .xlsx-print-preview__pages { display: block; overflow: visible; padding: 0; }
  .xlsx-print-page {
    break-after: page;
    box-shadow: none;
    height: var(--xlsx-page-height) !important;
    margin: 0;
    overflow: hidden;
    page-break-after: always;
    width: var(--xlsx-page-width) !important;
  }
  .xlsx-print-page:last-child { break-after: auto; page-break-after: auto; }
  .xlsx-print-page__sheet { transform: none !important; }
}
</style>
