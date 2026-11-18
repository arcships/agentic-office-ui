<template>
  <div
    ref="viewerRef"
    class="xlsx-viewer"
    :class="{ 'xlsx-viewer--print-preview': controller.isPrintPreviewOpen }"
    data-testid="xlsx-viewer"
    :data-state="controller.isLoading ? 'loading' : controller.error || controller.sourceError ? 'error' : controller.isLoadDeferred ? 'deferred' : controller.activeSheet || controller.activeTab?.kind === 'chartsheet' ? 'ready' : 'empty'"
    :style="viewerStyle"
    @keydown="onKeydown"
    tabindex="0"
  >
    <XlsxToolbar
      v-if="showDefaultToolbar"
      ref="toolbarRef"
      :controller="controller"
      :is-dark="isDark ?? false"
      :search="surfaceSearch"
      :show-upload="showUpload"
      @search-close="viewerRef?.focus()"
      @upload="emit('upload')"
    />
    <XlsxRibbon
      v-if="showRibbon"
      :controller="controller"
      :is-dark="isDark ?? false"
      :read-only="effectiveReadOnly"
      @authoring="openAuthoring"
      @update:read-only="onReadOnlyChange"
    />
    <XlsxFormulaBar
      v-if="showFormulaBar"
      :controller="controller"
      :is-dark="isDark ?? false"
      :read-only="effectiveReadOnly"
    />
    <div
      v-if="controller.calculationStatus"
      class="xlsx-viewer__calculation-status"
      data-testid="xlsx-calculation-status"
      role="status"
    >
      <slot
        name="calculation-status"
        :status="controller.calculationStatus"
      >
        {{ calculationMessage }}
      </slot>
    </div>
    <OfficeLiveStatus
      test-id="xlsx-live-status"
      :message="liveStatusMessage"
      :busy="controller.isLoading"
    />
    <div ref="bodyRef" class="xlsx-viewer__body">
      <div
        v-if="controller.isLoading"
        class="xlsx-viewer__loading"
      >
        <slot name="loading">
          <span class="xlsx-viewer__loading-text">{{ m.xlsx.viewer.loading }}</span>
        </slot>
      </div>
      <div
        v-else-if="controller.error || controller.sourceError"
        class="xlsx-viewer__error"
      >
        <slot
          name="error"
          :error="localizedError"
        >
          <span class="xlsx-viewer__error-text">{{ localizedError?.message }}</span>
        </slot>
      </div>
      <XlsxChartsheetSurface
        v-else-if="controller.activeTab?.kind === 'chartsheet'"
        :controller="controller"
        :is-dark="isDark ?? false"
      />
      <div
        v-else-if="controller.isLoadDeferred"
        class="xlsx-viewer__deferred"
        data-testid="xlsx-deferred-load"
      >
        <slot
          name="deferredLoad"
          :file-name="deferredLoadFileName"
          :bytes="deferredLoadBytes"
          :size-label="deferredLoadSizeLabel"
          :resume="onDeferredResume"
          :cancel="onDeferredCancel"
        >
          <div
            class="xlsx-viewer__deferred-card"
            :class="{ 'xlsx-viewer__deferred-card--dark': isDark }"
          >
            <div class="xlsx-viewer__deferred-title">{{ m.xlsx.viewer.deferredTitle }}</div>
            <div
              class="xlsx-viewer__deferred-file"
              data-testid="xlsx-deferred-file-name"
            >
              {{ deferredLoadFileName }}
            </div>
            <div
              class="xlsx-viewer__deferred-size"
              data-testid="xlsx-deferred-size"
            >
              {{ formatOfficeMessage(m.xlsx.viewer.deferredDetail, { size: deferredLoadSizeLabel }) }}
            </div>
            <div class="xlsx-viewer__deferred-actions">
              <button
                type="button"
                class="xlsx-viewer__deferred-button xlsx-viewer__deferred-button--primary"
                data-testid="xlsx-deferred-continue"
                :disabled="controller.canLoadDeferred === false"
                @click="onDeferredResume"
              >
                {{ m.xlsx.viewer.resumeParsing }}
              </button>
              <button
                type="button"
                class="xlsx-viewer__deferred-button"
                data-testid="xlsx-deferred-cancel"
                @click="onDeferredCancel"
              >
                {{ m.xlsx.viewer.cancel }}
              </button>
            </div>
          </div>
        </slot>
      </div>
      <div
        v-else-if="!controller.activeSheet"
        class="xlsx-viewer__empty"
      >
        <slot name="empty">
          <span class="xlsx-viewer__empty-text">{{ m.xlsx.viewer.empty }}</span>
        </slot>
      </div>
      <template v-else>
        <XlsxGrid
          ref="gridRef"
          :controller="controller"
          :get-cell-style="getCellStyle"
          :is-dark="isDark ?? false"
          :read-only="effectiveReadOnly"
          :selection-color="selectionColor"
          :selection-fill-color="selectionFillColor"
          @cell-double-click="onCellDoubleClick"
          @viewport-change="gridViewport = $event"
        />
        <XlsxChartOverlay
          :controller="controller"
          @chart-context-menu="onChartContextMenu"
          :is-dark="isDark ?? false"
          :scroll-left="gridViewport.scrollLeft"
          :scroll-top="gridViewport.scrollTop"
          :viewport-width="gridViewport.width"
          :viewport-height="gridViewport.height"
        />
        <XlsxImageLayer
          :controller="controller"
          :show-images="showImages ?? true"
          :scroll-left="gridViewport.scrollLeft"
          :scroll-top="gridViewport.scrollTop"
          :viewport-width="gridViewport.width"
          :viewport-height="gridViewport.height"
        />
        <XlsxDrawingLayer
          :controller="controller"
          :scroll-left="gridViewport.scrollLeft"
          :scroll-top="gridViewport.scrollTop"
          :viewport-width="gridViewport.width"
          :viewport-height="gridViewport.height"
        />
        <XlsxSelectionOverlay
          :controller="controller"
          :get-cell-style="getCellStyle"
          :selection-color="selectionColor"
          :selection-fill-color="selectionFillColor"
        />
        <XlsxContextMenu
          ref="contextMenuRef"
          :controller="controller"
          :target-element="gridElement"
          :suppress-context-menu="touchGestures.shouldSuppressContextMenu"
          @authoring="openAuthoring"
        />
      </template>
      <XlsxPrintPreview
        v-if="controller.isPrintPreviewOpen"
        :controller="controller"
        :is-dark="isDark ?? false"
        @close="controller.closePrintPreview()"
      />
    </div>
    <XlsxAuthoringDialog
      v-if="authoringKind"
      :controller="controller"
      :kind="authoringKind"
      :is-dark="isDark ?? false"
      @close="authoringKind = null"
    />
    <XlsxSheetTabs
      v-if="controller.tabs.length > 0"
      :controller="controller"
      :is-dark="isDark ?? false"
      :read-only="effectiveReadOnly"
    />
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch, type CSSProperties } from "vue";
import { clampSurfaceZoom, nextSurfaceZoom } from "@arcships/office-runtime/gesture-zoom";
import { createTouchGestureTracker, nextDoubleTapZoom } from "@arcships/office-runtime/touch-gestures";
import type { XlsxViewerController, XlsxCellAddress, XlsxCellStyleContext } from "@arcships/xlsx-core";
import type { OfficeLocale, OfficeMessagesOverride } from "@arcships/office-interaction";
import XlsxGrid from "./XlsxGrid.vue";
import XlsxToolbar from "./XlsxToolbar.vue";
import XlsxRibbon from "./XlsxRibbon.vue";
import XlsxFormulaBar from "./XlsxFormulaBar.vue";
import XlsxSheetTabs from "./XlsxSheetTabs.vue";
import XlsxChartOverlay from "./XlsxChartOverlay.vue";
import XlsxImageLayer from "./XlsxImageLayer.vue";
import XlsxDrawingLayer from "./XlsxDrawingLayer.vue";
import XlsxSelectionOverlay from "./XlsxSelectionOverlay.vue";
import XlsxContextMenu from "./XlsxContextMenu.vue";
import XlsxAuthoringDialog from "./XlsxAuthoringDialog.vue";
import XlsxPrintPreview from "./XlsxPrintPreview.vue";
import type { XlsxAuthoringDialogKind } from "./authoring-dialog";
import XlsxChartsheetSurface from "./XlsxChartsheetSurface.vue";
import { formatBinaryBytes } from "../composables/formatting";
import { OfficeLiveStatus } from "@arcships/vue-ui";
import { useXlsxSearch, type XlsxSearchState } from "../composables/useXlsxSearch";
import { formatOfficeMessage } from "@arcships/office-interaction";
import { provideXlsxMessages } from "../i18n";
import { localizeXlsxLoadError } from "../errors";

const props = withDefaults(
  defineProps<{
    controller: XlsxViewerController;
    getCellStyle?: ((cell: XlsxCellAddress, context?: XlsxCellStyleContext) => Partial<CSSProperties> | undefined) | null;
    height?: string;
    /** Optional minimum height; the viewer itself does not impose one. */
    minHeight?: string | number;
    isDark?: boolean | null;
    rounded?: boolean;
    readOnly?: boolean;
    selectionColor?: string;
    selectionFillColor?: string;
    showDefaultToolbar?: boolean;
    showRibbon?: boolean;
    showFormulaBar?: boolean;
    showImages?: boolean;
    showUpload?: boolean;
    /** Enables trackpad, pinch and double-tap zoom for the composite viewer. */
    enableGestureZoom?: boolean;
    locale?: OfficeLocale;
    /**
     * Locale for number and date formats that name none of their own
     * (`[$-407]`-style sections win). Defaults to the controller's locale.
     */
    documentLocale?: string;
    messages?: OfficeMessagesOverride;
  }>(),
  {
    getCellStyle: null,
    height: "100%",
    minHeight: undefined,
    isDark: false,
    rounded: true,
    readOnly: false,
    selectionColor: undefined,
    selectionFillColor: undefined,
    showDefaultToolbar: true,
    showRibbon: true,
    showFormulaBar: true,
    showImages: true,
    showUpload: false,
    enableGestureZoom: true,
    locale: "en",
    documentLocale: undefined,
    messages: undefined,
  }
);

const emit = defineEmits<{
  cellDoubleClick: [cell: XlsxCellAddress];
  upload: [];
  "update:readOnly": [value: boolean];
  searchStateChange: [state: XlsxSearchState];
  "load-deferred": [payload: XlsxDeferredLoadPayload];
  "deferred-resumed": [payload: XlsxDeferredLoadPayload];
  "deferred-cancel": [payload: XlsxDeferredLoadPayload];
}>();

type XlsxDeferredLoadPayload = { fileName: string; bytes: number };

const effectiveReadOnly = computed(() => props.readOnly || props.controller.isEditingLocked === true || props.controller.readOnly);
const m = provideXlsxMessages(() => props.locale, () => props.messages);
const localizedError = computed(() => {
  const sourceError = props.controller.sourceError;
  if (sourceError) {
    return { ...sourceError, message: localizeXlsxLoadError(sourceError, m.value.xlsx.errors) };
  }
  const error = props.controller.error;
  return error ? { ...error, message: error.message } : null;
});
const calculationMessage = computed(() => {
  const status = props.controller.calculationStatus;
  if (!status) return "";
  if (status.reason === "unresolved-sheet-refs") return m.value.xlsx.calculation.sheetMissing;
  if (status.reason === "recovery-snapshot-unavailable") return m.value.xlsx.calculation.snapshotUnavailable;
  return m.value.xlsx.calculation.failed;
});

// Number/date formats without a `[$-…]` section follow the host locale.
watch(
  () => props.documentLocale,
  (locale) => { props.controller.setDocumentLocale?.(locale); },
  { immediate: true },
);

const deferredLoadBytes = computed(() => props.controller.deferredLoadFileSize ?? 0);
const deferredLoadSizeLabel = computed(() => formatBinaryBytes(deferredLoadBytes.value));
const deferredLoadFileName = computed(() => props.controller.displayFileName?.trim() || m.value.xlsx.viewer.deferredUntitled);
const deferredLoadPayload = computed<XlsxDeferredLoadPayload>(() => ({
  fileName: deferredLoadFileName.value,
  bytes: deferredLoadBytes.value,
}));

function onDeferredResume() {
  if (!props.controller.isLoadDeferred) return;
  emit("deferred-resumed", deferredLoadPayload.value);
  props.controller.continueDeferredLoad();
}

function onDeferredCancel() {
  if (!props.controller.isLoadDeferred) return;
  emit("deferred-cancel", deferredLoadPayload.value);
  props.controller.cancelDeferredLoad?.();
}

watch(
  () => props.controller.isLoadDeferred,
  (deferred) => {
    if (deferred) emit("load-deferred", deferredLoadPayload.value);
  },
  { immediate: true },
);

// Loading, error and sheet switches were rendered visually only. This live
// region is the shared Office announcement channel for workbook state; the
// selected cell is announced by XlsxSelectionOverlay.
const liveStatusMessage = computed(() => {
  if (props.controller.isLoading) return m.value.xlsx.viewer.live.loading;
  if (props.controller.error || props.controller.sourceError) return formatOfficeMessage(m.value.xlsx.viewer.live.error, { message: localizedError.value?.message ?? props.controller.error?.message ?? "" });
  const sheet = props.controller.activeSheet;
  if (sheet) return formatOfficeMessage(m.value.xlsx.viewer.live.sheetReady, { name: sheet.name });
  if (props.controller.activeTab?.kind === "chartsheet") return m.value.xlsx.viewer.live.chartsheetReady;
  return m.value.xlsx.viewer.live.empty;
});

type GridInstance = InstanceType<typeof XlsxGrid>;
const viewerRef = ref<HTMLElement | null>(null);
const bodyRef = ref<HTMLElement | null>(null);
const toolbarRef = ref<InstanceType<typeof XlsxToolbar> | null>(null);
const gridRef = ref<GridInstance | null>(null);
const contextMenuRef = ref<{
  openAt(clientX: number, clientY: number, target?: unknown): boolean
  openForChart(chartId: string, clientX: number, clientY: number): boolean
} | null>(null);
const gridViewport = ref({ height: 0, scrollLeft: 0, scrollTop: 0, width: 0 });
const authoringKind = ref<XlsxAuthoringDialogKind | null>(null);

function openAuthoring(kind: XlsxAuthoringDialogKind) {
  authoringKind.value = kind;
}
const gridElement = computed<HTMLElement | null>(() => gridRef.value?.scrollContainer ?? null);
const surfaceSearch = useXlsxSearch(
  () => props.controller,
  {
    scrollToCell: async (cell) => {
      await nextTick();
      gridRef.value?.scrollToCell(cell);
    },
  },
);

const viewerStyle = computed<CSSProperties>(() => ({
  blockSize: props.height,
  borderRadius: props.rounded ? "12px" : "0px",
  display: "flex",
  flex: "1 1 auto",
  flexDirection: "column",
  inlineSize: "100%",
  isolation: "isolate",
  maxHeight: "100%",
  maxWidth: "100%",
  minHeight: typeof props.minHeight === "number" ? `${props.minHeight}px` : props.minHeight ?? "0px",
  minWidth: "0px",
  overflow: "hidden",
  position: "relative",
  width: "100%",
  backgroundColor: props.isDark ? "#18181b" : "#ffffff",
  color: props.isDark ? "#e4e4e7" : "#18181b",
  outline: "none",
}));

let printPageStyle: HTMLStyleElement | undefined;

function syncPrintPageSize(): void {
  const browserDocument = globalThis.document;
  if (!browserDocument) return;
  const layout = props.controller.printLayout;
  const firstPage = props.controller.isPrintPreviewOpen ? layout?.pages[0] : undefined;
  if (firstPage) {
    if (!printPageStyle || printPageStyle.isConnected === false) {
      printPageStyle = browserDocument.createElement("style");
      if (printPageStyle.dataset) printPageStyle.dataset.officePrintPage = "xlsx";
      browserDocument.head.append(printPageStyle);
    }
    printPageStyle.textContent = `@page { size: ${(firstPage.widthPx / 96).toFixed(2)}in ${(firstPage.heightPx / 96).toFixed(2)}in; margin: 0; }`;
    return;
  }
  const sheet = props.controller.activeSheet;
  if (!sheet || sheet.maxUsedCol < sheet.minUsedCol || sheet.maxUsedRow < sheet.minUsedRow) {
    printPageStyle?.remove?.();
    printPageStyle = undefined;
    return;
  }
  let width = 0;
  for (let column = sheet.minUsedCol; column <= sheet.maxUsedCol; column++) {
    width += sheet.colWidths[column] ?? sheet.defaultColWidthPx;
  }
  let height = 0;
  for (let row = sheet.minUsedRow; row <= sheet.maxUsedRow; row++) {
    height += sheet.rowHeights[row] ?? sheet.defaultRowHeightPx;
  }
  if (width <= 0 || height <= 0) return;
  if (!printPageStyle) {
    printPageStyle = browserDocument.createElement("style");
    if (printPageStyle.dataset) printPageStyle.dataset.officePrintPage = "xlsx";
    browserDocument.head.append(printPageStyle);
  }
  printPageStyle.textContent = `@page { size: ${Math.round(width)}px ${Math.round(height)}px; margin: 0; }`;
}

watch(
  () => [
    props.controller.activeSheet?.name,
    props.controller.activeSheet?.minUsedCol,
    props.controller.activeSheet?.maxUsedCol,
    props.controller.activeSheet?.minUsedRow,
    props.controller.activeSheet?.maxUsedRow,
    props.controller.isPrintPreviewOpen,
    props.controller.printLayout?.pages[0]?.widthPx,
    props.controller.printLayout?.pages[0]?.heightPx,
  ] as const,
  syncPrintPageSize,
  { immediate: true },
);

function onKeydown(event: KeyboardEvent) {
  if (event.key === "Escape" && props.controller.isPrintPreviewOpen) {
    event.preventDefault();
    props.controller.closePrintPreview();
    return;
  }
  if (
    props.showDefaultToolbar
    && (event.ctrlKey || event.metaKey)
    && event.key.toLocaleLowerCase() === "f"
  ) {
    event.preventDefault();
    toolbarRef.value?.openSearch();
    return;
  }

  if (!props.controller || effectiveReadOnly.value) return;

  if ((event.ctrlKey || event.metaKey) && event.key === "z") {
    event.preventDefault();
    props.controller.undo();
    return;
  }
  if ((event.ctrlKey || event.metaKey) && event.key === "y") {
    event.preventDefault();
    props.controller.redo();
    return;
  }
}

function onCellDoubleClick(cell: XlsxCellAddress) {
  emit("cellDoubleClick", cell);
}

function onReadOnlyChange(value: boolean) {
  emit("update:readOnly", value);
}

// ── Gesture zoom (composite viewer) ─────────────────────────────────
type GridZoomAnchor = NonNullable<ReturnType<GridInstance["captureZoomAnchor"]>>;

let pendingAnchor: { anchor: GridZoomAnchor; requestedZoom: number; token: number } | undefined;
let pendingZoom: number | undefined;
let gestureToken = 0;

function currentZoom(): number {
  return clampSurfaceZoom(props.controller.zoomScale / 100);
}

function requestViewerZoom(nextZoom: number, clientX: number, clientY: number): void {
  if (props.enableGestureZoom === false) return;
  const current = currentZoom();
  const target = clampSurfaceZoom(nextZoom);
  if (target === (pendingZoom ?? current)) return;
  const token = ++gestureToken;
  const anchor = gridRef.value?.captureZoomAnchor(clientX, clientY) ?? null;
  pendingZoom = target;
  pendingAnchor = anchor ? { anchor, requestedZoom: target, token } : undefined;
  props.controller.setZoomScale(target * 100);
}

const touchGestures = createTouchGestureTracker({
  isEnabled: () => props.enableGestureZoom !== false,
  isDoubleTapEnabled: () => effectiveReadOnly.value,
  getZoom: currentZoom,
  onZoom: (zoom, center) => requestViewerZoom(zoom, center.clientX, center.clientY),
  onDoubleTap: (pointer) => requestViewerZoom(nextDoubleTapZoom(currentZoom()), pointer.clientX, pointer.clientY),
  onLongPress: (pointer) => {
    contextMenuRef.value?.openAt(pointer.clientX, pointer.clientY, pointer.target);
  },
});

function onChartContextMenu(payload: { chartId: string; clientX: number; clientY: number }): void {
  contextMenuRef.value?.openForChart(payload.chartId, payload.clientX, payload.clientY);
}

function onTouchPointerDown(event: PointerEvent): void { touchGestures.pointerDown(event); }
function onTouchPointerMove(event: PointerEvent): void { touchGestures.pointerMove(event); }
function onTouchPointerUp(event: PointerEvent): void { touchGestures.pointerUp(event); }
function onTouchPointerCancel(event: PointerEvent): void { touchGestures.pointerCancel(event); }

function onWheel(event: WheelEvent): void {
  if (props.enableGestureZoom === false || event.defaultPrevented) return;
  // Chrome/Edge report trackpad pinch as ctrl+wheel; macOS reports Cmd+wheel
  // for explicit zoom shortcuts. Both must stay inside the viewer.
  if (!event.ctrlKey && !event.metaKey) return;
  event.preventDefault();
  requestViewerZoom(
    nextSurfaceZoom(pendingZoom ?? currentZoom(), event.deltaY, event.deltaMode),
    event.clientX,
    event.clientY,
  );
}

watch(() => props.controller.zoomScale, (value) => {
  if (typeof value !== "number" || !Number.isFinite(value)) return;
  const normalized = clampSurfaceZoom(value / 100);
  const restore = pendingAnchor?.requestedZoom === normalized ? pendingAnchor : undefined;
  pendingAnchor = undefined;
  pendingZoom = undefined;
  if (!restore) {
    gestureToken += 1;
    return;
  }
  void nextTick().then(() => {
    const apply = () => {
      if (restore.token === gestureToken) gridRef.value?.restoreZoomAnchor(restore.anchor);
    };
    if (typeof requestAnimationFrame === "function") requestAnimationFrame(apply);
    else apply();
  });
});

watch(() => props.controller.activeTabIndex, () => {
  gestureToken += 1;
  pendingAnchor = undefined;
  pendingZoom = undefined;
});

onMounted(() => {
  const element = bodyRef.value;
  if (!element) return;
  element.addEventListener("pointerdown", onTouchPointerDown, { passive: true });
  element.addEventListener("pointermove", onTouchPointerMove, { passive: true });
  element.addEventListener("pointerup", onTouchPointerUp);
  element.addEventListener("pointercancel", onTouchPointerCancel);
  element.addEventListener("wheel", onWheel, { passive: false });
});

onBeforeUnmount(() => {
  printPageStyle?.remove?.();
  printPageStyle = undefined;
  gestureToken += 1;
  touchGestures.reset();
  const element = bodyRef.value;
  element?.removeEventListener("pointerdown", onTouchPointerDown);
  element?.removeEventListener("pointermove", onTouchPointerMove);
  element?.removeEventListener("pointerup", onTouchPointerUp);
  element?.removeEventListener("pointercancel", onTouchPointerCancel);
  element?.removeEventListener("wheel", onWheel);
});

watch(surfaceSearch.searchState, (next) => {
  emit("searchStateChange", next);
}, { immediate: true });

defineExpose({
  search: surfaceSearch.search,
  activateSearchMatch: surfaceSearch.activateSearchMatch,
  searchNext: surfaceSearch.searchNext,
  searchPrevious: surfaceSearch.searchPrevious,
  clearSearch: surfaceSearch.clearSearch,
  getSearchState: surfaceSearch.getSearchState,
  scrollToCell: async (cell: XlsxCellAddress) => {
    await nextTick();
    gridRef.value?.scrollToCell(cell);
  },
});
</script>

<style scoped>
.xlsx-viewer {
  container: xlsx-viewer / inline-size;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  font-size: 12px;
}

.xlsx-viewer__body {
  display: flex;
  flex: 1;
  min-height: 0;
  min-width: 0;
  position: relative;
  touch-action: pan-x pan-y;
}

.xlsx-viewer__loading,
.xlsx-viewer__error,
.xlsx-viewer__empty {
  align-items: center;
  display: flex;
  flex: 1;
  justify-content: center;
}

.xlsx-viewer__calculation-status {
  background: rgba(245, 158, 11, 0.12);
  border-block-end: 1px solid rgba(245, 158, 11, 0.35);
  color: inherit;
  font-size: 12px;
  line-height: 1.4;
  padding: 6px 12px;
}

.xlsx-viewer__loading-text,
.xlsx-viewer__error-text,
.xlsx-viewer__empty-text {
  opacity: 0.6;
}

.xlsx-viewer__deferred {
  align-items: center;
  display: flex;
  flex: 1;
  justify-content: center;
  min-height: 0;
  min-width: 0;
  padding: 24px;
}

.xlsx-viewer__deferred-card {
  background: #ffffff;
  border: 1px solid #d4d4d8;
  border-radius: 12px;
  box-shadow: 0 12px 32px rgba(0, 0, 0, 0.08);
  display: flex;
  flex-direction: column;
  gap: 10px;
  max-width: 420px;
  padding: 20px;
  text-align: center;
  width: 100%;
}

.xlsx-viewer__deferred-card--dark {
  background: #27272a;
  border-color: #3f3f46;
}

.xlsx-viewer__deferred-title {
  font-size: 14px;
  font-weight: 600;
}

.xlsx-viewer__deferred-file {
  font-weight: 600;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.xlsx-viewer__deferred-size {
  color: #71717a;
  line-height: 1.5;
}

.xlsx-viewer__deferred-card--dark .xlsx-viewer__deferred-size {
  color: #a1a1aa;
}

.xlsx-viewer__deferred-actions {
  display: flex;
  gap: 8px;
  justify-content: center;
  margin-top: 4px;
}

.xlsx-viewer__deferred-button {
  background: transparent;
  border: 1px solid #d4d4d8;
  border-radius: 7px;
  color: inherit;
  cursor: pointer;
  font: inherit;
  font-size: 12px;
  padding: 7px 16px;
}

.xlsx-viewer__deferred-card--dark .xlsx-viewer__deferred-button {
  border-color: #52525b;
}

.xlsx-viewer__deferred-button:hover:not(:disabled) {
  background: rgba(113, 113, 122, 0.12);
}

.xlsx-viewer__deferred-button--primary {
  background: #2563eb;
  border-color: #2563eb;
  color: #ffffff;
}

.xlsx-viewer__deferred-button--primary:hover:not(:disabled) {
  background: #1d4ed8;
}

.xlsx-viewer__deferred-button:disabled {
  cursor: default;
  opacity: 0.5;
}

.xlsx-viewer__deferred-button:focus-visible {
  outline: 2px solid #2563eb;
  outline-offset: 2px;
}

@media print {
  .xlsx-viewer {
    border-radius: 0 !important;
    height: auto !important;
    min-height: 0 !important;
    overflow: visible !important;
  }
  .xlsx-toolbar,
  .xlsx-ribbon,
  .xlsx-formula-bar,
  .xlsx-sheet-tabs,
  .xlsx-context-menu,
  .xlsx-grid__edit-input,
  .xlsx-grid__comment-popover,
  .xlsx-viewer__deferred,
  .xlsx-viewer__calculation-status,
  .xlsx-live-status { display: none !important; }
  .xlsx-viewer__body { display: block !important; overflow: visible !important; }
  .xlsx-viewer--print-preview .xlsx-viewer__body > :not(.xlsx-print-preview) { display: none !important; }
  .xlsx-viewer--print-preview { background: #ffffff !important; }
}
</style>
