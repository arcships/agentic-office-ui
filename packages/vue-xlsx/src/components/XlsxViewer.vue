<template>
  <div
    class="xlsx-viewer"
    data-testid="xlsx-viewer"
    :data-state="controller.isLoading ? 'loading' : controller.error ? 'error' : controller.activeSheet || controller.activeTab?.kind === 'chartsheet' ? 'ready' : 'empty'"
    :style="viewerStyle"
    @keydown="onKeydown"
    tabindex="0"
  >
    <XlsxToolbar
      v-if="showDefaultToolbar"
      :controller="controller"
      :is-dark="isDark ?? false"
      :show-upload="showUpload"
      @upload="emit('upload')"
    />
    <XlsxRibbon
      v-if="showRibbon"
      :controller="controller"
      :is-dark="isDark ?? false"
      :read-only="effectiveReadOnly"
      @update:read-only="onReadOnlyChange"
    />
    <XlsxFormulaBar
      v-if="showFormulaBar"
      :controller="controller"
      :is-dark="isDark ?? false"
      :read-only="effectiveReadOnly"
    />
    <div class="xlsx-viewer__body">
      <div
        v-if="controller.isLoading"
        class="xlsx-viewer__loading"
      >
        <slot name="loading">
          <span class="xlsx-viewer__loading-text">加载中…</span>
        </slot>
      </div>
      <div
        v-else-if="controller.error"
        class="xlsx-viewer__error"
      >
        <slot
          name="error"
          :error="controller.error"
        >
          <span class="xlsx-viewer__error-text">{{ controller.error.message }}</span>
        </slot>
      </div>
      <XlsxChartsheetSurface
        v-else-if="controller.activeTab?.kind === 'chartsheet'"
        :controller="controller"
        :is-dark="isDark ?? false"
      />
      <div
        v-else-if="!controller.activeSheet"
        class="xlsx-viewer__empty"
      >
        <slot name="empty">
          <span class="xlsx-viewer__empty-text">打开一个 XLSX 文件开始查看</span>
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
          :is-dark="isDark ?? false"
          :scroll-left="gridViewport.scrollLeft"
          :scroll-top="gridViewport.scrollTop"
        />
        <XlsxImageLayer
          :controller="controller"
          :show-images="showImages ?? true"
          :scroll-left="gridViewport.scrollLeft"
          :scroll-top="gridViewport.scrollTop"
        />
        <XlsxSelectionOverlay
          :controller="controller"
          :get-cell-style="getCellStyle"
          :selection-color="selectionColor"
          :selection-fill-color="selectionFillColor"
        />
        <XlsxContextMenu
          :controller="controller"
          :target-element="gridElement"
        />
      </template>
    </div>
    <XlsxSheetTabs
      v-if="controller.tabs.length > 0"
      :controller="controller"
      :is-dark="isDark ?? false"
      :read-only="effectiveReadOnly"
    />
  </div>
</template>

<script setup lang="ts">
import { computed, ref, type ComponentPublicInstance, type CSSProperties } from "vue";
import type { XlsxViewerController, XlsxCellAddress, XlsxCellStyleContext } from "@arcships/xlsx-core";
import XlsxGrid from "./XlsxGrid.vue";
import XlsxToolbar from "./XlsxToolbar.vue";
import XlsxRibbon from "./XlsxRibbon.vue";
import XlsxFormulaBar from "./XlsxFormulaBar.vue";
import XlsxSheetTabs from "./XlsxSheetTabs.vue";
import XlsxChartOverlay from "./XlsxChartOverlay.vue";
import XlsxImageLayer from "./XlsxImageLayer.vue";
import XlsxSelectionOverlay from "./XlsxSelectionOverlay.vue";
import XlsxContextMenu from "./XlsxContextMenu.vue";
import XlsxChartsheetSurface from "./XlsxChartsheetSurface.vue";

const props = withDefaults(
  defineProps<{
    controller: XlsxViewerController;
    getCellStyle?: ((cell: XlsxCellAddress, context?: XlsxCellStyleContext) => Partial<CSSProperties> | undefined) | null;
    height?: string;
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
  }>(),
  {
    getCellStyle: null,
    height: "100%",
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
  }
);

const emit = defineEmits<{
  cellDoubleClick: [cell: XlsxCellAddress];
  upload: [];
  "update:readOnly": [value: boolean];
}>();

const effectiveReadOnly = computed(() => props.controller.readOnly || props.readOnly);

const gridRef = ref<ComponentPublicInstance | null>(null);
const gridViewport = ref({ scrollLeft: 0, scrollTop: 0 });
const gridElement = computed<HTMLElement | null>(() => {
  const element = gridRef.value?.$el;
  return typeof HTMLElement !== "undefined" && element instanceof HTMLElement ? element : null;
});

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
  minHeight: "0px",
  minWidth: "0px",
  overflow: "hidden",
  position: "relative",
  width: "100%",
  backgroundColor: props.isDark ? "#18181b" : "#ffffff",
  color: props.isDark ? "#e4e4e7" : "#18181b",
  outline: "none",
}));

function onKeydown(event: KeyboardEvent) {
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
</script>

<style scoped>
.xlsx-viewer {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  font-size: 12px;
}

.xlsx-viewer__body {
  display: flex;
  flex: 1;
  min-height: 0;
  min-width: 0;
  position: relative;
}

.xlsx-viewer__loading,
.xlsx-viewer__error,
.xlsx-viewer__empty {
  align-items: center;
  display: flex;
  flex: 1;
  justify-content: center;
}

.xlsx-viewer__loading-text,
.xlsx-viewer__error-text,
.xlsx-viewer__empty-text {
  opacity: 0.6;
}
</style>
