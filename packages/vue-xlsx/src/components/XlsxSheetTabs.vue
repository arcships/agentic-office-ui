<template>
  <div class="xlsx-sheettabs" :style="tabsStyle">
    <button
      class="xlsx-sheettabs__nav-btn"
      :disabled="activeTabIndex <= 0"
      @click="prevTab"
    >
      ‹
    </button>
    <div class="xlsx-sheettabs__list">
      <div
        v-for="(tab, index) in controller.tabs"
        :key="tab.id ?? index"
        class="xlsx-sheettabs__tab-wrapper"
        @mouseenter="onTabHover(index, $event)"
        @mouseleave="onTabLeave"
      >
        <button
          class="xlsx-sheettabs__tab"
          data-testid="xlsx-sheet-tab"
          :data-sheet-name="tab.name"
          role="tab"
          :aria-selected="index === activeTabIndex"
          :class="{
            'xlsx-sheettabs__tab--active': index === activeTabIndex,
          }"
          :style="tabStyle(index)"
          @click="onTabClick(index)"
          @dblclick="onTabDblClick(tab, index)"
        >
          <span class="xlsx-sheettabs__tab-name">{{ tab.name }}</span>
        </button>
        <div
          v-if="hoveredTabIndex === index && thumbnailComputed[index]"
          class="xlsx-sheettabs__thumbnail"
          :style="thumbnailPopupStyle"
        >
          <canvas
            :ref="(el) => setThumbnailCanvas(index, el)"
            :height="thumbnailComputed[index]?.height"
            :width="thumbnailComputed[index]?.width"
          />
        </div>
      </div>
      <button
        v-if="!controller.readOnly"
        class="xlsx-sheettabs__add-btn"
        title="添加工作表"
        @click="controller.addSheet()"
      >
        +
      </button>
    </div>
    <button
      class="xlsx-sheettabs__nav-btn"
      :disabled="activeTabIndex >= controller.tabs.length - 1"
      @click="nextTab"
    >
      ›
    </button>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, watch, nextTick, onMounted, onUnmounted, type CSSProperties } from "vue";
import type { XlsxViewerController, XlsxSheetThumbnail, UseXlsxViewerThumbnailsOptions } from "@extend-ai/xlsx-core";
import { useXlsxViewerThumbnails } from "../composables/useXlsxViewerThumbnails";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
  thumbnailOptions?: UseXlsxViewerThumbnailsOptions;
}>();

const activeTabIndex = computed(() => props.controller.activeTabIndex);

const { thumbnails: thumbnailComputedRaw } = useXlsxViewerThumbnails(
  () => props.controller,
  props.thumbnailOptions ?? { includeHeaders: true, resolution: { maxHeight: 132, maxWidth: 200 } },
);

const thumbnailComputed = computed(() => (thumbnailComputedRaw as unknown as XlsxSheetThumbnail[]));

const hoveredTabIndex = ref<number | null>(null);
const thumbnailCanvasRefs = ref<Map<number, HTMLCanvasElement | null>>(new Map());
const popupX = ref(0);
const popupY = ref(0);
let popupTimer: ReturnType<typeof setTimeout> | null = null;

const thumbnailPopupStyle = computed<CSSProperties>(() => ({
  left: `${popupX.value}px`,
  position: "fixed",
  top: `${popupY.value}px`,
  zIndex: 9999,
  transform: "translate(-50%, calc(-100% - 8px))",
  backgroundColor: props.isDark ? "#27272a" : "#fff",
  border: `1px solid ${props.isDark ? "#3f3f46" : "#e4e4e7"}`,
  borderRadius: "8px",
  boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
  padding: "8px",
  pointerEvents: "none",
}));

function setThumbnailCanvas(index: number, el: unknown) {
  const canvas = el as HTMLCanvasElement | null;
  thumbnailCanvasRefs.value.set(index, canvas);
}

function onTabHover(index: number, event: MouseEvent) {
  if (popupTimer) clearTimeout(popupTimer);
  const rect = (event.currentTarget as HTMLElement).getBoundingClientRect();
  popupX.value = rect.left + rect.width / 2;
  popupY.value = rect.top;
  hoveredTabIndex.value = index;

  nextTick(() => {
    const canvas = thumbnailCanvasRefs.value.get(index);
    const thumb = thumbnailComputed.value[index];
    if (canvas && thumb) {
      thumb.paint(canvas as HTMLCanvasElement);
    }
  });
}

function onTabLeave() {
  popupTimer = setTimeout(() => {
    hoveredTabIndex.value = null;
  }, 150);
}

function onTabClick(index: number) {
  props.controller.setActiveTabIndex(index);
}

function onTabDblClick(_tab: unknown, _index: number) {
  // Future: rename sheet
}

function prevTab() {
  if (activeTabIndex.value > 0) {
    props.controller.setActiveTabIndex(activeTabIndex.value - 1);
  }
}

function nextTab() {
  if (activeTabIndex.value < props.controller.tabs.length - 1) {
    props.controller.setActiveTabIndex(activeTabIndex.value + 1);
  }
}

const tabsStyle = computed<CSSProperties>(() => ({
  alignItems: "center",
  backgroundColor: props.isDark ? "#27272a" : "#f4f4f5",
  borderTop: `1px solid ${props.isDark ? "#3f3f46" : "#e4e4e7"}`,
  display: "flex",
  flexShrink: "0",
  gap: "2px",
  overflow: "hidden",
  padding: "0 4px",
}));

function tabStyle(index: number): CSSProperties {
  const isActive = index === activeTabIndex.value;
  return {
    backgroundColor: isActive
      ? props.isDark ? "#3f3f46" : "#ffffff"
      : "transparent",
    border: isActive
      ? `1px solid ${props.isDark ? "#52525b" : "#d4d4d8"}`
      : "1px solid transparent",
    borderBottom: "none",
    borderRadius: "4px 4px 0 0",
    color: isActive
      ? "inherit"
      : props.isDark ? "#a1a1aa" : "#71717a",
    cursor: "pointer",
    fontSize: "12px",
    maxWidth: "120px",
    padding: "4px 10px",
    whiteSpace: "nowrap",
  };
}

onMounted(() => {
  // Preload thumbnails
  nextTick(() => {
    thumbnailComputed.value.forEach((thumb, index) => {
      const canvas = document.createElement("canvas");
      thumb.paint(canvas);
      thumbnailCanvasRefs.value.set(index, canvas);
    });
  });
});

onUnmounted(() => {
  if (popupTimer) clearTimeout(popupTimer);
});
</script>

<style scoped>
.xlsx-sheettabs {
  user-select: none;
}

.xlsx-sheettabs__list {
  display: flex;
  flex: 1;
  gap: 2px;
  overflow-x: auto;
}

.xlsx-sheettabs__tab-wrapper {
  position: relative;
  flex-shrink: 0;
}

.xlsx-sheettabs__tab {
  align-items: center;
  background: transparent;
  border: none;
  display: flex;
  font-family: inherit;
  gap: 4px;
  outline: none;
  overflow: hidden;
}

.xlsx-sheettabs__tab:hover {
  background: rgba(128, 128, 128, 0.1);
}

.xlsx-sheettabs__tab-name {
  overflow: hidden;
  text-overflow: ellipsis;
}

.xlsx-sheettabs__thumbnail {
  display: flex;
  flex-direction: column;
  gap: 4px;
  max-width: 212px;
}

.xlsx-sheettabs__thumbnail canvas {
  border-radius: 6px;
  display: block;
  height: auto;
  width: 100%;
}

.xlsx-sheettabs__add-btn {
  align-items: center;
  background: transparent;
  border: 1px solid transparent;
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  display: flex;
  font-family: inherit;
  font-size: 16px;
  height: 28px;
  justify-content: center;
  width: 28px;
}

.xlsx-sheettabs__add-btn:hover {
  background: rgba(128, 128, 128, 0.15);
}

.xlsx-sheettabs__nav-btn {
  align-items: center;
  background: transparent;
  border: none;
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  display: flex;
  flex-shrink: 0;
  font-size: 16px;
  height: 28px;
  justify-content: center;
  width: 24px;
}

.xlsx-sheettabs__nav-btn:hover:not(:disabled) {
  background: rgba(128, 128, 128, 0.15);
}

.xlsx-sheettabs__nav-btn:disabled {
  cursor: default;
  opacity: 0.3;
}
</style>
