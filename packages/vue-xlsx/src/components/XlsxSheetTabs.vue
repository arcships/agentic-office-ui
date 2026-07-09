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
      <button
        v-for="(tab, index) in controller.tabs"
        :key="tab.id ?? index"
        class="xlsx-sheettabs__tab"
        :class="{
          'xlsx-sheettabs__tab--active': index === activeTabIndex,
        }"
        :style="tabStyle(index)"
        @click="onTabClick(index)"
        @dblclick="onTabDblClick(tab, index)"
      >
        <span class="xlsx-sheettabs__tab-name">{{ tab.name }}</span>
      </button>
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
import { computed, type CSSProperties } from "vue";
import type { XlsxViewerController } from "@extend-ai/xlsx-core";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
}>();

const activeTabIndex = computed(() => props.controller.activeTabIndex);

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
