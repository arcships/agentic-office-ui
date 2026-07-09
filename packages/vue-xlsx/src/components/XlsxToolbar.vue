<template>
  <div class="xlsx-toolbar" :style="toolbarStyle">
    <div class="xlsx-toolbar__left">
      <span class="xlsx-toolbar__filename" :title="displayFileName">
        {{ displayFileName }}
      </span>
    </div>
    <div class="xlsx-toolbar__center">
      <button
        class="xlsx-toolbar__btn"
        :disabled="!controller.canZoomOut"
        title="缩小"
        @click="controller.zoomOut()"
      >
        −
      </button>
      <select
        class="xlsx-toolbar__zoom-select"
        :value="zoomPercent"
        @change="onZoomChange"
      >
        <option
          v-for="z in zoomPresets"
          :key="z"
          :value="z"
        >
          {{ z }}%
        </option>
      </select>
      <button
        class="xlsx-toolbar__btn"
        :disabled="!controller.canZoomIn"
        title="放大"
        @click="controller.zoomIn()"
      >
        +
      </button>
      <button
        class="xlsx-toolbar__btn"
        title="重置缩放"
        @click="controller.resetZoom()"
      >
        ↺
      </button>
    </div>
    <div class="xlsx-toolbar__right">
      <button
        v-if="controller.canDownload"
        class="xlsx-toolbar__btn"
        title="下载原始文件"
        @click="controller.download()"
      >
        ⬇
      </button>
      <button
        v-if="controller.canExport"
        class="xlsx-toolbar__btn"
        title="导出 XLSX"
        @click="controller.exportXlsx()"
      >
        💾
      </button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, type CSSProperties } from "vue";
import type { XlsxViewerController } from "@extend-ai/xlsx-core";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
}>();

const zoomPresets = [50, 75, 100, 125, 150, 200];

const toolbarStyle = computed<CSSProperties>(() => ({
  alignItems: "center",
  backgroundColor: props.isDark ? "#27272a" : "#f4f4f5",
  borderBottom: `1px solid ${props.isDark ? "#3f3f46" : "#e4e4e7"}`,
  display: "flex",
  flexShrink: "0",
  gap: "8px",
  justifyContent: "space-between",
  padding: "4px 12px",
}));

const displayFileName = computed(() => props.controller.displayFileName);

const zoomPercent = computed(() => Math.round(props.controller.zoomScale));

function onZoomChange(event: Event) {
  const target = event.target as HTMLSelectElement;
  const value = Number(target.value);
  if (Number.isFinite(value)) {
    props.controller.setZoomScale(value);
  }
}
</script>

<style scoped>
.xlsx-toolbar {
  user-select: none;
}

.xlsx-toolbar__left,
.xlsx-toolbar__center,
.xlsx-toolbar__right {
  align-items: center;
  display: flex;
  gap: 6px;
}

.xlsx-toolbar__filename {
  font-size: 13px;
  font-weight: 500;
  max-width: 200px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.xlsx-toolbar__btn {
  align-items: center;
  background: transparent;
  border: 1px solid transparent;
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  display: flex;
  font-size: 14px;
  height: 28px;
  justify-content: center;
  min-width: 28px;
  padding: 0 6px;
}

.xlsx-toolbar__btn:hover:not(:disabled) {
  background: rgba(128, 128, 128, 0.15);
}

.xlsx-toolbar__btn:disabled {
  cursor: default;
  opacity: 0.3;
}

.xlsx-toolbar__zoom-select {
  background: transparent;
  border: 1px solid rgba(128, 128, 128, 0.25);
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  font-size: 12px;
  height: 28px;
  outline: none;
  padding: 0 4px;
}
</style>
