<template>
  <div class="page" data-testid="pptx-surface-page">
    <header class="page-header">
      <div>
        <h2>PPTX Surface — 最小嵌入组件</h2>
        <p class="desc">
          垂直连续渲染全部幻灯片，无 Toolbar/Sidebar/搜索栏/播放控制。
          宿主通过 usePptxDocument 自行管理跳页与缩放；单页左右切换只用于播放模式。
        </p>
      </div>
    </header>

    <div class="toolbar-zone">
      <label class="ctrl">
        <input ref="fileInputRef" data-testid="pptx-surface-file-input" type="file" accept=".pptx" @change="onFileChange" />
        选择 PPTX
      </label>

      <span class="sep" />

      <div class="ctrl">
        <button :disabled="state !== 'ready'" @click="void documentApi.previousSlide()">◀</button>
        <span class="page-value">{{ documentApi.document.value ? `${activeIndex + 1} / ${documentApi.document.value.slides.length}` : "—" }}</span>
        <button :disabled="state !== 'ready'" @click="void documentApi.nextSlide()">▶</button>
      </div>

      <span class="sep" />

      <div class="ctrl">
        <button :disabled="state !== 'ready'" @click="void documentApi.setZoom(Math.max(25, zoomPercent - 25))">−</button>
        <span class="zoom-value">{{ zoomPercent }}%</span>
        <button :disabled="state !== 'ready'" @click="void documentApi.setZoom(Math.min(400, zoomPercent + 25))">+</button>
      </div>

      <span class="sep" />

      <span class="ctrl" data-testid="pptx-surface-status">{{ statusLabel }}</span>
    </div>

    <div class="stage-container" :class="{ 'stage-container--dark': true }">
      <PptxStage
        ref="stage"
        class="pptx-surface-stage"
        data-testid="pptx-surface-stage"
        @context-menu="onContextMenu"
        @object-click="onObjectClick"
        @selection-change="onSelectionChange"
      />
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, useTemplateRef, watch } from "vue"
import {
  PptxStage,
  usePptxDocument,
  type PptxStageContextMenu,
  type PptxStageExpose,
  type PptxStageObjectClick,
  type PptxStageSelection,
} from "@arcships/vue-pptx"

// ── State ────────────────────────────────────────────────────────────
const source = ref<File | null>(null)
const fileInputRef = ref<HTMLInputElement>()

const stage = useTemplateRef<PptxStageExpose>("stage")
const stageElement = computed(() => stage.value?.element ?? null)
const documentApi = usePptxDocument(stageElement, {
  source,
  session: {
    renderMode: "list",
    listOptions: {
      windowed: true,
      initialSlides: 4,
      overscanViewport: 1.5,
    },
  },
})

const { activeIndex, zoomPercent, state } = documentApi

// ── Status ───────────────────────────────────────────────────────────
const statusLabel = ref("选择一个 PPTX 文件。")
watch(state, (next) => {
  if (next === "waiting-for-stage") statusLabel.value = "等待舞台挂载…"
  else if (next === "loading") statusLabel.value = "加载中…"
  else if (next === "ready") statusLabel.value = `已加载 ${documentApi.document.value?.slides.length ?? 0} 页`
  else if (next === "error") statusLabel.value = documentApi.error.value?.message ?? "加载失败"
})

// ── Actions ──────────────────────────────────────────────────────────
function onFileChange(event: Event): void {
  source.value = (event.target as HTMLInputElement).files?.[0] ?? null
}

function onSelectionChange(selection: PptxStageSelection): void {
  statusLabel.value = `已选中第 ${selection.slideIndex + 1} 页`
}

function onObjectClick(object: PptxStageObjectClick): void {
  statusLabel.value = `第 ${object.slideIndex + 1} 页对象：${object.objectKey}`
}

function onContextMenu(ctx: PptxStageContextMenu): void {
  const target = ctx.kind === "object" ? `对象 ${ctx.objectKey}` : "幻灯片"
  statusLabel.value = `第 ${ctx.slideIndex + 1} 页${target}右键 (${ctx.clientX}, ${ctx.clientY})`
}
</script>

<style scoped>
.page { padding: 16px; max-width: 1440px; margin: 0 auto; width: 100%; min-width: 0; display: flex; flex-direction: column; height: calc(100vh - 60px); }
h2 { margin-bottom: 4px; }
.desc { color: var(--muted-foreground); margin-bottom: 0; font-size: 13px; line-height: 1.5; }

.toolbar-zone {
  display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
  padding: 10px 12px; margin: 12px 0;
  border: 1px solid var(--border); border-radius: var(--radius);
  background: var(--background); flex-shrink: 0;
}
.ctrl { display: flex; align-items: center; gap: 4px; font-size: 13px; white-space: nowrap; }
.ctrl input[type="file"] { font-size: 12px; }
.toolbar-zone button {
  padding: 4px 10px; border: 1px solid var(--border); border-radius: 4px;
  background: var(--background); cursor: pointer; font-size: 13px;
}
.toolbar-zone button:disabled { opacity: .4; cursor: default; }
.page-value, .zoom-value { min-width: 60px; text-align: center; font-variant-numeric: tabular-nums; }
.sep { width: 1px; height: 20px; background: var(--border); }

.stage-container {
  flex: 1; min-height: 0; overflow: auto;
}
.stage-container--dark { background: #18181b; }
.pptx-surface-stage { box-sizing: border-box; inline-size: min(100%, 1280px); margin-inline: auto; padding: 24px; }
</style>
