<template>
  <div class="pptx-demo">
    <header class="pptx-demo__header">
      <div>
        <h2>PPTX 网页预览与播放</h2>
        <p>同一份文件可以切换浏览和演示模式，演示模式支持动画、媒体、键盘与全屏控制。</p>
      </div>
      <div class="pptx-demo__actions">
        <div class="pptx-demo__mode" aria-label="PPTX 显示模式">
          <button type="button" :aria-pressed="mode === 'browse'" data-testid="pptx-mode-browse" @click="mode = 'browse'">浏览</button>
          <button type="button" :aria-pressed="mode === 'present'" data-testid="pptx-mode-present" @click="mode = 'present'">演示</button>
        </div>
        <label class="pptx-demo__upload">
          <span>选择 PPTX</span>
          <input data-testid="pptx-file-input" type="file" accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation" @change="onFileChange" />
        </label>
      </div>
    </header>

    <p v-if="fileName" class="pptx-demo__file">当前文件：{{ fileName }}</p>
    <PptxViewer
      :source="source"
      :mode="mode"
      :autoplay="false"
      height="calc(100vh - 190px)"
      @action="status = `动作：${$event.kind}`"
      @capability="status = `发现 ${$event.discovered} 项播放能力，严格执行 ${$event.strict} 项`"
      @load-error="status = $event.message"
      @load-start="status = '正在加载…'"
      @load-success="status = `已加载 ${$event.slides.length} 页`"
      @media-request="status = `媒体 ${$event} 等待用户继续`"
      @playback-state-change="status = `演示状态：${$event.status}，第 ${$event.slideIndex + 1} 页`"
      @playback-error="status = `演示错误：${$event.message}`"
    />
    <p class="pptx-demo__status" data-testid="pptx-demo-status" role="status">{{ status }}</p>
  </div>
</template>

<script setup lang="ts">
import { ref } from "vue"
import { PptxViewer } from "@arcships/vue-pptx"

const source = ref<File | null>(null)
const mode = ref<"browse" | "present">("browse")
const fileName = ref("")
const status = ref("请选择一个 PPTX 文件。")

function onFileChange(event: Event): void {
  const file = (event.target as HTMLInputElement).files?.[0] ?? null
  source.value = file
  fileName.value = file?.name ?? ""
  if (!file) status.value = "请选择一个 PPTX 文件。"
}
</script>

<style scoped>
.pptx-demo { display: flex; flex: 1; flex-direction: column; gap: 10px; min-height: 0; padding: 18px 24px 24px; }
.pptx-demo__header { align-items: center; display: flex; gap: 24px; justify-content: space-between; }
.pptx-demo__header h2 { font-size: 20px; margin: 0 0 4px; }
.pptx-demo__header p,
.pptx-demo__file,
.pptx-demo__status { color: var(--muted-foreground); font-size: 13px; margin: 0; }
.pptx-demo__upload { background: var(--primary); border-radius: 8px; color: var(--primary-foreground); cursor: pointer; flex: 0 0 auto; font-size: 13px; font-weight: 600; padding: 8px 12px; }
.pptx-demo__upload input { display: none; }
.pptx-demo__actions,
.pptx-demo__mode { align-items: center; display: flex; gap: 8px; }
.pptx-demo__mode { background: var(--muted); border-radius: 8px; padding: 3px; }
.pptx-demo__mode button { background: transparent; border: 0; border-radius: 6px; color: var(--muted-foreground); cursor: pointer; font: inherit; padding: 6px 10px; }
.pptx-demo__mode button[aria-pressed="true"] { background: var(--background); color: var(--foreground); font-weight: 600; }
.pptx-demo :deep(.pptx-viewer) { border: 1px solid var(--border); border-radius: 10px; flex: 1; }
@media (max-width: 640px) {
  .pptx-demo { padding: 12px; }
  .pptx-demo__header { align-items: flex-start; flex-direction: column; gap: 10px; }
}
</style>
