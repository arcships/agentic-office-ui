<template>
  <div class="pptx-headless-demo">
    <header class="pptx-headless-demo__header">
      <div>
        <h2>PowerPoint 最小组合验证</h2>
        <p>本页不使用 PptxViewer，只组合文档、播放和舞台三个公开接口。</p>
      </div>
      <label class="pptx-headless-demo__upload">
        <span>选择 PowerPoint</span>
        <input
          data-testid="pptx-headless-file"
          type="file"
          accept=".pptx,.pptm,.ppsx,.ppsm,.potx,.potm,application/vnd.openxmlformats-officedocument.presentationml.presentation"
          @change="onFileChange"
        />
      </label>
    </header>

    <div class="pptx-headless-demo__controls">
      <button type="button" :disabled="state !== 'ready'" @click="void playback.previous()">上一步</button>
      <button type="button" :disabled="state !== 'ready'" data-testid="pptx-headless-next" @click="void playback.next()">下一步</button>
      <button type="button" :disabled="state !== 'ready'" data-testid="pptx-headless-play" @click="void playback.play()">播放</button>
      <button type="button" :disabled="state !== 'ready'" data-testid="pptx-headless-pause" @click="playback.pause()">暂停</button>
      <button type="button" :disabled="state !== 'ready'" data-testid="pptx-headless-resume" @click="void playback.resume()">继续</button>
      <button type="button" :disabled="state !== 'ready'" data-testid="pptx-headless-reset" @click="void playback.reset()">重播</button>
      <button type="button" :disabled="state !== 'ready'" data-testid="pptx-headless-slide-8" @click="void playback.goToSlide(7)">跳到第 8 页</button>
      <button type="button" :disabled="state !== 'ready'" data-testid="pptx-headless-slide-6" @click="void playback.goToSlide(5)">跳到第 6 页</button>
      <button type="button" data-testid="pptx-headless-remount" @click="mounted = !mounted">
        {{ mounted ? "卸载舞台" : "重新挂载" }}
      </button>
      <button type="button" data-testid="pptx-headless-second-toggle" @click="secondMounted = !secondMounted">
        {{ secondMounted ? "关闭第二舞台" : "打开第二舞台" }}
      </button>
    </div>

    <div v-if="secondMounted" class="pptx-headless-demo__second">
      <button type="button" data-testid="pptx-headless-second-next" @click="void secondPlayback.next()">第二舞台下一步</button>
      <PptxStage
        ref="secondStage"
        class="pptx-headless-demo__stage"
        data-testid="pptx-headless-second-stage"
        :data-document-state="secondDocument.state.value"
        :data-slide-index="secondDocument.activeIndex.value"
        :data-click-boundary="secondPlayback.snapshot.value?.clickBoundary ?? 0"
      />
    </div>

    <div class="pptx-headless-demo__stage-wrap">
      <PptxStage
        v-if="mounted"
        ref="stage"
        class="pptx-headless-demo__stage"
        :data-document-state="state"
        :data-file-name="documentApi.document.value?.fileName"
        :data-slide-count="documentApi.document.value?.slides.length ?? 0"
        :data-slide-index="activeIndex"
        :data-playback-status="playbackStatus"
        :data-click-boundary="snapshot?.clickBoundary ?? 0"
        aria-label="PPTX 最小播放舞台"
        @click="onStageClick"
      />
      <p v-else>舞台已卸载。</p>
    </div>

    <p class="pptx-headless-demo__status" data-testid="pptx-headless-status" role="status">
      {{ status }}
    </p>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, useTemplateRef, watch } from "vue"
import {
  PptxStage,
  usePptxDocument,
  usePptxPlayback,
  type PptxStageExpose,
} from "@arcships/vue-pptx"

const source = ref<File | null>(null)
const mounted = ref(true)
const secondMounted = ref(false)
const stage = useTemplateRef<PptxStageExpose>("stage")
const element = computed(() => stage.value?.element ?? null)
const documentApi = usePptxDocument(element, { source })
const playback = usePptxPlayback(documentApi, {
  autoplay: false,
  onEvent(event) {
    if (event.type === "warning") status.value = `警告：${event.warning.message}`
    if (event.type === "error") status.value = `播放错误：${event.error.message}`
    if (event.type === "mediarequest") status.value = `媒体 ${event.mediaId} 等待继续`
  },
})
const secondStage = useTemplateRef<PptxStageExpose>("secondStage")
const secondElement = computed(() => secondStage.value?.element ?? null)
const secondDocument = usePptxDocument(secondElement, { source })
const secondPlayback = usePptxPlayback(secondDocument, { autoplay: false })
const { activeIndex, state } = documentApi
const { snapshot, status: playbackStatus } = playback
const status = ref("请选择一个 PowerPoint 文件。")

watch(state, (next) => {
  if (next === "waiting-for-stage") status.value = "正在等待舞台挂载。"
  if (next === "loading") status.value = "正在加载。"
  if (next === "ready") status.value = `已加载 ${documentApi.document.value?.slides.length ?? 0} 页。`
  if (next === "error") status.value = documentApi.error.value?.message ?? "加载失败。"
})

function onFileChange(event: Event): void {
  source.value = (event.target as HTMLInputElement).files?.[0] ?? null
}

async function onStageClick(event: MouseEvent): Promise<void> {
  if (state.value !== "ready") return
  const object = (event.target as Element | null)
    ?.closest<HTMLElement>("[data-pptx-object-key]")
  const handled = object?.dataset.pptxObjectKey
    ? await playback.activateObject(object.dataset.pptxObjectKey)
    : false
  if (!handled) await playback.next()
}
</script>

<style scoped>
.pptx-headless-demo { display: flex; flex: 1; flex-direction: column; gap: 12px; min-height: 0; padding: 18px 24px 24px; }
.pptx-headless-demo__header { align-items: center; display: flex; gap: 24px; justify-content: space-between; }
.pptx-headless-demo__header h2 { font-size: 20px; margin: 0 0 4px; }
.pptx-headless-demo__header p,
.pptx-headless-demo__status { color: var(--muted-foreground); font-size: 13px; margin: 0; }
.pptx-headless-demo__upload { background: var(--primary); border-radius: 8px; color: var(--primary-foreground); cursor: pointer; font-size: 13px; font-weight: 600; padding: 8px 12px; }
.pptx-headless-demo__upload input { display: none; }
.pptx-headless-demo__controls { display: flex; flex-wrap: wrap; gap: 8px; }
.pptx-headless-demo__controls button { border: 1px solid var(--border); border-radius: 7px; cursor: pointer; padding: 7px 10px; }
.pptx-headless-demo__controls button:disabled { cursor: default; opacity: 0.45; }
.pptx-headless-demo__stage-wrap { align-items: center; background: #18181b; border-radius: 10px; display: flex; flex: 1; justify-content: center; min-height: 480px; overflow: hidden; }
.pptx-headless-demo__stage { inline-size: min(100%, 1280px); max-block-size: 100%; padding: 0; }
.pptx-headless-demo__second { align-items: center; background: #27272a; display: flex; gap: 10px; min-height: 220px; padding: 10px; }
.pptx-headless-demo__second .pptx-headless-demo__stage { flex: 1; max-block-size: 240px; }
@media (max-width: 640px) {
  .pptx-headless-demo { padding: 12px; }
  .pptx-headless-demo__header { align-items: flex-start; flex-direction: column; }
}
</style>
