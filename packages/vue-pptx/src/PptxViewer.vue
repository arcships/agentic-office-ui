<template>
  <section
    ref="viewerRef"
    class="pptx-viewer"
    :class="{
      'pptx-viewer--dark': isDark,
      'pptx-viewer--present': isPresent,
    }"
    :style="{ height, maxHeight: height }"
    :data-state="state"
    :data-playback-status="playbackSnapshot?.status"
    :data-slide-index="activeIndex"
    :data-click-boundary="playbackSnapshot?.clickBoundary"
    data-testid="pptx-viewer"
    tabindex="0"
    @keydown="onKeydown"
  >
    <header v-if="showToolbar && !isPresent" class="pptx-viewer__toolbar">
      <div class="pptx-viewer__toolbar-group">
        <button type="button" :disabled="!canGoPrevious" aria-label="上一页" @click="void goTo(activeIndex - 1)">‹</button>
        <span class="pptx-viewer__counter" data-testid="pptx-page-counter">
          {{ document ? `${activeIndex + 1} / ${document.slides.length}` : "0 / 0" }}
        </span>
        <button type="button" :disabled="!canGoNext" aria-label="下一页" @click="void goTo(activeIndex + 1)">›</button>
        <span v-if="activeSlide?.hidden" class="pptx-viewer__hidden-badge">隐藏页</span>
      </div>

      <form v-if="showSearch" class="pptx-viewer__search" role="search" @submit.prevent="runSearch">
        <input
          ref="searchInputRef"
          v-model="searchQuery"
          type="search"
          placeholder="搜索幻灯片"
          aria-label="搜索幻灯片"
          @keydown.esc.prevent="closeSearch"
        />
        <button type="submit" :disabled="state !== 'ready' || !searchQuery.trim()">搜索</button>
        <span v-if="searchResults.length" class="pptx-viewer__search-count">
          {{ searchCursor + 1 }} / {{ searchResults.length }}
        </span>
        <button v-if="searchResults.length" type="button" aria-label="上一个搜索结果" @click="void moveSearch(-1)">↑</button>
        <button v-if="searchResults.length" type="button" aria-label="下一个搜索结果" @click="void moveSearch(1)">↓</button>
      </form>

      <div class="pptx-viewer__toolbar-group pptx-viewer__zoom">
        <button type="button" :disabled="state !== 'ready' || zoom <= 50" aria-label="缩小" @click="void setZoom(zoom - 25)">−</button>
        <button type="button" :disabled="state !== 'ready'" aria-label="恢复适合窗口" @click="void setZoom(100)">{{ zoom }}%</button>
        <button type="button" :disabled="state !== 'ready' || zoom >= 200" aria-label="放大" @click="void setZoom(zoom + 25)">＋</button>
      </div>
    </header>

    <div class="pptx-viewer__body">
      <aside v-if="showSidebar && !isPresent && document" class="pptx-viewer__sidebar" aria-label="幻灯片列表">
        <PptxThumbnail
          v-for="slide in document.slides"
          :key="`${loadGeneration}:${slide.index}`"
          :session="requireSession()"
          :slide="slide"
          :active="slide.index === activeIndex"
          :width="thumbnailWidth"
          @select="void goTo($event)"
        />
      </aside>

      <main ref="stageScrollRef" class="pptx-viewer__stage-wrap">
        <div v-if="state === 'empty'" class="pptx-viewer__message" data-testid="pptx-empty">
          <slot name="empty">打开一个 PPTX 文件开始查看</slot>
        </div>
        <div v-else-if="state === 'loading'" class="pptx-viewer__message" data-testid="pptx-loading">
          <slot name="loading">正在加载 PPTX…</slot>
        </div>
        <div v-else-if="state === 'error'" class="pptx-viewer__message pptx-viewer__message--error" data-testid="pptx-error">
          <slot name="error" :error="error">{{ error?.message }}</slot>
        </div>
        <PptxStage
          ref="stageRef"
          class="pptx-viewer__stage"
          :aria-hidden="state !== 'ready'"
          :zoom="zoom / 100"
          :enable-gesture-zoom="!isPresent"
          :scroll-container="stageScrollRef"
          @update:zoom="void setZoom($event * 100)"
          @click="onStageClick"
        />
      </main>
    </div>

    <footer v-if="isPresent && showPlaybackControls" class="pptx-viewer__playback-controls">
      <div class="pptx-viewer__playback-group">
        <button type="button" :disabled="state !== 'ready'" aria-label="上一步" @click="void invoke(previous)">上一步</button>
        <button type="button" :disabled="state !== 'ready'" :aria-label="playbackToggleLabel" @click="togglePlayback">
          {{ playbackToggleLabel }}
        </button>
        <button type="button" :disabled="state !== 'ready'" aria-label="下一步" @click="void invoke(next)">下一步</button>
        <button type="button" :disabled="state !== 'ready'" aria-label="重新播放" @click="void invoke(reset)">重新播放</button>
      </div>
      <span class="pptx-viewer__playback-status">
        {{ document ? `${activeIndex + 1} / ${document.slides.length}` : "0 / 0" }}
        · {{ playbackStatusText }}
      </span>
      <span v-if="showCapabilityStatus && capability" class="pptx-viewer__capability">
        精确 {{ capability.strict }}/{{ capability.discovered }}
        <template v-if="capability.approximate"> · 近似 {{ capability.approximate }}</template>
        <template v-if="capability.static"> · 静态 {{ capability.static }}</template>
        <template v-if="capability.unparsed"> · 未解析 {{ capability.unparsed }}</template>
      </span>
      <div class="pptx-viewer__playback-group">
        <button
          v-if="playbackSnapshot?.blockedMediaIds.length"
          type="button"
          aria-label="继续播放媒体"
          @click="void invoke(() => resumeBlockedMedia())"
        >继续媒体</button>
        <button type="button" aria-label="全屏" @click="void invoke(enterFullscreen)">全屏</button>
      </div>
    </footer>
  </section>
</template>

<script setup lang="ts">
import { computed, ref, watch } from "vue"
import {
  PptxPlaybackError,
  PptxPreviewError,
  type PptxCapabilityReport,
  type PptxPlaybackWarning,
  type PptxPreviewDocument,
  type PptxPreviewLimits,
} from "@arcships/pptx-core"
import {
  createPptxDocumentSession,
  createPptxPreviewSession,
  type PptxActionRequest,
  type PptxApproximationPolicy,
  type PptxDocumentSessionFactory,
  type PptxPlaybackController,
  type PptxPlaybackEvent,
  type PptxPlaybackSnapshot,
  type PptxPreviewSession,
  type PptxPreviewSessionFactory,
  type PptxPreviewSource,
} from "@arcships/pptx-core/browser"
import type { PptxSearchState, UsePptxDocumentOptions, PptxStageExpose } from "./headless-types"
import { usePptxDocument } from "./composables/usePptxDocument"
import { usePptxPlayback } from "./composables/usePptxPlayback"
import PptxStage from "./PptxStage.vue"
import PptxThumbnail from "./PptxThumbnail.vue"

const props = withDefaults(defineProps<{
  source?: PptxPreviewSource | null
  mode?: "browse" | "present"
  height?: string
  initialSlide?: number
  isDark?: boolean
  limits?: PptxPreviewLimits
  autoplay?: boolean
  approximation?: PptxApproximationPolicy
  documentSessionFactory?: PptxDocumentSessionFactory
  externalMedia?: "disabled" | "allowed"
  sessionFactory?: PptxPreviewSessionFactory
  showCapabilityStatus?: boolean
  showHiddenSlides?: boolean
  showPlaybackControls?: boolean
  showSearch?: boolean
  showSidebar?: boolean
  showToolbar?: boolean
  thumbnailWidth?: number
}>(), {
  source: null,
  mode: "browse",
  height: "100%",
  initialSlide: 0,
  isDark: false,
  limits: undefined,
  autoplay: true,
  approximation: "off",
  documentSessionFactory: undefined,
  externalMedia: "disabled",
  sessionFactory: undefined,
  showCapabilityStatus: true,
  showHiddenSlides: false,
  showPlaybackControls: true,
  showSearch: true,
  showSidebar: true,
  showToolbar: true,
  thumbnailWidth: 148,
})

const emit = defineEmits<{
  loadStart: []
  loadSuccess: [document: PptxPreviewDocument]
  loadError: [error: PptxPreviewError]
  slideChange: [index: number]
  playbackReady: [controller: PptxPlaybackController]
  playbackStateChange: [snapshot: PptxPlaybackSnapshot]
  stepChange: [slideIndex: number, boundary: number]
  playbackWarning: [warning: PptxPlaybackWarning]
  capability: [report: PptxCapabilityReport]
  mediaRequest: [mediaId: string]
  action: [action: PptxActionRequest]
  playbackError: [error: PptxPlaybackError]
  searchStateChange: [state: PptxSearchState]
}>()

type ViewerState = "empty" | "loading" | "ready" | "error"

const viewerRef = ref<HTMLElement | null>(null)
const searchInputRef = ref<HTMLInputElement | null>(null)
const stageScrollRef = ref<HTMLElement | null>(null)
const stageRef = ref<PptxStageExpose | null>(null)
const stageElement = computed(() => stageRef.value?.element ?? null)
const isPresent = computed(() => props.mode === "present")

const documentOptions = {
  source: () => props.source,
  initialSlide: () => props.initialSlide,
  session: () => ({
    fitMode: "contain" as const,
    zoomPercent: 100,
    scrollContainer: stageScrollRef.value ?? undefined,
    renderMode: props.mode === "present" ? "slide" as const : "list" as const,
    listOptions: props.mode === "present" ? undefined : {
      windowed: true,
      initialSlides: 4,
      overscanViewport: 1.5,
    },
    limits: props.limits,
    lazyMedia: true,
    lazySlides: true,
    approximation: props.approximation,
    externalMedia: props.externalMedia,
  }),
  get factory() {
    if (props.mode === "present") {
      return props.documentSessionFactory ?? createPptxDocumentSession
    }
    return props.sessionFactory ?? createPptxPreviewSession
  },
} as UsePptxDocumentOptions

const pptxDocument = usePptxDocument(stageElement, documentOptions)
const {
  activeIndex,
  document,
  error,
  zoomPercent: zoom,
} = pptxDocument

const state = computed<ViewerState>(() => {
  if (pptxDocument.state.value === "error") return "error"
  if (pptxDocument.state.value === "ready") return "ready"
  if (pptxDocument.state.value === "idle" || pptxDocument.state.value === "disposed") return "empty"
  return "loading"
})

const playback = usePptxPlayback(pptxDocument, {
  enabled: isPresent,
  autoplay: props.autoplay,
  skipHiddenSlides: !props.showHiddenSlides,
  approximation: props.approximation,
  onEvent: onPlaybackEvent,
})
const { capability, controller, snapshot: playbackSnapshot } = playback

const searchQuery = ref("")
const searchResults = computed(() => pptxDocument.searchState.value.matches)
const searchCursor = computed(() => pptxDocument.searchState.value.activeIndex)
const loadGeneration = ref(0)

const activeSlide = computed(() => document.value?.slides[activeIndex.value])
const canGoPrevious = computed(() => state.value === "ready" && activeIndex.value > 0)
const canGoNext = computed(() =>
  state.value === "ready"
  && document.value !== null
  && activeIndex.value < document.value.slides.length - 1,
)

const playbackToggleLabel = computed(() => playbackSnapshot.value?.status === "running" ? "暂停" : "播放")
const playbackStatusText = computed(() => {
  const labels: Record<string, string> = {
    idle: "尚未开始",
    ready: "已就绪",
    transitioning: "正在切换",
    running: "正在播放",
    waiting: "等待下一步",
    paused: "已暂停",
    blocked: "等待媒体授权",
    ended: "播放结束",
    disposed: "已关闭",
  }
  return labels[playbackSnapshot.value?.status ?? "idle"] ?? "尚未开始"
})

function requireController(): PptxPlaybackController {
  if (!controller.value) throw new PptxPlaybackError("PLAYBACK_NOT_READY", "PPTX 播放控制器尚未准备好。")
  return controller.value
}

function toPlaybackError(reason: unknown): PptxPlaybackError {
  return reason instanceof PptxPlaybackError
    ? reason
    : new PptxPlaybackError(
      "UNSUPPORTED_FEATURE",
      reason instanceof Error ? reason.message : "PPTX 播放操作失败。",
      { cause: reason },
    )
}

async function invoke(command: () => void | Promise<void>): Promise<void> {
  try {
    await command()
  } catch (reason) {
    emit("playbackError", toPlaybackError(reason))
  }
}

function onPlaybackEvent(event: PptxPlaybackEvent): void {
  if (event.type === "statechange") {
    emit("playbackStateChange", event.snapshot)
  } else if (event.type === "stepchange") {
    emit("stepChange", event.slideIndex, event.boundary)
  } else if (event.type === "warning") {
    emit("playbackWarning", event.warning)
  } else if (event.type === "capability") {
    emit("capability", event.report)
  } else if (event.type === "mediarequest") {
    emit("mediaRequest", event.mediaId)
  } else if (event.type === "action") {
    emit("action", event.action)
  } else if (event.type === "error") {
    emit("playbackError", event.error)
  }
}

function requireSession(): PptxPreviewSession {
  const session = pptxDocument.getSession()
  if (!session) throw new PptxPreviewError("RENDER_FAILED", "PPTX 预览会话尚未准备好。")
  return session
}

function resetSearch(): void {
  pptxDocument.clearSearch()
}

function clearSearch(): void {
  searchQuery.value = ""
  resetSearch()
}

function closeSearch(): void {
  clearSearch()
  viewerRef.value?.focus()
}

async function goTo(index: number): Promise<void> {
  if (isPresent.value) {
    await playback.goToSlide(index, { includeHidden: props.showHiddenSlides })
  } else {
    pptxDocument.getSession()?.clearSearchHighlights()
    await pptxDocument.goTo(index)
  }
}

async function next(): Promise<void> {
  if (isPresent.value) await playback.next()
  else await pptxDocument.nextSlide()
}

async function previous(): Promise<void> {
  if (isPresent.value) await playback.previous()
  else await pptxDocument.previousSlide()
}

async function play(): Promise<void> {
  await playback.play()
}

function pause(): void {
  playback.pause()
}

async function resume(): Promise<void> {
  await playback.resume()
}

async function reset(): Promise<void> {
  await playback.reset()
}

async function goToSlide(index: number): Promise<void> {
  await goTo(index)
}

async function resumeBlockedMedia(mediaId?: string): Promise<void> {
  await playback.resumeBlockedMedia(mediaId)
}

function togglePlayback(): void {
  if (playbackSnapshot.value?.status === "running") {
    void invoke(() => pause())
  } else if (playbackSnapshot.value?.status === "paused") {
    void invoke(resume)
  } else {
    void invoke(play)
  }
}

async function enterFullscreen(): Promise<void> {
  const element = viewerRef.value
  if (!element?.requestFullscreen) {
    throw new PptxPlaybackError("FULLSCREEN_REJECTED", "当前环境不支持全屏播放。")
  }
  try {
    await element.requestFullscreen()
  } catch (reason) {
    throw new PptxPlaybackError("FULLSCREEN_REJECTED", "浏览器拒绝进入全屏。", { cause: reason })
  }
}

async function exitFullscreen(): Promise<void> {
  const browserDocument = globalThis.document
  if (!browserDocument?.fullscreenElement) return
  try {
    await browserDocument.exitFullscreen()
  } catch (reason) {
    throw new PptxPlaybackError("FULLSCREEN_REJECTED", "浏览器拒绝退出全屏。", { cause: reason })
  }
}

async function onStageClick(event: MouseEvent): Promise<void> {
  if (!isPresent.value || state.value !== "ready") return
  viewerRef.value?.focus()
  const target = event.target as Element | null
  const objectElement = typeof target?.closest === "function"
    ? target.closest<HTMLElement>("[data-pptx-object-key]")
    : null
  await invoke(async () => {
    const handled = objectElement?.dataset.pptxObjectKey
      ? await requireController().activateObject(objectElement.dataset.pptxObjectKey)
      : false
    if (!handled) await next()
  })
}

async function setZoom(value: number): Promise<void> {
  await pptxDocument.setZoom(Math.min(200, Math.max(50, value)))
}

function runSearch(): void {
  void pptxDocument.search(searchQuery.value).catch(() => undefined)
}

async function moveSearch(direction: 1 | -1): Promise<void> {
  if (!searchResults.value.length) return
  if (direction === 1) await pptxDocument.searchNext()
  else await pptxDocument.searchPrevious()
}

function onKeydown(event: KeyboardEvent): void {
  if (
    !isPresent.value
    && props.showToolbar
    && props.showSearch
    && (event.ctrlKey || event.metaKey)
    && event.key.toLocaleLowerCase() === "f"
  ) {
    event.preventDefault()
    searchInputRef.value?.focus()
    return
  }
  const target = event.target
  if (
    (typeof HTMLInputElement !== "undefined" && target instanceof HTMLInputElement)
    || (typeof HTMLTextAreaElement !== "undefined" && target instanceof HTMLTextAreaElement)
  ) return
  if (isPresent.value && event.key === "Escape") {
    event.preventDefault()
    void invoke(exitFullscreen)
  } else if (isPresent.value && event.key === "Home") {
    event.preventDefault()
    void invoke(() => goToSlide(0))
  } else if (isPresent.value && event.key === "End") {
    event.preventDefault()
    void invoke(() => goToSlide(Math.max((document.value?.slides.length ?? 1) - 1, 0)))
  } else if ((isPresent.value && event.key === "ArrowLeft") || event.key === "ArrowUp" || event.key === "PageUp") {
    event.preventDefault()
    void invoke(previous)
  } else if ((isPresent.value && event.key === "ArrowRight") || event.key === "ArrowDown" || event.key === "PageDown" || isPresent.value && event.key === " ") {
    event.preventDefault()
    void invoke(next)
  } else if (!isPresent.value && (event.ctrlKey || event.metaKey) && (event.key === "+" || event.key === "=")) {
    event.preventDefault()
    void setZoom(zoom.value + 25)
  } else if ((event.ctrlKey || event.metaKey) && event.key === "-") {
    event.preventDefault()
    void setZoom(zoom.value - 25)
  }
}

watch(() => props.initialSlide, (index) => {
  if (state.value === "ready") void goTo(index)
})

watch(() => props.mode, () => {
  resetSearch()
  if (props.source) void pptxDocument.open(props.source).catch(() => undefined)
})

watch(() => pptxDocument.state.value, (next, previous) => {
  if (next === "loading") {
    resetSearch()
    loadGeneration.value += 1
    emit("loadStart")
  } else if (next === "error" && error.value) {
    emit("loadError", error.value)
  } else if (next === "ready" && previous !== "ready" && document.value) {
    loadGeneration.value += 1
    emit("loadSuccess", document.value)
    emit("slideChange", activeIndex.value)
  }
})

watch(activeIndex, (index, previous) => {
  if (state.value === "ready" && index !== previous) emit("slideChange", index)
})

watch(controller, (nextController) => {
  if (nextController) emit("playbackReady", nextController)
}, { flush: "sync" })

watch(pptxDocument.searchState, (next) => {
  emit("searchStateChange", next)
}, { immediate: true })

defineExpose({
  getController: () => controller.value,
  next,
  previous,
  play,
  pause,
  resume,
  reset,
  goToSlide,
  search: pptxDocument.search,
  activateSearchMatch: pptxDocument.activateSearchMatch,
  searchNext: pptxDocument.searchNext,
  searchPrevious: pptxDocument.searchPrevious,
  clearSearch,
  getSearchState: pptxDocument.getSearchState,
  enterFullscreen,
  exitFullscreen,
})
</script>
