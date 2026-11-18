<template>
  <section
    ref="viewerRef"
    class="pptx-viewer"
    :class="{
      'pptx-viewer--dark': isDark,
      'pptx-viewer--present': isPresent,
    }"
    :style="rootStyle"
    :data-state="state"
    :data-playback-status="playbackSnapshot?.status"
    :data-slide-index="activeIndex"
    :data-slide-error="slideFailure?.index"
    :data-click-boundary="playbackSnapshot?.clickBoundary"
    data-testid="pptx-viewer"
    tabindex="0"
    @keydown="onKeydown"
  >
    <header v-if="showToolbar && !isPresent" class="pptx-viewer__toolbar">
      <div class="pptx-viewer__toolbar-group">
        <button type="button" :disabled="!canGoPrevious" :aria-label="m.pptx.viewer.previousSlide" @click="run(() => goTo(activeIndex - 1))">‹</button>
        <span class="pptx-viewer__counter" data-testid="pptx-page-counter">
          {{ document ? `${activeIndex + 1} / ${document.slides.length}` : "0 / 0" }}
        </span>
        <button type="button" :disabled="!canGoNext" :aria-label="m.pptx.viewer.nextSlide" @click="run(() => goTo(activeIndex + 1))">›</button>
        <span v-if="activeSlide?.hidden" class="pptx-viewer__hidden-badge">{{ m.pptx.viewer.hiddenSlide }}</span>
      </div>

      <form v-if="showSearch" class="pptx-viewer__search" role="search" @submit.prevent="runSearch">
        <input
          ref="searchInputRef"
          v-model="searchQuery"
          type="search"
          :placeholder="m.pptx.viewer.searchPlaceholder"
          :aria-label="m.pptx.viewer.searchPlaceholder"
          @keydown.esc.prevent="closeSearch"
        />
        <button type="submit" :disabled="state !== 'ready' || !searchQuery.trim()">{{ m.pptx.viewer.search }}</button>
        <span v-if="searchResults.length" class="pptx-viewer__search-count" aria-hidden="true">
          {{ searchCursor + 1 }} / {{ searchResults.length }}
        </span>
        <OfficeLiveStatus test-id="pptx-search-status" :message="searchStatusMessage" />
        <button v-if="searchResults.length" type="button" :aria-label="m.common.search.previous" @click="run(() => moveSearch(-1))">↑</button>
        <button v-if="searchResults.length" type="button" :aria-label="m.common.search.next" @click="run(() => moveSearch(1))">↓</button>
      </form>

      <div class="pptx-viewer__toolbar-group pptx-viewer__zoom">
        <button type="button" :disabled="state !== 'ready' || zoom <= 50" :aria-label="m.common.zoom.out" @click="run(() => setZoom(zoom - 25))">−</button>
        <button type="button" :disabled="state !== 'ready'" :aria-label="m.common.zoom.reset" @click="run(() => setZoom(100))">{{ zoom }}%</button>
        <button type="button" :disabled="state !== 'ready' || zoom >= 200" :aria-label="m.common.zoom.in" @click="run(() => setZoom(zoom + 25))">＋</button>
      </div>
    </header>

    <div v-if="warnings.length" class="pptx-viewer__warnings" role="status" data-testid="pptx-warnings">
      <slot name="warnings" :warnings="warnings">
        <p
          v-for="warning in warnings"
          :key="`${warning.code}:${warning.slideIndex ?? 'all'}`"
          class="pptx-viewer__warning"
        >{{ warning.message }}</p>
      </slot>
    </div>

    <div class="pptx-viewer__body">
      <OfficeLiveStatus
        test-id="pptx-live-status"
        :message="liveStatusMessage"
        :busy="state === 'loading'"
      />
      <aside
        v-if="showSidebar && !isPresent && document"
        ref="sidebarRef"
        class="pptx-viewer__sidebar"
        :aria-label="m.pptx.viewer.sidebar"
        data-testid="pptx-sidebar"
        :style="{ '--pptx-thumbnail-item-height': `${thumbnailItemHeight}px` }"
        @scroll.passive="onSidebarScroll"
      >
        <div
          v-if="sidebarWindow.start > 0"
          class="pptx-viewer__sidebar-spacer"
          :style="{ height: `${sidebarWindow.start * thumbnailItemHeight}px` }"
          aria-hidden="true"
        />
        <PptxThumbnail
          v-for="slide in visibleSlides"
          :key="`${loadGeneration}:${thumbnailFontEpoch}:${slide.index}`"
          :session="requireSession()"
          :slide="slide"
          :active="slide.index === activeIndex"
          :width="thumbnailWidth"
          @select="run(() => goTo($event))"
        />
        <div
          v-if="document.slides.length > sidebarWindow.end"
          class="pptx-viewer__sidebar-spacer"
          :style="{ height: `${(document.slides.length - sidebarWindow.end) * thumbnailItemHeight}px` }"
          aria-hidden="true"
        />
      </aside>

      <main ref="stageScrollRef" class="pptx-viewer__stage-wrap">
        <div v-if="state === 'empty'" class="pptx-viewer__message" data-testid="pptx-empty">
          <slot name="empty">{{ m.pptx.viewer.empty }}</slot>
        </div>
        <div v-else-if="state === 'waiting'" class="pptx-viewer__message" data-testid="pptx-waiting">
          <slot name="waiting">{{ m.pptx.viewer.waiting }}</slot>
        </div>
        <div v-else-if="state === 'loading'" class="pptx-viewer__message" data-testid="pptx-loading">
          <slot name="loading">{{ m.pptx.viewer.loading }}</slot>
        </div>
        <div v-else-if="state === 'cancelled'" class="pptx-viewer__message" data-testid="pptx-cancelled">
          <slot name="cancelled" :error="error">{{ error?.message ?? m.pptx.viewer.cancelled }}</slot>
          <button type="button" data-testid="pptx-cancelled-retry" @click="run(reopen)">{{ m.pptx.viewer.reload }}</button>
        </div>
        <div v-else-if="state === 'error'" class="pptx-viewer__message pptx-viewer__message--error" data-testid="pptx-error">
          <slot name="error" :error="error">{{ error?.message }}</slot>
          <button v-if="props.source" type="button" data-testid="pptx-error-retry" @click="run(reopen)">{{ m.pptx.viewer.retry }}</button>
        </div>
        <div
          v-if="slideFailure"
          class="pptx-viewer__slide-error"
          role="alert"
          data-testid="pptx-slide-error"
        >
          <slot name="slide-error" :error="slideFailure.error" :index="slideFailure.index">
            <span>{{ formatOfficeMessage(m.pptx.viewer.slideRenderFailed, { page: slideFailure.index + 1, message: slideFailure.error.message }) }}</span>
          </slot>
          <span class="pptx-viewer__slide-error-actions">
            <button type="button" data-testid="pptx-slide-retry" @click="run(retrySlide)">{{ m.pptx.viewer.retry }}</button>
            <button type="button" data-testid="pptx-slide-skip" @click="run(skipSlide)">{{ m.pptx.viewer.skip }}</button>
          </span>
        </div>
        <PptxStage
          ref="stageRef"
          class="pptx-viewer__stage"
          :style="stageStyle"
          :aria-hidden="state !== 'ready'"
          :zoom="zoom / 100"
          :enable-gesture-zoom="!isPresent"
          :locale="locale"
          :messages="messages"
          :scroll-container="stageScrollRef"
          @update:zoom="run(() => setZoom($event * 100))"
          @click="run(() => onStageClick($event))"
        />
      </main>
    </div>

    <footer v-if="isPresent && showPlaybackControls" class="pptx-viewer__playback-controls">
      <div class="pptx-viewer__playback-group">
        <button type="button" :disabled="state !== 'ready'" :aria-label="m.pptx.viewer.previousStep" @click="void invoke(previous)">{{ m.pptx.viewer.previousStep }}</button>
        <button type="button" :disabled="state !== 'ready'" :aria-label="playbackToggleLabel" @click="togglePlayback">
          {{ playbackToggleLabel }}
        </button>
        <button type="button" :disabled="state !== 'ready'" :aria-label="m.pptx.viewer.nextStep" @click="void invoke(next)">{{ m.pptx.viewer.nextStep }}</button>
        <button type="button" :disabled="state !== 'ready'" :aria-label="m.pptx.viewer.replay" @click="void invoke(reset)">{{ m.pptx.viewer.replay }}</button>
      </div>
      <span class="pptx-viewer__playback-status">
        {{ document ? `${activeIndex + 1} / ${document.slides.length}` : "0 / 0" }}
        · {{ playbackStatusText }}
      </span>
      <span v-if="showCapabilityStatus && capability" class="pptx-viewer__capability">
        {{ formatOfficeMessage(m.pptx.viewer.capability.strict, { count: capability.strict, total: capability.discovered }) }}
        <template v-if="capability.approximate"> · {{ formatOfficeMessage(m.pptx.viewer.capability.approximate, { count: capability.approximate }) }}</template>
        <template v-if="capability.static"> · {{ formatOfficeMessage(m.pptx.viewer.capability.static, { count: capability.static }) }}</template>
        <template v-if="capability.unparsed"> · {{ formatOfficeMessage(m.pptx.viewer.capability.unparsed, { count: capability.unparsed }) }}</template>
      </span>
      <div class="pptx-viewer__playback-group">
        <button
          v-if="playbackSnapshot?.blockedMediaIds.length"
          type="button"
          :aria-label="m.pptx.viewer.resumeMediaAction"
          @click="void invoke(() => resumeBlockedMedia())"
        >{{ m.pptx.viewer.resumeMedia }}</button>
        <button type="button" :aria-label="m.pptx.viewer.fullscreen" @click="void invoke(enterFullscreen)">{{ m.pptx.viewer.fullscreen }}</button>
      </div>
    </footer>
  </section>
</template>

<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, provide, ref, watch } from "vue"
import {
  PptxPlaybackError,
  PptxPreviewError,
  type PptxCapabilityReport,
  type PptxPlaybackWarning,
  type PptxPreviewDocument,
  type PptxPreviewLimits,
  type PptxPreviewWarning,
} from "@arcships/pptx-core"
import {
  createPptxDocumentSession,
  createPptxPreviewSession,
  type PptxActionRequest,
  type PptxApproximationPolicy,
  type PptxDocumentSessionFactory,
  type PptxExternalMediaUrlPolicy,
  type PptxPlaybackController,
  type PptxPlaybackEvent,
  type PptxPlaybackSnapshot,
  type PptxPreviewSession,
  type PptxPreviewSessionFactory,
  type PptxPreviewSource,
} from "@arcships/pptx-core/browser"
import type { PptxSearchState, UsePptxDocumentOptions, PptxStageExpose } from "./headless-types"
import { OfficeLiveStatus, isEditableViewerTarget, isInteractiveViewerTarget, useViewerKeys } from "@arcships/vue-ui"
import type { ViewerKeyCommand } from "@arcships/vue-ui"
import type { OfficeLocale, OfficeMessagesOverride } from "@arcships/office-interaction"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { providePptxMessages } from "./i18n"
import { usePptxDocument } from "./composables/usePptxDocument"
import { usePptxPlayback } from "./composables/usePptxPlayback"
import { watchFontLoading } from "./composables/font-loading"
import {
  PPTX_THUMBNAIL_OBSERVER,
  createPptxThumbnailObserver,
} from "./composables/thumbnail-observer"
import PptxStage from "./PptxStage.vue"
import PptxThumbnail from "./PptxThumbnail.vue"

const props = withDefaults(defineProps<{
  source?: PptxPreviewSource | null
  mode?: "browse" | "present"
  height?: string
  /** Optional minimum height; the viewer itself does not impose one. */
  minHeight?: string | number
  initialSlide?: number
  isDark?: boolean
  limits?: PptxPreviewLimits
  autoplay?: boolean
  approximation?: PptxApproximationPolicy
  documentSessionFactory?: PptxDocumentSessionFactory
  externalMedia?: "disabled" | "allowed"
  externalMediaUrlPolicy?: PptxExternalMediaUrlPolicy
  sessionFactory?: PptxPreviewSessionFactory
  openTimeoutMs?: number
  showCapabilityStatus?: boolean
  showHiddenSlides?: boolean
  showPlaybackControls?: boolean
  showSearch?: boolean
  showSidebar?: boolean
  showToolbar?: boolean
  showWarnings?: boolean
  thumbnailWidth?: number
  locale?: OfficeLocale
  messages?: OfficeMessagesOverride
}>(), {
  source: null,
  mode: "browse",
  height: "100%",
  minHeight: undefined,
  initialSlide: 0,
  isDark: false,
  limits: undefined,
  autoplay: true,
  approximation: "off",
  documentSessionFactory: undefined,
  externalMedia: "disabled",
  externalMediaUrlPolicy: undefined,
  sessionFactory: undefined,
  openTimeoutMs: 30_000,
  showCapabilityStatus: true,
  showHiddenSlides: false,
  showPlaybackControls: true,
  showSearch: true,
  showSidebar: true,
  showToolbar: true,
  showWarnings: true,
  thumbnailWidth: 148,
  locale: "en",
  messages: undefined,
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
  slideError: [error: PptxPreviewError, slideIndex: number]
  documentError: [error: PptxPreviewError]
}>()

type ViewerState = "empty" | "waiting" | "loading" | "ready" | "cancelled" | "error"

/** Padding, border, label and gap around the thumbnail canvas, kept in sync with `style.css`. */
const THUMBNAIL_CHROME_PX = 48
const THUMBNAIL_MIN_CANVAS_PX = 84
const SIDEBAR_OVERSCAN_ROWS = 3
const SIDEBAR_FALLBACK_VIEWPORT_PX = 640

const m = providePptxMessages(() => props.locale, () => props.messages)

const viewerRef = ref<HTMLElement | null>(null)
const searchInputRef = ref<HTMLInputElement | null>(null)
const stageScrollRef = ref<HTMLElement | null>(null)
const stageRef = ref<PptxStageExpose | null>(null)
const stageElement = computed(() => stageRef.value?.element ?? null)
const isPresent = computed(() => props.mode === "present")

// One laziness gate shared by every sidebar thumbnail.
const thumbnailObserver = createPptxThumbnailObserver()
provide(PPTX_THUMBNAIL_OBSERVER, thumbnailObserver)
onBeforeUnmount(() => thumbnailObserver.dispose())

// Thumbnails painted before a webfont loaded keep fallback metrics; remount them once per font batch.
const thumbnailFontEpoch = ref(0)
const stopFontLoadingWatcher = watchFontLoading(() => {
  thumbnailFontEpoch.value += 1
})
onBeforeUnmount(stopFontLoadingWatcher)

const documentOptions = {
  source: () => props.source,
  initialSlide: () => props.initialSlide,
  openTimeoutMs: () => props.openTimeoutMs,
  messages: () => m.value.pptx.documentErrors,
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
    externalMediaUrlPolicy: props.externalMediaUrlPolicy,
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
  switch (pptxDocument.state.value) {
    case "error": return "error"
    case "ready": return "ready"
    case "cancelled": return "cancelled"
    case "waiting-for-stage": return "waiting"
    case "loading": return "loading"
    default: return "empty"
  }
})

const warnings = computed<readonly PptxPreviewWarning[]>(() =>
  props.showWarnings ? document.value?.warnings ?? [] : [])
const slideFailure = computed(() => pptxDocument.slideError.value)

const isFullscreen = ref(false)
const printPageStyle = ref<HTMLStyleElement | undefined>()

function toCssLength(value: string | number | undefined): string | undefined {
  if (value === undefined) return undefined
  return typeof value === "number" ? `${value}px` : value
}

const rootStyle = computed(() => {
  if (isFullscreen.value) return undefined
  const style: Record<string, string> = { height: props.height, maxHeight: props.height }
  const minHeight = toCssLength(props.minHeight)
  if (minHeight) style.minHeight = minHeight
  return style
})

function syncPrintPageSize(width?: number, height?: number): void {
  const browserDocument = globalThis.document
  if (!browserDocument || typeof browserDocument.createElement !== "function") return
  if (!width || !height || width <= 0 || height <= 0) {
    printPageStyle.value?.remove?.()
    printPageStyle.value = undefined
    return
  }
  if (!printPageStyle.value) {
    const style = browserDocument.createElement("style")
    if (style.dataset) style.dataset.officePrintPage = "pptx"
    browserDocument.head?.append(style)
    printPageStyle.value = style
  }
  printPageStyle.value.textContent = `@page { size: ${Math.round(width)}px ${Math.round(height)}px; margin: 0; }`
}

watch([() => document.value?.width, () => document.value?.height], ([width, height]) => {
  syncPrintPageSize(width, height)
}, { immediate: true })

const stageStyle = computed(() => {
  const loaded = document.value
  if (!isPresent.value || !loaded || !loaded.width || !loaded.height) return undefined
  return {
    "--pptx-slide-aspect": `${loaded.width} / ${loaded.height}`,
    "--pptx-slide-ratio": String(Number((loaded.width / loaded.height).toFixed(4))),
  }
})

function syncFullscreen(): void {
  const active = globalThis.document?.fullscreenElement ?? null
  const element = viewerRef.value
  if (!active || !element) {
    isFullscreen.value = false
    return
  }
  isFullscreen.value = active === element
    || (typeof active.contains === "function" && active.contains(element))
}

onMounted(() => {
  globalThis.document?.addEventListener?.("fullscreenchange", syncFullscreen)
})

onBeforeUnmount(() => {
  printPageStyle.value?.remove?.()
  printPageStyle.value = undefined
  globalThis.document?.removeEventListener?.("fullscreenchange", syncFullscreen)
})

const playback = usePptxPlayback(pptxDocument, {
  enabled: isPresent,
  autoplay: props.autoplay,
  skipHiddenSlides: () => !props.showHiddenSlides,
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

// Sidebar virtualization: keep the mounted thumbnail count proportional to the
// visible rows instead of the deck length.
const sidebarRef = ref<HTMLElement | null>(null)
const sidebarScrollTop = ref(0)
const sidebarViewportHeight = ref(SIDEBAR_FALLBACK_VIEWPORT_PX)
const thumbnailItemHeight = computed(() => {
  const slideshow = document.value
  const ratio = slideshow && slideshow.width > 0 ? slideshow.height / slideshow.width : 9 / 16
  return Math.max(THUMBNAIL_MIN_CANVAS_PX, Math.round(props.thumbnailWidth * ratio)) + THUMBNAIL_CHROME_PX
})
const sidebarWindow = computed(() => {
  const count = document.value?.slides.length ?? 0
  const itemHeight = thumbnailItemHeight.value
  const firstVisible = Math.floor(Math.max(0, sidebarScrollTop.value) / itemHeight)
  const visibleRows = Math.ceil(sidebarViewportHeight.value / itemHeight) + 1
  const start = Math.max(0, Math.min(firstVisible, Math.max(count - 1, 0)) - SIDEBAR_OVERSCAN_ROWS)
  const end = Math.max(start, Math.min(count, start + visibleRows + SIDEBAR_OVERSCAN_ROWS * 2))
  return { start, end }
})
const visibleSlides = computed(() =>
  document.value?.slides.slice(sidebarWindow.value.start, sidebarWindow.value.end) ?? [],
)

const playbackToggleLabel = computed(() => playbackSnapshot.value?.status === "running"
  ? m.value.pptx.viewer.pause
  : m.value.pptx.viewer.play)
const playbackStatusText = computed(() => {
  const labels: Record<string, string> = {
    idle: m.value.pptx.viewer.status.idle,
    ready: m.value.pptx.viewer.status.ready,
    transitioning: m.value.pptx.viewer.status.transitioning,
    running: m.value.pptx.viewer.status.running,
    waiting: m.value.pptx.viewer.status.waiting,
    paused: m.value.pptx.viewer.status.paused,
    blocked: m.value.pptx.viewer.status.blocked,
    ended: m.value.pptx.viewer.status.ended,
    disposed: m.value.pptx.viewer.status.disposed,
  }
  return labels[playbackSnapshot.value?.status ?? "idle"] ?? m.value.pptx.viewer.status.idle
})

function requireController(): PptxPlaybackController {
  if (!controller.value) throw new PptxPlaybackError("PLAYBACK_NOT_READY", m.value.pptx.viewer.error.controllerNotReady)
  return controller.value
}

// Slide changes, playback and load failures were visible only to sighted users.
// One stable live region announces them for assistive technology.
const liveStatusMessage = computed(() => {
  if (state.value === "empty") return m.value.pptx.viewer.live.empty
  if (state.value === "loading") return m.value.pptx.viewer.live.loading
  if (state.value === "error") {
    return formatOfficeMessage(m.value.pptx.viewer.live.error, {
      message: error.value?.message ?? m.value.pptx.viewer.live.unknownError,
    })
  }
  const total = document.value?.slides.length ?? 0
  const position = formatOfficeMessage(m.value.pptx.viewer.live.ready, {
    current: activeIndex.value + 1,
    total,
  })
  const hidden = activeSlide.value?.hidden ? m.value.pptx.viewer.live.hiddenSuffix : ""
  const playbackStatus = isPresent.value
    ? formatOfficeMessage(m.value.pptx.viewer.live.playbackSuffix, { status: playbackStatusText.value })
    : ""
  return `${position}${hidden}${playbackStatus}`
})

// The search hit counter used to be a visual-only span. DOCX, XLSX and PDF all
// announce hits through a status region; the PPTX viewer now speaks the same
// common search wording so a slide search is announced the same way in every
// format.
const searchStatusMessage = computed(() => {
  const search = pptxDocument.searchState.value
  if (search.status === "searching") return m.value.common.search.searching
  if (search.status === "error") return m.value.common.search.error
  if (!search.query.trim()) return ""
  const total = search.matches.length
  if (!total) return m.value.common.search.noResults
  const current = Math.min(Math.max(search.activeIndex + 1, 1), total)
  return formatOfficeMessage(m.value.common.search.resultOf, { current, total })
})

function toPlaybackError(reason: unknown): PptxPlaybackError {
  return reason instanceof PptxPlaybackError
    ? reason
    : new PptxPlaybackError(
      "UNSUPPORTED_FEATURE",
      reason instanceof Error ? reason.message : m.value.pptx.viewer.error.operationFailed,
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

function toDocumentError(reason: unknown): PptxPreviewError {
  return reason instanceof PptxPreviewError
    ? reason
    : new PptxPreviewError(
      "RENDER_FAILED",
      reason instanceof Error ? reason.message : m.value.pptx.viewer.error.documentOperationFailed,
      reason,
    )
}

function reportDocumentError(reason: unknown): void {
  const failure = toDocumentError(reason)
  // Superseded or aborted work is expected when the host switches files or
  // modes; it is not a user-visible failure.
  if (failure.code === "ABORTED" || failure.code === "STALE_RESULT") return
  const currentSlideFailure = pptxDocument.slideError.value
  if (currentSlideFailure && currentSlideFailure.error === failure) return
  emit("documentError", failure)
}

/** Run a document command and surface its failure as an event instead of an unhandled rejection. */
function run(command: () => Promise<unknown> | void): void {
  try {
    const result: Promise<unknown> | void = command()
    if (result instanceof Promise) {
      void result.catch(reportDocumentError)
    }
  } catch (reason) {
    reportDocumentError(reason)
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
  if (!session) throw new PptxPreviewError("RENDER_FAILED", m.value.pptx.viewer.error.previewNotReady)
  return session
}

function measureSidebar(): void {
  const element = sidebarRef.value
  if (!element || typeof element.clientHeight !== "number" || element.clientHeight <= 0) return
  sidebarViewportHeight.value = element.clientHeight
}

function setSidebarScrollTop(value: number): void {
  sidebarScrollTop.value = Math.max(0, value)
  const element = sidebarRef.value
  if (element && typeof element.scrollTop === "number") element.scrollTop = sidebarScrollTop.value
}

function revealSidebarSlide(index: number): void {
  if (!sidebarRef.value) return
  const itemHeight = thumbnailItemHeight.value
  const top = Math.max(0, index) * itemHeight
  const bottom = top + itemHeight
  const current = sidebarScrollTop.value
  const viewport = sidebarViewportHeight.value
  const next = top < current ? top : bottom > current + viewport ? bottom - viewport : current
  if (next !== current) setSidebarScrollTop(next)
}

function onSidebarScroll(event: Event): void {
  const target = event.target as { scrollTop?: unknown; clientHeight?: unknown } | null
  if (typeof target?.scrollTop === "number") sidebarScrollTop.value = Math.max(0, target.scrollTop)
  if (typeof target?.clientHeight === "number" && target.clientHeight > 0) {
    sidebarViewportHeight.value = target.clientHeight
  }
}

watch(document, () => {
  setSidebarScrollTop(0)
  measureSidebar()
}, { flush: "post" })

watch(activeIndex, (index) => {
  measureSidebar()
  revealSidebarSlide(index)
}, { flush: "post" })

watch(thumbnailItemHeight, () => {
  revealSidebarSlide(activeIndex.value)
}, { flush: "post" })

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

async function reopen(): Promise<void> {
  const source = props.source
  if (!source) return
  await pptxDocument.open(source, { initialSlide: pptxDocument.activeIndex.value })
}

async function retrySlide(): Promise<void> {
  const failure = pptxDocument.slideError.value
  if (!failure) return
  await goTo(failure.index)
}

async function skipSlide(): Promise<void> {
  const failure = pptxDocument.slideError.value
  const slides = document.value?.slides
  if (!failure || !slides?.length) return
  const nextIndex = failure.index + 1 < slides.length ? failure.index + 1 : failure.index - 1
  if (nextIndex < 0 || nextIndex === failure.index) {
    pptxDocument.clearSlideError()
    return
  }
  await goTo(nextIndex)
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
    throw new PptxPlaybackError("FULLSCREEN_REJECTED", m.value.pptx.viewer.error.fullscreenUnsupported)
  }
  try {
    await element.requestFullscreen()
  } catch (reason) {
    throw new PptxPlaybackError("FULLSCREEN_REJECTED", m.value.pptx.viewer.error.fullscreenEnterRejected, { cause: reason })
  }
}

async function exitFullscreen(): Promise<void> {
  const browserDocument = globalThis.document
  if (!browserDocument?.fullscreenElement) return
  try {
    await browserDocument.exitFullscreen()
  } catch (reason) {
    throw new PptxPlaybackError("FULLSCREEN_REJECTED", m.value.pptx.viewer.error.fullscreenExitRejected, { cause: reason })
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
  run(() => pptxDocument.search(searchQuery.value))
}

async function moveSearch(direction: 1 | -1): Promise<void> {
  if (!searchResults.value.length) return
  if (direction === 1) await pptxDocument.searchNext()
  else await pptxDocument.searchPrevious()
}

function onKeydown(event: KeyboardEvent): void {
  if (handleBrowseScrollKey(event)) return
  viewerKeys.onKeydown(event)
}

const BROWSE_PAGE_SCROLL_RATIO = 0.9
const BROWSE_LINE_SCROLL_PX = 48
const BROWSE_SCROLL_KEYS = new Set([
  "PageDown",
  "PageUp",
  "Home",
  "End",
  "ArrowDown",
  "ArrowUp",
  " ",
  "Space",
  "Spacebar",
])

/**
 * PPTX keyboard map:
 * - browse: arrows/PageUp/PageDown/Home/End/Space scroll the slide list; the
 *   shared map adds Ctrl/Cmd zoom, Ctrl/Cmd+F search and Escape.
 * - present: PageUp/PageDown/Home/End and arrows/Space navigate slides; Escape
 *   leaves fullscreen; Ctrl/Cmd +/−/0 zoom.
 */
const viewerKeys = useViewerKeys({
  mode: () => (isPresent.value ? "present" : "browse"),
  pageNavigation: () => isPresent.value,
  zoom: true,
  search: () => !isPresent.value && props.showToolbar && props.showSearch,
  escape: true,
  onCommand: onViewerKeyCommand,
})

function onViewerKeyCommand(command: ViewerKeyCommand): boolean {
  switch (command) {
    case "page-next":
      void invoke(next)
      return true
    case "page-previous":
      void invoke(previous)
      return true
    case "page-first":
      void invoke(() => goToSlide(0))
      return true
    case "page-last":
      void invoke(() => goToSlide(Math.max((document.value?.slides.length ?? 1) - 1, 0)))
      return true
    case "zoom-in":
      if (zoom.value >= 200) return false
      run(() => setZoom(zoom.value + 25))
      return true
    case "zoom-out":
      if (zoom.value <= 50) return false
      run(() => setZoom(zoom.value - 25))
      return true
    case "zoom-reset":
      if (zoom.value === 100) return false
      run(() => setZoom(100))
      return true
    case "search":
      if (isPresent.value || !props.showToolbar || !props.showSearch) return false
      searchInputRef.value?.focus()
      return true
    case "escape":
      if (isPresent.value) {
        void invoke(exitFullscreen)
        return true
      }
      if (searchQuery.value || searchResults.value.length) {
        closeSearch()
        return true
      }
      return false
    default:
      return false
  }
}

/** Browse mode owns scrolling, not slide navigation. */
function handleBrowseScrollKey(event: KeyboardEvent): boolean {
  if (isPresent.value || state.value !== "ready" || event.defaultPrevented) return false
  if (event.ctrlKey || event.metaKey || event.altKey || event.shiftKey) return false
  if (!BROWSE_SCROLL_KEYS.has(event.key) || isEditableViewerTarget(event.target)) return false
  const space = event.key === " " || event.key === "Space" || event.key === "Spacebar"
  if (space && isInteractiveViewerTarget(event.target)) return false
  const container = stageScrollRef.value
  if (!container) return false
  const viewport = typeof container.clientHeight === "number" ? container.clientHeight : 0
  if (event.key !== "Home" && event.key !== "End" && event.key !== "ArrowDown" && event.key !== "ArrowUp" && viewport <= 0) {
    return false
  }
  const pageStep = Math.max(1, Math.round(viewport * BROWSE_PAGE_SCROLL_RATIO))
  switch (event.key) {
    case "PageDown":
    case " ":
    case "Space":
    case "Spacebar":
      scrollBrowseBy(pageStep)
      break
    case "PageUp":
      scrollBrowseBy(-pageStep)
      break
    case "Home":
      setBrowseScrollTop(0)
      break
    case "End":
      setBrowseScrollTop(Number.MAX_SAFE_INTEGER)
      break
    case "ArrowDown":
      scrollBrowseBy(BROWSE_LINE_SCROLL_PX)
      break
    case "ArrowUp":
      scrollBrowseBy(-BROWSE_LINE_SCROLL_PX)
      break
    default:
      return false
  }
  event.preventDefault?.()
  return true
}

function setBrowseScrollTop(top: number): void {
  const container = stageScrollRef.value
  if (!container) return
  const scrollHeight = typeof container.scrollHeight === "number" ? container.scrollHeight : 0
  const clientHeight = typeof container.clientHeight === "number" ? container.clientHeight : 0
  const max = Math.max(0, scrollHeight - clientHeight)
  container.scrollTop = Math.min(max, Math.max(0, top))
}

function scrollBrowseBy(delta: number): void {
  const container = stageScrollRef.value
  if (!container) return
  const top = typeof container.scrollTop === "number" ? container.scrollTop : 0
  setBrowseScrollTop(top + delta)
}

watch(() => props.initialSlide, (index) => {
  if (state.value === "ready") run(() => goTo(index))
})

watch(() => props.mode, () => {
  const source = props.source
  const index = pptxDocument.activeIndex.value
  resetSearch()
  if (!source) return
  // Re-opening builds a new session for the other render mode, so carry the
  // current slide across instead of dropping the viewer back to slide 1.
  run(() => pptxDocument.open(source, { initialSlide: index }))
})

watch(() => pptxDocument.slideError.value, (failure) => {
  if (failure) emit("slideError", failure.error, failure.index)
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
  refresh: pptxDocument.refresh,
  enterFullscreen,
  exitFullscreen,
})
</script>
