<template>
  <section
    ref="viewerRef"
    class="pptx-viewer"
    :class="{
      'pptx-viewer--dark': isDark,
      'pptx-viewer--present': isPresent,
      'pptx-viewer--presenter': presenterVisible,
    }"
    :style="rootStyle"
    :data-state="state"
    :data-office-theme="isDark || isPresent ? 'dark' : 'light'"
    :data-encrypted="isEncryptedError ? 'true' : undefined"
    :data-file-drop-active="isDragOver ? 'true' : undefined"
    :data-playback-status="playbackSnapshot?.status"
    :data-slide-index="activeIndex"
    :data-slide-error="slideFailure?.index"
    :data-click-boundary="playbackSnapshot?.clickBoundary"
    data-testid="pptx-viewer"
    tabindex="0"
    @keydown="onKeydown"
    @dragenter="onDragEnter"
    @dragover="onDragOver"
    @dragleave="onDragLeave"
    @drop="onDrop"
  >
    <OfficeToolbar
      v-if="showToolbar && !isPresent"
      class="pptx-viewer__toolbar"
      :label="m.pptx.viewer.toolbar"
    >
      <OfficeToolbarGroup gap="sm" :wrap="false">
        <OfficeIconButton
          :label="m.pptx.viewer.previousSlide"
          icon="chevron-left"
          size="icon-sm"
          :disabled="!canGoPrevious"
          data-testid="pptx-previous-slide"
          @click="run(() => goTo(activeIndex - 1))"
        />
        <OfficePageControl
          :model-value="activeIndex + 1"
          :total-pages="slideCount"
          :disabled="state !== 'ready'"
          data-testid="pptx-page-counter"
          @update:model-value="onPageChange"
        />
        <OfficeIconButton
          :label="m.pptx.viewer.nextSlide"
          icon="chevron-right"
          size="icon-sm"
          :disabled="!canGoNext"
          data-testid="pptx-next-slide"
          @click="run(() => goTo(activeIndex + 1))"
        />
        <OfficeBadge
          v-if="activeSlide?.hidden"
          variant="secondary"
          data-testid="pptx-hidden-badge"
        >{{ m.pptx.viewer.hiddenSlide }}</OfficeBadge>
      </OfficeToolbarGroup>

      <OfficeToolbarSeparator />

      <OfficeToolbarGroup
        gap="sm"
        :wrap="false"
        :label="m.pptx.viewer.viewMode"
      >
        <OfficeIconButton
          :label="m.pptx.viewer.viewModeSingle"
          icon="file-text"
          size="icon-sm"
          :active="!isSorter && !isPresent"
          :disabled="state !== 'ready'"
          data-testid="pptx-view-mode-single"
          @click="setViewMode('single')"
        />
        <OfficeIconButton
          :label="m.pptx.viewer.viewModeSorter"
          icon="layout-grid"
          size="icon-sm"
          :active="isSorter"
          :disabled="state !== 'ready'"
          data-testid="pptx-view-mode-sorter"
          @click="setViewMode('sorter')"
        />
        <OfficeIconButton
          v-if="showPresentToggle"
          :label="m.pptx.viewer.viewModePresent"
          icon="play"
          size="icon-sm"
          :disabled="state !== 'ready'"
          data-testid="pptx-view-mode-present"
          @click="setPresentMode('present')"
        />
      </OfficeToolbarGroup>

      <OfficeIconButton
        v-if="showNotesToggle"
        :label="m.pptx.viewer.notesToggle"
        icon="sticky-note"
        size="icon-sm"
        :active="notesOpen"
        :disabled="state !== 'ready'"
        data-testid="pptx-notes-toggle"
        @click="toggleNotes"
      />

      <template #center>
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
      </template>

      <template #trailing>
        <OfficeToolbarGroup gap="sm" :wrap="false">
          <OfficeIconButton
            :label="m.common.zoom.out"
            icon="minus"
            size="icon-sm"
            :disabled="state !== 'ready' || zoom <= 50"
            data-testid="pptx-zoom-out"
            @click="run(() => setZoom(zoom - 25))"
          />
          <OfficeButton
            variant="outline"
            size="sm"
            class="pptx-viewer__zoom-value"
            :aria-label="m.common.zoom.reset"
            :disabled="state !== 'ready'"
            data-testid="pptx-zoom-reset"
            @click="run(() => setZoom(100))"
          >{{ zoom }}%</OfficeButton>
          <OfficeIconButton
            :label="m.common.zoom.in"
            icon="plus"
            size="icon-sm"
            :disabled="state !== 'ready' || zoom >= 200"
            data-testid="pptx-zoom-in"
            @click="run(() => setZoom(zoom + 25))"
          />
          <OfficeToolbarSeparator v-if="showFullscreenControl" />
          <OfficeIconButton
            v-if="showFullscreenControl"
            :label="m.pptx.viewer.fullscreen"
            :icon="isFullscreen ? 'minimize' : 'maximize'"
            size="icon-sm"
            data-testid="pptx-toolbar-fullscreen"
            @click="void invoke(isFullscreen ? exitFullscreen : enterFullscreen)"
          />
        </OfficeToolbarGroup>
      </template>
    </OfficeToolbar>

    <div v-if="warnings.length" class="pptx-viewer__warnings" role="status" data-testid="pptx-warnings">
      <slot name="warnings" :warnings="warnings">
        <p
          v-for="warning in warnings"
          :key="`${warning.code}:${warning.slideIndex ?? 'all'}`"
          class="pptx-viewer__warning"
        >{{ warning.message }}</p>
      </slot>
    </div>

    <div v-if="dropError" class="pptx-viewer__drop-error" data-testid="pptx-drop-error" role="alert">
      {{ dropError }}
    </div>

    <div class="pptx-viewer__body">
      <OfficeLiveStatus
        test-id="pptx-live-status"
        :message="liveStatusMessage"
        :busy="state === 'loading'"
      />
      <aside
        v-if="showSidebar && !isPresent && !isSorter && document"
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
        <div v-if="state === 'empty'" class="pptx-viewer__message">
          <OfficeEmptyState
            icon="file-text"
            data-testid="pptx-empty"
          >
            <slot name="empty">{{ m.pptx.viewer.empty }}</slot>
          </OfficeEmptyState>
        </div>
        <div v-else-if="state === 'waiting'" class="pptx-viewer__message">
          <OfficeLoadingState
            :title="m.pptx.viewer.waiting"
            data-testid="pptx-waiting"
          >
            <slot name="waiting">{{ m.common.chrome.loadingDescription }}</slot>
          </OfficeLoadingState>
        </div>
        <div v-else-if="state === 'loading'" class="pptx-viewer__message">
          <OfficeLoadingState
            :title="m.pptx.viewer.loading"
            data-testid="pptx-loading"
          >
            <slot name="loading">{{ m.common.chrome.loadingDescription }}</slot>
          </OfficeLoadingState>
        </div>
        <div v-else-if="state === 'cancelled'" class="pptx-viewer__message">
          <OfficeErrorState
            :title="m.pptx.viewer.cancelled"
            icon="circle-x"
            :show-retry="false"
            data-testid="pptx-cancelled"
          >
            <slot name="cancelled" :error="error">{{ error?.message ?? m.common.chrome.errorDescription }}</slot>
            <template #actions>
              <OfficeButton
                variant="outline"
                size="sm"
                data-testid="pptx-cancelled-retry"
                @click="run(reopen)"
              >{{ m.pptx.viewer.reload }}</OfficeButton>
            </template>
          </OfficeErrorState>
        </div>
        <div v-else-if="isEncryptedError" class="pptx-viewer__message">
          <OfficeErrorState
            :title="m.pptx.viewer.encryptedTitle"
            :description="m.pptx.viewer.encryptedDescription"
            icon="lock"
            :show-retry="false"
            data-testid="pptx-encrypted"
          />
        </div>
        <div v-else-if="state === 'error'" class="pptx-viewer__message pptx-viewer__message--error">
          <OfficeErrorState
            :title="m.common.chrome.errorTitle"
            :error="error?.message"
            :show-retry="false"
            data-testid="pptx-error"
          >
            <slot name="error" :error="error">{{ error?.message }}</slot>
            <template #actions>
              <OfficeButton
                v-if="effectiveSource"
                variant="outline"
                size="sm"
                data-testid="pptx-error-retry"
                @click="run(reopen)"
              >{{ m.pptx.viewer.retry }}</OfficeButton>
            </template>
          </OfficeErrorState>
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
            <OfficeButton variant="outline" size="xs" data-testid="pptx-slide-retry" @click="run(retrySlide)">{{ m.pptx.viewer.retry }}</OfficeButton>
            <OfficeButton variant="ghost" size="xs" data-testid="pptx-slide-skip" @click="run(skipSlide)">{{ m.pptx.viewer.skip }}</OfficeButton>
          </span>
        </div>
        <PptxStage
          ref="stageRef"
          class="pptx-viewer__stage"
          :style="stageStyle"
          :aria-hidden="state !== 'ready' || isSorter"
          :zoom="zoom / 100"
          :enable-gesture-zoom="!isPresent"
          :locale="locale"
          :messages="messages"
          :scroll-container="stageScrollRef"
          @update:zoom="run(() => setZoom($event * 100))"
          @click="run(() => onStageClick($event))"
          @context-menu="onStageContextMenu"
          @link-activate="run(() => onLinkActivate($event))"
        />
      </main>

      <div
        v-if="sorterVisible"
        ref="sorterRef"
        class="pptx-viewer__sorter"
        data-testid="pptx-sorter"
        role="listbox"
        :aria-label="m.pptx.viewer.sorter"
        :style="{ '--pptx-sorter-item-height': `${sorterItemHeight}px` }"
        @scroll.passive="onSorterScroll"
      >
        <div
          class="pptx-viewer__sorter-spacer"
          :style="{ height: `${sorterWindow.startRow * sorterItemHeight}px` }"
          aria-hidden="true"
        />
        <div
          class="pptx-viewer__sorter-grid"
          data-testid="pptx-sorter-grid"
          :style="{ gridTemplateColumns: `repeat(${sorterColumns}, minmax(0, 1fr))` }"
        >
          <OfficeCard
            v-for="slide in sorterSlides"
            :key="`${loadGeneration}:${thumbnailFontEpoch}:${slide.index}`"
            class="pptx-viewer__sorter-item"
            data-testid="pptx-sorter-item"
            :data-slide-index="slide.index"
            :data-active="slide.index === activeIndex ? 'true' : 'false'"
            role="option"
            :aria-selected="slide.index === activeIndex"
          >
            <PptxThumbnail
              :session="requireSession()"
              :slide="slide"
              :active="slide.index === activeIndex"
              :width="sorterThumbnailWidth"
              @select="run(() => goTo($event))"
            />
          </OfficeCard>
        </div>
        <div
          class="pptx-viewer__sorter-spacer"
          :style="{ height: `${Math.max(0, sorterRowCount - sorterWindow.endRow) * sorterItemHeight}px` }"
          aria-hidden="true"
        />
      </div>
      <PptxPresenterPanel
        v-if="presenterVisible"
        class="pptx-viewer__presenter"
        :session="pptxDocument.getSession()"
        :slide="activeSlide"
        :next-slide="nextSlide"
        :present="isPresent"
        :running="isPresent && state === 'ready'"
      />
    </div>

    <PptxLinkMenu
      :open="linkMenuOpen"
      :x="linkMenuPosition.x"
      :y="linkMenuPosition.y"
      :items="linkMenuItems"
      :label="m.pptx.viewer.link.menu"
      @update:open="onLinkMenuOpenChange"
      @select="run(() => onLinkMenuSelect($event))"
    />

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
        <button
          v-if="showPresentToggle"
          type="button"
          data-testid="pptx-exit-present"
          :aria-label="m.pptx.viewer.viewModeBrowse"
          @click="exitPresent"
        >{{ m.pptx.viewer.viewModeBrowse }}</button>
      </div>
    </footer>
  </section>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, provide, ref, shallowRef, watch } from "vue"
import {
  PptxPlaybackError,
  PptxPreviewError,
  type PptxCapabilityReport,
  type PptxPlaybackDocument,
  type PptxPlaybackWarning,
  type PptxPreviewDocument,
  type PptxPreviewLimits,
  type PptxPreviewWarning,
  type PptxSlideAction,
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
import type { PptxStageContextMenu, PptxStageLinkActivate } from "./headless-types"
import { OfficeLiveStatus, isEditableViewerTarget, isInteractiveViewerTarget, useFileDrop, useViewerKeys, FILE_DROP_OUTLINE } from "@arcships/vue-ui"
import {
  OfficeBadge,
  OfficeButton,
  OfficeCard,
  OfficeEmptyState,
  OfficeErrorState,
  OfficeIconButton,
  OfficeLoadingState,
  OfficePageControl,
  OfficeToolbar,
  OfficeToolbarGroup,
  OfficeToolbarSeparator,
} from "@arcships/vue-ui"
import type { OfficeMenuEntry, ViewerKeyCommand } from "@arcships/vue-ui"
import type { OfficeLocale, OfficeMessagesOverride } from "@arcships/office-interaction"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { providePptxMessages } from "./i18n"
import {
  createPptxActionRequest,
  createPptxLinkIndex,
  openPptxLink,
  pptxLinkCopyValue,
  resolvePptxLinkEntry,
  resolvePptxLinkTarget,
  sanitizePptxActionRequest,
  type PptxLinkEntry,
  type PptxLinkTarget,
} from "./links"
import { PPTX_FILE_EXTENSIONS, PPTX_FILE_MIME_TYPES } from "./pptx-file-types"
import { usePptxDocument } from "./composables/usePptxDocument"
import { usePptxPlayback } from "./composables/usePptxPlayback"
import { watchFontLoading } from "./composables/font-loading"
import {
  PPTX_THUMBNAIL_OBSERVER,
  createPptxThumbnailObserver,
} from "./composables/thumbnail-observer"
import PptxStage from "./PptxStage.vue"
import PptxThumbnail from "./PptxThumbnail.vue"
import PptxPresenterPanel from "./PptxPresenterPanel.vue"
import PptxLinkMenu from "./PptxLinkMenu.vue"

const props = withDefaults(defineProps<{
  source?: PptxPreviewSource | null
  /** Display mode. Leave undefined to let the toolbar toggle control it. */
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
  showNotes?: boolean
  showPlaybackControls?: boolean
  showSearch?: boolean
  showSidebar?: boolean
  showToolbar?: boolean
  showWarnings?: boolean
  /** Toolbar toggle for the speaker-notes / presenter rail. */
  showNotesToggle?: boolean
  /** Toolbar button that switches to presentation mode. */
  showPresentToggle?: boolean
  /** Toolbar button that enters / leaves fullscreen in browse mode. */
  showFullscreenControl?: boolean
  /** Built-in link menu (open link / copy link) on right click. */
  showContextMenu?: boolean
  /**
   * `built-in` opens approved links and copies them from the context menu;
   * `host` only emits `action`/`contextMenu` and leaves the policy to the host.
   */
  linkHandling?: "built-in" | "host"
  thumbnailWidth?: number
  /** Browse layout: one slide list (`single`) or a slide-sorter grid (`sorter`). */
  viewMode?: "single" | "sorter" | null
  /** Browse layout used while `viewMode` is uncontrolled. Defaults to single. */
  defaultViewMode?: "single" | "sorter"
  locale?: OfficeLocale
  messages?: OfficeMessagesOverride
}>(), {
  source: null,
  mode: undefined,
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
  showNotes: false,
  showPlaybackControls: true,
  showSearch: true,
  showSidebar: true,
  showToolbar: true,
  showWarnings: true,
  showNotesToggle: true,
  showPresentToggle: true,
  showFullscreenControl: true,
  showContextMenu: true,
  linkHandling: "built-in" as "built-in" | "host",
  thumbnailWidth: 148,
  viewMode: null,
  defaultViewMode: "single" as "single" | "sorter",
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
  /** A supported presentation was dropped; the viewer opened it through `usePptxDocument`. */
  fileDrop: [file: File]
  /** Right-click on a slide or object; hosts can replace the built-in link menu. */
  contextMenu: [ctx: PptxStageContextMenu]
  /** The browse layout toggle changed; hosts can drive the `viewMode` prop. */
  "update:viewMode": [mode: "single" | "sorter"]
  /** The toolbar switched between browse and presentation mode. */
  "update:mode": [mode: "browse" | "present"]
  /** The toolbar notes toggle changed the presenter rail. */
  "update:showNotes": [visible: boolean]
}>()

type ViewerState = "empty" | "waiting" | "loading" | "ready" | "cancelled" | "error"

/** Padding, border, label and gap around the thumbnail canvas, kept in sync with `style.css`. */
const THUMBNAIL_CHROME_PX = 48
const THUMBNAIL_MIN_CANVAS_PX = 84
const SIDEBAR_OVERSCAN_ROWS = 3
const SIDEBAR_FALLBACK_VIEWPORT_PX = 640

const m = providePptxMessages(() => props.locale, () => props.messages)

// ── Drag-and-drop file import ──────────────────────────────────────
// A dropped presentation becomes the effective source; `usePptxDocument`
// opens it through the same session pipeline as the `source` prop.
const droppedSource = shallowRef<PptxPreviewSource | null>(null)
const dropError = ref("")
const effectiveSource = computed<PptxPreviewSource | null>(() => droppedSource.value ?? props.source ?? null)
const fileDrop = useFileDrop({
  accept: { extensions: PPTX_FILE_EXTENSIONS, mimeTypes: PPTX_FILE_MIME_TYPES },
  onFile: (file) => {
    dropError.value = ""
    emit("fileDrop", file)
    droppedSource.value = file
  },
  onReject: () => {
    dropError.value = m.value.pptx.viewer.unsupportedDrop
  },
})
const { isDragOver, onDragEnter, onDragOver, onDragLeave, onDrop } = fileDrop

// A host-owned source supersedes a file dropped earlier. This watcher is
// registered before `usePptxDocument` opens its own source watcher below, so
// the dropped source is already cleared when that one reads the new value.
watch(() => props.source, () => { droppedSource.value = null; dropError.value = "" })

const viewerRef = ref<HTMLElement | null>(null)
const searchInputRef = ref<HTMLInputElement | null>(null)
const stageScrollRef = ref<HTMLElement | null>(null)
const stageRef = ref<PptxStageExpose | null>(null)
const stageElement = computed(() => stageRef.value?.element ?? null)

// ── Browse / present mode ──────────────────────────────────────────
// `mode` stays the host-controlled source of truth. When it is omitted the
// toolbar switches an internal mode, which is what the shared view-mode cluster
// binds to; either way `update:mode` reports the change.
const internalMode = ref<"browse" | "present">("browse")
const effectiveMode = computed<"browse" | "present">(() =>
  props.mode === "present" || props.mode === "browse" ? props.mode : internalMode.value,
)
const isPresent = computed(() => effectiveMode.value === "present")

function setPresentMode(mode: "browse" | "present"): void {
  if (mode === effectiveMode.value) return
  if (props.mode == null) internalMode.value = mode
  emit("update:mode", mode)
}

function exitPresent(): void {
  if (!isPresent.value) return
  setPresentMode("browse")
  if (isFullscreen.value) void invoke(exitFullscreen)
}

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
  source: () => effectiveSource.value,
  initialSlide: () => props.initialSlide,
  openTimeoutMs: () => props.openTimeoutMs,
  messages: () => m.value.pptx.documentErrors,
  session: () => ({
    fitMode: "contain" as const,
    zoomPercent: 100,
    scrollContainer: stageScrollRef.value ?? undefined,
    renderMode: isPresent.value ? "slide" as const : "list" as const,
    listOptions: isPresent.value ? undefined : {
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
    if (isPresent.value) {
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
/**
 * Password-protected presentations (and legacy OLE/CFB binaries renamed to
 * `.pptx`) are rejected by the core before the zip reader runs. They get the
 * dedicated "protected" state instead of the generic parse-failure copy.
 */
const isEncryptedError = computed(
  () => state.value === "error" && error.value?.code === "UNSUPPORTED_CONTENT",
)
const slideCount = computed(() => document.value?.slides.length ?? 0)

function onPageChange(page: number): void {
  const last = Math.max(slideCount.value - 1, 0)
  run(() => goTo(Math.min(Math.max(page - 1, 0), last)))
}

const isFullscreen = ref(false)
const printPageStyle = ref<HTMLStyleElement | undefined>()

// ── Browse layout (single slide list / slide sorter) ────────────────
// The sorter reuses the sidebar thumbnails in a windowed grid. The stage stays
// mounted (a hidden overlay covers it) because the document session renders
// into it; that is also what keeps the current slide across mode switches.
const SORTER_MIN_COLUMNS = 2
const SORTER_MAX_COLUMNS = 5
const SORTER_FALLBACK_COLUMNS = 3
const SORTER_GRID_GAP_PX = 12
const SORTER_OVERSCAN_ROWS = 1
const SORTER_FALLBACK_VIEWPORT_PX = 640
const internalViewMode = ref<"single" | "sorter">(
  props.defaultViewMode === "sorter" ? "sorter" : "single",
)
const effectiveViewMode = computed<"single" | "sorter">(() =>
  props.viewMode === "sorter" || props.viewMode === "single"
    ? props.viewMode
    : internalViewMode.value,
)
const isSorter = computed(() => !isPresent.value && effectiveViewMode.value === "sorter")
const sorterVisible = computed(
  () => isSorter.value && state.value === "ready" && document.value !== null,
)

const sorterRef = ref<HTMLElement | null>(null)
const sorterScrollTop = ref(0)
const sorterViewportHeight = ref(SORTER_FALLBACK_VIEWPORT_PX)
const sorterViewportWidth = ref(0)
const sorterItemWidth = computed(() => Math.max(120, props.thumbnailWidth))
const sorterThumbnailWidth = computed(() => Math.max(96, sorterItemWidth.value - 16))
const sorterItemHeight = computed(() => {
  const slideshow = document.value
  const ratio = slideshow && slideshow.width > 0 ? slideshow.height / slideshow.width : 9 / 16
  return Math.max(THUMBNAIL_MIN_CANVAS_PX, Math.round(sorterThumbnailWidth.value * ratio)) + THUMBNAIL_CHROME_PX
})
const sorterColumns = computed(() => {
  const width = sorterViewportWidth.value
  if (!(width > 0)) return SORTER_FALLBACK_COLUMNS
  const columns = Math.floor((width - 24 + SORTER_GRID_GAP_PX) / (sorterItemWidth.value + SORTER_GRID_GAP_PX))
  return Math.min(SORTER_MAX_COLUMNS, Math.max(SORTER_MIN_COLUMNS, columns))
})
const sorterRowCount = computed(() => {
  const count = document.value?.slides.length ?? 0
  return Math.ceil(count / sorterColumns.value)
})
const sorterWindow = computed(() => {
  const count = document.value?.slides.length ?? 0
  const columns = sorterColumns.value
  const rowHeight = sorterItemHeight.value
  const firstVisible = Math.floor(Math.max(0, sorterScrollTop.value) / Math.max(1, rowHeight))
  const visibleRows = Math.ceil(sorterViewportHeight.value / Math.max(1, rowHeight)) + 1
  const startRow = Math.max(0, Math.min(firstVisible, Math.max(0, sorterRowCount.value - 1)) - SORTER_OVERSCAN_ROWS)
  const endRow = Math.min(
    sorterRowCount.value,
    Math.max(startRow, firstVisible + visibleRows + SORTER_OVERSCAN_ROWS),
  )
  return {
    start: startRow * columns,
    end: Math.min(count, endRow * columns),
    startRow,
    endRow,
  }
})
const sorterSlides = computed(() =>
  document.value?.slides.slice(sorterWindow.value.start, sorterWindow.value.end) ?? [],
)

function measureSorter(): void {
  const element = sorterRef.value
  if (!element) return
  const height = typeof element.clientHeight === "number" ? element.clientHeight : 0
  if (height > 0) sorterViewportHeight.value = height
  const width = typeof element.clientWidth === "number" ? element.clientWidth : 0
  if (width > 0) sorterViewportWidth.value = width
}

function setSorterScrollTop(value: number): void {
  sorterScrollTop.value = Math.max(0, value)
  const element = sorterRef.value
  if (element && typeof element.scrollTop === "number") element.scrollTop = sorterScrollTop.value
}

function revealSorterSlide(index: number): void {
  const element = sorterRef.value
  if (!element) return
  measureSorter()
  const columns = sorterColumns.value
  const row = Math.floor(Math.max(0, index) / Math.max(1, columns))
  const top = row * sorterItemHeight.value
  const bottom = top + sorterItemHeight.value
  const current = sorterScrollTop.value
  const viewport = sorterViewportHeight.value
  const next = top < current ? top : bottom > current + viewport ? bottom - viewport : current
  if (next !== current) setSorterScrollTop(next)
}

function onSorterScroll(event: Event): void {
  const target = event.target as { scrollTop?: unknown; clientHeight?: unknown; clientWidth?: unknown } | null
  if (typeof target?.scrollTop === "number") sorterScrollTop.value = Math.max(0, target.scrollTop)
  if (typeof target?.clientHeight === "number" && target.clientHeight > 0) {
    sorterViewportHeight.value = target.clientHeight
  }
  if (typeof target?.clientWidth === "number" && target.clientWidth > 0) {
    sorterViewportWidth.value = target.clientWidth
  }
}

function setViewMode(mode: "single" | "sorter"): void {
  const next = mode === "sorter" ? "sorter" : "single"
  if (next === effectiveViewMode.value) return
  if (props.viewMode == null) internalViewMode.value = next
  emit("update:viewMode", next)
  // Switching back to the single view must not lose the selected slide.
  if (next === "sorter") void nextTick(() => revealSorterSlide(activeIndex.value))
}

function toCssLength(value: string | number | undefined): string | undefined {
  if (value === undefined) return undefined
  return typeof value === "number" ? `${value}px` : value
}

const rootStyle = computed(() => {
  const outline = isDragOver.value ? FILE_DROP_OUTLINE : undefined
  if (isFullscreen.value) return outline ? { outline } : undefined
  const style: Record<string, string> = { height: props.height, maxHeight: props.height }
  if (outline) style.outline = outline
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

// ── Speaker notes and presenter view ────────────────────────────────
// `showNotes` (or the toolbar toggle) mounts the presenter rail: notes for the
// current slide in browse mode, plus the next-slide preview and the elapsed
// clock in present mode.
const notesOpen = ref(props.showNotes)
watch(() => props.showNotes, (visible) => { notesOpen.value = visible })

function toggleNotes(): void {
  notesOpen.value = !notesOpen.value
  emit("update:showNotes", notesOpen.value)
}

const presenterVisible = computed(() =>
  notesOpen.value && state.value === "ready" && !isSorter.value)
const nextSlide = computed(() => {
  const slides = document.value?.slides
  if (!slides) return null
  const skipHidden = !props.showHiddenSlides
  for (let index = activeIndex.value + 1; index < slides.length; index += 1) {
    const candidate = slides[index]
    if (!skipHidden || !candidate.hidden) return candidate
  }
  return null
})

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
    // Present mode: the controller already navigated internal jumps, so only
    // external targets need the built-in opener. The host still sees the same
    // action request, re-checked against the shared URL policy.
    const action = sanitizePptxActionRequest(event.action)
    if (props.linkHandling !== "host" && (action.kind === "open-url" || action.kind === "mailto")) {
      openPptxLink(action.url)
    }
    emit("action", action)
  } else if (event.type === "error") {
    emit("playbackError", event.error)
  }
}

// ── Hyperlinks and the built-in link menu ───────────────────────────
// The playback parser resolves shape links for every render mode; browse mode
// has no playback controller, so the viewer routes them itself. Both modes
// share `resolvePptxLinkTarget()` (the office link policy) and both keep
// emitting `action` for hosts. The menu surface itself is the shared
// `OfficeDropdownMenu` (see `PptxLinkMenu.vue`), anchored at the right-click
// point so focus, typeahead, arrow keys and Escape stay primitive-owned.
interface ContextMenuState {
  slideIndex: number
  objectKey: string
  x: number
  y: number
  action: PptxSlideAction
  target: PptxLinkTarget
}

const EMPTY_LINK_INDEX: ReadonlyMap<string, PptxLinkEntry> = new Map()
const contextMenu = shallowRef<ContextMenuState | null>(null)
const linkMenuOpen = ref(false)
const linkMenuPosition = ref({ x: 0, y: 0 })
const linkMenuVisible = computed(() => contextMenu.value !== null && props.showContextMenu !== false)
const contextMenuCopyValue = computed(() => {
  const target = contextMenu.value?.target
  return target ? pptxLinkCopyValue(target) : undefined
})
const contextMenuOpenLabel = computed(() => {
  const target = contextMenu.value?.target
  if (target?.kind === "slide") {
    return formatOfficeMessage(m.value.pptx.viewer.link.openSlide, { number: target.slideIndex + 1 })
  }
  return m.value.pptx.viewer.link.open
})
const linkMenuItems = computed<OfficeMenuEntry[]>(() => {
  const target = contextMenu.value?.target
  const items: OfficeMenuEntry[] = [{
    id: "open",
    label: contextMenuOpenLabel.value,
    icon: target?.kind === "slide" ? "arrow-right" : "link",
  }]
  if (contextMenuCopyValue.value) {
    items.push({ id: "copy", label: m.value.pptx.viewer.link.copy, icon: "copy" })
  }
  return items
})

function currentPlaybackDocument(): PptxPlaybackDocument | null {
  return pptxDocument.getSession()?.playbackDocument ?? null
}

// Reads `document` so a newly opened deck builds a fresh index.
const linkIndex = computed<ReadonlyMap<string, PptxLinkEntry>>(() => {
  if (state.value !== "ready" || !document.value) return EMPTY_LINK_INDEX
  return createPptxLinkIndex(currentPlaybackDocument())
})

function linkActionFor(slideIndex: number, objectKey: string | undefined, trigger: "click" | "hover" = "click"): PptxSlideAction | undefined {
  if (!objectKey || !Number.isInteger(slideIndex)) return undefined
  return resolvePptxLinkEntry(linkIndex.value, slideIndex, objectKey)?.[trigger]
}

async function applyLinkTarget(target: PptxLinkTarget): Promise<void> {
  if (props.linkHandling === "host") return
  if (target.kind === "slide") {
    await goTo(target.slideIndex)
  } else if (target.kind === "url") {
    openPptxLink(target.url)
  }
}

/** Activate a parsed link: report it to the host, then run the built-in policy. */
async function activateLink(action: PptxSlideAction): Promise<void> {
  const target = resolvePptxLinkTarget(action)
  emit("action", createPptxActionRequest(action, target))
  await applyLinkTarget(target)
}

function objectKeyFromEvent(event: MouseEvent): { slideIndex: number; objectKey?: string } | null {
  const target = event.target as Element | null
  if (typeof target?.closest !== "function") return null
  const slideElement = target.closest<HTMLElement>("[data-slide-index]")
  const slideIndex = Number(slideElement?.dataset.slideIndex)
  if (!Number.isInteger(slideIndex) || slideIndex < 0) return null
  return {
    slideIndex,
    objectKey: target.closest<HTMLElement>("[data-pptx-object-key]")?.dataset.pptxObjectKey || undefined,
  }
}

/**
 * Browse-mode link clicks.
 *
 * The renderer already turns some text runs into anchors; once the shape has a
 * parsed action the viewer owns the click so the URL policy applies once and
 * the link does not also open natively.
 */
function onBrowseLinkClick(event: MouseEvent): void {
  const hit = objectKeyFromEvent(event)
  if (!hit?.objectKey) return
  const action = linkActionFor(hit.slideIndex, hit.objectKey)
  if (!action) return
  if (props.linkHandling !== "host") event.preventDefault?.()
  run(() => activateLink(action))
}

async function onLinkActivate(target: PptxStageLinkActivate): Promise<void> {
  if (isPresent.value) {
    // Present mode goes through the controller so it also runs the click-trigger
    // animation; the controller emits the action the viewer then opens.
    await invoke(async () => {
      await requireController().activateObject(target.objectKey)
    })
    return
  }
  const action = linkActionFor(target.slideIndex, target.objectKey)
  if (action) await activateLink(action)
}

function onStageContextMenu(context: PptxStageContextMenu): void {
  emit("contextMenu", context)
  if (props.showContextMenu === false || state.value !== "ready" || context.kind !== "object") {
    closeContextMenu()
    return
  }
  const action = linkActionFor(context.slideIndex, context.objectKey)
  const target = resolvePptxLinkTarget(action)
  if (!action || target.kind === "unsupported") {
    closeContextMenu()
    return
  }
  contextMenu.value = {
    slideIndex: context.slideIndex,
    objectKey: context.objectKey,
    x: context.clientX,
    y: context.clientY,
    action,
    target,
  }
  linkMenuPosition.value = { x: context.clientX, y: context.clientY }
  linkMenuOpen.value = true
}

function closeContextMenu(): void {
  linkMenuOpen.value = false
  contextMenu.value = null
}

function onLinkMenuOpenChange(open: boolean): void {
  if (open) {
    linkMenuOpen.value = true
    return
  }
  const wasOpen = linkMenuOpen.value
  closeContextMenu()
  // The Reka trigger is the zero-size pointer anchor; hand focus back to the
  // viewer instead of leaving it on a hidden helper.
  if (wasOpen) void nextTick(() => viewerRef.value?.focus())
}

async function onLinkMenuSelect(id: string): Promise<void> {
  if (id === "open") await openContextMenuLink()
  else if (id === "copy") await copyContextMenuLink()
}

async function openContextMenuLink(): Promise<void> {
  const state = contextMenu.value
  closeContextMenu()
  if (!state) return
  await activateLink(state.action)
}

async function copyContextMenuLink(): Promise<void> {
  const value = contextMenuCopyValue.value
  closeContextMenu()
  if (!value) return
  const clipboard = (globalThis.navigator as { clipboard?: { writeText?: (text: string) => Promise<void> } } | undefined)?.clipboard
  try {
    await clipboard?.writeText?.(value)
  } catch {
    // Copying is a convenience; a blocked clipboard leaves the link untouched.
  }
}

// A menu that outlives its target would act on the wrong slide.
watch([document, activeIndex, state, effectiveMode], closeContextMenu)

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
  if (isSorter.value) revealSorterSlide(index)
}, { flush: "post" })

// The sorter is a windowed grid; measure it once it is actually on screen so
// the column count matches the width the host gave the viewer.
watch(sorterVisible, (visible) => {
  if (!visible) return
  void nextTick(() => {
    measureSorter()
    revealSorterSlide(activeIndex.value)
  })
})

watch(() => props.viewMode, (mode) => {
  if (mode !== "single" && mode !== "sorter") return
  if (mode === "sorter") void nextTick(() => revealSorterSlide(activeIndex.value))
})

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
  const source = effectiveSource.value
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
  if (state.value !== "ready") return
  if (!isPresent.value) {
    onBrowseLinkClick(event)
    return
  }
  viewerRef.value?.focus()
  const target = event.target as Element | null
  const objectElement = typeof target?.closest === "function"
    ? target.closest<HTMLElement>("[data-pptx-object-key]")
    : null
  const objectKey = objectElement?.dataset.pptxObjectKey
  // A nested text anchor would open natively in addition to the action; when
  // the shape has a parsed link the viewer owns the single, policy-checked open.
  if (objectKey && linkActionFor(playbackSnapshot.value?.slideIndex ?? activeIndex.value, objectKey)) {
    event.preventDefault?.()
  }
  await invoke(async () => {
    const handled = objectKey ? await requireController().activateObject(objectKey) : false
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
  // The sorter owns scrolling while it is open; the slide stage is hidden behind it.
  const container = isSorter.value ? sorterRef.value : stageScrollRef.value
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
  if (isSorter.value) {
    const element = sorterRef.value
    const scrollHeight = typeof element?.scrollHeight === "number" ? element.scrollHeight : 0
    const clientHeight = typeof element?.clientHeight === "number" ? element.clientHeight : 0
    const max = Math.max(0, scrollHeight - clientHeight)
    setSorterScrollTop(Math.min(max, Math.max(0, top)))
    return
  }
  const container = stageScrollRef.value
  if (!container) return
  const scrollHeight = typeof container.scrollHeight === "number" ? container.scrollHeight : 0
  const clientHeight = typeof container.clientHeight === "number" ? container.clientHeight : 0
  const max = Math.max(0, scrollHeight - clientHeight)
  container.scrollTop = Math.min(max, Math.max(0, top))
}

function scrollBrowseBy(delta: number): void {
  if (isSorter.value) {
    setBrowseScrollTop(sorterScrollTop.value + delta)
    return
  }
  const container = stageScrollRef.value
  if (!container) return
  const top = typeof container.scrollTop === "number" ? container.scrollTop : 0
  setBrowseScrollTop(top + delta)
}

watch(() => props.initialSlide, (index) => {
  if (state.value === "ready") run(() => goTo(index))
})

watch(effectiveMode, () => {
  const source = effectiveSource.value
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
