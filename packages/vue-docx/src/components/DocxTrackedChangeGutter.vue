<template>
  <div
    v-if="positioned.length > 0 || composer"
    class="docx-tracked-change-gutter"
    :data-office-theme="documentTheme"
    data-office-density="comfortable"
    :class="{
      'docx-tracked-change-gutter--drawer': drawerMode,
      'docx-tracked-change-gutter--drawer-closed': drawerMode && !drawerOpen,
    }"
    data-docx-thumbnail-exclude="true"
    data-docx-annotation-gutter="true"
    :data-docx-annotation-gutter-mode="drawerMode ? 'drawer' : 'column'"
    :data-docx-annotation-drawer-open="drawerMode ? String(drawerOpen) : undefined"
    :data-docx-gutter-card-count="positioned.length"
    :data-docx-gutter-zoom="zoomFactor"
    :style="paneStyle"
  >
    <!--
      Narrow viewports have no room for a column beside the paper, so the pane
      becomes an overlay drawer with a tab: the paper keeps the full viewport
      width and the cards slide over its right edge when opened.
    -->
    <button
      v-if="drawerMode"
      type="button"
      class="docx-gutter-drawer-tab"
      data-docx-gutter-drawer-tab="true"
      :aria-expanded="drawerOpen ? 'true' : 'false'"
      :title="drawerOpen ? m.docx.gutter.hideComments : m.docx.gutter.showComments"
      @click.stop="toggleDrawer"
      @keydown.stop
    >
      <span class="docx-gutter-drawer-tab-count">{{ positioned.length }}</span>
      <span class="docx-gutter-drawer-tab-label">{{ drawerOpen ? m.docx.gutter.comments : "" }}</span>
    </button>
    <OfficeCard
      v-for="entry of visibleCards"
      :key="entry.annotation.id"
      :ref="(element) => registerCardElement(entry.annotation.id, cardElement(element))"
      class="docx-gutter-card"
      :class="cardClasses(entry).join(' ')"
      :style="cardStyle(entry)"
      :data-docx-gutter-card-id="entry.annotation.id"
      :data-docx-gutter-card-index="entry.annotation.ordinal"
      :data-docx-gutter-anchor-y="entry.anchorY"
      :data-docx-gutter-card-top="entry.top"
      :data-docx-tracked-change-kind="entry.annotation.trackedChange?.kind"
      :data-docx-gutter-balloon="gutterBalloonFlag(entry)"
      role="button"
      tabindex="0"
      @click="emit('activate', entry.annotation.id)"
      @keydown.enter.prevent="emit('activate', entry.annotation.id)"
      @keydown.space.prevent="emit('activate', entry.annotation.id)"
    >
      <!--
        A host that provides `render-comment-card` / `render-tracked-change-card`
        replaces the card body; the positioned wrapper, activation, measurement
        and drawer visibility stay owned by this pane so cards keep anchoring.
      -->
      <HostCard
        v-if="customCards[entry.annotation.id]"
        v-bind="customCards[entry.annotation.id]"
      />
      <template v-else>
        <div class="docx-gutter-card-header">
          <span
            class="docx-gutter-card-index"
            :style="{ backgroundColor: accentColor(entry) }"
          >{{ entry.annotation.ordinal }}</span>
          <span class="docx-gutter-card-type">{{ cardTypeLabel(entry) }}</span>
          <span class="docx-gutter-card-author">{{ authorOf(entry) }}</span>
        </div>
        <div class="docx-gutter-card-text">
          {{ cardText(entry) }}
        </div>
        <div
          v-if="canDecide(entry)"
          class="docx-gutter-card-actions"
          @click.stop
          @keydown.enter.stop
          @keydown.space.stop
        >
          <button
            type="button"
            class="docx-gutter-card-action docx-gutter-card-action--accept"
            data-testid="docx-gutter-accept"
            :title="m.docx.gutter.acceptTitle"
            @click.stop="emit('decide', entry.annotation.id, 'accept')"
          >
            ✓ {{ m.docx.gutter.accept }}
          </button>
          <button
            type="button"
            class="docx-gutter-card-action docx-gutter-card-action--reject"
            data-testid="docx-gutter-reject"
            :title="m.docx.gutter.rejectTitle"
            @click.stop="emit('decide', entry.annotation.id, 'reject')"
          >
            ✕ {{ m.docx.gutter.reject }}
          </button>
        </div>
        <div
          v-if="commentActions && entry.annotation.comment"
          class="docx-gutter-card-actions"
          @click.stop
          @keydown.enter.stop
          @keydown.space.stop
        >
          <button
            type="button"
            class="docx-gutter-card-action docx-gutter-card-action--comment"
            data-testid="docx-gutter-comment-reply"
            :title="m.docx.gutter.replyTitle"
            @click.stop="emit('reply', entry.annotation.comment.commentId)"
          >
            {{ m.docx.gutter.reply }}
          </button>
          <button
            v-if="isThreadRoot(entry)"
            type="button"
            class="docx-gutter-card-action docx-gutter-card-action--comment"
            data-testid="docx-gutter-comment-resolve"
            :title="entry.annotation.comment.resolved ? m.docx.gutter.reopenTitle : m.docx.gutter.resolveTitle"
            @click.stop="emit('resolve', entry.annotation.comment.commentId, !entry.annotation.comment.resolved)"
          >
            {{ entry.annotation.comment.resolved ? m.docx.gutter.reopen : m.docx.gutter.resolve }}
          </button>
          <button
            type="button"
            class="docx-gutter-card-action docx-gutter-card-action--comment"
            data-testid="docx-gutter-comment-delete"
            :title="m.docx.gutter.deleteTitle"
            @click.stop="emit('delete', entry.annotation.comment.commentId)"
          >
            {{ m.docx.gutter.delete }}
          </button>
        </div>
        <!--
          The reply editor lives inside the card it replies to, so the new text
          appears in the thread instead of a detached dialog.
        -->
        <div
          v-if="replyComposerFor(entry)"
          class="docx-gutter-card-composer"
          data-testid="docx-gutter-comment-composer"
          :data-docx-comment-composer-mode="composerMode"
          @click.stop
          @keydown.enter.stop
          @keydown.space.stop
        >
          <div class="docx-gutter-card-composer-title">
            {{ m.docx.gutter.composerReplyTitle }}
          </div>
          <textarea
            class="docx-gutter-card-composer-input"
            data-testid="docx-gutter-comment-composer-input"
            :aria-label="m.docx.gutter.composerReplyPlaceholder"
            :placeholder="m.docx.gutter.composerReplyPlaceholder"
            :value="draft"
            rows="2"
            @input="onComposerInput"
          />
          <div class="docx-gutter-card-actions">
            <button
              type="button"
              class="docx-gutter-card-action docx-gutter-card-action--comment"
              data-testid="docx-gutter-comment-composer-submit"
              @click.stop="submitComposer"
            >
              {{ m.docx.gutter.composerSubmit }}
            </button>
            <button
              type="button"
              class="docx-gutter-card-action"
              data-testid="docx-gutter-comment-composer-cancel"
              @click.stop="cancelComposer"
            >
              {{ m.docx.gutter.composerCancel }}
            </button>
          </div>
        </div>
        <div v-if="entry.annotation.comment?.resolved" class="docx-gutter-card-resolved">
          ✓ {{ m.docx.trackedChange.resolved }}
        </div>
        <div v-if="dateOf(entry)" class="docx-gutter-card-date">
          {{ formatDate(dateOf(entry)) }}
        </div>
      </template>
    </OfficeCard>

    <!--
      A new comment has no gutter card yet, so its editor sits at the top of
      the pane until the comment is submitted.
    -->
    <div
      v-if="composer && composer.mode === 'new' && (!drawerMode || drawerOpen)"
      class="docx-gutter-card docx-gutter-card--comment docx-gutter-card--composer"
      data-testid="docx-gutter-comment-composer"
      :data-docx-comment-composer-mode="composerMode"
      @click.stop
      @keydown.enter.stop
      @keydown.space.stop
    >
      <div class="docx-gutter-card-header">
        <span class="docx-gutter-card-type">{{ m.docx.gutter.composerTitle }}</span>
      </div>
      <textarea
        class="docx-gutter-card-composer-input"
        data-testid="docx-gutter-comment-composer-input"
        :aria-label="m.docx.gutter.composerPlaceholder"
        :placeholder="m.docx.gutter.composerPlaceholder"
        :value="draft"
        rows="3"
        @input="onComposerInput"
      />
      <div class="docx-gutter-card-actions">
        <button
          type="button"
          class="docx-gutter-card-action docx-gutter-card-action--comment"
          data-testid="docx-gutter-comment-composer-submit"
          @click.stop="submitComposer"
        >
          {{ m.docx.gutter.composerSubmit }}
        </button>
        <button
          type="button"
          class="docx-gutter-card-action"
          data-testid="docx-gutter-comment-composer-cancel"
          @click.stop="cancelComposer"
        >
          {{ m.docx.gutter.composerCancel }}
        </button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, defineComponent, nextTick, onMounted, ref, watch } from "vue"
import type { PropType, VNodeChild } from "vue"
import {
  commentAccentColor,
  trackedChangeAccentColor,
  trackedChangeUsesGutterBalloon,
  type DocxCommentCardRenderProps,
  type DocxDocumentTheme,
  type DocxTrackedChange,
  type DocxTrackedChangeCardRenderProps,
  type PositionedGutterAnnotation,
} from "@arcships/docx-core"
import {
  annotationGutterCardTopPx,
  annotationGutterDrawerPaneStyle,
  annotationGutterPaneStyle,
} from "./gutter-metrics"
import { formatOfficeDate } from "@arcships/office-interaction"
import { OfficeCard } from "@arcships/vue-ui"
import { useDocxLocale, useDocxMessages } from "../i18n"

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    /** Cards placed against their anchors, in document order. */
    positioned: readonly PositionedGutterAnnotation[]
    /** Viewer zoom factor of the page subtree this pane lives in. */
    zoomFactor: number
    /** Page height in page space, used to size the pane. */
    pageHeightPx: number
    documentTheme?: DocxDocumentTheme
    activeAnnotationId?: string
    /** True when the host can accept/reject revisions on the cards. */
    decidable?: boolean
    /** Render as an overlay drawer instead of a column beside the paper. */
    drawerMode?: boolean
    /** True when the host can reply to, resolve and delete comments. */
    commentActions?: boolean
    /** Open inline comment editor rendered by this pane. */
    composer?: { mode: "new" | "reply"; commentId?: number }
    /** Zero-based page index these cards belong to. */
    pageIndex?: number
    /**
     * Host renderer that replaces every comment card body. The pane keeps the
     * positioned wrapper, activation and measurement.
     */
    renderCommentCard?: (props: DocxCommentCardRenderProps) => unknown
    /** Host renderer that replaces every tracked-change card body. */
    renderTrackedChangeCard?: (
      props: DocxTrackedChangeCardRenderProps
    ) => unknown
  }>(),
  {
    zoomFactor: 1,
    pageHeightPx: 0,
    documentTheme: "light" as DocxDocumentTheme,
    decidable: false,
    drawerMode: false,
    commentActions: false,
    composer: undefined,
    pageIndex: 0,
    renderCommentCard: undefined,
    renderTrackedChangeCard: undefined,
  }
)

const emit = defineEmits<{
  activate: [annotationId: string]
  cardHeight: [annotationId: string, heightPx: number]
  decide: [annotationId: string, decision: "accept" | "reject"]
  reply: [commentId: number]
  resolve: [commentId: number, resolved: boolean]
  delete: [commentId: number]
  composerSubmit: [text: string]
  composerCancel: []
}>()
const m = useDocxMessages()
const locale = useDocxLocale()

// ── Host-rendered cards ────────────────────────────────────────────
// A host may replace the body of comment and tracked-change cards. The card
// wrapper below still owns positioning, activation, measurement and drawer
// visibility, so custom cards line up with the same anchors as the defaults.
type HostCardRenderer =
  | ((props: DocxCommentCardRenderProps) => unknown)
  | ((props: DocxTrackedChangeCardRenderProps) => unknown)

interface HostCardEntry {
  render: HostCardRenderer
  renderProps: DocxCommentCardRenderProps | DocxTrackedChangeCardRenderProps
}

const HostCard = defineComponent({
  name: "DocxGutterHostCard",
  props: {
    render: {
      type: Function as PropType<HostCardRenderer>,
      required: true,
    },
    renderProps: {
      type: Object as PropType<
        DocxCommentCardRenderProps | DocxTrackedChangeCardRenderProps
      >,
      required: true,
    },
  },
  setup(card) {
    return () =>
      (card.render as unknown as (props: unknown) => VNodeChild)(
        card.renderProps
      )
  },
})

/** Positioning style handed to host cards; the pane already applies it. */
function cardRenderStyle(
  entry: PositionedGutterAnnotation
): Record<string, string | number | undefined> {
  return {
    position: "absolute",
    left: 0,
    right: 0,
    top: `${annotationGutterCardTopPx(entry.top, props.zoomFactor)}px`,
  }
}

function hostCardFor(entry: PositionedGutterAnnotation): HostCardEntry | undefined {
  const comment = entry.annotation.comment
  if (comment && props.renderCommentCard) {
    return {
      render: props.renderCommentCard,
      renderProps: {
        comment,
        snippet: cardText(entry),
        formattedDate: formatDate(comment.date) || undefined,
        accentColor: accentColor(entry),
        documentTheme: props.documentTheme,
        pageIndex: props.pageIndex,
        style: cardRenderStyle(entry),
        canModify: props.commentActions,
        activate: () => emit("activate", entry.annotation.id),
        setResolved: (resolved: boolean) => {
          if (!props.commentActions) return
          emit("resolve", comment.commentId, resolved)
        },
      } satisfies DocxCommentCardRenderProps,
    }
  }

  const change = entry.annotation.trackedChange
  if (change && props.renderTrackedChangeCard) {
    return {
      render: props.renderTrackedChangeCard,
      renderProps: {
        change,
        kindLabel: changeLabel(change),
        snippet: cardText(entry),
        formattedDate: formatDate(change.date) || undefined,
        accentColor: accentColor(entry),
        documentTheme: props.documentTheme,
        pageIndex: props.pageIndex,
        style: cardRenderStyle(entry),
        canDecide: canDecide(entry),
        activate: () => emit("activate", entry.annotation.id),
        accept: () => {
          if (!canDecide(entry)) return
          emit("decide", entry.annotation.id, "accept")
        },
        reject: () => {
          if (!canDecide(entry)) return
          emit("decide", entry.annotation.id, "reject")
        },
      } satisfies DocxTrackedChangeCardRenderProps,
    }
  }

  return undefined
}

const customCards = computed(() => {
  const cards: Record<string, HostCardEntry> = {}
  for (const entry of visibleCards.value) {
    const card = hostCardFor(entry)
    if (card) cards[entry.annotation.id] = card
  }
  return cards
})

// ── Revision decisions ─────────────────────────────────────────────
// Only tracked-change cards expose decisions; comment cards stay activate-only.
function canDecide(entry: PositionedGutterAnnotation): boolean {
  return props.decidable && Boolean(entry.annotation.trackedChange)
}

// ── Comment actions & inline editor ────────────────────────────────
// Resolving is a thread-level action, so only the root card offers it; every
// comment in the thread can be replied to (and deleted individually).
function isThreadRoot(entry: PositionedGutterAnnotation): boolean {
  const comment = entry.annotation.comment
  return Boolean(comment) && comment?.parentId === undefined
}

function replyComposerFor(entry: PositionedGutterAnnotation): boolean {
  return (
    props.composer?.mode === "reply" &&
    props.composer.commentId !== undefined &&
    props.composer.commentId === entry.annotation.comment?.commentId
  )
}

const composerMode = computed(() => props.composer?.mode ?? "")

/** Draft text belongs to the open editor, so it resets when the editor moves. */
const draft = ref("")
watch(
  () => [props.composer?.mode, props.composer?.commentId] as const,
  () => {
    draft.value = ""
  }
)

function onComposerInput(event: Event): void {
  draft.value = (event.target as HTMLTextAreaElement | null)?.value ?? ""
}

function submitComposer(): void {
  if (draft.value.trim().length === 0) {
    return
  }
  emit("composerSubmit", draft.value)
}

function cancelComposer(): void {
  draft.value = ""
  emit("composerCancel")
}

// ── Pane geometry ──────────────────────────────────────────────────
// Counter-scaled so the pane keeps Word's fixed annotation-pane metrics at
// every zoom level instead of shrinking/growing with the paper.
const drawerOpen = ref(true)

const paneStyle = computed(() =>
  props.drawerMode
    ? annotationGutterDrawerPaneStyle(
        props.zoomFactor,
        props.pageHeightPx,
        drawerOpen.value
      )
    : annotationGutterPaneStyle(props.zoomFactor, props.pageHeightPx)
)

// A closed drawer keeps the tab clickable and hides the cards entirely, so the
// paper underneath stays usable.
const visibleCards = computed(() =>
  props.drawerMode && !drawerOpen.value ? [] : props.positioned
)

function toggleDrawer(): void {
  drawerOpen.value = !drawerOpen.value
}

watch(
  () => props.drawerMode,
  () => {
    drawerOpen.value = true
  }
)

function cardStyle(entry: PositionedGutterAnnotation): Record<string, string> {
  return {
    top: `${annotationGutterCardTopPx(entry.top, props.zoomFactor)}px`,
  }
}

// ── Card content ───────────────────────────────────────────────────
function changeLabel(change: DocxTrackedChange): string {
  switch (change.kind) {
    case "insertion":
      return m.value.docx.trackedChange.inserted
    case "deletion":
      return m.value.docx.trackedChange.deleted
    case "move-from":
      return m.value.docx.trackedChange.movedFrom
    case "move-to":
      return m.value.docx.trackedChange.movedTo
    case "format-change":
      return m.value.docx.trackedChange.format
    case "paragraph-format-change":
      return m.value.docx.trackedChange.paragraphFormatting
    case "table-format-change":
      return m.value.docx.trackedChange.tableFormatting
    case "row-format-change":
      return m.value.docx.trackedChange.rowFormatting
    case "cell-format-change":
      return m.value.docx.trackedChange.cellFormatting
    default:
      return change.kind
  }
}

function accentColor(entry: PositionedGutterAnnotation): string {
  if (entry.annotation.comment) {
    return commentAccentColor(
      props.documentTheme,
      entry.annotation.comment.commentId
    )
  }

  return entry.annotation.trackedChange
    ? trackedChangeAccentColor(entry.annotation.trackedChange.kind, props.documentTheme)
    : "#9ca3af"
}

function cardClasses(entry: PositionedGutterAnnotation): string[] {
  const classes = ["docx-gutter-card--annotation"]
  if (entry.annotation.comment) {
    classes.push("docx-gutter-card--comment")
    if (entry.annotation.comment.resolved) {
      classes.push("docx-gutter-card--resolved")
    }
  }

  const change = entry.annotation.trackedChange
  if (change) {
    switch (change.kind) {
      case "insertion":
        classes.push("docx-gutter-card--insertion")
        break
      case "deletion":
        classes.push("docx-gutter-card--deletion")
        break
      case "move-from":
      case "move-to":
        classes.push("docx-gutter-card--move")
        break
      case "format-change":
      case "paragraph-format-change":
      case "table-format-change":
      case "row-format-change":
      case "cell-format-change":
        classes.push("docx-gutter-card--format")
        break
      default:
        break
    }
  }

  if (props.activeAnnotationId === entry.annotation.id) {
    classes.push("docx-gutter-card--active")
  }

  return classes
}

function cardTypeLabel(entry: PositionedGutterAnnotation): string {
  if (entry.annotation.comment) {
    return m.value.docx.trackedChange.comment
  }

  const change = entry.annotation.trackedChange
  return change ? changeLabel(change) : m.value.docx.trackedChange.annotation
}

function gutterBalloonFlag(entry: PositionedGutterAnnotation): string | undefined {
  const change = entry.annotation.trackedChange
  return change
    ? trackedChangeUsesGutterBalloon(change)
      ? "true"
      : "false"
    : undefined
}

function authorOf(entry: PositionedGutterAnnotation): string {
  return (
    entry.annotation.comment?.author ??
    entry.annotation.trackedChange?.author ??
    ""
  )
}

function cardText(entry: PositionedGutterAnnotation): string {
  return (
    entry.annotation.comment?.text ??
    entry.annotation.trackedChange?.text ??
    ""
  )
}

function dateOf(entry: PositionedGutterAnnotation): string | undefined {
  return (
    entry.annotation.comment?.date ?? entry.annotation.trackedChange?.date
  )
}

function formatDate(dateStr?: string): string {
  if (!dateStr) return ""
  return formatOfficeDate(dateStr, locale.value) || dateStr
}

// ── Card measurement ───────────────────────────────────────────────
// Cards are counter-scaled, so `offsetHeight` (layout px) equals their screen
// height; the layout pipeline converts it back to page space.
const cardElements = new Map<string, HTMLElement>()

function isMeasurableElement(element: unknown): element is HTMLElement {
  return (
    typeof element === "object" &&
    element !== null &&
    typeof (element as HTMLElement).offsetHeight === "number"
  )
}

function registerCardElement(annotationId: string, element: unknown): void {
  if (isMeasurableElement(element)) {
    cardElements.set(annotationId, element)
    return
  }

  cardElements.delete(annotationId)
}

/** A component ref points at the component instance; measure its root. */
function cardElement(element: unknown): unknown {
  if (!element || typeof element !== "object") return element
  if (isMeasurableElement(element)) return element
  return (element as { $el?: unknown }).$el
}

function measureCards(): void {
  for (const [annotationId, element] of cardElements) {
    emit("cardHeight", annotationId, element.offsetHeight)
  }

}

async function measureCardsAfterRender(): Promise<void> {
  await nextTick()
  measureCards()
}

onMounted(measureCardsAfterRender)
watch(() => props.positioned, measureCardsAfterRender, { flush: "post" })
// Host cards can change size without the annotation list changing (a renamed
// action, a longer label), so measurement follows the resolved card entries.
watch(customCards, measureCardsAfterRender, { flush: "post" })

// Activating an annotation from the page (marker or highlighted text) brings
// its card into view inside the pane.
watch(
  () => props.activeAnnotationId,
  async (annotationId) => {
    if (!annotationId) return
    await nextTick()
    cardElements.get(annotationId)?.scrollIntoView?.({ block: "nearest" })
  }
)
</script>

<style scoped>
.docx-tracked-change-gutter {
  box-sizing: border-box;
  /* Default beside-the-paper placement at 100% zoom. The pane counter-scales, so
     `annotationGutterPaneStyle` overrides `left`/`width`/`height`/`transform`
     per zoom level to keep the visible gap and the card text constant. */
  left: calc(100% + 16px);
  overflow-x: hidden;
  overflow-y: auto;
  padding: 0;
  /* Cards are absolutely positioned in screen px by
     `annotationGutterCardTopPx`; the pane only provides the scrolling frame. */
  position: absolute;
  top: 0;
  width: 240px;
  z-index: 8;
}
.docx-gutter-card {
  background: var(--office-card, #f9fafb);
  border: 0;
  border-left: 3px solid #d1d5db;
  border-radius: 4px;
  box-sizing: border-box;
  cursor: pointer;
  font-size: 12px;
  left: 0;
  margin-bottom: 4px;
  padding: 6px 8px;
  position: absolute;
  right: 0;
}
.docx-gutter-card :deep(.office-card__content) {
  padding: 0;
}
.docx-gutter-card:focus-visible {
  outline: 2px solid #2563eb;
  outline-offset: 1px;
}
.docx-gutter-card--active {
  box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.45);
}
.docx-gutter-card--insertion {
  background: #f0fdf4;
  border-left-color: #22c55e;
}
.docx-gutter-card--deletion {
  background: #fef2f2;
  border-left-color: #ef4444;
}
.docx-gutter-card--format {
  background: #eff6ff;
  border-left-color: #3b82f6;
}
.docx-gutter-card--move {
  background: #f0fdf4;
  border-left-color: #70ad47;
}
.docx-gutter-card--comment {
  background: #fffbeb;
  border-left-color: #f59e0b;
}
.docx-gutter-card--resolved {
  opacity: 0.7;
}
.docx-gutter-card-header {
  align-items: baseline;
  display: flex;
  gap: 6px;
  margin-bottom: 2px;
}
.docx-gutter-card-index {
  border-radius: 999px;
  color: #ffffff;
  flex: 0 0 auto;
  font-size: 10px;
  font-weight: 600;
  line-height: 14px;
  min-width: 14px;
  padding: 0 4px;
  text-align: center;
}
.docx-gutter-card-type {
  color: #374151;
  flex: 1 1 auto;
  font-weight: 600;
}
.docx-gutter-card-author {
  color: #6b7280;
  flex: 0 0 auto;
  font-size: 11px;
}
.docx-gutter-card-text {
  color: #4b5563;
  white-space: pre-wrap;
  word-break: break-word;
}
.docx-gutter-card-actions {
  display: flex;
  gap: 4px;
  margin-top: 6px;
}
.docx-gutter-card-action {
  background: #ffffff;
  border: 1px solid #d1d5db;
  border-radius: 3px;
  color: #374151;
  cursor: pointer;
  font-size: 11px;
  line-height: 1.4;
  padding: 1px 6px;
}
.docx-gutter-card-action--accept:hover {
  background: #ecfdf5;
  border-color: #22c55e;
}
.docx-gutter-card-action--reject:hover {
  background: #fef2f2;
  border-color: #ef4444;
}
.docx-gutter-card-action:focus-visible {
  outline: 2px solid #2563eb;
  outline-offset: 1px;
}
.docx-gutter-card-resolved {
  color: #22c55e;
  font-size: 11px;
  margin-top: 4px;
}
.docx-gutter-card--composer {
  cursor: default;
  top: 0;
  z-index: 1;
}
.docx-gutter-card-composer {
  border-top: 1px dashed #d1d5db;
  margin-top: 6px;
  padding-top: 6px;
}
.docx-gutter-card-composer-title {
  color: #374151;
  font-weight: 600;
  margin-bottom: 4px;
}
.docx-gutter-card-composer-input {
  box-sizing: border-box;
  font-family: inherit;
  font-size: 12px;
  min-height: 40px;
  padding: 4px 6px;
  resize: vertical;
  width: 100%;
}
.docx-gutter-card-action--comment:hover {
  background: #fffbeb;
  border-color: #f59e0b;
}
.docx-gutter-card-date {
  color: #9ca3af;
  font-size: 10px;
  margin-top: 2px;
}

/* Drawer mode: an overlay that slides over the paper's right edge. The pane is
   counter-scaled by `annotationGutterDrawerPaneStyle`, so these sizes are
   screen px at every zoom level. */
.docx-tracked-change-gutter--drawer {
  background: rgba(255, 255, 255, 0.97);
  border-left: 1px solid #e5e7eb;
  box-shadow: -4px 0 16px rgba(15, 23, 42, 0.12);
  overflow-x: hidden;
  overflow-y: auto;
  padding: 28px 8px 8px 6px;
  pointer-events: auto;
}
.docx-tracked-change-gutter--drawer-closed {
  background: transparent;
  border-left: none;
  box-shadow: none;
  overflow: visible;
  padding: 0;
}
.docx-gutter-drawer-tab {
  align-items: center;
  background: #1f2937;
  border: none;
  border-radius: 4px 0 0 4px;
  color: #ffffff;
  cursor: pointer;
  display: flex;
  font-family: inherit;
  font-size: 11px;
  gap: 4px;
  justify-content: center;
  padding: 4px 6px;
  position: absolute;
  right: 0;
  top: 8px;
  z-index: 2;
}
.docx-gutter-drawer-tab:hover {
  background: #374151;
}
.docx-gutter-drawer-tab:focus-visible {
  outline: 2px solid #2563eb;
  outline-offset: 1px;
}
.docx-gutter-drawer-tab-count {
  background: rgba(255, 255, 255, 0.22);
  border-radius: 999px;
  font-weight: 600;
  min-width: 14px;
  padding: 0 4px;
  text-align: center;
}
.docx-gutter-drawer-tab-label {
  font-weight: 600;
}
/*
 * Keep the gutter beside the paper at every viewport width. The page frame is
 * exactly one paper tall and clips the page surface, so flowing the gutter
 * below the paper would overlap the next page; DocxViewerRoot reserves the
 * horizontal space for this gutter instead.
 */
</style>
