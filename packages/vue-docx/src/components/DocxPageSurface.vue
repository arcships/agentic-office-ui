<template>
  <div class="docx-page-frame" :style="frameStyle">
    <div
      ref="pageSurfaceRef"
      data-docx-page-surface="true"
      :data-docx-page-paper-height="pageLayout.pageHeightPx"
      class="docx-page-surface"
      :style="surfaceStyle"
      @click="onSurfaceClick"
    >
      <!-- Cover layer (full-page background images) -->
      <div
        v-if="coverEntries.length > 0"
        data-docx-page-cover-layer="true"
        class="docx-page-cover-layer"
      >
        <img
          v-for="cover of coverEntries"
          :key="cover.key"
          :src="cover.src"
          alt=""
          aria-hidden="true"
          draggable="false"
          class="docx-page-cover-image"
          :style="cover.style"
        />
      </div>

      <!-- Border overlay -->
      <div
        v-if="borderStyle"
        data-docx-page-border-overlay="true"
        class="docx-page-border-overlay"
        :style="borderStyle"
      />

      <!-- Body -->
      <DocxPageBody
        :page-index="pageIndex"
        :page-node-segments="pageNodeSegments"
        :page-layout="pageLayout"
        :page-content-width-px="pageContentWidthPx"
        :model="model"
        :controller="controller"
        :editable="editable"
        :theme="theme"
        :page-number="pageNumber"
        :total-pages="totalPages"
        :page-number-format="pageNumberFormat"
        :tracked-changes-enabled="trackedChangesEnabled"
        :comments-enabled="commentsEnabled"
        :note-marker-labels="noteMarkerLabels"
      >
        <template v-if="hasHeader" #header>
          <DocxPageHeader
            :page-index="pageIndex"
            :section="headerSection"
            :page-layout="pageLayout"
            :page-content-width-px="pageContentWidthPx"
            :theme="theme"
            :model="model"
            :controller="controller"
            :page-number="pageNumber"
            :total-pages="totalPages"
            :page-number-format="pageNumberFormat"
            :tracked-changes-enabled="trackedChangesEnabled"
            :comments-enabled="commentsEnabled"
          />
        </template>
        <template v-if="footnotes.length" #footnotes>
          <DocxNotesSection
            kind="footnote"
            :notes="footnotes"
            :labels="footnoteLabels"
            :numbering-definitions="model.metadata.numberingDefinitions"
            :theme="theme"
          />
        </template>
        <template v-if="endnotes.length" #endnotes>
          <DocxNotesSection
            kind="endnote"
            :notes="endnotes"
            :labels="endnoteLabels"
            :numbering-definitions="model.metadata.numberingDefinitions"
            :theme="theme"
          />
        </template>
        <template v-if="hasFooter" #footer>
          <DocxPageFooter
            :page-index="pageIndex"
            :section="footerSection"
            :page-layout="pageLayout"
            :page-content-width-px="pageContentWidthPx"
            :theme="theme"
            :model="model"
            :controller="controller"
            :page-number="pageNumber"
            :total-pages="totalPages"
            :page-number-format="pageNumberFormat"
            :tracked-changes-enabled="trackedChangesEnabled"
            :comments-enabled="commentsEnabled"
          />
        </template>
      </DocxPageBody>

      <!-- Floating image layer -->
      <DocxImageLayer
        v-if="editable && controller"
        :page-index="pageIndex"
        :page-width-px="pageWidthPx"
        :page-height-px="pageLayout.pageHeightPx"
        :controller="controller"
        :page-node-segments="pageNodeSegments"
        :editable="editable"
        :zoom-factor="zoomFactor"
      />

      <!-- Form field layer -->
      <DocxFormFieldLayer
        v-if="editable && controller && hasFormFields"
        :page-index="pageIndex"
        :controller="controller"
        :page-node-segments="pageNodeSegments"
      />
    </div>

    <!--
      Tracked change / comment gutter lives beside the paper, so it must stay
      outside the clipped page surface or it would disappear when annotations
      are enabled.
    -->

    <!--
      Anchor markers and leader lines sit in the page frame as well: they use
      page-space coordinates but must be able to draw past the paper edge,
      which the clipped surface would cut off.
    -->
    <DocxGutterAnchorMarkers
      v-if="showGutter"
      :positioned="positionedAnnotations"
      :zoom-factor="annotationZoomFactor"
      :page-width-px="pageWidthPx"
      :content-right-px="pageContentRightPx"
      :document-theme="theme"
      :active-annotation-id="activeAnnotationId"
      @activate="activateAnnotation($event, { scrollToAnchor: false })"
    />
    <DocxTrackedChangeGutter
      v-if="showGutter"
      :positioned="positionedAnnotations"
      :zoom-factor="annotationZoomFactor"
      :page-height-px="annotationPageHeightPx"
      :document-theme="theme"
      :active-annotation-id="activeAnnotationId"
      @activate="activateAnnotation($event, { scrollToAnchor: true })"
      @card-height="reportGutterCardHeight"
    />
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, onUnmounted, watch } from "vue"
import type {
  DocModel,
  DocxComment,
  DocxEditorController,
  DocxTrackedChange,
  DocumentNoteDefinition,
  DocumentPageNodeSegment,
  FooterSection,
  HeaderSection,
  TableCellContentNode,
} from "@arcships/docx-core"
import {
  buildPageGutterAnnotations,
  docModelThumbnailMetadataSignature,
  docNodeContentSignature,
  findGutterAnnotationAnchorElementInPage,
  resolveDocumentLayout,
  pageMarginPaddingStyle,
} from "@arcships/docx-core"
import DocxPageHeader from "./DocxPageHeader.vue"
import DocxPageFooter from "./DocxPageFooter.vue"
import DocxPageBody from "./DocxPageBody"
import DocxImageLayer from "./DocxImageLayer.vue"
import DocxFormFieldLayer from "./DocxFormFieldLayer.vue"
import DocxTrackedChangeGutter from "./DocxTrackedChangeGutter.vue"
import DocxGutterAnchorMarkers from "./DocxGutterAnchorMarkers.vue"
import DocxNotesSection from "./DocxNotesSection.vue"
import { useDocxGutterLayout } from "../composables/gutter-layout"
import {
  ensureDocxViewerPageSurfaceRegistry,
  notifyDocxViewerPageSurfaceListeners,
} from "../composables/page-surface-registry"

// ── Types ──────────────────────────────────────────────────────────
interface DocxPageSurfaceProps {
  pageIndex: number
  pageLayout: ReturnType<typeof resolveDocumentLayout>
  pageNodeSegments: DocumentPageNodeSegment[]
  model: DocModel
  controller?: DocxEditorController
  editable?: boolean
  /** Zoom factor applied to the page subtree; the annotation pane counter-scales. */
  zoomFactor?: number
  pageWidthPx: number
  pageContentWidthPx: number
  pageNumber: number
  totalPages: number
  pageNumberFormat?: string
  headerSection?: HeaderSection
  footerSection?: FooterSection
  trackedChangesEnabled?: boolean
  commentsEnabled?: boolean
  trackedChanges?: readonly DocxTrackedChange[]
  comments?: readonly DocxComment[]
  footnotes?: DocumentNoteDefinition[]
  endnotes?: DocumentNoteDefinition[]
  /** Note id → marker label, so note bodies match their in-text markers. */
  footnoteLabels?: Map<number, string>
  endnoteLabels?: Map<number, string>
  noteMarkerLabels?: { footnote?: Map<number, string>; endnote?: Map<number, string> }
  theme?: "light" | "dark"
}

interface CoverEntry {
  key: string
  src: string
  style: Record<string, string>
}

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<DocxPageSurfaceProps>(),
  {
    editable: true,
    trackedChangesEnabled: false,
    commentsEnabled: false,
    footnotes: () => [],
    endnotes: () => [],
    theme: "light" as const,
    zoomFactor: 1,
  }
)

const emit = defineEmits<{
  measure: [event: { heightPx: number }]
}>()

// ── Refs ───────────────────────────────────────────────────────────
const pageSurfaceRef = ref<HTMLDivElement | undefined>()

let publishedSurface:
  | {
      controller: DocxEditorController
      element: HTMLDivElement
      pageIndex: number
    }
  | undefined

function unpublishPageSurface(): void {
  if (!publishedSurface) return
  const { controller, element, pageIndex } = publishedSurface
  const registry = ensureDocxViewerPageSurfaceRegistry(controller)
  if (registry.pageElements.get(pageIndex) === element) {
    registry.pageElements.delete(pageIndex)
    registry.pageContentKeys.delete(pageIndex)
    registry.pageSizes.delete(pageIndex)
    registry.pageThumbnailSnapshots.delete(pageIndex)
    notifyDocxViewerPageSurfaceListeners(controller)
  }
  publishedSurface = undefined
}

function publishPageSurface(): void {
  unpublishPageSurface()
  const controller = props.controller
  const element = pageSurfaceRef.value
  if (!controller || !element) return

  const registry = ensureDocxViewerPageSurfaceRegistry(controller)
  registry.pageElements.set(props.pageIndex, element)
  registry.pageSizes.set(props.pageIndex, {
    widthPx: props.pageWidthPx,
    heightPx: props.pageLayout.pageHeightPx,
  })
  registry.pageContentKeys.set(
    props.pageIndex,
    [
      controller.documentLoadNonce,
      docModelThumbnailMetadataSignature(props.model.metadata),
      ...props.pageNodeSegments.map((segment) =>
        docNodeContentSignature(props.model.nodes[segment.nodeIndex])
      ),
    ].join(":")
  )
  publishedSurface = { controller, element, pageIndex: props.pageIndex }
  notifyDocxViewerPageSurfaceListeners(controller)
}

// ── Computed ───────────────────────────────────────────────────────
const hasHeader = computed(() => (props.headerSection?.nodes.length ?? 0) > 0)
const hasFooter = computed(() => (props.footerSection?.nodes.length ?? 0) > 0)

function nodesContainFormField(nodes: TableCellContentNode[]): boolean {
  return nodes.some((node) => {
    if (node.type === "paragraph") {
      return node.children.some((child) => child.type === "form-field")
    }
    return node.rows.some((row) =>
      row.cells.some((cell) => nodesContainFormField(cell.nodes))
    )
  })
}

const hasFormFields = computed(() =>
  props.pageNodeSegments.some((segment) => {
    const node = props.model.nodes[segment.nodeIndex]
    if (!node) return false
    if (node.type === "paragraph") {
      return node.children.some((child) => child.type === "form-field")
    }
    return node.rows.some((row) =>
      row.cells.some((cell) => nodesContainFormField(cell.nodes))
    )
  })
)

const showGutter = computed(
  () => props.trackedChangesEnabled || props.commentsEnabled
)

// ── Annotation gutter ──────────────────────────────────────────────
// Annotations are numbered and ordered in document order, then placed against
// their page-space anchors: the DOM resolver measures each anchor inside the
// page (independent of zoom) and the pure layout stacks the cards from there.
const pageAnnotations = computed(() =>
  buildPageGutterAnnotations({
    comments: props.comments ?? props.controller?.comments ?? [],
    trackedChanges: props.trackedChanges ?? props.controller?.trackedChanges ?? [],
    pageSegments: props.pageNodeSegments,
    showComments: props.commentsEnabled,
    showTrackedChanges: props.trackedChangesEnabled,
    headerPartName: props.headerSection?.partName,
    footerPartName: props.footerSection?.partName,
  })
)

const pageMeasuredHeightPx = ref(0)

const {
  positioned: positionedAnnotations,
  zoomFactor: annotationZoomFactor,
  reportCardHeight: reportGutterCardHeight,
  refresh: refreshGutterLayout,
} = useDocxGutterLayout({
  pageElement: pageSurfaceRef,
  annotations: () => pageAnnotations.value,
  layoutSignal: () =>
    [props.model, props.pageNodeSegments, props.pageWidthPx, props.theme] as const,
  pageWidthPx: () => props.pageWidthPx,
  pageHeightPx: () =>
    Math.max(props.pageLayout.pageHeightPx, pageMeasuredHeightPx.value),
  zoomFactor: () => props.zoomFactor,
})

// Height of the annotation surface in page space; cards clamp and stack
// against it so they never land outside the paper.
const annotationPageHeightPx = computed(() =>
  Math.max(props.pageLayout.pageHeightPx, pageMeasuredHeightPx.value)
)

const pageContentRightPx = computed(
  () =>
    props.pageWidthPx -
    (props.pageLayout.marginsPx?.right ?? 72)
)

const activeAnnotationId = ref<string | undefined>()

function annotationIdForTarget(target: EventTarget | null): string | undefined {
  const element = (target as HTMLElement | null)?.closest?.(
    "[data-docx-comment-ids], [data-docx-comment-id], [data-docx-tracked-change-id]"
  ) as HTMLElement | null
  if (!element) return undefined

  const dataset = element.dataset ?? {}
  const commentIds = (dataset.docxCommentIds ?? dataset.docxCommentId ?? "")
    .split(/\s+/)
    .filter(Boolean)
  for (const commentId of commentIds) {
    const annotation = pageAnnotations.value.find(
      (entry) => entry.comment && String(entry.comment.commentId) === commentId
    )
    if (annotation) return annotation.id
  }

  const changeId = dataset.docxTrackedChangeId
  if (changeId) {
    const annotation = pageAnnotations.value.find(
      (entry) =>
        entry.trackedChange &&
        (entry.trackedChange.id === changeId ||
          entry.trackedChange.inlineAnchorId === changeId)
    )
    if (annotation) return annotation.id
  }

  return undefined
}

function activateAnnotation(
  annotationId: string,
  options?: { scrollToAnchor?: boolean }
): void {
  activeAnnotationId.value = annotationId
  if (!options?.scrollToAnchor) return

  const pageElement = pageSurfaceRef.value
  const annotation = pageAnnotations.value.find(
    (entry) => entry.id === annotationId
  )
  if (!pageElement || !annotation) return

  const anchorElement = findGutterAnnotationAnchorElementInPage(
    pageElement,
    annotation
  )
  anchorElement?.scrollIntoView?.({ block: "center", inline: "nearest" })
}

const coverEntries = computed<CoverEntry[]>(() => {
  // Collect full-page cover background images from paragraph segments
  // Simplified: empty for now — full implementation requires image run detection
  return []
})

const borderStyle = computed<Record<string, string> | undefined>(() => {
  // Page border overlay — simplified: no border by default
  return undefined
})

const frameStyle = computed(() => ({
  position: "relative" as const,
  width: `${props.pageWidthPx}px`,
  height: `${props.pageLayout.pageHeightPx}px`,
  flex: "0 0 auto" as const,
}))

// A page is exactly one paper tall: any content pagination mis-assigned to it
// is clipped at the paper edge instead of growing the box into the next page.
//
// TODO(#95 / #25 / #47): clipping is the rendering safety net, not the fix for
// content assignment. Pagination must still move overflowing paragraphs to the
// following page; until it does, clipped content is intentionally hidden.
const surfaceStyle = computed(() => ({
  width: `${props.pageWidthPx}px`,
  height: `${props.pageLayout.pageHeightPx}px`,
  overflow: "hidden" as const,
  backgroundColor: props.theme === "dark" ? "#1f2937" : "#ffffff",
  color: props.theme === "dark" ? "#f9fafb" : "#111827",
  position: "relative" as const,
  boxSizing: "border-box" as const,
  // Draw the paper outline inside the box so the content area keeps exactly
  // the page size pagination lays out (border-box height minus margins).
  boxShadow: "inset 0 0 0 1px #d4d4d4, 0 8px 24px rgba(0, 0, 0, 0.08)",
  ...pageMarginPaddingStyle(props.pageLayout.marginsPx ?? { top: 72, bottom: 72, left: 72, right: 72 }),
}))

// ── Surface click → select paragraph / focus annotation ────────────
function onSurfaceClick(event: MouseEvent): void {
  // Clicks on a commented / revised run link the text to its gutter card.
  const annotationId = annotationIdForTarget(event.target)
  if (annotationId) {
    activateAnnotation(annotationId)
  }
}

// ── Measure page height ────────────────────────────────────────────
let resizeObserver: ResizeObserver | null = null

function measurePage(): void {
  if (!pageSurfaceRef.value) return
  // offsetHeight is layout height, unaffected by ancestor CSS transforms.
  // getBoundingClientRect().height would already include scale(zoomFactor),
  // causing a double scale when the viewer multiplies by zoomFactor again.
  // The surface has a fixed paper height, so this measurement stays stable
  // even when the page body has clipped overflowing content.
  const measuredHeight = pageSurfaceRef.value.offsetHeight
  const heightPx = measuredHeight > 0 ? measuredHeight : props.pageLayout.pageHeightPx
  if (heightPx !== pageMeasuredHeightPx.value) {
    pageMeasuredHeightPx.value = heightPx
    // Paragraph anchors move when content reflows; re-measure them.
    refreshGutterLayout()
  }
  emit("measure", { heightPx })
}

onMounted(() => {
  publishPageSurface()
  const ResizeObserverConstructor = globalThis.ResizeObserver
  if (pageSurfaceRef.value && typeof ResizeObserverConstructor === "function") {
    resizeObserver = new ResizeObserverConstructor(() => {
      measurePage()
    })
    resizeObserver.observe(pageSurfaceRef.value)
    measurePage()
  }
})

onUnmounted(() => {
  unpublishPageSurface()
  if (resizeObserver) {
    resizeObserver.disconnect()
    resizeObserver = null
  }
})

watch(
  () => [
    props.controller,
    props.controller?.documentLoadNonce,
    props.model,
    props.pageIndex,
    props.pageWidthPx,
    props.pageLayout.pageHeightPx,
  ] as const,
  () => publishPageSurface(),
  { flush: "post" }
)
</script>

<style scoped>
.docx-page-surface {
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
}
/*
 * The page body shrinks to the content area and clips mispaginated content.
 * Its direct children must keep their natural height so overflow is clipped at
 * the body edge instead of compressing paragraphs on top of each other.
 */
.docx-page-surface :deep(.docx-page-body > *) {
  flex-shrink: 0;
}
.docx-page-cover-layer {
  position: absolute;
  left: 0;
  top: 0;
  right: 0;
  bottom: 0;
  overflow: clip;
  pointer-events: none;
  z-index: 0;
}
.docx-page-cover-image {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  display: block;
  object-fit: cover;
}
.docx-page-border-overlay {
  position: absolute;
  inset: 0;
  pointer-events: none;
  z-index: 5;
}
</style>
