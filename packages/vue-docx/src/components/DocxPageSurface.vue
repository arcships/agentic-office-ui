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
      >
        <div
          v-for="line of borderEdgeLines"
          :key="line.edge"
          :data-docx-page-border-edge="line.edge"
          class="docx-page-border-edge"
          :style="line.style"
        />
      </div>

      <!-- Body -->
      <DocxPageBody
        :page-index="pageIndex"
        :page-node-segments="pageNodeSegments"
        :page-columns="pageColumns"
        :column-layout="columnLayout"
        :page-layout="pageLayout"
        :page-content-width-px="pageContentWidthPx"
        :floating-page-origin-px="floatingPageOriginPx"
        :model="model"
        :controller="controller"
        :editable="editable"
        :theme="theme"
        :page-number="pageNumber"
        :total-pages="totalPages"
        :page-number-format="pageNumberFormat"
        :section-pages="resolvedSectionPages"
        :resolve-style-ref-field-value="fieldResolvers.resolveStyleRefFieldValue"
        :resolve-bookmark-field-value="fieldResolvers.resolveBookmarkFieldValue"
        :resolve-bookmark-page-number="fieldResolvers.resolveBookmarkPageNumber"
        :tracked-changes-enabled="trackedChangesEnabled"
        :comments-enabled="commentsEnabled"
        :note-marker-labels="noteMarkerLabels"
        :cover-image-child-indexes-by-node="coverImageChildIndexesByNode"
        :header-reserve-px="pageFurnitureReservePx.header"
        :footer-reserve-px="pageFurnitureReservePx.footer"
        @paragraph-measure="(nodeIndex: number, heightPx: number) => emit('paragraphMeasure', nodeIndex, heightPx)"
        @table-row-measure="(nodeIndex: number, rowHeightsPx: number[]) => emit('tableRowMeasure', nodeIndex, rowHeightsPx)"
        @content-measure="(heightPx: number) => emit('contentMeasure', heightPx)"
      >
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
      </DocxPageBody>

      <template v-if="hasHeader">
        <DocxPageHeader
          :page-index="pageIndex"
          :section="headerSection"
          :page-layout="pageLayout"
          :page-content-width-px="pageContentWidthPx"
          :floating-page-origin-px="floatingPageOriginPx"
          :theme="theme"
          :model="model"
          :controller="controller"
          :page-number="pageNumber"
          :total-pages="totalPages"
          :page-number-format="pageNumberFormat"
          :section-pages="resolvedSectionPages"
          :page-start-node-index="pageStartNodeIndex"
          :resolve-style-ref-field-value="fieldResolvers.resolveStyleRefFieldValue"
          :resolve-bookmark-field-value="fieldResolvers.resolveBookmarkFieldValue"
          :resolve-bookmark-page-number="fieldResolvers.resolveBookmarkPageNumber"
          :tracked-changes-enabled="trackedChangesEnabled"
          :comments-enabled="commentsEnabled"
          :header-distance-px="pageLayout.headerDistancePx"
        />
      </template>
      <template v-if="hasFooter">
        <DocxPageFooter
          :page-index="pageIndex"
          :section="footerSection"
          :page-layout="pageLayout"
          :page-content-width-px="pageContentWidthPx"
          :floating-page-origin-px="floatingPageOriginPx"
          :theme="theme"
          :model="model"
          :controller="controller"
          :page-number="pageNumber"
          :total-pages="totalPages"
          :page-number-format="pageNumberFormat"
          :section-pages="resolvedSectionPages"
          :page-start-node-index="pageStartNodeIndex"
          :resolve-style-ref-field-value="fieldResolvers.resolveStyleRefFieldValue"
          :resolve-bookmark-field-value="fieldResolvers.resolveBookmarkFieldValue"
          :resolve-bookmark-page-number="fieldResolvers.resolveBookmarkPageNumber"
          :tracked-changes-enabled="trackedChangesEnabled"
          :comments-enabled="commentsEnabled"
          :footer-distance-px="pageLayout.footerDistancePx"
        />
      </template>

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
        :skip-image-keys="coverImageKeys"
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
      :drawer-mode="annotationDrawerMode"
      :decidable="editable && Boolean(controller)"
      @activate="activateAnnotation($event, { scrollToAnchor: true })"
      @card-height="reportGutterCardHeight"
      @decide="onGutterDecide"
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
  DocumentPageBorders,
  FooterSection,
  HeaderSection,
} from "@arcships/docx-core"
import {
  buildPageGutterAnnotations,
  docModelThumbnailMetadataSignature,
  docNodeContentSignature,
  findGutterAnnotationAnchorElementInPage,
  fullPageCoverAbsoluteFloatingImageStyle,
  isLikelyFullPageCoverFloatingImage,
  estimateDocNodeHeightPx,
  parseSectionPageBorders,
  resolveDocumentBackgroundColor,
  resolveDocumentLayout,
  resolveDocumentSectionPropertiesXml,
  resolveFooterPaginationReservePx,
  resolveFloatingPageOriginPx,
  resolveHeaderPaginationReservePx,
  resolveRenderableImageSource,
  pageMarginPaddingStyle,
  pagePatternPaintStyle,
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
import { createDocxFieldResolvers } from "../composables/docx-field-resolvers"
import {
  ensureDocxViewerPageSurfaceRegistry,
  notifyDocxViewerPageSurfaceListeners,
} from "../composables/page-surface-registry"
import { collectPageFormFields } from "../composables/docx-form-fields"

// ── Types ──────────────────────────────────────────────────────────
interface DocxPageSurfaceProps {
  pageIndex: number
  pageLayout: ReturnType<typeof resolveDocumentLayout>
  pageNodeSegments: DocumentPageNodeSegment[]
  /** Per-body-column segments; omitted for single-column pages. */
  pageColumns?: DocumentPageNodeSegment[][]
  columnLayout?: { count: number; gapPx: number; widthsPx: number[] }
  model: DocModel
  controller?: DocxEditorController
  editable?: boolean
  /** Zoom factor applied to the page subtree; the annotation pane counter-scales. */
  zoomFactor?: number
  /** Render the annotation pane as an overlay drawer (narrow viewports). */
  annotationDrawerMode?: boolean
  pageWidthPx: number
  pageContentWidthPx: number
  pageNumber: number
  totalPages: number
  sectionPages?: number
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
  pageBackgroundColor?: string
  pageBorders?: DocumentPageBorders | null
  theme?: "light" | "dark"
}

interface CoverEntry {
  key: string
  src: string
  style: Record<string, string | number | undefined>
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
    annotationDrawerMode: false,
  }
)

const emit = defineEmits<{
  measure: [event: { heightPx: number }]
  paragraphMeasure: [nodeIndex: number, heightPx: number]
  tableRowMeasure: [nodeIndex: number, rowHeightsPx: number[]]
  contentMeasure: [heightPx: number]
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

function unpublishPageSurface(notify = true): void {
  if (!publishedSurface) return
  const { controller, element, pageIndex } = publishedSurface
  publishedSurface = undefined
  const registry = ensureDocxViewerPageSurfaceRegistry(controller)
  if (registry.pageElements.get(pageIndex) === element) {
    registry.pageElements.delete(pageIndex)
    registry.pageContentKeys.delete(pageIndex)
    registry.pageSizes.delete(pageIndex)
    if (notify) {
      notifyDocxViewerPageSurfaceListeners(controller, {
        pageIndex,
        kind: "unpublish",
      })
    }
  }
}

function publishPageSurface(): void {
  const controller = props.controller
  const element = pageSurfaceRef.value
  if (!controller || !element) {
    unpublishPageSurface()
    return
  }

  // Replacing a surface for the same page is one publish, not an unpublish
  // notification followed by a publish notification.
  const registry = ensureDocxViewerPageSurfaceRegistry(controller)
  const previousElement = registry.pageElements.get(props.pageIndex)
  const previousContentKey = registry.pageContentKeys.get(props.pageIndex)
  const previousSize = registry.pageSizes.get(props.pageIndex)
  unpublishPageSurface(false)
  const pageSize = {
    widthPx: props.pageWidthPx,
    heightPx: props.pageLayout.pageHeightPx,
  }
  const contentKey = [
    controller.documentLoadNonce,
    docModelThumbnailMetadataSignature(props.model.metadata),
    ...props.pageNodeSegments.map((segment) =>
      docNodeContentSignature(props.model.nodes[segment.nodeIndex])
    ),
  ].join(":")

  registry.pageElements.set(props.pageIndex, element)
  registry.pageSizes.set(props.pageIndex, pageSize)
  registry.pageContentKeys.set(props.pageIndex, contentKey)
  publishedSurface = { controller, element, pageIndex: props.pageIndex }

  const changed =
    previousElement !== element ||
    previousContentKey !== contentKey ||
    previousSize?.widthPx !== pageSize.widthPx ||
    previousSize?.heightPx !== pageSize.heightPx
  if (changed) {
    notifyDocxViewerPageSurfaceListeners(controller, {
      pageIndex: props.pageIndex,
      kind: "publish",
      contentKey,
    })
  }
}

// ── Computed ───────────────────────────────────────────────────────
const hasHeader = computed(() => (props.headerSection?.nodes.length ?? 0) > 0)
const hasFooter = computed(() => (props.footerSection?.nodes.length ?? 0) > 0)
const resolvedSectionPages = computed(() => props.sectionPages ?? props.totalPages)
const pageStartNodeIndex = computed(() => props.pageNodeSegments[0]?.nodeIndex)
const fieldResolvers = computed(() =>
  createDocxFieldResolvers(props.model, {
    pageNumberByNodeIndex: new Map(
      props.pageNodeSegments.map((segment) => [segment.nodeIndex, props.pageNumber]),
    ),
  }),
)

// Extra body padding needed when a tall header/footer grows past the nominal
// margin. The helpers return the overlap beyond the margin, so the body area
// shrinks instead of the header/footer pushing the paper taller.
const pageFurnitureReservePx = computed(() => {
  const reserveOptions = {
    estimateDocNodeHeightPx,
    minimumVisibleFlowFooterReservePx: 0,
  }
  return {
    header: hasHeader.value
      ? resolveHeaderPaginationReservePx(
          [props.headerSection as HeaderSection],
          props.pageLayout,
          reserveOptions
        )
      : 0,
    footer: hasFooter.value
      ? resolveFooterPaginationReservePx(
          [props.footerSection as FooterSection],
          props.pageLayout,
          reserveOptions
        )
      : 0,
  }
})

// Origin for margin/page/column anchored floating images. The surface is the
// positioned ancestor of the page body, and its padding box starts at the page
// edge, so `page` anchors are measured from 0/0 and margin anchors add the
// page margins.
const floatingPageOriginPx = computed(() =>
  resolveFloatingPageOriginPx(props.pageWidthPx, props.pageLayout.marginsPx)
)

// The layer and its mount gate share one collector: table-cell fields are
// found recursively and a table split across pages only reports the fields in
// the rows rendered on this page.
const hasFormFields = computed(
  () => collectPageFormFields(props.model, props.pageNodeSegments).length > 0
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

/**
 * Accept or reject the revision behind a gutter card. The gutter only reports
 * the annotation id; the owning controller decides how the model is rewritten.
 */
function onGutterDecide(
  annotationId: string,
  decision: "accept" | "reject"
): void {
  const controller = props.controller
  const annotation = pageAnnotations.value.find(
    (entry) => entry.id === annotationId
  )
  const changeId = annotation?.trackedChange?.id
  if (!controller || !changeId) return

  if (decision === "accept") {
    controller.acceptTrackedChange(changeId)
    return
  }
  controller.rejectTrackedChange(changeId)
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
  const entries: CoverEntry[] = []
  const seen = new Set<string>()
  for (const segment of props.pageNodeSegments) {
    const node = props.model.nodes[segment.nodeIndex]
    if (!node || node.type !== "paragraph") continue
    node.children.forEach((child, childIndex) => {
      if (child.type !== "image") return
      if (
        child.floating?.behindDocument !== true ||
        !isLikelyFullPageCoverFloatingImage(
          child,
          props.pageWidthPx,
          props.pageLayout.pageHeightPx
        )
      ) {
        return
      }
      const src = resolveRenderableImageSource(child)
      if (!src) return
      const key = `${segment.nodeIndex}:${childIndex}`
      if (seen.has(key)) return
      seen.add(key)
      entries.push({
        key,
        src,
        style: fullPageCoverAbsoluteFloatingImageStyle(
          child,
          props.pageLayout,
          { anchorToPageSurface: true }
        ),
      })
    })
  }
  return entries
})

const coverImageKeys = computed(() => coverEntries.value.map((entry) => entry.key))

const coverImageChildIndexesByNode = computed<Record<number, number[]>>(() => {
  const indexes: Record<number, number[]> = {}
  for (const key of coverImageKeys.value) {
    const [nodeIndex, childIndex] = key.split(":").map(Number)
    if (!Number.isFinite(nodeIndex) || !Number.isFinite(childIndex)) continue
    if (!indexes[nodeIndex]) indexes[nodeIndex] = []
    indexes[nodeIndex].push(childIndex)
  }
  return indexes
})

type PageBorderEdgeName = "top" | "right" | "bottom" | "left"

interface PageBorderEdgeLine {
  edge: PageBorderEdgeName
  style: Record<string, string>
}

const resolvedPageBorders = computed<DocumentPageBorders | undefined>(() => {
  if (props.pageBorders !== undefined) {
    return props.pageBorders ?? undefined
  }
  return parseSectionPageBorders(
    resolveDocumentSectionPropertiesXml(props.model)
  )
})

const borderStyle = computed<Record<string, string> | undefined>(() => {
  if (!resolvedPageBorders.value) return undefined
  return {
    position: "absolute",
    inset: "0",
    pointerEvents: "none",
    zIndex: "5",
  }
})

const borderEdgeLines = computed<PageBorderEdgeLine[]>(() => {
  const borders = resolvedPageBorders.value
  if (!borders) return []
  const lines: PageBorderEdgeLine[] = []
  for (const edge of ["top", "right", "bottom", "left"] as const) {
    const border = borders[edge]
    if (!border?.cssBorder) continue
    const offset = `${border.offsetPx ?? 0}px`
    const style: Record<string, string> = { position: "absolute" }
    if (edge === "top" || edge === "bottom") {
      style.left = "0"
      style.right = "0"
      style.height = "0"
      style[edge] = offset
      style[edge === "top" ? "borderTop" : "borderBottom"] = border.cssBorder
    } else {
      style.top = "0"
      style.bottom = "0"
      style.width = "0"
      style[edge] = offset
      style[edge === "left" ? "borderLeft" : "borderRight"] = border.cssBorder
    }
    if (border.shadow) {
      style.color = border.color ?? "#000000"
      style.filter = "drop-shadow(1px 1px 0 currentColor)"
    }
    lines.push({ edge, style })
  }
  return lines
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
const pagePatternStyle = computed(() =>
  pagePatternPaintStyle(props.model.metadata.documentBackgroundPattern)
)

const surfaceStyle = computed(() => ({
  width: `${props.pageWidthPx}px`,
  height: `${props.pageLayout.pageHeightPx}px`,
  overflow: "hidden" as const,
  backgroundColor:
    props.pageBackgroundColor ??
    pagePatternStyle.value.backgroundColor ??
    resolveDocumentBackgroundColor(props.model.metadata.documentBackgroundColor) ??
    (props.theme === "dark" ? "#1f2937" : "#ffffff"),
  backgroundImage: pagePatternStyle.value.backgroundImage,
  backgroundSize: pagePatternStyle.value.backgroundSize,
  backgroundRepeat: pagePatternStyle.value.backgroundRepeat,
  color: props.theme === "dark" ? "#f9fafb" : "#111827",
  position: "relative" as const,
  // Own stacking context: the behind-text image layer uses a negative z-index so
  // it paints under the page text but still above the page background.
  isolation: "isolate" as const,
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
    props.pageNodeSegments,
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
  z-index: -1;
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

@media print {
  /* Paper must stay printable when the viewer is in dark mode. The inline
     surface style carries the on-screen theme, so print needs !important. */
  .docx-page-surface {
    background: #fff !important;
    border: 0 !important;
    box-shadow: none !important;
    color: #000 !important;
  }
  .docx-page-border-overlay,
  .docx-tracked-change-gutter { display: none !important; }
}
</style>
