<template>
  <div
    class="docx-page-frame"
    :style="frameStyle"
    :data-docx-page-label="pageLabel ?? pageNumber"
    :data-docx-page-parity-blank="parityBlankPage ? 'true' : 'false'"
  >
    <div
      ref="pageSurfaceRef"
      data-docx-page-surface="true"
      :data-docx-page-paper-width="pageLayout.pageWidthPx"
      :data-docx-page-paper-height="pageLayout.pageHeightPx"
      class="docx-page-surface"
      :style="surfaceStyle"
      @click="onSurfaceClick"
    >
      <!-- Cover layer (full-page background images) -->
      <div
        v-if="coverEntries.length > 0"
        data-docx-page-cover-layer="true"
        :data-docx-pagination-ignore="true"
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
        :data-docx-pagination-ignore="true"
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
        :zoom-factor="zoomFactor"
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
        :revision-display="renderRevisionDisplay"
        :comments-enabled="commentsEnabled"
        :note-marker-labels="noteMarkerLabels"
        :cover-image-child-indexes-by-node="coverImageChildIndexesByNode"
        :header-reserve-px="pageFurnitureReservePx.header"
        :footer-reserve-px="pageFurnitureReservePx.footer"
        :on-internal-link-click="onInternalLinkClick"
        :bookmark-target="bookmarkTarget"
        @paragraph-measure="(nodeIndex: number, heightPx: number) => emit('paragraphMeasure', nodeIndex, heightPx)"
        @table-row-measure="(nodeIndex: number, rowHeightsPx: number[]) => emit('tableRowMeasure', nodeIndex, rowHeightsPx)"
        @content-measure="(heightPx: number) => emit('contentMeasure', heightPx)"
      >
        <template v-if="footnotes.length" #footnotes>
          <DocxNotesSection
            kind="footnote"
            :notes="footnotes"
            :labels="footnoteLabels"
            :model="model"
            :numbering-definitions="model.metadata.numberingDefinitions"
            :theme="theme"
            :on-internal-link-click="onInternalLinkClick"
          />
        </template>
        <template v-if="endnotes.length" #endnotes>
          <DocxNotesSection
            kind="endnote"
            :notes="endnotes"
            :labels="endnoteLabels"
            :model="model"
            :numbering-definitions="model.metadata.numberingDefinitions"
            :theme="theme"
            :on-internal-link-click="onInternalLinkClick"
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
          :revision-display="revisionDisplay"
          :comments-enabled="commentsEnabled"
          :header-distance-px="pageLayout.headerDistancePx"
          :editable="headerEditing"
          :on-internal-link-click="onInternalLinkClick"
          @text-input="onHeaderTextInput"
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
          :revision-display="revisionDisplay"
          :comments-enabled="commentsEnabled"
          :footer-distance-px="pageLayout.footerDistancePx"
          :editable="footerEditing"
          :on-internal-link-click="onInternalLinkClick"
          @text-input="onFooterTextInput"
        />
      </template>

      <!--
        Header/footer activation. Word shows a header/footer band you click to
        start editing; the same affordance creates the part when the section has
        none yet. It is absolutely positioned and ignored by pagination.
      -->
      <template v-if="editable && controller">
        <div
          v-for="region of pageFurnitureRegions"
          :key="`section-region-${region}-${pageIndex}`"
          data-docx-pagination-ignore="true"
          class="docx-page-section-activate"
          :class="`docx-page-section-activate--${region}`"
          :style="regionActivateStyle(region)"
        >
          <button
            type="button"
            class="docx-page-section-activate__button"
            :data-docx-section-region="region"
            :data-testid="`docx-page-${region}-edit`"
            :data-docx-section-empty="hasRegionPart(region) ? 'false' : 'true'"
            :aria-label="regionActionLabel(region)"
            :title="regionActionLabel(region)"
            :aria-pressed="activeSectionRegion === region"
            @click.stop="toggleRegion(region)"
          >
            {{ region === "header" ? m.docx.headerFooter.header : m.docx.headerFooter.footer }}
          </button>
          <span
            v-if="activeSectionRegion === region"
            class="docx-page-section-activate__status"
            :data-testid="`docx-page-${region}-editing`"
          >
            {{ region === "header" ? m.docx.headerFooter.editingHeader : m.docx.headerFooter.editingFooter }}
          </span>
          <button
            v-if="activeSectionRegion === region"
            type="button"
            class="docx-page-section-activate__done"
            :data-testid="`docx-page-${region}-done`"
            :aria-label="m.docx.headerFooter.closeEditing"
            :title="m.docx.headerFooter.closeEditing"
            @click.stop="closeRegion()"
          >
            {{ m.docx.headerFooter.closeEditing }}
          </button>
        </div>
      </template>

      <!-- Floating image layer -->
      <DocxImageLayer
        v-if="editable && controller"
        :data-docx-pagination-ignore="true"
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
        :data-docx-pagination-ignore="true"
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
      :comment-actions="editable && Boolean(controller)"
      :composer="pageCommentComposer"
      :page-index="pageIndex"
      :render-comment-card="renderCommentCard"
      :render-tracked-change-card="renderTrackedChangeCard"
      @activate="activateAnnotation($event, { scrollToAnchor: true })"
      @card-height="reportGutterCardHeight"
      @decide="onGutterDecide"
      @reply="onGutterCommentReply"
      @resolve="onGutterCommentResolve"
      @delete="onGutterCommentDelete"
      @composer-submit="onGutterComposerSubmit"
      @composer-cancel="onGutterComposerCancel"
    />
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, onUnmounted, watch } from "vue"
import type {
  DocModel,
  DocxComment,
  DocxCommentCardRenderProps,
  DocxEditorController,
  DocxRevisionDisplay,
  DocxTrackedChange,
  DocxTrackedChangeCardRenderProps,
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
  resolveDocumentLayoutDefaults,
  trackedChangeDuplicatesInlineMarkup,
  resolveDocumentSectionPropertiesXml,
  resolveClippedBodyRenderedBottomPx,
  resolveClippedMeasuredHeightPx,
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
import { useDocxMessages } from "../i18n"
import { createDocxFieldResolvers } from "../composables/docx-field-resolvers"
import type { DocxBookmarkTargetHighlight } from "../composables/docx-bookmarks"
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
  /** Label Word shows for the page (`iv`, `12`); `w:fmt="none"` → empty. */
  pageLabel?: string
  /** True when the page only exists so a section starts on an even/odd page. */
  parityBlankPage?: boolean
  headerSection?: HeaderSection
  footerSection?: FooterSection
  /** Section this page belongs to; header/footer authoring targets it. */
  sectionIndex?: number
  /** `w:type` variant this page looks for (`default`, `first`, `even`). */
  headerReferenceType?: string
  footerReferenceType?: string
  trackedChangesEnabled?: boolean
  /** Review display: inline markup, gutter balloons, or final text only. */
  revisionDisplay?: DocxRevisionDisplay
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
  /** Internal `#bookmark` link clicked anywhere on this page. */
  onInternalLinkClick?: (bookmarkName: string) => void
  /** Transient highlight after the viewer revealed a bookmark on this page. */
  bookmarkTarget?: DocxBookmarkTargetHighlight
  /** Host renderer that replaces every comment card on this page. */
  renderCommentCard?: (props: DocxCommentCardRenderProps) => unknown
  /** Host renderer that replaces every tracked-change card on this page. */
  renderTrackedChangeCard?: (
    props: DocxTrackedChangeCardRenderProps
  ) => unknown
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
    revisionDisplay: "balloons" as DocxRevisionDisplay,
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

// Header/footer authoring: one region at a time on this page. The model is
// shared by every page of the section, so a commit updates all of them.
const activeSectionRegion = ref<"header" | "footer" | undefined>(undefined)
const headerEditing = computed(
  () => activeSectionRegion.value === "header" && Boolean(props.controller)
)
const footerEditing = computed(
  () => activeSectionRegion.value === "footer" && Boolean(props.controller)
)
const pageFurnitureRegions = ["header", "footer"] as const

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

// ── Header/footer authoring ────────────────────────────────────────
const m = useDocxMessages()

function sectionEntry(region: "header" | "footer"): HeaderSection | FooterSection | undefined {
  return region === "header" ? props.headerSection : props.footerSection
}

function hasRegionPart(region: "header" | "footer"): boolean {
  return (sectionEntry(region)?.nodes.length ?? 0) > 0
}

function regionActionLabel(region: "header" | "footer"): string {
  const group = m.value.docx.headerFooter
  if (hasRegionPart(region)) {
    return region === "header" ? group.editHeader : group.editFooter
  }
  return region === "header" ? group.addHeader : group.addFooter
}

function targetSectionIndex(): number {
  const index = Number(props.sectionIndex)
  return Number.isFinite(index) ? Math.max(0, Math.round(index)) : 0
}

/**
 * The band sits at the top/bottom edge of the paper, above the header distance,
 * so it never covers the header text it activates.
 */
function regionActivateStyle(region: "header" | "footer"): Record<string, string> {
  const style: Record<string, string> = {
    left: `${props.pageLayout.marginsPx?.left ?? 72}px`,
    right: `${props.pageLayout.marginsPx?.right ?? 72}px`,
  }
  if (region === "header") style.top = "2px"
  else style.bottom = "2px"
  return style
}

function closeRegion(): void {
  activeSectionRegion.value = undefined
}

function toggleRegion(region: "header" | "footer"): void {
  if (!props.editable || !props.controller) return
  if (activeSectionRegion.value === region) {
    closeRegion()
    return
  }
  if (!hasRegionPart(region)) {
    // No part to edit yet: create the variant this page looks for (default,
    // title page or even page) and open it for editing.
    const referenceType =
      region === "header" ? props.headerReferenceType : props.footerReferenceType
    props.controller.addHeaderFooter(targetSectionIndex(), region, { referenceType })
  }
  activeSectionRegion.value = region
}

function onHeaderTextInput(nodeIndex: number, text: string): void {
  const partName = props.headerSection?.partName
  if (!partName || !props.controller) return
  props.controller.commitSectionParagraphText(
    { region: "header", partName, nodeIndex },
    text
  )
}

function onFooterTextInput(nodeIndex: number, text: string): void {
  const partName = props.footerSection?.partName
  if (!partName || !props.controller) return
  props.controller.commitSectionParagraphText(
    { region: "footer", partName, nodeIndex },
    text
  )
}

/** Clicks inside a header/footer story must keep the region open. */
function isSectionStoryTarget(target: EventTarget | null): boolean {
  const element = target as (HTMLElement & { closest?: (selector: string) => unknown }) | null
  if (!element || typeof element.closest !== "function") return false
  return Boolean(
    element.closest(
      "[data-docx-page-header], [data-docx-page-footer], .docx-page-section-activate"
    )
  )
}
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
  // Header/footer paragraphs lay out with the document's own spacing defaults
  // (automatic paragraph spacing and its tab grid).
  const layoutDefaults = resolveDocumentLayoutDefaults(props.model.metadata)
  const reserveOptions = {
    estimateDocNodeHeightPx: (
      node: DocModel["nodes"][number],
      availableWidthPx?: number,
      numberingDefinitions?: DocModel["metadata"]["numberingDefinitions"],
      docGridLinePitchPx?: number
    ) =>
      estimateDocNodeHeightPx(
        node,
        availableWidthPx,
        numberingDefinitions,
        docGridLinePitchPx,
        layoutDefaults
      ),
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
  () =>
    props.commentsEnabled ||
    (props.trackedChangesEnabled && props.revisionDisplay === "balloons")
)

// Hidden revisions mean the final text, regardless of the preferred markup
// mode. The gutter selectors keep their value for when revisions are shown.
const renderRevisionDisplay = computed<DocxRevisionDisplay>(() =>
  props.trackedChangesEnabled ? props.revisionDisplay : "none"
)

// Inline display keeps revisions in the text flow; none hides them. Only
// balloons reserve gutter cards for tracked changes.
const trackedChangeCardsEnabled = computed(
  () => props.trackedChangesEnabled && props.revisionDisplay === "balloons"
)

// Insertions and move-to content stay inline in balloons mode, so only
// revisions without inline markup get a card (deletions, formatting, table
// structure).
const pageTrackedChanges = computed(() =>
  (props.trackedChanges ?? props.controller?.trackedChanges ?? []).filter(
    (change) => !trackedChangeDuplicatesInlineMarkup(change)
  )
)

// ── Annotation gutter ──────────────────────────────────────────────
// Annotations are numbered and ordered in document order, then placed against
// their page-space anchors: the DOM resolver measures each anchor inside the
// page (independent of zoom) and the pure layout stacks the cards from there.
const pageAnnotations = computed(() =>
  buildPageGutterAnnotations({
    comments: props.comments ?? props.controller?.comments ?? [],
    trackedChanges: pageTrackedChanges.value,
    pageSegments: props.pageNodeSegments,
    showComments: props.commentsEnabled,
    showTrackedChanges: trackedChangeCardsEnabled.value,
    headerPartName: props.headerSection?.partName,
    footerPartName: props.footerSection?.partName,
  })
)

const pageMeasuredHeightPx = ref(0)

/**
 * The inline comment editor when its anchor lives on this page.
 *
 * A new comment is placed beside the paragraph its selection starts in; a
 * reply is placed beside the card of the comment it answers. Other pages get
 * `undefined` and render no editor.
 */
const pageCommentComposer = computed<
  { mode: "new" | "reply"; commentId?: number } | undefined
>(() => {
  const composer = props.controller?.commentComposer
  if (!composer) return undefined

  if (composer.mode === "reply") {
    const commentId = composer.commentId
    if (commentId === undefined) return undefined
    const onThisPage = pageAnnotations.value.some(
      (entry) => entry.comment?.commentId === commentId
    )
    return onThisPage ? { mode: "reply", commentId } : undefined
  }

  const anchor = composer.anchor
  if (!anchor) return undefined
  const nodeIndex =
    anchor.kind === "paragraph" ? anchor.nodeIndex : anchor.tableIndex
  return props.pageNodeSegments.some(
    (segment) => segment.nodeIndex === nodeIndex
  )
    ? { mode: "new" }
    : undefined
})

function onGutterCommentReply(commentId: number): void {
  props.controller?.beginCommentReply(commentId)
}

function onGutterCommentResolve(commentId: number, resolved: boolean): void {
  props.controller?.resolveComment(commentId, resolved)
}

function onGutterCommentDelete(commentId: number): void {
  props.controller?.deleteComment(commentId)
}

function onGutterComposerSubmit(text: string): void {
  props.controller?.submitCommentComposer(text)
}

function onGutterComposerCancel(): void {
  props.controller?.cancelCommentComposer()
}

const {
  positioned: positionedAnnotations,
  zoomFactor: annotationZoomFactor,
  reportCardHeight: reportGutterCardHeight,
  refresh: refreshGutterLayout,
} = useDocxGutterLayout({
  pageElement: pageSurfaceRef,
  annotations: () => pageAnnotations.value,
  layoutSignal: () =>
    [
      props.model,
      props.pageNodeSegments,
      props.pageWidthPx,
      props.theme,
      props.revisionDisplay,
    ] as const,
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
  // Clicking the page body finishes header/footer editing; clicks inside a
  // story (or its activation band) keep it open.
  if (activeSectionRegion.value && !isSectionStoryTarget(event.target)) {
    closeRegion()
  }
  // Clicks on a commented / revised run link the text to its gutter card.
  const annotationId = annotationIdForTarget(event.target)
  if (annotationId) {
    activateAnnotation(annotationId)
  }
}

// ── Measure page height ────────────────────────────────────────────
let resizeObserver: ResizeObserver | null = null

function measurePage(): void {
  const surface = pageSurfaceRef.value
  if (!surface) return
  // offsetHeight is layout height, unaffected by ancestor CSS transforms.
  // getBoundingClientRect().height would already include scale(zoomFactor),
  // causing a double scale when the viewer multiplies by zoomFactor again.
  const fallbackHeightPx =
    surface.offsetHeight > 0 ? surface.offsetHeight : props.pageLayout.pageHeightPx
  // Measure what the page shows, not its scroll area: content that a clipping
  // ancestor (the page body, a sliced table cell, a column) cuts off must not
  // inflate the height the viewer uses for page offsets and thumbnails.
  const heightPx =
    resolveClippedMeasuredHeightPx({
      // A page is never shorter than one paper, even when its content is.
      minimumHeightPx: props.pageLayout.pageHeightPx,
      fallbackHeightPx,
      clippedContentBottomPx: resolveClippedBodyRenderedBottomPx(surface, {
        zoomFactor: props.zoomFactor,
      }),
    }) ?? fallbackHeightPx
  if (heightPx !== pageMeasuredHeightPx.value) {
    pageMeasuredHeightPx.value = heightPx
    // Paragraph anchors move when content reflows; re-measure them.
    refreshGutterLayout()
  }
  emit("measure", { heightPx })
}

onMounted(() => {
  publishPageSurface()
  measurePage()
  const ResizeObserverConstructor = globalThis.ResizeObserver
  if (pageSurfaceRef.value && typeof ResizeObserverConstructor === "function") {
    resizeObserver = new ResizeObserverConstructor(() => {
      measurePage()
    })
    resizeObserver.observe(pageSurfaceRef.value)
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
  position: relative;
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

/*
 * Header/footer activation band. Absolutely positioned at the paper edge so it
 * never changes page flow or covers the header text it activates.
 */
.docx-page-section-activate {
  position: absolute;
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 6px;
  z-index: 6;
  pointer-events: none;
}
.docx-page-section-activate > * {
  pointer-events: auto;
}
.docx-page-section-activate__button,
.docx-page-section-activate__done {
  padding: 1px 8px;
  border: 1px dashed rgb(100 116 139 / 55%);
  border-radius: 4px;
  background: rgb(255 255 255 / 78%);
  color: rgb(51 65 85);
  font: 11px/1.5 system-ui, -apple-system, "Segoe UI", sans-serif;
  cursor: pointer;
  opacity: 0.72;
}
.docx-page-section-activate__button:hover,
.docx-page-section-activate__button:focus-visible,
.docx-page-section-activate__done:hover,
.docx-page-section-activate__done:focus-visible {
  opacity: 1;
  border-style: solid;
  outline: none;
}
.docx-page-section-activate__status {
  color: rgb(100 116 139);
  font: 10px/1.5 system-ui, -apple-system, "Segoe UI", sans-serif;
}
@media print {
  .docx-page-section-activate {
    display: none !important;
  }
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
