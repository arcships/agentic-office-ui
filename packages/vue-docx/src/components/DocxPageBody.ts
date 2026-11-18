// DocxPageBody — Vue render function component for page body content.
//
// Renders paragraphs and tables within a single page's body area.
// Uses Vue's h() render function, not a template, for maximum
// compatibility with the dynamically-computed node structure.
//
// Migration source: @extend-ai/react-docx commit 6f70b92,
//   react-viewer/src/editor.tsx — renderDocumentNode / renderPageBody sections.

import {
  h,
  defineComponent,
  nextTick,
  onMounted,
  onUnmounted,
  ref,
  watch,
  type PropType,
  type VNode,
} from "vue"
import {
  type DocxEditorController,
  type DocxFloatingPageOriginPx,
  type DocModel,
  type ParagraphNode,
  type TableNode,
  type DocumentPageNodeSegment,
} from "@arcships/docx-core"
import {
  DEFAULT_DOC_PAGE_MARGIN,
  resolveFloatingPageOriginPx,
  resolveWrappedImageFlowClearancePx,
  resolveDefaultTabStopPx,
  resolveDocumentLayoutDefaults,
  resolveParagraphAutoSpacingInStory,
  resolveClippedBodyRenderedBottomPx,
  resolveClippedMeasuredHeightPx,
  docxNodeSourceId,
  type DocxPageFlowWrapClearance,
  type DocumentLayoutDefaults,
} from "@arcships/docx-core"
import DocxParagraphHost from "./DocxParagraphHost.vue"
import DocxTableHost from "./DocxTableHost.vue"
import {
  restoreDomSelectionFromTextRange,
  syncDomSelectionToEditor,
} from "../composables/editor-selection"

// ── Types ──────────────────────────────────────────────────────────
interface PageLayoutInfo {
  pageWidthPx: number
  pageHeightPx: number
  marginsPx?: { top: number; bottom: number; left: number; right: number }
}

const SELECTION_TRACKING_OWNER = Symbol("docx-selection-tracking-owner")
type SelectionTrackingRoot = HTMLElement & {
  [SELECTION_TRACKING_OWNER]?: {
    owner: object
    detach: () => void
  }
}

// ── Component ──────────────────────────────────────────────────────
export default defineComponent({
  name: "DocxPageBody",
  components: {
    DocxParagraphHost,
    DocxTableHost,
  },
  props: {
    pageIndex: { type: Number, required: true },
    /** Zoom applied to the page subtree; measured rects are unscaled by it. */
    zoomFactor: { type: Number, default: 1 },
    pageNodeSegments: {
      type: Array as PropType<DocumentPageNodeSegment[]>,
      required: true,
    },
    /** Per-body-column segments; omitted for single-column pages. */
    pageColumns: {
      type: Array as PropType<DocumentPageNodeSegment[][]>,
      required: false,
    },
    columnLayout: {
      type: Object as PropType<{
        count: number
        gapPx: number
        widthsPx: number[]
      }>,
      required: false,
    },
    pageLayout: {
      type: Object as PropType<PageLayoutInfo>,
      required: true,
    },
    pageContentWidthPx: { type: Number, required: true },
    /** Page-relative origin for margin/page/column anchored floating images. */
    floatingPageOriginPx: {
      type: Object as PropType<DocxFloatingPageOriginPx>,
      required: false,
    },
    controller: {
      type: Object as PropType<DocxEditorController>,
      required: false,
    },
    model: {
      type: Object as PropType<DocModel>,
      required: true,
    },
    editable: { type: Boolean, default: true },
    theme: {
      type: String as PropType<"light" | "dark">,
      default: "light",
    },
    trackedChangesEnabled: { type: Boolean, default: false },
    commentsEnabled: { type: Boolean, default: false },
    pageNumber: { type: Number, required: true },
    totalPages: { type: Number, required: true },
    pageNumberFormat: { type: String, required: false },
    sectionPages: { type: Number, required: false },
    resolveStyleRefFieldValue: {
      type: Function as PropType<
        (target: string, beforeNodeIndex?: number) => string | undefined
      >,
      required: false,
    },
    resolveBookmarkFieldValue: {
      type: Function as PropType<(target: string) => string | undefined>,
      required: false,
    },
    resolveBookmarkPageNumber: {
      type: Function as PropType<(target: string) => number | undefined>,
      required: false,
    },
    noteMarkerLabels: {
      type: Object as PropType<{
        footnote?: Map<number, string>
        endnote?: Map<number, string>
      }>,
      required: false,
    },
    coverImageChildIndexesByNode: {
      type: Object as PropType<Record<number, readonly number[]>>,
      required: false,
    },
    headerReservePx: { type: Number, default: 0 },
    footerReservePx: { type: Number, default: 0 },
  },
  emits: ["measure", "paragraphMeasure", "tableRowMeasure", "contentMeasure"],
  setup(props, { emit, slots }) {
    const selectionOwner = {}
    const bodyElement = ref<HTMLElement>()

    function emitContentMeasure(): void {
      const body = bodyElement.value
      if (!body) return
      // `scrollHeight` counts content that a clipping ancestor cut off, so a
      // page whose body, column or sliced cell clips overflowing paragraphs
      // would report a taller body and push pagination away from Word. Measure
      // the visible bottom instead and keep `scrollHeight` as the fallback.
      const heightPx = resolveClippedMeasuredHeightPx({
        fallbackHeightPx: body.scrollHeight,
        // A body never reports less than the height it occupies on screen.
        minimumHeightPx: body.clientHeight,
        clippedContentBottomPx: resolveClippedBodyRenderedBottomPx(body, {
          zoomFactor: props.zoomFactor,
        }),
      })
      if (heightPx !== undefined && Number.isFinite(heightPx) && heightPx > 0) {
        emit("contentMeasure", heightPx)
      }
    }

    onMounted(() => { void nextTick(emitContentMeasure) })
    watch(
      () => props.pageNodeSegments,
      () => { void nextTick(emitContentMeasure) },
      { flush: "post" }
    )
    let selectionDocument: Document | undefined
    let selectionRoot: HTMLElement | undefined
    let toolbarPointerDown = false
    let toolbarPointerReset: ReturnType<typeof setTimeout> | undefined

    function editorRoot(): HTMLElement | undefined {
      const body = bodyElement.value
      return (body?.closest?.(".docx-editor-viewer") as HTMLElement | undefined) ?? body
    }

    function resetToolbarPointerState(): void {
      toolbarPointerDown = false
      if (toolbarPointerReset !== undefined) clearTimeout(toolbarPointerReset)
      toolbarPointerReset = undefined
    }

    function onEditorPointerDown(event: Event): void {
      const target = event.target as Element | null
      toolbarPointerDown = Boolean(
        target?.closest?.(".docx-toolbar, .docx-toolbar-shell")
      )
      if (toolbarPointerReset !== undefined) clearTimeout(toolbarPointerReset)
      toolbarPointerReset = setTimeout(resetToolbarPointerState, 0)
    }

    function syncSelection(): void {
      const root = editorRoot()
      if (!props.editable || !props.controller || !root) return
      syncDomSelectionToEditor(props.controller, root, undefined, {
        preserveExpandedSelection: toolbarPointerDown,
      })
    }

    function detachSelectionTracking(): void {
      selectionDocument?.removeEventListener("selectionchange", syncSelection)
      selectionRoot?.removeEventListener("pointerdown", onEditorPointerDown, true)
      const ownedRoot = selectionRoot as SelectionTrackingRoot | undefined
      if (ownedRoot?.[SELECTION_TRACKING_OWNER]?.owner === selectionOwner) {
        delete ownedRoot[SELECTION_TRACKING_OWNER]
      }
      selectionDocument = undefined
      selectionRoot = undefined
      resetToolbarPointerState()
    }

    function activateSelectionTracking(): void {
      const root = editorRoot()
      const ownerDocument = root?.ownerDocument
      if (!props.editable || !props.controller || !root || !ownerDocument) return
      if (selectionDocument === ownerDocument && selectionRoot === root) return
      const trackingRoot = root as SelectionTrackingRoot
      const previousOwner = trackingRoot[SELECTION_TRACKING_OWNER]
      if (previousOwner?.owner !== selectionOwner) previousOwner?.detach()
      detachSelectionTracking()
      selectionDocument = ownerDocument
      selectionRoot = root
      ownerDocument.addEventListener("selectionchange", syncSelection)
      root.addEventListener("pointerdown", onEditorPointerDown, true)
      trackingRoot[SELECTION_TRACKING_OWNER] = {
        owner: selectionOwner,
        detach: detachSelectionTracking,
      }
    }

    function restoreRequestedSelection(focus: boolean): void {
      const controller = props.controller
      const root = editorRoot()
      const range = controller?.historyRestoreRequest?.activeTextRange ??
        controller?.activeTextRange
      if (!props.editable || !controller || !root || !range) return
      void nextTick(() => {
        const restored = restoreDomSelectionFromTextRange(root, range, {
          focus,
          owner: bodyElement.value,
        })
        if (restored) {
          // Focusing a contenteditable can synchronously collapse its native
          // selection and run the paragraph focus handler before the browser
          // emits selectionchange. Confirm the restored DOM range back into
          // the public controller immediately so undo/redo never exposes that
          // transient collapsed range as the final editor state.
          syncDomSelectionToEditor(controller, root)
        }
      })
    }

    watch(
      () => props.controller?.historyRestoreRequest?.nonce,
      () => restoreRequestedSelection(true)
    )
    watch(
      () => [props.editable, props.controller] as const,
      ([editable, controller]) => {
        if (!editable || !controller) detachSelectionTracking()
      }
    )
    onUnmounted(detachSelectionTracking)

    // ── Helpers ────────────────────────────────────────────────────
    function segmentKey(segment: DocumentPageNodeSegment, index: number): string {
      const base = `node-${segment.nodeIndex}`
      if (segment.tableRowRange) {
        return `${base}-rows-${segment.tableRowRange.startRowIndex}-${segment.tableRowRange.endRowIndex}`
      }
      if (segment.paragraphLineRange) {
        return `${base}-lines-${segment.paragraphLineRange.startLineIndex}-${segment.paragraphLineRange.endLineIndex}`
      }
      return `${base}-${index}`
    }

    function segmentNode(
      segment: DocumentPageNodeSegment
    ): ParagraphNode | TableNode | undefined {
      return props.model.nodes[segment.nodeIndex]
    }

    function segmentHeightPx(_segment: DocumentPageNodeSegment): number {
      // Return 0 to let content determine height dynamically
      return 0
    }

    /**
     * Top clearance for each paragraph on this page: a float that outlives its
     * anchor paragraph would otherwise be covered by the following paragraph,
     * because every paragraph is an isolated flex item that CSS floats cannot
     * reach across. Only whole-paragraph pages are cleared — a page whose
     * paragraphs were split by pagination has no reliable flow origin.
     */
    function wrappedImageClearances(): DocxPageFlowWrapClearance[] {
      const paragraphs: ParagraphNode[] = []
      for (const segment of props.pageNodeSegments) {
        const node = segmentNode(segment)
        if (!node || node.type !== "paragraph" || segment.paragraphLineRange) {
          return []
        }
        paragraphs.push(node)
      }
      if (paragraphs.length < 2) return []
      // Fast path: most pages have no floating image at all.
      const hasFloatingImage = paragraphs.some((paragraph) =>
        paragraph.children.some(
          (child) => child.type === "image" && Boolean(child.floating)
        )
      )
      if (!hasFloatingImage) return []
      return resolveWrappedImageFlowClearancePx(
        paragraphs,
        props.pageContentWidthPx,
        props.model.metadata.numberingDefinitions
      )
    }

    // Floating images anchored to the page, its margins, or the text column must
    // be placed from the page origin. Without it `absoluteFloatingImageStyle`
    // degrades every margin/page anchor to the (0, 0) page corner.
    function floatingPageOrigin(): DocxFloatingPageOriginPx {
      if (props.floatingPageOriginPx) return props.floatingPageOriginPx
      return resolveFloatingPageOriginPx(
        props.pageLayout.pageWidthPx,
        props.pageLayout.marginsPx ?? {
          top: DEFAULT_DOC_PAGE_MARGIN,
          right: DEFAULT_DOC_PAGE_MARGIN,
          bottom: DEFAULT_DOC_PAGE_MARGIN,
          left: DEFAULT_DOC_PAGE_MARGIN,
        }
      )
    }

    // Automatic (HTML) paragraph spacing and document tab grid, resolved once
    // per page so every host renders the numbers the paginator measured.
    const layoutDefaults: DocumentLayoutDefaults = resolveDocumentLayoutDefaults(
      props.model.metadata
    )
    const defaultTabStopPx = resolveDefaultTabStopPx(
      layoutDefaults.defaultTabStopTwips
    )

    function paragraphAutoSpacing(nodeIndex: number, paragraph: ParagraphNode) {
      return resolveParagraphAutoSpacingInStory(
        props.model.nodes,
        nodeIndex,
        paragraph,
        layoutDefaults
      )
    }

    function ensureParagraphSelection(nodeIndex: number): void {
      const controller = props.controller
      if (!controller) return
      const range = controller.activeTextRange
      const rangeStartsHere = range?.start.location.kind === "paragraph" &&
        range.start.location.nodeIndex === nodeIndex
      const rangeEndsHere = range?.end.location.kind === "paragraph" &&
        range.end.location.nodeIndex === nodeIndex
      // The browser range is the more precise selection. Do not clear a
      // caret or expanded range just because the coarse paragraph selection
      // has not caught up with the focus event yet.
      if (rangeStartsHere && rangeEndsHere) return
      const selection = controller.selection
      if (selection.kind === "paragraph" && selection.nodeIndex === nodeIndex) return
      controller.selectParagraph(nodeIndex)
    }

    function renderSegmentList(
      segments: DocumentPageNodeSegment[],
      keyPrefix: string,
      paragraphSlotRef: { value: number },
      wrapClearances: DocxPageFlowWrapClearance[],
      columnIndex?: number
    ): VNode[] {
      const children: VNode[] = []
      for (let i = 0; i < segments.length; i++) {
        const segment = segments[i]
        const node = segmentNode(segment)
        const key = `${keyPrefix}${segmentKey(segment, i)}`

        if (!node) {
          children.push(
            h("div", {
              key,
              class: "docx-page-body-unknown",
              style: {
                padding: "8px",
                color: "#9ca3af",
                fontSize: "12px",
              },
            }, "Unsupported node type")
          )
        } else if (node.type === "paragraph") {
          const wrapClearanceTopPx =
            wrapClearances[paragraphSlotRef.value++]?.marginTopPx ?? 0
          const autoSpacing = paragraphAutoSpacing(
            segment.nodeIndex,
            node as ParagraphNode
          )
          children.push(
            h(DocxParagraphHost, {
              key,
              paragraph: node as ParagraphNode,
              paragraphIndex: segment.nodeIndex,
              sourceId: docxNodeSourceId(segment.nodeIndex),
              // Showing revisions is a display choice: it must not turn the
              // document read-only, otherwise a document with revisions can
              // never be edited or cleaned up.
              editable: props.editable,
              documentTheme: props.theme,
              controller: props.controller,
              floatingPageOriginPx: floatingPageOrigin(),
              wrapClearanceTopPx,
              lineRange: segment.paragraphLineRange,
              showTrackedChanges: props.trackedChangesEnabled,
              showCommentHighlights: props.commentsEnabled,
              numberingDefinitions: props.model.metadata.numberingDefinitions,
              autoSpacingBeforePx: autoSpacing.beforePx,
              autoSpacingAfterPx: autoSpacing.afterPx,
              defaultTabStopPx,
              pageNumber: props.pageNumber,
              totalPages: props.totalPages,
              pageNumberFormat: props.pageNumberFormat,
              sectionPages: props.sectionPages,
              resolveStyleRefFieldValue: props.resolveStyleRefFieldValue,
              resolveBookmarkFieldValue: props.resolveBookmarkFieldValue,
              resolveBookmarkPageNumber: props.resolveBookmarkPageNumber,
              fieldResolutionNodeIndex: segment.nodeIndex,
              noteMarkerLabels: props.noteMarkerLabels,
              skipImageChildIndexes:
                props.coverImageChildIndexesByNode?.[segment.nodeIndex],
              onTextInput: (nodeIndex: number, text: string) => {
                props.controller?.commitParagraphText(nodeIndex, text)
              },
              onFocus: (nodeIndex: number) => {
                ensureParagraphSelection(nodeIndex)
              },
              onClick: (nodeIndex: number, _event: MouseEvent) => {
                ensureParagraphSelection(nodeIndex)
              },
              onMeasure: (nodeIndex: number, heightPx: number) => {
                emit("paragraphMeasure", nodeIndex, heightPx)
              },
            })
          )
        } else if (node.type === "table") {
          children.push(
            h(DocxTableHost, {
              key,
              table: node as TableNode,
              tableIndex: segment.nodeIndex,
              sourceId: docxNodeSourceId(segment.nodeIndex),
              editable: props.editable,
              controller: props.controller,
              rowRange: segment.tableRowRange,
              rowSlice: segment.tableRowSlice,
              documentTheme: props.theme,
              floatingPageOriginPx: floatingPageOrigin(),
              showTrackedChanges: props.trackedChangesEnabled,
              showCommentHighlights: props.commentsEnabled,
              numberingDefinitions: props.model.metadata.numberingDefinitions,
              pageNumber: props.pageNumber,
              totalPages: props.totalPages,
              pageNumberFormat: props.pageNumberFormat,
              sectionPages: props.sectionPages,
              resolveStyleRefFieldValue: props.resolveStyleRefFieldValue,
              resolveBookmarkFieldValue: props.resolveBookmarkFieldValue,
              resolveBookmarkPageNumber: props.resolveBookmarkPageNumber,
              model: props.model,
              maxAvailableWidthPx:
                (columnIndex !== undefined
                  ? props.columnLayout?.widthsPx?.[columnIndex]
                  : undefined) ?? props.pageContentWidthPx,
              noteMarkerLabels: props.noteMarkerLabels,
              onMeasure: (tableIndex: number, rowHeightsPx: number[]) => {
                emit("tableRowMeasure", tableIndex, rowHeightsPx)
              },
            })
          )
        }
      }
      return children
    }

    // ── Render ─────────────────────────────────────────────────────
    return () => {
      const segments = props.pageNodeSegments
      const bodyChildren: VNode[] = []
      const wrapClearances = wrappedImageClearances()
      const paragraphSlotRef = { value: 0 }
      const columns =
        !props.editable && props.pageColumns && props.pageColumns.length > 1
          ? props.pageColumns
          : undefined

      if (segments.length === 0 && props.editable && props.controller) {
        // Empty page hint
        bodyChildren.push(
          h("div", {
            class: "docx-page-body-empty-hint",
            style: {
              minHeight: "28px",
              marginTop: "10px",
              cursor: "text",
              color: "#9ca3af",
              fontSize: "13px",
            },
            onClick: () => {
              props.controller?.appendParagraph("")
            },
          }, "Click to add content")
        )
      } else if (columns) {
        const margins = props.pageLayout.marginsPx ?? {
          top: 0,
          right: 0,
          bottom: 0,
          left: 0,
        }
        const columnHeightPx = Math.max(
          0,
          Math.round(
            props.pageLayout.pageHeightPx -
              margins.top -
              margins.bottom -
              (props.headerReservePx ?? 0) -
              (props.footerReservePx ?? 0)
          )
        )
        bodyChildren.push(
          h(
            "div",
            {
              class: "docx-page-columns",
              "data-docx-page-columns": columns.length,
              style: {
                display: "flex",
                alignItems: "flex-start",
                gap: `${Math.max(0, props.columnLayout?.gapPx ?? 0)}px`,
                height: `${columnHeightPx}px`,
                overflow: "hidden",
              },
            },
            columns.map((columnSegments, columnIndex) =>
              h(
                "div",
                {
                  key: `column-${columnIndex}`,
                  class: "docx-page-column",
                  style: {
                    flex: "0 0 auto",
                    width: `${Math.max(
                      24,
                      Math.round(
                        props.columnLayout?.widthsPx?.[columnIndex] ??
                          props.pageContentWidthPx
                      )
                    )}px`,
                    height: "100%",
                    overflow: "hidden",
                  },
                },
                renderSegmentList(
                  columnSegments,
                  `col-${columnIndex}-`,
                  paragraphSlotRef,
                  [],
                  columnIndex
                )
              )
            )
          )
        )
      } else {
        bodyChildren.push(
          ...renderSegmentList(segments, "", paragraphSlotRef, wrapClearances)
        )
      }

      return h(
        "div",
        {
          ref: bodyElement,
          "data-docx-page-body": "true",
          "data-testid": "docx-page",
          "data-page": props.pageNumber,
          class: "docx-page-body",
          onPointerdown: activateSelectionTracking,
          onFocusin: activateSelectionTracking,
          onMouseup: syncSelection,
          onKeyup: syncSelection,
          onClick: syncSelection,
          style: {
            isolation: "isolate",
            display: "flex",
            flexDirection: "column",
            // Shrink the column to the page body box, then clip at that box.
            // Its children opt out of shrinking (see DocxPageSurface styles)
            // so overflowing paragraphs are clipped instead of compressed.
            flex: "0 1 auto",
            minHeight: "0px",
            overflow: "hidden",
            paddingTop: `${Math.max(0, Math.round(props.headerReservePx))}px`,
            paddingBottom: `${Math.max(0, Math.round(props.footerReservePx))}px`,
            color: props.theme === "dark" ? "#f9fafb" : "#111827",
          },
        },
        [
          ...(slots.header?.() ?? []),
          ...bodyChildren,
          ...(slots.footnotes?.() ?? []),
          ...(slots.endnotes?.() ?? []),
          ...(slots.footer?.() ?? []),
        ]
      )
    }
  },
})
