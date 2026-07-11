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
  onUnmounted,
  ref,
  watch,
  type PropType,
  type VNode,
} from "vue"
import {
  paragraphText,
  type DocxEditorController,
  type DocModel,
  type ParagraphNode,
  type TableNode,
  type DocumentPageNodeSegment,
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
    pageNodeSegments: {
      type: Array as PropType<DocumentPageNodeSegment[]>,
      required: true,
    },
    pageLayout: {
      type: Object as PropType<PageLayoutInfo>,
      required: true,
    },
    pageContentWidthPx: { type: Number, required: true },
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
    searchQuery: { type: String, default: "" },
    activeSearchNodeIndex: { type: Number, required: false },
    pageNumber: { type: Number, required: true },
    totalPages: { type: Number, required: true },
    pageNumberFormat: { type: String, required: false },
  },
  emits: ["measure"],
  setup(props, { emit, slots }) {
    const selectionOwner = {}
    const bodyElement = ref<HTMLElement>()
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

    // ── Render ─────────────────────────────────────────────────────
    return () => {
      const segments = props.pageNodeSegments
      const bodyChildren: VNode[] = []

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
      } else {
        for (let i = 0; i < segments.length; i++) {
          const segment = segments[i]
          const node = segmentNode(segment)
          const key = segmentKey(segment, i)

          if (!node) {
            bodyChildren.push(
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
            bodyChildren.push(
              h(DocxParagraphHost, {
                key,
                paragraph: node as ParagraphNode,
                paragraphIndex: segment.nodeIndex,
                editable: props.editable && !props.trackedChangesEnabled,
                documentTheme: props.theme,
                controller: props.controller,
                showTrackedChanges: props.trackedChangesEnabled,
                showCommentHighlights: props.commentsEnabled,
                numberingDefinitions: props.model.metadata.numberingDefinitions,
                pageNumber: props.pageNumber,
                totalPages: props.totalPages,
                pageNumberFormat: props.pageNumberFormat,
                "data-docx-search-match": props.searchQuery && paragraphText(node).toLocaleLowerCase().includes(props.searchQuery.toLocaleLowerCase()) ? "true" : undefined,
                "data-docx-search-active": props.activeSearchNodeIndex === segment.nodeIndex ? "true" : undefined,
                style: props.searchQuery && paragraphText(node).toLocaleLowerCase().includes(props.searchQuery.toLocaleLowerCase())
                  ? {
                      background: props.activeSearchNodeIndex === segment.nodeIndex ? "#fde68a" : "#fef3c7",
                      boxShadow: props.activeSearchNodeIndex === segment.nodeIndex ? "0 0 0 2px #f59e0b" : undefined,
                    }
                  : undefined,
                onTextInput: (nodeIndex: number, text: string) => {
                  props.controller?.commitParagraphText(nodeIndex, text)
                },
                onFocus: (nodeIndex: number) => {
                  ensureParagraphSelection(nodeIndex)
                },
                onClick: (nodeIndex: number, _event: MouseEvent) => {
                  ensureParagraphSelection(nodeIndex)
                },
              })
            )
          } else if (node.type === "table") {
            bodyChildren.push(
              h(DocxTableHost, {
                key,
                table: node as TableNode,
                tableIndex: segment.nodeIndex,
                editable: props.editable && !props.trackedChangesEnabled,
                controller: props.controller,
                rowRange: segment.tableRowRange,
                documentTheme: props.theme,
                showTrackedChanges: props.trackedChangesEnabled,
                showCommentHighlights: props.commentsEnabled,
                numberingDefinitions: props.model.metadata.numberingDefinitions,
                pageNumber: props.pageNumber,
                totalPages: props.totalPages,
                pageNumberFormat: props.pageNumberFormat,
                searchQuery: props.searchQuery,
                searchActive: props.activeSearchNodeIndex === segment.nodeIndex,
              })
            )
          }
        }
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
            minHeight: "0px",
            overflow: "visible",
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
