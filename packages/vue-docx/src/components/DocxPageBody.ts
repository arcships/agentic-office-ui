// DocxPageBody — Vue render function component for page body content.
//
// Renders paragraphs and tables within a single page's body area.
// Uses Vue's h() render function, not a template, for maximum
// compatibility with the dynamically-computed node structure.
//
// Migration source: @extend-ai/react-docx commit 6f70b92,
//   react-viewer/src/editor.tsx — renderDocumentNode / renderPageBody sections.

import { h, defineComponent, type PropType, type VNode } from "vue"
import type {
  DocxEditorController,
  DocModel,
  ParagraphNode,
  TableNode,
  DocumentPageNodeSegment,
} from "@extend-ai/docx-core"
import DocxParagraphHost from "./DocxParagraphHost.vue"
import DocxTableHost from "./DocxTableHost.vue"

// ── Types ──────────────────────────────────────────────────────────
interface PageLayoutInfo {
  pageWidthPx: number
  pageHeightPx: number
  marginsPx?: { top: number; bottom: number; left: number; right: number }
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
      required: true,
    },
    editable: { type: Boolean, default: true },
    theme: {
      type: String as PropType<"light" | "dark">,
      default: "light",
    },
    trackedChangesEnabled: { type: Boolean, default: false },
    commentsEnabled: { type: Boolean, default: false },
  },
  emits: ["measure"],
  setup(props, { emit }) {
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
      const model = props.controller.model
      if (!model) return undefined
      return model.nodes[segment.nodeIndex]
    }

    function segmentHeightPx(_segment: DocumentPageNodeSegment): number {
      // Return 0 to let content determine height dynamically
      return 0
    }

    // ── Render ─────────────────────────────────────────────────────
    return () => {
      const segments = props.pageNodeSegments
      const bodyChildren: VNode[] = []

      if (segments.length === 0 && props.editable) {
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
              props.controller.appendParagraph("")
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
                onTextInput: (nodeIndex: number, text: string) => {
                  props.controller.commitParagraphText(nodeIndex, text)
                },
                onFocus: (nodeIndex: number) => {
                  props.controller.selectParagraph(nodeIndex)
                },
                onClick: (nodeIndex: number, _event: MouseEvent) => {
                  props.controller.selectParagraph(nodeIndex)
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
              })
            )
          }
        }
      }

      return h(
        "div",
        {
          "data-docx-page-body": "true",
          class: "docx-page-body",
          style: {
            isolation: "isolate",
            display: "flex",
            flexDirection: "column",
            minHeight: "0px",
            overflow: "visible",
            color: props.theme === "dark" ? "#f9fafb" : "#111827",
          },
        },
        bodyChildren
      )
    }
  },
})
