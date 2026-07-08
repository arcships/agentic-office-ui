// DocxPageBody — renders a page's block list (paragraphs and tables).
// Render function component (no template) for performance.
// Uses LayoutBlock from the layout engine (read-only viewer path).

import { h, defineComponent, type VNode, type PropType } from "vue"
import type {
  LayoutBlock,
  LayoutParagraphBlock,
  LayoutTableBlock,
  LayoutRun,
} from "@extend-ai/docx-core"

const HEADING_FONT_SIZES: Record<number, string> = {
  1: "2rem",
  2: "1.6rem",
  3: "1.35rem",
  4: "1.2rem",
  5: "1.05rem",
  6: "0.95rem",
}

const HIGHLIGHT_TO_CSS: Record<string, string> = {
  yellow: "#fff59d",
  green: "#bbf7d0",
  cyan: "#a5f3fc",
  magenta: "#f5d0fe",
  blue: "#bfdbfe",
  red: "#fecaca",
  black: "#111827",
  white: "#ffffff",
  darkgray: "#9ca3af",
  lightgray: "#e5e7eb",
}

const SCRIPT_FONT_SCALE = 0.65

function resolveHighlightColor(value?: string): string | undefined {
  if (!value) return undefined
  const normalized = value.trim().toLowerCase()
  if (!normalized) return undefined
  if (normalized.startsWith("#")) return normalized
  return HIGHLIGHT_TO_CSS[normalized] ?? normalized
}

function runTextStyle(run: LayoutRun): Record<string, string | undefined> {
  if (run.kind === "image") return {}
  const hasScript =
    run.style?.verticalAlign === "superscript" ||
    run.style?.verticalAlign === "subscript"
  const verticalAlign =
    run.style?.verticalAlign === "superscript"
      ? "super"
      : run.style?.verticalAlign === "subscript"
        ? "sub"
        : undefined
  const textDecoration = [
    run.style?.underline ? "underline" : "",
    run.style?.strike ? "line-through" : "",
  ]
    .filter(Boolean)
    .join(" ") || "none"

  return {
    fontWeight: run.style?.bold ? "700" : undefined,
    fontStyle: run.style?.italic ? "italic" : undefined,
    textDecoration,
    color: run.style?.color,
    backgroundColor: resolveHighlightColor(run.style?.highlight),
    fontSize: run.style?.fontSizePt
      ? `${Number((run.style.fontSizePt * (hasScript ? SCRIPT_FONT_SCALE : 1)).toFixed(3))}pt`
      : hasScript
        ? `${SCRIPT_FONT_SCALE}em`
        : undefined,
    fontFamily: run.style?.fontFamily,
    verticalAlign,
    whiteSpace: "pre-wrap",
  }
}

function renderRun(run: LayoutRun): VNode {
  if (run.kind === "image") {
    const wPx = run.widthPx
    const hPx = run.heightPx
    if (!run.src) {
      return h("span", { key: run.id, style: imagePlaceholderStyle(run) }, "Missing image")
    }
    return h("img", {
      key: run.id,
      src: run.src,
      alt: run.alt ?? "DOCX image",
      style: {
        maxWidth: wPx ? `${wPx}px` : "100%",
        maxHeight: hPx ? `${hPx}px` : undefined,
        verticalAlign: "middle",
        marginInline: "4px",
      },
    })
  }

  const style = runTextStyle(run)
  if (run.link) {
    return h("a", {
      key: run.id,
      href: run.link,
      target: run.link.startsWith("#") ? undefined : "_blank",
      rel: run.link.startsWith("#") ? undefined : "noreferrer noopener",
      style: {
        ...style,
        color: style.color ?? "inherit",
      },
    }, run.text)
  }

  return h("span", { key: run.id, style }, run.text)
}

function imagePlaceholderStyle(run: LayoutRun): Record<string, string | undefined> {
  const wPx = run.kind === "image" ? run.widthPx : undefined
  const hPx = run.kind === "image" ? run.heightPx : undefined
  const w = (wPx ?? 0) <= 56 && (hPx ?? 0) <= 56 ? 12 : 10
  return {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    width: wPx ? `${wPx}px` : "1.8em",
    height: hPx ? `${hPx}px` : "1.8em",
    minWidth: "16px",
    minHeight: "16px",
    border: "1px solid #d1d5db",
    borderRadius: "3px",
    backgroundColor: "#ffffff",
    color: "#0f172a",
    fontSize: `${w}px`,
    fontWeight: "700",
    textTransform: "lowercase",
    fontFamily: "Arial, sans-serif",
    lineHeight: "1",
    verticalAlign: "middle",
    marginInline: "4px",
  }
}

function renderTableBlock(block: LayoutTableBlock): VNode {
  const tableStyle = {
    width: "100%",
    borderCollapse: "collapse" as const,
    tableLayout: "fixed" as const,
    marginBottom: "8px",
  }
  return h("table", { class: "docx-table", style: tableStyle }, [
    h("tbody", null, block.rows.map((row) =>
      h("tr", { key: row.id }, row.cells.map((cell) =>
        h("td", {
          key: cell.id,
          colspan: cell.colSpan,
          style: {
            border: "1px solid #d1d5db",
            padding: "8px",
            backgroundColor: cell.backgroundColor,
            verticalAlign: "top",
            minWidth: "0",
            wordWrap: "break-word",
            overflowWrap: "break-word",
            wordBreak: "break-word",
          },
        }, cell.paragraphs.map((p) =>
          h("p", {
            key: p.id,
            style: {
              margin: "0",
              textAlign: p.align,
              fontWeight: p.headingLevel ? 700 : undefined,
              fontSize: p.headingLevel ? HEADING_FONT_SIZES[p.headingLevel] : undefined,
            },
          }, p.runs.map(renderRun))
        ))
      ))
    ))
  ])
}

function renderParagraphBlock(block: LayoutParagraphBlock): VNode {
  return h("p", {
    key: block.id,
    class: "docx-paragraph",
    style: {
      margin: "0",
      minHeight: block.height ? `${block.height}px` : undefined,
      textAlign: block.align,
      fontWeight: block.headingLevel ? 700 : undefined,
      fontSize: block.headingLevel ? HEADING_FONT_SIZES[block.headingLevel] : undefined,
    },
  }, block.runs.map(renderRun))
}

export default defineComponent({
  name: "DocxPageBody",
  props: {
    blocks: { type: Array as PropType<LayoutBlock[]>, required: true },
  },
  setup(props) {
    return (): VNode => h("div", { class: "docx-page-body" },
      props.blocks.map((block) => {
        if (block.kind === "table") {
          return renderTableBlock(block)
        }
        return renderParagraphBlock(block)
      })
    )
  },
})
