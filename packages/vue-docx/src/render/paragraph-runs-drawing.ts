// Drawing-object and equation run rendering.
//
// The model carries shapes, charts, SmartArt, embedded OLE objects, 3D models and
// equations as paragraph children (issue #33). None of them can be painted by
// the library itself, so the renderer paints what it parsed — text box content,
// chart data, a 3D poster frame — and a labelled frame at the anchored position
// for the rest. The frame is a placeholder, never a silent drop.

import { h, type VNode } from "vue"
import { officeMessagesEn, type OfficeDocxMessages } from "@arcships/office-interaction"
import type {
  DrawingRunNode,
  MathRunNode,
  ParagraphNode,
  TextBoxContent,
} from "@arcships/docx-core"
import {
  chartPreviewDataUri,
  drawingObjectContentStyle,
  drawingObjectFrameStyle,
  drawingObjectIsBehindText,
  drawingObjectPlaceholderStyle,
  drawingObjectSizePx,
  runStyleToCss,
} from "@arcships/docx-core"

export type DocxDrawingLabels = OfficeDocxMessages["drawing"]

export interface DrawingRunRenderContext {
  drawing: DrawingRunNode
  key: string
  childIndex: number
  labels?: DocxDrawingLabels
  paragraphSourceId?: string
  documentTheme?: "light" | "dark"
}

export interface MathRunRenderContext {
  math: MathRunNode
  key: string
  childIndex: number
  paragraphSourceId?: string
}

function drawingKindLabel(
  drawing: DrawingRunNode,
  labels: DocxDrawingLabels,
): string {
  switch (drawing.kind) {
    case "textBox":
      return labels.textBox
    case "wordArt":
      return labels.wordArt
    case "group":
      return labels.group
    case "smartArt":
      return labels.smartArt
    case "chart":
      return labels.chart
    case "oleObject":
      return drawing.ole?.progId
        ? labels.embeddedObject.replace("{progId}", drawing.ole.progId)
        : labels.oleObject
    case "model3d":
      return labels.model3d
    case "shape":
    default:
      return labels.shape
  }
}

function drawingAriaLabel(
  drawing: DrawingRunNode,
  labels: DocxDrawingLabels,
): string {
  return drawing.alt || drawing.name || drawingKindLabel(drawing, labels)
}

function textBoxParagraphStyle(): Record<string, string | number | undefined> {
  return {
    margin: 0,
    padding: 0,
    whiteSpace: "pre-wrap",
  }
}

function textBoxContentVNodes(
  content: TextBoxContent,
  keySeed: string,
  theme: "light" | "dark",
): VNode[] {
  return content.paragraphs.map((paragraph, paragraphIndex) => {
    const style = textBoxParagraphStyle()
    return h(
      "div",
      {
        key: `${keySeed}-textbox-p-${paragraphIndex}`,
        "data-docx-drawing-textbox-paragraph": "true",
        style: {
          ...style,
          textAlign: paragraph.align ?? undefined,
          lineHeight: `${paragraph.lineHeightPx}px`,
        },
      },
      paragraph.segments.map((segment, segmentIndex) =>
        h(
          "span",
          {
            key: `${keySeed}-textbox-p-${paragraphIndex}-s-${segmentIndex}`,
            style: segment.style
              ? (runStyleToCss(segment.style, theme) as Record<string, string | number>)
              : undefined,
          },
          segment.text,
        ),
      ),
    )
  })
}

function drawingPreviewSrc(drawing: DrawingRunNode): string | undefined {
  if (drawing.src) return drawing.src
  if (drawing.kind === "chart") {
    const { widthPx, heightPx } = drawingObjectSizePx(drawing)
    return chartPreviewDataUri(drawing.chart, widthPx, heightPx)
  }
  return undefined
}

/** Renders one drawing object child of a paragraph. */
export function renderDrawingRun(ctx: DrawingRunRenderContext): VNode {
  const { drawing, key, childIndex } = ctx
  const labels = ctx.labels ?? officeMessagesEn.docx.drawing
  const theme = ctx.documentTheme ?? "light"
  const label = drawingKindLabel(drawing, labels)
  const ariaLabel = drawingAriaLabel(drawing, labels)
  const previewSrc = drawingPreviewSrc(drawing)

  const attributes: Record<string, unknown> = {
    key,
    role: "img",
    "aria-label": ariaLabel,
    "data-docx-drawing": "true",
    "data-docx-drawing-kind": drawing.kind,
    "data-docx-drawing-child-index": String(childIndex),
    "data-docx-drawing-behind": drawingObjectIsBehindText(drawing) ? "true" : undefined,
    ...(ctx.paragraphSourceId
      ? { "data-docx-source-id": `${ctx.paragraphSourceId}/child/${childIndex}` }
      : {}),
    style: {
      ...drawingObjectPlaceholderStyle(drawing),
      ...drawingObjectFrameStyle(drawing),
    },
  }

  if (previewSrc) {
    return h("span", attributes, [
      h("img", {
        key: `${key}-preview`,
        src: previewSrc,
        alt: ariaLabel,
        draggable: false,
        loading: "lazy",
        decoding: "async",
        "data-docx-drawing-preview": "true",
        style: { width: "100%", height: "100%", objectFit: "contain" },
      }),
    ])
  }

  if (drawing.kind === "textBox" && drawing.textBox) {
    return h("span", attributes, [
      h(
        "span",
        {
          key: `${key}-textbox`,
          "data-docx-drawing-textbox": "true",
          style: { ...drawingObjectContentStyle(), alignItems: "flex-start" },
        },
        textBoxContentVNodes(drawing.textBox, key, theme),
      ),
    ])
  }

  const detail =
    drawing.kind === "chart"
      ? drawing.chart?.title ?? `${label}`
      : drawing.ole?.progId && drawing.kind === "oleObject"
        ? label
        : drawing.name && drawing.name !== label
          ? drawing.name
          : label

  return h(
    "span",
    attributes,
    h("span", {
      key: `${key}-label`,
      "data-docx-drawing-placeholder": "true",
      style: drawingObjectContentStyle(),
    }, detail),
  )
}

const MATH_STYLE: Record<string, string | number> = {
  fontFamily: "\"Cambria Math\", \"Latin Modern Math\", serif",
  fontStyle: "italic",
  lineHeight: 1.2,
}

/** Renders one equation child of a paragraph as its linearized text. */
export function renderMathRun(ctx: MathRunRenderContext): VNode {
  const { math, key, childIndex } = ctx
  const styleOverrides: Record<string, string | number> = {}
  if (math.style?.fontSizePt) {
    styleOverrides.fontSize = `${math.style.fontSizePt}pt`
  }
  if (math.style?.color) styleOverrides.color = math.style.color

  return h(
    "span",
    {
      key,
      role: "math",
      "aria-label": math.text,
      "data-docx-math": "true",
      "data-docx-math-display": math.display ? "true" : undefined,
      "data-docx-drawing-child-index": String(childIndex),
      ...(ctx.paragraphSourceId
        ? { "data-docx-source-id": `${ctx.paragraphSourceId}/child/${childIndex}` }
        : {}),
      style: {
        ...MATH_STYLE,
        ...styleOverrides,
        display: math.display ? "block" : "inline",
        margin: math.display ? "4px 0" : undefined,
        textAlign: math.display ? "center" : undefined,
      },
    },
    math.text || officeMessagesEn.docx.drawing.equation,
  )
}

/** Drop-cap metrics for a paragraph's first letter, or undefined. */
export function dropCapFirstTextChild(
  paragraph: ParagraphNode,
): { childIndex: number; firstChar: string; remainder: string } | undefined {
  for (let index = 0; index < paragraph.children.length; index += 1) {
    const child = paragraph.children[index]
    if (!child || child.type !== "text") continue
    if (child.text.length === 0) continue
    const first = Array.from(child.text)[0] ?? ""
    return {
      childIndex: index,
      firstChar: first,
      remainder: child.text.slice(first.length),
    }
  }
  return undefined
}
