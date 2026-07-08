// Text run rendering for renderParagraphRuns.
// Upstream editor.tsx: 18327-19082 (appendPlainTextWithSoftBreakControl,
// renderRun text branch, hyperlink/note markers).
//
// Translates JSX to Vue h() calls.

import { h, type VNode } from "vue"
import type { ParagraphNode, TextRunNode } from "@extend-ai/docx-core"
import {
  runStyleToCss,
  linkStyleToCss,
  trackedInlineStyle,
  commentHighlightStyle,
  updateEstimatedLineWidthPxForText,
  resolveTabSpacerWidthPx,
  noteMarkerLabel,
} from "@extend-ai/docx-core"
import type { DocxDocumentTheme } from "@extend-ai/docx-core"
import type { ParagraphTrackedInlineChange } from "@extend-ai/docx-core"

export interface TextRunRenderContext {
  documentTheme: DocxDocumentTheme
  paragraph: ParagraphNode
  useTabLeaderLayout: boolean
  useAnchoredTabLayout: boolean
  approximateLineWidthPx: number
  tabStopPositionsPx: number[]
  fallbackTabWidthPx: number
  checkboxChoiceRow: boolean
  onInternalLinkClick?: (bookmarkName: string) => void
  hasPageField: boolean
  hasStyleRefField: boolean
  pageFieldSequence: unknown[]
  pageFieldValueSequence: unknown[]
  styleRefFieldValueSequence: unknown[]
  consumedPageFieldValues: number
  consumedStyleRefFieldValues: number
  pageNumber?: number
  totalPages?: number
  pageNumberFormat?: string
  resolveStyleRefFieldValue?: (target: string) => string | undefined
  tocLinkColor?: string | undefined
}

export interface TextRunRenderResult {
  vnodes: VNode[]
  consumedPageFieldValues: number
  consumedStyleRefFieldValues: number
  newLineWidthPx: number
}

/**
 * Render a text run child as Vue VNodes. Handles soft breaks,
 * hyperlinks, note markers, and tab styling.
 */
export function renderTextRun(
  child: TextRunNode,
  key: string,
  textOverride: string,
  ctx: TextRunRenderContext,
  annotationAttributes: Record<string, string>,
  trackedInlineChange: ParagraphTrackedInlineChange | undefined
): TextRunRenderResult {
  const textStyle = {
    ...trackedInlineStyle(
      runStyleToCss(child.style, ctx.documentTheme),
      trackedInlineChange
    ),
    ...commentHighlightStyle(ctx.documentTheme, undefined),
  }

  const noteLabel = noteMarkerLabel(
    child.noteReference,
    new Map(),
    new Map()
  )

  // Note marker (footnote/endnote)
  if (noteLabel) {
    return {
      vnodes: [
        h(
          "span",
          {
            key,
            ...annotationAttributes,
            style: {
              ...textStyle,
              verticalAlign: "super",
              fontSize: "0.75em",
            },
          },
          noteLabel
        ),
      ],
      consumedPageFieldValues: ctx.consumedPageFieldValues,
      consumedStyleRefFieldValues: ctx.consumedStyleRefFieldValues,
      newLineWidthPx: updateEstimatedLineWidthPxForText(
        ctx.approximateLineWidthPx,
        noteLabel,
        child.style
      ),
    }
  }

  // Tab handling
  if (
    textOverride === "\t" &&
    !ctx.useTabLeaderLayout &&
    !ctx.useAnchoredTabLayout
  ) {
    const tabWidthPx = resolveTabSpacerWidthPx(
      ctx.tabStopPositionsPx,
      ctx.approximateLineWidthPx,
      ctx.fallbackTabWidthPx,
      ctx.checkboxChoiceRow
    )
    const hasUnderline = Boolean(child.style?.underline)
    return {
      vnodes: [
        h(
          "span",
          {
            key,
            ...annotationAttributes,
            style: buildTabStyle(textStyle, hasUnderline, tabWidthPx),
          },
          "\u00a0"
        ),
      ],
      consumedPageFieldValues: ctx.consumedPageFieldValues,
      consumedStyleRefFieldValues: ctx.consumedStyleRefFieldValues,
      newLineWidthPx: ctx.approximateLineWidthPx + tabWidthPx,
    }
  }

  // Hyperlink
  if (child.link) {
    return renderHyperlinkRun(
      child,
      key,
      textOverride,
      ctx,
      annotationAttributes,
      trackedInlineChange
    )
  }

  // Plain text with soft-break control
  return renderPlainTextRun(
    child,
    key,
    textOverride,
    textStyle,
    ctx,
    annotationAttributes
  )
}

// ---- Internal helpers ----

function buildTabStyle(
  textStyle: Record<string, string | number | undefined>,
  hasUnderline: boolean,
  tabWidthPx: number
): Record<string, string | number | undefined> {
  return {
    ...textStyle,
    display: "inline-block",
    width: tabWidthPx,
    minWidth: tabWidthPx,
    whiteSpace: "pre",
    textDecoration: hasUnderline ? "none" : textStyle.textDecoration,
    borderBottom: hasUnderline ? "1px solid currentColor" : undefined,
    lineHeight: "1em",
  }
}

function renderHyperlinkRun(
  child: TextRunNode,
  key: string,
  textOverride: string,
  ctx: TextRunRenderContext,
  annotationAttributes: Record<string, string>,
  trackedInlineChange: ParagraphTrackedInlineChange | undefined
): TextRunRenderResult {
  const linkHref = child.link!
  const isInternalLink = linkHref.startsWith("#")
  const isToc = false // isTableOfContentsParagraph handled by caller

  const linkStyle = isToc
    ? {
        ...trackedInlineStyle(
          {
            ...runStyleToCss(child.style, ctx.documentTheme),
            color: ctx.tocLinkColor
              ? ctx.tocLinkColor
              : "inherit",
            textDecoration: "none",
          },
          trackedInlineChange
        ),
        ...commentHighlightStyle(ctx.documentTheme, undefined),
      }
    : {
        ...trackedInlineStyle(
          linkStyleToCss(child.style, ctx.documentTheme),
          trackedInlineChange
        ),
        ...commentHighlightStyle(ctx.documentTheme, undefined),
      }

  return {
    vnodes: [
      h(
        "a",
        {
          key,
          ...annotationAttributes,
          href: linkHref,
          target: isInternalLink ? undefined : "_blank",
          rel: isInternalLink ? undefined : "noreferrer noopener",
          onMouseDown(event: MouseEvent) {
            if (!isInternalLink || !ctx.onInternalLinkClick) return
            event.preventDefault()
            event.stopPropagation()
          },
          onClick(event: MouseEvent) {
            if (!isInternalLink || !ctx.onInternalLinkClick) return
            event.preventDefault()
            event.stopPropagation()
            ctx.onInternalLinkClick(linkHref.slice(1))
          },
          style: linkStyle,
        },
        textOverride
      ),
    ],
    consumedPageFieldValues: ctx.consumedPageFieldValues,
    consumedStyleRefFieldValues: ctx.consumedStyleRefFieldValues,
    newLineWidthPx: updateEstimatedLineWidthPxForText(
      ctx.approximateLineWidthPx,
      textOverride,
      child.style
    ),
  }
}

function renderPlainTextRun(
  child: TextRunNode,
  key: string,
  textOverride: string,
  textStyle: Record<string, string | number | undefined>,
  ctx: TextRunRenderContext,
  annotationAttributes: Record<string, string>
): TextRunRenderResult {
  const shouldControlSoftBreakStretch =
    ctx.paragraph.style?.align === "justify" && textOverride.includes("\n")

  if (!shouldControlSoftBreakStretch) {
    return {
      vnodes: [
        h(
          "span",
          {
            key,
            ...annotationAttributes,
            style: textStyle,
          },
          textOverride
        ),
      ],
      consumedPageFieldValues: ctx.consumedPageFieldValues,
      consumedStyleRefFieldValues: ctx.consumedStyleRefFieldValues,
      newLineWidthPx: updateEstimatedLineWidthPxForText(
        ctx.approximateLineWidthPx,
        textOverride,
        child.style
      ),
    }
  }

  // Soft break control for justified text
  const segments = textOverride.split("\n")
  const vnodes: VNode[] = []
  let newLineWidth = ctx.approximateLineWidthPx
  segments.forEach((segment, segmentIndex) => {
    const isLastSegment = segmentIndex === segments.length - 1
    if (segmentIndex > 0) {
      vnodes.push(h("br", { key: `${key}-soft-break-${segmentIndex}` }))
    }

    if (segment.length > 0) {
      vnodes.push(
        h(
          "span",
          {
            key: `${key}-segment-${segmentIndex}`,
            ...annotationAttributes,
            style: isLastSegment
              ? {
                  ...textStyle,
                  display: "inline-block",
                  maxWidth: "100%",
                  whiteSpace: "pre-wrap",
                  verticalAlign: "baseline",
                }
              : textStyle,
          },
          segment
        )
      )
    }

    newLineWidth = updateEstimatedLineWidthPxForText(
      newLineWidth,
      segment,
      child.style
    )
    if (!isLastSegment) {
      newLineWidth = 0
    }
  })

  return {
    vnodes,
    consumedPageFieldValues: ctx.consumedPageFieldValues,
    consumedStyleRefFieldValues: ctx.consumedStyleRefFieldValues,
    newLineWidthPx: newLineWidth,
  }
}
