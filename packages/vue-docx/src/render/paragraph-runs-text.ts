// Text run rendering for renderParagraphRuns.
// Upstream editor.tsx: 18327-19082 (appendPlainTextWithSoftBreakControl,
// renderRun text branch, hyperlink/note markers).
//
// Translates JSX to Vue h() calls.

import { h, type VNode } from "vue"
import { sanitizeOfficeLink } from "@arcships/office-runtime/link-policy"
import type { ParagraphNode, TextRunNode } from "@arcships/docx-core"
import {
  runStyleToCss,
  linkStyleToCss,
  trackedInlineStyle,
  commentHighlightStyle,
  updateEstimatedLineWidthPxForText,
  resolveTabSpacerWidthPx,
  noteMarkerLabel,
  isJustifiedAlignment,
  docxScriptSegmentStyle,
  segmentTextByDocxScriptFont,
} from "@arcships/docx-core"
import type { DocxDocumentTheme } from "@arcships/docx-core"
import type { DocxScriptFontSegment } from "@arcships/docx-core"
import type { ParagraphTrackedInlineChange } from "@arcships/docx-core"

/** Replace only the font properties of a span with the script slot's font. */
function scriptSegmentStyle(
  baseStyle: Record<string, string | number | undefined>,
  runStyle: TextRunNode["style"] | undefined,
  segment: DocxScriptFontSegment,
  theme: DocxDocumentTheme
): Record<string, string | number | undefined> {
  const scriptStyle = runStyleToCss(
    docxScriptSegmentStyle(runStyle, segment),
    theme
  )
  return {
    ...baseStyle,
    ...(scriptStyle.fontFamily ? { fontFamily: scriptStyle.fontFamily } : {}),
    ...(scriptStyle.fontSize ? { fontSize: scriptStyle.fontSize } : {}),
    fontKerning: scriptStyle.fontKerning,
    fontFeatureSettings: scriptStyle.fontFeatureSettings,
  }
}

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
  /** Pre-formatted marker labels keyed by raw w:id (see buildDocumentPageNotes). */
  noteMarkerLabels?: { footnote?: Map<number, string>; endnote?: Map<number, string> }
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
      trackedInlineChange,
      ctx.documentTheme
    ),
    ...commentHighlightStyle(ctx.documentTheme, undefined),
  }

  const noteRef = child.noteReference
  const suppliedLabel = noteRef
    ? noteRef.kind === "footnote"
      ? ctx.noteMarkerLabels?.footnote?.get(noteRef.id)
      : ctx.noteMarkerLabels?.endnote?.get(noteRef.id)
    : undefined
  // Fall back to raw display indexes / w:id for callers that supply no labels.
  const noteLabel = suppliedLabel ?? noteMarkerLabel(noteRef, new Map(), new Map())

  // Note marker (footnote/endnote)
  if (noteLabel) {
    return {
      vnodes: [
        h(
          "span",
          {
            key,
            ...annotationAttributes,
            "data-docx-note-marker": "true",
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
  const linkHref = child.link ? sanitizeOfficeLink(child.link) : undefined
  if (linkHref) {
    return renderHyperlinkRun(
      child,
      linkHref,
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
  linkHref: string,
  key: string,
  textOverride: string,
  ctx: TextRunRenderContext,
  annotationAttributes: Record<string, string>,
  trackedInlineChange: ParagraphTrackedInlineChange | undefined
): TextRunRenderResult {
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
          trackedInlineChange,
          ctx.documentTheme
        ),
        ...commentHighlightStyle(ctx.documentTheme, undefined),
      }
    : {
        ...trackedInlineStyle(
          linkStyleToCss(child.style, ctx.documentTheme),
          trackedInlineChange,
          ctx.documentTheme
        ),
        ...commentHighlightStyle(ctx.documentTheme, undefined),
      }

  const scriptSegments = segmentTextByDocxScriptFont(textOverride, child.style)
  const content =
    scriptSegments.length <= 1
      ? textOverride
      : scriptSegments.map((segment) =>
          h(
            "span",
            {
              key: `${key}-script-${segment.startOffset}`,
              style: scriptSegmentStyle(
                linkStyle,
                child.style,
                segment,
                ctx.documentTheme
              ),
            },
            segment.text
          )
        )

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
        content
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
    isJustifiedAlignment(ctx.paragraph.style?.align) &&
    textOverride.includes("\n")

  if (!shouldControlSoftBreakStretch) {
    const scriptSegments = segmentTextByDocxScriptFont(
      textOverride,
      child.style
    )
    const vnodes =
      scriptSegments.length <= 1
        ? [
            h(
              "span",
              {
                key,
                ...annotationAttributes,
                style: textStyle,
              },
              textOverride
            ),
          ]
        : scriptSegments.map((segment) =>
            h(
              "span",
              {
                key: `${key}-script-${segment.startOffset}`,
                ...annotationAttributes,
                style: scriptSegmentStyle(
                  textStyle,
                  child.style,
                  segment,
                  ctx.documentTheme
                ),
              },
              segment.text
            )
          )
    return {
      vnodes,
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
      const lineStyle = isLastSegment
        ? {
            ...textStyle,
            display: "inline-block",
            maxWidth: "100%",
            whiteSpace: "pre-wrap",
            verticalAlign: "baseline",
          }
        : textStyle
      const scriptSegments = segmentTextByDocxScriptFont(
        segment,
        child.style
      )
      if (scriptSegments.length <= 1) {
        vnodes.push(
          h(
            "span",
            {
              key: `${key}-segment-${segmentIndex}`,
              ...annotationAttributes,
              style: lineStyle,
            },
            segment
          )
        )
      } else {
        for (const scriptSegment of scriptSegments) {
          vnodes.push(
            h(
              "span",
              {
                key: `${key}-segment-${segmentIndex}-script-${scriptSegment.startOffset}`,
                ...annotationAttributes,
                style: scriptSegmentStyle(
                  lineStyle,
                  child.style,
                  scriptSegment,
                  ctx.documentTheme
                ),
              },
              scriptSegment.text
            )
          )
        }
      }
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
