// Form field run rendering for renderParagraphRuns.
// Upstream editor.tsx: 18609-18637 (form-field branch of renderRun).
//
// Translates JSX to Vue h() calls.

import { h, type VNode } from "vue"
import type { FormFieldRunNode } from "@arcships/docx-core"
import {
  formFieldDisplayValue,
  runStyleToCss,
  linkStyleToCss,
  trackedInlineStyle,
  commentHighlightStyle,
  updateEstimatedLineWidthPxForText,
  resolveTabSpacerWidthPx,
} from "@arcships/docx-core"
import type { DocxDocumentTheme } from "@arcships/docx-core"
import type { ParagraphTrackedInlineChange } from "@arcships/docx-core"

export interface FieldRunRenderContext {
  useTabLeaderLayout: boolean
  useAnchoredTabLayout: boolean
  onInternalLinkClick?: (bookmarkName: string) => void
  documentTheme: DocxDocumentTheme
  approximateLineWidthPx: number
  tabStopPositionsPx: number[]
  fallbackTabWidthPx: number
  checkboxChoiceRow: boolean
}

export interface FieldRunRenderResult {
  vnode: VNode
  advanceLineWidthPx?: number
}

/**
 * Render a single form-field run child as a Vue VNode.
 */
export function renderFormFieldRun(
  child: FormFieldRunNode,
  key: string,
  ctx: FieldRunRenderContext,
  annotationAttributes: Record<string, string>,
  trackedInlineChange: ParagraphTrackedInlineChange | undefined
): FieldRunRenderResult {
  const textValue = (formFieldDisplayValue(child) || "\u00a0").replace(
    /\t/g,
    " "
  )

  if (child.link) {
    return renderLinkedFormField(child, key, textValue, ctx, annotationAttributes, trackedInlineChange)
  }

  return renderPlainFormField(child, key, textValue, ctx, annotationAttributes, trackedInlineChange)
}

function renderLinkedFormField(
  child: FormFieldRunNode,
  key: string,
  textValue: string,
  ctx: FieldRunRenderContext,
  annotationAttributes: Record<string, string>,
  trackedInlineChange: ParagraphTrackedInlineChange | undefined
): FieldRunRenderResult {
  const linkHref = child.link!
  const isInternalLink = linkHref.startsWith("#")
  const linkStyle = {
    ...trackedInlineStyle(
      linkStyleToCss(child.style, ctx.documentTheme),
      trackedInlineChange
    ),
    ...commentHighlightStyle(ctx.documentTheme, undefined),
  }

  return {
    vnode: h(
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
      textValue
    ),
    advanceLineWidthPx: updateEstimatedLineWidthPxForText(
      ctx.approximateLineWidthPx,
      textValue,
      child.style
    ),
  }
}

function renderPlainFormField(
  child: FormFieldRunNode,
  key: string,
  textValue: string,
  ctx: FieldRunRenderContext,
  annotationAttributes: Record<string, string>,
  trackedInlineChange: ParagraphTrackedInlineChange | undefined
): FieldRunRenderResult {
  const trackedStyle = {
    ...trackedInlineStyle(
      runStyleToCss(child.style, ctx.documentTheme),
      trackedInlineChange
    ),
    ...commentHighlightStyle(ctx.documentTheme, undefined),
  }

  if (textValue === "\t" && !ctx.useTabLeaderLayout && !ctx.useAnchoredTabLayout) {
    const tabWidthPx = resolveTabSpacerWidthPx(
      ctx.tabStopPositionsPx,
      ctx.approximateLineWidthPx,
      ctx.fallbackTabWidthPx,
      ctx.checkboxChoiceRow
    )
    const hasUnderline = Boolean(child.style?.underline)
    return {
      vnode: h(
        "span",
        {
          key,
          ...annotationAttributes,
          style: {
            ...trackedStyle,
            display: "inline-block",
            width: tabWidthPx,
            minWidth: tabWidthPx,
            whiteSpace: "pre",
            textDecoration: hasUnderline ? "none" : trackedStyle.textDecoration,
            borderBottom: hasUnderline ? "1px solid currentColor" : undefined,
            lineHeight: "1em",
          },
        },
        "\u00a0"
      ),
    }
  }

  return {
    vnode: h(
      "span",
      {
        key,
        ...annotationAttributes,
        style: trackedStyle,
      },
      textValue
    ),
    advanceLineWidthPx: updateEstimatedLineWidthPxForText(
      ctx.approximateLineWidthPx,
      textValue,
      child.style
    ),
  }
}
