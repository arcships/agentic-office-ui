// docx-style-preview.ts — CSS for the paragraph-style hover preview.
//
// The toolbar select cannot render what a style looks like, so hovering a
// style in the preview list shows a sample paragraph built from the style's own
// run/paragraph properties. The mapping is pure so hosts can reuse it for their
// own style gallery.

import type { DocxDocumentTheme, ParagraphStyleDefinition } from "@arcships/docx-core"
import { runStyleToCss } from "@arcships/docx-core"

export interface DocxStylePreview {
  /** CSS for the sample text. */
  text: Record<string, string | number | undefined>
  /** CSS for the sample paragraph box (alignment, spacing, shading). */
  paragraph: Record<string, string | number | undefined>
}

function ptToPx(value: number | undefined): number | undefined {
  if (value === undefined || !Number.isFinite(value) || value <= 0) return undefined
  return Math.round(value * (96 / 72) * 100) / 100
}

function spacingToPx(twips: number | undefined): number | undefined {
  if (twips === undefined || !Number.isFinite(twips)) return undefined
  return Math.round((twips / 15) * 100) / 100
}

/**
 * Preview CSS for one paragraph style.
 *
 * `undefined` (the picker's "Normal" entry) previews the document default: the
 * caller keeps its own body typography instead of forcing a look.
 */
export function docxStylePreviewCss(
  style: ParagraphStyleDefinition | undefined,
  documentTheme: DocxDocumentTheme = "light"
): DocxStylePreview {
  if (!style) return { text: {}, paragraph: {} }

  const text: Record<string, string | number | undefined> = {
    ...runStyleToCss(style.runStyle, documentTheme),
    fontWeight: style.runStyle?.bold || style.headingLevel ? "700" : undefined,
  }
  const fontSizePx = ptToPx(style.runStyle?.fontSizePt)
  if (fontSizePx !== undefined) text.fontSize = `${fontSizePx}px`

  const paragraph: Record<string, string | number | undefined> = {
    textAlign: style.align,
    backgroundColor: style.backgroundColor,
  }
  const before = spacingToPx(style.spacing?.beforeTwips)
  const after = spacingToPx(style.spacing?.afterTwips)
  if (before !== undefined) paragraph.marginTop = `${before}px`
  if (after !== undefined) paragraph.marginBottom = `${after}px`
  if (style.indent?.leftTwips !== undefined) {
    const left = spacingToPx(style.indent.leftTwips)
    if (left !== undefined) paragraph.marginLeft = `${left}px`
  }

  return { text, paragraph }
}
