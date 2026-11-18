// Paragraph-level style-to-CSS mapping for DocxParagraphHost.
//
// The model carries spacing, indents, line-spacing rules, shading and borders
// (ParagraphStyle), but none of it used to reach the DOM. This module mirrors
// the geometry used by the pagination helpers in docx-core so that the space a
// paragraph reserves while paginating matches the space it occupies on screen:
//
//   vertical   = marginTop + borderTop + paddingTop + content + ...
//   horizontal = borderLeft + paddingLeft + content + paddingRight + ...
//
// Indents are emitted as padding (and first-line/hanging as text-indent), so
// borders stay outside the text box and do not change the text measure.

import {
  DEFAULT_PARAGRAPH_LINE_MULTIPLE,
  DEFAULT_UNSPECIFIED_DOCX_FONT_FAMILY,
  cssFontFamily,
  normalizeFontFamilyToken,
  paragraphBaseFontSizePx,
  paragraphBorderPaddingPx,
  paragraphBorderToCss,
  paragraphDominantFontFamily,
  resolveAutoLineSpacingMultiple,
  resolveListParagraphIndent,
  twipsToPixels,
  twipsToSignedPixels,
} from "@arcships/docx-core"
import type {
  NumberingDefinitionSet,
  ParagraphBorderStyle,
  ParagraphNode,
  ParagraphSpacing,
} from "@arcships/docx-core"

export type ParagraphCssStyle = Record<string, string | number | undefined>

export interface ParagraphStyleCssOptions {
  numberingDefinitions?: NumberingDefinitionSet
  /**
   * Resolved `w:beforeAutospacing` / `w:afterAutospacing` pixels. The caller
   * resolves them with the document compatibility mode and the neighbouring
   * list items, so the margins match what the paginator measured.
   */
  autoSpacingBeforePx?: number
  autoSpacingAfterPx?: number
}

const HEADING_FONT_WEIGHT = 700

function pixelValue(value: number | undefined): string | undefined {
  return value === undefined ? undefined : `${value}px`
}

/**
 * Pixel string with sub-pixel precision, used where a spacing value is
 * fractional (Word's 14 pt automatic gap is 18.67 px at 96 dpi).
 */
function precisePixelValue(value: number | undefined): string | undefined {
  if (value === undefined) return undefined
  return `${Math.round(value * 100) / 100}px`
}

function nonZeroPixelValue(value: number | undefined): string | undefined {
  return value !== undefined && value > 0 ? `${value}px` : undefined
}

function borderValue(
  border: ParagraphBorderStyle | undefined
): string | undefined {
  return border ? paragraphBorderToCss(border) : undefined
}

// paragraphDominantFontFamily() returns a normalized (lower-case) token. Walk
// the runs once more to keep the font's original spelling in the CSS stack.
function paragraphFontFamilyCss(
  paragraph: ParagraphNode
): string | undefined {
  const dominant = paragraphDominantFontFamily(paragraph)
  if (!dominant) {
    return undefined
  }

  for (const child of paragraph.children) {
    if (child.type !== "text" && child.type !== "form-field") {
      continue
    }
    const family = child.style?.fontFamily
    if (family && normalizeFontFamilyToken(family) === dominant) {
      return cssFontFamily(family)
    }
  }

  return cssFontFamily(dominant)
}

/**
 * CSS line-height for an OOXML paragraph spacing rule.
 *
 * - `exact` / `atLeast` use the explicit twips value as a pixel length.
 * - `auto` is a multiple of the paragraph's single line (240 twips per line),
 *   emitted as a unitless CSS line-height so it scales with the font size.
 * - no line rule/value leaves line-height inherited (CSS `normal`).
 */
export function paragraphLineHeightCss(
  spacing: ParagraphSpacing | undefined
): string | number | undefined {
  const lineTwips = spacing?.lineTwips
  if (!Number.isFinite(lineTwips) || (lineTwips as number) <= 0) {
    return undefined
  }

  const lineRule = spacing?.lineRule ?? "auto"
  if (lineRule === "exact" || lineRule === "atLeast") {
    return pixelValue(twipsToPixels(lineTwips))
  }

  return resolveAutoLineSpacingMultiple(
    lineTwips,
    DEFAULT_PARAGRAPH_LINE_MULTIPLE
  )
}

/**
 * Build the inline style object for one paragraph.
 *
 * Callers pass the numbering definitions from the document so list paragraphs
 * pick up the same effective indent the paginator measured.
 */
export function paragraphStyleToCss(
  paragraph: ParagraphNode,
  options: ParagraphStyleCssOptions = {}
): ParagraphCssStyle {
  const style = paragraph.style
  const spacing = style?.spacing
  const borders = style?.borders
  const indent = resolveListParagraphIndent(
    paragraph,
    options.numberingDefinitions
  )

  const leftIndentPx = Math.max(0, twipsToSignedPixels(indent?.leftTwips) ?? 0)
  const rightIndentPx = Math.max(
    0,
    twipsToSignedPixels(indent?.rightTwips) ?? 0
  )
  const firstLineIndentPx = twipsToSignedPixels(indent?.firstLineTwips)
  const hangingIndentPx = twipsToSignedPixels(indent?.hangingTwips)
  const textIndentPx =
    firstLineIndentPx !== undefined
      ? firstLineIndentPx
      : hangingIndentPx !== undefined
      ? -hangingIndentPx
      : undefined

  // Border "space" is the gap between the border stroke and the text. The
  // stroke itself comes from the border-* property, so together they equal the
  // border inset the pagination helpers budget for.
  const topBorderSpacePx = paragraphBorderPaddingPx(borders?.top)
  const rightBorderSpacePx = paragraphBorderPaddingPx(borders?.right)
  const bottomBorderSpacePx = paragraphBorderPaddingPx(borders?.bottom)
  const leftBorderSpacePx = paragraphBorderPaddingPx(borders?.left)

  // `bar` renders as a vertical rule beside quoted paragraphs; use it as the
  // left border when the paragraph has no explicit left border.
  const leftBorder = borders?.left ?? borders?.bar

  const fontFamily =
    paragraphFontFamilyCss(paragraph) ??
    cssFontFamily(DEFAULT_UNSPECIFIED_DOCX_FONT_FAMILY)

  // Automatic (HTML) spacing adds to the explicit `w:before` / `w:after`
  // exactly like the paginator's `paragraphBeforeSpacingPx`.
  const autoBeforeSpacingPx = Math.max(0, options.autoSpacingBeforePx ?? 0)
  const autoAfterSpacingPx = Math.max(0, options.autoSpacingAfterPx ?? 0)
  const explicitBeforeSpacingPx = twipsToPixels(spacing?.beforeTwips)
  const explicitAfterSpacingPx = twipsToPixels(spacing?.afterTwips)
  const beforeSpacingPx =
    explicitBeforeSpacingPx === undefined && autoBeforeSpacingPx === 0
      ? undefined
      : (explicitBeforeSpacingPx ?? 0) + autoBeforeSpacingPx
  const afterSpacingPx =
    explicitAfterSpacingPx === undefined && autoAfterSpacingPx === 0
      ? undefined
      : (explicitAfterSpacingPx ?? 0) + autoAfterSpacingPx

  return {
    marginTop: precisePixelValue(beforeSpacingPx),
    marginBottom: precisePixelValue(afterSpacingPx),
    textAlign: style?.align,
    fontWeight: style?.headingLevel ? HEADING_FONT_WEIGHT : undefined,
    // Keep the DOM base font in sync with the measurement path so line breaks
    // computed by the paginator match the ones the browser produces.
    fontSize: `${paragraphBaseFontSizePx(paragraph)}px`,
    fontFamily,
    lineHeight: paragraphLineHeightCss(spacing),
    paddingLeft: nonZeroPixelValue(leftIndentPx + (leftBorderSpacePx ?? 0)),
    paddingRight: nonZeroPixelValue(
      rightIndentPx + (rightBorderSpacePx ?? 0)
    ),
    paddingTop: nonZeroPixelValue(topBorderSpacePx),
    paddingBottom: nonZeroPixelValue(bottomBorderSpacePx),
    textIndent: pixelValue(textIndentPx),
    backgroundColor: style?.backgroundColor,
    borderTop: borderValue(borders?.top),
    borderRight: borderValue(borders?.right),
    borderBottom: borderValue(borders?.bottom),
    borderLeft: borderValue(leftBorder),
  }
}
