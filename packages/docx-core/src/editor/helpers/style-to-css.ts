// Run-style to CSS conversion (font, color, decoration, highlight, link,
// tracked-change inline styling).
// Upstream editor.tsx: lines 14888-14904 (resolveHighlightColor) and
// 15636-15863 (themedRunColor ... trackedDeletedStyle).

import type {
  FormFieldRunNode,
  NumberingDefinitionSet,
  ParagraphNode,
  TextRunNode
} from "../../engine/types";
import {
  normalizeBorderType,
  ooxmlBorderTypeToCss,
  ooxmlBorderWidthPx
} from "../../engine/border-css";
import { HIGHLIGHT_TO_CSS, SCRIPT_FONT_SCALE } from "./constants";
import { resolveDocxFontFeatureSettings } from "../../viewer/font-features";
import { resolveDocxFontKerning } from "../../viewer/font-kerning";
import { resolveDocxTextFontFamily } from "../../viewer/script-fonts";
import type { DocxDocumentTheme } from "./editor-types";
import type { ParagraphTrackedInlineChange } from "./tracked-changes";
import { resolveListParagraphIndent } from "./xml-parsing-extra";
import { twipsToSignedPixels } from "./ooxml-helpers";
import {
  measureTextWidthPx,
  paragraphBaseFontSizePx
} from "./line-height";
import { tableOfContentsLeadingLeftTabStopPx } from "./field-helpers";
import { isTableOfContentsParagraph } from "./paragraph-toc";
import type { ParagraphNumberingLabel } from "./numbering";
import { clampNumber } from "./zoom-utils";

export function resolveHighlightColor(value?: string): string | undefined {
  if (!value) {
    return undefined;
  }

  const normalized = value.trim().toLowerCase();
  if (!normalized) {
    return undefined;
  }

  if (normalized.startsWith("#")) {
    return normalized;
  }

  return HIGHLIGHT_TO_CSS[normalized] ?? normalized;
}

export function themedRunColor(
  color: string | undefined,
  documentTheme: DocxDocumentTheme
): string | undefined {
  if (documentTheme !== "dark") {
    return color;
  }

  if (!color) {
    return "#f3f4f6";
  }

  const normalized = color.trim().toLowerCase();
  if (
    normalized === "#000" ||
    normalized === "#000000" ||
    normalized === "#111111" ||
    normalized === "#111827" ||
    normalized === "black" ||
    normalized === "rgb(0,0,0)" ||
    normalized === "rgb(0, 0, 0)"
  ) {
    return "#f3f4f6";
  }

  return color;
}

function fontFamilyTokens(value: string | undefined): string[] {
  if (!value) {
    return [];
  }

  return value
    .split(",")
    .map((part) => part.trim().replace(/^['"]+|['"]+$/g, ""))
    .filter(
      (part) =>
        part.length > 0 &&
        !/^(sans-serif|serif|monospace)$/i.test(part)
    );
}

export function cssFontFamily(
  fontFamily?: string,
  options?: { eastAsia?: string; complexScript?: string }
): string | undefined {
  const families: string[] = [];
  const seen = new Set<string>();
  for (const candidate of [
    fontFamily,
    options?.eastAsia,
    options?.complexScript
  ]) {
    for (const token of fontFamilyTokens(candidate)) {
      const key = token.toLowerCase();
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      families.push(token);
    }
  }

  if (families.length === 0) {
    return undefined;
  }

  const lower = families.join(" ").toLowerCase();
  const genericFamily = /(mono|consolas|courier|menlo|code)/.test(lower)
    ? "monospace"
    : /(times|cambria|georgia|garamond|baskerville|serif|simsun|song|mincho|ming)/.test(
        lower
      )
    ? "serif"
    : "sans-serif";

  const tokens = families.map((family) => {
    const escaped = family.replace(/"/g, '\\"');
    return /\s/.test(escaped) ? `"${escaped}"` : escaped;
  });

  return [...tokens, genericFamily].join(", ");
}

export function runStyleToCss(
  style?: TextRunNode["style"],
  documentTheme: DocxDocumentTheme = "light",
  /**
   * Text covered by the span. When supplied, the first strong-script slot
   * supplies the family (so a CJK-only form field does not fall back to the
   * ASCII family). Mixed-script text is rendered as one span per script by the
   * DOCX renderer instead.
   */
  text?: string
): Record<string, string | number | undefined> {
  const hasScriptVerticalAlign =
    style?.verticalAlign === "superscript" ||
    style?.verticalAlign === "subscript";
  const verticalAlign =
    style?.verticalAlign === "superscript"
      ? "super"
      : style?.verticalAlign === "subscript"
      ? "sub"
      : undefined;
  const textDecorationTokens = [
    style?.underline ? "underline" : "",
    style?.strike ? "line-through" : "",
  ].filter(Boolean);
  const textDecoration =
    textDecorationTokens.length > 0 ? textDecorationTokens.join(" ") : "none";
  const decorationVariant = style?.underline
    ? style.underlineStyle
    : style?.strike
    ? style.strikeStyle
    : undefined;
  const decorationStyle =
    decorationVariant === "double"
      ? "double"
      : decorationVariant === "dotted" || decorationVariant === "dottedHeavy"
      ? "dotted"
      : decorationVariant === "dash" ||
        decorationVariant === "dashedHeavy" ||
        decorationVariant === "dashLong" ||
        decorationVariant === "dashLongHeavy" ||
        decorationVariant === "dotDash" ||
        decorationVariant === "dashDotHeavy" ||
        decorationVariant === "dotDotDash" ||
        decorationVariant === "dashDotDotHeavy"
      ? "dashed"
      : decorationVariant === "wave" ||
        decorationVariant === "wavyHeavy" ||
        decorationVariant === "wavyDouble"
      ? "wavy"
      : undefined;
  const decorationThickness =
    decorationVariant === "thick" ||
    decorationVariant === "dottedHeavy" ||
    decorationVariant === "dashedHeavy" ||
    decorationVariant === "dashLongHeavy" ||
    decorationVariant === "dashDotHeavy" ||
    decorationVariant === "dashDotDotHeavy" ||
    decorationVariant === "wavyHeavy"
      ? "3px"
      : decorationVariant === "wavyDouble"
      ? "2px"
      : undefined;
  const borderType = normalizeBorderType(style?.runBorder?.type);
  const borderMapping = ooxmlBorderTypeToCss(borderType);
  const borderStyle =
    borderMapping && borderMapping.style !== "none"
      ? borderMapping.style
      : undefined;
  const borderWidthPx =
    borderStyle && style?.runBorder
      ? Math.max(1, ooxmlBorderWidthPx(style.runBorder))
      : undefined;
  const borderPaddingPt =
    Number.isFinite(style?.runBorder?.spacePt) &&
    (style?.runBorder?.spacePt as number) >= 0
      ? Math.max(0, Math.round(style?.runBorder?.spacePt as number))
      : undefined;

  return {
    fontWeight: style?.bold ? 700 : undefined,
    fontStyle: style?.italic ? "italic" : undefined,
    textDecoration,
    textDecorationStyle: decorationStyle,
    textDecorationThickness: decorationThickness,
    textTransform: style?.caps ? "uppercase" : undefined,
    fontVariantCaps:
      !style?.caps && style?.smallCaps ? "small-caps" : undefined,
    color: themedRunColor(style?.color, documentTheme),
    backgroundColor:
      style?.backgroundColor ?? resolveHighlightColor(style?.highlight),
    fontSize: style?.fontSizePt
      ? `${Number(
          (
            style.fontSizePt * (hasScriptVerticalAlign ? SCRIPT_FONT_SCALE : 1)
          ).toFixed(3)
        )}pt`
      : hasScriptVerticalAlign
      ? `${SCRIPT_FONT_SCALE}em`
      : undefined,
    fontFamily:
      (text
        ? cssFontFamily(resolveDocxTextFontFamily(text, style), {
            eastAsia: style?.fontFamilyEastAsia,
            complexScript: style?.fontFamilyComplexScript
          })
        : undefined) ??
      cssFontFamily(style?.fontFamily, {
        eastAsia: style?.fontFamilyEastAsia,
        complexScript: style?.fontFamilyComplexScript
      }),
    fontKerning: resolveDocxFontKerning(style, style?.fontSizePt),
    fontFeatureSettings: resolveDocxFontFeatureSettings(style),
    letterSpacing: Number.isFinite(style?.characterSpacingTwips)
      ? `${Number(
          ((style?.characterSpacingTwips as number) / 20).toFixed(3)
        )}pt`
      : undefined,
    direction:
      style?.direction === "rtl"
        ? "rtl"
        : style?.direction === "ltr"
        ? "ltr"
        : undefined,
    unicodeBidi: style?.direction ? "embed" : undefined,
    textShadow: style?.textShadow
      ? `${style.textShadow.offsetXPt ?? 1}pt ${
          style.textShadow.offsetYPt ?? 1
        }pt ${style.textShadow.blurPt ?? 1}pt ${
          style.textShadow.color ?? "rgba(0, 0, 0, 0.35)"
        }`
      : undefined,
    WebkitTextStroke: style?.textOutline
      ? `${style.textOutline.widthPt ?? 0.5}pt ${
          style.textOutline.color ?? "currentColor"
        }`
      : undefined,
    verticalAlign,
    display:
      borderStyle ||
      (Number.isFinite(style?.characterScalePercent) &&
        (style?.characterScalePercent as number) !== 100)
        ? "inline-block"
        : undefined,
    transform:
      Number.isFinite(style?.characterScalePercent) &&
      (style?.characterScalePercent as number) !== 100
        ? `scaleX(${Number(
            ((style?.characterScalePercent as number) / 100).toFixed(3)
          )})`
        : undefined,
    transformOrigin:
      Number.isFinite(style?.characterScalePercent) &&
      (style?.characterScalePercent as number) !== 100
        ? "left center"
        : undefined,
    borderStyle,
    borderWidth: borderWidthPx ? `${borderWidthPx}px` : undefined,
    borderColor: borderStyle
      ? style?.runBorder?.color ?? "currentColor"
      : undefined,
    ...(borderPaddingPt !== undefined
      ? {
          paddingTop: `${borderPaddingPt}pt`,
          paddingRight: `${borderPaddingPt}pt`,
          paddingBottom: `${borderPaddingPt}pt`,
          paddingLeft: `${borderPaddingPt}pt`,
        }
      : undefined),
    boxDecorationBreak: borderStyle ? "clone" : undefined,
    whiteSpace: "pre-wrap",
  };
}

export function linkStyleToCss(
  style?: TextRunNode["style"],
  documentTheme: DocxDocumentTheme = "light"
): Record<string, string | number | undefined> {
  const base = runStyleToCss(style, documentTheme);
  const resolvedTextDecoration =
    typeof base.textDecoration === "string" &&
    base.textDecoration.trim().length > 0
      ? base.textDecoration
      : "none";
  return {
    ...base,
    color: base.color ?? "inherit",
    textDecoration: resolvedTextDecoration,
  };
}

export function mergeTextDecorations(
  baseDecoration: Record<string, string | number | undefined>["textDecoration"],
  decoration: string
): string {
  const tokens = new Set<string>();
  // `none` is exclusive in CSS: "none underline" is invalid and the whole
  // declaration is dropped. Only emit it when no real decoration remains.
  const addTokens = (value: unknown): void => {
    if (typeof value !== "string") return;
    value
      .split(/\s+/)
      .map((token) => token.trim())
      .filter((token) => token && token !== "none")
      .forEach((token) => tokens.add(token));
  };

  addTokens(baseDecoration);
  addTokens(decoration);

  return tokens.size > 0 ? Array.from(tokens).join(" ") : "none";
}

export function trackedInlineStyle(
  baseStyle: Record<string, string | number | undefined>,
  change: ParagraphTrackedInlineChange | undefined,
  documentTheme: DocxDocumentTheme = "light"
): Record<string, string | number | undefined> {
  if (!change) {
    return baseStyle;
  }

  if (change.kind === "insertion" || change.kind === "move-to") {
    // Insertions use a distinct green accent (the gutter already uses green
    // for inserted content) plus an underline, so they remain readable
    // without relying on colour alone. Deletions stay red and struck through.
    const accentColor =
      change.kind === "move-to"
        ? documentTheme === "dark"
          ? "#86efac"
          : "#70ad47"
        : documentTheme === "dark"
        ? "#4ade80"
        : "#15803d";
    return {
      ...baseStyle,
      color: accentColor,
      textDecoration: mergeTextDecorations(
        baseStyle.textDecoration,
        "underline"
      ),
      textDecorationColor: accentColor,
    };
  }

  if (change.kind === "deletion" || change.kind === "move-from") {
    const accentColor = change.kind === "move-from" ? "#70ad47" : "#b91c1c";
    return {
      ...baseStyle,
      color: accentColor,
      textDecoration: mergeTextDecorations(
        baseStyle.textDecoration,
        "line-through"
      ),
      opacity: 0.95,
    };
  }

  return baseStyle;
}

export function trackedDeletedStyle(
  documentTheme: DocxDocumentTheme,
  baseRunStyle?: TextRunNode["style"] | FormFieldRunNode["style"]
): Record<string, string | number | undefined> {
  const baseStyle = runStyleToCss(baseRunStyle, documentTheme);
  const accentColor = documentTheme === "dark" ? "#fca5a5" : "#b91c1c";
  return {
    ...baseStyle,
    color: accentColor,
    textDecoration: mergeTextDecorations(
      baseStyle.textDecoration,
      "line-through"
    ),
    textDecorationColor: accentColor,
    whiteSpace: baseStyle.whiteSpace ?? "pre-wrap",
    lineHeight: "inherit",
    opacity: 0.95,
  };
}


// ---- Numbering marker CSS helpers ----
// Upstream editor.tsx: 15325-15432

export function resolveNumberingMarkerBoxWidthPx(
  paragraph: ParagraphNode,
  numberingDefinitions?: NumberingDefinitionSet,
  numberingLabel?: ParagraphNumberingLabel
): number | undefined {
  const resolvedIndent = resolveListParagraphIndent(
    paragraph,
    numberingDefinitions
  );
  const hangingIndentPx = twipsToSignedPixels(resolvedIndent?.hangingTwips);
  const firstLineIndentPx = twipsToSignedPixels(resolvedIndent?.firstLineTwips);
  const candidateWidthPx =
    Number.isFinite(hangingIndentPx) && Math.abs(hangingIndentPx as number) > 0
      ? Math.abs(hangingIndentPx as number)
      : Number.isFinite(firstLineIndentPx) && (firstLineIndentPx as number) < 0
      ? Math.abs(firstLineIndentPx as number)
      : undefined;

  const labelTextForWidth = numberingLabel?.imageSrc
    ? numberingLabel.trailingText ?? ""
    : numberingLabel?.text ?? "";
  const normalizedLabelTextForWidth = labelTextForWidth
    .replace(/\t/g, " ")
    .replace(/\u00a0/g, " ");
  const measuredLabelWidthPx =
    normalizedLabelTextForWidth.length > 0
      ? Math.ceil(
          measureTextWidthPx(
            normalizedLabelTextForWidth,
            numberingLabel?.style,
            paragraphBaseFontSizePx(paragraph)
          )
        ) + 4
      : 0;

  const imageWidthPx = numberingLabel?.imageWidthPx;
  const tocLeadingLeftTabStopPx =
    tableOfContentsLeadingLeftTabStopPx(paragraph);
  const tocLabelGapPx = isTableOfContentsParagraph(paragraph) ? 6 : 0;
  const tocMarkerTargetWidthPx =
    Number.isFinite(tocLeadingLeftTabStopPx) &&
    (tocLeadingLeftTabStopPx as number) > 0
      ? Math.min(
          Math.ceil(tocLeadingLeftTabStopPx as number),
          Math.max(
            measuredLabelWidthPx + tocLabelGapPx,
            measuredLabelWidthPx + 16
          )
        )
      : 0;
  const minimumVisualWidthPx = Math.max(
    measuredLabelWidthPx + tocLabelGapPx,
    Number.isFinite(imageWidthPx) && (imageWidthPx as number) > 0
      ? Math.ceil(imageWidthPx as number) + 2
      : 0,
    tocMarkerTargetWidthPx
  );

  if (!Number.isFinite(candidateWidthPx) || (candidateWidthPx as number) < 8) {
    return minimumVisualWidthPx > 0
      ? clampNumber(Math.round(minimumVisualWidthPx), 8, 220)
      : undefined;
  }

  return Math.max(
    minimumVisualWidthPx,
    clampNumber(Math.round(candidateWidthPx as number), 8, 220)
  );
}

export function numberingMarkerStyle(
  paragraph: ParagraphNode,
  numberingDefinitions: NumberingDefinitionSet | undefined,
  numberingLabel: ParagraphNumberingLabel,
  baseStyle: Record<string, string | number | undefined> | undefined,
  documentTheme: DocxDocumentTheme
): Record<string, string | number | undefined> {
  const markerBoxWidthPx = resolveNumberingMarkerBoxWidthPx(
    paragraph,
    numberingDefinitions,
    numberingLabel
  );
  const markerGapPx = isTableOfContentsParagraph(paragraph)
    ? 6
    : markerBoxWidthPx
    ? 0
    : 2;

  const labelAlignment = numberingLabel.alignment;
  const justifyContent =
    labelAlignment === "left" ||
    labelAlignment === "start" ||
    labelAlignment === "low"
      ? "flex-start"
      : labelAlignment === "center"
      ? "center"
      : "flex-end";

  return {
    ...(baseStyle ?? {}),
    display: "inline-flex",
    alignItems: "baseline",
    justifyContent,
    verticalAlign: "baseline",
    width: markerBoxWidthPx ? `${markerBoxWidthPx}px` : undefined,
    minWidth: markerBoxWidthPx ? `${markerBoxWidthPx}px` : "1.1em",
    marginRight: markerGapPx,
    whiteSpace: "pre",
    fontFamily: cssFontFamily(
      numberingLabel.fontFamily ?? numberingLabel.style?.fontFamily,
      {
        eastAsia: numberingLabel.style?.fontFamilyEastAsia,
        complexScript: numberingLabel.style?.fontFamilyComplexScript
      }
    ),
    color: themedRunColor(
      numberingLabel.color ?? numberingLabel.style?.color,
      documentTheme
    ),
  };
}
