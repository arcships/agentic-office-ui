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
import { HIGHLIGHT_TO_CSS, SCRIPT_FONT_SCALE } from "./constants";
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

export function cssFontFamily(fontFamily?: string): string | undefined {
  if (!fontFamily) {
    return undefined;
  }

  const trimmed = fontFamily.trim();
  if (!trimmed) {
    return undefined;
  }

  if (trimmed.includes(",")) {
    return trimmed;
  }

  const normalized = trimmed.replace(/^['"]+|['"]+$/g, "");
  if (!normalized) {
    return undefined;
  }

  const escaped = normalized.replace(/"/g, '\\"');
  const familyToken = /\s/.test(escaped) ? `"${escaped}"` : escaped;
  const lower = normalized.toLowerCase();
  const genericFamily = /(mono|consolas|courier|menlo|code)/.test(lower)
    ? "monospace"
    : /(times|cambria|georgia|garamond|baskerville|serif)/.test(lower)
    ? "serif"
    : "sans-serif";

  return `${familyToken}, ${genericFamily}`;
}

export function runStyleToCss(
  style?: TextRunNode["style"],
  documentTheme: DocxDocumentTheme = "light"
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
  const borderType = style?.runBorder?.type?.trim().toLowerCase();
  const borderStyle =
    borderType === "single"
      ? "solid"
      : borderType === "nil" || borderType === "none"
      ? undefined
      : borderType;
  const borderWidthPx =
    Number.isFinite(style?.runBorder?.sizeEighthPt) &&
    (style?.runBorder?.sizeEighthPt as number) > 0
      ? Math.max(
          1,
          Number(
            (
              ((style?.runBorder?.sizeEighthPt as number) / 8) *
              (96 / 72)
            ).toFixed(2)
          )
        )
      : borderStyle
      ? 1
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
    fontFamily: cssFontFamily(style?.fontFamily),
    letterSpacing: Number.isFinite(style?.characterSpacingTwips)
      ? `${Number(
          ((style?.characterSpacingTwips as number) / 20).toFixed(3)
        )}pt`
      : undefined,
    verticalAlign,
    display: borderStyle ? "inline-block" : undefined,
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
  if (typeof baseDecoration === "string") {
    baseDecoration
      .split(/\s+/)
      .map((token) => token.trim())
      .filter(Boolean)
      .forEach((token) => tokens.add(token));
  }

  decoration
    .split(/\s+/)
    .map((token) => token.trim())
    .filter(Boolean)
    .forEach((token) => tokens.add(token));

  return tokens.size > 0 ? Array.from(tokens).join(" ") : "none";
}

export function trackedInlineStyle(
  baseStyle: Record<string, string | number | undefined>,
  change: ParagraphTrackedInlineChange | undefined
): Record<string, string | number | undefined> {
  if (!change) {
    return baseStyle;
  }

  if (change.kind === "insertion" || change.kind === "move-to") {
    const accentColor = change.kind === "move-to" ? "#70ad47" : "#dc2626";
    return {
      ...baseStyle,
      color: accentColor,
      textDecoration: mergeTextDecorations(
        baseStyle.textDecoration,
        "underline"
      ),
    };
  }

  return baseStyle;
}

export function trackedDeletedStyle(
  documentTheme: DocxDocumentTheme,
  baseRunStyle?: TextRunNode["style"] | FormFieldRunNode["style"]
): Record<string, string | number | undefined> {
  const baseStyle = runStyleToCss(baseRunStyle, documentTheme);
  return {
    ...baseStyle,
    color: documentTheme === "dark" ? "#fca5a5" : "#b91c1c",
    textDecoration: mergeTextDecorations(
      baseStyle.textDecoration,
      "line-through"
    ),
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

  return {
    ...(baseStyle ?? {}),
    display: "inline-flex",
    alignItems: "baseline",
    justifyContent: "flex-end",
    verticalAlign: "baseline",
    width: markerBoxWidthPx ? `${markerBoxWidthPx}px` : undefined,
    minWidth: markerBoxWidthPx ? `${markerBoxWidthPx}px` : "1.1em",
    marginRight: markerGapPx,
    whiteSpace: "pre",
    fontFamily: cssFontFamily(
      numberingLabel.fontFamily ?? numberingLabel.style?.fontFamily
    ),
    color: themedRunColor(
      numberingLabel.color ?? numberingLabel.style?.color,
      documentTheme
    ),
  };
}
