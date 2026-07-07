// Paragraph line-height estimation, font analysis, text-width measurement,
// and tab-stop resolution.
// Upstream editor.tsx: lines 8923-9340, 10056-10389.
//
// Ports the canonical `resolveParagraphFirstLineLeftTabStopsPx` (9310-9323)
// and `resolveParagraphFirstLineOriginPx` (9292-9308), the paragraph
// font-analysis cluster (8923-9157), canvas-backed text-width measurement
// (9165-9277), tab-stop helpers (9279-9340), pretext block-height helpers
// (10056-10202), and the auto-line-spacing / doc-grid / line-height estimators
// (10204-10389). The wrap-line-count cluster (9342-10054) and
// estimateParagraphHeightPx depend on pretext-integration / paragraph-tracked /
// style-block-css; those entry points land in the table-height / pretext-build
// modules which import from this one.

import type {
  FormFieldRunNode,
  ParagraphNode,
  TextRunNode
} from "../../engine/types";
import { twipsToPixels, TWIPS_PER_PIXEL } from "../../viewer/section-layout";
import type { PretextVariableWidthLayout } from "../../viewer/pretext-layout";
import {
  paragraphBaseFontSizePxByParagraph,
  paragraphDominantFontFamilyByParagraph,
  paragraphMeasureCanvasContext,
  setCacheEntry,
  textWidthByFontAndValue,
  setParagraphMeasureCanvasContext
} from "./cache-utils";
import {
  DEFAULT_PARAGRAPH_FONT_SIZE_PT,
  DEFAULT_UNSPECIFIED_DOCX_FONT_FAMILY,
  DEFAULT_TAB_STOP_PX,
  MIN_AUTO_LINE_MULTIPLE,
  SCRIPT_FONT_SCALE,
  TEXT_MEASURE_CACHE_MAX_ENTRIES,
  WORD_AUTO_LINE_SCALE_BLEND_END_MULTIPLE,
  WORD_SINGLE_LINE_AUTO_SCALE,
  WORD_SINGLE_LINE_AUTO_SCALE_SANS,
  WORD_SINGLE_LINE_AUTO_SCALE_SERIF,
  WORD_EMPTY_PARAGRAPH_LINE_SCALE,
  WORD_EMPTY_PARAGRAPH_LINE_SCALE_SERIF,
  WORD_EMPTY_PARAGRAPH_LINE_SCALE_SANS,
  DEFAULT_PARAGRAPH_LINE_MULTIPLE
} from "./constants";
import { twipsToSignedPixels } from "./ooxml-helpers";
import { cssFontFamily } from "./style-to-css";
import { formFieldDisplayValue } from "./paragraph-inspect";
import { isTableOfContentsParagraph } from "./paragraph-toc";
import {
  paragraphHasCheckboxFormField,
  paragraphHasOnlyWhitespaceText,
  paragraphIsFloatingImageAnchorOnly
} from "./paragraph-geometry";

// --- Paragraph font analysis (upstream 8923-9157) ---

export function paragraphDominantFontSizePt(
  paragraph: ParagraphNode
): number | undefined {
  const weightByFontSizePt = new Map<number, number>();

  const addWeight = (fontSizePt: number | undefined, weight: number): void => {
    if (
      !Number.isFinite(fontSizePt) ||
      (fontSizePt as number) <= 0 ||
      weight <= 0
    ) {
      return;
    }

    const key = Number((fontSizePt as number).toFixed(2));
    weightByFontSizePt.set(key, (weightByFontSizePt.get(key) ?? 0) + weight);
  };

  paragraph.children.forEach((child) => {
    if (child.type !== "text" && child.type !== "form-field") {
      return;
    }

    const text =
      child.type === "text"
        ? child.text.replace(/\u2063/g, "")
        : formFieldDisplayValue(child);
    const textWeight = Math.max(1, text.length);
    addWeight(child.style?.fontSizePt, textWeight);
  });

  if (weightByFontSizePt.size === 0) {
    return undefined;
  }

  let dominantFontSizePt: number | undefined;
  let dominantWeight = -1;

  for (const [fontSizePt, weight] of weightByFontSizePt) {
    if (
      weight > dominantWeight ||
      (weight === dominantWeight &&
        (dominantFontSizePt === undefined || fontSizePt > dominantFontSizePt))
    ) {
      dominantWeight = weight;
      dominantFontSizePt = fontSizePt;
    }
  }

  return dominantFontSizePt;
}

export function paragraphBaseFontSizePx(paragraph: ParagraphNode): number {
  const cached = paragraphBaseFontSizePxByParagraph.get(paragraph);
  if (cached !== undefined) {
    return cached;
  }

  const dominantRunFontSizePt = paragraphDominantFontSizePt(paragraph);
  const fontSizePt =
    dominantRunFontSizePt && dominantRunFontSizePt > 0
      ? dominantRunFontSizePt
      : Number.isFinite(paragraph.style?.headingLevel)
      ? DEFAULT_PARAGRAPH_FONT_SIZE_PT +
        Math.max(0, 6 - (paragraph.style?.headingLevel ?? 6))
      : DEFAULT_PARAGRAPH_FONT_SIZE_PT;

  const baseFontSizePx = Math.max(10, Math.round((fontSizePt * 96) / 72));
  paragraphBaseFontSizePxByParagraph.set(paragraph, baseFontSizePx);
  return baseFontSizePx;
}

export function normalizeFontFamilyToken(fontFamily?: string): string | undefined {
  if (!fontFamily) {
    return undefined;
  }

  const [firstToken] = fontFamily.split(",");
  const normalized = firstToken
    ?.trim()
    .replace(/^['"]+|['"]+$/g, "")
    .toLowerCase();
  return normalized || undefined;
}

export function paragraphDominantFontFamily(
  paragraph: ParagraphNode
): string | undefined {
  const cached = paragraphDominantFontFamilyByParagraph.get(paragraph);
  if (cached !== undefined) {
    return cached ?? undefined;
  }

  const weightByFamily = new Map<string, number>();
  const addWeight = (fontFamily: string | undefined, weight: number): void => {
    const normalizedFamily = normalizeFontFamilyToken(fontFamily);
    if (!normalizedFamily) {
      return;
    }

    const safeWeight = Math.max(1, Math.round(weight));
    weightByFamily.set(
      normalizedFamily,
      (weightByFamily.get(normalizedFamily) ?? 0) + safeWeight
    );
  };

  paragraph.children.forEach((child) => {
    if (child.type !== "text" && child.type !== "form-field") {
      return;
    }

    const text =
      child.type === "text"
        ? child.text.replace(/\u2063/g, "")
        : formFieldDisplayValue(child);
    addWeight(child.style?.fontFamily, text.length);
  });

  if (weightByFamily.size === 0) {
    paragraphDominantFontFamilyByParagraph.set(paragraph, null);
    return undefined;
  }

  let dominantFamily: string | undefined;
  let dominantWeight = -1;
  for (const [family, weight] of weightByFamily) {
    if (weight > dominantWeight) {
      dominantFamily = family;
      dominantWeight = weight;
    }
  }

  paragraphDominantFontFamilyByParagraph.set(paragraph, dominantFamily ?? null);
  return dominantFamily;
}

export function singleLineAutoScaleForFontFamily(fontFamily?: string): number {
  const normalized =
    normalizeFontFamilyToken(fontFamily) ?? fontFamily?.toLowerCase();
  if (!normalized) {
    return WORD_SINGLE_LINE_AUTO_SCALE;
  }

  if (
    normalized === "times roman" ||
    normalized === "times new roman" ||
    normalized === "cambria" ||
    normalized === "garamond" ||
    normalized === "georgia" ||
    normalized === "book antiqua" ||
    normalized === "palatino linotype"
  ) {
    return WORD_SINGLE_LINE_AUTO_SCALE_SERIF;
  }

  if (normalized === "arial") {
    return WORD_SINGLE_LINE_AUTO_SCALE_SANS;
  }

  return WORD_SINGLE_LINE_AUTO_SCALE;
}

export function emptyParagraphLineScaleForFontFamily(fontFamily?: string): number {
  const normalized =
    normalizeFontFamilyToken(fontFamily) ?? fontFamily?.toLowerCase();
  if (!normalized) {
    return WORD_EMPTY_PARAGRAPH_LINE_SCALE;
  }

  if (
    normalized === "times roman" ||
    normalized === "times new roman" ||
    normalized === "cambria" ||
    normalized === "garamond" ||
    normalized === "georgia" ||
    normalized === "book antiqua" ||
    normalized === "palatino linotype"
  ) {
    return WORD_EMPTY_PARAGRAPH_LINE_SCALE_SERIF;
  }

  if (normalized === "arial") {
    return WORD_EMPTY_PARAGRAPH_LINE_SCALE_SANS;
  }

  return WORD_EMPTY_PARAGRAPH_LINE_SCALE;
}

// A paragraph whose only content is whitespace and/or floating anchors lays
// out as one empty line at the mark font's natural metrics — floating
// objects never contribute to the line box.
export function paragraphRendersTextFreeLine(paragraph: ParagraphNode): boolean {
  return (
    paragraphHasOnlyWhitespaceText(paragraph) ||
    paragraphIsFloatingImageAnchorOnly(paragraph)
  );
}

export function paragraphContainsExplicitLineBreakText(
  paragraph: ParagraphNode
): boolean {
  return paragraph.children.some(
    (child) => child.type === "text" && /[\r\n]/.test(child.text)
  );
}

export function resolveParagraphSingleLineAutoScale(
  paragraph: ParagraphNode,
  fontFamily?: string
): number {
  if (paragraphRendersTextFreeLine(paragraph)) {
    return emptyParagraphLineScaleForFontFamily(fontFamily);
  }

  const baseScale = singleLineAutoScaleForFontFamily(fontFamily);
  return paragraphHasCheckboxFormField(paragraph)
    ? Math.max(1.08, baseScale)
    : baseScale;
}

// --- Text-width measurement (upstream 9165-9277) ---

export function estimatedGlyphWidthPx(character: string, fontSizePx: number): number {
  if (/\s/.test(character)) {
    return fontSizePx * 0.34;
  }
  if (/[A-Z]/.test(character)) {
    return fontSizePx * 0.66;
  }
  if (/[a-z]/.test(character)) {
    return fontSizePx * 0.55;
  }
  if (/[0-9]/.test(character)) {
    return fontSizePx * 0.57;
  }
  if (/[\u2e80-\u9fff\u3040-\u30ff\uac00-\ud7af]/.test(character)) {
    return fontSizePx;
  }

  return fontSizePx * 0.48;
}

export function fallbackMeasureTextWidthPx(text: string, fontSizePx: number): number {
  let widthPx = 0;
  for (const character of text) {
    widthPx += estimatedGlyphWidthPx(character, fontSizePx);
  }
  return widthPx;
}

export function resolveMeasureFontSizePx(
  style: TextRunNode["style"] | FormFieldRunNode["style"] | undefined,
  paragraphBaseFontPx: number
): number {
  const runFontSizePx =
    Number.isFinite(style?.fontSizePt) && (style?.fontSizePt as number) > 0
      ? ((style?.fontSizePt as number) * 96) / 72
      : paragraphBaseFontPx;
  const verticalAlignScale =
    style?.verticalAlign === "superscript" ||
    style?.verticalAlign === "subscript"
      ? SCRIPT_FONT_SCALE
      : 1;
  return Math.max(8, Math.round(runFontSizePx * verticalAlignScale));
}

export function resolveMeasureFont(
  style: TextRunNode["style"] | FormFieldRunNode["style"] | undefined,
  paragraphBaseFontPx: number
): string {
  const fontStyle = style?.italic ? "italic" : "normal";
  const fontWeight = style?.bold ? "700" : "400";
  const fontSizePx = resolveMeasureFontSizePx(style, paragraphBaseFontPx);
  const fontFamily =
    cssFontFamily(style?.fontFamily) ??
    cssFontFamily(DEFAULT_UNSPECIFIED_DOCX_FONT_FAMILY) ??
    '"Times New Roman", serif';
  return `${fontStyle} normal ${fontWeight} ${fontSizePx}px ${fontFamily}`;
}

export function measureTextWidthPx(
  text: string,
  style: TextRunNode["style"] | FormFieldRunNode["style"] | undefined,
  paragraphBaseFontPx: number
): number {
  if (!text) {
    return 0;
  }

  const font = resolveMeasureFont(style, paragraphBaseFontPx);
  const characterSpacingPx = Number.isFinite(style?.characterSpacingTwips)
    ? (style?.characterSpacingTwips as number) / TWIPS_PER_PIXEL
    : 0;
  const cacheKey = `${font}\u0000${characterSpacingPx}\u0000${text}`;
  const cached = textWidthByFontAndValue.get(cacheKey);
  if (cached !== undefined) {
    return cached;
  }

  const fontSizePx = resolveMeasureFontSizePx(style, paragraphBaseFontPx);
  let measuredWidthPx = fallbackMeasureTextWidthPx(text, fontSizePx);

  if (typeof document !== "undefined") {
    try {
      let context = paragraphMeasureCanvasContext;
      if (!context) {
        const canvas = document.createElement("canvas");
        context = canvas.getContext("2d") ?? undefined;
        if (context) {
          setParagraphMeasureCanvasContext(context);
        }
      }
      if (context) {
        context.font = font;
        measuredWidthPx = context.measureText(text).width;
      }
    } catch {
      // Keep deterministic fallback width when canvas measurement is unavailable.
    }
  }

  if (characterSpacingPx !== 0 && text.length > 1) {
    measuredWidthPx += characterSpacingPx * Math.max(0, text.length - 1);
  }

  setCacheEntry(textWidthByFontAndValue, cacheKey, measuredWidthPx);
  while (textWidthByFontAndValue.size > TEXT_MEASURE_CACHE_MAX_ENTRIES) {
    const firstKey = textWidthByFontAndValue.keys().next().value as
      | string
      | undefined;
    if (!firstKey) {
      break;
    }
    textWidthByFontAndValue.delete(firstKey);
  }

  return measuredWidthPx;
}

// --- Tab-stop resolution (upstream 9279-9340) ---

export function resolveParagraphTabStopsPx(paragraph: ParagraphNode): number[] {
  const stopsPx = (paragraph.style?.tabStops ?? [])
    .map((tabStop) => twipsToPixels(tabStop.positionTwips))
    .filter(
      (value): value is number =>
        Number.isFinite(value) && (value as number) > 0
    )
    .map((value) => Math.round(value))
    .sort((left, right) => left - right);

  return stopsPx;
}

export function resolveParagraphFirstLineOriginPx(paragraph: ParagraphNode): number {
  const leftIndentPx = twipsToSignedPixels(paragraph.style?.indent?.leftTwips);
  const firstLineIndentPx = twipsToSignedPixels(
    paragraph.style?.indent?.firstLineTwips
  );
  const hangingIndentPx = twipsToSignedPixels(
    paragraph.style?.indent?.hangingTwips
  );
  const textIndentPx =
    firstLineIndentPx ??
    (Number.isFinite(hangingIndentPx) ? -(hangingIndentPx as number) : 0);

  return (
    (Number.isFinite(leftIndentPx) ? (leftIndentPx as number) : 0) +
    (Number.isFinite(textIndentPx) ? (textIndentPx as number) : 0)
  );
}

export function resolveParagraphFirstLineLeftTabStopsPx(
  paragraph: ParagraphNode
): number[] {
  const firstLineOriginPx = resolveParagraphFirstLineOriginPx(paragraph);
  return (paragraph.style?.tabStops ?? [])
    .filter((tabStop) => tabStop.alignment !== "right")
    .map((tabStop) => twipsToPixels(tabStop.positionTwips))
    .filter(
      (value): value is number =>
        Number.isFinite(value) && (value as number) > firstLineOriginPx + 0.5
    )
    .map((value) => Math.round(value - firstLineOriginPx))
    .sort((left, right) => left - right);
}

export function resolveNextTabStopPx(
  currentLineWidthPx: number,
  tabStopsPx: number[]
): number {
  const nextExplicit = tabStopsPx.find(
    (stopPx) => stopPx > currentLineWidthPx + 0.5
  );
  if (nextExplicit !== undefined) {
    return nextExplicit;
  }

  const tabStepPx = DEFAULT_TAB_STOP_PX;
  const nextMultiple =
    Math.floor(Math.max(0, currentLineWidthPx) / tabStepPx + 1) * tabStepPx;
  return Math.max(tabStepPx, nextMultiple);
}

// --- Pretext block-height helpers (upstream 10056-10202) ---

export function pretextLayoutContentBottomPx(
  layout: PretextVariableWidthLayout
): number {
  if (layout.lines.length === 0) {
    return 0;
  }

  const lineHeightPx = Math.max(1, Math.round(layout.lineHeightPx ?? 1));
  return (layout.lines[layout.lines.length - 1]?.y ?? 0) + lineHeightPx;
}

export function topAndBottomExclusionCanOverflowParagraphBox(
  layout: PretextVariableWidthLayout
): boolean {
  const containerWidthPx = Math.max(
    1,
    Math.round(layout.containerWidthPx ?? 0)
  );
  if (containerWidthPx <= 0 || layout.lines.length === 0) {
    return false;
  }

  return (layout.exclusions ?? []).some((exclusion) => {
    const left = Math.round(exclusion.left);
    const right = Math.round(exclusion.right);
    return left <= 0 && right >= containerWidthPx;
  });
}

export function wrappedPretextParagraphBlockHeightPx(
  layout: PretextVariableWidthLayout
): number {
  const contentBottomPx = pretextLayoutContentBottomPx(layout);
  if (topAndBottomExclusionCanOverflowParagraphBox(layout)) {
    return Math.max(1, contentBottomPx);
  }

  return Math.max(1, Math.round(layout.height));
}

export function resolvePretextLineRangeContentHeightPx(
  layout: PretextVariableWidthLayout,
  startLineIndex: number,
  endLineIndex: number
): number {
  if (layout.lines.length === 0) {
    return 0;
  }

  const lineHeightPx = Math.max(1, Math.round(layout.lineHeightPx ?? 1));
  const safeStart = Math.max(
    0,
    Math.min(Math.round(startLineIndex), layout.lines.length)
  );
  const safeEnd = Math.max(
    safeStart,
    Math.min(Math.round(endLineIndex), layout.lines.length)
  );
  if (safeEnd <= safeStart) {
    return 0;
  }

  const firstLineTopPx = layout.lines[safeStart]?.y ?? safeStart * lineHeightPx;
  const lastLineTopPx =
    layout.lines[safeEnd - 1]?.y ?? (safeEnd - 1) * lineHeightPx;
  return Math.max(1, Math.round(lastLineTopPx - firstLineTopPx + lineHeightPx));
}

export function resolveMaxPretextLineRangeEndIndexThatFits(
  layout: PretextVariableWidthLayout,
  startLineIndex: number,
  maxEndLineIndex: number,
  availableHeightPx: number
): number {
  const safeStart = Math.max(
    0,
    Math.min(Math.round(startLineIndex), layout.lines.length)
  );
  const safeMaxEnd = Math.max(
    safeStart,
    Math.min(Math.round(maxEndLineIndex), layout.lines.length)
  );
  const safeAvailableHeightPx = Math.max(0, Math.round(availableHeightPx));
  if (safeAvailableHeightPx <= 0 || safeMaxEnd <= safeStart) {
    return safeStart;
  }

  let low = safeStart;
  let high = safeMaxEnd;
  let bestEnd = safeStart;
  while (low <= high) {
    const mid = Math.floor((low + high) / 2);
    const candidateHeightPx = resolvePretextLineRangeContentHeightPx(
      layout,
      safeStart,
      mid
    );
    if (candidateHeightPx <= safeAvailableHeightPx) {
      bestEnd = mid;
      low = mid + 1;
      continue;
    }
    high = mid - 1;
  }

  return bestEnd;
}

// --- Auto-line-spacing / doc-grid / line-height estimation (upstream 10204-10389) ---

export function resolveAutoLineSpacingMultiple(
  lineTwips: number | undefined,
  fallbackMultiple: number
): number {
  if (!Number.isFinite(lineTwips)) {
    return Math.max(MIN_AUTO_LINE_MULTIPLE, fallbackMultiple);
  }

  return Math.max(MIN_AUTO_LINE_MULTIPLE, (lineTwips as number) / 240);
}

export function autoLineHeightScaleForMultiple(
  multiple: number,
  singleLineScale: number
): number {
  const safeSingleLineScale = Math.max(
    MIN_AUTO_LINE_MULTIPLE,
    Number.isFinite(singleLineScale)
      ? singleLineScale
      : WORD_SINGLE_LINE_AUTO_SCALE
  );
  if (!Number.isFinite(multiple)) {
    return safeSingleLineScale;
  }

  if (multiple <= 1) {
    return safeSingleLineScale;
  }

  if (multiple >= WORD_AUTO_LINE_SCALE_BLEND_END_MULTIPLE) {
    return 1;
  }

  const blendProgress =
    (multiple - 1) / (WORD_AUTO_LINE_SCALE_BLEND_END_MULTIPLE - 1);
  return Number(
    (safeSingleLineScale + (1 - safeSingleLineScale) * blendProgress).toFixed(4)
  );
}

export function calibrateAutoLineSpacingMultiple(
  multiple: number,
  fontFamily?: string,
  singleLineScaleOverride?: number
): number {
  const normalizedMultiple = Math.max(MIN_AUTO_LINE_MULTIPLE, multiple);
  const singleLineScale =
    singleLineScaleOverride ?? singleLineAutoScaleForFontFamily(fontFamily);
  return Math.max(
    MIN_AUTO_LINE_MULTIPLE,
    Number(
      (
        normalizedMultiple *
        autoLineHeightScaleForMultiple(normalizedMultiple, singleLineScale)
      ).toFixed(3)
    )
  );
}

export function paragraphDocGridSnapState(
  paragraph: ParagraphNode
): "inherit" | "snap" | "disable" {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml) {
    return "inherit";
  }

  const paragraphPropertiesXml =
    sourceXml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/i)?.[0] ??
    sourceXml.match(/<w:pPr\b[^>]*\/>/i)?.[0] ??
    "";
  if (!paragraphPropertiesXml) {
    return "inherit";
  }

  const snapToGridTag =
    paragraphPropertiesXml.match(/<w:snapToGrid\b[^>]*\/?>/i)?.[0] ?? "";
  if (!snapToGridTag) {
    return "inherit";
  }

  const valueMatch = snapToGridTag
    .match(/\bw:val="([^"]+)"/i)?.[1]
    ?.trim()
    .toLowerCase();
  if (!valueMatch) {
    return "snap";
  }

  return valueMatch === "0" || valueMatch === "false" || valueMatch === "off"
    ? "disable"
    : "snap";
}

export function resolveParagraphDocGridLinePitchPx(
  paragraph: ParagraphNode,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false
): number | undefined {
  const docGridSnapState = paragraphDocGridSnapState(paragraph);
  if (
    disableDocGridSnap ||
    docGridSnapState === "disable" ||
    !Number.isFinite(docGridLinePitchPx) ||
    (docGridLinePitchPx as number) <= 0
  ) {
    return undefined;
  }

  return Math.max(1, Math.round(docGridLinePitchPx as number));
}

export function estimateParagraphLineHeightPx(
  paragraph: ParagraphNode,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false
): number {
  const lineTwips = paragraph.style?.spacing?.lineTwips;
  const lineRule = paragraph.style?.spacing?.lineRule ?? "auto";
  const docGridMinimumLineHeightPx = resolveParagraphDocGridLinePitchPx(
    paragraph,
    docGridLinePitchPx,
    disableDocGridSnap
  );
  const baseFontPx = paragraphBaseFontSizePx(paragraph);
  const baseFontFamily = paragraphDominantFontFamily(paragraph);
  const singleLineScale = resolveParagraphSingleLineAutoScale(
    paragraph,
    baseFontFamily
  );
  const defaultLineMultiple = isTableOfContentsParagraph(paragraph)
    ? 1.05
    : DEFAULT_PARAGRAPH_LINE_MULTIPLE;
  const normalLineHeightPx = Math.max(
    1,
    Math.round(
      baseFontPx *
        calibrateAutoLineSpacingMultiple(
          DEFAULT_PARAGRAPH_LINE_MULTIPLE,
          baseFontFamily,
          singleLineScale
        )
    )
  );

  if (lineRule !== "auto" && Number.isFinite(lineTwips)) {
    const explicitLineHeightPx = Math.max(1, twipsToPixels(lineTwips) ?? 1);
    if (lineRule === "exact") {
      return explicitLineHeightPx;
    }

    return Math.max(
      explicitLineHeightPx,
      normalLineHeightPx,
      docGridMinimumLineHeightPx ?? 0
    );
  }

  const resolvedAutoMultiple = resolveAutoLineSpacingMultiple(
    lineTwips,
    defaultLineMultiple
  );
  // A text-free paragraph renders exactly one line at the mark font's natural
  // metrics, and an explicit auto multiple scales that natural line (Word
  // semantics). The wrapped-text blend toward bare font-size lines exists to
  // offset wrapped-line overcounting, which cannot happen here.
  const multiple = paragraphRendersTextFreeLine(paragraph)
    ? Math.max(
        MIN_AUTO_LINE_MULTIPLE,
        Number((resolvedAutoMultiple * singleLineScale).toFixed(3))
      )
    : calibrateAutoLineSpacingMultiple(
        resolvedAutoMultiple,
        baseFontFamily,
        singleLineScale
      );
  const autoLineHeightPx = Math.max(1, Math.round(baseFontPx * multiple));
  const minimumReadableAutoLineHeightPx = paragraph.style?.numbering
    ? Math.ceil(baseFontPx)
    : 0;
  return Math.max(
    autoLineHeightPx,
    minimumReadableAutoLineHeightPx,
    docGridMinimumLineHeightPx ?? 0
  );
}
