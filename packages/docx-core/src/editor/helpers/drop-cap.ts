// Drop-cap detection, font-size resolution, text-advance-width estimation,
// tab-spacer width, and interactive form-field width sizing.
// Upstream editor.tsx: lines 543-615, 7625-7935.
//
// These helpers compute per-character advance-width estimates (cache-backed),
// resolve drop-cap visual sizing, and size interactive form fields (checkbox /
// date / dropdown / default) by estimated text width. estimateTextAdvanceWidthPx
// is the runtime-functional fallback used by synthetic-text-box fit-to-width
// scaling and form-field intrinsic width sizing.

import type {
  FormFieldRunNode,
  ParagraphNode,
  TextRunNode
} from "../../engine/types";
import {
  estimatedTextAdvanceWidthByFontAndValue,
  paragraphDropCapBySourceXml,
  setCacheEntry
} from "./cache-utils";
import {
  DEFAULT_PARAGRAPH_FONT_SIZE_PT,
  TEXT_MEASURE_CACHE_MAX_ENTRIES
} from "./constants";
import { xmlAttribute } from "./ooxml-helpers";

export function paragraphDropCap(
  paragraph: ParagraphNode
): NonNullable<NonNullable<ParagraphNode["style"]>["dropCap"]> | undefined {
  if (paragraph.style?.dropCap) {
    return paragraph.style.dropCap;
  }

  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return undefined;
  }

  const cached = paragraphDropCapBySourceXml.get(xml);
  if (cached !== undefined) {
    return cached ?? undefined;
  }

  const framePrTag = xml.match(/<w:framePr\b[^>]*\/?>/i)?.[0];
  if (!framePrTag) {
    setCacheEntry(paragraphDropCapBySourceXml, xml, null);
    return undefined;
  }

  const dropCapRaw = xmlAttribute(framePrTag, "w:dropCap")
    ?.trim()
    .toLowerCase();
  let type: "drop" | "margin" | undefined;
  if (dropCapRaw === "drop" || dropCapRaw === "margin") {
    type = dropCapRaw;
  }
  if (!type) {
    setCacheEntry(paragraphDropCapBySourceXml, xml, null);
    return undefined;
  }

  const linesRaw = xmlAttribute(framePrTag, "w:lines");
  const parsedLines = linesRaw ? Number(linesRaw) : Number.NaN;
  const wrap = xmlAttribute(framePrTag, "w:wrap")?.trim();
  const horizontalAnchor = xmlAttribute(framePrTag, "w:hAnchor")?.trim();
  const verticalAnchor = xmlAttribute(framePrTag, "w:vAnchor")?.trim();
  const xTwipsRaw = xmlAttribute(framePrTag, "w:x");
  const yTwipsRaw = xmlAttribute(framePrTag, "w:y");
  const horizontalSpaceTwipsRaw = xmlAttribute(framePrTag, "w:hSpace");
  const verticalSpaceTwipsRaw = xmlAttribute(framePrTag, "w:vSpace");
  const xTwips = xTwipsRaw ? Number(xTwipsRaw) : Number.NaN;
  const yTwips = yTwipsRaw ? Number(yTwipsRaw) : Number.NaN;
  const horizontalSpaceTwips = horizontalSpaceTwipsRaw
    ? Number(horizontalSpaceTwipsRaw)
    : Number.NaN;
  const verticalSpaceTwips = verticalSpaceTwipsRaw
    ? Number(verticalSpaceTwipsRaw)
    : Number.NaN;
  const resolved = {
    type,
    lines:
      Number.isFinite(parsedLines) && parsedLines > 0
        ? Math.round(parsedLines)
        : undefined,
    ...(wrap ? { wrap } : undefined),
    ...(horizontalAnchor ? { horizontalAnchor } : undefined),
    ...(verticalAnchor ? { verticalAnchor } : undefined),
    ...(Number.isFinite(xTwips) ? { xTwips: Math.round(xTwips) } : undefined),
    ...(Number.isFinite(yTwips) ? { yTwips: Math.round(yTwips) } : undefined),
    ...(Number.isFinite(horizontalSpaceTwips)
      ? { horizontalSpaceTwips: Math.round(horizontalSpaceTwips) }
      : undefined),
    ...(Number.isFinite(verticalSpaceTwips)
      ? { verticalSpaceTwips: Math.round(verticalSpaceTwips) }
      : undefined),
  };
  setCacheEntry(paragraphDropCapBySourceXml, xml, resolved);
  return resolved;
}

export function checkboxChoiceRowTabWidthPx(paragraph: ParagraphNode): number {
  const checkboxCount = paragraph.children.filter(
    (child) => child.type === "form-field" && child.fieldType === "checkbox"
  ).length;
  const labelCharCount = paragraph.children
    .filter((child): child is TextRunNode => child.type === "text")
    .map((child) => child.text.replace(/\t/g, ""))
    .join("")
    .replace(/\s+/g, "").length;

  if (checkboxCount >= 3) {
    return 6;
  }

  if (labelCharCount <= 8) {
    return 8;
  }

  return 10;
}

export function attachTextToPreviousCheckbox(
  paragraph: ParagraphNode,
  childIndex: number,
  text: string
): string {
  if (!text || text[0] !== " ") {
    return text;
  }

  const previousChild =
    childIndex > 0 ? paragraph.children[childIndex - 1] : undefined;
  if (
    previousChild?.type !== "form-field" ||
    previousChild.fieldType !== "checkbox"
  ) {
    return text;
  }

  return `\u00a0${text.slice(1)}`;
}

export function runFontSizePx(
  style?: TextRunNode["style"] | FormFieldRunNode["style"]
): number {
  if (
    style?.fontSizePt &&
    Number.isFinite(style.fontSizePt) &&
    style.fontSizePt > 0
  ) {
    return Math.max(9, (style.fontSizePt * 96) / 72);
  }

  return Math.max(9, (DEFAULT_PARAGRAPH_FONT_SIZE_PT * 96) / 72);
}

export function explicitRunFontSizePx(
  style?: TextRunNode["style"] | FormFieldRunNode["style"]
): number | undefined {
  if (
    style?.fontSizePt &&
    Number.isFinite(style.fontSizePt) &&
    style.fontSizePt > 0
  ) {
    return Math.max(9, (style.fontSizePt * 96) / 72);
  }

  return undefined;
}

export function resolveDropCapFontSizePx(
  style: TextRunNode["style"] | undefined,
  lineHeightPx: number,
  lineCount: number,
  previewFontSizePx?: number
): number {
  if (Number.isFinite(previewFontSizePx) && (previewFontSizePx as number) > 0) {
    return Math.max(12, Math.round(previewFontSizePx as number));
  }

  const explicitFontSizePx = explicitRunFontSizePx(style);
  if (
    Number.isFinite(explicitFontSizePx) &&
    (explicitFontSizePx as number) > 0
  ) {
    return Math.max(12, Math.round(explicitFontSizePx as number));
  }

  const baseFontSizePx = runFontSizePx(style);
  const derivedFromLinesPx = Math.max(
    18,
    Math.round(lineHeightPx * Math.max(1.6, lineCount * 0.92))
  );
  return Math.max(derivedFromLinesPx, Math.round(baseFontSizePx * 1.8));
}

export function resolveDropCapVisualHeightPx(
  style: TextRunNode["style"] | undefined,
  lineHeightPx: number,
  lineCount: number,
  currentFontSizePx?: number
): number {
  const baseHeightPx = Math.max(18, Math.round(lineHeightPx * lineCount));
  const baselineFontSizePx = resolveDropCapFontSizePx(
    style,
    lineHeightPx,
    lineCount
  );
  const effectiveFontSizePx =
    Number.isFinite(currentFontSizePx) && (currentFontSizePx as number) > 0
      ? Math.max(12, Math.round(currentFontSizePx as number))
      : baselineFontSizePx;
  const scale =
    baselineFontSizePx > 0 ? effectiveFontSizePx / baselineFontSizePx : 1;

  return Math.max(
    18,
    Math.round(baseHeightPx * scale),
    Math.round(effectiveFontSizePx * 1.05)
  );
}

export function estimateTextAdvanceWidthPx(
  text: string,
  style?: TextRunNode["style"] | FormFieldRunNode["style"]
): number {
  if (!text) {
    return 0;
  }

  const fontSizePx = runFontSizePx(style);
  const normalized = text.includes("\u00a0")
    ? text.replace(/\u00a0/g, " ")
    : text;
  const cacheKey = `${fontSizePx}\u0000${normalized}`;
  const cached = estimatedTextAdvanceWidthByFontAndValue.get(cacheKey);
  if (cached !== undefined) {
    return cached;
  }

  let total = 0;

  for (let index = 0; index < normalized.length; index += 1) {
    const code = normalized.charCodeAt(index);
    if (code === 10 || code === 13) {
      continue;
    }
    if (code === 9) {
      total += fontSizePx * 2;
      continue;
    }
    if (code === 32) {
      total += fontSizePx * 0.33;
      continue;
    }
    if (
      code === 33 ||
      code === 39 ||
      code === 44 ||
      code === 46 ||
      code === 49 ||
      code === 58 ||
      code === 59 ||
      code === 73 ||
      code === 96 ||
      code === 105 ||
      code === 108 ||
      code === 124
    ) {
      total += fontSizePx * 0.29;
      continue;
    }
    if ((code >= 65 && code <= 90) || (code >= 48 && code <= 57)) {
      total += fontSizePx * 0.6;
      continue;
    }
    if (code >= 0x3000 && code <= 0x9fff) {
      total += fontSizePx * 0.95;
      continue;
    }
    total += fontSizePx * 0.54;
  }

  const estimatedWidthPx = Math.max(0, Math.round(total));
  setCacheEntry(
    estimatedTextAdvanceWidthByFontAndValue,
    cacheKey,
    estimatedWidthPx
  );
  while (
    estimatedTextAdvanceWidthByFontAndValue.size >
    TEXT_MEASURE_CACHE_MAX_ENTRIES
  ) {
    const firstKey = estimatedTextAdvanceWidthByFontAndValue.keys().next()
      .value as string | undefined;
    if (!firstKey) {
      break;
    }
    estimatedTextAdvanceWidthByFontAndValue.delete(firstKey);
  }

  return estimatedWidthPx;
}

export function updateEstimatedLineWidthPxForText(
  currentLineWidthPx: number,
  text: string,
  style?: TextRunNode["style"] | FormFieldRunNode["style"]
): number {
  if (!text) {
    return currentLineWidthPx;
  }

  if (!text.includes("\n")) {
    return currentLineWidthPx + estimateTextAdvanceWidthPx(text, style);
  }

  const segments = text.split("\n");
  const trailingSegment = segments[segments.length - 1] ?? "";
  return estimateTextAdvanceWidthPx(trailingSegment, style);
}

export function resolveTabSpacerWidthPx(
  tabStopPositionsPx: number[],
  currentLineWidthPx: number,
  fallbackWidthPx: number,
  fixedFallback = false
): number {
  const safeFallback = Math.max(12, Math.round(fallbackWidthPx));
  if (tabStopPositionsPx.length === 0) {
    if (fixedFallback) {
      return safeFallback;
    }
    // Word's default tab behavior: advance to the next multiple of the
    // default tab stop from the current position (not a fixed-width gap).
    const nextStop =
      (Math.floor((currentLineWidthPx + 0.5) / safeFallback) + 1) *
      safeFallback;
    return Math.max(2, Math.round(nextStop - currentLineWidthPx));
  }

  const nextStop = tabStopPositionsPx.find(
    (stop) => stop > currentLineWidthPx + 0.5
  );
  if (nextStop !== undefined) {
    return Math.max(8, Math.round(nextStop - currentLineWidthPx));
  }

  const lastStop = tabStopPositionsPx[tabStopPositionsPx.length - 1] ?? 0;
  const overflow = Math.max(0, currentLineWidthPx - lastStop);
  const stepCount = Math.floor(overflow / safeFallback) + 1;
  const projectedStop = lastStop + stepCount * safeFallback;
  return Math.max(8, Math.round(projectedStop - currentLineWidthPx));
}

export function estimateInteractiveFieldWidthPx(
  field: FormFieldRunNode
): number {
  if (field.fieldType === "checkbox") {
    return resolveCheckboxFieldWidthPx(field);
  }

  if (field.fieldType === "date") {
    const text = field.value?.trim() || "MM/DD/YYYY";
    return Math.max(
      68,
      Math.min(220, estimateTextAdvanceWidthPx(text, field.style) + 12)
    );
  }

  if (field.fieldType === "dropdown") {
    const text =
      field.value?.trim() ||
      field.options?.[0]?.displayText?.trim() ||
      field.title?.trim() ||
      "Select";
    return Math.max(
      58,
      Math.min(240, estimateTextAdvanceWidthPx(text, field.style) + 16)
    );
  }

  const defaultText = field.widget?.text?.defaultText?.trim();
  const textValue =
    field.value?.trim() || defaultText || field.title?.trim() || "Click here.";
  return Math.max(
    42,
    Math.min(280, estimateTextAdvanceWidthPx(textValue, field.style) + 12)
  );
}

export function resolveCheckboxFieldWidthPx(field: FormFieldRunNode): number {
  const exactWidgetWidthPx =
    field.widget?.checkbox?.sizeMode === "exact" &&
    Number.isFinite(field.widget.checkbox.sizePt)
      ? Math.max(
          14,
          Math.round(((field.widget.checkbox.sizePt as number) * 96) / 72 + 4)
        )
      : undefined;
  if (Number.isFinite(exactWidgetWidthPx)) {
    return exactWidgetWidthPx as number;
  }

  const checkedSymbol = field.checkedSymbol ?? "☒";
  const uncheckedSymbol = field.uncheckedSymbol ?? "☐";
  return Math.max(
    14,
    estimateTextAdvanceWidthPx(checkedSymbol, field.style),
    estimateTextAdvanceWidthPx(uncheckedSymbol, field.style)
  );
}
