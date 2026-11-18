// DOCX XML parsing helpers: run-style parsing, tracked-change token parsing,
// balanced-tag extraction, and drawing/image transform parsing.
// Upstream editor.tsx: lines 16061-16261, 16263-16413, 17032-17156.

import type {
  DocxThemeFonts,
  ImageRunNode,
  ParagraphAlignment,
  ParagraphDirection,
  TextRunNode
} from "../../engine/types";
import {
  parseParagraphAlignmentFromParagraphProperties,
  parseParagraphDirectionFromParagraphProperties,
  parseRunTypographyFromRunXml
} from "../../engine/text-typography";
import { xmlAttribute } from "./ooxml-helpers";
import type { DocxInlineTrackedChangeKind } from "./editor-types";

export function decodeXmlText(text: string): string {
  if (!text) {
    return text;
  }

  const withNumericEntities = text
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => {
      const codePoint = Number.parseInt(hex, 16);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : _;
    })
    .replace(/&#([0-9]+);/g, (_, decimal: string) => {
      const codePoint = Number.parseInt(decimal, 10);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : _;
    });

  return withNumericEntities
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

export interface XmlBalancedTagRange {
  start: number;
  end: number;
  tagName: string;
  openTag: string;
}

export function extractBalancedTagRanges(
  xml: string,
  tagName: string
): XmlBalancedTagRange[] {
  if (!xml) {
    return [];
  }

  const escapedTagName = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`<(/?)${escapedTagName}(?=[\\s>/])[^>]*>`, "gi");
  const stack: Array<{ start: number; openTag: string }> = [];
  const ranges: XmlBalancedTagRange[] = [];

  for (const match of xml.matchAll(pattern)) {
    const fullMatch = match[0] ?? "";
    if (!fullMatch) {
      continue;
    }

    const start = match.index ?? 0;
    const isClosing = match[1] === "/";
    const isSelfClosing = !isClosing && /\/>\s*$/i.test(fullMatch);
    if (isSelfClosing) {
      ranges.push({
        start,
        end: start + fullMatch.length,
        tagName,
        openTag: fullMatch,
      });
      continue;
    }

    if (!isClosing) {
      stack.push({
        start,
        openTag: fullMatch,
      });
      continue;
    }

    const opener = stack.pop();
    if (!opener) {
      continue;
    }

    ranges.push({
      start: opener.start,
      end: start + fullMatch.length,
      tagName,
      openTag: opener.openTag,
    });
  }

  return ranges;
}

/**
 * Returns only balanced ranges that are not nested inside another range of the
 * same tag. Tables, rows and paragraphs can contain nested structures of the
 * same name; callers that index top-level containers (rendered rows, cells or
 * paragraph properties) need the outer ranges only.
 */
export function extractTopLevelTagRanges(
  xml: string,
  tagName: string
): XmlBalancedTagRange[] {
  const ranges = extractBalancedTagRanges(xml, tagName);
  return ranges.filter(
    (range) =>
      !ranges.some(
        (other) =>
          other !== range &&
          other.start <= range.start &&
          other.end >= range.end &&
          (other.start < range.start || other.end > range.end)
      )
  );
}

export interface RevisionTagRange extends XmlBalancedTagRange {
  kind: DocxInlineTrackedChangeKind;
  revisionId?: string;
  author?: string;
  date?: string;
}

export function trackedChangeKindFromTagName(
  tagName: string
): DocxInlineTrackedChangeKind | undefined {
  const normalized = tagName.trim().toLowerCase();
  switch (normalized) {
    case "w:ins":
      return "insertion";
    case "w:del":
      return "deletion";
    case "w:movefrom":
      return "move-from";
    case "w:moveto":
      return "move-to";
    default:
      return undefined;
  }
}

export function normalizeTrackedChangeSnippet(value?: string): string | undefined {
  if (!value) {
    return undefined;
  }
  const normalized = value
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return normalized.length > 0 ? normalized : undefined;
}

export function formatTrackedChangeDate(value?: string): string | undefined {
  if (!value) {
    return undefined;
  }

  const trimmed = value.trim();
  if (!trimmed) {
    return undefined;
  }

  const parsed = new Date(trimmed);
  if (Number.isNaN(parsed.getTime())) {
    return undefined;
  }

  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "2-digit",
    year: "numeric",
    timeZone: "UTC",
  }).format(parsed);
}

export function stripTextBoxContentFromRunXml(runXml: string): string {
  if (!runXml.includes("w:txbxContent")) {
    return runXml;
  }
  return runXml.replace(/<w:txbxContent\b[\s\S]*?<\/w:txbxContent>/gi, "");
}

export function parseTrackedRunTokens(
  runXml: string,
  includeDeletedText: boolean
): Array<{ text: string; isNote: boolean }> {
  if (!runXml) {
    return [];
  }

  const tokens: Array<{ text: string; isNote: boolean }> = [];
  const pattern =
    /<w:delText\b[^>]*>([\s\S]*?)<\/w:delText>|<(?:w|a):t\b[^>]*>([\s\S]*?)<\/(?:w|a):t>|<w:tab\b[^>]*\/?>|<w:(?:br|cr)\b[^>]*\/?>|<w:footnoteReference\b[^>]*\/?>|<w:endnoteReference\b[^>]*\/?>/gi;

  for (const match of runXml.matchAll(pattern)) {
    if (match[1] !== undefined) {
      if (!includeDeletedText) {
        continue;
      }
      const decoded = decodeXmlText(match[1] ?? "");
      tokens.push({ text: decoded, isNote: false });
      continue;
    }

    if (match[2] !== undefined) {
      const decoded = decodeXmlText(match[2] ?? "");
      tokens.push({ text: decoded, isNote: false });
      continue;
    }

    const tagXml = match[0] ?? "";
    if (/^<w:tab\b/i.test(tagXml)) {
      tokens.push({ text: "\t", isNote: false });
      continue;
    }
    if (/^<w:(?:br|cr)\b/i.test(tagXml)) {
      tokens.push({ text: "\n", isNote: false });
      continue;
    }
    if (/^<w:(?:footnoteReference|endnoteReference)\b/i.test(tagXml)) {
      tokens.push({ text: "\u2063", isNote: true });
    }
  }

  return tokens;
}
export function xmlBooleanFlag(tagXml: string | undefined): boolean {
  if (!tagXml) {
    return false;
  }

  const raw = xmlAttribute(tagXml, "w:val")?.trim().toLowerCase();
  if (!raw) {
    return true;
  }

  return !(raw === "false" || raw === "0" || raw === "off" || raw === "none");
}

export function xmlColorValue(tagXml: string | undefined): string | undefined {
  if (!tagXml) {
    return undefined;
  }

  const raw = xmlAttribute(tagXml, "w:val")?.trim();
  if (!raw || /^auto$/i.test(raw)) {
    return undefined;
  }

  if (/^#[0-9a-f]{6}$/i.test(raw)) {
    return raw.toLowerCase();
  }

  if (/^[0-9a-f]{6}$/i.test(raw)) {
    return `#${raw.toLowerCase()}`;
  }

  return raw;
}

export function parseRunStyleFromRunXml(
  runXml: string,
  options: { themeFonts?: DocxThemeFonts } = {}
): TextRunNode["style"] | undefined {
  return parseRunTypographyFromRunXml(runXml, options);
}

export function balancedTagXmlBlocks(xml: string, tagName: string): string[] {
  return extractBalancedTagRanges(xml, tagName).map((range) =>
    xml.slice(range.start, range.end)
  );
}

export function mergeTextRunStyles(
  base?: TextRunNode["style"],
  override?: TextRunNode["style"]
): TextRunNode["style"] | undefined {
  if (!base && !override) {
    return undefined;
  }

  const merged: TextRunNode["style"] = {
    fontFamily: override?.fontFamily ?? base?.fontFamily,
    fontFamilyEastAsia:
      override?.fontFamilyEastAsia ?? base?.fontFamilyEastAsia,
    fontFamilyComplexScript:
      override?.fontFamilyComplexScript ?? base?.fontFamilyComplexScript,
    fontSizePt: override?.fontSizePt ?? base?.fontSizePt,
    bold: override?.bold ?? base?.bold,
    italic: override?.italic ?? base?.italic,
    underline: override?.underline ?? base?.underline,
    underlineStyle: override?.underlineStyle ?? base?.underlineStyle,
    strike: override?.strike ?? base?.strike,
    strikeStyle: override?.strikeStyle ?? base?.strikeStyle,
    caps: override?.caps ?? base?.caps,
    smallCaps: override?.smallCaps ?? base?.smallCaps,
    color: override?.color ?? base?.color,
    highlight: override?.highlight ?? base?.highlight,
    characterSpacingTwips:
      override?.characterSpacingTwips ?? base?.characterSpacingTwips,
    characterScalePercent:
      override?.characterScalePercent ?? base?.characterScalePercent,
    direction: override?.direction ?? base?.direction,
    verticalAlign: override?.verticalAlign ?? base?.verticalAlign,
    runBorder: override?.runBorder ?? base?.runBorder,
    textShadow: override?.textShadow ?? base?.textShadow,
    textOutline: override?.textOutline ?? base?.textOutline,
  };

  return Object.values(merged).some((value) => value !== undefined)
    ? merged
    : undefined;
}

export function parseParagraphAlignmentFromXml(
  paragraphPropertiesXml: string
): ParagraphAlignment | undefined {
  return parseParagraphAlignmentFromParagraphProperties(paragraphPropertiesXml);
}

export function parseParagraphDirectionFromXml(
  paragraphPropertiesXml: string
): ParagraphDirection | undefined {
  return parseParagraphDirectionFromParagraphProperties(paragraphPropertiesXml);
}

export function parseDrawingImageTransformFromSourceXml(sourceXml?: string):
  | {
      rotationDegrees?: number;
      flipH?: boolean;
      flipV?: boolean;
    }
  | undefined {
  if (!sourceXml) {
    return undefined;
  }

  const transformXml =
    balancedTagXmlBlocks(sourceXml, "a:xfrm")[0] ??
    sourceXml.match(/<a:xfrm\b[^>]*\/?>/i)?.[0] ??
    "";
  if (!transformXml) {
    return undefined;
  }

  const rotationRaw = Number(xmlAttribute(transformXml, "rot"));
  const rotationDegrees = Number.isFinite(rotationRaw)
    ? Number(((rotationRaw as number) / 60000).toFixed(3))
    : undefined;
  const flipHRaw = xmlAttribute(transformXml, "flipH")?.trim().toLowerCase();
  const flipVRaw = xmlAttribute(transformXml, "flipV")?.trim().toLowerCase();
  const flipH = flipHRaw === "1" || flipHRaw === "true";
  const flipV = flipVRaw === "1" || flipVRaw === "true";

  if (!Number.isFinite(rotationDegrees) && !flipH && !flipV) {
    return undefined;
  }

  return {
    rotationDegrees,
    flipH: flipH || undefined,
    flipV: flipV || undefined,
  };
}

export function joinCssTransforms(
  ...parts: Array<string | undefined>
): string | undefined {
  const resolved = parts.filter((part): part is string =>
    Boolean(part && part.trim().length > 0)
  );
  return resolved.length > 0 ? resolved.join(" ") : undefined;
}

/**
 * Normalized quarter-turn rotation of a drawing (90 or 270), or `undefined`
 * when the image is not rotated by a quarter turn.
 */
export function resolveImageRotationQuarterTurnDegrees(
  rotationDegrees?: number
): 90 | 270 | undefined {
  if (!Number.isFinite(rotationDegrees)) {
    return undefined;
  }

  const degrees = rotationDegrees as number;
  if (Math.abs((Math.abs(degrees) % 180) - 90) >= 0.5) {
    return undefined;
  }

  const normalized = (((Math.round(degrees / 90) % 4) + 4) % 4) * 90;
  return normalized === 90 || normalized === 270 ? normalized : undefined;
}

/**
 * Resolve the layout frame an image occupies after its drawing rotation.
 *
 * A 90°/270° rotation turns the visual footprint a quarter turn, so the frame
 * swaps width and height; rendering the image inside the unrotated frame makes
 * it overflow and paint over neighbouring content. Quarter-turned frames are
 * reported with `rotationSwapped` so callers can rotate the image into the
 * frame (`resolveImageRenderTransformStyle` with `fillFrame: true`).
 */
export function resolveImageFrameSizePx(
  image: ImageRunNode,
  widthPx?: number,
  heightPx?: number
): { widthPx?: number; heightPx?: number; rotationSwapped: boolean } {
  const quarterTurn = resolveImageRotationQuarterTurnDegrees(
    parseDrawingImageTransformFromSourceXml(image.sourceXml)?.rotationDegrees
  );
  const safeWidthPx = Number.isFinite(widthPx)
    ? Math.max(1, Math.round(widthPx as number))
    : undefined;
  const safeHeightPx = Number.isFinite(heightPx)
    ? Math.max(1, Math.round(heightPx as number))
    : undefined;
  if (
    quarterTurn === undefined ||
    safeWidthPx === undefined ||
    safeHeightPx === undefined
  ) {
    return { widthPx, heightPx, rotationSwapped: false };
  }

  return {
    widthPx: safeHeightPx,
    heightPx: safeWidthPx,
    rotationSwapped: true,
  };
}

export function resolveImageRenderTransformStyle(
  image: ImageRunNode,
  options?: {
    frameWidthPx?: number;
    frameHeightPx?: number;
    fillFrame?: boolean;
    baseTransform?: string;
  }
): Record<string, string | number | undefined> {
  const transform = parseDrawingImageTransformFromSourceXml(image.sourceXml);
  const rotationDegrees = transform?.rotationDegrees;
  const frameWidthPx = Number.isFinite(options?.frameWidthPx)
    ? Math.max(1, Math.round(options?.frameWidthPx as number))
    : undefined;
  const frameHeightPx = Number.isFinite(options?.frameHeightPx)
    ? Math.max(1, Math.round(options?.frameHeightPx as number))
    : undefined;
  const fillFrame = options?.fillFrame === true;
  const normalizedQuarterTurn = resolveImageRotationQuarterTurnDegrees(
    rotationDegrees
  );
  const flipScaleTransform =
    transform?.flipH || transform?.flipV
      ? `scale(${transform?.flipH ? -1 : 1}, ${transform?.flipV ? -1 : 1})`
      : undefined;

  if (
    fillFrame &&
    Number.isFinite(frameWidthPx) &&
    Number.isFinite(frameHeightPx) &&
    Number.isFinite(normalizedQuarterTurn)
  ) {
    const safeFrameWidthPx = frameWidthPx as number;
    const safeFrameHeightPx = frameHeightPx as number;
    const rotationTransform =
      normalizedQuarterTurn === 90
        ? `translate(${safeFrameWidthPx}px, 0px) rotate(90deg)`
        : `translate(0px, ${safeFrameHeightPx}px) rotate(-90deg)`;

    return {
      width: `${safeFrameHeightPx}px`,
      height: `${safeFrameWidthPx}px`,
      maxWidth: "none",
      transformOrigin: "top left",
      transform: joinCssTransforms(
        options?.baseTransform,
        rotationTransform,
        flipScaleTransform
      ),
    };
  }

  const rotationTransform =
    Number.isFinite(rotationDegrees) &&
    Math.abs(rotationDegrees as number) >= 0.01
      ? `rotate(${rotationDegrees}deg)`
      : undefined;
  const transformValue = joinCssTransforms(
    options?.baseTransform,
    rotationTransform,
    flipScaleTransform
  );
  if (!transformValue) {
    return {};
  }

  return {
    transformOrigin: "center center",
    transform: transformValue,
  };
}
