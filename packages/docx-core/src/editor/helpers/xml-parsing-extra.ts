// DOCX XML parsing helpers: balanced-tag extraction, tracked-change token parsing,
// run-style parsing, and image transform helpers.
// Upstream editor.tsx: lines 16061-17156.

import type {
  ImageRunNode,
  ParagraphAlignment,
  TextRunNode,
} from "../../engine/types";
import type { DocxTrackedChangeKind } from "./editor-types";
import { xmlAttribute } from "./ooxml-helpers";
import { twipsToPixels } from "../../viewer/section-layout";

const XML_NAME_ESCAPE_PATTERN = /[.*+?^${}()|[\]\\]/g;

function decodeXmlText(text: string): string {
  if (!text) {
    return text;
  }

  return text
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => {
      const codePoint = Number.parseInt(hex, 16);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : _;
    })
    .replace(/&#([0-9]+);/g, (_, decimal: string) => {
      const codePoint = Number.parseInt(decimal, 10);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : _;
    })
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

export function extractBalancedTagRanges(
  xml: string,
  tagName: string
): Array<{ start: number; end: number; tagName: string; openTag: string }> {
  if (!xml) {
    return [];
  }

  const escapedTagName = tagName.replace(XML_NAME_ESCAPE_PATTERN, "\\$&");
  const pattern = new RegExp(`<(/?)${escapedTagName}(?=[\\s>/])[^>]*>`, "gi");
  const stack: Array<{ start: number; openTag: string }> = [];
  const ranges: Array<{ start: number; end: number; tagName: string; openTag: string }> = [];

  for (const match of xml.matchAll(pattern)) {
    const fullMatch = match[0] ?? "";
    if (!fullMatch) {
      continue;
    }

    const start = match.index ?? 0;
    const isClosing = match[1] === "/";
    const isSelfClosing = !isClosing && /\/>\s*$/i.test(fullMatch);
    if (isSelfClosing) {
      ranges.push({ start, end: start + fullMatch.length, tagName, openTag: fullMatch });
      continue;
    }

    if (!isClosing) {
      stack.push({ start, openTag: fullMatch });
      continue;
    }

    const opener = stack.pop();
    if (!opener) {
      continue;
    }

    ranges.push({ start: opener.start, end: start + fullMatch.length, tagName, openTag: opener.openTag });
  }

  return ranges;
}

export function trackedChangeKindFromTagName(
  tagName: string
): Exclude<DocxTrackedChangeKind, "format-change" | "paragraph-format-change"> | undefined {
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

  const normalized = value.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
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
  const pattern = /<w:delText\b[^>]*>([\s\S]*?)<\/w:delText>|<(?:w|a):t\b[^>]*>([\s\S]*?)<\/(?:w|a):t>|<w:tab\b[^>]*\/?>|<w:(?:br|cr)\b[^>]*\/?>|<w:footnoteReference\b[^>]*\/?>|<w:endnoteReference\b[^>]*\/?>/gi;

  for (const match of runXml.matchAll(pattern)) {
    if (match[1] !== undefined) {
      if (!includeDeletedText) {
        continue;
      }
      tokens.push({ text: decodeXmlText(match[1] ?? ""), isNote: false });
      continue;
    }

    if (match[2] !== undefined) {
      tokens.push({ text: decodeXmlText(match[2] ?? ""), isNote: false });
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
  runXml: string
): TextRunNode["style"] | undefined {
  const rPrRange = extractBalancedTagRanges(runXml, "w:rPr")[0];
  const rPrXml = rPrRange ? runXml.slice(rPrRange.start, rPrRange.end) : "";
  if (!rPrXml) {
    return undefined;
  }

  const rFontsTag = rPrXml.match(/<w:rFonts\b[^>]*\/?>/i)?.[0];
  const rFontsAscii = rFontsTag ? xmlAttribute(rFontsTag, "w:ascii") : undefined;
  const rFontsHAnsi = rFontsTag ? xmlAttribute(rFontsTag, "w:hAnsi") : undefined;
  const rFontsEastAsia = rFontsTag ? xmlAttribute(rFontsTag, "w:eastAsia") : undefined;
  const rFontsCs = rFontsTag ? xmlAttribute(rFontsTag, "w:cs") : undefined;
  const fontFamily = rFontsAscii ?? rFontsHAnsi ?? rFontsEastAsia ?? rFontsCs;

  const sizeTag =
    rPrXml.match(/<w:sz\b[^>]*\/?>/i)?.[0] ?? rPrXml.match(/<w:szCs\b[^>]*\/?>/i)?.[0];
  const sizeHalfPoints = sizeTag ? Number(xmlAttribute(sizeTag, "w:val")) : Number.NaN;
  const fontSizePt =
    Number.isFinite(sizeHalfPoints) && sizeHalfPoints > 0
      ? Number((sizeHalfPoints / 2).toFixed(2))
      : undefined;

  const bold = xmlBooleanFlag(rPrXml.match(/<w:b(?:Cs)?\b[^>]*\/?>/i)?.[0]);
  const italic = xmlBooleanFlag(rPrXml.match(/<w:i(?:Cs)?\b[^>]*\/?>/i)?.[0]);
  const underlineTag = rPrXml.match(/<w:u\b[^>]*\/?>/i)?.[0];
  const underline = xmlBooleanFlag(underlineTag);
  const strike = xmlBooleanFlag(rPrXml.match(/<w:strike\b[^>]*\/?>/i)?.[0]);
  const color = xmlColorValue(rPrXml.match(/<w:color\b[^>]*\/?>/i)?.[0]);
  const highlightTag = rPrXml.match(/<w:highlight\b[^>]*\/?>/i)?.[0];
  const highlight = highlightTag ? xmlAttribute(highlightTag, "w:val") : undefined;
  const verticalAlignTag = rPrXml.match(/<w:vertAlign\b[^>]*\/?>/i)?.[0];
  const verticalAlign = verticalAlignTag ? xmlAttribute(verticalAlignTag, "w:val") : undefined;

  const style: NonNullable<TextRunNode["style"]> = {
    fontFamily: fontFamily?.trim() || undefined,
    fontSizePt,
    bold: bold || undefined,
    italic: italic || undefined,
    underline: underline || undefined,
    strike: strike || undefined,
    color,
    highlight: highlight?.trim() || undefined,
    verticalAlign:
      verticalAlign === "superscript" || verticalAlign === "subscript"
        ? verticalAlign
        : undefined,
  };

  return Object.values(style).some((value) => value !== undefined) ? style : undefined;
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
    fontSizePt: override?.fontSizePt ?? base?.fontSizePt,
    bold: override?.bold ?? base?.bold,
    italic: override?.italic ?? base?.italic,
    underline: override?.underline ?? base?.underline,
    strike: override?.strike ?? base?.strike,
    color: override?.color ?? base?.color,
    highlight: override?.highlight ?? base?.highlight,
    verticalAlign: override?.verticalAlign ?? base?.verticalAlign,
  };

  return Object.values(merged).some((value) => value !== undefined) ? merged : undefined;
}

export function parseParagraphAlignmentFromXml(
  paragraphPropertiesXml: string
): ParagraphAlignment | undefined {
  const jcTag = paragraphPropertiesXml.match(/<w:jc\b[^>]*\/?>/i)?.[0];
  const raw = jcTag ? xmlAttribute(jcTag, "w:val")?.trim().toLowerCase() : undefined;
  if (
    raw === "center" ||
    raw === "right" ||
    raw === "justify" ||
    raw === "left"
  ) {
    return raw;
  }
  return undefined;
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

  return { rotationDegrees, flipH: flipH || undefined, flipV: flipV || undefined };
}

export function joinCssTransforms(...parts: Array<string | undefined>): string | undefined {
  const resolved = parts.filter((part): part is string => Boolean(part && part.trim().length > 0));
  return resolved.length > 0 ? resolved.join(" ") : undefined;
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
  const normalizedQuarterTurn =
    Number.isFinite(rotationDegrees) &&
    Math.abs((Math.abs(rotationDegrees as number) % 180) - 90) < 0.5
      ? (((Math.round((rotationDegrees as number) / 90) % 4) + 4) % 4) * 90
      : undefined;
  const flipScaleTransform =
    transform?.flipH || transform?.flipV
      ? `scale(${transform?.flipH ? -1 : 1}, ${transform?.flipV ? -1 : 1})`
      : undefined;
  void twipsToPixels(0);

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
        : normalizedQuarterTurn === 180
          ? `translate(${safeFrameWidthPx}px, ${safeFrameHeightPx}px) rotate(180deg)`
          : normalizedQuarterTurn === 270
            ? `translate(0px, ${safeFrameHeightPx}px) rotate(-90deg)`
            : undefined;

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
    Number.isFinite(rotationDegrees) && Math.abs(rotationDegrees as number) >= 0.01
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
