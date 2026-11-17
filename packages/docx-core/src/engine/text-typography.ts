// OOXML character/paragraph typography parsing shared by the TS editor
// pipeline and the post-import model enrichment pass.
//
// The WASM model builder historically collapsed most `w:rPr` vocabulary into a
// few booleans.  These helpers keep the full OOXML vocabulary available to the
// renderer: underline/strike variants, run borders, caps, character spacing and
// scale, RTL, per-script fonts, theme fonts and text effects.

import type {
  DocxThemeFonts,
  ParagraphAlignment,
  ParagraphDirection,
  TextOutlineStyle,
  TextRunBorderStyle,
  TextShadowStyle,
  TextStyle,
  TextStrikeStyle,
  TextUnderlineStyle
} from "./types";

const TAG_PATTERN_CACHE = new Map<string, RegExp>();

function tagPattern(tagName: string): RegExp {
  let pattern = TAG_PATTERN_CACHE.get(tagName);
  if (!pattern) {
    const escaped = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    pattern = new RegExp(`<(/?)${escaped}(?=[\\s>/])[^>]*>`, "gi");
    TAG_PATTERN_CACHE.set(tagName, pattern);
  }
  return pattern;
}

/** Extract the value of an XML attribute from a tag string. */
export function xmlAttribute(
  tagXml: string,
  attribute: string
): string | undefined {
  const escaped = attribute.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = tagXml.match(
    new RegExp(`${escaped}=(?:"([^"]+)"|'([^']+)')`, "i")
  );
  return match?.[1] ?? match?.[2];
}

export function decodeXmlText(value: string): string {
  if (!value) {
    return value;
  }

  return value
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => {
      const codePoint = Number.parseInt(hex, 16);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : _;
    })
    .replace(/&#([0-9]+);/g, (_, decimal: string) => {
      const codePoint = Number.parseInt(decimal, 10);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : _;
    })
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

export interface XmlTagRange {
  start: number;
  end: number;
  tagName: string;
  openTag: string;
}

/** Index balanced (including nested) tag ranges in document order. */
export function extractBalancedTagRanges(
  xml: string,
  tagName: string
): XmlTagRange[] {
  if (!xml) {
    return [];
  }

  const pattern = tagPattern(tagName);
  pattern.lastIndex = 0;
  const stack: Array<{ start: number; openTag: string }> = [];
  const ranges: XmlTagRange[] = [];

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
        openTag: fullMatch
      });
      continue;
    }

    if (!isClosing) {
      stack.push({ start, openTag: fullMatch });
      continue;
    }

    const opener = stack.pop();
    if (opener) {
      ranges.push({
        start: opener.start,
        end: start + fullMatch.length,
        tagName,
        openTag: opener.openTag
      });
    }
  }

  return ranges;
}

export function balancedTagXmlBlocks(xml: string, tagName: string): string[] {
  return extractBalancedTagRanges(xml, tagName).map((range) =>
    xml.slice(range.start, range.end)
  );
}

function firstTagXml(xml: string, tagName: string): string | undefined {
  const range = extractBalancedTagRanges(xml, tagName)[0];
  return range ? xml.slice(range.start, range.end) : undefined;
}

/** OOXML boolean property (`<w:b/>` is true, `<w:b w:val="0"/>` is false). */
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

/** OOXML boolean property that is true when the tag is absent (default on). */
export function xmlBooleanFlagDefaultOn(tagXml: string | undefined): boolean {
  if (!tagXml) {
    return true;
  }

  return !(
    xmlAttribute(tagXml, "w:val")?.trim().toLowerCase() === "false" ||
    xmlAttribute(tagXml, "w:val")?.trim() === "0" ||
    xmlAttribute(tagXml, "w:val")?.trim().toLowerCase() === "off"
  );
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

const UNDERLINE_STYLES = new Set<string>([
  "single",
  "words",
  "double",
  "thick",
  "dotted",
  "dottedheavy",
  "dash",
  "dashedheavy",
  "dashlong",
  "dashlongheavy",
  "dotdash",
  "dashdotheavy",
  "dotdotdash",
  "dashdotdotheavy",
  "wave",
  "wavyheavy",
  "wavydouble"
]);

const UNDERLINE_STYLE_CANONICAL: Record<string, TextUnderlineStyle> = {
  single: "single",
  words: "words",
  double: "double",
  thick: "thick",
  dotted: "dotted",
  dottedheavy: "dottedHeavy",
  dash: "dash",
  dashedheavy: "dashedHeavy",
  dashlong: "dashLong",
  dashlongheavy: "dashLongHeavy",
  dotdash: "dotDash",
  dashdotheavy: "dashDotHeavy",
  dotdotdash: "dotDotDash",
  dashdotdotheavy: "dashDotDotHeavy",
  wave: "wave",
  wavyheavy: "wavyHeavy",
  wavydouble: "wavyDouble"
};

export function normalizeUnderlineStyle(
  value: string | undefined
): TextUnderlineStyle | undefined {
  const normalized = value?.trim().toLowerCase();
  if (!normalized || !UNDERLINE_STYLES.has(normalized)) {
    return undefined;
  }
  return UNDERLINE_STYLE_CANONICAL[normalized];
}

export function parseUnderlineFromRunProperties(rPrXml: string): {
  underline?: boolean;
  underlineStyle?: TextUnderlineStyle;
} {
  const tag = firstTagXml(rPrXml, "w:u");
  if (!tag) {
    return {};
  }

  const raw = xmlAttribute(tag, "w:val")?.trim().toLowerCase();
  if (raw === "none" || raw === "nil" || raw === "0" || raw === "false") {
    return {};
  }

  if (!raw || raw === "single" || raw === "1" || raw === "true") {
    return { underline: true, underlineStyle: "single" };
  }

  const style = normalizeUnderlineStyle(raw);
  return style ? { underline: true, underlineStyle: style } : { underline: true };
}

export function parseStrikeFromRunProperties(rPrXml: string): {
  strike?: boolean;
  strikeStyle?: TextStrikeStyle;
} {
  const doubleStrikeTag = firstTagXml(rPrXml, "w:dstrike");
  if (xmlBooleanFlag(doubleStrikeTag)) {
    return { strike: true, strikeStyle: "double" };
  }

  const strikeTag = firstTagXml(rPrXml, "w:strike");
  if (xmlBooleanFlag(strikeTag)) {
    return { strike: true, strikeStyle: "single" };
  }

  return {};
}

export interface ParsedRunFonts {
  fontFamily?: string;
  fontFamilyEastAsia?: string;
  fontFamilyComplexScript?: string;
}

function themeFontFamily(
  themeValue: string | undefined,
  themeFonts: DocxThemeFonts | undefined
): string | undefined {
  const normalized = themeValue?.trim().toLowerCase();
  if (!normalized || !themeFonts) {
    return undefined;
  }

  const isMajor = normalized.startsWith("major");
  const set = isMajor ? themeFonts.major : themeFonts.minor;
  if (!set) {
    return undefined;
  }

  if (normalized.includes("eastasia")) {
    return set.eastAsia;
  }
  if (normalized.includes("bidi") || normalized.endsWith("cs")) {
    return set.complexScript;
  }
  return set.ascii;
}

export function parseRunFontsFromXml(
  rFontsTag: string | undefined,
  themeFonts?: DocxThemeFonts
): ParsedRunFonts {
  if (!rFontsTag) {
    return {};
  }

  const ascii = xmlAttribute(rFontsTag, "w:ascii")?.trim();
  const hAnsi = xmlAttribute(rFontsTag, "w:hAnsi")?.trim();
  const eastAsia = xmlAttribute(rFontsTag, "w:eastAsia")?.trim();
  const complexScript = xmlAttribute(rFontsTag, "w:cs")?.trim();
  const asciiTheme = xmlAttribute(rFontsTag, "w:asciiTheme");
  const hAnsiTheme = xmlAttribute(rFontsTag, "w:hAnsiTheme");
  const eastAsiaTheme = xmlAttribute(rFontsTag, "w:eastAsiaTheme");
  const complexScriptTheme =
    xmlAttribute(rFontsTag, "w:cstheme") ??
    xmlAttribute(rFontsTag, "w:csTheme");

  return {
    fontFamily:
      ascii ||
      hAnsi ||
      themeFontFamily(hAnsiTheme, themeFonts) ||
      themeFontFamily(asciiTheme, themeFonts) ||
      undefined,
    fontFamilyEastAsia:
      eastAsia || themeFontFamily(eastAsiaTheme, themeFonts) || undefined,
    fontFamilyComplexScript:
      complexScript ||
      themeFontFamily(complexScriptTheme, themeFonts) ||
      undefined
  };
}

export function parseRunBorderFromXml(
  rPrXml: string
): TextRunBorderStyle | undefined {
  const tag = firstTagXml(rPrXml, "w:bdr");
  if (!tag) {
    return undefined;
  }

  const type = xmlAttribute(tag, "w:val")?.trim();
  if (!type) {
    return undefined;
  }

  const sizeRaw = Number(xmlAttribute(tag, "w:sz"));
  const spaceRaw = Number(xmlAttribute(tag, "w:space"));
  const colorRaw = xmlAttribute(tag, "w:color")?.trim();
  const color =
    colorRaw && !/^auto$/i.test(colorRaw)
      ? /^[0-9a-f]{6}$/i.test(colorRaw)
        ? `#${colorRaw.toLowerCase()}`
        : colorRaw.toLowerCase()
      : undefined;
  return {
    type,
    color,
    sizeEighthPt: Number.isFinite(sizeRaw) && sizeRaw > 0 ? sizeRaw : undefined,
    spacePt: Number.isFinite(spaceRaw) && spaceRaw >= 0 ? spaceRaw : undefined,
    frame: /w:frame="(?:1|true|on)"/i.test(tag) ? true : undefined,
    shadow: /w:shadow="(?:1|true|on)"/i.test(tag) ? true : undefined
  };
}

function parseEffectContainer(rPrXml: string): string {
  return firstTagXml(rPrXml, "w:effect") ?? rPrXml;
}

export function parseTextShadowFromRunProperties(
  rPrXml: string
): TextShadowStyle | undefined {
  const effectXml = parseEffectContainer(rPrXml);
  const tag = firstTagXml(effectXml, "w:shadow") ?? firstTagXml(rPrXml, "w:shadow");
  if (!xmlBooleanFlag(tag)) {
    return undefined;
  }

  return {
    blurPt: 1,
    offsetXPt: 1,
    offsetYPt: 1,
    color: "rgba(0, 0, 0, 0.35)"
  };
}

export function parseTextOutlineFromRunProperties(
  rPrXml: string
): TextOutlineStyle | undefined {
  const effectXml = parseEffectContainer(rPrXml);
  const tag =
    firstTagXml(effectXml, "w:outline") ?? firstTagXml(rPrXml, "w:outline");
  if (!xmlBooleanFlag(tag)) {
    return undefined;
  }

  return {
    widthPt: 0.5,
    color: xmlColorValue(tag)
  };
}

export interface ParseRunTypographyOptions {
  themeFonts?: DocxThemeFonts;
}

export function parseRunTypographyFromRunXml(
  runXml: string,
  options: ParseRunTypographyOptions = {}
): TextStyle | undefined {
  const rPrRange = extractBalancedTagRanges(runXml, "w:rPr")[0];
  const rPrXml = rPrRange ? runXml.slice(rPrRange.start, rPrRange.end) : "";
  if (!rPrXml) {
    return undefined;
  }

  const rFontsTag = extractBalancedTagRanges(rPrXml, "w:rFonts")[0];
  const fontFamilies = parseRunFontsFromXml(
    rFontsTag ? rPrXml.slice(rFontsTag.start, rFontsTag.end) : undefined,
    options.themeFonts
  );

  const sizeTag = firstTagXml(rPrXml, "w:sz") ?? firstTagXml(rPrXml, "w:szCs");
  const sizeHalfPoints = sizeTag
    ? Number(xmlAttribute(sizeTag, "w:val"))
    : Number.NaN;
  const fontSizePt =
    Number.isFinite(sizeHalfPoints) && sizeHalfPoints > 0
      ? Number((sizeHalfPoints / 2).toFixed(2))
      : undefined;

  const bold =
    xmlBooleanFlag(firstTagXml(rPrXml, "w:b")) ||
    xmlBooleanFlag(firstTagXml(rPrXml, "w:bCs"));
  const italic =
    xmlBooleanFlag(firstTagXml(rPrXml, "w:i")) ||
    xmlBooleanFlag(firstTagXml(rPrXml, "w:iCs"));
  const underline = parseUnderlineFromRunProperties(rPrXml);
  const strike = parseStrikeFromRunProperties(rPrXml);
  const color = xmlColorValue(firstTagXml(rPrXml, "w:color"));
  const highlightTag = firstTagXml(rPrXml, "w:highlight");
  const highlight = highlightTag
    ? xmlAttribute(highlightTag, "w:val")
    : undefined;
  const verticalAlignTag = firstTagXml(rPrXml, "w:vertAlign");
  const verticalAlign = verticalAlignTag
    ? xmlAttribute(verticalAlignTag, "w:val")?.trim().toLowerCase()
    : undefined;

  const caps = xmlBooleanFlag(firstTagXml(rPrXml, "w:caps"));
  const smallCaps = xmlBooleanFlag(firstTagXml(rPrXml, "w:smallCaps"));
  const spacingTag = firstTagXml(rPrXml, "w:spacing");
  const spacingTwips = spacingTag
    ? Number(xmlAttribute(spacingTag, "w:val"))
    : Number.NaN;
  const widthTag = firstTagXml(rPrXml, "w:w");
  const widthRaw = widthTag ? Number(xmlAttribute(widthTag, "w:val")) : Number.NaN;
  const characterScalePercent =
    Number.isFinite(widthRaw) && widthRaw > 0 && widthRaw !== 100
      ? Math.max(1, Math.min(600, Math.round(widthRaw)))
      : undefined;

  const rtl = xmlBooleanFlag(firstTagXml(rPrXml, "w:rtl"));

  const style: TextStyle = {
    fontFamily: fontFamilies.fontFamily,
    fontFamilyEastAsia: fontFamilies.fontFamilyEastAsia,
    fontFamilyComplexScript: fontFamilies.fontFamilyComplexScript,
    fontSizePt,
    bold: bold || undefined,
    italic: italic || undefined,
    underline: underline.underline,
    underlineStyle: underline.underlineStyle,
    strike: strike.strike,
    strikeStyle: strike.strikeStyle,
    caps: caps || undefined,
    smallCaps: smallCaps || undefined,
    color,
    highlight: highlight?.trim() || undefined,
    characterSpacingTwips:
      Number.isFinite(spacingTwips) && spacingTwips !== 0
        ? Math.round(spacingTwips)
        : undefined,
    characterScalePercent,
    direction: rtl ? "rtl" : undefined,
    verticalAlign:
      verticalAlign === "superscript" || verticalAlign === "subscript"
        ? verticalAlign
        : undefined,
    runBorder: parseRunBorderFromXml(rPrXml),
    textShadow: parseTextShadowFromRunProperties(rPrXml),
    textOutline: parseTextOutlineFromRunProperties(rPrXml)
  };

  return Object.values(style).some((value) => value !== undefined)
    ? style
    : undefined;
}

const PARAGRAPH_ALIGNMENTS = new Set<string>([
  "left",
  "center",
  "right",
  "justify",
  "both",
  "distribute",
  "thaidistribute",
  "low",
  "start",
  "end"
]);

export function parseParagraphAlignmentFromParagraphProperties(
  paragraphPropertiesXml: string
): ParagraphAlignment | undefined {
  const jcTag = firstTagXml(paragraphPropertiesXml, "w:jc");
  const raw = jcTag ? xmlAttribute(jcTag, "w:val")?.trim() : undefined;
  if (!raw) {
    return undefined;
  }

  const normalized = raw.toLowerCase();
  if (normalized === "thaidistribute") {
    return "thaiDistribute";
  }
  return PARAGRAPH_ALIGNMENTS.has(normalized)
    ? (normalized as ParagraphAlignment)
    : undefined;
}

export function parseParagraphDirectionFromParagraphProperties(
  paragraphPropertiesXml: string
): ParagraphDirection | undefined {
  const bidiTag = firstTagXml(paragraphPropertiesXml, "w:bidi");
  if (bidiTag) {
    return xmlBooleanFlag(bidiTag) ? "rtl" : "ltr";
  }

  const rtlTag = firstTagXml(paragraphPropertiesXml, "w:rtl");
  if (rtlTag) {
    return xmlBooleanFlag(rtlTag) ? "rtl" : "ltr";
  }

  return undefined;
}

function themeFontSetFromXml(
  themeXml: string,
  familyTag: string
): { ascii?: string; eastAsia?: string; complexScript?: string } | undefined {
  const block = firstTagXml(themeXml, familyTag);
  if (!block) {
    return undefined;
  }

  const latin = firstTagXml(block, "a:latin");
  const eastAsia = firstTagXml(block, "a:ea");
  const complexScript = firstTagXml(block, "a:cs");
  const result = {
    ascii: latin ? xmlAttribute(latin, "typeface")?.trim() || undefined : undefined,
    eastAsia: eastAsia
      ? xmlAttribute(eastAsia, "typeface")?.trim() || undefined
      : undefined,
    complexScript: complexScript
      ? xmlAttribute(complexScript, "typeface")?.trim() || undefined
      : undefined
  };

  return result.ascii || result.eastAsia || result.complexScript
    ? result
    : undefined;
}

export function parseThemeFontsFromXml(
  themeXml?: string
): DocxThemeFonts | undefined {
  if (!themeXml) {
    return undefined;
  }

  const major = themeFontSetFromXml(themeXml, "a:majorFont");
  const minor = themeFontSetFromXml(themeXml, "a:minorFont");
  return major || minor ? { major, minor } : undefined;
}
