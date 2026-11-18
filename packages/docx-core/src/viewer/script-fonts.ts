/**
 * OOXML per-script font selection.
 *
 * A run carries up to four font slots (`w:ascii`, `w:hAnsi`, `w:eastAsia`,
 * `w:cs`). Word picks the slot per character and per run language; a single
 * DOM span with one `font-family` therefore cannot reproduce a mixed
 * Chinese/English/Arabic paragraph. This module classifies each code point,
 * keeps neutral punctuation with the surrounding strong script, and splits the
 * run into segments whose UTF-16 offsets stay aligned with selection and
 * pretext layout.
 */

import type { TextStyle } from "../engine/types";

export type DocxFontScript =
  | "ascii"
  | "highAnsi"
  | "eastAsia"
  | "complexScript";

export interface DocxScriptFontSegment {
  text: string;
  startOffset: number;
  endOffset: number;
  script: DocxFontScript;
  fontFamily?: string;
  fontSizePt?: number;
}

const EAST_ASIA_SCRIPT_RE =
  /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}\p{Script=Yi}]/u;
const COMPLEX_SCRIPT_RE =
  /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gurmukhi}\p{Script=Gujarati}\p{Script=Oriya}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Sinhala}\p{Script=Thai}\p{Script=Lao}\p{Script=Tibetan}\p{Script=Myanmar}\p{Script=Khmer}]/u;
const NEUTRAL_OR_CONTINUATION_RE =
  /[\p{Mark}\p{Separator}\p{Punctuation}\u200c\u200d\ufe00-\ufe0f]/u;
const EAST_ASIA_LANGUAGE_RE = /^(?:ja|ko|zh)(?:-|$)/i;
const COMPLEX_SCRIPT_LANGUAGE_RE =
  /^(?:ar|ckb|dv|fa|he|hi|ks|ku|pa-arab|ps|sd|syr|ug|ur|yi)(?:-|$)/i;
const GENERIC_FAMILIES =
  /^(?:serif|sans-serif|monospace|cursive|fantasy|system-ui)$/i;

function normalizeFamily(value?: string): string | undefined {
  const normalized = value?.trim();
  return normalized || undefined;
}

function normalizedFamilyKey(value?: string): string | undefined {
  return normalizeFamily(value)
    ?.replace(/^['"]+|['"]+$/g, "")
    .toLowerCase();
}

function hasLegacyFontFamilyOverride(style?: TextStyle): boolean {
  const current = normalizedFamilyKey(style?.fontFamily);
  const imported = normalizedFamilyKey(style?.sourceFontFamily);
  if (!current) {
    return false;
  }
  if (imported) {
    return current !== imported;
  }

  return !(
    normalizeFamily(style?.fontFamilyAscii) ||
    normalizeFamily(style?.fontFamilyHAnsi) ||
    normalizeFamily(style?.fontFamilyEastAsia) ||
    normalizeFamily(style?.fontFamilyComplexScript)
  );
}

function normalizedFontHint(style?: TextStyle): string {
  return (
    style?.fontHint
      ?.trim()
      .replace(/[-_\s]/g, "")
      .toLowerCase() ?? ""
  );
}

function isRightToLeft(style?: TextStyle): boolean {
  return style?.rightToLeft === true || style?.direction === "rtl";
}

function neutralScript(style?: TextStyle): DocxFontScript {
  const hint = normalizedFontHint(style);
  if (hint === "eastasia") {
    return "eastAsia";
  }
  if (hint === "cs" || hint === "complexscript" || hint === "bidi") {
    return "complexScript";
  }
  const language = style?.language?.trim();
  if (
    isRightToLeft(style) ||
    COMPLEX_SCRIPT_LANGUAGE_RE.test(language ?? "") ||
    (style?.rightToLeft !== false &&
      style?.direction !== "ltr" &&
      !language &&
      COMPLEX_SCRIPT_LANGUAGE_RE.test(style?.languageBidi?.trim() ?? ""))
  ) {
    return "complexScript";
  }
  if (EAST_ASIA_LANGUAGE_RE.test(style?.languageEastAsia?.trim() ?? "")) {
    return "eastAsia";
  }
  return "highAnsi";
}

/**
 * Selects the OOXML font slot for one Unicode code point. An explicit
 * complex-script property forces complex-script formatting, while an
 * explicitly disabled property lets mixed RTL runs keep per-script fonts.
 */
export function classifyDocxFontScript(
  value: string,
  style?: TextStyle
): DocxFontScript | undefined {
  const character = Array.from(value)[0];
  if (!character) {
    return undefined;
  }

  if (
    style?.complexScript === true ||
    (isRightToLeft(style) && style?.complexScript !== false)
  ) {
    return "complexScript";
  }

  const codePoint = character.codePointAt(0) ?? 0;
  if (codePoint <= 0x7f) {
    return "ascii";
  }
  if (
    EAST_ASIA_SCRIPT_RE.test(character) ||
    (codePoint >= 0x3000 && codePoint <= 0x303f) ||
    (codePoint >= 0xff00 && codePoint <= 0xffef)
  ) {
    return "eastAsia";
  }
  if (COMPLEX_SCRIPT_RE.test(character)) {
    return "complexScript";
  }
  if (NEUTRAL_OR_CONTINUATION_RE.test(character)) {
    return undefined;
  }
  return "highAnsi";
}

/** Resolve one OOXML slot, retaining the legacy single-family API as fallback. */
export function resolveDocxScriptFontFamily(
  style: TextStyle | undefined,
  script: DocxFontScript
): string | undefined {
  const legacyFamily = normalizeFamily(style?.fontFamily);
  if (hasLegacyFontFamilyOverride(style)) {
    return legacyFamily;
  }

  switch (script) {
    case "ascii":
      return (
        normalizeFamily(style?.fontFamilyAscii) ??
        normalizeFamily(style?.fontFamilyHAnsi) ??
        normalizeFamily(style?.fontFamilyEastAsia) ??
        normalizeFamily(style?.fontFamilyComplexScript) ??
        legacyFamily
      );
    case "highAnsi":
      return (
        normalizeFamily(style?.fontFamilyHAnsi) ??
        normalizeFamily(style?.fontFamilyAscii) ??
        normalizeFamily(style?.fontFamilyEastAsia) ??
        normalizeFamily(style?.fontFamilyComplexScript) ??
        legacyFamily
      );
    case "eastAsia":
      return (
        normalizeFamily(style?.fontFamilyEastAsia) ??
        normalizeFamily(style?.fontFamilyHAnsi) ??
        normalizeFamily(style?.fontFamilyAscii) ??
        normalizeFamily(style?.fontFamilyComplexScript) ??
        legacyFamily
      );
    case "complexScript":
      return (
        normalizeFamily(style?.fontFamilyComplexScript) ??
        normalizeFamily(style?.fontFamilyHAnsi) ??
        normalizeFamily(style?.fontFamilyAscii) ??
        normalizeFamily(style?.fontFamilyEastAsia) ??
        legacyFamily
      );
  }
}

function resolveDocxScriptFontSizePt(
  style: TextStyle | undefined,
  script: DocxFontScript
): number | undefined {
  return script === "complexScript"
    ? style?.fontSizeCsPt ?? style?.fontSizePt
    : style?.fontSizePt;
}

interface ScriptToken {
  text: string;
  startOffset: number;
  endOffset: number;
  script?: DocxFontScript;
}

/**
 * Splits where the resolved font family or size changes. UTF-16 offsets remain
 * aligned with DOM selection and pretext layout offsets.
 */
export function segmentTextByDocxScriptFont(
  text: string,
  style?: TextStyle
): DocxScriptFontSegment[] {
  if (!text) {
    return [];
  }

  if (/^[\x00-\x7f]+$/.test(text)) {
    const script = classifyDocxFontScript(text[0], style) ?? "ascii";
    return [
      {
        text,
        startOffset: 0,
        endOffset: text.length,
        script,
        fontFamily: resolveDocxScriptFontFamily(style, script),
        ...(style?.fontSizeCsPt !== undefined
          ? { fontSizePt: resolveDocxScriptFontSizePt(style, script) }
          : {}),
      },
    ];
  }

  const tokens: ScriptToken[] = [];
  let offset = 0;
  for (const character of text) {
    const startOffset = offset;
    offset += character.length;
    tokens.push({
      text: character,
      startOffset,
      endOffset: offset,
      script: classifyDocxFontScript(character, style),
    });
  }

  const fallbackScript = neutralScript(style);
  const nextStrongScript: Array<DocxFontScript | undefined> = new Array(
    tokens.length
  );
  let nextScript: DocxFontScript | undefined;
  for (let index = tokens.length - 1; index >= 0; index -= 1) {
    nextStrongScript[index] = nextScript;
    if (tokens[index]?.script) {
      nextScript = tokens[index]?.script;
    }
  }

  let previousScript: DocxFontScript | undefined;
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (!token) {
      continue;
    }
    token.script =
      token.script ??
      previousScript ??
      nextStrongScript[index] ??
      fallbackScript;
    previousScript = token.script;
  }

  const segments: DocxScriptFontSegment[] = [];
  const families = new Map<DocxFontScript, { family?: string; key?: string }>();
  let previousFamilyKey: string | undefined;
  for (const token of tokens) {
    const script = token.script ?? fallbackScript;
    let resolved = families.get(script);
    if (!resolved) {
      const family = resolveDocxScriptFontFamily(style, script);
      resolved = { family, key: normalizedFamilyKey(family) };
      families.set(script, resolved);
    }
    const fontFamily = resolved.family;
    const fontSizePt =
      style?.fontSizeCsPt !== undefined
        ? resolveDocxScriptFontSizePt(style, script)
        : undefined;
    const previous = segments[segments.length - 1];
    if (
      previous &&
      previousFamilyKey === resolved.key &&
      previous.fontSizePt === fontSizePt
    ) {
      previous.text += token.text;
      previous.endOffset = token.endOffset;
      continue;
    }
    previousFamilyKey = resolved.key;
    segments.push({
      text: token.text,
      startOffset: token.startOffset,
      endOffset: token.endOffset,
      script,
      fontFamily,
      ...(fontSizePt !== undefined ? { fontSizePt } : {}),
    });
  }

  return segments;
}

/**
 * Font size Word uses for `text`: the complex-script size for a complex-script
 * run, otherwise the run size.
 */
export function resolveDocxTextFontSizePt(
  text: string,
  style?: TextStyle
): number | undefined {
  for (const character of text) {
    const script = classifyDocxFontScript(character, style);
    if (script) return resolveDocxScriptFontSizePt(style, script);
  }
  return resolveDocxScriptFontSizePt(
    style,
    style?.complexScript === true ||
      (isRightToLeft(style) && style?.complexScript !== false)
      ? "complexScript"
      : "highAnsi"
  );
}

/** Font family Word uses for `text`, based on its first strong-script run. */
export function resolveDocxTextFontFamily(
  text: string,
  style?: TextStyle
): string | undefined {
  let hasStrongScript = false;
  for (const character of text) {
    const script = classifyDocxFontScript(character, style);
    if (!script) continue;
    hasStrongScript = true;
    const family = resolveDocxScriptFontFamily(style, script);
    if (family) return family;
  }
  return (
    (text && !hasStrongScript
      ? resolveDocxScriptFontFamily(style, neutralScript(style))
      : undefined) ?? normalizeFamily(style?.fontFamily)
  );
}

/**
 * Style for one script segment: the segment's family/size replaces the run's
 * legacy family so `runStyleToCss()` emits the same font the layout measured.
 */
export function docxScriptSegmentStyle(
  style: TextStyle | undefined,
  segment: DocxScriptSegmentStyleSource
): TextStyle | undefined {
  if (!style && !segment.fontFamily) {
    return style;
  }
  return {
    ...style,
    fontFamily: segment.fontFamily ?? style?.fontFamily,
    sourceFontFamily: segment.fontFamily ?? style?.sourceFontFamily,
    fontFamilyAscii: undefined,
    fontFamilyHAnsi: undefined,
    fontFamilyEastAsia: undefined,
    fontFamilyComplexScript: undefined,
    fontSizePt: segment.fontSizePt ?? style?.fontSizePt,
    fontSizeCsPt: segment.fontSizePt ?? style?.fontSizeCsPt,
  };
}

export interface DocxScriptSegmentStyleSource {
  fontFamily?: string;
  fontSizePt?: number;
}

function cssFontFamilyToken(family: string): string {
  const trimmed = family.trim();
  if (!trimmed) return "sans-serif";
  if (trimmed.includes(",") || GENERIC_FAMILIES.test(trimmed)) {
    return trimmed;
  }
  const escaped = trimmed.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  return /\s/.test(escaped) ? `"${escaped}"` : escaped;
}

/** Canonical `font-family` value for one resolved script family. */
export function docxScriptFamilyCss(family?: string): string | undefined {
  const normalized = normalizeFamily(family);
  return normalized ? cssFontFamilyToken(normalized) : undefined;
}
