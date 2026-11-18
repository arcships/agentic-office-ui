// Document-level paragraph defaults from word/settings.xml.
//
// The WASM model builder does not surface `w:compat/w:compatSetting`
// (compatibility mode) or `w:defaultTabStop`, and it flattens `w:spacing` down
// to the four numeric attributes, so the two layout inputs Word needs beyond
// the explicit values are recovered from the package here:
//
//   - `metadata.compatibility.compatibilityMode` — decides whether the
//     automatic (HTML) paragraph spacing below is honoured (Word only does so
//     in compatibility modes 12 and 15).
//   - `metadata.defaultTabStopTwips` — the fallback tab grid; Word's default is
//     720 twips (0.5"), but Chinese templates commonly ship 420 twips.
//
// Paragraph-level `w:beforeAutospacing` / `w:afterAutospacing` are read by
// `parseParagraphAutoSpacing` from the paragraph properties XML kept on
// imported paragraphs.

import type { DocModel } from "./types";
import type { OoxmlPackage } from "./ooxml-core";
import { xmlAttribute } from "./text-typography";

const SETTINGS_PART = "word/settings.xml";

const DEFAULT_TAB_STOP_TAG_PATTERN = /<w:defaultTabStop\b[^>]*\/?>/i;
const COMPAT_SETTING_TAG_PATTERN = /<w:compatSetting\b[^>]*\/?>/gi;
const PARAGRAPH_PROPERTIES_TAG_PATTERN = /<w:pPr\b[\s\S]*?<\/w:pPr>/i;
const SPACING_TAG_PATTERN = /<w:spacing\b[^>]*\/?>/i;

/** OOXML `ST_OnOff` values that mean "on". */
function parseOnOff(value: string | undefined): boolean | undefined {
  if (value === undefined) {
    return undefined;
  }
  const normalized = value.trim().toLowerCase();
  if (normalized === "1" || normalized === "true" || normalized === "on") {
    return true;
  }
  if (normalized === "0" || normalized === "false" || normalized === "off") {
    return false;
  }
  return undefined;
}

function positiveInteger(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const parsed = Number(value.trim());
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return undefined;
  }
  return Math.round(parsed);
}

/**
 * `w:defaultTabStop/@w:val` in twips. Word writes 720 for the 0.5" default;
 * templates that change it (420 in many CJK documents) shift every tab.
 */
export function parseDefaultTabStopTwips(
  settingsXml: string | undefined
): number | undefined {
  if (!settingsXml) {
    return undefined;
  }
  const tag = settingsXml.match(DEFAULT_TAB_STOP_TAG_PATTERN)?.[0];
  if (!tag) {
    return undefined;
  }
  return positiveInteger(xmlAttribute(tag, "w:val"));
}

/**
 * `w:compat/w:compatSetting[@w:name='compatibilityMode']/@w:val`.
 * Only the Word-namespaced setting counts; other compatSetting names (e.g.
 * `overrideTableStyleFontSizeAndJustification`) carry unrelated values.
 */
export function parseCompatibilityMode(
  settingsXml: string | undefined
): number | undefined {
  if (!settingsXml) {
    return undefined;
  }
  const compatStart = settingsXml.search(/<w:compat\b/i);
  if (compatStart < 0) {
    return undefined;
  }
  // `w:compat` is a small block; scanning to the end keeps the parser from
  // depending on balanced-tag bookkeeping for a lookup this narrow.
  const compatXml = settingsXml.slice(compatStart);
  for (const match of compatXml.matchAll(COMPAT_SETTING_TAG_PATTERN)) {
    const tag = match[0];
    const name = xmlAttribute(tag, "w:name")?.trim();
    if (name !== "compatibilityMode") {
      continue;
    }
    const mode = Number(xmlAttribute(tag, "w:val"));
    if (Number.isFinite(mode) && mode > 0) {
      return Math.round(mode);
    }
  }
  return undefined;
}

export interface ParsedParagraphAutoSpacing {
  beforeAutoSpacing?: boolean;
  afterAutoSpacing?: boolean;
}

/**
 * `w:spacing/@w:beforeAutospacing` and `@w:afterAutospacing` of a paragraph.
 * Returns `undefined` when the paragraph names neither attribute, so callers
 * can leave an already-populated model untouched.
 */
export function parseParagraphAutoSpacing(
  paragraphPropertiesXml: string | undefined
): ParsedParagraphAutoSpacing | undefined {
  if (!paragraphPropertiesXml) {
    return undefined;
  }
  // Accept a whole `<w:p>` (the importer's usual input) or the properties
  // element itself.
  const paragraphProperties =
    paragraphPropertiesXml.match(PARAGRAPH_PROPERTIES_TAG_PATTERN)?.[0] ??
    paragraphPropertiesXml;
  const spacingTag = paragraphProperties.match(SPACING_TAG_PATTERN)?.[0];
  if (!spacingTag) {
    return undefined;
  }

  const beforeAutoSpacing = parseOnOff(
    xmlAttribute(spacingTag, "w:beforeAutospacing")
  );
  const afterAutoSpacing = parseOnOff(
    xmlAttribute(spacingTag, "w:afterAutospacing")
  );
  if (beforeAutoSpacing === undefined && afterAutoSpacing === undefined) {
    return undefined;
  }
  return { beforeAutoSpacing, afterAutoSpacing };
}

/**
 * Attaches the layout inputs that live in word/settings.xml to the model.
 */
export function attachDocumentDefaults(
  model: DocModel,
  pkg: OoxmlPackage
): DocModel {
  const settingsXml = pkg.parts.get(SETTINGS_PART)?.content;
  if (!settingsXml) {
    return model;
  }

  const defaultTabStopTwips = parseDefaultTabStopTwips(settingsXml);
  const compatibilityMode = parseCompatibilityMode(settingsXml);
  const currentCompatibility = model.metadata.compatibility;
  const hasCompatibilityMode =
    currentCompatibility?.compatibilityMode !== undefined;

  if (
    defaultTabStopTwips === undefined &&
    (compatibilityMode === undefined || hasCompatibilityMode)
  ) {
    return model;
  }

  const metadata = { ...model.metadata };
  if (defaultTabStopTwips !== undefined && metadata.defaultTabStopTwips === undefined) {
    metadata.defaultTabStopTwips = defaultTabStopTwips;
  }
  if (compatibilityMode !== undefined && !hasCompatibilityMode) {
    metadata.compatibility = {
      ...currentCompatibility,
      compatibilityMode,
    };
  }
  return { ...model, metadata };
}
