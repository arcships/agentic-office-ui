// DOCX XML parsing helpers (continued): paragraph indent extraction from
// source XML, list- numbering indent resolution.
// Upstream editor.tsx: lines 15081-15324.

import type {
  NumberingDefinitionSet,
  ParagraphIndent,
  ParagraphNode,
} from "../../engine/types";
import { LIST_LEVEL_STEP_TWIPS } from "./constants";
import { paragraphExplicitIndentBySourceXml, setCacheEntry } from "./cache-utils";
import {
  effectiveNumberingNumIdForParagraph,
  findNumberingLevelDefinition,
} from "./numbering";

/**
 * Extract explicit indent values from the w:ind tag in a paragraph's source XML.
 *
 * Returns the indent defined directly in the paragraph properties, or undefined
 * when no w:ind tag is present or all indent attributes are absent/non-finite.
 * Results are cached on paragraphExplicitIndentBySourceXml.
 */
export function paragraphExplicitIndentTwips(
  paragraph: ParagraphNode
): ParagraphIndent | undefined {
  const sourceXml = paragraph.sourceXml;
  if (!sourceXml) {
    return undefined;
  }

  const cached = paragraphExplicitIndentBySourceXml.get(sourceXml);
  if (cached !== undefined) {
    return cached === null ? undefined : cached;
  }

  const paragraphPropertiesXml =
    sourceXml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/i)?.[0] ??
    sourceXml.match(/<w:pPr\b[^>]*\/>/i)?.[0];
  if (!paragraphPropertiesXml) {
    setCacheEntry(paragraphExplicitIndentBySourceXml, sourceXml, null as never);
    return undefined;
  }

  const indentTag = paragraphPropertiesXml.match(/<w:ind\b[^>]*\/?>/i)?.[0];
  if (!indentTag) {
    setCacheEntry(paragraphExplicitIndentBySourceXml, sourceXml, null as never);
    return undefined;
  }

  const parseIndentTwips = (attribute: string): number | undefined => {
    const match = indentTag.match(
      new RegExp(`\\b${attribute}="(-?\\d+)"`, "i")
    );
    if (!match?.[1]) {
      return undefined;
    }

    const parsed = Number.parseInt(match[1], 10);
    return Number.isFinite(parsed) ? parsed : undefined;
  };

  let firstLineTwips = parseIndentTwips("w:firstLine");
  const hangingTwips = parseIndentTwips("w:hanging");
  if (Number.isFinite(firstLineTwips) && Number.isFinite(hangingTwips)) {
    // ECMA-376: when both firstLine and hanging are present, firstLine is ignored.
    firstLineTwips = undefined;
  }

  const explicitIndent: ParagraphIndent = {
    leftTwips: parseIndentTwips("w:left"),
    firstLineTwips,
    hangingTwips,
  };
  const hasAnyExplicitIndent =
    Number.isFinite(explicitIndent.leftTwips) ||
    Number.isFinite(explicitIndent.firstLineTwips) ||
    Number.isFinite(explicitIndent.hangingTwips);
  if (!hasAnyExplicitIndent) {
    setCacheEntry(paragraphExplicitIndentBySourceXml, sourceXml, null as never);
    return undefined;
  }

  setCacheEntry(paragraphExplicitIndentBySourceXml, sourceXml, explicitIndent);
  return explicitIndent;
}

/**
 * Resolve the effective list paragraph indent, combining numbering-level
 * indentation with paragraph-level explicit indent and style indent.
 *
 * When a paragraph has active numbering, this function merges indent values
 * from the numbering level definition, the paragraph style, and any explicit
 * w:ind on the paragraph itself, following ECMA-376 priority rules:
 *
 * 1. Explicit paragraph w:ind (left) overrides style/numbering left.
 * 2. When both numbering level and paragraph style provide left indent, the
 *    level offset relative to the base level is added to the style value.
 * 3. firstLine/hanging from explicit w:ind override style and numbering.
 * 4. ECMA-376: firstLine is ignored when both firstLine and hanging are present.
 */
export function resolveListParagraphIndent(
  paragraph: ParagraphNode,
  numberingDefinitions?: NumberingDefinitionSet
): ParagraphIndent | undefined {
  const numbering = paragraph.style?.numbering;
  if (!numbering || !Number.isFinite(numbering.numId) || numbering.numId <= 0) {
    return paragraph.style?.indent;
  }

  if (!numberingDefinitions) {
    return paragraph.style?.indent;
  }

  const effectiveNumId =
    effectiveNumberingNumIdForParagraph(paragraph, numberingDefinitions) ??
    numbering.numId;
  const numberingRecoveryActive = effectiveNumId !== numbering.numId;
  const ilvl = Math.max(0, Math.round(numbering.ilvl ?? 0));
  const level = findNumberingLevelDefinition(
    numberingDefinitions,
    effectiveNumId,
    ilvl
  );
  const baseLevel = findNumberingLevelDefinition(
    numberingDefinitions,
    effectiveNumId,
    0
  );
  const levelIndent = level?.indent;
  const styleIndent = paragraph.style?.indent;
  const numberingHasVisibleMarker = Boolean(
    (level?.text && level.text.trim().length > 0) || level?.pictureBullet?.src
  );
  const numberingProvidesUsableIndent = Boolean(
    Number.isFinite(levelIndent?.leftTwips) ||
      Number.isFinite(levelIndent?.firstLineTwips) ||
      Number.isFinite(levelIndent?.hangingTwips)
  );
  const styleLeftTwips = styleIndent?.leftTwips;
  const explicitParagraphIndent = paragraphExplicitIndentTwips(paragraph);
  const explicitParagraphLeftTwips = explicitParagraphIndent?.leftTwips;
  const explicitParagraphFirstLineTwips =
    explicitParagraphIndent?.firstLineTwips;
  const explicitParagraphHangingTwips = explicitParagraphIndent?.hangingTwips;
  const hasExplicitParagraphFirstLineTwips = Number.isFinite(
    explicitParagraphFirstLineTwips
  );
  const hasExplicitParagraphHangingTwips = Number.isFinite(
    explicitParagraphHangingTwips
  );
  const preferRecoveredNumberingTextIndent = numberingRecoveryActive;
  if (!numberingHasVisibleMarker && !numberingProvidesUsableIndent) {
    return styleIndent;
  }
  const baseLevelLeftTwips = Number.isFinite(baseLevel?.indent?.leftTwips)
    ? baseLevel?.indent?.leftTwips ?? 0
    : Number.isFinite(styleLeftTwips)
    ? styleLeftTwips
    : undefined;
  const levelLeftTwips = levelIndent?.leftTwips;
  const hasExplicitLevelLeftTwips = Number.isFinite(levelLeftTwips);
  const hasExplicitStyleLeftTwips = Number.isFinite(styleLeftTwips);
  const hasExplicitParagraphLeftTwips = Number.isFinite(
    explicitParagraphLeftTwips
  );
  let nextLeftTwips = styleLeftTwips;

  if (hasExplicitParagraphLeftTwips) {
    nextLeftTwips = explicitParagraphLeftTwips;
  } else if (
    Number.isFinite(levelLeftTwips) &&
    Number.isFinite(baseLevelLeftTwips) &&
    Number.isFinite(styleLeftTwips)
  ) {
    const levelOffsetTwips = Math.max(
      0,
      (levelLeftTwips ?? 0) - (baseLevelLeftTwips ?? 0)
    );
    // When paragraph style indentation comes from defaults (often 0), it should not
    // suppress the numbering level indentation from DOCX.
    const styleUsesListBaseIndent =
      (styleLeftTwips as number) >= (baseLevelLeftTwips as number) - 120;
    nextLeftTwips = styleUsesListBaseIndent
      ? (styleLeftTwips ?? 0) + levelOffsetTwips
      : levelLeftTwips ?? 0;
  } else if (Number.isFinite(levelLeftTwips)) {
    nextLeftTwips = levelLeftTwips;
  } else if (ilvl > 0) {
    nextLeftTwips = ilvl * LIST_LEVEL_STEP_TWIPS;
  }

  // Some documents provide numbering but omit usable paragraph indents for list
  // paragraphs. Preserve explicit zero indents from OOXML and only synthesize
  // a fallback list indent when both style and numbering level omit left indents.
  if (!Number.isFinite(nextLeftTwips)) {
    if (hasExplicitParagraphLeftTwips) {
      nextLeftTwips = explicitParagraphLeftTwips;
    } else if (hasExplicitLevelLeftTwips) {
      nextLeftTwips = levelLeftTwips;
    } else if (hasExplicitStyleLeftTwips) {
      nextLeftTwips = styleLeftTwips;
    } else {
      nextLeftTwips = Math.max(
        LIST_LEVEL_STEP_TWIPS,
        (ilvl + 1) * LIST_LEVEL_STEP_TWIPS
      );
    }
  } else if ((nextLeftTwips as number) <= 0) {
    if (hasExplicitParagraphLeftTwips) {
      nextLeftTwips = explicitParagraphLeftTwips;
    } else if (hasExplicitLevelLeftTwips) {
      nextLeftTwips = levelLeftTwips;
    } else if (hasExplicitStyleLeftTwips) {
      nextLeftTwips = styleLeftTwips;
    } else if (
      Number.isFinite(levelLeftTwips) &&
      (levelLeftTwips as number) > 0
    ) {
      nextLeftTwips = levelLeftTwips;
    } else if (
      Number.isFinite(styleLeftTwips) &&
      (styleLeftTwips as number) > 0
    ) {
      nextLeftTwips = styleLeftTwips;
    } else {
      nextLeftTwips = Math.max(
        LIST_LEVEL_STEP_TWIPS,
        (ilvl + 1) * LIST_LEVEL_STEP_TWIPS
      );
    }
  }

  if (!Number.isFinite(nextLeftTwips)) {
    return styleIndent;
  }

  const nextLeftTwipsRounded = Math.max(0, Math.round(nextLeftTwips ?? 0));
  let nextFirstLineTwips: number | undefined;
  let nextHangingTwips: number | undefined;
  if (preferRecoveredNumberingTextIndent) {
    nextFirstLineTwips = levelIndent?.firstLineTwips;
    nextHangingTwips = levelIndent?.hangingTwips;
  } else if (
    hasExplicitParagraphFirstLineTwips ||
    hasExplicitParagraphHangingTwips
  ) {
    // Paragraph-level w:ind overrides numbering/style indentation semantics.
    // If firstLine is explicitly present without hanging, do not inherit hanging.
    // If hanging is explicitly present without firstLine, do not inherit firstLine.
    nextFirstLineTwips = hasExplicitParagraphFirstLineTwips
      ? explicitParagraphFirstLineTwips
      : undefined;
    nextHangingTwips = hasExplicitParagraphHangingTwips
      ? explicitParagraphHangingTwips
      : undefined;
  } else {
    nextFirstLineTwips = Number.isFinite(styleIndent?.firstLineTwips)
      ? styleIndent?.firstLineTwips
      : levelIndent?.firstLineTwips;
    nextHangingTwips = Number.isFinite(styleIndent?.hangingTwips)
      ? styleIndent?.hangingTwips
      : levelIndent?.hangingTwips;
  }

  if (
    Number.isFinite(nextFirstLineTwips) &&
    Number.isFinite(nextHangingTwips)
  ) {
    // ECMA-376: firstLine is ignored when both values are present.
    nextFirstLineTwips = undefined;
  }

  return {
    ...styleIndent,
    leftTwips: nextLeftTwipsRounded,
    firstLineTwips: nextFirstLineTwips,
    hangingTwips: nextHangingTwips,
  };
}
