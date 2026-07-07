// Table-of-contents (TOC) paragraph predicates and field-instruction parsers.
//
// Extracted as a leaf module to break the field-helpers <-> line-height and
// field-helpers -> line-height -> paragraph-geometry -> synthetic-textbox
// circular module dependencies (review #2 F2). Both field-helpers and
// line-height (via the longer cycle through paragraph-geometry ->
// synthetic-textbox) need these predicates/parsers; housing them here lets
// every consumer import from a single acyclic leaf.
//
// Upstream editor.tsx: lines 15864-15899 (TOC style/level detection),
// 17653-17700 (PAGE/NUMPAGES/StyleRef field-instruction parsing).

import type { ParagraphNode } from "../../engine/types";
import { decodeXmlText } from "./xml-parsing";

const TABLE_OF_CONTENTS_STYLE_ID = /^toc(?:[\s_-]*\d+)?$/i;

export function isTableOfContentsStyle(styleId?: string): boolean {
  if (!styleId) {
    return false;
  }
  return TABLE_OF_CONTENTS_STYLE_ID.test(styleId.trim());
}

export function tableOfContentsLevel(paragraph: ParagraphNode): number | undefined {
  const candidates = [paragraph.style?.styleId, paragraph.style?.styleName];
  for (const candidate of candidates) {
    if (!candidate) {
      continue;
    }
    const match = candidate.trim().match(/^toc(?:[\s_-]*(\d+))?$/i);
    if (!match) {
      continue;
    }
    const parsedLevel = match[1] ? Number.parseInt(match[1], 10) : 1;
    if (Number.isFinite(parsedLevel) && parsedLevel > 0) {
      return Math.round(parsedLevel);
    }
    return 1;
  }
  return undefined;
}

export function isTableOfContentsParagraph(paragraph: ParagraphNode): boolean {
  return (
    isTableOfContentsStyle(paragraph.style?.styleId) ||
    isTableOfContentsStyle(paragraph.style?.styleName)
  );
}

export type PageFieldKind = "PAGE" | "NUMPAGES";

export function instructionTextToPageFieldKind(
  rawInstruction: string
): PageFieldKind | undefined {
  const normalized = decodeXmlText(rawInstruction)
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
  if (!normalized || normalized.includes("PAGEREF")) {
    return undefined;
  }

  if (/\bNUMPAGES\b/.test(normalized)) {
    return "NUMPAGES";
  }
  if (/\bPAGE\b/.test(normalized)) {
    return "PAGE";
  }

  return undefined;
}

export function instructionTextToStyleRefTarget(
  rawInstruction: string
): string | undefined {
  const normalized = decodeXmlText(rawInstruction).replace(/\s+/g, " ").trim();
  if (!normalized) {
    return undefined;
  }

  const match = normalized.match(/\bSTYLEREF\b\s+(?:"([^"]+)"|([^\s\\]+))/i);
  const target = (match?.[1] ?? match?.[2] ?? "").trim();
  return target.length > 0 ? target : undefined;
}
