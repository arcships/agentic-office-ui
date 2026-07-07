// Field helpers: TOC/field detection, tab-leader layout, PAGE/NUMPAGES/StyleRef
// field-sequence parsing, tab-leader styling, note-marker labels, date helpers.
// Upstream editor.tsx: lines 15864-16060, 17653-18107.

import type { ParagraphNode, TextRunNode } from "../../engine/types";
import { twipsToPixels } from "../../viewer/section-layout";
import { setCacheEntry } from "./cache-utils";
import { xmlAttribute } from "./ooxml-helpers";
import { decodeXmlText } from "./xml-parsing";
import { numberToRoman } from "./numbering";
import { formFieldDisplayValue } from "./paragraph-inspect";
import { resolveParagraphFirstLineLeftTabStopsPx } from "./line-height";
import {
  isTableOfContentsParagraph,
  isTableOfContentsStyle,
  tableOfContentsLevel,
  instructionTextToPageFieldKind,
  instructionTextToStyleRefTarget
} from "./paragraph-toc";
import type { PageFieldKind } from "./paragraph-toc";

// Re-export the TOC predicates and field-instruction parsers from the
// paragraph-toc leaf module so existing barrel consumers are unaffected.
// The canonical declarations live in paragraph-toc.ts (acyclic leaf).
export {
  isTableOfContentsStyle,
  tableOfContentsLevel,
  isTableOfContentsParagraph,
  instructionTextToPageFieldKind,
  instructionTextToStyleRefTarget
};
export type { PageFieldKind } from "./paragraph-toc";

export function paragraphUsesTabLeaders(paragraph: ParagraphNode): boolean {
  if (!isTableOfContentsParagraph(paragraph)) {
    return false;
  }

  const tabStops = paragraph.style?.tabStops ?? [];
  if (
    tabStops.some(
      (tabStop) => tabStop.alignment === "right" || tabStop.leader === "dot"
    )
  ) {
    return true;
  }

  return paragraph.children.some((child) => {
    if (child.type === "text") {
      return child.text.includes("\t");
    }
    if (child.type === "form-field") {
      return formFieldDisplayValue(child).includes("\t");
    }
    return false;
  });
}

export function paragraphLeadingTabStop(paragraph: ParagraphNode):
  | {
      alignment?: "left" | "center" | "right" | "decimal" | "bar";
      leader?: "none" | "dot" | "hyphen" | "underscore" | "middleDot";
      positionTwips?: number;
    }
  | undefined {
  const tabStops = paragraph.style?.tabStops ?? [];
  const explicitTabStop = tabStops.find(
    (tabStop) => tabStop.alignment === "right" || tabStop.leader === "dot"
  );
  if (explicitTabStop) {
    return explicitTabStop;
  }
  if (isTableOfContentsParagraph(paragraph)) {
    return {
      alignment: "right",
      leader: "dot",
    };
  }
  return undefined;
}

export function tableOfContentsLeadingLeftTabStopPx(
  paragraph: ParagraphNode
): number | undefined {
  if (!isTableOfContentsParagraph(paragraph)) {
    return undefined;
  }

  return resolveParagraphFirstLineLeftTabStopsPx(paragraph)[0];
}

export function paragraphContainsTabCharacter(paragraph: ParagraphNode): boolean {
  return paragraph.children.some((child) => {
    if (child.type === "text") {
      return child.text.includes("\t");
    }
    if (child.type === "form-field") {
      return formFieldDisplayValue(child).includes("\t");
    }
    return false;
  });
}

export function paragraphTabCharacterCount(paragraph: ParagraphNode): number {
  return paragraph.children.reduce((count, child) => {
    const text =
      child.type === "text"
        ? child.text
        : child.type === "form-field"
        ? formFieldDisplayValue(child)
        : "";
    if (!text) {
      return count;
    }
    const matches = text.match(/\t/g);
    return count + (matches ? matches.length : 0);
  }, 0);
}

export type ParagraphAnchoredTabLayout = "none" | "center-right" | "center" | "right";

export function paragraphAnchoredTabLayout(
  paragraph: ParagraphNode,
  options?: {
    withinHeaderFooter?: boolean;
  }
): ParagraphAnchoredTabLayout {
  const tabStops = paragraph.style?.tabStops ?? [];
  const hasLeft = tabStops.some((tabStop) => tabStop.alignment === "left");
  const hasCenter = tabStops.some((tabStop) => tabStop.alignment === "center");
  const hasRight = tabStops.some((tabStop) => tabStop.alignment === "right");
  const tabCount = paragraphTabCharacterCount(paragraph);
  const withinHeaderFooter = options?.withinHeaderFooter === true;

  if (hasLeft && (hasCenter || hasRight)) {
    return "none";
  }

  if (hasCenter && hasRight) {
    if (tabCount >= 2) {
      return "center-right";
    }
    if (withinHeaderFooter && tabCount >= 1) {
      return "center-right";
    }
    return "none";
  }
  if (hasCenter && tabCount === 1) {
    return "center";
  }
  if (hasRight && tabCount === 1) {
    return "right";
  }

  return "none";
}

export function paragraphFirstTabStopPx(
  paragraph: ParagraphNode,
  alignment: "center" | "right"
): number | undefined {
  return (paragraph.style?.tabStops ?? [])
    .filter((tabStop) => tabStop.alignment === alignment)
    .map((tabStop) => twipsToPixels(tabStop.positionTwips))
    .filter(
      (positionPx): positionPx is number =>
        Number.isFinite(positionPx) && (positionPx as number) > 0
    )
    .sort((left, right) => left - right)[0];
}

export function paragraphUsesCenterTabLayout(paragraph: ParagraphNode): boolean {
  return paragraphAnchoredTabLayout(paragraph) === "center";
}

export function paragraphUsesRightTabLayout(paragraph: ParagraphNode): boolean {
  return paragraphAnchoredTabLayout(paragraph) === "right";
}

export function paragraphUsesCenterRightTabLayout(paragraph: ParagraphNode): boolean {
  return paragraphAnchoredTabLayout(paragraph) === "center-right";
}

export interface PageFieldValueToken {
  kind: PageFieldKind;
  rawText: string;
}

export interface StyleRefFieldValueToken {
  target: string;
  rawText: string;
}

export function paragraphPageFieldSequence(paragraph: ParagraphNode): PageFieldKind[] {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return [];
  }

  const fields: PageFieldKind[] = [];
  for (const instructionMatch of xml.matchAll(
    /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/gi
  )) {
    const kind = instructionTextToPageFieldKind(instructionMatch[1] ?? "");
    if (kind) {
      fields.push(kind);
    }
  }
  for (const simpleFieldMatch of xml.matchAll(
    /<w:fldSimple\b[^>]*\bw:instr="([^"]+)"[^>]*>/gi
  )) {
    const kind = instructionTextToPageFieldKind(simpleFieldMatch[1] ?? "");
    if (kind) {
      fields.push(kind);
    }
  }

  return fields;
}

const pageFieldValueSequenceBySourceXml = new Map<
  string,
  PageFieldValueToken[]
>();

export function paragraphPageFieldValueSequence(
  paragraph: ParagraphNode
): PageFieldValueToken[] {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return [];
  }

  const cached = pageFieldValueSequenceBySourceXml.get(xml);
  if (cached) {
    return cached;
  }

  const values: PageFieldValueToken[] = [];
  const fieldStack: Array<{ kind?: PageFieldKind; inResult: boolean }> = [];
  const tokenPattern =
    /<w:fldSimple\b[^>]*\bw:instr="([^"]+)"[^>]*>[\s\S]*?<\/w:fldSimple>|<w:r\b[\s\S]*?<\/w:r>/gi;

  for (const tokenMatch of xml.matchAll(tokenPattern)) {
    const tokenXml = tokenMatch[0] ?? "";
    if (!tokenXml) {
      continue;
    }

    if (/^<w:fldSimple\b/i.test(tokenXml)) {
      const kind = instructionTextToPageFieldKind(tokenMatch[1] ?? "");
      if (!kind) {
        continue;
      }

      for (const textMatch of tokenXml.matchAll(
        /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
      )) {
        values.push({
          kind,
          rawText: decodeXmlText(textMatch[1] ?? ""),
        });
      }
      continue;
    }

    const beginCount =
      tokenXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="begin"[^>]*\/?>/gi)
        ?.length ?? 0;
    for (let index = 0; index < beginCount; index += 1) {
      fieldStack.push({ inResult: false });
    }

    for (const instructionMatch of tokenXml.matchAll(
      /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/gi
    )) {
      const kind = instructionTextToPageFieldKind(instructionMatch[1] ?? "");
      if (!kind || fieldStack.length === 0) {
        continue;
      }

      let assigned = false;
      for (
        let stackIndex = fieldStack.length - 1;
        stackIndex >= 0;
        stackIndex -= 1
      ) {
        if (fieldStack[stackIndex].kind === undefined) {
          fieldStack[stackIndex].kind = kind;
          assigned = true;
          break;
        }
      }
      if (!assigned) {
        fieldStack[fieldStack.length - 1].kind = kind;
      }
    }

    const separateCount =
      tokenXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="separate"[^>]*\/?>/gi)
        ?.length ?? 0;
    for (let index = 0; index < separateCount; index += 1) {
      for (
        let stackIndex = fieldStack.length - 1;
        stackIndex >= 0;
        stackIndex -= 1
      ) {
        if (!fieldStack[stackIndex].inResult) {
          fieldStack[stackIndex].inResult = true;
          break;
        }
      }
    }

    const activeFieldKind = (() => {
      for (
        let stackIndex = fieldStack.length - 1;
        stackIndex >= 0;
        stackIndex -= 1
      ) {
        const stackEntry = fieldStack[stackIndex];
        if (stackEntry.inResult && stackEntry.kind) {
          return stackEntry.kind;
        }
      }
      return undefined;
    })();

    if (activeFieldKind) {
      for (const textMatch of tokenXml.matchAll(
        /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
      )) {
        values.push({
          kind: activeFieldKind,
          rawText: decodeXmlText(textMatch[1] ?? ""),
        });
      }
    }

    const endCount =
      tokenXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="end"[^>]*\/?>/gi)
        ?.length ?? 0;
    for (let index = 0; index < endCount; index += 1) {
      fieldStack.pop();
    }
  }

  setCacheEntry(pageFieldValueSequenceBySourceXml, xml, values);
  return values;
}

const styleRefFieldValueSequenceBySourceXml = new Map<
  string,
  StyleRefFieldValueToken[]
>();

export function paragraphStyleRefFieldValueSequence(
  paragraph: ParagraphNode
): StyleRefFieldValueToken[] {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return [];
  }

  const cached = styleRefFieldValueSequenceBySourceXml.get(xml);
  if (cached) {
    return cached;
  }

  const values: StyleRefFieldValueToken[] = [];
  const fieldStack: Array<{ target?: string; inResult: boolean }> = [];
  const tokenPattern =
    /<w:fldSimple\b[^>]*\bw:instr="([^"]+)"[^>]*>[\s\S]*?<\/w:fldSimple>|<w:r\b[\s\S]*?<\/w:r>/gi;

  for (const tokenMatch of xml.matchAll(tokenPattern)) {
    const tokenXml = tokenMatch[0] ?? "";
    if (!tokenXml) {
      continue;
    }

    if (/^<w:fldSimple\b/i.test(tokenXml)) {
      const target = instructionTextToStyleRefTarget(tokenMatch[1] ?? "");
      if (!target) {
        continue;
      }

      for (const textMatch of tokenXml.matchAll(
        /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
      )) {
        values.push({
          target,
          rawText: decodeXmlText(textMatch[1] ?? ""),
        });
      }
      continue;
    }

    const beginCount =
      tokenXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="begin"[^>]*\/?>/gi)
        ?.length ?? 0;
    for (let index = 0; index < beginCount; index += 1) {
      fieldStack.push({ inResult: false });
    }

    for (const instructionMatch of tokenXml.matchAll(
      /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/gi
    )) {
      const target = instructionTextToStyleRefTarget(instructionMatch[1] ?? "");
      if (!target || fieldStack.length === 0) {
        continue;
      }

      let assigned = false;
      for (
        let stackIndex = fieldStack.length - 1;
        stackIndex >= 0;
        stackIndex -= 1
      ) {
        if (fieldStack[stackIndex].target === undefined) {
          fieldStack[stackIndex].target = target;
          assigned = true;
          break;
        }
      }
      if (!assigned) {
        fieldStack[fieldStack.length - 1].target = target;
      }
    }

    const separateCount =
      tokenXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="separate"[^>]*\/?>/gi)
        ?.length ?? 0;
    for (let index = 0; index < separateCount; index += 1) {
      for (
        let stackIndex = fieldStack.length - 1;
        stackIndex >= 0;
        stackIndex -= 1
      ) {
        if (!fieldStack[stackIndex].inResult) {
          fieldStack[stackIndex].inResult = true;
          break;
        }
      }
    }

    const activeFieldTarget = (() => {
      for (
        let stackIndex = fieldStack.length - 1;
        stackIndex >= 0;
        stackIndex -= 1
      ) {
        const stackEntry = fieldStack[stackIndex];
        if (stackEntry.inResult && stackEntry.target) {
          return stackEntry.target;
        }
      }
      return undefined;
    })();

    if (activeFieldTarget) {
      for (const textMatch of tokenXml.matchAll(
        /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
      )) {
        values.push({
          target: activeFieldTarget,
          rawText: decodeXmlText(textMatch[1] ?? ""),
        });
      }
    }

    const endCount =
      tokenXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="end"[^>]*\/?>/gi)
        ?.length ?? 0;
    for (let index = 0; index < endCount; index += 1) {
      fieldStack.pop();
    }
  }

  setCacheEntry(styleRefFieldValueSequenceBySourceXml, xml, values);
  return values;
}

export function normalizeStyleRefTarget(value: string): string {
  return value.replace(/\s+/g, " ").trim().toLowerCase();
}

export function paragraphStyledRunText(
  paragraph: ParagraphNode,
  styleRefTarget: string
): string | undefined {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return undefined;
  }

  const normalizedTarget = normalizeStyleRefTarget(styleRefTarget);
  if (!normalizedTarget) {
    return undefined;
  }

  const textChunks: string[] = [];
  for (const runMatch of xml.matchAll(/<w:r\b[\s\S]*?<\/w:r>/gi)) {
    const runXml = runMatch[0] ?? "";
    if (!runXml) {
      continue;
    }

    const runStyleTag = runXml.match(/<w:rStyle\b[^>]*>/i)?.[0] ?? "";
    const runStyleValue = xmlAttribute(runStyleTag, "w:val");
    if (normalizeStyleRefTarget(runStyleValue ?? "") !== normalizedTarget) {
      continue;
    }

    for (const textMatch of runXml.matchAll(
      /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
    )) {
      const value = decodeXmlText(textMatch[1] ?? "");
      if (value.length > 0) {
        textChunks.push(value);
      }
    }

    if (/<w:tab\b[^>]*\/?>/i.test(runXml)) {
      textChunks.push("\t");
    }
  }

  const joined = textChunks.join("").trim();
  return joined.length > 0 ? joined : undefined;
}

export function tabLeaderStyle(
  leader: string | undefined,
  color: string | undefined
): Record<string, string | number | undefined> {
  const normalizedLeader = leader === "middleDot" ? "dot" : leader;
  const resolvedColor = color || "currentColor";
  if (normalizedLeader === "hyphen") {
    return {
      backgroundImage: `linear-gradient(to right, transparent 0, transparent 4px, ${resolvedColor} 4px, ${resolvedColor} 5px, transparent 5px, transparent 8px)`,
      backgroundSize: "8px 1px",
      backgroundPosition: "0 80%",
      backgroundRepeat: "repeat-x",
    };
  }
  if (normalizedLeader === "underscore") {
    return {
      backgroundImage: `linear-gradient(to right, ${resolvedColor} 0, ${resolvedColor} 1px, transparent 1px, transparent 4px)`,
      backgroundSize: "4px 1px",
      backgroundPosition: "0 90%",
      backgroundRepeat: "repeat-x",
    };
  }

  return {
    backgroundImage: `radial-gradient(${resolvedColor} 1px, transparent 1px)`,
    backgroundSize: "7px 9px",
    backgroundPosition: "0 72%",
    backgroundRepeat: "repeat-x",
  };
}

export function noteMarkerLabel(
  noteReference: TextRunNode["noteReference"],
  footnoteDisplayIndexById: Map<number, number>,
  endnoteDisplayIndexById: Map<number, number>
): string | undefined {
  if (!noteReference) {
    return undefined;
  }

  const index =
    noteReference.kind === "footnote"
      ? footnoteDisplayIndexById.get(noteReference.id)
      : endnoteDisplayIndexById.get(noteReference.id);

  if (noteReference.kind === "footnote") {
    const value = index ?? Math.max(1, Math.round(noteReference.id));
    if (!Number.isFinite(value) || value <= 0) {
      return undefined;
    }
    return String(Math.round(value));
  }

  const romanValue = index ?? Math.max(1, Math.round(noteReference.id));
  if (!Number.isFinite(romanValue) || romanValue <= 0) {
    return undefined;
  }

  return numberToRoman(Math.round(romanValue)).toLowerCase();
}

export function normalizeDateInputValue(value?: string): string | undefined {
  if (!value) {
    return undefined;
  }

  const trimmed = value.trim();
  if (!trimmed) {
    return undefined;
  }

  const directMatch = trimmed.match(/^(\d{4}-\d{2}-\d{2})/);
  if (directMatch?.[1]) {
    return directMatch[1];
  }

  const parsed = new Date(trimmed);
  if (Number.isNaN(parsed.getTime())) {
    return undefined;
  }

  return parsed.toISOString().slice(0, 10);
}
