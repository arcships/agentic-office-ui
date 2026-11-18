// Inline revision cursor mapping for DOCX paragraphs.
//
// The tracked-change and comment collectors parse revision markup from a
// paragraph's `sourceXml`, but the run renderer walks `paragraph.children`.
// Counting XML tokens therefore drifts whenever one model child represents
// several XML tokens (a merged tab + text run, a form field, ...). This module
// scans the visible XML units and maps each one onto the model child that
// actually renders it, so both sides share the renderer's visible-child cursor.

import type { ParagraphNode } from "../../engine/types";
import { decodeXmlText, extractBalancedTagRanges } from "./xml-parsing";

export type ParagraphVisibleUnitKind = "text" | "tab" | "break" | "note" | "image";

export interface ParagraphVisibleUnit {
  kind: ParagraphVisibleUnitKind;
  /** Text the renderer will produce for this unit ("\t", "\n", note marker, ...). */
  text: string;
  isNote: boolean;
  /** offset in the paragraph sourceXml */
  start: number;
  /** end offset in the paragraph sourceXml */
  end: number;
  /**
   * Index into the renderer's tracked cursor space (paragraph children with
   * form fields skipped). Undefined when the model has no matching child.
   */
  visibleChildIndex?: number;
}

const VISIBLE_UNIT_PATTERN =
  /<(?:w|a):t\b[^>]*>([\s\S]*?)<\/(?:w|a):t>|<w:tab\b[^>]*\/?>|<w:(?:br|cr)\b[^>]*\/?>|<w:(?:footnoteReference|endnoteReference)\b[^>]*\/?>|<(?:w:drawing|w:pict)\b[^>]*>[\s\S]*?<\/(?:w:drawing|w:pict)>/gi;

function rangeContains(
  ranges: Array<{ start: number; end: number }>,
  offset: number
): boolean {
  return ranges.some((range) => offset >= range.start && offset < range.end);
}

/**
 * Content controls that the model builder turns into a single form-field child
 * instead of exposing their inner runs as text children.
 */
function formFieldContentRanges(
  sourceXml: string
): Array<{ start: number; end: number }> {
  if (!sourceXml.includes("w:sdt")) {
    return [];
  }

  return extractBalancedTagRanges(sourceXml, "w:sdt").filter((range) => {
    const sdtXml = sourceXml.slice(range.start, range.end);
    const sdtPrRange = extractBalancedTagRanges(sdtXml, "w:sdtPr")[0];
    if (!sdtPrRange) {
      return false;
    }
    const sdtProperties = sdtXml.slice(sdtPrRange.start, sdtPrRange.end);
    return /<w:(?:text|checkbox|date|dropDownList|comboBox)\b|<w:ffData\b/i.test(
      sdtProperties
    );
  });
}

/**
 * Scans a paragraph's XML for the units that produce renderable children:
 * text runs, tabs, line breaks, note references and images. Text-box content
 * is skipped because the model builder renders it elsewhere.
 */
export function collectParagraphVisibleUnits(
  sourceXml: string
): ParagraphVisibleUnit[] {
  if (!sourceXml) {
    return [];
  }

  const textBoxRanges = extractBalancedTagRanges(sourceXml, "w:txbxContent");
  const formFieldRanges = formFieldContentRanges(sourceXml);
  const units: ParagraphVisibleUnit[] = [];

  for (const match of sourceXml.matchAll(VISIBLE_UNIT_PATTERN)) {
    const start = match.index ?? 0;
    if (
      rangeContains(textBoxRanges, start) ||
      rangeContains(formFieldRanges, start)
    ) {
      continue;
    }

    const fullMatch = match[0] ?? "";
    const end = start + fullMatch.length;

    if (match[1] !== undefined) {
      const text = decodeXmlText(match[1] ?? "");
      // Empty <w:t/> elements do not create model children.
      if (text.length === 0) {
        continue;
      }
      units.push({ kind: "text", text, isNote: false, start, end });
      continue;
    }

    if (/^<w:tab\b/i.test(fullMatch)) {
      units.push({ kind: "tab", text: "\t", isNote: false, start, end });
      continue;
    }
    if (/^<w:(?:br|cr)\b/i.test(fullMatch)) {
      units.push({ kind: "break", text: "\n", isNote: false, start, end });
      continue;
    }
    if (/^<w:(?:footnoteReference|endnoteReference)\b/i.test(fullMatch)) {
      units.push({ kind: "note", text: "", isNote: true, start, end });
      continue;
    }
    if (/^<(?:w:drawing|w:pict)\b/i.test(fullMatch)) {
      units.push({ kind: "image", text: "", isNote: false, start, end });
    }
  }

  return units;
}

function isTextLikeUnit(unit: ParagraphVisibleUnit): boolean {
  return unit.kind === "text" || unit.kind === "tab" || unit.kind === "break";
}

function unitMatchesChild(
  unit: ParagraphVisibleUnit,
  child: ParagraphNode["children"][number]
): boolean {
  if (unit.kind === "image") {
    return child.type === "image";
  }
  if (child.type !== "text") {
    return false;
  }
  if (unit.isNote) {
    return Boolean(child.noteReference);
  }
  if (child.noteReference) {
    return false;
  }
  return child.text === unit.text;
}

/**
 * Aligns visible XML units with the paragraph's model children. Several units
 * may map to one child (merged runs); form-field children consume no cursor
 * step because the run renderer skips them.
 */
export function mapVisibleUnitsToParagraphChildren(
  units: ParagraphVisibleUnit[],
  children: ParagraphNode["children"]
): void {
  let childIndex = 0;
  let visibleChildIndex = 0;
  let pendingMergedChild:
    | { visibleIndex: number; remainingText: string }
    | undefined;

  for (const unit of units) {
    if (pendingMergedChild) {
      if (
        isTextLikeUnit(unit) &&
        pendingMergedChild.remainingText.startsWith(unit.text)
      ) {
        unit.visibleChildIndex = pendingMergedChild.visibleIndex;
        pendingMergedChild.remainingText =
          pendingMergedChild.remainingText.slice(unit.text.length);
        if (pendingMergedChild.remainingText.length === 0) {
          pendingMergedChild = undefined;
        }
        continue;
      }
      pendingMergedChild = undefined;
    }

    while (childIndex < children.length) {
      const child = children[childIndex];
      if (child.type === "form-field") {
        // Renderer skips form fields without advancing the tracked cursor.
        childIndex += 1;
        continue;
      }

      if (unitMatchesChild(unit, child)) {
        unit.visibleChildIndex = visibleChildIndex;
        childIndex += 1;
        visibleChildIndex += 1;
        break;
      }

      if (
        child.type === "text" &&
        !child.noteReference &&
        isTextLikeUnit(unit) &&
        unit.text.length > 0 &&
        child.text.startsWith(unit.text)
      ) {
        unit.visibleChildIndex = visibleChildIndex;
        const remainingText = child.text.slice(unit.text.length);
        if (remainingText.length > 0) {
          pendingMergedChild = { visibleIndex: visibleChildIndex, remainingText };
        }
        childIndex += 1;
        visibleChildIndex += 1;
        break;
      }

      childIndex += 1;
      visibleChildIndex += 1;
    }
  }
}

/**
 * Number of rendered children that precede `xmlPosition`. Used to place
 * deleted text segments that have no model child of their own.
 */
export function paragraphVisibleCursorBefore(
  units: ParagraphVisibleUnit[],
  xmlPosition: number
): number {
  let cursor = 0;
  for (const unit of units) {
    if (unit.start >= xmlPosition) {
      break;
    }
    if (unit.visibleChildIndex !== undefined) {
      cursor = Math.max(cursor, unit.visibleChildIndex + 1);
    }
  }
  return cursor;
}
