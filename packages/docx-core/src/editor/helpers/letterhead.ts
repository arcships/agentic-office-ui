// Letterhead layout: legacy-frame detection, left/right float-side resolution
// from paragraph indents, and letterhead column-group grouping.
// Upstream editor.tsx: lines 617-831.
//
// Word letterheads are realized through legacy <w:framePr> floating paragraphs
// with a large one-sided indent. This module detects the float side from the
// indent asymmetry, confirms a nearby legacy frame exists, and groups
// consecutive same-page left/right segments into a LetterheadColumnGroup for
// two-column overlay rendering.

import type { DocModel, ParagraphNode } from "../../engine/types";
import { paragraphText } from "./paragraph-inspect";
import {
  LETTERHEAD_FRAME_NEARBY_NODE_DISTANCE,
  LETTERHEAD_INDENT_MAX_TWIPS,
  LETTERHEAD_INDENT_MIN_TWIPS,
  LETTERHEAD_MAX_TEXT_LENGTH
} from "./constants";

export type LetterheadFloatSide = "left" | "right";

// The segment-grouping entry point (paragraphLetterheadColumnGroupAtSegmentOffset)
// lives in layout/page-segmentation-core.ts as a callback-driven version
// (LetterheadColumnSegmentGroup) so the layout layer stays decoupled from the
// node-inspection helpers here. This module provides the detection primitives
// the layout layer invokes through its resolveFloatSideAtNodeIndex callback.

export function paragraphHasLegacyFrame(paragraph: ParagraphNode): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  return /<w:framePr\b[^>]*\/?>/i.test(xml);
}

export function paragraphLetterheadSideFromIndent(
  paragraph: ParagraphNode
): LetterheadFloatSide | undefined {
  const leftIndentTwipsRaw = paragraph.style?.indent?.leftTwips;
  const rightIndentTwipsRaw = paragraph.style?.indent?.rightTwips;
  const leftIndentTwips =
    Number.isFinite(leftIndentTwipsRaw) && (leftIndentTwipsRaw as number) > 0
      ? Math.round(leftIndentTwipsRaw as number)
      : 0;
  const rightIndentTwips =
    Number.isFinite(rightIndentTwipsRaw) && (rightIndentTwipsRaw as number) > 0
      ? Math.round(rightIndentTwipsRaw as number)
      : 0;
  const leftCandidate =
    rightIndentTwips >= LETTERHEAD_INDENT_MIN_TWIPS &&
    rightIndentTwips <= LETTERHEAD_INDENT_MAX_TWIPS;
  const rightCandidate =
    leftIndentTwips >= LETTERHEAD_INDENT_MIN_TWIPS &&
    leftIndentTwips <= LETTERHEAD_INDENT_MAX_TWIPS;

  if (leftCandidate && !rightCandidate) {
    return "left";
  }

  if (rightCandidate && !leftCandidate) {
    return "right";
  }

  if (leftCandidate && rightCandidate) {
    return rightIndentTwips >= leftIndentTwips ? "left" : "right";
  }

  return undefined;
}

function paragraphHasNearbyLegacyFrame(
  nodes: DocModel["nodes"],
  nodeIndex: number
): boolean {
  const startIndex = Math.max(
    0,
    nodeIndex - LETTERHEAD_FRAME_NEARBY_NODE_DISTANCE
  );
  const endIndex = Math.min(
    nodes.length - 1,
    nodeIndex + LETTERHEAD_FRAME_NEARBY_NODE_DISTANCE
  );

  for (let index = startIndex; index <= endIndex; index += 1) {
    const node = nodes[index];
    if (!node || node.type !== "paragraph") {
      continue;
    }
    if (paragraphHasLegacyFrame(node)) {
      return true;
    }
  }

  return false;
}

export function paragraphLetterheadFloatSideAtNodeIndex(
  nodes: DocModel["nodes"],
  nodeIndex: number
): LetterheadFloatSide | undefined {
  const node = nodes[nodeIndex];
  if (!node || node.type !== "paragraph") {
    return undefined;
  }

  const side = paragraphLetterheadSideFromIndent(node);
  if (!side) {
    return undefined;
  }

  const normalizedText = paragraphText(node).replace(/\s+/g, " ").trim();
  if (normalizedText.length > LETTERHEAD_MAX_TEXT_LENGTH) {
    return undefined;
  }

  if (
    paragraphHasLegacyFrame(node) ||
    paragraphHasNearbyLegacyFrame(nodes, nodeIndex)
  ) {
    return side;
  }

  return undefined;
}

export function letterheadColumnGroupContainerStyle(): Record<
  string,
  string | number | undefined
> {
  return {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    columnGap: 28,
    alignItems: "start",
    width: "100%",
    maxWidth: "100%",
    boxSizing: "border-box",
  };
}

export function letterheadColumnStackStyle(): Record<
  string,
  string | number | undefined
> {
  return {
    minWidth: 0,
    maxWidth: "100%",
  };
}

export function letterheadParagraphStyleAdjustments(
  side: LetterheadFloatSide,
  oppositeSide?: LetterheadFloatSide,
  options?: {
    suppressLetterheadColumnLayout?: boolean;
  }
): Record<string, string | number | undefined> {
  void side;
  void oppositeSide;
  if (options?.suppressLetterheadColumnLayout) {
    return {
      width: "100%",
      boxSizing: "border-box",
      marginLeft: 0,
      marginRight: 0,
    };
  }

  return {};
}
