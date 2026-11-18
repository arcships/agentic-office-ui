// Paragraph border style-to-CSS conversion helpers.
// Upstream editor.tsx: lines 15036-15079.

import type {
  ParagraphAlignment,
  ParagraphBorderStyle,
  ParagraphStyle,
  TableBorderStyle
} from "../../engine/types";
import { pointsToPixels } from "./ooxml-helpers";
import { normalizeBorderType, tableBorderToCss } from "./table-utils";

const JUSTIFIED_ALIGNMENTS = new Set<ParagraphAlignment>([
  "justify",
  "both",
  "distribute",
  "thaiDistribute"
]);

/** Map an OOXML `w:jc` value to a CSS `text-align` keyword. */
function paragraphAlignmentToCss(
  align: ParagraphAlignment | undefined
): string | undefined {
  if (!align) {
    return undefined;
  }
  if (JUSTIFIED_ALIGNMENTS.has(align)) {
    return "justify";
  }
  if (align === "low") {
    return "left";
  }
  return align;
}

export function isJustifiedAlignment(
  align: ParagraphAlignment | undefined
): boolean {
  return align !== undefined && JUSTIFIED_ALIGNMENTS.has(align);
}

/** CSS for paragraph direction/RTL plus extended alignment values. */
function paragraphDirectionStyle(
  style: ParagraphStyle | undefined
): Record<string, string | undefined> {
  if (!style?.direction) {
    return {};
  }
  return {
    direction: style.direction,
    unicodeBidi: "embed"
  };
}

export function paragraphTypographyStyle(
  style: ParagraphStyle | undefined
): Record<string, string | undefined> {
  return {
    textAlign: paragraphAlignmentToCss(style?.align),
    ...paragraphDirectionStyle(style)
  };
}

export function paragraphBorderToCss(
  border: ParagraphBorderStyle | undefined
): string | undefined {
  return tableBorderToCss(border as TableBorderStyle | undefined);
}

export function paragraphBorderPaddingPx(
  border: ParagraphBorderStyle | undefined
): number | undefined {
  const type = normalizeBorderType(border?.type);
  if (!type || type === "none" || type === "nil") {
    return undefined;
  }

  return pointsToPixels(border?.spacePt);
}

export function paragraphBorderStrokeWidthPx(
  border: ParagraphBorderStyle | undefined
): number {
  const type = normalizeBorderType(border?.type);
  if (!type || type === "none" || type === "nil") {
    return 0;
  }

  const sizeEighthPt = border?.sizeEighthPt;
  return Number.isFinite(sizeEighthPt) && (sizeEighthPt as number) > 0
    ? Math.max(0.5, Number(((sizeEighthPt as number) / 6).toFixed(2)))
    : 1;
}

export function paragraphBorderInsetPx(
  border: ParagraphBorderStyle | undefined
): number {
  const type = normalizeBorderType(border?.type);
  if (!type || type === "none" || type === "nil") {
    return 0;
  }

  return (
    paragraphBorderStrokeWidthPx(border) +
    (paragraphBorderPaddingPx(border) ?? 0)
  );
}
