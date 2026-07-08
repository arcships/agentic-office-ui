// Paragraph border style-to-CSS conversion helpers.
// Upstream editor.tsx: lines 15036-15079.

import type { ParagraphBorderStyle, TableBorderStyle } from "../../engine/types";
import { pointsToPixels } from "./ooxml-helpers";
import { normalizeBorderType, tableBorderToCss } from "./table-utils";

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
