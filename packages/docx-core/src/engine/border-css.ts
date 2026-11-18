// Shared OOXML border-type to CSS mapping.
//
// OOXML exposes dozens of border line styles (`wave`, `thick`,
// `thinThickSmallGap`, `triple`, ...) while CSS only accepts a small set of
// border-style keywords.  Passing an OOXML keyword straight through makes the
// whole declaration invalid, so every consumer (paragraph/table/run borders)
// goes through this single mapping.

export interface OoxmlBorderLike {
  type: string;
  color?: string;
  sizeEighthPt?: number;
}

export type CssBorderStyle =
  | "none"
  | "solid"
  | "double"
  | "dotted"
  | "dashed"
  | "groove"
  | "ridge"
  | "inset"
  | "outset";

export function normalizeBorderType(
  type: string | undefined
): string | undefined {
  if (!type) {
    return undefined;
  }

  const normalized = type.trim().toLowerCase();
  return normalized.length > 0 ? normalized : undefined;
}

interface CssBorderTypeMapping {
  style: CssBorderStyle;
  /** Minimum stroke width in CSS px; used by `thick` and compound lines. */
  minWidthPx?: number;
}

export function ooxmlBorderTypeToCss(
  type: string | undefined
): CssBorderTypeMapping | undefined {
  const normalized = normalizeBorderType(type);
  if (!normalized) {
    return undefined;
  }

  if (normalized === "none" || normalized === "nil") {
    return { style: "none" };
  }

  switch (normalized) {
    case "double":
    case "triple":
    case "thinthicksmallgap":
    case "thickthinsmallgap":
    case "thinthickmediumgap":
    case "thickthinmediumgap":
    case "thinthicklargegap":
    case "thickthinlargegap":
      return { style: "double" };
    case "dotted":
      return { style: "dotted" };
    case "dashed":
    case "dashsmallgap":
    case "dashlong":
    case "dashdotstroked":
    case "dotdash":
    case "dotdotdash":
      return { style: "dashed" };
    case "threedemboss":
      return { style: "ridge" };
    case "threedengrave":
      return { style: "groove" };
    case "inset":
      return { style: "inset" };
    case "outset":
      return { style: "outset" };
    case "thick":
      return { style: "solid", minWidthPx: 2 };
    case "wave":
    case "doublewave":
    case "wavyheavy":
    case "wavydouble":
      // CSS has no wavy border style; keep the run border visible instead of
      // emitting an invalid keyword that would drop the declaration.
      return { style: "solid", minWidthPx: 2 };
    default:
      return { style: "solid" };
  }
}

export function ooxmlBorderWidthPx(border: OoxmlBorderLike | undefined): number {
  const mapping = ooxmlBorderTypeToCss(border?.type);
  if (!mapping || mapping.style === "none") {
    return 0;
  }

  const sizeEighthPt = border?.sizeEighthPt;
  const computed =
    Number.isFinite(sizeEighthPt) && (sizeEighthPt as number) > 0
      ? (sizeEighthPt as number) / 6
      : 1;
  return Number(
    Math.max(0.5, mapping.minWidthPx ?? 0, computed).toFixed(2)
  );
}

export function ooxmlBorderToCss(
  border: OoxmlBorderLike | undefined
): string | undefined {
  const mapping = ooxmlBorderTypeToCss(border?.type);
  if (!mapping) {
    return undefined;
  }
  if (mapping.style === "none") {
    return "none";
  }

  const widthPx = ooxmlBorderWidthPx(border);
  const color = border?.color ?? "#000000";
  return `${widthPx}px ${mapping.style} ${color}`;
}
