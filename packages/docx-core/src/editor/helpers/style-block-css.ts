// Block-level style to CSS conversion extracted from upstream editor.tsx.
// Upstream editor.tsx: style-block-css helper range referenced by docs/docx-editor-helpers-split-plan.md.

import type {
  ParagraphBorderSet,
  ParagraphIndent,
  ParagraphStyleDefinition,
  TableBorderSet,
  TableCellStyle,
  TableRowStyle,
} from "../../engine/types";
import { twipsToPixels } from "../../viewer/section-layout";

type CssStyle = Record<string, string | number | undefined>;

function toFiniteNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function twipsToCssPx(value: number | undefined): string | undefined {
  if (value === undefined) {
    return undefined;
  }
  return `${twipsToPixels(value)}px`;
}

function borderStyleToCss(
  border:
    | { type?: string | null; sizeEighthPt?: number | null; color?: string | null }
    | undefined
): Record<string, string | number | undefined> {
  if (!border) {
    return {};
  }

  const type = border.type?.trim().toLowerCase();
  if (!type || type === "nil" || type === "none") {
    return {};
  }

  const width =
    toFiniteNumber(border.sizeEighthPt) && (border.sizeEighthPt as number) > 0
      ? Math.max(1, Math.round(((border.sizeEighthPt as number) / 8) * (96 / 72)))
      : 1;

  return {
    borderStyle: type === "single" ? "solid" : type,
    borderWidth: `${width}px`,
    borderColor: border.color ?? "currentColor",
  };
}

export function paragraphStyleToCss(
  style?: ParagraphStyleDefinition,
  indent?: ParagraphIndent,
  borders?: ParagraphBorderSet
): CssStyle {
  return {
    marginTop: twipsToCssPx(style?.spacing?.beforeTwips),
    marginBottom: twipsToCssPx(style?.spacing?.afterTwips),
    ...(indent?.leftTwips !== undefined ? { marginLeft: twipsToCssPx(indent.leftTwips) } : {}),
    ...(indent?.rightTwips !== undefined ? { marginRight: twipsToCssPx(indent.rightTwips) } : {}),
    ...(indent?.firstLineTwips !== undefined
      ? { textIndent: twipsToCssPx(indent.firstLineTwips) }
      : {}),
    ...(borders?.top ? { borderTopWidth: borderStyleToCss(borders.top).borderWidth } : {}),
    ...(borders?.top ? { borderTopStyle: borderStyleToCss(borders.top).borderStyle } : {}),
    ...(borders?.top ? { borderTopColor: borderStyleToCss(borders.top).borderColor } : {}),
    ...(borders?.right ? { borderRightWidth: borderStyleToCss(borders.right).borderWidth } : {}),
    ...(borders?.right ? { borderRightStyle: borderStyleToCss(borders.right).borderStyle } : {}),
    ...(borders?.right ? { borderRightColor: borderStyleToCss(borders.right).borderColor } : {}),
    ...(borders?.bottom ? { borderBottomWidth: borderStyleToCss(borders.bottom).borderWidth } : {}),
    ...(borders?.bottom ? { borderBottomStyle: borderStyleToCss(borders.bottom).borderStyle } : {}),
    ...(borders?.bottom ? { borderBottomColor: borderStyleToCss(borders.bottom).borderColor } : {}),
    ...(borders?.left ? { borderLeftWidth: borderStyleToCss(borders.left).borderWidth } : {}),
    ...(borders?.left ? { borderLeftStyle: borderStyleToCss(borders.left).borderStyle } : {}),
    ...(borders?.left ? { borderLeftColor: borderStyleToCss(borders.left).borderColor } : {}),
  };
}

export function tableCellStyleToCss(
  style?: TableCellStyle,
  rowStyle?: TableRowStyle,
  tableBorders?: TableBorderSet
): CssStyle {
  return {
    verticalAlign: style?.verticalAlign,
    backgroundColor: style?.backgroundColor,
    ...(style?.marginTwips?.topTwips !== undefined
      ? { paddingTop: twipsToCssPx(style.marginTwips.topTwips) }
      : {}),
    ...(style?.marginTwips?.rightTwips !== undefined
      ? { paddingRight: twipsToCssPx(style.marginTwips.rightTwips) }
      : {}),
    ...(style?.marginTwips?.bottomTwips !== undefined
      ? { paddingBottom: twipsToCssPx(style.marginTwips.bottomTwips) }
      : {}),
    ...(style?.marginTwips?.leftTwips !== undefined
      ? { paddingLeft: twipsToCssPx(style.marginTwips.leftTwips) }
      : {}),
    ...(rowStyle?.heightTwips !== undefined ? { height: twipsToCssPx(rowStyle.heightTwips) } : {}),
    ...(tableBorders?.top ? { borderTopWidth: borderStyleToCss(tableBorders.top).borderWidth } : {}),
    ...(tableBorders?.top ? { borderTopStyle: borderStyleToCss(tableBorders.top).borderStyle } : {}),
    ...(tableBorders?.top ? { borderTopColor: borderStyleToCss(tableBorders.top).borderColor } : {}),
    ...(tableBorders?.right ? { borderRightWidth: borderStyleToCss(tableBorders.right).borderWidth } : {}),
    ...(tableBorders?.right ? { borderRightStyle: borderStyleToCss(tableBorders.right).borderStyle } : {}),
    ...(tableBorders?.right ? { borderRightColor: borderStyleToCss(tableBorders.right).borderColor } : {}),
    ...(tableBorders?.bottom ? { borderBottomWidth: borderStyleToCss(tableBorders.bottom).borderWidth } : {}),
    ...(tableBorders?.bottom ? { borderBottomStyle: borderStyleToCss(tableBorders.bottom).borderStyle } : {}),
    ...(tableBorders?.bottom ? { borderBottomColor: borderStyleToCss(tableBorders.bottom).borderColor } : {}),
    ...(tableBorders?.left ? { borderLeftWidth: borderStyleToCss(tableBorders.left).borderWidth } : {}),
    ...(tableBorders?.left ? { borderLeftStyle: borderStyleToCss(tableBorders.left).borderStyle } : {}),
    ...(tableBorders?.left ? { borderLeftColor: borderStyleToCss(tableBorders.left).borderColor } : {}),
  };
}
