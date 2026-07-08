import type {
  ParagraphNode,
  TableNode,
  TableRowNode,
  TableCellNode,
  TableCellStyle,
  ParagraphStyle,
  TableCellContentNode
} from "../../engine/types";
import { twipsToPixels } from "../../viewer/section-layout";
import {
  estimateParagraphLineHeightPx,
  paragraphBaseFontSizePx,
  paragraphDominantFontFamily,
  paragraphRendersTextFreeLine,
  paragraphContainsExplicitLineBreakText,
  resolveParagraphSingleLineAutoScale,
  resolveParagraphTabStopsPx,
  resolveParagraphFirstLineOriginPx,
  resolveParagraphFirstLineLeftTabStopsPx,
  resolveNextTabStopPx,
  paragraphDocGridSnapState,
  resolveParagraphDocGridLinePitchPx
} from "./line-height";
import {
  tableCellParagraphsRecursively,
  tableCellParagraphs,
  isTableCellTableContentNode,
  isParagraphCellContentNode
} from "./paragraph-inspect";

export type RecordStyle = Record<string, string | number | undefined>;

export type TableLineHeightContext = {
  table: TableNode;
  row: TableRowNode;
  cell: TableCellNode;
  cellStyle: TableCellStyle | undefined;
  paragraphStyle: ParagraphStyle | undefined;
};

// TODO: port upstream table-height-specific helpers that depend on the missing table-height module.

export function resolveTableParagraphLineHeightPx(
  paragraph: ParagraphNode,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false
): number {
  return estimateParagraphLineHeightPx(paragraph, docGridLinePitchPx, disableDocGridSnap);
}

export function resolveTableRowLineHeightPx(
  table: TableNode,
  row: TableRowNode,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false
): number {
  let maxHeightPx = 0;
  for (const cell of row.cells ?? []) {
    for (const paragraph of tableCellParagraphsRecursively(cell.nodes)) {
      maxHeightPx = Math.max(
        maxHeightPx,
        resolveTableParagraphLineHeightPx(paragraph, docGridLinePitchPx, disableDocGridSnap)
      );
    }
  }
  return maxHeightPx;
}

export function estimateTableLineHeightPx(
  table: TableNode,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false
): number {
  return Math.max(
    1,
    ...(table.rows ?? []).map((row) =>
      resolveTableRowLineHeightPx(table, row, docGridLinePitchPx, disableDocGridSnap)
    )
  );
}

export function resolveTableCellBlockStyle(
  style: TableCellStyle | undefined
): RecordStyle {
  return {
    paddingTop: twipsToPixels((style as { paddingTopTwips?: number } | undefined)?.paddingTopTwips),
    paddingRight: twipsToPixels((style as { paddingRightTwips?: number } | undefined)?.paddingRightTwips),
    paddingBottom: twipsToPixels((style as { paddingBottomTwips?: number } | undefined)?.paddingBottomTwips),
    paddingLeft: twipsToPixels((style as { paddingLeftTwips?: number } | undefined)?.paddingLeftTwips)
  };
}

export function isLineHeightRelevantParagraph(paragraph: ParagraphNode): boolean {
  return paragraphRendersTextFreeLine(paragraph) || paragraphContainsExplicitLineBreakText(paragraph);
}

export function resolveTableParagraphLineHeightContext(
  paragraph: ParagraphNode
): number {
  const fontSizePx = paragraphBaseFontSizePx(paragraph);
  const fontFamily = paragraphDominantFontFamily(paragraph);
  const scale = resolveParagraphSingleLineAutoScale(paragraph, fontFamily);
  return Math.max(1, Math.round(fontSizePx * scale));
}

export function resolveTableLineHeightTabStopsPx(paragraph: ParagraphNode): number[] {
  const originPx = resolveParagraphFirstLineOriginPx(paragraph);
  return resolveParagraphFirstLineLeftTabStopsPx(paragraph).map((value) => value + originPx);
}

export function resolveTableParagraphDocGridLinePitchPx(
  paragraph: ParagraphNode,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false
): number | undefined {
  return resolveParagraphDocGridLinePitchPx(paragraph, docGridLinePitchPx, disableDocGridSnap);
}

export function resolveTableParagraphTabStopsPx(paragraph: ParagraphNode): number[] {
  return resolveParagraphTabStopsPx(paragraph);
}

export function resolveTableParagraphFirstLineOriginPx(paragraph: ParagraphNode): number {
  return resolveParagraphFirstLineOriginPx(paragraph);
}

export function resolveTableParagraphFirstLineLeftTabStopsPx(paragraph: ParagraphNode): number[] {
  return resolveParagraphFirstLineLeftTabStopsPx(paragraph);
}

export function resolveTableParagraphNextTabStopPx(
  currentLineWidthPx: number,
  tabStopsPx: number[]
): number {
  return resolveNextTabStopPx(currentLineWidthPx, tabStopsPx);
}

export function resolveTableParagraphDocGridSnapState(paragraph: ParagraphNode): string {
  return paragraphDocGridSnapState(paragraph);
}
