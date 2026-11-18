// Table column-width computation, floating-table side resolution, embedded
// table runtime keys, and table spacing/border helpers.
// Upstream editor.tsx: lines 20511-21352.

import type {
  ParagraphBorderSet,
  ParagraphBorderStyle,
  ParagraphNode,
  TableBorderSet,
  TableBorderStyle,
  TableBoxSpacing,
  TableCellStyle,
  TableNode,
  TableRowStyle,
  TableStyle
} from "../../engine/types";
import {
  normalizeBorderType,
  ooxmlBorderToCss,
  ooxmlBorderWidthPx
} from "../../engine/border-css";
import { twipsToPixels, TWIPS_PER_PIXEL } from "../../viewer/section-layout";
import {
  DEFAULT_TOOLBAR_BORDER_COLOR,
  DEFAULT_TOOLBAR_BORDER_SIZE_EIGHTH_PT,
  WORD_TABLE_CELL_FALLBACK_PADDING_PX
} from "./constants";
import type { TableSpacingTwips } from "./cache-utils";
import { paragraphText } from "./paragraph-inspect";

export function tableCellText(paragraphs: ParagraphNode[]): string {
  return paragraphs.map(paragraphText).join("\n");
}

export function tableColumnCount(table: TableNode): number {
  return Math.max(
    1,
    ...table.rows.map((row) =>
      row.cells.reduce(
        (total, cell) =>
          total +
          (cell.style?.gridSpan && cell.style.gridSpan > 1
            ? cell.style.gridSpan
            : 1),
        0
      )
    )
  );
}

export function resolveFloatingTableSide(
  table: TableNode
): "left" | "right" | undefined {
  const floating = table.style?.floating;
  if (!floating) {
    return undefined;
  }

  const horizontalAlign = floating.horizontalAlign?.toLowerCase();
  if (horizontalAlign === "right" || horizontalAlign === "outside") {
    return "right";
  }
  if (horizontalAlign === "left" || horizontalAlign === "inside") {
    return "left";
  }

  // An explicit `w:horzAnchor` makes `w:tblpX` an offset from that anchor's
  // left edge, so a large offset is a table dropped to the right — not a
  // right-anchored table. Only the anchor-less case keeps the old heuristic.
  if (typeof floating.horizontalAnchor === "string" && floating.horizontalAnchor) {
    return "left";
  }

  if (Number.isFinite(floating.xTwips) && (floating.xTwips as number) > 1440) {
    return "right";
  }

  return "left";
}

export function tableWrapperStyle(
  table: TableNode,
  indentPx: number
): Record<string, string | number | undefined> {
  const floating = table.style?.floating;
  if (!floating) {
    return {
      marginLeft: indentPx,
      position: "relative",
    };
  }

  const side = resolveFloatingTableSide(table) ?? "left";
  // A table dragged with the move handle carries its drop position in `tblpX` /
  // `tblpY`. The wrapper keeps its float and its text distances and shifts by
  // that offset, so the table lands where it was dropped: the offset counts
  // from the margin on the side the text wraps away from, like Word's
  // `horzAnchor="margin"`, and from the anchor paragraph vertically.
  const offsetLeftPx = signedTwipsToPixels(floating.xTwips);
  const offsetTopPx = signedTwipsToPixels(floating.yTwips);
  const marginTop = (twipsToPixels(floating.topFromTextTwips) ?? 0) + offsetTopPx;
  const marginBottom = twipsToPixels(floating.bottomFromTextTwips) ?? 8;
  const marginLeftFromText = twipsToPixels(floating.leftFromTextTwips) ?? 8;
  const marginRightFromText = twipsToPixels(floating.rightFromTextTwips) ?? 8;

  return {
    float: side,
    clear: "none",
    marginTop,
    marginBottom,
    marginLeft:
      (side === "left" ? marginLeftFromText + indentPx : marginLeftFromText) +
      (side === "left" ? offsetLeftPx : 0),
    marginRight:
      (side === "right" ? marginRightFromText + indentPx : marginRightFromText) +
      (side === "right" ? offsetLeftPx : 0),
    position: "relative",
    zIndex: 1,
  };
}

/** Twips to CSS pixels without clamping: drop offsets may be negative. */
function signedTwipsToPixels(twips?: number): number {
  if (!Number.isFinite(twips)) {
    return 0;
  }
  return Math.round((twips as number) / TWIPS_PER_PIXEL);
}

/** Fields a caller may change on `TableStyle.floating` (`setTableFloating`). */
export interface DocxTableFloatingUpdate {
  xTwips?: number;
  yTwips?: number;
  horizontalAnchor?: "text" | "margin" | "page";
  verticalAnchor?: "text" | "margin" | "page";
  horizontalAlign?: "left" | "center" | "right" | "inside" | "outside";
  verticalAlign?: "top" | "center" | "bottom" | "inside" | "outside";
  leftFromTextTwips?: number;
  rightFromTextTwips?: number;
  topFromTextTwips?: number;
  bottomFromTextTwips?: number;
}

const TABLE_HORIZONTAL_ANCHORS = new Set(["text", "margin", "page"]);
const TABLE_HORIZONTAL_ALIGNMENTS = new Set([
  "left",
  "center",
  "right",
  "inside",
  "outside",
]);
const TABLE_VERTICAL_ALIGNMENTS = new Set([
  "top",
  "center",
  "bottom",
  "inside",
  "outside",
]);

function normalizedTableOption(
  value: unknown,
  allowed: ReadonlySet<string>
): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const normalized = value.trim().toLowerCase();
  return allowed.has(normalized) ? normalized : undefined;
}

/**
 * Merges a `setTableFloating` patch into an existing floating style.
 *
 * Non-finite numbers and unknown anchor/alignment names are refused instead of
 * being written to `w:tblpPr`, and a patch that changes nothing leaves the
 * table's current style untouched. Omitted fields keep their value, so dragging
 * a table a second time only moves it.
 */
export function mergeTableFloatingStyle(
  base: TableStyle["floating"],
  update: DocxTableFloatingUpdate
): TableStyle["floating"] {
  const next: NonNullable<TableStyle["floating"]> = { ...(base ?? {}) };
  const num = (value: unknown): number | undefined =>
    typeof value === "number" && Number.isFinite(value)
      ? Math.round(value)
      : undefined;
  const assign = (
    key: keyof DocxTableFloatingUpdate,
    value: string | number | undefined
  ): void => {
    if (value === undefined) return;
    (next as Record<string, string | number>)[key] = value;
  };

  assign("xTwips", num(update.xTwips));
  assign("yTwips", num(update.yTwips));
  assign("leftFromTextTwips", num(update.leftFromTextTwips));
  assign("rightFromTextTwips", num(update.rightFromTextTwips));
  assign("topFromTextTwips", num(update.topFromTextTwips));
  assign("bottomFromTextTwips", num(update.bottomFromTextTwips));
  assign(
    "horizontalAnchor",
    normalizedTableOption(update.horizontalAnchor, TABLE_HORIZONTAL_ANCHORS)
  );
  assign(
    "verticalAnchor",
    normalizedTableOption(update.verticalAnchor, TABLE_HORIZONTAL_ANCHORS)
  );
  assign(
    "horizontalAlign",
    normalizedTableOption(update.horizontalAlign, TABLE_HORIZONTAL_ALIGNMENTS)
  );
  assign(
    "verticalAlign",
    normalizedTableOption(update.verticalAlign, TABLE_VERTICAL_ALIGNMENTS)
  );

  return Object.keys(next).length > 0 ? next : undefined;
}

/** Rendered width of a floating table that has to stay reachable on its page. */
const MIN_TABLE_VISIBLE_PX = 42;

export interface DocxTableFloatingDrop {
  /** Floating offsets the table already has, in CSS pixels from its anchors. */
  startXPx?: number;
  startYPx?: number;
  /** Pointer movement since the drag started, in CSS pixels. */
  deltaXPx: number;
  deltaYPx: number;
  /** Page width in CSS pixels; together with the width it clamps the drop. */
  pageWidthPx?: number;
  /** Rendered table width in CSS pixels. */
  tableWidthPx?: number;
}

/**
 * The `setTableFloating` patch for a table dropped at `start + delta`.
 *
 * The drop anchors the table to the page margin horizontally and to its own
 * paragraph vertically — the pair Word writes when a table is dragged out of
 * the text flow — and carries the drop position in twips. The existing wrapper
 * style renders that position and keeps the surrounding text wrapping on the
 * other side of the table.
 */
export function resolveTableFloatingDropPatch(
  drop: DocxTableFloatingDrop
): DocxTableFloatingUpdate {
  const deltaXPx = Number.isFinite(drop.deltaXPx) ? Math.round(drop.deltaXPx) : 0;
  const deltaYPx = Number.isFinite(drop.deltaYPx) ? Math.round(drop.deltaYPx) : 0;
  const startXPx = Number.isFinite(drop.startXPx) ? Math.round(drop.startXPx as number) : 0;
  const startYPx = Number.isFinite(drop.startYPx) ? Math.round(drop.startYPx as number) : 0;

  let xPx = startXPx + deltaXPx;
  const pageWidthPx = Number.isFinite(drop.pageWidthPx)
    ? Math.round(drop.pageWidthPx as number)
    : undefined;
  if (pageWidthPx !== undefined && pageWidthPx > 0) {
    const tableWidthPx = Number.isFinite(drop.tableWidthPx)
      ? Math.max(0, Math.round(drop.tableWidthPx as number))
      : 0;
    const minXPx = tableWidthPx > 0 ? MIN_TABLE_VISIBLE_PX - tableWidthPx : 0;
    const maxXPx = Math.max(minXPx, pageWidthPx - MIN_TABLE_VISIBLE_PX);
    xPx = Math.min(Math.max(xPx, minXPx), maxXPx);
  }

  return {
    xTwips: Math.round(xPx * TWIPS_PER_PIXEL),
    yTwips: Math.round((startYPx + deltaYPx) * TWIPS_PER_PIXEL),
    horizontalAnchor: "margin",
    verticalAnchor: "text",
  };
}

export interface EmbeddedTableRuntimeKeySegment {
  rowIndex: number;
  cellIndex: number;
  contentIndex: number;
}

export interface EmbeddedTableRuntimeKeyLocation {
  hostTableIndex: number;
  hostRowIndex: number;
  hostCellIndex: number;
  rootContentIndex: number;
  descendants: EmbeddedTableRuntimeKeySegment[];
}

export function parseEmbeddedTableRuntimeKey(
  tableRuntimeKey: string
): EmbeddedTableRuntimeKeyLocation | undefined {
  const tokens = tableRuntimeKey.split("-");
  const parseIndex = (token: string | undefined): number | undefined => {
    const value = Number(token);
    if (!Number.isInteger(value) || value < 0) {
      return undefined;
    }
    return value;
  };

  if (
    tokens.length < 8 ||
    (tokens[0] !== "body" && tokens[0] !== "active") ||
    tokens[1] !== "cell" ||
    tokens[2] !== "nested" ||
    tokens[3] !== "table"
  ) {
    return undefined;
  }

  const hostTableIndex = parseIndex(tokens[4]);
  const hostRowIndex = parseIndex(tokens[5]);
  const hostCellIndex = parseIndex(tokens[6]);
  const rootContentIndex = parseIndex(tokens[7]);
  if (
    hostTableIndex === undefined ||
    hostRowIndex === undefined ||
    hostCellIndex === undefined ||
    rootContentIndex === undefined
  ) {
    return undefined;
  }

  const descendants: EmbeddedTableRuntimeKeySegment[] = [];
  let cursor = 8;
  while (cursor < tokens.length) {
    if (tokens[cursor] !== "nested" || tokens[cursor + 1] !== "table") {
      return undefined;
    }

    const rowIndex = parseIndex(tokens[cursor + 2]);
    const cellIndex = parseIndex(tokens[cursor + 3]);
    const contentIndex = parseIndex(tokens[cursor + 4]);
    if (
      rowIndex === undefined ||
      cellIndex === undefined ||
      contentIndex === undefined
    ) {
      return undefined;
    }

    descendants.push({ rowIndex, cellIndex, contentIndex });
    cursor += 5;
  }

  return {
    hostTableIndex,
    hostRowIndex,
    hostCellIndex,
    rootContentIndex,
    descendants,
  };
}

const columnWidthsByTable = new WeakMap<
  TableNode,
  Map<number, number[] | undefined>
>();

export function columnWidthsFromTableDefinition(
  table: TableNode,
  columnCount: number
): number[] | undefined {
  const cachedByCount = columnWidthsByTable.get(table);
  if (cachedByCount?.has(columnCount)) {
    return cachedByCount.get(columnCount);
  }

  const resolved = computeColumnWidthsFromTableDefinition(table, columnCount);
  const cache = cachedByCount ?? new Map<number, number[] | undefined>();
  cache.set(columnCount, resolved);
  columnWidthsByTable.set(table, cache);
  return resolved;
}

export function computeColumnWidthsFromTableDefinition(
  table: TableNode,
  columnCount: number
): number[] | undefined {
  const gridWidths = table.style?.columnWidthsTwips;
  const rowDerivedWidths = deriveColumnWidthsFromTableRows(table, columnCount);

  if (gridWidths && gridWidths.length === columnCount) {
    // Some generators emit a placeholder uniform grid while the real column
    // geometry lives in per-cell tcW. Word's fixed-layout algorithm trusts
    // cell widths over the grid, so when the two disagree on most measured
    // cells, prefer the row-derived widths.
    if (
      rowDerivedWidths &&
      rowDerivedWidths.length > 0 &&
      gridConflictsWithRowWidths(table, gridWidths)
    ) {
      return rowDerivedWidths;
    }
    return normalizeColumnWidthsTwips(gridWidths, columnCount);
  }

  if (rowDerivedWidths && rowDerivedWidths.length > 0) {
    return rowDerivedWidths;
  }

  if (gridWidths && gridWidths.length > 0) {
    return normalizeColumnWidthsTwips(gridWidths, columnCount);
  }

  return undefined;
}

export function gridConflictsWithRowWidths(
  table: TableNode,
  gridWidths: number[]
): boolean {
  let conflictRows = 0;
  let measuredRows = 0;

  for (const row of table.rows) {
    let columnCursor = 0;
    let measuredCells = 0;
    let conflictCells = 0;

    for (const cell of row.cells) {
      const span =
        cell.style?.gridSpan && cell.style.gridSpan > 1
          ? cell.style.gridSpan
          : 1;
      const expected = gridWidths
        .slice(columnCursor, columnCursor + span)
        .reduce((sum, value) => sum + Math.max(0, value), 0);
      columnCursor += span;

      const actual = cell.style?.widthTwips;
      if (!actual || actual <= 0 || expected <= 0) {
        continue;
      }
      measuredCells += 1;
      if (Math.abs(actual - expected) / expected > 0.2) {
        conflictCells += 1;
      }
    }

    if (measuredCells > 0) {
      measuredRows += 1;
      if (conflictCells * 2 > measuredCells) {
        conflictRows += 1;
      }
    }
  }

  return measuredRows > 0 && conflictRows * 2 > measuredRows;
}

export function deriveColumnWidthsFromTableRows(
  table: TableNode,
  columnCount: number
): number[] | undefined {
  let bestCandidate: number[] | undefined;
  let bestPositiveCount = -1;
  let bestTotalWidth = -1;

  for (const row of table.rows) {
    const candidate: number[] = [];
    for (const cell of row.cells) {
      const span =
        cell.style?.gridSpan && cell.style.gridSpan > 1
          ? cell.style.gridSpan
          : 1;
      const cellWidth = cell.style?.widthTwips;

      if (cellWidth && cellWidth > 0) {
        const perColumn = cellWidth / span;
        for (let index = 0; index < span; index += 1) {
          candidate.push(perColumn);
        }
        continue;
      }

      for (let index = 0; index < span; index += 1) {
        candidate.push(0);
      }
    }

    if (candidate.length !== columnCount || candidate.length === 0) {
      continue;
    }

    const positiveCount = candidate.filter((value) => value > 0).length;
    if (positiveCount <= 0) {
      continue;
    }

    const totalWidth = candidate.reduce(
      (sum, value) => sum + (value > 0 ? value : 0),
      0
    );
    if (
      positiveCount > bestPositiveCount ||
      (positiveCount === bestPositiveCount && totalWidth > bestTotalWidth)
    ) {
      bestCandidate = candidate;
      bestPositiveCount = positiveCount;
      bestTotalWidth = totalWidth;
    }
  }

  if (!bestCandidate) {
    return undefined;
  }

  return normalizeColumnWidthsTwips(bestCandidate, columnCount);
}

export function normalizeColumnWidthsTwips(
  widths: number[],
  columnCount: number
): number[] {
  const fallback = 1440 / Math.max(1, columnCount);
  const sanitized = Array.from({ length: columnCount }, (_, index) => {
    const raw = widths[index];
    if (!Number.isFinite(raw)) {
      return fallback;
    }
    return Math.max(1, raw);
  });

  if (sanitized.every((value) => value <= 0)) {
    return Array.from({ length: columnCount }, () => fallback);
  }

  return sanitized;
}

export function normalizeColumnWidthsPx(
  widths: number[],
  columnCount: number,
  fallbackTableWidthPx?: number,
  minimumWidthPx = 24
): number[] {
  const fallbackWidth =
    Number.isFinite(fallbackTableWidthPx) &&
    (fallbackTableWidthPx as number) > 0
      ? (fallbackTableWidthPx as number) / Math.max(1, columnCount)
      : 140;

  return Array.from({ length: columnCount }, (_, index) => {
    const raw = widths[index];
    if (!Number.isFinite(raw) || (raw as number) <= 0) {
      return Math.max(minimumWidthPx, Math.round(fallbackWidth));
    }
    return Math.max(minimumWidthPx, Math.round(raw as number));
  });
}

export function defaultColumnWidthsPx(
  columnCount: number,
  tableWidthPx?: number
): number[] {
  const fallbackWidth =
    Number.isFinite(tableWidthPx) && (tableWidthPx as number) > 0
      ? (tableWidthPx as number) / Math.max(1, columnCount)
      : 140;
  return Array.from({ length: columnCount }, () =>
    Math.max(24, Math.round(fallbackWidth))
  );
}

export function clampTableWidthPx(widthPx: number, maxWidthPx?: number): number {
  if (!Number.isFinite(widthPx) || widthPx <= 0) {
    return 1;
  }

  if (!Number.isFinite(maxWidthPx) || (maxWidthPx as number) <= 0) {
    return Math.max(1, Math.round(widthPx));
  }

  return Math.max(
    1,
    Math.min(Math.round(widthPx), Math.round(maxWidthPx as number))
  );
}

export function fitColumnWidthsToWidth(
  columnWidths: number[],
  targetWidthPx: number
): number[] {
  if (columnWidths.length === 0) {
    return [];
  }

  if (!Number.isFinite(targetWidthPx) || targetWidthPx <= 0) {
    return [...columnWidths];
  }

  const sanitized = columnWidths.map((value) =>
    Number.isFinite(value) && (value as number) > 0 ? (value as number) : 1
  );
  const currentTotal = sanitized.reduce((sum, value) => sum + value, 0);
  if (!Number.isFinite(currentTotal) || currentTotal <= 0) {
    const even = Math.max(1, targetWidthPx / sanitized.length);
    return Array.from({ length: sanitized.length }, () => even);
  }

  if (Math.abs(currentTotal - targetWidthPx) <= 0.5) {
    return sanitized;
  }

  if (currentTotal < targetWidthPx) {
    const scale = targetWidthPx / currentTotal;
    return sanitized.map((value) => Math.max(1, value * scale));
  }

  const minimumWidthPx = 8;
  if (sanitized.length * minimumWidthPx >= targetWidthPx) {
    const even = Math.max(1, targetWidthPx / sanitized.length);
    return Array.from({ length: sanitized.length }, () => even);
  }

  const scaled = sanitized.map((value) =>
    Math.max(minimumWidthPx, (value / currentTotal) * targetWidthPx)
  );
  let overflow = scaled.reduce((sum, value) => sum + value, 0) - targetWidthPx;
  let guard = 0;

  while (overflow > 0.25 && guard < 64) {
    const adjustableIndexes = scaled
      .map((value, index) => ({ value, index }))
      .filter((entry) => entry.value > minimumWidthPx + 0.01);
    if (adjustableIndexes.length === 0) {
      break;
    }

    const adjustableTotal = adjustableIndexes.reduce(
      (sum, entry) => sum + (entry.value - minimumWidthPx),
      0
    );
    if (adjustableTotal <= 0) {
      break;
    }

    for (const entry of adjustableIndexes) {
      const share =
        ((entry.value - minimumWidthPx) / adjustableTotal) * overflow;
      scaled[entry.index] = Math.max(minimumWidthPx, entry.value - share);
    }

    overflow = scaled.reduce((sum, value) => sum + value, 0) - targetWidthPx;
    guard += 1;
  }

  return scaled;
}

export function rowGridSpanCount(
  row: TableNode["rows"][number],
  maxColumnCount: number
): number {
  const span = row.cells.reduce((total, cell) => {
    const cellSpan =
      cell.style?.gridSpan && cell.style.gridSpan > 1 ? cell.style.gridSpan : 1;
    return total + cellSpan;
  }, 0);

  return Math.max(0, Math.min(maxColumnCount, span));
}

export function resolveFittedTableColumnWidths(
  table: TableNode,
  rawColumnWidthsPx: number[],
  targetWidthPx: number
): {
  columnWidthsPx: number[];
  effectiveColumnCount: number;
} {
  const columnCount = rawColumnWidthsPx.length;
  if (columnCount === 0) {
    return {
      columnWidthsPx: [],
      effectiveColumnCount: 0,
    };
  }

  const fallback = fitColumnWidthsToWidth(rawColumnWidthsPx, targetWidthPx);
  const rawTotalWidthPx = rawColumnWidthsPx.reduce(
    (sum, widthPx) => sum + widthPx,
    0
  );
  if (
    table.style?.layout !== "fixed" ||
    table.rows.length < 2 ||
    !Number.isFinite(rawTotalWidthPx) ||
    rawTotalWidthPx <= 0 ||
    !Number.isFinite(targetWidthPx) ||
    targetWidthPx <= 0 ||
    rawTotalWidthPx <= targetWidthPx + 0.5
  ) {
    return {
      columnWidthsPx: fallback,
      effectiveColumnCount: columnCount,
    };
  }

  const rowSpanCounts = table.rows.map((row) =>
    rowGridSpanCount(row, columnCount)
  );
  const spanFrequency = new Map<number, number>();
  rowSpanCounts.forEach((spanCount) => {
    if (spanCount <= 0 || spanCount >= columnCount) {
      return;
    }
    spanFrequency.set(spanCount, (spanFrequency.get(spanCount) ?? 0) + 1);
  });

  if (spanFrequency.size === 0) {
    return {
      columnWidthsPx: fallback,
      effectiveColumnCount: columnCount,
    };
  }

  let dominantSpanCount = columnCount;
  let dominantSpanFrequency = 0;
  for (const [spanCount, frequency] of spanFrequency.entries()) {
    if (
      frequency > dominantSpanFrequency ||
      (frequency === dominantSpanFrequency && spanCount > dominantSpanCount)
    ) {
      dominantSpanCount = spanCount;
      dominantSpanFrequency = frequency;
    }
  }

  const dominantCoverage =
    dominantSpanFrequency / Math.max(1, rowSpanCounts.length);
  if (
    dominantCoverage < 0.6 ||
    dominantSpanCount <= 0 ||
    dominantSpanCount >= columnCount
  ) {
    return {
      columnWidthsPx: fallback,
      effectiveColumnCount: columnCount,
    };
  }

  const trailingColumnWidthsPx = rawColumnWidthsPx.slice(dominantSpanCount);
  const trailingWidthPx = trailingColumnWidthsPx.reduce(
    (sum, widthPx) => sum + widthPx,
    0
  );
  const trailingRatio = trailingWidthPx / rawTotalWidthPx;
  if (
    !Number.isFinite(trailingWidthPx) ||
    trailingWidthPx < 24 ||
    trailingRatio < 0.12
  ) {
    return {
      columnWidthsPx: fallback,
      effectiveColumnCount: columnCount,
    };
  }

  const leadingWidthsPx = rawColumnWidthsPx.slice(0, dominantSpanCount);
  const fittedLeadingWidthsPx = fitColumnWidthsToWidth(
    leadingWidthsPx,
    targetWidthPx
  );
  return {
    columnWidthsPx: [
      ...fittedLeadingWidthsPx,
      ...Array.from({ length: columnCount - dominantSpanCount }, () => 0),
    ],
    effectiveColumnCount: dominantSpanCount,
  };
}

export function resolveTableSpacingPaddingPx(spacing?: TableSpacingTwips): {
  top: number;
  right: number;
  bottom: number;
  left: number;
} {
  return {
    top:
      twipsToPixels(spacing?.topTwips) ??
      WORD_TABLE_CELL_FALLBACK_PADDING_PX.top,
    right:
      twipsToPixels(spacing?.rightTwips) ??
      WORD_TABLE_CELL_FALLBACK_PADDING_PX.right,
    bottom:
      twipsToPixels(spacing?.bottomTwips) ??
      WORD_TABLE_CELL_FALLBACK_PADDING_PX.bottom,
    left:
      twipsToPixels(spacing?.leftTwips) ??
      WORD_TABLE_CELL_FALLBACK_PADDING_PX.left,
  };
}

export function tableSpacingPaddingStyle(
  spacing?: TableSpacingTwips
): Record<string, string | number | undefined> {
  const { top, right, bottom, left } = resolveTableSpacingPaddingPx(spacing);

  return {
    paddingTop: top,
    paddingRight: right,
    paddingBottom: bottom,
    paddingLeft: left,
  };
}

export function mergeTableSpacing(
  baseSpacing?: TableSpacingTwips,
  overrideSpacing?: TableSpacingTwips
): TableSpacingTwips | undefined {
  if (!baseSpacing && !overrideSpacing) {
    return undefined;
  }

  return {
    topTwips: overrideSpacing?.topTwips ?? baseSpacing?.topTwips,
    rightTwips: overrideSpacing?.rightTwips ?? baseSpacing?.rightTwips,
    bottomTwips: overrideSpacing?.bottomTwips ?? baseSpacing?.bottomTwips,
    leftTwips: overrideSpacing?.leftTwips ?? baseSpacing?.leftTwips,
  };
}

export type TableBorderSide = "top" | "right" | "bottom" | "left";

export { normalizeBorderType };

export function tableBorderToCss(
  border: TableBorderStyle | undefined
): string | undefined {
  return ooxmlBorderToCss(border);
}

export function tableBorderStrokeWidthPx(
  border: TableBorderStyle | undefined
): number {
  return ooxmlBorderWidthPx(border);
}

export function borderTypeVisible(type: string | undefined): boolean {
  const normalizedType = normalizeBorderType(type);
  return Boolean(
    normalizedType && normalizedType !== "none" && normalizedType !== "nil"
  );
}

export function paragraphBorderVisible(
  border: ParagraphBorderStyle | undefined
): boolean {
  return borderTypeVisible(border?.type);
}

export function tableBorderVisible(border: TableBorderStyle | undefined): boolean {
  return borderTypeVisible(border?.type);
}

export function tableBorderSetHasVisibleEdges(
  borders: TableBorderSet | undefined
): boolean {
  return Boolean(
    tableBorderVisible(borders?.top) ||
      tableBorderVisible(borders?.right) ||
      tableBorderVisible(borders?.bottom) ||
      tableBorderVisible(borders?.left)
  );
}

export function tableUsesSeparateBorderModel(table: TableNode): boolean {
  const explicitCellSpacingPx = twipsToPixels(table.style?.cellSpacingTwips);
  if (
    Number.isFinite(explicitCellSpacingPx) &&
    (explicitCellSpacingPx as number) > 0
  ) {
    return true;
  }

  const tableBorders = table.style?.borders;
  if (
    tableBorderVisible(tableBorders?.insideH) ||
    tableBorderVisible(tableBorders?.insideV)
  ) {
    return false;
  }

  return table.rows.some((row) =>
    row.cells.some((cell) => tableBorderSetHasVisibleEdges(cell.style?.borders))
  );
}

export function resolveTableSeparateBorderSpacingPx(table: TableNode): number {
  const explicitCellSpacingPx = twipsToPixels(table.style?.cellSpacingTwips);
  if (
    Number.isFinite(explicitCellSpacingPx) &&
    (explicitCellSpacingPx as number) > 0
  ) {
    return Math.max(0, Math.round(explicitCellSpacingPx as number));
  }

  if (
    tableUsesSeparateBorderModel(table) &&
    tableBorderSetHasVisibleEdges(table.style?.borders)
  ) {
    return 1;
  }

  return 0;
}

export function tableElementBorderStyle(
  table: TableNode,
  borderSpacingPx = resolveTableSeparateBorderSpacingPx(table)
): Record<string, string | number | undefined> {
  if (!tableUsesSeparateBorderModel(table)) {
    return {
      borderCollapse: "collapse",
    };
  }

  return {
    borderCollapse: "separate",
    borderSpacing: `${Math.max(0, Math.round(borderSpacingPx))}px`,
  };
}

export function resolvePreferredParagraphBorder(
  borders: ParagraphBorderSet | undefined
): ParagraphBorderStyle | undefined {
  if (!borders) {
    return undefined;
  }

  return (
    borders.top ??
    borders.right ??
    borders.bottom ??
    borders.left ??
    borders.between ??
    borders.bar
  );
}

export function resolvePreferredTableBorder(
  borders: TableBorderSet | undefined
): TableBorderStyle | undefined {
  if (!borders) {
    return undefined;
  }

  return (
    borders.top ??
    borders.right ??
    borders.bottom ??
    borders.left ??
    borders.insideH ??
    borders.insideV ??
    borders.tl2br ??
    borders.tr2bl
  );
}

export function toolbarParagraphBorderStyle(
  seed: ParagraphBorderStyle | undefined
): ParagraphBorderStyle {
  return {
    type: borderTypeVisible(seed?.type) ? (seed?.type as string) : "single",
    sizeEighthPt:
      Number.isFinite(seed?.sizeEighthPt) && (seed?.sizeEighthPt as number) > 0
        ? Math.round(seed?.sizeEighthPt as number)
        : DEFAULT_TOOLBAR_BORDER_SIZE_EIGHTH_PT,
    color: seed?.color ?? DEFAULT_TOOLBAR_BORDER_COLOR,
    ...(Number.isFinite(seed?.spacePt)
      ? { spacePt: Math.max(0, Math.round(seed?.spacePt as number)) }
      : undefined),
  };
}

export function toolbarTableBorderStyle(
  seed: TableBorderStyle | undefined
): TableBorderStyle {
  return {
    type: borderTypeVisible(seed?.type) ? (seed?.type as string) : "single",
    sizeEighthPt:
      Number.isFinite(seed?.sizeEighthPt) && (seed?.sizeEighthPt as number) > 0
        ? Math.round(seed?.sizeEighthPt as number)
        : DEFAULT_TOOLBAR_BORDER_SIZE_EIGHTH_PT,
    color: seed?.color ?? DEFAULT_TOOLBAR_BORDER_COLOR,
  };
}

export function nilParagraphBorderStyle(): ParagraphBorderStyle {
  return { type: "nil" };
}

export function nilTableBorderStyle(): TableBorderStyle {
  return { type: "nil" };
}

